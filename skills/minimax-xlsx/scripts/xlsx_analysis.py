#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Read-only, type-preserving XLSX region profiling and grouped analysis."""
import argparse
from collections import Counter
from decimal import Decimal, DecimalException, Inexact, ROUND_HALF_EVEN, localcontext
import json
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parent))
from stdlib_xlsx import MAX_CELLS, Workbook, address, read_spec
from xlsx_tabular import column_name


def typed_value(cell):
    value = cell.get('value')
    if value is None:
        return ('blank', None)
    if cell.get('type') == 'd':
        return ('date_text', value)
    if isinstance(value, bool):
        return ('boolean', value)
    if isinstance(value, Decimal):
        return ('number', value)
    return ('text', value)


def is_blank(item):
    return item[0] == 'blank' or item == ('text', '')


def statistics(values):
    numbers = [value for kind, value in values if kind == 'number']
    if not numbers:
        return {'count': 0, 'sum': None, 'min': None, 'max': None, 'mean': None, 'mean_rounded': False}
    # Never silently round a sum. Fail explicitly if this precision bound is exceeded.
    with localcontext() as context:
        context.prec = 1000
        context.traps[Inexact] = True
        try:
            total = sum(numbers, Decimal(0))
        except DecimalException as exc:
            raise ValueError('Numeric sum exceeds exact 1000-digit arithmetic limit') from exc
    with localcontext() as context:
        context.prec = 34
        context.rounding = ROUND_HALF_EVEN
        context.traps[Inexact] = False
        context.clear_flags()
        mean = total / len(numbers)
        rounded = context.flags[Inexact]
    return {'count': len(numbers), 'sum': total, 'min': min(numbers), 'max': max(numbers),
            'mean': mean, 'mean_rounded': rounded}


def region(source, sheet, cell_range=None, header=True, formulas='reject'):
    if formulas not in ('reject', 'cached'):
        raise ValueError('Formula policy must be reject or cached')
    book = Workbook.load(source)
    if sheet not in book.cells:
        raise ValueError('Unknown worksheet: ' + sheet)
    cells = book.cells[sheet]
    if cell_range:
        ends = cell_range.split(':')
        if len(ends) not in (1, 2):
            raise ValueError('Range must be A1 or A1:C10')
        r1, c1 = address(ends[0])
        r2, c2 = address(ends[-1])
    elif cells:
        positions = [address(ref) for ref in cells]
        r1, c1 = 1, 1
        r2, c2 = max(p[0] for p in positions), max(p[1] for p in positions)
    else:
        raise ValueError('Empty worksheet; supply an explicit range')
    if r2 < r1 or c2 < c1 or (r2-r1+1)*(c2-c1+1) > MAX_CELLS:
        raise ValueError('Reversed range or region exceeds 200000 cells')
    rows = []
    cached = 0
    for row in range(r1, r2+1):
        values = []
        for col in range(c1, c2+1):
            ref = column_name(col) + str(row)
            cell = cells.get(ref, {})
            if cell.get('formula') is not None:
                if formulas == 'reject':
                    raise ValueError('Formula at ' + ref + '; choose --formulas cached explicitly')
                if cell.get('cache_status') != 'unverified' or cell.get('value') is None:
                    raise ValueError('Missing formula cache at ' + ref)
                cached += 1
            if cell.get('type') == 'e':
                raise ValueError('Excel error at ' + ref + ': ' + str(cell.get('value')))
            values.append(typed_value(cell))
        rows.append((row, values))
    columns = [column_name(col) for col in range(c1, c2+1)]
    headers = rows.pop(0)[1] if header else None
    metadata = {'sheet': sheet, 'range': column_name(c1)+str(r1)+':'+column_name(c2)+str(r2),
                'header': header, 'source_rows': len(rows), 'unverified_formula_caches': cached,
                'date_system': '1904' if book.date1904 else '1900',
                'date_policy': 'numeric serials remain numbers; ISO date cells remain date_text; no format inference',
                'precision': 'exact sums up to 1000 digits; means 34 significant digits, ROUND_HALF_EVEN',
                'decimals': 'JSON strings; use type tags to distinguish numbers from text',
                'visibility': 'includes hidden rows/columns and filtered-out rows; merges are not expanded'}
    return columns, headers, rows, metadata


def profile(source, sheet, cell_range=None, header=True, formulas='reject'):
    columns, headers, rows, metadata = region(source, sheet, cell_range, header, formulas)
    result = []
    for index, name in enumerate(columns):
        values = [row[index] for _, row in rows]
        populated = [item for item in values if not is_blank(item)]
        types = Counter(kind for kind, _ in values)
        result.append({'column': name, 'header': None if headers is None else {'type': headers[index][0], 'value': headers[index][1]},
                       'types': dict(sorted(types.items())), 'mixed_types': len({kind for kind, _ in populated}) > 1,
                       'missing_cells': types['blank'], 'empty_text': values.count(('text', '')),
                       'blank_count': sum(is_blank(item) for item in values),
                       'distinct_nonblank': len(set(populated)),
                       'duplicate_nonblank': len(populated)-len(set(populated)),
                       'numeric': statistics(values)})
    unique_rows = {tuple(row) for _, row in rows}
    return {'status': 'profiled', **metadata, 'duplicate_rows': len(rows)-len(unique_rows), 'columns': result}


def config_value(value):
    if value is None:
        return ('blank', None)
    if isinstance(value, bool):
        return ('boolean', value)
    if isinstance(value, (int, Decimal)):
        return ('number', Decimal(value))
    if isinstance(value, str):
        return ('text', value)
    raise ValueError('Filter values must be scalar JSON values')


def validate_query(query, columns):
    if not isinstance(query, dict) or set(query) - {'group_by', 'measure', 'filters', 'mixed'}:
        raise ValueError('Query allows only group_by, measure, filters and mixed')
    groups = query.get('group_by', [])
    if not isinstance(groups, list) or any(not isinstance(col, str) or col not in columns for col in groups) or len(groups) != len(set(groups)):
        raise ValueError('group_by must contain unique column letters inside the selected range')
    measure = query.get('measure')
    if measure is not None and (not isinstance(measure, str) or measure not in columns):
        raise ValueError('measure must be a column letter inside the selected range')
    mixed = query.get('mixed', 'reject')
    if mixed not in ('reject', 'ignore'):
        raise ValueError('mixed must be reject or ignore')
    filters = query.get('filters', [])
    if not isinstance(filters, list):
        raise ValueError('filters must be a list')
    parsed = []
    for condition in filters:
        if not isinstance(condition, dict) or set(condition) != {'column', 'op', 'value'}:
            raise ValueError('Each filter requires column, op and value')
        col, op = condition['column'], condition['op']
        if not isinstance(col, str) or col not in columns or op not in ('eq', 'ne', 'gt', 'ge', 'lt', 'le'):
            raise ValueError('Invalid filter column or operator')
        value = config_value(condition['value'])
        if op not in ('eq', 'ne') and value[0] != 'number':
            raise ValueError('Ordering filters require a JSON number')
        parsed.append((columns.index(col), op, value))
    return groups, measure, mixed, parsed


def matches(row, filters):
    for index, op, expected in filters:
        actual = row[index]
        if op == 'eq':
            keep = actual == expected
        elif op == 'ne':
            keep = actual != expected
        elif actual[0] != 'number':
            keep = False
        else:
            left, right = actual[1], expected[1]
            keep = {'gt': left > right, 'ge': left >= right,
                    'lt': left < right, 'le': left <= right}[op]
        if not keep:
            return False
    return True


def summarize(source, sheet, query, cell_range=None, header=True, formulas='reject'):
    columns, _, rows, metadata = region(source, sheet, cell_range, header, formulas)
    group_columns, measure, mixed, filters = validate_query(query, columns)
    indices = [columns.index(col) for col in group_columns]
    measure_index = None if measure is None else columns.index(measure)
    buckets = {}
    selected = 0
    for row_number, row in rows:
        if not matches(row, filters):
            continue
        selected += 1
        key = tuple(row[index] for index in indices)
        if key not in buckets:
            if len(buckets) >= 10000:
                raise ValueError('More than 10000 groups; narrow the input range or filters')
            buckets[key] = {'rows': 0, 'values': [], 'blank': 0, 'ignored': 0}
        bucket = buckets[key]
        bucket['rows'] += 1
        if measure_index is not None:
            value = row[measure_index]
            if is_blank(value):
                bucket['blank'] += 1
            elif value[0] != 'number':
                if mixed == 'reject':
                    raise ValueError('Non-numeric measure at ' + measure + str(row_number) + '; choose mixed: ignore explicitly')
                bucket['ignored'] += 1
            bucket['values'].append(value)
    result = []
    for key, bucket in buckets.items():
        result.append({'keys': [{'column': col, 'type': item[0], 'value': item[1]}
                                for col, item in zip(group_columns, key)],
                       'rows': bucket['rows'], 'blank_measure': bucket['blank'],
                       'ignored_non_numeric': bucket['ignored'],
                       'numeric': None if measure is None else statistics(bucket['values'])})
    return {'status': 'summarized', **metadata, 'selected_rows': selected,
            'excluded_rows': len(rows)-selected, 'group_by': group_columns, 'measure': measure,
            'mixed': mixed, 'groups': result}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command', choices=('profile', 'summarize'))
    parser.add_argument('input')
    parser.add_argument('--sheet', required=True)
    parser.add_argument('--range', dest='cell_range')
    parser.add_argument('--no-header', action='store_true')
    parser.add_argument('--formulas', choices=('reject', 'cached'), default='reject')
    parser.add_argument('--query', help='JSON grouping/filter configuration for summarize')
    args = parser.parse_args()
    try:
        if args.command == 'profile':
            if args.query:
                raise ValueError('--query is only valid for summarize')
            result = profile(args.input, args.sheet, args.cell_range, not args.no_header, args.formulas)
        else:
            if not args.query:
                raise ValueError('summarize requires --query')
            if Path(args.query).stat().st_size > 65536:
                raise ValueError('Query exceeds 64 KiB')
            result = summarize(args.input, args.sheet, read_spec(args.query), args.cell_range, not args.no_header, args.formulas)
        print(json.dumps(result, ensure_ascii=False, default=str))
        return 0
    except (ValueError, OSError, LookupError, DecimalException, ET.ParseError, zipfile.BadZipFile) as exc:
        print(json.dumps({'status': 'failed', 'message': str(exc)}, ensure_ascii=False))
        return 1


if __name__ == '__main__':
    sys.exit(main())
