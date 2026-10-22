#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Filter, deduplicate and stably sort typed worksheet rows into JSON."""
import argparse
from collections import Counter
from decimal import DecimalException
import json
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parent))
from xlsx_analysis import region, validate_query, matches, is_blank
from xlsx_headers import resolve_column
from stdlib_xlsx import read_spec

MAX_QUERY_BYTES = 65536
TYPE_ORDER = {'boolean': 0, 'number': 1, 'date_text': 2, 'text': 3}


def read_query(path):
    return read_spec(path, max_bytes=MAX_QUERY_BYTES)


def validate_config(config, columns, schema=None):
    if not isinstance(config, dict) or set(config) - {'filters', 'deduplicate', 'sort'}:
        raise ValueError('Allowed query keys: filters, deduplicate, sort')
    _, _, _, filters = validate_query({'filters': config.get('filters', [])}, columns, schema)
    deduplicate = config.get('deduplicate')
    if deduplicate is not None:
        if not isinstance(deduplicate, dict) or set(deduplicate) - {'columns', 'keep'}:
            raise ValueError('deduplicate requires columns and optional keep')
        deduplicate = dict(deduplicate)
        keys = deduplicate.get('columns')
        if isinstance(keys, list):
            keys = [resolve_column(key, columns, schema) for key in keys]
            deduplicate['columns'] = keys
        if not isinstance(keys, list) or not keys or any(not isinstance(key, str) or key not in columns for key in keys) or len(keys) != len(set(keys)):
            raise ValueError('Deduplication columns must be a nonempty unique list inside the region')
        if deduplicate.get('keep', 'first') not in ('first', 'last', 'unique'):
            raise ValueError('keep must be first, last or unique')
    sorting = config.get('sort', [])
    if not isinstance(sorting, list):
        raise ValueError('sort must be a list')
    seen = set()
    resolved_sorting = []
    for item in sorting:
        if not isinstance(item, dict) or set(item) - {'column', 'descending', 'blanks'}:
            raise ValueError('Sort entries require column and optional descending/blanks')
        item = dict(item)
        column = resolve_column(item.get('column'), columns, schema)
        item['column'] = column
        resolved_sorting.append(item)
        if not isinstance(column, str) or column not in columns or column in seen:
            raise ValueError('Sort columns must be unique and inside the region')
        seen.add(column)
        if type(item.get('descending', False)) is not bool or item.get('blanks', 'last') not in ('first', 'last'):
            raise ValueError('descending must be boolean; blanks must be first or last')
    return filters, deduplicate, resolved_sorting


def select_rows(source, sheet, config, cell_range=None, header=True, formulas='reject',
                header_rows=None, header_exclude=(), include_spanning_headings=False):
    columns, headers, rows, metadata = region(source, sheet, cell_range, header, formulas,
                                                header_rows, header_exclude, include_spanning_headings)
    filters, deduplicate, sorting = validate_config(config, columns, metadata.get('header_schema'))
    selected = [(number, row) for number, row in rows if matches(row, filters)]
    filtered_count = len(selected)
    duplicate_groups = 0
    if deduplicate is not None:
        indices = [columns.index(key) for key in deduplicate['columns']]
        def key(record):
            return tuple(record[1][index] for index in indices)
        counts = Counter(key(record) for record in selected)
        duplicate_groups = sum(count > 1 for count in counts.values())
        keep = deduplicate.get('keep', 'first')
        chosen = {}
        if keep == 'unique':
            selected = [record for record in selected if counts[key(record)] == 1]
        else:
            for record in selected:
                marker = key(record)
                if keep == 'last' or marker not in chosen:
                    chosen[marker] = record[0]
            selected = [record for record in selected if chosen[key(record)] == record[0]]
    removed = filtered_count - len(selected)
    # Reversed stable passes implement multi-column ordering. Blank placement is
    # independent of ascending/descending and preserves earlier tie order.
    for item in reversed(sorting):
        index = columns.index(item['column'])
        blanks, values = [], []
        for record in selected:
            (blanks if is_blank(record[1][index]) else values).append(record)
        values.sort(key=lambda record: (TYPE_ORDER[record[1][index][0]], record[1][index][1]),
                    reverse=item.get('descending', False))
        selected = blanks + values if item.get('blanks', 'last') == 'first' else values + blanks
    def encode(item):
        return {'type': item[0], 'value': item[1]}
    return {'status': 'selected', **metadata, 'requested_query': config, 'columns': columns,
            'headers': None if headers is None else [encode(item) for item in headers],
            'filtered_out_rows': len(rows)-filtered_count, 'removed_duplicate_rows': removed,
            'duplicate_key_groups': duplicate_groups, 'output_rows': len(selected),
            'sort_type_order': list(TYPE_ORDER),
            'rows': [{'source_row': number, 'cells': [encode(item) for item in values]}
                     for number, values in selected]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input')
    parser.add_argument('--sheet', required=True)
    parser.add_argument('--query', required=True)
    parser.add_argument('--range', dest='cell_range')
    parser.add_argument('--no-header', action='store_true')
    parser.add_argument('--header-rows')
    parser.add_argument('--header-exclude', type=int, nargs='*', default=[])
    parser.add_argument('--include-spanning-headings', action='store_true')
    parser.add_argument('--formulas', choices=('reject', 'cached'), default='reject')
    args = parser.parse_args()
    try:
        result = select_rows(args.input, args.sheet, read_query(args.query), args.cell_range,
                             not args.no_header, args.formulas,
                             args.header_rows, args.header_exclude, args.include_spanning_headings)
        print(json.dumps(result, ensure_ascii=False, default=str))
        return 0
    except (ValueError, OSError, LookupError, DecimalException, ET.ParseError, zipfile.BadZipFile) as exc:
        print(json.dumps({'status': 'failed', 'message': str(exc)}, ensure_ascii=False))
        return 1


if __name__ == '__main__':
    sys.exit(main())
