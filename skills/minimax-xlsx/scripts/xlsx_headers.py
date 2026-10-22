#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Explicit multi-row header reconstruction with merge provenance and typed paths."""
import argparse
from decimal import Decimal, DecimalException
import json
from pathlib import Path
import re
import sys
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parent))
from stdlib_xlsx import MAX_CELLS, NS, address, xml
from xlsx_tabular import column_name


def header_bounds(spec, first_row, last_row):
    if not isinstance(spec, str) or not re.fullmatch(r'[1-9][0-9]{0,6}(?::[1-9][0-9]{0,6})?', spec):
        raise ValueError('header_rows must be an explicit row or row range, such as 3:5')
    ends = [int(item) for item in spec.split(':')]
    start, end = ends[0], ends[-1]
    if not first_row <= start <= end <= last_row:
        raise ValueError('Header rows must be fully inside the selected region')
    return start, end


def reconstruct(book, sheet, bounds, header_rows, read_cell, exclude=(), include_spanning=False, merge_nodes=None):
    r1, c1, r2, c2 = bounds
    start, end = header_bounds(header_rows, r1, r2)
    if any(type(row) is not int or not start <= row <= end for row in exclude):
        raise ValueError('Excluded header rows must lie inside header_rows')
    excluded = set(exclude)
    if len(excluded) == end-start+1:
        raise ValueError('At least one header row must remain')
    if merge_nodes is None:
        tree = xml(book.parts[book.sheets[sheet]])
        merge_nodes = tree.findall(NS + 'mergeCells/' + NS + 'mergeCell')
    merges = merge_nodes
    if len(merges) > 10000:
        raise ValueError('More than 10000 merged regions')
    occupied = {}
    contexts = []
    for node in merges:
        ref = node.get('ref', '')
        parts = ref.split(':')
        if len(parts) != 2:
            raise ValueError('Malformed merged region: ' + ref)
        top, left = address(parts[0])
        bottom, right = address(parts[1])
        if bottom < top or right < left or (bottom, right) == (top, left):
            raise ValueError('Malformed merged region: ' + ref)
        if bottom < start or top > end or right < c1 or left > c2:
            continue
        if top < start or bottom > end:
            raise ValueError('Merge crosses the declared header boundary: ' + ref)
        if left < c1 or right > c2:
            raise ValueError('Region clips a header merge; select the complete merge: ' + ref)
        if (bottom-top+1)*(right-left+1) > MAX_CELLS:
            raise ValueError('Merged header exceeds cell limit')
        anchor = column_name(left)+str(top)
        spanning = left == c1 and right == c2 and left < right and top == bottom
        if spanning:
            contexts.append({'row': top, 'source': anchor, 'merge': ref,
                             'included_in_path': include_spanning and top not in excluded,
                             'reason': 'full-width merged heading; semantic role requires explicit choice'})
        for row in range(top, bottom+1):
            for col in range(left, right+1):
                key = (row, col)
                if key in occupied:
                    raise ValueError('Overlapping header merges at ' + column_name(col)+str(row))
                coord = column_name(col)+str(row)
                cell = book.cells[sheet].get(coord, {})
                if coord != anchor and (cell.get('formula') is not None or cell.get('value') not in (None, '')):
                    raise ValueError('Non-anchor merged header cell contains data: ' + coord)
                occupied[key] = (anchor, ref, spanning)
    columns = []
    for col in range(c1, c2+1):
        letter = column_name(col)
        levels, path, seen_sources = [], [], set()
        for row in range(start, end+1):
            coordinate = letter+str(row)
            source, merge, spanning = occupied.get((row, col), (coordinate, None, False))
            kind, value = read_cell(source)
            role = 'excluded' if row in excluded else 'spanning_heading' if spanning and not include_spanning else 'label'
            component = {'type': kind, 'value': value}
            levels.append({'row': row, 'coordinate': coordinate, 'source': source,
                           'merge': merge, 'propagated': coordinate != source,
                           'role': role, **component})
            if role == 'label' and source not in seen_sources:
                path.append(component)
                seen_sources.add(source)
        columns.append({'id': 'col:'+letter, 'column': letter, 'path': path, 'levels': levels})
    return {'rows': {'first': start, 'last': end}, 'excluded_rows': sorted(excluded),
            'ignored_leading_rows': list(range(r1, start)), 'data_first_row': end+1,
            'columns': columns, 'context_headings': contexts,
            'policy': 'explicit rows; merge-only propagation; blanks retained; vertical merge sources appear once in path'}


def selector_value(value):
    if value is None:
        return ('blank', None)
    if isinstance(value, bool):
        return ('boolean', value)
    if isinstance(value, (int, Decimal)):
        return ('number', Decimal(value))
    if isinstance(value, str):
        return ('text', value)
    if isinstance(value, dict) and set(value) == {'type', 'value'}:
        kind = value['type']
        item = value['value']
        if kind == 'date_text' and isinstance(item, str):
            return (kind, item)
        if kind == 'number' and isinstance(item, str):
            number = Decimal(item)
            if number.is_finite():
                return (kind, number)
        inferred = selector_value(item) if not isinstance(item, dict) else None
        if inferred is not None and inferred[0] == kind:
            return inferred
    raise ValueError('Path components must be scalar values or typed {type,value} objects')


def resolve_column(selector, columns, schema=None):
    if isinstance(selector, str):
        letter = selector[4:] if selector.startswith('col:') else selector
        if letter in columns:
            return letter
    elif isinstance(selector, dict) and set(selector) == {'id'}:
        if isinstance(selector['id'], str) and selector['id'].startswith('col:'):
            return resolve_column(selector['id'], columns, schema)
    elif isinstance(selector, dict) and set(selector) == {'path'} and isinstance(selector['path'], list):
        if schema is None:
            raise ValueError('Path selection requires explicit --header-rows')
        wanted = [selector_value(value) for value in selector['path']]
        matches = [item['column'] for item in schema['columns']
                   if [(part['type'], part['value']) for part in item['path']] == wanted]
        if len(matches) == 1:
            return matches[0]
        raise ValueError('Header path matched %d columns; select an explicit col: ID' % len(matches))
    raise ValueError('Unknown column selector; use a column letter, col: ID or {path:[...]}')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input')
    parser.add_argument('--sheet', required=True)
    parser.add_argument('--range', dest='cell_range')
    parser.add_argument('--header-rows', required=True)
    parser.add_argument('--header-exclude', type=int, nargs='*', default=[])
    parser.add_argument('--include-spanning-headings', action='store_true')
    parser.add_argument('--formulas', choices=('reject', 'cached'), default='reject')
    args = parser.parse_args()
    try:
        from xlsx_analysis import region
        _, _, _, metadata = region(args.input, args.sheet, args.cell_range, True, args.formulas,
                                    args.header_rows, args.header_exclude, args.include_spanning_headings)
        print(json.dumps({'status': 'headers_reconstructed', **metadata}, ensure_ascii=False, default=str))
        return 0
    except (ValueError, OSError, LookupError, DecimalException, ET.ParseError, zipfile.BadZipFile) as exc:
        print(json.dumps({'status': 'failed', 'message': str(exc)}, ensure_ascii=False))
        return 1


if __name__ == '__main__':
    sys.exit(main())
