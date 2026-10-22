#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Explicit, bounded CSV/TSV interchange using the local OOXML implementation."""
import argparse
import csv
import io
import json
import re
import sys
from decimal import InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook, MAX_CELLS, address, numeric

MAX_INPUT_BYTES = 16 * 1024 * 1024


def column_name(number):
    result = ''
    while number:
        number, digit = divmod(number - 1, 26)
        result = chr(65 + digit) + result
    return result


def delimiter(path):
    suffix = Path(path).suffix.lower()
    if suffix not in ('.csv', '.tsv'):
        raise ValueError('Tabular file must end in .csv or .tsv')
    return '\t' if suffix == '.tsv' else ','


def import_table(source, output, sheet='Sheet1', encoding='utf-8-sig', numeric_columns=(), header_rows=1):
    """Import literal text; convert only explicitly selected column letters."""
    if type(header_rows) is not int or not 0 <= header_rows <= 1048576:
        raise ValueError('header_rows must be between 0 and 1048576')
    columns = set()
    for column in numeric_columns:
        if not re.fullmatch(r'[A-Za-z]{1,3}', column):
            raise ValueError('Numeric columns must be column letters, such as B or AA')
        columns.add(address(column + '1')[1])
    separator = delimiter(source)
    with open(source, 'rb') as handle:
        raw = handle.read(MAX_INPUT_BYTES + 1)
    if len(raw) > MAX_INPUT_BYTES:
        raise ValueError('Tabular input exceeds 16 MiB limit')
    text = raw.decode(encoding)
    cells = {}
    width = None
    rows = 0
    for rows, values in enumerate(csv.reader(io.StringIO(text, newline=''), delimiter=separator, strict=True), 1):
        if rows > 1048576:
            raise ValueError('Too many rows')
        if width is None:
            width = len(values)
        if not width or len(values) != width:
            raise ValueError('Expected a rectangular table; inconsistent row %d' % rows)
        if width > 16384 or rows * width > MAX_CELLS:
            raise ValueError('Table exceeds supported cell limit')
        for col, value in enumerate(values, 1):
            ref = column_name(col) + str(rows)
            if col in columns and rows > header_rows and value != '':
                if not re.fullmatch(r'[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?', value):
                    raise ValueError('Invalid numeric value at ' + ref)
                value = numeric(value)
                if len(value.normalize().as_tuple().digits) > 15:
                    raise ValueError('More than 15 significant digits at ' + ref + '; import as text')
            cells[ref] = {'value': value}
    if not rows:
        raise ValueError('Empty tabular input')
    if any(col > width for col in columns):
        raise ValueError('Numeric column lies outside input table')
    book = Workbook.create({'sheets': [{'name': sheet, 'cells': cells}]})
    book.save(output)
    return {'status': 'imported', 'rows': rows, 'columns': width, 'output_written': True,
            'numeric_columns': sorted(columns), 'formulas_created': 0}


def export_table(source, output, sheet, cell_range=None, formula_mode='reject', text_mode='safe'):
    """Export raw scalar values, with explicit policies for formulas and text."""
    if formula_mode not in ('reject', 'cached') or text_mode not in ('safe', 'raw'):
        raise ValueError('Invalid formula/text policy')
    separator = delimiter(output)
    book = Workbook.load(source)
    if sheet not in book.cells:
        raise ValueError('Unknown sheet: ' + sheet)
    cells = book.cells[sheet]
    if cell_range:
        ends = cell_range.split(':')
        if len(ends) not in (1, 2):
            raise ValueError('Range must be A1 or A1:C10')
        first = address(ends[0])
        last = address(ends[-1])
    elif cells:
        first = (1, 1)
        positions = [address(ref) for ref in cells]
        last = (max(p[0] for p in positions), max(p[1] for p in positions))
    else:
        raise ValueError('Empty sheet; specify a range to export blanks')
    r1, c1 = first
    r2, c2 = last
    if r2 < r1 or c2 < c1 or (r2-r1+1)*(c2-c1+1) > MAX_CELLS:
        raise ValueError('Reversed range or export exceeds supported cell limit')
    buffer = io.StringIO(newline='')
    writer = csv.writer(buffer, delimiter=separator, lineterminator='\r\n')
    cached = escaped = 0
    for row in range(r1, r2+1):
        values = []
        for col in range(c1, c2+1):
            ref = column_name(col) + str(row)
            cell = cells.get(ref, {})
            value = cell.get('value')
            if cell.get('formula') is not None:
                if formula_mode == 'reject':
                    raise ValueError('Formula at ' + ref + '; cached export requires --formulas cached')
                if cell.get('cache_status') != 'unverified' or value is None:
                    raise ValueError('Missing formula cache at ' + ref)
                cached += 1
            if cell.get('type') == 'e':
                raise ValueError('Excel error at ' + ref + ': ' + str(value))
            if isinstance(value, str) and text_mode == 'safe' and (value.lstrip().startswith(('=', '+', '-', '@')) or value.startswith(('\t', '\r', '\n'))):
                value = "'" + value
                escaped += 1
            values.append('' if value is None else 'TRUE' if value is True else 'FALSE' if value is False else str(value))
        writer.writerow(values)
    # Encode and validate everything before exclusive output creation.
    payload = buffer.getvalue().encode('utf-8-sig')
    created = False
    try:
        with open(output, 'xb') as handle:
            created = True
            handle.write(payload)
    except BaseException:
        if created:
            Path(output).unlink(missing_ok=True)
        raise
    return {'status': 'exported', 'rows': r2-r1+1, 'columns': c2-c1+1,
            'output_written': True, 'unverified_formula_caches': cached,
            'escaped_text_cells': escaped, 'text_mode': text_mode,
            'values': 'raw values; number formats and date display are not applied'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    load = commands.add_parser('import')
    load.add_argument('input')
    load.add_argument('output')
    load.add_argument('--sheet', default='Sheet1')
    load.add_argument('--encoding', default='utf-8-sig')
    load.add_argument('--numeric-columns', nargs='*', default=[])
    load.add_argument('--header-rows', type=int, default=1)
    export = commands.add_parser('export')
    export.add_argument('input')
    export.add_argument('output')
    export.add_argument('--sheet', required=True)
    export.add_argument('--range', dest='cell_range')
    export.add_argument('--formulas', choices=('reject', 'cached'), default='reject')
    export.add_argument('--text', choices=('safe', 'raw'), default='safe')
    args = parser.parse_args()
    try:
        if args.command == 'import':
            result = import_table(args.input, args.output, args.sheet, args.encoding,
                                  args.numeric_columns, args.header_rows)
        else:
            result = export_table(args.input, args.output, args.sheet, args.cell_range,
                                  args.formulas, args.text)
        print(json.dumps(result, ensure_ascii=False))
        return 0
    except (ValueError, OSError, LookupError, csv.Error, InvalidOperation, zipfile.BadZipFile, ET.ParseError) as exc:
        print(json.dumps({'status': 'failed', 'output_written': False, 'message': str(exc)}, ensure_ascii=False))
        return 1


if __name__ == '__main__':
    sys.exit(main())
