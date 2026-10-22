#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Bounded header-range suggestions, never automatic data-boundary selection."""
import argparse
from decimal import DecimalException
import json
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook, NS, address, xml
from xlsx_analysis import typed_value, is_blank
from xlsx_headers import reconstruct
from xlsx_tabular import column_name


def suggest(source, sheet, scan_rows=32, scan_columns=32, max_depth=4):
    if any(type(value) is not int for value in (scan_rows, scan_columns, max_depth)) or not 2 <= scan_rows <= 64 or not 1 <= scan_columns <= 64 or not 1 <= max_depth <= 4:
        raise ValueError('Limits: scan rows 2..64, columns 1..64, header depth 1..4')
    book = Workbook.load(source)
    if sheet not in book.cells:
        raise ValueError('Unknown worksheet: ' + sheet)
    cells = book.cells[sheet]
    positions = [address(ref) for ref in cells]
    if not positions:
        return {'status': 'ambiguous', 'candidates': [], 'ambiguities': ['empty worksheet'], 'requires_confirmation': True}
    last_row = min(scan_rows, max(row for row, _ in positions))
    last_col = min(scan_columns, max(col for _, col in positions))
    def read_cell(ref):
        cell = cells.get(ref, {})
        if cell.get('formula') is not None or cell.get('type') == 'e':
            raise ValueError('Formula/error evidence at ' + ref + '; declare rows and cache policy manually')
        return typed_value(cell)
    rows = {}
    uncertain = []
    for row in range(1, last_row+1):
        items = []
        for col in range(1, last_col+1):
            ref = column_name(col)+str(row)
            try:
                items.append(read_cell(ref))
            except ValueError:
                uncertain.append(ref)
                items.append(('unknown', None))
        rows[row] = items
    active = [row for row, items in rows.items() if any(not is_blank(item) for item in items)]
    blocks = []
    for row in active:
        if not blocks or row > blocks[-1]['last']+1:
            blocks.append({'first': row, 'last': row})
        else:
            blocks[-1]['last'] = row
    merge_nodes = xml(book.parts[book.sheets[sheet]]).findall(NS+'mergeCells/'+NS+'mergeCell')
    candidates = []
    rejected = 0
    for block_index, block in enumerate(blocks):
        for start in range(block['first'], block['last']):
            for end in range(start, min(start+max_depth-1,block['last']-1)+1):
                try:
                    schema = reconstruct(book, sheet, (1,1,last_row,last_col), '%d:%d' % (start,end), read_cell, merge_nodes=merge_nodes)
                except ValueError:
                    rejected += 1
                    continue
                if schema['context_headings']:
                    continue
                heads = [item for row in range(start,end+1) for item in rows[row] if not is_blank(item)]
                following = [item for row in range(end+1,min(end+3,block['last'])+1) for item in rows[row] if not is_blank(item)]
                if not heads or not following:
                    continue
                header_text = sum(kind == 'text' for kind,_ in heads)/len(heads)
                header_numeric = sum(kind == 'number' for kind,_ in heads)/len(heads)
                data_numeric = sum(kind == 'number' for kind,_ in following)/len(following)
                merge_refs = {level['merge'] for col in schema['columns'] for level in col['levels'] if level['merge']}
                transition = header_text >= .5 and data_numeric > header_numeric + .25
                evidence = []
                cautions = []
                score = 0
                if merge_refs:
                    score += 4 + min(len(merge_refs),4)
                    evidence.append('complete merged header structure: '+', '.join(sorted(merge_refs)))
                if transition:
                    score += 5
                    evidence.append('text-rich rows followed by a higher numeric-value share')
                if header_text >= .5:
                    score += 1
                    evidence.append('header-range stored values are mostly text')
                if end > start and not merge_refs:
                    score -= end-start
                    cautions.append('unmerged extra rows might be data or notes')
                if header_numeric == 1:
                    cautions.append('all-numeric rows could be headers or ordinary data')
                if not transition and not merge_refs:
                    cautions.append('no structural/type transition; text-only or headerless data remain possible')
                candidates.append({'header_rows': '%d:%d' % (start,end), 'data_first_row':end+1,
                                   'score':score, 'block':block_index, 'evidence':evidence,
                                   'cautions':cautions, 'strong_evidence':bool(transition or merge_refs),
                                   'paths': [col['path'] for col in schema['columns']]})
    candidates.sort(key=lambda item: (-item['score'],int(item['header_rows'].split(':')[0]),item['header_rows']))
    ambiguities = ['heuristic ranking, not calibrated probability; user confirmation required']
    if len(blocks)>1:
        ambiguities.append('multiple nonblank row blocks; could be separate tables, notes or gaps')
    if uncertain:
        ambiguities.append('formula/error cells were not interpreted as label or data evidence')
    if not candidates or not candidates[0]['strong_evidence']:
        ambiguities.append('insufficient evidence to distinguish headers from data')
    elif len(candidates)>1 and candidates[0]['score']-candidates[1]['score']<2:
        ambiguities.append('several similarly ranked ranges; review alternate depths or repeated headers')
    signatures = {}
    for item in candidates:
        signature = json.dumps(item['paths'],ensure_ascii=False,default=str,sort_keys=True)
        signatures.setdefault(signature,set()).add(item['header_rows'].split(':')[0])
    if any(len(starts)>1 for starts in signatures.values()):
        ambiguities.append('repeated header paths at different row positions are possible')
    status = 'suggested' if len(ambiguities)==1 else 'ambiguous'
    return {'status':status, 'requires_confirmation':True, 'sheet':sheet,
            'scan_range':'A1:'+column_name(last_col)+str(last_row),
            'scan_clipped': any(row>scan_rows or col>scan_columns for row,col in positions),
            'blocks':blocks, 'candidates':candidates[:8], 'candidate_count':len(candidates),
            'rejected_structures':rejected, 'uncertain_cells':uncertain,
            'ambiguities':ambiguities, 'automatic_boundary_applied':False}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input')
    parser.add_argument('--sheet',required=True)
    parser.add_argument('--scan-rows',type=int,default=32)
    parser.add_argument('--scan-columns',type=int,default=32)
    parser.add_argument('--max-depth',type=int,default=4)
    args=parser.parse_args()
    try:
        print(json.dumps(suggest(args.input,args.sheet,args.scan_rows,args.scan_columns,args.max_depth),ensure_ascii=False,default=str))
        return 0
    except (ValueError,OSError,LookupError,DecimalException,ET.ParseError,zipfile.BadZipFile) as exc:
        print(json.dumps({'status':'failed','message':str(exc)},ensure_ascii=False))
        return 1


if __name__=='__main__':
    sys.exit(main())
