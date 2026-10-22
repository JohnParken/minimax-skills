#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Bounded composite-key matching with explicit cardinality and typed diagnostics."""
import argparse
from collections import defaultdict
from decimal import DecimalException
import json
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import MAX_CELLS
from xlsx_analysis import region, is_blank
from xlsx_headers import resolve_column
from xlsx_rows import read_query


def load_side(source, spec, formulas):
    allowed={'sheet','keys','range','header_rows','header','header_exclude','include_spanning_headings'}
    if not isinstance(spec,dict) or set(spec)-allowed or not isinstance(spec.get('sheet'),str):
        raise ValueError('Each side requires sheet/keys and only documented region/header options')
    for option in ('header','include_spanning_headings'):
        if option in spec and type(spec[option]) is not bool:
            raise ValueError(option+' must be boolean')
    excluded=spec.get('header_exclude',[])
    if not isinstance(excluded,list):
        raise ValueError('header_exclude must be a list')
    columns,headers,rows,metadata=region(source,spec['sheet'],spec.get('range'),spec.get('header',True),formulas,
                                         spec.get('header_rows'),excluded,spec.get('include_spanning_headings',False))
    selectors=spec.get('keys')
    if not isinstance(selectors,list) or not selectors:
        raise ValueError('Each side needs a nonempty key selector list')
    keys=[resolve_column(selector,columns,metadata.get('header_schema')) for selector in selectors]
    if len(keys)!=len(set(keys)):
        raise ValueError('Key columns must be unique')
    return {'columns':columns,'headers':headers,'rows':rows,'metadata':metadata,
            'key_columns':keys,'indices':[columns.index(key) for key in keys]}


def index_rows(side, nulls):
    index=defaultdict(list)
    key_by_row={}
    null_rows=[]
    for number,values in side['rows']:
        key=tuple(values[col] for col in side['indices'])
        if any(is_blank(item) for item in key):
            null_rows.append(number)
            if nulls=='never':
                key_by_row[number]=None
                continue
        key_by_row[number]=key
        index[key].append((number,values))
    duplicates=[records for records in index.values() if len(records)>1]
    return index,key_by_row,{'null_key_rows':len(null_rows),'null_key_row_sample':null_rows[:20],
                            'duplicate_key_groups':len(duplicates),
                            'duplicate_key_row_samples':[[record[0] for record in records[:20]] for records in duplicates[:20]]}


def match_tables(left_source,right_source,config):
    if not isinstance(config,dict) or set(config)-{'left','right','relationship','how','nulls','formulas','max_rows'}:
        raise ValueError('Unknown matching configuration keys')
    relationship=config.get('relationship')
    how=config.get('how','left')
    nulls=config.get('nulls','never')
    formulas=config.get('formulas','reject')
    max_rows=config.get('max_rows',10000)
    if relationship not in ('one_to_one','one_to_many','many_to_one'):
        raise ValueError('Explicit relationship required: one_to_one, one_to_many or many_to_one')
    if how not in ('left','inner','full') or nulls not in ('never','match'):
        raise ValueError('how must be left/inner/full; nulls must be never/match')
    if type(max_rows) is not int or not 1<=max_rows<=10000:
        raise ValueError('max_rows must be 1..10000')
    left=load_side(left_source,config.get('left'),formulas)
    right=load_side(right_source,config.get('right'),formulas)
    if len(left['indices'])!=len(right['indices']):
        raise ValueError('Left and right composite keys need the same number of components')
    _,left_keys,left_diagnostics=index_rows(left,nulls)
    right_index,_,right_diagnostics=index_rows(right,nulls)
    violations=[]
    if relationship in ('one_to_one','one_to_many') and left_diagnostics['duplicate_key_groups']:
        violations.append('left keys are not unique')
    if relationship in ('one_to_one','many_to_one') and right_diagnostics['duplicate_key_groups']:
        violations.append('right keys are not unique')
    if violations:
        return {'status':'failed','message':'; '.join(violations),'left':left_diagnostics,'right':right_diagnostics,'rows':[]}
    result=[]
    unmatched_left=[]
    matched_right=set()
    width=len(left['columns'])+len(right['columns'])
    def encode(values):
        return None if values is None else [{'type':kind,'value':value} for kind,value in values]
    def emit(left_record,right_record):
        if len(result)>=max_rows or (len(result)+1)*width>MAX_CELLS:
            raise ValueError('Match output exceeds row/cell budget; narrow inputs or use a smaller range')
        result.append({'left_row':None if left_record is None else left_record[0],
                       'right_row':None if right_record is None else right_record[0],
                       'left_cells':None if left_record is None else encode(left_record[1]),
                       'right_cells':None if right_record is None else encode(right_record[1])})
    for record in left['rows']:
        key=left_keys[record[0]]
        matches=[] if key is None else right_index.get(key,[])
        if not matches:
            unmatched_left.append(record[0])
            if how!='inner':
                emit(record,None)
        for other in matches:
            matched_right.add(other[0])
            emit(record,other)
    unmatched_right=[record for record in right['rows'] if record[0] not in matched_right]
    if how=='full':
        for record in unmatched_right:
            emit(None,record)
    return {'status':'matched','relationship':relationship,'how':how,'nulls':nulls,
            'left':{**left['metadata'],**left_diagnostics,'columns':left['columns'],'key_columns':left['key_columns'],
                    'unmatched_rows':len(unmatched_left),'unmatched_row_sample':unmatched_left[:20]},
            'right':{**right['metadata'],**right_diagnostics,'columns':right['columns'],'key_columns':right['key_columns'],
                     'unmatched_rows':len(unmatched_right),'unmatched_row_sample':[record[0] for record in unmatched_right[:20]]},
            'output_rows':len(result),'rows':result}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('left')
    parser.add_argument('right')
    parser.add_argument('--query',required=True)
    args=parser.parse_args()
    try:
        result=match_tables(args.left,args.right,read_query(args.query))
        print(json.dumps(result,ensure_ascii=False,default=str))
        return 1 if result['status']=='failed' else 0
    except (ValueError,OSError,LookupError,DecimalException,ET.ParseError,zipfile.BadZipFile) as exc:
        print(json.dumps({'status':'failed','message':str(exc),'rows':[]},ensure_ascii=False))
        return 1


if __name__=='__main__':
    sys.exit(main())
