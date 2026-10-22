#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Explicit within-group ranking and signed numeric-sum shares; no contribution model."""
import argparse
from decimal import Decimal, DecimalException
import json
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import MAX_CELLS
from xlsx_analysis import region,validate_query,matches,is_blank
from xlsx_advanced import exact,rounded,precision_option
from xlsx_rows import read_query


def analyze(source,sheet,config,**options):
    allowed={'group_by','measure','filters','rank','descending','missing','non_numeric','denominator','precision'}
    if not isinstance(config,dict) or set(config)-allowed:
        raise ValueError('Unknown rank/share configuration keys')
    method=config.get('rank','competition')
    descending=config.get('descending',True)
    missing=config.get('missing','omit')
    non_numeric=config.get('non_numeric','reject')
    if method not in ('competition','dense','ordinal') or type(descending) is not bool:
        raise ValueError('rank must be competition/dense/ordinal; descending must be boolean')
    if missing not in ('omit','reject') or non_numeric not in ('omit','reject'):
        raise ValueError('missing/non_numeric must be omit or reject')
    if config.get('denominator','group_numeric_sum')!='group_numeric_sum':
        raise ValueError('Only denominator: group_numeric_sum is supported')
    precision=precision_option(config)
    columns,headers,rows,metadata=region(source,sheet,**options)
    groups,measure,_,filters=validate_query({key:config[key] for key in ('group_by','measure','filters') if key in config},columns,metadata.get('header_schema'))
    if measure is None:
        raise ValueError('An explicit measure selector is required')
    selected=[record for record in rows if matches(record[1],filters)]
    if len(columns)+4>16384 or len(selected)*(len(columns)+4)>MAX_CELLS:
        raise ValueError('Rank/share typed output exceeds Excel column or 200000 cell budget')
    indices=[columns.index(column) for column in groups]
    measure_index=columns.index(measure)
    buckets={}
    outcomes={}
    for number,values in selected:
        key=tuple(values[index] for index in indices)
        if key not in buckets:
            if len(buckets)>=10000:
                raise ValueError('More than 10000 ranking groups')
            buckets[key]={'valid':[],'missing':0,'nonnumeric':0}
        bucket=buckets[key]
        item=values[measure_index]
        if is_blank(item):
            if missing=='reject':
                raise ValueError('Missing ranking measure at '+measure+str(number))
            bucket['missing']+=1
            outcomes[number]=(None,None,'missing_omitted',None)
        elif item[0]!='number':
            if non_numeric=='reject':
                raise ValueError('Non-numeric ranking measure at '+measure+str(number))
            bucket['nonnumeric']+=1
            outcomes[number]=(None,None,'non_numeric_omitted',None)
        else:
            bucket['valid'].append((number,item[1]))
    summaries=[]
    for key,bucket in buckets.items():
        valid=bucket['valid']
        total=exact(lambda:sum((value for _,value in valid),Decimal(0))) if valid else None
        ordered=sorted(valid,key=lambda record:record[1],reverse=descending)
        previous=None
        dense=competition=0
        for position,(number,value) in enumerate(ordered,1):
            if previous is None or value!=previous:
                dense+=1
                competition=position
            rank=position if method=='ordinal' else dense if method=='dense' else competition
            ratio=None if total==0 else rounded(lambda:value/total,precision)
            outcomes[number]=(Decimal(rank),None if ratio is None else ratio['value'],
                              'zero_total' if ratio is None else 'ok',None if ratio is None else ratio['rounded'])
            previous=value
        summaries.append({'keys':[{'column':column,'type':item[0],'value':item[1]} for column,item in zip(groups,key)],
                          'numeric_count':len(valid),'numeric_total':total,'negative_numeric_count':sum(value<0 for _,value in valid),
                          'missing_omitted':bucket['missing'],'non_numeric_omitted':bucket['nonnumeric'],
                          'status':'insufficient_numeric_samples' if total is None else 'zero_total' if total==0 else 'ok'})
    generated=[]
    for identifier,label in (('calc:rank','Rank'),('calc:share','Group share'),('calc:share_status','Share status'),('calc:share_rounded','Share rounded')):
        generated.append({'column':identifier,'id':identifier,'path':[{'type':'text','value':label}],
                          'levels':[{'type':'text','value':label,'role':'derived'}]})
    output=[]
    for number,values in selected:
        rank,share,status,inexact=outcomes[number]
        derived=[('blank',None) if rank is None else ('number',rank),
                 ('blank',None) if share is None else ('number',share),('text',status),
                 ('blank',None) if inexact is None else ('boolean',inexact)]
        output.append({'source_row':number,'cells':[{'type':kind,'value':value} for kind,value in values+derived]})
    metadata.pop('precision',None)
    return {'status':'selected',**metadata,'analysis':'rank_share','requested_query':config,
            'columns':columns+[item['column'] for item in generated],
            'headers':(headers_as_cells(headers,columns)+[{'type':'text','value':item['path'][0]['value']} for item in generated]),
            'generated_columns':generated,'output_rows':len(output),'filtered_out_rows':len(rows)-len(selected),'rows':output,
            'ranking_definition':{'method':method,'descending':descending,'ordinal_tie_break':'original source row order','data_order_changed':False},
            'share_definition':{'denominator':'sum of valid numeric observations in each group after filters',
                                'negative_values':'signed sum; shares may be negative or greater than one',
                                'zero_total':'null share with zero_total diagnostic; ranks still computed',
                                'not_growth_contribution':True,'unit':'fraction'},
            'group_statistics':summaries,'arithmetic':{'exact_sum_precision_limit':1000,'result_precision':precision,
                                                       'intermediate_precision':precision+16,'rounding':'ROUND_HALF_EVEN'}}


def headers_as_cells(headers,columns):
    return [{'type':'text','value':column} for column in columns] if headers is None else [{'type':kind,'value':value} for kind,value in headers]


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input')
    parser.add_argument('--sheet',required=True)
    parser.add_argument('--query',required=True)
    parser.add_argument('--range',dest='cell_range')
    parser.add_argument('--no-header',action='store_true')
    parser.add_argument('--header-rows')
    parser.add_argument('--header-exclude',type=int,nargs='*',default=[])
    parser.add_argument('--include-spanning-headings',action='store_true')
    parser.add_argument('--formulas',choices=('reject','cached'),default='reject')
    args=parser.parse_args()
    try:
        result=analyze(args.input,args.sheet,read_query(args.query),cell_range=args.cell_range,header=not args.no_header,
                       header_rows=args.header_rows,header_exclude=args.header_exclude,
                       include_spanning_headings=args.include_spanning_headings,formulas=args.formulas)
        print(json.dumps(result,ensure_ascii=False,default=str))
        return 0
    except (ValueError,OSError,LookupError,DecimalException,ET.ParseError,zipfile.BadZipFile) as exc:
        print(json.dumps({'status':'failed','message':str(exc)},ensure_ascii=False))
        return 1


if __name__=='__main__':
    sys.exit(main())
