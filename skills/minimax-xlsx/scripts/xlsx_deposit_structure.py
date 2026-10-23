#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Explicit deposit classification shares and account/customer Top N concentration."""
import argparse
from decimal import Decimal,DecimalException
import json
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook,read_spec
from xlsx_deposits import snapshot
from xlsx_daily import calendar_date
from xlsx_analysis import region
from xlsx_headers import resolve_column
from xlsx_advanced import exact,rounded,precision_option
from xlsx_result_export import Report,MAX_INPUT_BYTES


def analyze(source,config):
    required={'snapshot','date','scope','concentration_grain','top_n','classifications'}
    if not isinstance(config,dict) or not required<=set(config) or set(config)-required-{'customer_column','precision'}:
        raise ValueError('Structure requires snapshot/date/scope/concentration_grain/top_n/classifications')
    calendar_date(config['date'],'date'); scope=config['scope']
    if not isinstance(scope,dict) or set(scope)!={'currency','unit','version','level'} or any(not isinstance(value,str) or not value.strip() for value in scope.values()):
        raise ValueError('Require one explicit currency/unit/version/level')
    grain=config['concentration_grain']
    if grain not in ('account','customer') or (grain=='customer')!=('customer_column' in config):
        raise ValueError('Customer concentration requires stable anonymized customer_column; account concentration must omit it')
    top=config['top_n']
    if not isinstance(top,list) or not top or len(top)>10 or any(type(value) is not int or not 1<=value<=1000 for value in top) or len(set(top))!=len(top):
        raise ValueError('top_n requires 1..10 unique positive integers no greater than 1000')
    precision=precision_option(config)
    accounts,metadata=snapshot(source,config['snapshot'],config['date'],scope)
    customers={}
    if grain=='customer':
        spec=config['snapshot']; columns,_,rows,meta=region(source,spec['sheet'],cell_range=spec.get('range'),header_rows=spec.get('header_rows'))
        column=resolve_column(config['customer_column'],columns,meta.get('header_schema'))
        if column in metadata['columns'].values(): raise ValueError('Customer column must differ from account and metadata columns')
        by_row={number:values[columns.index(column)] for number,values in rows}
        for account,record in accounts.items():
            value=by_row[record['source_row']]
            if value[0]!='text' or not value[1].strip(): raise ValueError('Customer IDs must be nonempty literal text, never inferred from accounts')
            customers[account]=value[1]
    classifications=config['classifications']
    if not isinstance(classifications,list) or not 1<=len(classifications)<=10:
        raise ValueError('Declare 1..10 explicit independent classification axes')
    ids=set()
    for axis in classifications:
        if not isinstance(axis,dict) or set(axis)!={'id','field','mapping'} or not isinstance(axis['id'],str) or not axis['id'] or axis['id'] in ids:
            raise ValueError('Classification requires unique id, field and mapping')
        ids.add(axis['id'])
        if axis['field'] not in ('product','entity','customer') or axis['field']=='customer' and grain!='customer':
            raise ValueError('Classification field must be explicit product/entity or available customer')
        mapping=axis['mapping']
        if not isinstance(mapping,dict) or not mapping or any(not isinstance(key,str) or not isinstance(value,str) or not key or not value for key,value in mapping.items()):
            raise ValueError('Mapping must assign each source value exactly one literal category; no lists/overlaps')
    total=exact(lambda:sum((row['balance'] for row in accounts.values()),Decimal(0)))
    structure=[]; lineage=[]; entities={}
    def share(value):
        if total==0: return {'value':None,'rounded':False,'status':'undefined_zero_total'}
        result=rounded(lambda:value/total,precision)
        return {**result,'status':'calculated'}
    for account,record in accounts.items():
        entity=customers[account] if grain=='customer' else account
        entities.setdefault(entity,[]).append(record['balance'])
        lineage.append({'account':account,'customer':customers.get(account),'file':metadata['file'],'sheet':metadata['sheet'],'source_row':record['source_row']})
    for axis in classifications:
        groups={}
        for account,record in accounts.items():
            value=customers[account] if axis['field']=='customer' else record[axis['field']]
            if value not in axis['mapping']: raise ValueError('Unmapped classification value for '+axis['id']+': '+value)
            category=axis['mapping'][value]
            groups.setdefault(category,[]).append(record['balance'])
        classified=Decimal(0)
        for category,values in groups.items():
            value=exact(lambda:sum(values,Decimal(0))); classified=exact(lambda:classified+value)
            structure.append({'axis':axis['id'],'category':category,'balance':value,'account_count':len(values),'share':share(value)})
        if classified!=total: raise ValueError('Classification total failed reconciliation')
    balances=[{'entity':entity,'balance':exact(lambda:sum(values,Decimal(0))),'account_count':len(values)} for entity,values in entities.items()]
    balances.sort(key=lambda item:item['balance'],reverse=True)
    concentration=[]
    for number in top:
        chosen=balances[:number]; value=exact(lambda:sum((row['balance'] for row in chosen),Decimal(0)))
        concentration.append({'requested_top_n':number,'effective_n':len(chosen),'population':len(balances),'grain':grain,'balance':value,'share':share(value),'selected_entities':[row['entity'] for row in chosen]})
    return {'status':'deposit_structure','schema_version':'deposit_structure_v1','requested_query':config,'source':metadata,
            'total_balance':total,'structure':structure,'concentration':concentration,'ranked_entities':balances,'lineage':lineage,
            'quality_flags':['zero_total_balance'] if total==0 else [],
            'policy':{'classification':'Independent axes; each account appears exactly once per axis; never sum across axes',
                      'concentration':'Aggregate account balances by declared stable entity before Top N; exact N, stable first-observed tie order, zero balances retained',
                      'scope':scope,'grain':grain,'precision':precision,'rounding':'ROUND_HALF_EVEN','exact_precision_limit':1000,
                      'unsupported':['interest expense/cost','maturity/renewal','stability inference','regulatory thresholds']}}


def export_structure(result,output,readable=True):
    if result.get('schema_version')!='deposit_structure_v1': raise ValueError('Unknown deposit structure schema')
    report=Report(); scope=result['policy']['scope']; config=result['requested_query']
    prefix=[scope[key] for key in ('currency','unit','version','level')]+[config['date']]
    context=['Currency','Unit','Version','Level','Date']
    structure=[context+['Axis','Category','Balance','Account count','Share','Share rounded','Share status']]
    for row in result['structure']:
        structure.append(prefix+[row['axis'],row['category'],report.number(row['balance'],'structure balance'),row['account_count'],report.number(row['share']['value'],'share'),row['share']['rounded'],row['share']['status']])
    concentration=[context+['Grain','Requested N','Effective N','Population','Balance','Share','Share rounded','Share status']]
    for row in result['concentration']:
        concentration.append(prefix+[row['grain'],row['requested_top_n'],row['effective_n'],row['population'],report.number(row['balance'],'Top N balance'),report.number(row['share']['value'],'share'),row['share']['rounded'],row['share']['status']])
    entities=[['Declared grain','Stable entity','Rank position','Balance','Account count']]+[[result['policy']['grain'],row['entity'],index,report.number(row['balance'],'entity balance'),row['account_count']] for index,row in enumerate(result['ranked_entities'],1)]
    lineage=[['Account','Customer','File','Sheet','Source row']]+[[row[key] for key in ('account','customer','file','sheet','source_row')] for row in result['lineage']]
    snapshot_json=json.dumps(result,ensure_ascii=False,allow_nan=False,default=str)
    if len(snapshot_json.encode('utf-8'))>MAX_INPUT_BYTES: raise ValueError('Structure snapshot exceeds 16 MiB')
    metadata=[['Section','Part','JSON/text']]+[['structure_json',index+1,snapshot_json[start:start+30000]] for index,start in enumerate(range(0,len(snapshot_json),30000))]
    sheets=[report.sheet(name,rows) for name,rows in [('Structure',structure),('Concentration',concentration),('Entities',entities),('Lineage',lineage),('Metadata',metadata)]]
    sheets.append(report.sheet('Warnings',[['Location','Issue','Original']]+[['quality',flag,None] for flag in result['quality_flags']]+report.warnings))
    appearance=None if not readable else {'preset':'readable','sheets':{sheet['name']:{'header_rows':'1','column_widths':{'A':24}} for sheet in sheets}}
    book=Workbook.create({'sheets':sheets},appearance=appearance); book.save(output)
    return {'status':'exported','output_written':True,'output':str(output),'formulas_created':0,'total_cells':report.count}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input'); parser.add_argument('--query',required=True); parser.add_argument('--xlsx'); parser.add_argument('--plain',action='store_true')
    args=parser.parse_args()
    try:
        result=analyze(args.input,read_spec(args.query,max_bytes=65536))
        if args.xlsx: result['export']=export_structure(result,args.xlsx,not args.plain)
        print(json.dumps(result,ensure_ascii=False,default=str)); return 0
    except (ValueError,OSError,LookupError,DecimalException,ET.ParseError,zipfile.BadZipFile) as exc:
        print(json.dumps({'status':'failed','message':str(exc),'output_written':False},ensure_ascii=False)); return 1


if __name__=='__main__': sys.exit(main())
