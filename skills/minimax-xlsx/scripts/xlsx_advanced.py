#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Explicit Decimal descriptive statistics and calendar-month comparisons."""
import argparse
from decimal import Decimal, DecimalException, Inexact, ROUND_FLOOR, ROUND_HALF_EVEN, localcontext
import json
from pathlib import Path
import re
import sys
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0,str(Path(__file__).resolve().parent))
from xlsx_analysis import region, is_blank
from xlsx_headers import resolve_column
from xlsx_rows import read_query


def precision_option(config):
    precision=config.get('precision',34)
    if type(precision) is not int or not 16<=precision<=100:
        raise ValueError('precision must be an integer from 16 to 100 significant digits')
    return precision


def rounded(operation,precision):
    with localcontext() as context:
        context.prec=precision+16
        context.rounding=ROUND_HALF_EVEN
        context.traps[Inexact]=False
        context.clear_flags()
        value=operation()
        intermediate=context.flags[Inexact]
        context.prec=precision
        context.clear_flags()
        value=+value
        return {'value':value,'rounded':intermediate or context.flags[Inexact]}


def exact(operation):
    with localcontext() as context:
        context.prec=1000
        context.traps[Inexact]=True
        try:
            return operation()
        except DecimalException as exc:
            raise ValueError('Exact arithmetic exceeds the supported 1000-digit limit') from exc


def resolve_many(selectors,columns,schema):
    if not isinstance(selectors,list) or not selectors:
        raise ValueError('Provide a nonempty list of column selectors')
    result=[resolve_column(selector,columns,schema) for selector in selectors]
    if len(result)!=len(set(result)):
        raise ValueError('Column selectors must resolve to unique columns')
    return result


def quantile(numbers,probability,precision):
    # Hyndman-Fan type 7: h=(n-1)*p; interpolate sorted x[floor(h)]..x[ceil(h)].
    def interpolate():
        rank=(len(numbers)-1)*probability
        lower=int(rank.to_integral_value(rounding=ROUND_FLOOR))
        if lower==len(numbers)-1:
            return numbers[-1]
        fraction=rank-lower
        return numbers[lower]+fraction*(numbers[lower+1]-numbers[lower])
    value=exact(interpolate)
    return rounded(lambda:value,precision)


def describe(source,sheet,config,**options):
    if not isinstance(config,dict) or set(config)-{'columns','missing','non_numeric','quantiles','variance','precision'}:
        raise ValueError('Unknown descriptive-statistics configuration keys')
    missing=config.get('missing','omit')
    non_numeric=config.get('non_numeric','reject')
    variance=config.get('variance','both')
    precision=precision_option(config)
    if missing not in ('omit','reject') or non_numeric not in ('omit','reject') or variance not in ('sample','population','both'):
        raise ValueError('Invalid missing/non_numeric/variance policy')
    probabilities=config.get('quantiles',[Decimal('0.25'),Decimal('0.5'),Decimal('0.75')])
    if not isinstance(probabilities,list) or len(probabilities)>21:
        raise ValueError('quantiles must be a list with at most 21 probabilities')
    if any(isinstance(value,bool) or not isinstance(value,(int,Decimal)) or not Decimal(value).is_finite() or not 0<=value<=1 for value in probabilities):
        raise ValueError('Quantile probabilities must be JSON numbers between 0 and 1')
    probabilities=[Decimal(value) for value in probabilities]
    if len(probabilities)!=len(set(probabilities)):
        raise ValueError('Duplicate quantile probabilities')
    columns,_,rows,metadata=region(source,sheet,**options)
    metadata.pop('precision',None)
    chosen=resolve_many(config.get('columns'),columns,metadata.get('header_schema'))
    if len(chosen)*(len(probabilities)+6)>20000:
        raise ValueError('Descriptive result exceeds 20000 metric values; reduce columns/quantiles')
    output=[]
    for column in chosen:
        index=columns.index(column)
        values=[]
        missing_count=omitted_count=0
        for number,row in rows:
            item=row[index]
            if is_blank(item):
                if missing=='reject':
                    raise ValueError('Missing value at '+column+str(number))
                missing_count+=1
            elif item[0]!='number':
                if non_numeric=='reject':
                    raise ValueError('Non-numeric value at '+column+str(number))
                omitted_count+=1
            else:
                values.append(item[1])
        values.sort()
        count=len(values)
        result={'column':column,'column_id':'col:'+column,'count':count,'missing_omitted':missing_count,
                'non_numeric_omitted':omitted_count,'status':'ok' if count else 'insufficient_numeric_samples',
                'mean':None,'median':None,'quantiles':[],'variance':{}}
        if count:
            total=exact(lambda:sum(values,Decimal(0)))
            result['mean']=rounded(lambda:total/count,precision)
            result['median']=quantile(values,Decimal('.5'),precision)
            result['quantiles']=[{'probability':probability,**quantile(values,probability,precision)} for probability in probabilities]
            numerator=exact(lambda:count*sum((value*value for value in values),Decimal(0))-total*total)
            modes=('sample','population') if variance=='both' else (variance,)
            for mode in modes:
                enough=count>=2 if mode=='sample' else count>=1
                denominator=count*(count-1) if mode=='sample' else count*count
                result['variance'][mode]={'status':'ok' if enough else 'insufficient_numeric_samples',
                    'ddof':1 if mode=='sample' else 0,
                    'variance':rounded(lambda:numerator/denominator,precision) if enough else None,
                    'standard_deviation':rounded(lambda:(numerator/denominator).sqrt(),precision) if enough else None}
        else:
            for mode in (('sample','population') if variance=='both' else (variance,)):
                result['variance'][mode]={'status':'insufficient_numeric_samples','ddof':1 if mode=='sample' else 0,
                                          'variance':None,'standard_deviation':None}
        output.append(result)
    return {'status':'described',**metadata,'columns':output,'missing':missing,'non_numeric':non_numeric,
            'quantile_algorithm':'Hyndman-Fan type 7; h=(n-1)*p, linear interpolation',
            'arithmetic':{'exact_precision_limit':1000,'result_precision':precision,
                          'intermediate_precision':precision+16,'rounding':'ROUND_HALF_EVEN',
                          'rounded_flag':'true if intermediate or final result was inexact'}}


def parse_month(item,coordinate):
    if item[0]!='text' or not re.fullmatch(r'[0-9]{4}-(?:0[1-9]|1[0-2])',item[1]):
        raise ValueError('Period must be literal YYYY-MM text at '+coordinate)
    year,month=map(int,item[1].split('-'))
    if year<1:
        raise ValueError('Period year must be 0001..9999 at '+coordinate)
    return (year-1)*12+month-1


def month_text(number):
    return None if number<0 else '%04d-%02d' % (number//12+1,number%12+1)


def compare(source,sheet,config,**options):
    if not isinstance(config,dict) or set(config)-{'entity_keys','period','metrics','period_format','comparisons','non_numeric','precision'}:
        raise ValueError('Unknown period-comparison configuration keys')
    if config.get('period_format')!='YYYY-MM':
        raise ValueError('An explicit period_format: YYYY-MM is required')
    kinds=config.get('comparisons',['mom','yoy'])
    if not isinstance(kinds,list) or not kinds or any(kind not in ('mom','yoy') for kind in kinds) or len(kinds)!=len(set(kinds)):
        raise ValueError('comparisons must contain unique mom/yoy entries')
    non_numeric=config.get('non_numeric','reject')
    if non_numeric not in ('reject','diagnose'):
        raise ValueError('non_numeric must be reject or diagnose')
    precision=precision_option(config)
    columns,_,rows,metadata=region(source,sheet,**options)
    metadata.pop('precision',None)
    schema=metadata.get('header_schema')
    entities=config.get('entity_keys')
    if not isinstance(entities,list):
        raise ValueError('Explicit entity_keys list required; [] means one entity')
    entities=resolve_many(entities,columns,schema) if entities else []
    period=resolve_column(config.get('period'),columns,schema)
    metrics=resolve_many(config.get('metrics'),columns,schema)
    if len(rows)*len(metrics)*len(kinds)>20000:
        raise ValueError('Comparison exceeds 20000 metric results; narrow the region/metrics')
    entity_indices=[columns.index(col) for col in entities]
    period_index=columns.index(period)
    metric_indices=[columns.index(col) for col in metrics]
    indexed={}
    records=[]
    for row_number,row in rows:
        key=tuple(row[index] for index in entity_indices)
        if any(is_blank(item) for item in key):
            raise ValueError('Missing entity key at source row '+str(row_number))
        month=parse_month(row[period_index],period+str(row_number))
        marker=(key,month)
        if marker in indexed:
            raise ValueError('Duplicate entity/month at rows %d and %d' % (indexed[marker][0],row_number))
        for col,index in zip(metrics,metric_indices):
            if not is_blank(row[index]) and row[index][0]!='number' and non_numeric=='reject':
                raise ValueError('Non-numeric metric at '+col+str(row_number))
        indexed[marker]=(row_number,row)
        records.append((key,month,row_number,row))
    output=[]
    for key,month,row_number,row in records:
        for kind in kinds:
            base_month=month-(1 if kind=='mom' else 12)
            base=indexed.get((key,base_month))
            results=[]
            for column,index in zip(metrics,metric_indices):
                current=row[index]
                previous=None if base is None else base[1][index]
                result={'column':column,'current':current[1] if current[0]=='number' else None,
                        'base':previous[1] if previous is not None and previous[0]=='number' else None,
                        'difference':None,'growth_rate':None,'negative_base':False,'status':'ok'}
                if is_blank(current):
                    result['status']='current_missing'
                elif current[0]!='number':
                    result['status']='current_non_numeric'
                elif base is None:
                    result['status']='missing_period'
                elif is_blank(previous):
                    result['status']='base_missing'
                elif previous[0]!='number':
                    result['status']='base_non_numeric'
                else:
                    difference=exact(lambda:current[1]-previous[1])
                    result['difference']=difference
                    result['negative_base']=previous[1]<0
                    if previous[1]==0:
                        result['status']='zero_base'
                    else:
                        result['growth_rate']=rounded(lambda:difference/previous[1],precision)
                results.append(result)
            output.append({'entity':[{'column':column,'type':item[0],'value':item[1]} for column,item in zip(entities,key)],
                           'period':month_text(month),'comparison':kind,'base_period':month_text(base_month),
                           'source_row':row_number,'base_source_row':None if base is None else base[0],'metrics':results})
    return {'status':'compared',**metadata,'period_format':'YYYY-MM','entity_keys':entities,'period_column':period,
            'growth_definition':'(current-base)/base; signed denominator; a negative base can reverse intuitive improvement signs',
            'arithmetic':{'exact_difference_precision_limit':1000,'result_precision':precision,'intermediate_precision':precision+16,
                          'rounding':'ROUND_HALF_EVEN','growth_unit':'fraction, not percent'},
            'duplicate_policy':'reject','missing_period_policy':'diagnose; no previous-row substitution','rows':output}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command',choices=('describe','compare'))
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
        operation=describe if args.command=='describe' else compare
        result=operation(args.input,args.sheet,read_query(args.query),cell_range=args.cell_range,header=not args.no_header,
                         header_rows=args.header_rows,header_exclude=args.header_exclude,
                         include_spanning_headings=args.include_spanning_headings,formulas=args.formulas)
        print(json.dumps(result,ensure_ascii=False,default=str))
        return 0
    except (ValueError,OSError,LookupError,DecimalException,ET.ParseError,zipfile.BadZipFile) as exc:
        print(json.dumps({'status':'failed','message':str(exc)},ensure_ascii=False))
        return 1


if __name__=='__main__':
    sys.exit(main())
