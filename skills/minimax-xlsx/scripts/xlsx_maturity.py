#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Explicit deposit maturity partitions and overlapping forward windows, no renewal inference."""
import argparse
from decimal import Decimal,DecimalException
import json
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import read_spec
from xlsx_analysis import region,is_blank
from xlsx_headers import resolve_column
from xlsx_daily import calendar_date
from xlsx_advanced import exact,rounded,precision_option
from xlsx_management import column_id
from xlsx_result_export import export_result

FIELDS={'contract','principal','maturity_date','observation_date','institution','currency','unit','version','level','status'}


def metric_id(*parts):
    return json.dumps(list(parts),ensure_ascii=False,separators=(',',':'))


def analyze(source,config):
    required={'sheet','columns','observation_date','scope','institutions','status_mapping','buckets','cumulative_windows','share_denominator'}
    if not isinstance(config,dict) or not required<=set(config) or set(config)-required-{'range','header_rows','precision'}:
        raise ValueError('Require explicit maturity input columns/date/scope/institutions/status mapping/buckets/windows/denominator')
    day=calendar_date(config['observation_date'],'observation_date'); precision=precision_option(config)
    scope=config['scope']
    if not isinstance(scope,dict) or set(scope)!={'currency','unit','version','definition_version','level','completeness'} or any(not isinstance(value,str) or not value.strip() for value in scope.values()) or scope['completeness'] not in ('declared_complete','unknown'):
        raise ValueError('Declare one currency/unit/version/definition_version/level and completeness')
    if scope['level']=='bank_scope': raise ValueError('bank_scope is reserved for separate all-institution output')
    institutions=config['institutions']
    if not isinstance(institutions,list) or not 1<=len(institutions)<=100 or any(not isinstance(value,str) or not value.strip() for value in institutions) or len(set(institutions))!=len(institutions):
        raise ValueError('Declare 1..100 unique institutions at one level; parent/leaf mixed input is unsupported')
    status_map=config['status_mapping']
    if not isinstance(status_map,dict) or not status_map or any(not isinstance(key,str) or not key or value not in ('include','exclude') for key,value in status_map.items()):
        raise ValueError('Map every source status explicitly to include/exclude')
    if config['share_denominator']!='eligible_total_principal': raise ValueError('Shares use all included principal, including past/today/unknown dates')
    buckets=config['buckets']
    if not isinstance(buckets,list) or not 1<=len(buckets)<=20: raise ValueError('Declare 1..20 disjoint future day buckets')
    next_day=1; ids={'past_date','today','unknown_date'}
    for index,bucket in enumerate(buckets):
        if not isinstance(bucket,dict) or set(bucket)!={'id','min_days','max_days'} or not isinstance(bucket['id'],str) or not bucket['id'] or bucket['id'] in ids:
            raise ValueError('Buckets need unique nonreserved id/min_days/max_days')
        ids.add(bucket['id']); lo=bucket['min_days']; hi=bucket['max_days']
        if type(lo) is not int or lo!=next_day or hi is not None and (type(hi) is not int or not lo<=hi<=36500):
            raise ValueError('Future buckets must be contiguous inclusive nonoverlapping intervals from day 1')
        if hi is None and index!=len(buckets)-1 or hi is not None and index==len(buckets)-1:
            raise ValueError('Final future bucket must be unbounded (max_days null), and only final may be unbounded')
        next_day=None if hi is None else hi+1
    windows=config['cumulative_windows']
    if not isinstance(windows,list) or len(windows)>20 or any(type(value) is not int or not 1<=value<=36500 for value in windows) or len(set(windows))!=len(windows):
        raise ValueError('Cumulative windows must be unique positive days <=36500')
    selectors=config['columns']
    if not isinstance(selectors,dict) or set(selectors)!=FIELDS: raise ValueError('Require contract/principal/maturity/observation/institution/currency/unit/version/level/status columns')
    columns,_,rows,metadata=region(source,config['sheet'],cell_range=config.get('range'),header_rows=config.get('header_rows'))
    mapping={name:resolve_column(selector,columns,metadata.get('header_schema')) for name,selector in selectors.items()}
    if len(set(mapping.values()))!=len(mapping): raise ValueError('Maturity fields must be distinct columns')
    included=[]; excluded=[]; seen=set()
    for number,values in rows:
        fields={name:values[columns.index(column)] for name,column in mapping.items()}
        for name in FIELDS-{'principal','maturity_date'}:
            if fields[name][0]!='text' or not fields[name][1].strip(): raise ValueError('Maturity identifiers/metadata must be explicit nonempty text at row '+str(number))
        contract=fields['contract'][1]
        if contract in seen: raise ValueError('Duplicate contract observation; resolve snapshot grain explicitly')
        seen.add(contract)
        if len(seen)>10000: raise ValueError('At most 10000 contract observations')
        if fields['observation_date']!=('text',config['observation_date']) or any(fields[key]!=('text',scope[key]) for key in ('currency','unit','version','level')):
            raise ValueError('Observation date/currency/unit/version/level differs from scope')
        institution=fields['institution'][1]
        if institution not in institutions: raise ValueError('Unknown scoped institution')
        principal=fields['principal']
        if principal[0]!='number' or principal[1]<0: raise ValueError('Principal must be nonnegative numeric; blank/text/negative principal is unsupported')
        maturity=fields['maturity_date']
        if is_blank(maturity): distance=None
        elif maturity[0] not in ('text','date_text'): raise ValueError('Maturity date must be ISO text, never inferred serials')
        else: distance=(calendar_date(maturity[1],'maturity row '+str(number))-day).days
        status=fields['status'][1]
        if status not in status_map: raise ValueError('Unmapped contract status')
        if status_map[status]=='exclude':
            excluded.append({'source_row':number,'status':status}); continue
        category='unknown_date' if distance is None else 'past_date' if distance<0 else 'today' if distance==0 else next(bucket['id'] for bucket in buckets if bucket['min_days']<=distance and (bucket['max_days'] is None or distance<=bucket['max_days']))
        included.append({'source_row':number,'institution':institution,'principal':principal[1],'distance':distance,'bucket':category})
    definitions=[]; categories=['past_date','today','unknown_date']+[bucket['id'] for bucket in buckets]
    def add(identifier,label,unit,kind,aggregation,**extras):
        definitions.append({'metric_id':identifier,'label':label,'unit':unit,'kind':kind,'time_basis':'ratio' if kind=='ratio' else 'point_in_time','aggregation':aggregation,**extras})
    add('eligible_principal','Included principal',scope['unit'],'stock','SUM',additive_dimensions=['institution'])
    add('eligible_contracts','Included contract observations','contracts','count','COUNT',additive_dimensions=['institution'])
    for role,labels in [('bucket',categories),('window',[str(value) for value in windows])]:
        for label in labels:
            principal_id=metric_id(role,label,'principal'); count_id=metric_id(role,label,'contracts'); share_id=metric_id(role,label,'share')
            add(principal_id,role+' '+label+' principal',scope['unit'],'stock','SUM',additive_dimensions=['institution'])
            add(count_id,role+' '+label+' contract observations','contracts','count','COUNT',additive_dimensions=['institution'])
            add(share_id,role+' '+label+' principal share','fraction','ratio','RATIO',numerator=principal_id,denominator='eligible_principal',scale=1)
    roles=[('dimension',name) for name in ('period','observation_date','institution','report_level','currency','version')]+[('metric',item['metric_id'],'actual') for item in definitions]+[('quality','flags')]
    identifiers=[column_id(*role) for role in roles]
    labels=[role[1] if role[0]!='metric' else next(item['label'] for item in definitions if item['metric_id']==role[1]) for role in roles]
    generated=[{'id':identifier,'column':identifier,'path':[{'type':'text','value':label}],'levels':[{'type':'text','value':label,'role':'maturity_distribution'}]} for identifier,label in zip(identifiers,labels)]
    output=[]; lineage=[]; checks=[]; details=[]
    for kind,institution in [('bank_scope','all')]+[(scope['level'],value) for value in institutions]:
        records=included if kind=='bank_scope' else [row for row in included if row['institution']==institution]
        total=exact(lambda:sum((row['principal'] for row in records),Decimal(0))); flags=[]
        if not records: flags.append('no_eligible_observations')
        if scope['completeness']=='unknown': flags.append('observation_completeness_unknown')
        if any(row['distance'] is None for row in records): flags.append('unknown_maturity_dates')
        if total==0: flags.append('zero_principal_denominator')
        actuals={'eligible_principal':total,'eligible_contracts':len(records)}; partition_sum=Decimal(0)
        for role,names in [('bucket',categories),('window',[str(value) for value in windows])]:
            for label in names:
                selected=[row for row in records if row['bucket']==label] if role=='bucket' else [row for row in records if row['distance'] is not None and 1<=row['distance']<=int(label)]
                value=exact(lambda:sum((row['principal'] for row in selected),Decimal(0)))
                if role=='bucket': partition_sum=exact(lambda:partition_sum+value)
                share=None if total==0 else rounded(lambda:value/total,precision)
                actuals[metric_id(role,label,'principal')]=value; actuals[metric_id(role,label,'contracts')]=len(selected)
                actuals[metric_id(role,label,'share')]=None if share is None else share['value']
                details.append({'scope_kind':kind,'scope_id':institution,'role':role,'id':label,'principal':value,'contracts':len(selected),'denominator':total,'share_rounded':False if share is None else share['rounded']})
        residual=exact(lambda:total-partition_sum)
        if residual!=0: raise ValueError('Disjoint maturity partitions failed reconciliation')
        checks.append({'check':'partition_principal','status':'pass','detail':kind+'/'+institution+': residual '+str(residual)})
        values=[day.isoformat()[:7],day.isoformat(),institution,kind,scope['currency'],scope['version']]
        cells=[{'type':'text','value':value} for value in values]+[{'type':'blank' if actuals[item['metric_id']] is None else 'number','value':actuals[item['metric_id']]} for item in definitions]
        cells.append({'type':'text','value':json.dumps(flags,ensure_ascii=False)})
        row_id=len(output)+1
        output.append({'row_id':row_id,'cells':cells,'source_rows':[row['source_row'] for row in records],'quality_flags':flags})
        lineage += [{'output_row':row_id,'source_file':str(source),'sheet':config['sheet'],'source_row':row['source_row'],'metric_id':'eligible_principal'} for row in records]
    return {'status':'management_wide','schema_version':'management_wide_v1','analysis':'deposit_maturity','sheet':config['sheet'],'range':metadata['range'],
            'requested_query':config,'source_file':str(source),'source_metadata':metadata,'columns':identifiers,'headers':[{'type':'text','value':label} for label in labels],
            'generated_columns':generated,'rows':output,'output_rows':len(output),'metric_dictionary':definitions,'definition_version':scope['definition_version'],
            'report_level':'separate bank_scope and same-level institution slices','lineage':lineage,'details':details,'excluded_status_rows':excluded,
            'validation':checks+[{'check':'ratios_recomputed','status':'pass','detail':'Shares = scoped partition/window principal divided by all included scoped principal'},
                                 {'check':'additivity','status':'warning','detail':'Do not add bank_scope to institution slices or sum overlapping cumulative windows'}],
            'policies':{'days':'Gregorian calendar days = maturity date - explicit observation date; day 0 today; negative days past date; null unknown',
                        'buckets':'Disjoint inclusive future intervals start at day 1 and cover all future days; past/today/unknown separate',
                        'windows':'Overlapping future cumulative intervals [1,N]; exclude day 0, past and unknown; never add windows',
                        'status':'Mapped include/exclude only; duplicate IDs, malformed dates and negative amounts fail even in excluded records',
                        'interpretation':'Past maturity in snapshot is not proof of nonredemption/overdue; future maturity is not loss; no renewal calculation',
                        'hierarchy':'One configured institution level only; no parent/child input or automatic rollup; bank_scope separately aggregated',
                        'precision':precision,'rounding':'ROUND_HALF_EVEN','exact_precision_limit':1000}}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input'); parser.add_argument('--query',required=True); parser.add_argument('--xlsx'); parser.add_argument('--plain',action='store_true')
    args=parser.parse_args()
    try:
        result=analyze(args.input,read_spec(args.query,max_bytes=65536))
        if args.xlsx: result['export']=export_result(result,args.xlsx,not args.plain)
        print(json.dumps(result,ensure_ascii=False,default=str)); return 0
    except (ValueError,OSError,LookupError,DecimalException,ET.ParseError,zipfile.BadZipFile) as exc:
        print(json.dumps({'status':'failed','message':str(exc),'output_written':False},ensure_ascii=False)); return 1


if __name__=='__main__': sys.exit(main())
