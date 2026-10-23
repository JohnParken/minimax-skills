#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Observed deposit snapshot changes; no inferred openings, closures or customer intent."""
import argparse
from decimal import Decimal,DecimalException
import json
from pathlib import Path
import sys
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook,read_spec
from xlsx_analysis import region
from xlsx_headers import resolve_column
from xlsx_daily import calendar_date,summarize as daily_summary
from xlsx_advanced import exact,rounded,precision_option
from xlsx_result_export import Report,MAX_INPUT_BYTES

FIELDS={'entity','institution','product','balance','currency','unit','version','date','level'}


def snapshot(source,spec,day,scope):
    if not isinstance(spec,dict) or set(spec)-{'sheet','columns','range','header_rows'} or not isinstance(spec.get('sheet'),str):
        raise ValueError('Snapshot requires sheet/columns and optional range/header_rows')
    if not isinstance(spec.get('columns'),dict) or set(spec['columns'])!=FIELDS:
        raise ValueError('Snapshot columns require entity/institution/product/balance/currency/unit/version/date/level')
    columns,_,rows,metadata=region(source,spec['sheet'],cell_range=spec.get('range'),header_rows=spec.get('header_rows'))
    mapping={name:resolve_column(selector,columns,metadata.get('header_schema')) for name,selector in spec['columns'].items()}
    if len(set(mapping.values()))!=len(mapping): raise ValueError('Snapshot fields must use distinct columns')
    output={}
    for number,values in rows:
        record={name:values[columns.index(column)] for name,column in mapping.items()}
        for name in FIELDS-{'balance'}:
            if record[name][0]!='text' or not record[name][1].strip():
                raise ValueError('Snapshot identifiers/metadata must be explicit nonempty text at row '+str(number))
        if record['date']!=('text',day): raise ValueError('Snapshot date differs from explicit side date')
        if any(record[key]!=('text',value) for key,value in scope.items()):
            raise ValueError('Snapshot currency/unit/version/level differs from declared scope')
        value=record['balance']
        if value[0]!='number' or value[1]<0: raise ValueError('Deposit balances must be nonnegative numeric values, no text coercion')
        key=record['entity'][1]
        if key in output: raise ValueError('Duplicate stable entity key; preaggregate explicitly at declared grain')
        output[key]={'entity':key,'institution':record['institution'][1],'product':record['product'][1],
                     'balance':value[1],'source_row':number}
        if len(output)>5000: raise ValueError('At most 5000 snapshot entities per side')
    if not output: raise ValueError('Snapshot contains no observations')
    return output,{'file':str(source),'sheet':spec['sheet'],'region':metadata,'columns':mapping}


def analyze(left_source,right_source,config,daily_source=None):
    required={'grain','left','right','left_date','right_date','scope'}
    if not isinstance(config,dict) or not required<=set(config) or set(config)-required-{'daily','targets','precision'}:
        raise ValueError('Unknown deposit config; require grain, both snapshot specs/dates and scope')
    if config['grain'] not in ('account','customer'): raise ValueError('Declare account or customer grain using one stable anonymized key')
    start=calendar_date(config['left_date'],'left_date'); end=calendar_date(config['right_date'],'right_date')
    if start>=end: raise ValueError('right_date must be after left_date')
    scope=config['scope']
    if not isinstance(scope,dict) or set(scope)!={'currency','unit','version','level'} or any(not isinstance(value,str) or not value.strip() for value in scope.values()):
        raise ValueError('Require one explicit currency/unit/version/institution level scope')
    precision=precision_option(config)
    left,left_meta=snapshot(left_source,config['left'],config['left_date'],scope)
    right,right_meta=snapshot(right_source,config['right'],config['right_date'],scope)
    daily=None; averages={}
    if (daily_source is None)!=('daily' not in config): raise ValueError('Supply both daily source and daily configuration')
    if daily_source is not None:
        spec=config['daily']
        if not isinstance(spec,dict) or set(spec)!={'sheet','query','metric_id','grain'} or spec['grain']!=config['grain']:
            raise ValueError('Daily input requires sheet/query/metric_id and matching grain')
        query=spec['query']
        if not isinstance(query,dict) or query.get('coverage','complete')!='complete' or query.get('end')!=config['right_date'] or query.get('definition_version')!=scope['version'] or not isinstance(query.get('entity_keys'),list) or len(query['entity_keys'])!=1:
            raise ValueError('Daily query must use one stable entity key, complete coverage, matching version and right_date end')
        definitions=query.get('metrics')
        if not isinstance(definitions,list) or len(definitions)!=1 or not isinstance(definitions[0],dict) or definitions[0].get('metric_id')!=spec['metric_id'] or definitions[0].get('unit')!=scope['unit']:
            raise ValueError('Daily query must declare exactly the selected deposit metric with matching unit')
        daily=daily_summary(daily_source,spec['sheet'],query,minimum_value=Decimal(0))
        identities=[tuple(json.loads(value)) for value in daily['columns']]
        def value(row,*identity): return row['cells'][identities.index(tuple(identity))]['value']
        for row in daily['rows']:
            if value(row,'dimension','source_metric')!=spec['metric_id']:
                raise ValueError('Daily source must contain only the selected deposit metric')
            if any(value(row,'dimension',key)!=scope[key] for key in ('currency','unit','version')):
                raise ValueError('Daily currency/unit/version is incompatible')
            entity_roles=[role for role in identities if role[0]=='dimension' and role[1].startswith('entity:')]
            entity=value(row,*entity_roles[0])
            if entity in averages: raise ValueError('Duplicate daily entity interval')
            if entity not in left and entity not in right: raise ValueError('Daily entity absent from both snapshots; universe differs')
            averages[entity]={'value':value(row,'metric','daily_average','actual'),
                              'observed':value(row,'metric','observed_daily_average','actual'),
                              'quality_flags':row['quality_flags'],'daily_row_id':row['row_id']}
    targets=config.get('targets',[]); target_index={}
    if not isinstance(targets,list) or len(targets)>10000: raise ValueError('At most 10000 explicit targets')
    periods={'balance':config['right_date'],'net_change':config['left_date']+'..'+config['right_date']}
    if daily is not None: periods['daily_average']=config['daily']['query']['start']+'..'+config['daily']['query']['end']
    for target in targets:
        if not isinstance(target,dict) or set(target)!={'entity','metric','period','unit','currency','version','level','grain','value'}:
            raise ValueError('Target requires entity/metric/period/unit/currency/version/level/grain/value')
        if not isinstance(target['entity'],str) or target['entity'] not in left and target['entity'] not in right:
            raise ValueError('Target entity absent from snapshot universe')
        metric=target['metric']
        if not isinstance(metric,str) or metric not in periods or target['period']!=periods[metric] or target['grain']!=config['grain'] or any(target[key]!=value for key,value in scope.items()):
            raise ValueError('Target metric/period/grain/unit/currency/version/level incompatible; targets cannot be reused across measures')
        if isinstance(target['value'],bool) or not isinstance(target['value'],(int,Decimal)) or not Decimal(target['value']).is_finite():
            raise ValueError('Target must be a finite JSON numeric value')
        key=(target['entity'],metric)
        if key in target_index: raise ValueError('Duplicate entity/metric target')
        target_index[key]=Decimal(target['value'])
    records=[]; components={key:[] for key in ('continuing_increase','continuing_decrease','only_current_observed','only_previous_observed')}
    lineage=[]
    for entity in dict.fromkeys(list(left)+list(right)):
        before=left.get(entity); after=right.get(entity)
        old=None if before is None else before['balance']; new=None if after is None else after['balance']
        delta=exact(lambda:(Decimal(0) if new is None else new)-(Decimal(0) if old is None else old))
        category='only_current_observed' if before is None else 'only_previous_observed' if after is None else 'continuing_increase' if delta>0 else 'continuing_decrease' if delta<0 else 'continuing_unchanged'
        if category in components: components[category].append(delta)
        moved=before is not None and after is not None and before['institution']!=after['institution']
        changed_product=before is not None and after is not None and before['product']!=after['product']
        flags=[]
        if moved: flags.append('institution_transfer_not_bank_entry_exit')
        if changed_product: flags.append('product_changed_not_bank_entry_exit')
        if before is None or after is None: flags.append('observation_absence_not_lifecycle_evidence')
        average=averages.get(entity)
        mean=None if average is None else average['value']
        observed=None if average is None else average['observed']
        if daily is not None:
            flags+=['entity_daily_mean_not_attributed_to_institution_or_product']
            flags+=['daily_entity_not_observed'] if average is None else average['quality_flags']
        gap_mean=None if new is None or mean is None else exact(lambda:new-mean)
        record={'entity':entity,'before_institution':None if before is None else before['institution'],
                'after_institution':None if after is None else after['institution'],
                'before_product':None if before is None else before['product'],'after_product':None if after is None else after['product'],
                'before':old,'balance':new,'net_change':delta,'component':category,'institution_changed':moved,'product_changed':changed_product,
                'daily_average':mean,'observed_daily_average':observed,'end_minus_daily_average':gap_mean,'quality_flags':flags,'targets':{}}
        for metric in periods:
            target=target_index.get((entity,metric)); actual=record[metric]
            achievement=None; gap=None; status='not_configured'
            if target is not None:
                status='missing_actual' if actual is None else 'nonpositive_target' if target<=0 else 'evaluated_at_least'
                if actual is not None:
                    gap=exact(lambda:max(target-actual,Decimal(0)))
                    if target>0:
                        ratio=rounded(lambda:actual/target,precision); achievement=ratio['value']
                        if ratio['rounded']: flags.append('rounded_achievement:'+metric)
            record['targets'][metric]={'target':target,'achievement':achievement,'gap':gap,'status':status}
        records.append(record)
        for side,item,metadata in [('left',before,left_meta),('right',after,right_meta)]:
            if item is not None: lineage.append({'entity':entity,'side':side,'file':metadata['file'],'sheet':metadata['sheet'],'source_row':item['source_row']})
        if average is not None: lineage.append({'entity':entity,'side':'daily','daily_row_id':average['daily_row_id']})
        if len(records)>10000: raise ValueError('At most 10000 union entities')
    old_total=exact(lambda:sum((row['balance'] for row in left.values()),Decimal(0)))
    new_total=exact(lambda:sum((row['balance'] for row in right.values()),Decimal(0)))
    delta_total=exact(lambda:new_total-old_total)
    parts={key:exact(lambda:sum(values,Decimal(0))) for key,values in components.items()}
    residual=exact(lambda:delta_total-sum(parts.values(),Decimal(0)))
    if residual!=0: raise ValueError('Deposit component reconciliation failed')
    return {'status':'deposit_analysis','schema_version':'deposit_analysis_v1','requested_query':config,'records':records,
            'totals':{'before':old_total,'balance':new_total,'net_change':delta_total,'components':parts,'residual':residual},
            'sources':{'left':left_meta,'right':right_meta,'daily':daily},'lineage':lineage,
            'policy':{'scope':scope,'grain':config['grain'],'absence':'Not observed, never inferred opening/closure',
                      'transfers':'Stable IDs align across institution/product changes; no institution/product flow attribution',
                      'daily':'Complete inclusive coverage required; observed means separate; no zero/forward fill; no end-date manipulation inference',
                      'targets':'Entity-specific at_least only; separate metric and period; nonpositive targets never classified achieved',
                      'precision':precision,'rounding':'ROUND_HALF_EVEN','exact_precision_limit':1000}}


def export_deposits(result,output,readable=True):
    if result.get('schema_version')!='deposit_analysis_v1': raise ValueError('Unknown deposit schema')
    report=Report()
    text=['entity','before_institution','after_institution','before_product','after_product','component']
    numbers=['before','balance','net_change','daily_average','observed_daily_average','end_minus_daily_average']
    metrics=['balance','net_change','daily_average']
    headers=text+numbers+['currency','unit','version','level','grain','left_date','right_date']
    headers += [metric+'/'+field for metric in metrics for field in ('target','achievement','gap','status')]+['quality_flags']
    data=[headers]
    for record in result['records']:
        scope=result['policy']['scope']; config=result['requested_query']
        values=[record[key] for key in text]+[report.number(record[key],key) for key in numbers]
        values += [scope[key] for key in ('currency','unit','version','level')]+[config['grain'],config['left_date'],config['right_date']]
        for metric in metrics:
            target=record['targets'].get(metric,{'target':None,'achievement':None,'gap':None,'status':'not_configured'})
            values += [report.number(target[key],metric+'/'+key) for key in ('target','achievement','gap')]+[target['status']]
        data.append(values+[json.dumps(record['quality_flags'],ensure_ascii=False)])
    checks=[['Measure','Value']]
    checks += [[key,report.number(result['totals'][key],key)] for key in ('before','balance','net_change','residual')]
    checks += [[key,report.number(value,key)] for key,value in result['totals']['components'].items()]
    lineage=[['Entity','Side','File','Sheet','Source row','Daily row ID']]+[[row['entity'],row['side'],row.get('file'),row.get('sheet'),row.get('source_row'),row.get('daily_row_id')] for row in result['lineage']]
    quality=[['Entity','Quality flag']]+[[record['entity'],flag] for record in result['records'] for flag in record['quality_flags']]
    snapshot_json=json.dumps(result,ensure_ascii=False,allow_nan=False,default=str)
    if len(snapshot_json.encode('utf-8'))>MAX_INPUT_BYTES: raise ValueError('Deposit report snapshot exceeds 16 MiB')
    metadata=[['Section','Part','JSON/text']]+[['deposit_json',index+1,snapshot_json[start:start+30000]] for index,start in enumerate(range(0,len(snapshot_json),30000))]
    sheets=[report.sheet(name,rows) for name,rows in [('Deposits',data),('Checks',checks),('Quality',quality),('Lineage',lineage),('Metadata',metadata)]]
    sheets.append(report.sheet('Warnings',[['Location','Issue','Original']]+report.warnings))
    appearance=None if not readable else {'preset':'readable','sheets':{sheet['name']:{'header_rows':'1','column_widths':{'A':24}} for sheet in sheets}}
    if appearance: appearance['sheets']['Deposits']['category_columns']=['F']
    book=Workbook.create({'sheets':sheets},appearance=appearance); book.save(output)
    return {'status':'exported','output_written':True,'output':str(output),'formulas_created':0,'total_cells':report.count}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('left'); parser.add_argument('right'); parser.add_argument('--query',required=True)
    parser.add_argument('--daily'); parser.add_argument('--xlsx'); parser.add_argument('--plain',action='store_true')
    args=parser.parse_args()
    try:
        result=analyze(args.left,args.right,read_spec(args.query,max_bytes=65536),args.daily)
        if args.xlsx: result['export']=export_deposits(result,args.xlsx,not args.plain)
        print(json.dumps(result,ensure_ascii=False,default=str)); return 0
    except (ValueError,OSError,LookupError,DecimalException,ET.ParseError,zipfile.BadZipFile) as exc:
        print(json.dumps({'status':'failed','message':str(exc),'output_written':False},ensure_ascii=False)); return 1


if __name__=='__main__': sys.exit(main())
