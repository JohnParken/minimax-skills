#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Reconcile native management wide reports using explicit comparable composite keys."""
import argparse
from decimal import Decimal, DecimalException
import json
from pathlib import Path
import sys

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook, read_spec
from xlsx_dashboard import build as dashboard
from xlsx_advanced import exact, rounded, precision_option
from xlsx_result_export import Report, MAX_INPUT_BYTES


def reconcile(left,right,config):
    required={'left_period','right_period','keys','metrics','group_by','scope','tolerance'}
    if not isinstance(config,dict) or not required<=set(config) or set(config)-required-{'precision'}:
        raise ValueError('Require explicit periods, keys, metrics, group_by, scope and absolute tolerance')
    tolerance=config['tolerance']
    if isinstance(tolerance,bool) or not isinstance(tolerance,(int,Decimal)) or not Decimal(tolerance).is_finite() or tolerance<0:
        raise ValueError('Tolerance must be a nonnegative finite JSON number in metric units')
    precision=precision_option(config)
    scope=config['scope']
    if not isinstance(scope,dict) or set(scope)!={'currency','version','report_level'} or any(not isinstance(value,str) or not value for value in scope.values()):
        raise ValueError('Require explicit single currency/version/report_level scope')
    views=[dashboard(source,{'overview_period':config[name+'_period'],'metrics':config['metrics'],'trend_metrics':[]}) for name,source in [('left',left),('right',right)]]
    dimensions=views[0]['dimensions']
    if set(dimensions)!=set(views[1]['dimensions']):
        raise ValueError('Report dimensions differ; normalize explicitly before reconciliation')
    keys=config['keys']; groups=config['group_by']
    expected=set(dimensions)-{'period','institution_name'}
    if not isinstance(keys,list) or any(not isinstance(key,str) for key in keys) or len(set(keys))!=len(keys) or set(keys)!=expected:
        raise ValueError('Keys must explicitly retain every dimension except period/institution_name')
    if not isinstance(groups,list) or any(not isinstance(key,str) or key not in keys for key in groups) or len(set(groups))!=len(groups):
        raise ValueError('group_by must be a unique subset of composite keys')
    definitions=[]
    for source in (left,right):
        definitions.append({item['metric_id']:item for item in source['metric_dictionary']})
    if left.get('definition_version')!=right.get('definition_version') or left.get('hierarchy_summary')!=right.get('hierarchy_summary'):
        raise ValueError('Definition version or static institution mapping differs; no automatic restatement')
    for metric in config['metrics']:
        if definitions[0][metric]!=definitions[1][metric]:
            raise ValueError('Metric dictionaries differ: '+metric)
        if definitions[0][metric].get('aggregation')!='SUM' or definitions[0][metric]['kind'] not in ('stock','flow'):
            raise ValueError('Reconciliation totals require additive SUM stock/flow metrics; ratios/count/min/max unsupported')
    indexes=[]
    for view in views:
        index={}
        for record in view['overview']:
            if any(record['dimensions'][key]!=value for key,value in scope.items()):
                continue
            if record['actual'] is None or record['quality_flags']:
                raise ValueError('Missing or quality-flagged actuals cannot produce complete reconciliation')
            key=tuple(record['dimensions'][name] for name in keys)+(record['metric_id'],)
            if key in index: raise ValueError('Duplicate composite key')
            index[key]=record
        if not index: raise ValueError('No observations for explicit period and comparison scope')
        indexes.append(index)
    for metric in config['metrics']:
        selected=[record for index in indexes for record in index.values() if record['metric_id']==metric]
        varying={key for key in keys if len({record['dimensions'][key] for record in selected})>1}
        additive=definitions[0][metric].get('additive_dimensions',[])
        if not isinstance(additive,list) or any(not isinstance(key,str) for key in additive):
            raise ValueError('Invalid additive dimension declaration')
        if varying-set(additive):
            raise ValueError('Totals cross dimensions not declared additive: '+metric)
    output=[]; buckets={}; totals={}
    for key in dict.fromkeys(list(indexes[0])+list(indexes[1])):
        before=indexes[0].get(key); after=indexes[1].get(key)
        record=after or before; metric=record['metric_id']
        old=None if before is None else before['actual']; new=None if after is None else after['actual']
        delta=exact(lambda:(Decimal(0) if new is None else new)-(Decimal(0) if old is None else old))
        status='added' if before is None else 'removed' if after is None else 'unchanged_within_tolerance' if delta.copy_abs()<=tolerance else 'changed'
        dimensions={name:record['dimensions'][name] for name in keys}
        output.append({'record_id':len(output)+1,'keys':dimensions,'metric_id':metric,'unit':record['unit'],
                       'before':old,'after':new,'delta':delta,'status':status,
                       'left_wide_row':None if before is None else before['wide_row_id'],
                       'right_wide_row':None if after is None else after['wide_row_id']})
        group=(metric,tuple(dimensions[name] for name in groups))
        buckets.setdefault(group,[]).append(delta)
        totals.setdefault(metric,{'before':[],'after':[],'deltas':[]})
        if old is not None: totals[metric]['before'].append(old)
        if new is not None: totals[metric]['after'].append(new)
        totals[metric]['deltas'].append(delta)
        if len(output)>20000: raise ValueError('Reconciliation exceeds 20000 records')
    checks=[]; net={}
    for metric,values in totals.items():
        old=exact(lambda:sum(values['before'],Decimal(0))); new=exact(lambda:sum(values['after'],Decimal(0)))
        delta=exact(lambda:new-old); detail=exact(lambda:sum(values['deltas'],Decimal(0)))
        residual=exact(lambda:delta-detail)
        if residual.copy_abs()>tolerance: raise ValueError('Arithmetic reconciliation failed')
        net[metric]=delta
        checks.append({'metric_id':metric,'unit':definitions[0][metric]['unit'],'before':old,'after':new,'net_change':delta,
                       'detail_change':detail,'residual':residual,'status':'reconciled',
                       'net_status':'zero_net_change' if delta==0 else 'negative_net_change' if delta<0 else 'positive_net_change'})
    contributions=[]
    for (metric,group),values in buckets.items():
        delta=exact(lambda:sum(values,Decimal(0))); denominator=net[metric]
        share=None if denominator==0 else rounded(lambda:delta/denominator,precision)
        contributions.append({'metric_id':metric,'group':dict(zip(groups,group)),'unit':definitions[0][metric]['unit'],
                              'delta':delta,'net_change':denominator,'signed_share':None if share is None else share['value'],
                              'share_rounded':False if share is None else share['rounded'],
                              'share_status':'undefined_zero_net' if denominator==0 else 'signed_negative_net' if denominator<0 else 'signed_positive_net'})
    return {'status':'reconciled','schema_version':'management_reconciliation_v1','requested_query':config,
            'records':output,'checks':checks,'contributions':contributions,'sources':{'left':left,'right':right},
            'policies':{'absence':'Missing record stays null; zero used only to bridge added/removed arithmetic, not imputed observations',
                        'tolerance':'Absolute in each metric unit; small deltas retained in totals, never rounded away',
                        'contribution':'Signed group delta / signed net delta; may be negative or exceed 1; arithmetic allocation, not causal explanation',
                        'time':'Explicit period snapshots only; no cross-period balance or YTD accumulation',
                        'comparability':'Single explicit currency/version/level; identical dictionary and hierarchy; all remaining dimensions retained'},
            'arithmetic':{'exact_precision_limit':1000,'result_precision':precision,'rounding':'ROUND_HALF_EVEN'}}


def export_reconciliation(result,output):
    result=reconcile(result['sources']['left'],result['sources']['right'],result['requested_query'])
    report=Report(); keys=result['requested_query']['keys']; groups=result['requested_query']['group_by']
    records=[keys+['Metric','Unit','Before','After','Delta','Status','Left wide row','Right wide row']]
    for item in result['records']:
        records.append([item['keys'][key] for key in keys]+[item['metric_id'],item['unit']]+[report.number(item[key],key) for key in ('before','after','delta')]+[item['status'],item['left_wide_row'],item['right_wide_row']])
    checks=[['Metric','Unit','Before','After','Net change','Detail change','Residual','Check','Net status']]
    for item in result['checks']:
        checks.append([item['metric_id'],item['unit']]+[report.number(item[key],key) for key in ('before','after','net_change','detail_change','residual')]+[item['status'],item['net_status']])
    contributions=[groups+['Metric','Unit','Delta','Net change','Signed share','Share rounded','Share status']]
    for item in result['contributions']:
        contributions.append([item['group'][key] for key in groups]+[item['metric_id'],item['unit']]+[report.number(item[key],key) for key in ('delta','net_change','signed_share')]+[item['share_rounded'],item['share_status']])
    snapshot=json.dumps(result,ensure_ascii=False,allow_nan=False,default=str)
    if len(snapshot.encode('utf-8'))>MAX_INPUT_BYTES: raise ValueError('Reconciliation snapshot exceeds 16 MiB')
    metadata=[['Section','Part','JSON/text']]+[['reconciliation_json',index+1,snapshot[start:start+30000]] for index,start in enumerate(range(0,len(snapshot),30000))]
    sheets=[report.sheet(name,rows) for name,rows in [('Records',records),('Checks',checks),('Contributions',contributions),('Metadata',metadata)]]
    sheets.append(report.sheet('Warnings',[['Location','Issue','Original']]+report.warnings))
    appearance={'preset':'readable','sheets':{sheet['name']:{'header_rows':'1','column_widths':{'A':24},'wrap_text':sheet['name'] in ('Metadata','Warnings')} for sheet in sheets}}
    book=Workbook.create({'sheets':sheets},appearance=appearance); book.save(output)
    return {'status':'exported','output_written':True,'output':str(output),'formulas_created':0,'total_cells':report.count}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('left'); parser.add_argument('right'); parser.add_argument('--query',required=True); parser.add_argument('--xlsx')
    args=parser.parse_args()
    try:
        result=reconcile(read_spec(args.left,max_bytes=MAX_INPUT_BYTES),read_spec(args.right,max_bytes=MAX_INPUT_BYTES),read_spec(args.query,max_bytes=65536))
        if args.xlsx: result['export']=export_reconciliation(result,args.xlsx)
        print(json.dumps(result,ensure_ascii=False,default=str)); return 0
    except (ValueError,OSError,LookupError,DecimalException) as exc:
        print(json.dumps({'status':'failed','message':str(exc),'output_written':False},ensure_ascii=False)); return 1


if __name__=='__main__': sys.exit(main())
