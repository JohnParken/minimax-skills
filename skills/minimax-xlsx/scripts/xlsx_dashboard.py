#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Build a bounded, static management dashboard from native management-wide JSON."""
import argparse
from decimal import Decimal, DecimalException
import json
from pathlib import Path
import sys

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook, read_spec
from xlsx_advanced import parse_month
from xlsx_result_export import Report, MAX_INPUT_BYTES, management_table, source_columns
from xlsx_tabular import column_name

OPERATORS={'gt':lambda a,b:a>b,'ge':lambda a,b:a>=b,'lt':lambda a,b:a<b,
           'le':lambda a,b:a<=b,'eq':lambda a,b:a==b,'ne':lambda a,b:a!=b}


def month(value):
    if not isinstance(value,str):
        raise ValueError('Period must be explicit YYYY-MM text')
    return parse_month(('text',value),'dashboard period')


def build(source,config):
    if not isinstance(source,dict) or source.get('status')!='management_wide' or source.get('schema_version')!='management_wide_v1':
        raise ValueError('Dashboard requires native management_wide_v1')
    if not isinstance(config,dict) or set(config)-{'overview_period','metrics','ranks','thresholds','trend_metrics'}:
        raise ValueError('Unknown dashboard configuration fields')
    if len(json.dumps(source,ensure_ascii=False,allow_nan=False,default=str).encode('utf-8'))>MAX_INPUT_BYTES:
        raise ValueError('Source report exceeds 16 MiB')
    month(config.get('overview_period'))
    # Reuse the native export validator, including typed cells and sequential row IDs.
    validator=Report()
    management_table(source,validator,source_columns(source))
    if len(source['rows'])>5000:
        raise ValueError('Dashboard supports at most 5000 source rows')
    roles=[tuple(json.loads(identifier)) for identifier in source['columns']]
    if len(set(roles))!=len(roles):
        raise ValueError('Duplicate semantic column IDs')
    dims=[role[1] for role in roles if role[0]=='dimension']
    if not {'period','institution','currency','version','report_level'}<=set(dims):
        raise ValueError('Missing protected comparison dimensions')
    dictionary=source.get('metric_dictionary')
    if not isinstance(dictionary,list) or any(not isinstance(item,dict) or not all(isinstance(item.get(key),str) for key in ('metric_id','label','unit','kind','time_basis')) for item in dictionary):
        raise ValueError('Invalid metric dictionary')
    definitions={item['metric_id']:item for item in dictionary}
    if len(definitions)!=len(dictionary):
        raise ValueError('Duplicate dictionary metric IDs')
    metrics=config.get('metrics')
    if not isinstance(metrics,list) or not metrics or len(metrics)>20 or any(not isinstance(key,str) or key not in definitions for key in metrics) or len(set(metrics))!=len(metrics):
        raise ValueError('Select 1..20 unique dictionary metric IDs')
    trends=config.get('trend_metrics',metrics)
    if not isinstance(trends,list) or any(not isinstance(key,str) or key not in metrics for key in trends) or len(set(trends))!=len(trends):
        raise ValueError('trend_metrics must be a unique selected subset')
    ranks=config.get('ranks',{})
    if not isinstance(ranks,dict) or set(ranks)-set(metrics):
        raise ValueError('Rank metrics must be selected')
    for rule in ranks.values():
        if not isinstance(rule,dict) or set(rule)!={'method','descending'} or not isinstance(rule['method'],str) or rule['method'] not in ('competition','dense','ordinal') or type(rule['descending']) is not bool:
            raise ValueError('Ranks require explicit method and boolean descending')
    thresholds=config.get('thresholds',[])
    if not isinstance(thresholds,list) or len(thresholds)>100:
        raise ValueError('At most 100 explicit threshold rules')
    for rule in thresholds:
        if not isinstance(rule,dict) or set(rule)!={'metric_id','field','op','value','unit','severity','scope'}:
            raise ValueError('Threshold requires metric_id/field/op/value/unit/severity/scope')
        key=rule['metric_id']; field=rule['field']; scope=rule['scope']
        if not isinstance(key,str) or key not in metrics or not isinstance(field,str) or field not in ('actual','achievement','gap') or not isinstance(rule['op'],str) or rule['op'] not in OPERATORS:
            raise ValueError('Unknown threshold metric/field/operator')
        if rule['unit']!=('multiple' if field=='achievement' else definitions[key]['unit']):
            raise ValueError('Threshold unit differs from metric field')
        if not isinstance(rule['severity'],str) or not rule['severity'].strip():
            raise ValueError('Threshold severity must be explicit text')
        if not isinstance(scope,dict) or not {'currency','version','report_level'}<=set(scope) or set(scope)-set(dims) or any(not isinstance(value,str) or not value for value in scope.values()):
            raise ValueError('Threshold scope must retain currency/version/report_level')
        if isinstance(rule['value'],bool) or not isinstance(rule['value'],(int,Decimal)) or not Decimal(rule['value']).is_finite():
            raise ValueError('Threshold value must be a finite JSON number')
    records=[]; seen=set()
    for row in source['rows']:
        values={role:validator.typed(cell,'source row') for role,cell in zip(roles,row['cells'])}
        dimensions={key:values[('dimension',key)] for key in dims}
        if any(not isinstance(value,str) or not value for value in dimensions.values()):
            raise ValueError('Dashboard dimensions must be nonempty literal text')
        month(dimensions['period'])
        quality=values.get(('quality','flags'))
        if not isinstance(quality,str) or json.loads(quality)!=row['quality_flags']:
            raise ValueError('Quality column and row diagnostics disagree')
        context=tuple((key,value) for key,value in dimensions.items() if key!='institution_name')
        if context in seen:
            raise ValueError('Duplicate management comparison context')
        seen.add(context)
        flags=list(row['quality_flags'])
        if source.get('unverified_formula_caches'):
            flags.append('unverified_source_formula_caches')
        for key in metrics:
            fields={}
            for field in ('actual','target','achievement','gap'):
                role=('metric',key,field)
                if field=='actual' and role not in values:
                    raise ValueError('Selected metric lacks actual column')
                value=values.get(role)
                cell=row['cells'][roles.index(role)] if role in values else {'type':'blank'}
                if value is not None and cell['type']!='number':
                    raise ValueError('Metric fields require typed numbers, never numeric text')
                fields[field]=None if value is None else Decimal(value)
            definition=definitions[key]
            records.append({'wide_row_id':row['row_id'],'dimensions':dimensions,'metric_id':key,
                            'label':definition['label'],'unit':definition['unit'],'kind':definition['kind'],
                            'time_basis':definition['time_basis'],**fields,'quality_flags':flags,
                            'rank':None,'peer_count':None,'rank_status':'not_requested'})
            if len(records)>20000:
                raise ValueError('Dashboard exceeds 20000 metric records')
    overview=[record.copy() for record in records if record['dimensions']['period']==config['overview_period']]
    peers={}
    for record in overview:
        key=record['metric_id']
        if key not in ranks:
            continue
        record['rank_status']='excluded_quality' if record['quality_flags'] else 'missing_actual' if record['actual'] is None else 'ranked'
        if record['rank_status']=='ranked':
            context=(key,tuple((name,value) for name,value in record['dimensions'].items() if name not in ('institution','institution_name')))
            peers.setdefault(context,[]).append(record)
    for (key,_),group in peers.items():
        rule=ranks[key]; group.sort(key=lambda item:item['actual'],reverse=rule['descending'])
        previous=None; dense=0; rank=0
        for index,record in enumerate(group,1):
            if record['actual']!=previous:
                dense+=1; rank=index
            record['rank']=index if rule['method']=='ordinal' else dense if rule['method']=='dense' else rank
            record['peer_count']=len(group); previous=record['actual']
    alerts=[]; matched=set()
    for record in overview:
        for index,rule in enumerate(thresholds,1):
            if rule['metric_id']!=record['metric_id'] or any(record['dimensions'][key]!=value for key,value in rule['scope'].items()):
                continue
            matched.add(index); value=record[rule['field']]
            status='not_evaluated_missing' if value is None else 'not_evaluated_quality' if record['quality_flags'] else 'triggered' if OPERATORS[rule['op']](value,rule['value']) else 'within_configured_threshold'
            alerts.append({'rule_id':index,'wide_row_id':record['wide_row_id'],'metric_id':record['metric_id'],
                           'dimensions':record['dimensions'],'field':rule['field'],'value':value,'unit':rule['unit'],
                           'threshold':rule['value'],'operator':rule['op'],'severity':rule['severity'],
                           'status':status,'quality_flags':record['quality_flags']})
            if len(alerts)>20000:
                raise ValueError('Dashboard exceeds 20000 alert evaluations')
    trend=[record for record in records if record['metric_id'] in trends]
    trend.sort(key=lambda item:(tuple((key,value) for key,value in item['dimensions'].items() if key!='period'),item['metric_id'],month(item['dimensions']['period'])))
    warnings=[]
    if not overview: warnings.append('No observations for requested overview period')
    if len(matched)!=len(thresholds): warnings.append('Unmatched threshold rules: '+str(sorted(set(range(1,len(thresholds)+1))-matched)))
    if source.get('unverified_formula_caches'): warnings.append('Source formula caches are unverified')
    return {'status':'management_dashboard','schema_version':'management_dashboard_v1','requested_query':config,
            'dimensions':dims,'overview':overview,'alerts':alerts,'trends':trend,'warnings':warnings,
            'charts_created':False,'source_report':source,
            'policies':{'ranking':'clean rows only; all dimensions except institution/name retained; stable ordinal ties',
                        'thresholds':'User configured management rules, never regulatory defaults; no rule means not evaluated',
                        'trends':'Observed periods only; no interpolation, summation or charts',
                        'aggregation':'No dashboard totals; ancestor and descendant rows remain separate'}}


def export_dashboard(result,output):
    # Rebuild from audited input/config instead of trusting derived caller JSON.
    result=build(result['source_report'],result['requested_query'])
    report=Report(); dims=result['dimensions']
    numeric=('actual','target','achievement','gap','rank','peer_count')
    overview=[dims+['Wide row ID','Metric','Label','Unit','Kind','Time basis']+list(numeric)+['Rank status','Quality flags']]
    trend=[dims+['Wide row ID','Metric','Label','Unit','Kind','Time basis','Actual','Quality flags']]
    def prefix(record):
        return [record['dimensions'][key] for key in dims]+[record['wide_row_id'],record['metric_id'],record['label'],record['unit'],record['kind'],record['time_basis']]
    for record in result['overview']:
        overview.append(prefix(record)+[report.number(record[key],key) for key in numeric]+[record['rank_status'],json.dumps(record['quality_flags'],ensure_ascii=False)])
    for record in result['trends']:
        trend.append(prefix(record)+[report.number(record['actual'],'trend actual'),json.dumps(record['quality_flags'],ensure_ascii=False)])
    alerts=[dims+['Rule ID','Wide row ID','Metric','Field','Value','Unit','Operator','Threshold','Severity','Status','Quality flags']]
    for record in result['alerts']:
        alerts.append([record['dimensions'][key] for key in dims]+[record['rule_id'],record['wide_row_id'],record['metric_id'],record['field'],report.number(record['value'],'alert value'),record['unit'],record['operator'],report.number(record['threshold'],'threshold'),record['severity'],record['status'],json.dumps(record['quality_flags'],ensure_ascii=False)])
    snapshot=json.dumps(result,ensure_ascii=False,allow_nan=False,default=str)
    if len(snapshot.encode('utf-8'))>MAX_INPUT_BYTES:
        raise ValueError('Dashboard snapshot exceeds 16 MiB')
    metadata=[['Section','Part','JSON/text']]+[['dashboard_json',index+1,snapshot[start:start+30000]] for index,start in enumerate(range(0,len(snapshot),30000))]
    definitions=[['Metric','Label','Unit','Kind','Time basis','Definition JSON']]+[[item['metric_id'],item['label'],item['unit'],item['kind'],item['time_basis'],json.dumps(item,ensure_ascii=False,default=str)] for item in result['source_report']['metric_dictionary']]
    sheets=[report.sheet(name,rows) for name,rows in [('Overview',overview),('Alerts',alerts),('Trends',trend),('Definitions',definitions),('Metadata',metadata)]]
    sheets.append(report.sheet('Warnings',[['Location','Issue','Original']]+[['dashboard',warning,None] for warning in result['warnings']]+report.warnings))
    appearance={'preset':'readable','sheets':{sheet['name']:{'header_rows':'1','column_widths':{'A':24},'wrap_text':sheet['name'] in ('Metadata','Warnings')} for sheet in sheets}}
    appearance['sheets']['Alerts']['category_columns']=[column_name(len(dims)+10)]
    book=Workbook.create({'sheets':sheets},appearance=appearance); book.save(output)
    return {'status':'exported','output':str(output),'output_written':True,'formulas_created':0,'charts_created':False,'total_cells':report.count}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input'); parser.add_argument('--query',required=True); parser.add_argument('--xlsx')
    args=parser.parse_args()
    try:
        result=build(read_spec(args.input,max_bytes=MAX_INPUT_BYTES),read_spec(args.query,max_bytes=65536))
        if args.xlsx: result['export']=export_dashboard(result,args.xlsx)
        print(json.dumps(result,ensure_ascii=False,default=str)); return 0
    except (ValueError,OSError,LookupError,DecimalException) as exc:
        print(json.dumps({'status':'failed','message':str(exc),'output_written':False},ensure_ascii=False)); return 1


if __name__=='__main__':
    sys.exit(main())
