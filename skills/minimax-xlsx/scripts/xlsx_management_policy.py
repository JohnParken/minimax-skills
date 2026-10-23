#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Explicit monthly plan deviation and comparable peer benchmarking, without inferred policy."""
import argparse
import calendar
from decimal import Decimal,DecimalException
import json
from pathlib import Path
import sys

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook,read_spec
from xlsx_dashboard import build as dashboard
from xlsx_advanced import exact,rounded,precision_option,parse_month
from xlsx_result_export import Report,MAX_INPUT_BYTES,daily_table,source_columns
from xlsx_daily import calendar_date

TARGET_TYPES={('stock','point_in_time'):'balance',('flow','period_flow'):'period_flow',
              ('flow','ytd'):'ytd',('ratio','ratio'):'ratio',('stock','daily_average'):'daily_average'}



def native_view(source,config):
    if source.get('status')!='daily_summary':
        if 'daily_entity_role' in config: raise ValueError('daily_entity_role is only valid for a daily summary')
        view=dashboard(source,{'overview_period':config['periods'][-1],'metrics':[config['metric_id']]})
        definition=next(item for item in source['metric_dictionary'] if item['metric_id']==config['metric_id'])
        return view,definition
    if config.get('daily_entity_role')!='institution' or source.get('operation')!='average' or source.get('coverage_policy')!='complete':
        raise ValueError('Daily policy requires explicit institution entity role, average and complete coverage')
    daily_table(source,Report(),source_columns(source))
    query=source.get('requested_query',{})
    start=calendar_date(query.get('start'),'daily start'); end=calendar_date(query.get('end'),'daily end')
    if start.day!=1 or (start.year,start.month)!=(end.year,end.month) or end.day!=calendar.monthrange(end.year,end.month)[1]:
        raise ValueError('Daily management policy requires one complete natural-month interval')
    roles=[tuple(json.loads(value)) for value in source['columns']]
    entities=[role for role in roles if role[0]=='dimension' and role[1].startswith('entity:')]
    if len(entities)!=1: raise ValueError('Daily policy requires exactly one declared institution entity key')
    dictionary=source.get('source_metric_dictionary')
    if not isinstance(dictionary,list) or any(not isinstance(item,dict) for item in dictionary):
        raise ValueError('Daily source metric dictionary must be a list of objects')
    definitions=[item for item in dictionary if item.get('metric_id')==config['metric_id']]
    if len(definitions)!=1 or definitions[0].get('kind')!='stock' or definitions[0].get('time_basis')!='daily_snapshot':
        raise ValueError('Daily metric must be explicit stock/daily_snapshot')
    definition={**definitions[0],'time_basis':'daily_average','aggregation':'DAILY_MEAN'}
    records=[]; seen=set(); validator=Report()
    for row in source['rows']:
        values={role:validator.typed(cell,'daily source') for role,cell in zip(roles,row['cells'])}
        if values[('dimension','source_metric')]!=config['metric_id']: raise ValueError('Daily policy source contains additional metrics; narrow explicitly')
        dims={'institution':values[entities[0]],'period':start.isoformat()[:7],
              'report_level':config['scope']['report_level'],'currency':values[('dimension','currency')],'version':values[('dimension','version')]}
        if any(not isinstance(value,str) or not value for value in dims.values()) or values[('dimension','unit')]!=definition['unit']:
            raise ValueError('Invalid daily scope metadata')
        if values[('dimension','start')]!=start.isoformat() or values[('dimension','end')]!=end.isoformat():
            raise ValueError('Daily row interval differs from declared month')
        key=tuple(dims.items())
        if key in seen: raise ValueError('Duplicate daily institution/month scope')
        seen.add(key)
        flags=list(row['quality_flags'])
        if source.get('unverified_formula_caches'): flags.append('unverified_source_formula_caches')
        if json.loads(values[('quality','flags')])!=row['quality_flags']: raise ValueError('Daily quality diagnostics disagree')
        expected_days=Decimal((end-start).days+1)
        counts={}
        for name in ('reported_days','observed_days','expected_days'):
            role=('metric',name,'actual')
            if role not in values or row['cells'][roles.index(role)]['type']!='number':
                raise ValueError('Daily coverage counts must be explicit numbers')
            count=Decimal(values[role])
            if count!=count.to_integral_value() or count<0:
                raise ValueError('Daily coverage counts must be nonnegative integers')
            counts[name]=count
        if counts['expected_days']!=expected_days or not counts['observed_days']<=counts['reported_days']<=expected_days:
            raise ValueError('Daily coverage counts contradict the natural-month interval')
        actual=values[('metric','daily_average','actual')]
        complete=counts['observed_days']==expected_days
        if (actual is not None)!=complete or not complete and 'incomplete_calendar_coverage' not in flags:
            raise ValueError('Daily mean/quality contradicts complete-coverage counts')
        if actual is not None:
            cell=row['cells'][roles.index(('metric','daily_average','actual'))]
            if cell['type']!='number': raise ValueError('Daily actual requires typed number')
            actual=Decimal(actual)
        records.append({'dimensions':dims,'actual':actual,'quality_flags':flags,'wide_row_id':row['row_id']})
    return {'dimensions':['institution','period','report_level','currency','version'],'trends':records},definition


def analyze(source,config):
    required={'metric_id','scope','periods','institutions','context','plan_definition','plans','rank_method'}
    if not isinstance(config,dict) or not required<=set(config) or set(config)-required-{'precision','daily_entity_role'}:
        raise ValueError('Require metric/scope/periods/institutions/context/plan_definition/plans/rank_method')
    if not isinstance(source,dict) or len(json.dumps(source,ensure_ascii=False,allow_nan=False,default=str).encode('utf-8'))>MAX_INPUT_BYTES:
        raise ValueError('Require a native source result within 16 MiB')
    precision=precision_option(config)
    periods=config['periods']
    if not isinstance(periods,list) or not 1<=len(periods)<=36 or any(not isinstance(value,str) for value in periods):
        raise ValueError('Declare 1..36 consecutive monthly periods in order')
    numbers=[parse_month(('text',value),'policy period') for value in periods]
    if numbers!=list(range(numbers[0],numbers[0]+len(numbers))): raise ValueError('Configured periods must be unique, ascending, consecutive months')
    scope=config['scope']; required_scope={'report_level','currency','version','definition_version','unit','kind','time_basis','period_granularity'}
    if not isinstance(scope,dict) or set(scope)!=required_scope or any(not isinstance(value,str) or not value.strip() for value in scope.values()) or scope['period_granularity']!='month':
        raise ValueError('Explicit level/currency/version/definition_version/unit/kind/time_basis/month scope required')
    target_type=TARGET_TYPES.get((scope['kind'],scope['time_basis']))
    if target_type is None: raise ValueError('Unsupported source time basis; daily means must not be relabelled point-in-time balances')
    plan_definition=config['plan_definition']
    if not isinstance(plan_definition,dict) or set(plan_definition)!={'target_type','unit','kind','time_basis','currency','version','definition_version','plan_version','direction'}:
        raise ValueError('Declare complete comparable plan definition, version and direction')
    if any(plan_definition[key]!=scope[key] for key in ('unit','kind','time_basis','currency','version','definition_version')) or plan_definition['target_type']!=target_type:
        raise ValueError('Actual and plan units, target type, currency, basis or versions differ')
    if not isinstance(plan_definition['plan_version'],str) or not plan_definition['plan_version'].strip() or plan_definition['direction'] not in ('at_least','at_most'):
        raise ValueError('Plan version and at_least/at_most direction required')
    method=config['rank_method']
    if method not in ('competition','dense','ordinal'): raise ValueError('Explicit competition/dense/ordinal rank method required')
    institutions=config['institutions']
    if not isinstance(institutions,list) or not 1<=len(institutions)<=1000 or len(institutions)*len(periods)>10000:
        raise ValueError('At most 1000 explicit institutions and 10000 institution/period combinations')
    mapping={}
    for item in institutions:
        if not isinstance(item,dict) or set(item)!={'id','level','peer_group'} or any(not isinstance(value,str) or not value.strip() for value in item.values()):
            raise ValueError('Institution needs explicit id, level and peer_group')
        if item['id'] in mapping or item['level']!=scope['report_level']: raise ValueError('Duplicate institution or mixed reporting levels')
        mapping[item['id']]=item
    view,definition=native_view(source,config)
    if source.get('definition_version')!=scope['definition_version']: raise ValueError('Source definition version differs')
    if any(definition[key]!=scope[key] for key in ('unit','kind','time_basis')): raise ValueError('Source metric dictionary differs from scope')
    if scope['kind']=='ratio' and (not isinstance(definition.get('numerator'),str) or not isinstance(definition.get('denominator'),str)):
        raise ValueError('Ratio dictionary must retain numerator/denominator definitions')
    if scope['kind']=='ratio' and not any(item.get('check')=='ratios_recomputed' and item.get('status')=='pass' for item in source.get('validation',[])):
        raise ValueError('Ratio source must retain validated aggregate-then-divide evidence')
    context=config['context']; remaining=set(view['dimensions'])-{'institution','institution_name','period','currency','version','report_level'}
    if not isinstance(context,dict) or set(context)!=remaining or any(not isinstance(value,str) or not value for value in context.values()):
        raise ValueError('Context must explicitly fix every remaining comparison dimension')
    plans=config['plans']
    if not isinstance(plans,list) or len(plans)>10000: raise ValueError('At most 10000 explicit period plans; no inferred linear allocation')
    plan_index={}
    for plan in plans:
        if not isinstance(plan,dict) or set(plan)!={'institution','period','applicability','value'}:
            raise ValueError('Each plan needs institution/period/applicability/value')
        if not isinstance(plan['institution'],str) or plan['institution'] not in mapping or not isinstance(plan['period'],str) or plan['period'] not in periods or plan['applicability'] not in ('applicable','not_applicable'):
            raise ValueError('Plan institution/period/applicability outside declared scope')
        value=plan['value']
        if plan['applicability']=='not_applicable':
            if value is not None: raise ValueError('Not-applicable plan must have null value, never zero')
        elif isinstance(value,bool) or not isinstance(value,(int,Decimal)) or not Decimal(value).is_finite():
            raise ValueError('Applicable plan must have a finite explicit numeric value')
        key=(plan['institution'],plan['period'])
        if key in plan_index: raise ValueError('Duplicate institution/period plan')
        plan_index[key]=plan
    actuals={}; excluded=0
    for record in view['trends']:
        dims=record['dimensions']; key=(dims['institution'],dims['period'])
        if dims['institution'] not in mapping or dims['period'] not in periods or any(dims[name]!=value for name,value in context.items()):
            excluded+=1; continue
        if any(dims[name]!=scope[name] for name in ('currency','version','report_level')):
            raise ValueError('Selected institution/period has incompatible currency/version/level; narrow source explicitly')
        if key in actuals: raise ValueError('Ambiguous institution/period actuals')
        actuals[key]=record
    rows=[]; peer_groups={}; quality=[]; lineage=[]
    for institution,item in mapping.items():
        for period in periods:
            source_row=actuals.get((institution,period)); plan=plan_index.get((institution,period))
            actual=None if source_row is None else source_row['actual']
            flags=[] if source_row is None else list(source_row['quality_flags'])
            observation='missing_record' if source_row is None else 'missing_value' if actual is None else 'observed_zero' if actual==0 else 'observed'
            target=None if plan is None else plan['value']
            status='missing_plan' if plan is None else 'not_applicable' if plan['applicability']=='not_applicable' else 'unverified_or_quality_actual' if flags else 'missing_actual' if actual is None else 'nonpositive_plan' if target<=0 else 'comparable'
            deviation=None; gap=None; attainment=None; ratio_rounded=False; behind=None
            if status=='comparable':
                deviation=exact(lambda:actual-target if plan_definition['direction']=='at_least' else target-actual)
                gap=exact(lambda:max(-deviation,Decimal(0))); behind=deviation<0
                if plan_definition['direction']=='at_least' or actual>0:
                    ratio=rounded(lambda:actual/target if plan_definition['direction']=='at_least' else target/actual,precision)
                    attainment=ratio['value']; ratio_rounded=ratio['rounded']
                else: flags.append('undefined_attainment_nonpositive_actual')
            record={'institution':institution,'period':period,'peer_group':item['peer_group'],'report_level':scope['report_level'],
                    'metric_id':config['metric_id'],'unit':scope['unit'],'currency':scope['currency'],'version':scope['version'],
                    'actual':actual,'observation_status':observation,'planned_to_date':target,'comparison_status':status,
                    'favorable_deviation':deviation,'gap':gap,'attainment':attainment,'attainment_rounded':ratio_rounded,
                    'behind_plan':behind,'peer_rank':None,'peer_count':None,'rank_status':'excluded_not_comparable',
                    'quality_flags':flags,'source_report_row':None if source_row is None else source_row['wide_row_id']}
            rows.append(record)
            if status=='comparable' and attainment is not None and actual>=0:
                peer_groups.setdefault((period,item['peer_group']),[]).append(record)
            elif status=='comparable': record['rank_status']='excluded_undefined_or_negative_attainment'
            if status!='comparable': quality.append({'institution':institution,'period':period,'issue':status})
            for flag in flags: quality.append({'institution':institution,'period':period,'issue':flag})
            if source_row is not None: lineage.append({'institution':institution,'period':period,'source_report_row_id':source_row['wide_row_id']})
    for group in peer_groups.values():
        group.sort(key=lambda row:row['attainment'],reverse=True)
        previous=None; dense=0; rank=0
        for index,row in enumerate(group,1):
            if previous!=row['attainment']: dense+=1; rank=index
            row['peer_rank']=index if method=='ordinal' else dense if method=='dense' else rank
            row['peer_count']=len(group); row['rank_status']='ranked_attainment_in_explicit_peer_group'
            previous=row['attainment']
    continuity=[]
    for institution in mapping:
        history=[row for row in rows if row['institution']==institution]
        complete=all(row['behind_plan'] is not None for row in history)
        count=None
        if complete:
            count=0
            for row in reversed(history):
                if not row['behind_plan']: break
                count+=1
        continuity.append({'institution':institution,'start':periods[0],'end':periods[-1],
                           'trailing_behind_periods':count,'status':'complete_consecutive_window' if complete else 'unknown_incomplete_or_inapplicable_window'})
    return {'status':'management_policy','schema_version':'management_policy_v1','requested_query':config,
            'metric_dictionary':definition,'rows':rows,'continuity':continuity,'quality':quality,'lineage':lineage,
            'excluded_source_records':excluded,'source_report':source,'source_row_kind':'daily_summary_row' if source.get('status')=='daily_summary' else 'management_wide_row',
            'policies':{'plans':'Explicit per-period planned-to-date values; no linear allocation or cumulative reconstruction',
                        'time_basis':'point_in_time / period_flow / ytd / ratio remain distinct; no summing periods',
                        'deviation':'at_least actual-plan; at_most plan-actual; negative means behind',
                        'gap':'max(-favorable_deviation,0); achievement at_least actual/positive plan, at_most positive plan/positive actual',
                        'peers':'Rank only comparable nonnegative attainment in same explicit group, period and validated scope/context; no raw-size ranking',
                        'continuity':'Whole configured consecutive window must be comparable; missing/inapplicable periods yield unknown, never zero',
                        'ratios':'Use only upstream validated aggregate-then-divide actuals; no averaging ratios',
                        'authority':'User configured management definitions; no bank-policy or regulatory conformity claim',
                        'precision':precision,'rounding':'ROUND_HALF_EVEN','exact_precision_limit':1000}}


def export_policy(result,output):
    result=analyze(result['source_report'],result['requested_query']); report=Report()
    fields=['institution','period','peer_group','report_level','metric_id','unit','currency','version','actual','observation_status','planned_to_date','comparison_status','favorable_deviation','gap','attainment','attainment_rounded','behind_plan','peer_rank','peer_count','rank_status','source_report_row']
    numeric={'actual','planned_to_date','favorable_deviation','gap','attainment','peer_rank','peer_count','source_report_row'}
    progress=[fields+['quality_flags']]
    for row in result['rows']:
        progress.append([report.number(row[key],key) if key in numeric else row[key] for key in fields]+[json.dumps(row['quality_flags'],ensure_ascii=False)])
    continuity=[['Institution','Start','End','Trailing behind periods','Status']]+[[row['institution'],row['start'],row['end'],row['trailing_behind_periods'],row['status']] for row in result['continuity']]
    definitions=[['Section','JSON']]+[[key,json.dumps(result['requested_query'][key],ensure_ascii=False,default=str)] for key in ('scope','context','plan_definition')]+[['metric_dictionary',json.dumps(result['metric_dictionary'],ensure_ascii=False,default=str)]]
    definitions += [['institution',json.dumps(item,ensure_ascii=False)] for item in result['requested_query']['institutions']]
    calculations=[['Rule','Definition']]+[[key,str(value)] for key,value in result['policies'].items()]
    lineage=[['Institution','Period','Source report row ID']]+[[row['institution'],row['period'],row['source_report_row_id']] for row in result['lineage']]
    quality=[['Institution','Period','Issue']]+[[row['institution'],row['period'],row['issue']] for row in result['quality']]
    snapshot=json.dumps(result,ensure_ascii=False,allow_nan=False,default=str)
    if len(snapshot.encode('utf-8'))>MAX_INPUT_BYTES: raise ValueError('Policy snapshot exceeds 16 MiB')
    metadata=[['Section','Part','JSON/text']]+[['policy_json',index+1,snapshot[start:start+30000]] for index,start in enumerate(range(0,len(snapshot),30000))]
    sheets=[report.sheet(name,table) for name,table in [('Progress',progress),('Continuity',continuity),('Dictionary',definitions),('Calculations',calculations),('Lineage',lineage),('Quality',quality),('Metadata',metadata)]]
    sheets.append(report.sheet('Warnings',[['Location','Issue','Original']]+report.warnings))
    appearance={'preset':'readable','sheets':{sheet['name']:{'header_rows':'1','column_widths':{'A':24}} for sheet in sheets}}
    appearance['sheets']['Progress']['category_columns']=['L']
    book=Workbook.create({'sheets':sheets},appearance=appearance); book.save(output)
    return {'status':'exported','output_written':True,'output':str(output),'formulas_created':0,'total_cells':report.count}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input'); parser.add_argument('--query',required=True); parser.add_argument('--xlsx')
    args=parser.parse_args()
    try:
        result=analyze(read_spec(args.input,max_bytes=MAX_INPUT_BYTES),read_spec(args.query,max_bytes=65536))
        if args.xlsx: result['export']=export_policy(result,args.xlsx)
        print(json.dumps(result,ensure_ascii=False,default=str)); return 0
    except (ValueError,OSError,LookupError,DecimalException) as exc:
        print(json.dumps({'status':'failed','message':str(exc),'output_written':False},ensure_ascii=False)); return 1


if __name__=='__main__': sys.exit(main())
