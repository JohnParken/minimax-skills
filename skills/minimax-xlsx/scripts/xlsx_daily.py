#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Inclusive-calendar daily balance averages and daily-flow month/YTD sums."""
import argparse
from datetime import date,timedelta
from decimal import Decimal,DecimalException
import json
from pathlib import Path
import re
import sys
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import MAX_CELLS
from xlsx_analysis import region,validate_query,matches,is_blank
from xlsx_headers import resolve_column
from xlsx_rows import read_query
from xlsx_advanced import exact,rounded,precision_option
from xlsx_management import column_id,required_text


def calendar_date(value,location):
    if not isinstance(value,str) or not re.fullmatch(r'[0-9]{4}-[0-9]{2}-[0-9]{2}',value):
        raise ValueError('Explicit YYYY-MM-DD date required at '+location)
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError('Invalid calendar date at '+location) from exc


def windows(days,operation,periods):
    if operation=='average':
        return [('interval',days[0].isoformat()+'..'+days[-1].isoformat(),days)]
    output=[]
    for kind in periods:
        buckets={}
        for day in days:
            key='%04d-%02d' % (day.year,day.month) if kind=='month' else str(day.year)+'-YTD'
            buckets.setdefault(key,[]).append(day)
        for label,interval in buckets.items():
            role=('month_window' if interval[0].day==1 else 'partial_month_window') if kind=='month' else 'year_to_date'
            output.append((role,label,interval))
    return output


def summarize(source,sheet,config,operation='average',minimum_value=None,**options):
    allowed={'entity_keys','columns','metrics','start','end','date_format','definition_version','coverage','outside','filters','periods','precision'}
    if not isinstance(config,dict) or set(config)-allowed or operation not in ('average','accumulate'):
        raise ValueError('Unknown daily operation/configuration')
    if config.get('date_format')!='YYYY-MM-DD':
        raise ValueError('date_format: YYYY-MM-DD is required; serial dates are not inferred')
    start=calendar_date(config.get('start'),'start'); end=calendar_date(config.get('end'),'end')
    length=(end-start).days+1
    if not 1<=length<=366:
        raise ValueError('Daily interval must contain 1..366 inclusive calendar days')
    coverage=config.get('coverage','complete'); outside=config.get('outside','reject')
    if coverage not in ('complete','observed') or outside not in ('reject','exclude'):
        raise ValueError('coverage must be complete/observed; outside must be reject/exclude')
    periods=config.get('periods',['month','ytd'])
    if operation=='average' and 'periods' in config:
        raise ValueError('periods is only valid for daily-flow accumulation')
    if operation=='accumulate':
        if not isinstance(periods,list) or not periods or any(kind not in ('month','ytd') for kind in periods) or len(periods)!=len(set(periods)):
            raise ValueError('periods must be unique month/ytd entries')
        if 'ytd' in periods and (start.month,start.day)!=(1,1):
            raise ValueError('YTD requires an interval beginning January 1; partial years are not relabelled YTD')
    if minimum_value is not None and (isinstance(minimum_value,bool) or not isinstance(minimum_value,(int,Decimal)) or not Decimal(minimum_value).is_finite()):
        raise ValueError('Internal minimum_value must be a finite numeric bound')
    precision=precision_option(config); version=required_text(config.get('definition_version'),'definition_version')
    dictionary=config.get('metrics')
    if not isinstance(dictionary,list) or not 1<=len(dictionary)<=20:
        raise ValueError('Declare 1..20 daily source metrics')
    definitions={}
    expected_kind,expected_basis=('stock','daily_snapshot') if operation=='average' else ('flow','daily_flow')
    for definition in dictionary:
        if not isinstance(definition,dict) or set(definition)!={'metric_id','label','unit','kind','time_basis'}:
            raise ValueError('Daily metrics require metric_id/label/unit/kind/time_basis')
        for field in definition: required_text(definition[field],'daily metric/'+field)
        if definition['kind']!=expected_kind or definition['time_basis']!=expected_basis:
            raise ValueError('Expected '+expected_kind+'/'+expected_basis+'; month-end or YTD snapshots cannot substitute for daily inputs')
        if definition['metric_id'] in definitions: raise ValueError('Duplicate daily metric ID')
        definitions[definition['metric_id']]=definition
    required={'metric','date','value','unit','currency','version','basis'}
    selectors=config.get('columns'); entities=config.get('entity_keys')
    if not isinstance(selectors,dict) or set(selectors)!=required or not isinstance(entities,list) or not 1<=len(entities)<=8:
        raise ValueError('Daily columns require metric/date/value/unit/currency/version/basis and 1..8 entity selectors')
    columns,headers,rows,metadata=region(source,sheet,**options); schema=metadata.get('header_schema')
    fields={name:resolve_column(selector,columns,schema) for name,selector in selectors.items()}
    entity_columns=[resolve_column(selector,columns,schema) for selector in entities]
    if len(set(list(fields.values())+entity_columns))!=len(fields)+len(entity_columns):
        raise ValueError('Daily field/entity columns must be distinct')
    _,_,_,filters=validate_query({'filters':config.get('filters',[])},columns,schema)
    selected=[record for record in rows if matches(record[1],filters)]
    groups={}; outside_count=0
    for row_number,values in selected:
        item=values[columns.index(fields['date'])]
        if item[0] not in ('text','date_text'):
            raise ValueError('Daily dates must be explicit ISO text, not numeric serials at row '+str(row_number))
        day=calendar_date(item[1],fields['date']+str(row_number))
        if not start<=day<=end:
            if outside=='reject': raise ValueError('Date outside requested interval at row '+str(row_number))
            outside_count+=1; continue
        entity=[]
        for column in entity_columns:
            value=values[columns.index(column)]
            if value[0]!='text' or not value[1].strip(): raise ValueError('Daily entity IDs must be nonblank literal text')
            entity.append(value[1])
        metric=values[columns.index(fields['metric'])]
        if metric[0]!='text' or metric[1] not in definitions: raise ValueError('Unknown daily metric at row '+str(row_number))
        definition=definitions[metric[1]]
        unit=values[columns.index(fields['unit'])]; currency=values[columns.index(fields['currency'])]; actual_version=values[columns.index(fields['version'])]
        if unit!=('text',definition['unit']) or values[columns.index(fields['basis'])]!=('text',expected_basis):
            raise ValueError('Daily unit/time basis mismatch at row '+str(row_number))
        if currency[0]!='text' or not re.fullmatch(r'[A-Z]{3}',currency[1]) or actual_version!=('text',version):
            raise ValueError('Daily currency/definition version invalid at row '+str(row_number))
        key=tuple(entity)+(metric[1],currency[1],version)
        if key not in groups:
            if len(groups)>=1000: raise ValueError('More than 1000 daily groups')
            groups[key]={}
        if day in groups[key]:
            raise ValueError('Duplicate daily date '+day.isoformat()+' at rows %d and %d' % (groups[key][day][0],row_number))
        value=values[columns.index(fields['value'])]
        if is_blank(value): numeric=None
        elif value[0]!='number': raise ValueError('Daily values must be numeric without text coercion at row '+str(row_number))
        else: numeric=value[1]
        if numeric is not None and minimum_value is not None and numeric<minimum_value:
            raise ValueError('Negative daily deposit balances are outside this module scope at row '+str(row_number))
        groups[key][day]=(row_number,numeric)
    days=[start+timedelta(days=index) for index in range(length)]
    periods_to_emit=windows(days,operation,periods)
    generated=[]
    def add(role,identifier,label):
        identity=column_id(role,identifier) if role!='metric' else column_id(role,identifier,'actual')
        generated.append({'id':identity,'column':identity,'path':[{'type':'text','value':label}],
                          'levels':[{'type':'text','value':label,'role':'derived_daily'}]})
    for index,column in enumerate(entity_columns):
        add('dimension','entity:'+column,'实体 '+column)
    for name,label in (('source_metric','源指标'),('unit','单位'),('currency','币种'),('version','口径版本'),
                       ('period_kind','期间类型'),('period','期间'),('start','开始日期'),('end','结束日期')):
        add('dimension',name,label)
    primary='daily_average' if operation=='average' else 'period_flow_sum'
    observed='observed_daily_average' if operation=='average' else 'observed_flow_sum'
    for name,label in ((primary,'完整日均' if operation=='average' else '完整期间发生额'),(observed,'观测日均' if operation=='average' else '观测发生额'),
                       ('reported_days','有记录日数'),('observed_days','有效观测日数'),('expected_days','应覆盖日数'),('coverage_ratio','覆盖率 [倍]')):
        add('metric',name,label if coverage=='complete' or name!=primary else '观测口径主值')
    add('quality','flags','质量标记')
    if len(groups)*len(periods_to_emit)*len(generated)>MAX_CELLS:
        raise ValueError('Daily report exceeds 200000 typed output cells')
    output=[]; lineage=[]; details=[]
    for key,observations in groups.items():
        entity=list(key[:len(entity_columns)]); identifier,currency,actual_version=key[-3:]
        for kind,label,interval in periods_to_emit:
            subset={day:observations[day] for day in interval if day in observations}
            valid={day:record for day,record in subset.items() if record[1] is not None}
            missing_dates=[day.isoformat() for day in interval if day not in valid]
            total=exact(lambda:sum((record[1] for record in valid.values()),Decimal(0))) if valid else None
            observed_metric=None
            flags=[]
            if total is not None:
                if operation=='average':
                    mean=rounded(lambda:total/len(valid),precision); observed_metric=mean['value']
                    if mean['rounded']: flags.append('rounded_daily_average')
                else: observed_metric=total
            complete=not missing_dates
            primary_value=observed_metric if complete or coverage=='observed' else None
            if missing_dates: flags.append('incomplete_calendar_coverage')
            if coverage=='observed' and missing_dates: flags.append('observed_only_not_full_interval')
            if not valid: flags.append('no_numeric_observations')
            ratio=rounded(lambda:Decimal(len(valid))/len(interval),precision)
            values=[('text',value) for value in entity+[identifier,definitions[identifier]['unit'],currency,actual_version,kind,label,interval[0].isoformat(),interval[-1].isoformat()]]
            values += [('blank',None) if value is None else ('number',value) for value in
                       (primary_value,observed_metric,Decimal(len(subset)),Decimal(len(valid)),Decimal(len(interval)),ratio['value'])]
            values.append(('text',json.dumps(flags,ensure_ascii=False)))
            row_id=len(output)+1
            output.append({'row_id':row_id,'cells':[{'type':value_type,'value':value} for value_type,value in values],
                           'source_rows':[record[0] for record in subset.values()],'quality_flags':flags})
            details.append({'row_id':row_id,'expected_days':len(interval),'reported_days':len(subset),'observed_days':len(valid),
                            'missing_dates':missing_dates,'coverage_rounded':ratio['rounded']})
            for row_number,_ in subset.values():
                lineage.append({'output_row':row_id,'source_file':str(source),'sheet':sheet,'source_row':row_number,'metric_id':identifier})
                if len(lineage)>20000: raise ValueError('Daily audit exceeds 20000 lineage rows')
    summary_definitions=[]
    for item in generated:
        identity=json.loads(item['id'])
        if identity[0]!='metric': continue
        name=identity[1]
        unit='days' if name.endswith('_days') else 'fraction' if name=='coverage_ratio' else 'per-row source metric unit'
        aggregation={'reported_days':'REPORTED_DAY_COUNT','observed_days':'NUMERIC_DAY_COUNT','expected_days':'CALENDAR_DAY_COUNT',
                     'coverage_ratio':'NUMERIC_DAY_COUNT/CALENDAR_DAY_COUNT'}.get(name,'DAILY_MEAN' if operation=='average' else 'DAILY_FLOW_SUM')
        summary_definitions.append({'metric_id':item['id'],'label':item['path'][0]['value'],'unit':unit,
                                    'kind':'derived','time_basis':'inclusive_calendar_interval','aggregation':aggregation})
    validations=[{'check':'daily_unique_key','status':'pass','detail':'entity/metric/currency/version/date duplicates rejected'},
                 {'check':'daily_basis','status':'pass','detail':expected_kind+'/'+expected_basis+'; no YTD snapshots summed'},
                 {'check':'calendar_coverage','status':'warning' if any(item['missing_dates'] for item in details) else 'pass',
                  'detail':str(sum(bool(item['missing_dates']) for item in details))+' report intervals have missing/blank dates'}]
    if not groups:
        validations.append({'check':'observed_groups','status':'warning','detail':'No entities observed; missing entity universe is not inferred'})
    metadata.pop('precision',None)
    return {'status':'daily_summary','schema_version':'daily_summary_v1',**metadata,'source_file':str(source),
            'requested_query':config,'operation':operation,'definition_version':version,'report_level':'entity_daily_scope',
            'columns':[item['id'] for item in generated],'headers':[{'type':'text','value':item['path'][0]['value']} for item in generated],
            'generated_columns':generated,'rows':output,'output_rows':len(output),'lineage':lineage,'metric_dictionary':summary_definitions,
            'source_metric_dictionary':dictionary,'coverage_details':details,'coverage_policy':coverage,'outside_rows':outside_count,
            'filtered_out_rows':len(rows)-len(selected),'validation':validations,
            'arithmetic':{'exact_sum_precision_limit':1000,'result_precision':precision,'intermediate_precision':precision+16,'rounding':'ROUND_HALF_EVEN'},
            'policy':'inclusive Gregorian calendar days; no zero fill, forward fill, endpoint averaging or business-day guessing'}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('operation',choices=('average','accumulate')); parser.add_argument('input')
    parser.add_argument('--sheet',required=True); parser.add_argument('--query',required=True); parser.add_argument('--xlsx')
    parser.add_argument('--range',dest='cell_range'); parser.add_argument('--header-rows'); parser.add_argument('--no-header',action='store_true')
    parser.add_argument('--header-exclude',nargs='*',type=int,default=[]); parser.add_argument('--include-spanning-headings',action='store_true')
    parser.add_argument('--formulas',choices=('reject','cached'),default='reject')
    args=parser.parse_args()
    try:
        result=summarize(args.input,args.sheet,read_query(args.query),args.operation,cell_range=args.cell_range,header=not args.no_header,
                         header_rows=args.header_rows,header_exclude=args.header_exclude,include_spanning_headings=args.include_spanning_headings,formulas=args.formulas)
        if args.xlsx:
            from xlsx_result_export import export_result
            result['export']=export_result(result,args.xlsx)
        print(json.dumps(result,ensure_ascii=False,default=str)); return 0
    except (ValueError,OSError,LookupError,DecimalException,ET.ParseError,zipfile.BadZipFile) as exc:
        print(json.dumps({'status':'failed','message':str(exc),'output_written':False},ensure_ascii=False)); return 1


if __name__=='__main__': sys.exit(main())
