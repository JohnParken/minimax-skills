#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Aggregate observed customers and overlapping product categories without exporting ID lists."""
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
from xlsx_advanced import parse_month,rounded,precision_option
from xlsx_result_export import Report,MAX_INPUT_BYTES,management_table,source_columns
from xlsx_management import column_id

FIELDS={'account','customer','institution','product','holding','period','version','level','scope'}


def read_customers(source,spec,period,config,institutions,categories):
    if not isinstance(spec,dict) or set(spec)-{'sheet','columns','range','header_rows'} or not isinstance(spec.get('sheet'),str):
        raise ValueError('Each side needs sheet/columns and optional range/header_rows')
    expected=FIELDS|({'owner'} if config['ownership']=='designated' else set())
    if not isinstance(spec.get('columns'),dict) or set(spec['columns'])!=expected:
        raise ValueError('Customer columns require account/customer/institution/product/holding/period/version/level/scope and owner for designated mode')
    columns,_,rows,metadata=region(source,spec['sheet'],cell_range=spec.get('range'),header_rows=spec.get('header_rows'))
    mapping={name:resolve_column(selector,columns,metadata.get('header_schema')) for name,selector in spec['columns'].items()}
    if len(set(mapping.values()))!=len(mapping): raise ValueError('Customer field columns must be distinct')
    customers={}; seen=set(); accounts={}
    for number,values in rows:
        fields={name:values[columns.index(column)] for name,column in mapping.items()}
        for name in expected-{'holding'}:
            if fields[name][0]!='text' or not fields[name][1].strip():
                raise ValueError('Customer identifiers/metadata must be nonblank literal text at row '+str(number))
        if fields['holding'][0]!='boolean': raise ValueError('holding must be explicit boolean, never inferred from balance or account presence')
        if fields['period']!=('text',period) or fields['version']!=('text',config['scope']['version']) or fields['level']!=('text',config['scope']['level']) or fields['scope']!=('text',config['scope']['observation_scope']):
            raise ValueError('Period/version/level/observation scope differs from declared comparison')
        account=fields['account'][1]; customer=fields['customer'][1]; institution=fields['institution'][1]; product=fields['product'][1]
        if institution not in institutions: raise ValueError('Unmapped observed institution; narrow source explicitly')
        if product not in config['product_mapping']: raise ValueError('Unmapped product code at row '+str(number))
        owner=fields['owner'][1] if 'owner' in fields else institution
        if owner not in institutions: raise ValueError('Designated owner is outside institution scope')
        grain=(account,product)
        if grain in seen: raise ValueError('Duplicate account/product observation; no guessed deduplication')
        seen.add(grain)
        account_identity=(customer,institution,owner)
        if account in accounts and accounts[account]!=account_identity: raise ValueError('One account has conflicting customer/institution/owner mappings')
        accounts[account]=account_identity
        item=customers.setdefault(customer,{'institutions':set(),'owners':set(),'products':set(),'source_rows':[]})
        item['institutions'].add(institution); item['owners'].add(owner); item['source_rows'].append(number)
        if fields['holding'][1]: item['products'].update(config['product_mapping'][product])
        if len(customers)>5000 or len(seen)>20000: raise ValueError('At most 5000 customers/20000 account-product observations per period')
    for item in customers.values():
        if len(item['owners'])!=1: raise ValueError('Customer has multiple institution owners; configure one consistent designated owner')
        item['owner']=next(iter(item['owners']))
        if item['owner'] not in item['institutions']: raise ValueError('Designated owner must have an observed account for that customer')
    return customers,{'file':str(source),'sheet':spec['sheet'],'region':metadata,'columns':mapping,'observations':len(rows)},accounts


def analyze(left_source,right_source,config):
    required={'left','right','periods','scope','institutions','ownership','categories','product_mapping','coverage_denominator','fact_grain'}
    if not isinstance(config,dict) or not required<=set(config) or set(config)-required-{'precision'}:
        raise ValueError('Require both inputs, periods, scope, institutions, ownership, categories, product mapping, denominator and fact_grain')
    if config['fact_grain']!='account_product': raise ValueError('Explicit account_product observation grain required')
    periods=config['periods']
    if not isinstance(periods,dict) or set(periods)!={'left','right'} or any(not isinstance(value,str) for value in periods.values()): raise ValueError('Explicit left/right YYYY-MM periods required')
    months=[parse_month(('text',periods[name]),name+' period') for name in ('left','right')]
    if months[0]>=months[1]: raise ValueError('Right observation period must be later than left')
    scope=config['scope']
    if not isinstance(scope,dict) or set(scope)!={'level','version','definition_version','observation_scope','completeness'} or any(not isinstance(value,str) or not value.strip() for value in scope.values()) or scope['completeness'] not in ('declared_complete','unknown'):
        raise ValueError('Declare level/version/definition_version/observation_scope and completeness declared_complete/unknown')
    if config['ownership'] not in ('exclusive','designated'): raise ValueError('Ownership must be exclusive or explicit designated owner column')
    if config['coverage_denominator'] not in ('observed_customers','holding_customers'): raise ValueError('Coverage denominator must be observed_customers or holding_customers')
    precision=precision_option(config)
    entries=config['institutions']
    if not isinstance(entries,list) or not 1<=len(entries)<=100: raise ValueError('Declare 1..100 scoped institutions')
    institutions={}
    for item in entries:
        if not isinstance(item,dict) or set(item)!={'id','level','peer_group'} or any(not isinstance(value,str) or not value.strip() for value in item.values()) or item['level']!=scope['level'] or item['id'] in institutions:
            raise ValueError('Institutions need unique id, same level and explicit peer_group')
        institutions[item['id']]=item
    definitions=config['categories']
    if not isinstance(definitions,list) or not 1<=len(definitions)<=32 or any(not isinstance(item,dict) or set(item)!={'id','label'} or any(not isinstance(value,str) or not value.strip() for value in item.values()) for item in definitions):
        raise ValueError('Declare 1..32 categories with id/label')
    categories={item['id']:item['label'] for item in definitions}
    if len(categories)!=len(definitions): raise ValueError('Duplicate category IDs')
    mapping=config['product_mapping']
    if not isinstance(mapping,dict) or not mapping or len(mapping)>1000: raise ValueError('Declare bounded explicit product-to-category mapping')
    for product,labels in mapping.items():
        if not isinstance(product,str) or not product.strip() or not isinstance(labels,list) or not labels or any(not isinstance(label,str) or label not in categories for label in labels) or len(set(labels))!=len(labels):
            raise ValueError('Map each product to unique known category IDs; overlap between categories is allowed')
    left,left_meta,left_accounts=read_customers(left_source,config['left'],periods['left'],config,institutions,categories)
    right,right_meta,right_accounts=read_customers(right_source,config['right'],periods['right'],config,institutions,categories)
    if any(left_accounts[key][0]!=right_accounts[key][0] for key in left_accounts.keys()&right_accounts.keys()):
        raise ValueError('Stable account maps to a different customer across periods; resolve identity changes explicitly')
    scopes=[('bank','all')]+[('institution',key) for key in institutions]+[('peer_group',key) for key in dict.fromkeys(item['peer_group'] for item in entries)]
    rows=[]; coverage=[]; distributions=[]; changes=[]; lineage=[]
    def members(customers,kind,key):
        return {customer for customer,item in customers.items() if kind=='bank' or kind=='institution' and item['owner']==key or kind=='peer_group' and institutions[item['owner']]['peer_group']==key}
    def ratio(numerator,denominator):
        if denominator==0: return {'value':None,'rounded':False,'status':'undefined_zero_denominator'}
        return {**rounded(lambda:Decimal(numerator)/denominator,precision),'status':'calculated'}
    for kind,key in scopes:
        sets=[]
        for side,customers,metadata in [('left',left,left_meta),('right',right,right_meta)]:
            observed=members(customers,kind,key); holders={customer for customer in observed if customers[customer]['products']}
            denominator=len(observed if config['coverage_denominator']=='observed_customers' else holders)
            flags=[]
            if not denominator: flags.append('zero_coverage_denominator')
            if scope['completeness']=='unknown': flags.append('observation_completeness_unknown')
            rows.append({'scope_kind':kind,'scope_id':key,'period':periods[side],'customer_count':len(observed),'holding_customer_count':len(holders),'coverage_denominator':denominator,'quality_flags':flags})
            for category,label in categories.items():
                count=sum(category in customers[customer]['products'] for customer in observed)
                coverage.append({'scope_kind':kind,'scope_id':key,'period':periods[side],'category_id':category,'label':label,'customer_count':count,'denominator':denominator,'coverage':ratio(count,denominator)})
            for count in range(len(categories)+1):
                population=sum(len(customers[customer]['products'])==count for customer in observed)
                distributions.append({'scope_kind':kind,'scope_id':key,'period':periods[side],'held_category_count':count,'customer_count':population,'share_of_observed':ratio(population,len(observed))})
            for number in sorted({number for customer in observed for number in customers[customer]['source_rows']}):
                lineage.append({'scope_kind':kind,'scope_id':key,'period':periods[side],'file':metadata['file'],'sheet':metadata['sheet'],'source_row':number})
                if len(lineage)>60000: raise ValueError('Customer aggregate lineage exceeds 60000 rows; narrow input')
            sets.append(observed)
        previous,current=sets
        changes.append({'scope_kind':kind,'scope_id':key,'previous_customers':len(previous),'current_customers':len(current),
                        'continuing_observed':len(previous&current),'only_current_observed':len(current-previous),
                        'only_previous_observed':len(previous-current),'net_observed_change':len(current)-len(previous),
                        'residual':len(current)-len(previous)-len(current-previous)+len(previous-current)})
    migrations=sum(left[customer]['owner']!=right[customer]['owner'] for customer in left.keys()&right.keys())
    safe_config=config  # Configuration contains institution/product metadata, no customer ownership list.
    result={'status':'customer_structure','schema_version':'customer_structure_v1','requested_query':safe_config,
            'summary':rows,'coverage':coverage,'product_count_distribution':distributions,'changes':changes,'lineage':lineage,
            'sources':{'left':left_meta,'right':right_meta},'migration_count':migrations,
            'observation_counts':{'left_accounts':len(left_accounts),'right_accounts':len(right_accounts)},
            'policies':{'identity':'Stable anonymized customer key deduplicated across account/product observations; no exported customer/account lists',
                        'ownership':'Exclusive or consistent designated owner; all observed products attributed to that owner; no inferred primary relationship',
                        'additivity':'Institution and peer-group customer partitions are disjoint under owner policy; bank count is directly deduplicated, never summed across report levels',
                        'coverage':'Distinct customers holding category / explicit denominator; overlapping category ratios are not additive',
                        'distribution':'Number of distinct mapped categories with holding=true, not number of accounts; includes zero holdings among observed customers',
                        'change':'Set observations only; not acquisition/opening/loss/closure; owner migration is not bank-level entry/exit',
                        'privacy':'Aggregates and source coordinates only; no stable customer/account IDs or per-person decisions',
                        'precision':precision,'rounding':'ROUND_HALF_EVEN'}}
    result['management_wide']=management_wide(result)
    return result


def management_wide(result):
    config=result['requested_query']; scope=config['scope']
    counts=['observed_customers','holding_customers','coverage_denominator']
    category_metrics=[(json.dumps(['category_count',item['id']],ensure_ascii=False,separators=(',',':')),item) for item in config['categories']]
    coverage_metrics=[(json.dumps(['category_coverage',item['id']],ensure_ascii=False,separators=(',',':')),item) for item in config['categories']]
    definitions=[{'metric_id':key,'label':key,'unit':'customers','kind':'count','time_basis':'observed_snapshot','aggregation':'COUNT','additive_dimensions':['institution']} for key in counts]
    definitions += [{'metric_id':key,'label':item['label']+' held customers','unit':'customers','kind':'count','time_basis':'observed_snapshot','aggregation':'COUNT','additive_dimensions':['institution']} for key,item in category_metrics]
    definitions += [{'metric_id':key,'label':item['label']+' coverage','unit':'fraction','kind':'ratio','time_basis':'ratio','aggregation':'RATIO','numerator':category_metrics[index][0],'denominator':'coverage_denominator','scale':1} for index,(key,item) in enumerate(coverage_metrics)]
    roles=[('dimension',name) for name in ('period','institution','report_level','currency','version','peer_group','institution_level','observation_scope')]
    roles += [('metric',item['metric_id'],'actual') for item in definitions]+[('quality','flags')]
    identifiers=[column_id(*role) for role in roles]
    labels=[role[1] if role[0]!='metric' else next(item['label'] for item in definitions if item['metric_id']==role[1]) for role in roles]
    generated=[{'id':identifier,'column':identifier,'path':[{'type':'text','value':label}],'levels':[{'type':'text','value':label,'role':'customer_aggregate'}]} for identifier,label in zip(identifiers,labels)]
    output=[]; lineage=[]
    coverage_index={(row['scope_kind'],row['scope_id'],row['period'],row['category_id']):row for row in result['coverage']}
    source_index={}
    for row in result['lineage']: source_index.setdefault((row['scope_kind'],row['scope_id'],row['period']),[]).append(row)
    for index,row in enumerate(result['summary'],1):
        context=(row['scope_kind'],row['scope_id'],row['period'])
        institution=json.dumps([row['scope_kind'],row['scope_id']],ensure_ascii=False,separators=(',',':'))
        peer_group=next(item['peer_group'] for item in config['institutions'] if item['id']==row['scope_id']) if row['scope_kind']=='institution' else row['scope_id'] if row['scope_kind']=='peer_group' else 'not_applicable'
        values=[row['period'],institution,row['scope_kind'],'not_applicable',scope['version'],peer_group,scope['level'],scope['observation_scope']]
        cells=[{'type':'text','value':value} for value in values]
        values=[row['customer_count'],row['holding_customer_count'],row['coverage_denominator']]
        values += [coverage_index[context+(item['id'],)]['customer_count'] for _,item in category_metrics]
        values += [coverage_index[context+(item['id'],)]['coverage']['value'] for _,item in coverage_metrics]
        cells += [{'type':'blank' if value is None else 'number','value':value} for value in values]
        cells.append({'type':'text','value':json.dumps(row['quality_flags'],ensure_ascii=False)})
        sources=source_index.get(context,[])
        output.append({'row_id':index,'cells':cells,'source_rows':sorted({item['source_row'] for item in sources}),'quality_flags':row['quality_flags']})
        lineage += [{'output_row':index,'source_file':item['file'],'sheet':item['sheet'],'source_row':item['source_row'],'metric_id':'observed_customers'} for item in sources]
    return {'status':'management_wide','schema_version':'management_wide_v1','sheet':'aggregate customer observations','range':'configured input regions',
            'requested_query':config,'columns':identifiers,'headers':[{'type':'text','value':label} for label in labels],'generated_columns':generated,
            'rows':output,'output_rows':len(output),'metric_dictionary':definitions,'definition_version':scope['definition_version'],
            'report_level':'separate bank/institution/peer_group slices','lineage':lineage,
            'validation':[{'check':'ratios_recomputed','status':'pass','detail':'Deduplicated category customer counts divided by explicit customer denominator'},
                          {'check':'scope_additivity','status':'warning','detail':'Separate scope levels must not be summed; overlapping category ratios are not additive'}],
            'customer_policies':result['policies']}


def export_customers(result,output,readable=True):
    if result.get('schema_version')!='customer_structure_v1': raise ValueError('Unknown customer aggregate schema')
    report=Report(); config=result['requested_query']; scope=config['scope']
    context=[scope['level'],scope['version'],scope['definition_version'],scope['observation_scope'],config['coverage_denominator']]
    context_labels=['Level','Version','Definition version','Observation scope','Denominator policy']
    summary=[context_labels+['Scope kind','Scope ID','Period','Observed customers','Holding customers','Denominator','Quality flags']]
    for row in result['summary']:
        summary.append(context+[row[key] for key in ('scope_kind','scope_id','period','customer_count','holding_customer_count','coverage_denominator')]+[json.dumps(row['quality_flags'],ensure_ascii=False)])
    coverage=[['Scope kind','Scope ID','Period','Category','Label','Holding customers','Denominator','Coverage','Rounded','Status']]
    for row in result['coverage']:
        coverage.append([row[key] for key in ('scope_kind','scope_id','period','category_id','label','customer_count','denominator')]+[report.number(row['coverage']['value'],'coverage'),row['coverage']['rounded'],row['coverage']['status']])
    distribution=[['Scope kind','Scope ID','Period','Distinct held categories','Customers','Share of observed','Rounded','Status']]
    for row in result['product_count_distribution']:
        distribution.append([row[key] for key in ('scope_kind','scope_id','period','held_category_count','customer_count')]+[report.number(row['share_of_observed']['value'],'distribution share'),row['share_of_observed']['rounded'],row['share_of_observed']['status']])
    fields=['scope_kind','scope_id','previous_customers','current_customers','continuing_observed','only_current_observed','only_previous_observed','net_observed_change','residual']
    changes=[fields]+[[row[key] for key in fields] for row in result['changes']]
    lineage_fields=['scope_kind','scope_id','period','file','sheet','source_row']
    lineage=[lineage_fields]+[[row[key] for key in lineage_fields] for row in result['lineage']]
    policies=[['Rule','Definition']]+[[key,str(value)] for key,value in result['policies'].items()]+[['Owner migrations among continuing bank observations',result['migration_count']]]
    quality=[['Scope kind','Scope ID','Period','Flag']]+[[row['scope_kind'],row['scope_id'],row['period'],flag] for row in result['summary'] for flag in row['quality_flags']]
    snapshot=json.dumps(result,ensure_ascii=False,allow_nan=False,default=str)
    if len(snapshot.encode('utf-8'))>MAX_INPUT_BYTES: raise ValueError('Customer aggregate snapshot exceeds 16 MiB')
    metadata=[['Section','Part','JSON/text']]+[['customer_aggregate_json',index+1,snapshot[start:start+30000]] for index,start in enumerate(range(0,len(snapshot),30000))]
    wide,_,_=management_table(result['management_wide'],report,source_columns(result['management_wide']))
    sheets=[report.sheet('Wide',wide)]+[report.sheet(name,rows) for name,rows in [('Customers',summary),('Coverage',coverage),('ProductCounts',distribution),('Changes',changes),('Lineage',lineage),('Definitions',policies),('Quality',quality),('Metadata',metadata)]]
    sheets.append(report.sheet('Warnings',[['Location','Issue','Original']]+report.warnings))
    appearance=None if not readable else {'preset':'readable','sheets':{sheet['name']:{'header_rows':'1','column_widths':{'A':24}} for sheet in sheets}}
    book=Workbook.create({'sheets':sheets},appearance=appearance); book.save(output)
    return {'status':'exported','output_written':True,'output':str(output),'formulas_created':0,'total_cells':report.count}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('left'); parser.add_argument('right'); parser.add_argument('--query',required=True); parser.add_argument('--xlsx'); parser.add_argument('--plain',action='store_true')
    args=parser.parse_args()
    try:
        result=analyze(args.left,args.right,read_spec(args.query,max_bytes=65536))
        if args.xlsx: result['export']=export_customers(result,args.xlsx,not args.plain)
        print(json.dumps(result,ensure_ascii=False,default=str)); return 0
    except (ValueError,OSError,LookupError,DecimalException,ET.ParseError,zipfile.BadZipFile) as exc:
        print(json.dumps({'status':'failed','message':str(exc),'output_written':False},ensure_ascii=False)); return 1


if __name__=='__main__': sys.exit(main())
