#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Configured management long-to-wide reports; no bank-specific regulatory definitions."""
import argparse
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
from xlsx_advanced import exact,rounded,precision_option,parse_month

DIMENSIONS={'period','institution','line','currency','version','product','region','channel'}
PROTECTED={'period','currency','version'}


def column_id(*parts):
    return json.dumps(list(parts),ensure_ascii=False,separators=(',',':'))


def required_text(value,location):
    if not isinstance(value,str) or not value.strip() or len(value)>128:
        raise ValueError('Expected 1..128 character text at '+location)
    return value


def metric_dictionary(items,dimensions,dropped):
    if not isinstance(items,list) or not 1<=len(items)<=50:
        raise ValueError('metrics must contain 1..50 explicit definitions')
    definitions={}
    allowed={'metric_id','label','unit','kind','time_basis','aggregation','additive_dimensions','reduction_dimensions',
             'numerator','denominator','scale','allow_negative_denominator'}
    for item in items:
        if not isinstance(item,dict) or set(item)-allowed:
            raise ValueError('Unknown metric dictionary fields')
        for field in ('metric_id','label','unit','kind','time_basis','aggregation'):
            required_text(item.get(field),'metric/'+field)
        identifier=item['metric_id']
        if identifier in definitions:
            raise ValueError('Duplicate metric_id: '+identifier)
        for field in ('additive_dimensions','reduction_dimensions'):
            values=item.get(field,[])
            if not isinstance(values,list) or any(not isinstance(value,str) or value not in dimensions for value in values) or len(values)!=len(set(values)):
                raise ValueError('Invalid '+field+' for '+identifier)
            if set(values)&PROTECTED:
                raise ValueError('Period, currency and version cannot be declared additive/reduced')
        if item['kind']=='ratio':
            if item['aggregation']!='RATIO' or item['time_basis']!='ratio' or item.get('additive_dimensions'):
                raise ValueError('Ratios must be RATIO/ratio and are never directly additive')
            required_text(item.get('numerator'),'numerator')
            required_text(item.get('denominator'),'denominator')
            scale=item.get('scale')
            if isinstance(scale,bool) or not isinstance(scale,(int,Decimal)) or not Decimal(scale).is_finite() or scale<=0:
                raise ValueError('Ratio requires an explicit positive numeric scale')
            if type(item.get('allow_negative_denominator',False)) is not bool:
                raise ValueError('allow_negative_denominator must be boolean')
        else:
            if item['kind'] not in ('stock','flow') or item['aggregation'] not in ('SUM','COUNT','MIN','MAX'):
                raise ValueError('Base metrics need stock/flow and SUM/COUNT/MIN/MAX')
            if item['kind']=='stock' and item['time_basis']!='point_in_time' or item['kind']=='flow' and item['time_basis'] not in ('period_flow','ytd'):
                raise ValueError('Stock/flow time basis is inconsistent')
            if item['aggregation']=='COUNT' and item['unit']!='records':
                raise ValueError('COUNT counts numeric fact observations only; unit must be records')
            permission='additive_dimensions' if item['aggregation'] in ('SUM','COUNT') else 'reduction_dimensions'
            if not dropped<=set(item.get(permission,[])):
                raise ValueError('Metric '+identifier+' does not permit reduction across '+','.join(sorted(dropped)))
            if set(item)&{'numerator','denominator','scale','allow_negative_denominator'}:
                raise ValueError('Base metric cannot contain ratio definitions')
        definitions[identifier]=item
    for identifier,item in definitions.items():
        if item['kind']=='ratio':
            numerator=definitions.get(item['numerator']); denominator=definitions.get(item['denominator'])
            if numerator is None or denominator is None or numerator['kind']=='ratio' or denominator['kind']=='ratio':
                raise ValueError('Ratios require two defined base metrics, not ratio chains')
            if numerator['unit']!=denominator['unit'] or numerator['time_basis']!=denominator['time_basis'] or numerator['kind']!=denominator['kind']:
                raise ValueError('Mixed-unit/basis ratio components require a separate unsupported business definition')
    return definitions


def dimension_text(name,value,version,location):
    if value[0]!='text' or not value[1].strip():
        raise ValueError('Management dimensions must be nonempty literal text at '+location)
    text=value[1]
    if name=='period':
        parse_month(value,location)
    if name=='currency' and not re.fullmatch(r'[A-Z]{3}',text):
        raise ValueError('Currency must be an explicit three-letter code at '+location)
    if name=='version' and text!=version:
        raise ValueError('Definition version mismatch at '+location)
    return text


def target_index(targets,row_dimensions,definitions,version,report_level,hierarchy=None):
    if not isinstance(targets,list) or len(targets)>2000:
        raise ValueError('targets must be a list of at most 2000 records')
    output={}
    allowed={'dimensions','metric_id','unit','kind','time_basis','definition_version','report_level','mode','value'}
    for index,target in enumerate(targets):
        if not isinstance(target,dict) or set(target)!=allowed or not isinstance(target['dimensions'],dict) or set(target['dimensions'])!=set(row_dimensions):
            raise ValueError('Target fields/grain must match the complete report row dimensions')
        identifier=target['metric_id']
        if not isinstance(identifier,str) or identifier not in definitions:
            raise ValueError('Target names an unknown metric')
        definition=definitions[identifier]
        expected_level=report_level
        if hierarchy is not None:
            node=hierarchy.node(target['dimensions'].get('institution'),target['dimensions'].get('period'))
            expected_level=node['level']
            if expected_level not in hierarchy.report_levels:
                raise ValueError('Target level is not selected for this hierarchy report')
        if any(target[field]!=definition[field] for field in ('unit','kind','time_basis')) or target['definition_version']!=version or target['report_level']!=expected_level:
            raise ValueError('Target type/unit/basis/version/report level does not match actual metric')
        if target['mode'] not in ('at_least','at_most'):
            raise ValueError('Target mode must be explicit at_least or at_most')
        value=target['value']
        if isinstance(value,bool) or not isinstance(value,(int,Decimal)) or not Decimal(value).is_finite():
            raise ValueError('Target value must be a finite JSON number')
        values=tuple(dimension_text(name,('text',target['dimensions'][name]) if isinstance(target['dimensions'][name],str) else ('invalid',None),version,'target/'+name) for name in row_dimensions)
        marker=(values,identifier)
        if marker in output:
            raise ValueError('Duplicate target at configured grain')
        output[marker]={**target,'value':Decimal(value),'target_index':index}
    return output


def build(source,sheet,config,**options):
    allowed={'dimensions','dimension_labels','row_dimensions','fact_grain','metric_column','value_column','unit_column',
             'basis_column','record_scope','definition_version','period_format','metrics','filters','missing','targets','precision','hierarchy'}
    if not isinstance(config,dict) or set(config)-allowed:
        raise ValueError('Unknown management configuration keys')
    dimensions=config.get('dimensions')
    if not isinstance(dimensions,dict) or not PROTECTED|{'institution'}<=set(dimensions) or set(dimensions)-DIMENSIONS:
        raise ValueError('Dimensions require period/institution/currency/version and only supported management dimensions')
    row_dimensions=config.get('row_dimensions')
    if not isinstance(row_dimensions,list) or any(not isinstance(name,str) or name not in dimensions for name in row_dimensions) or len(row_dimensions)!=len(set(row_dimensions)) or not PROTECTED<=set(row_dimensions):
        raise ValueError('Report rows must retain period/currency/version with unique declared dimensions')
    grain=config.get('fact_grain')
    if not isinstance(grain,list) or any(not isinstance(name,str) for name in grain) or len(grain)!=len(set(grain)) or set(grain)!=set(dimensions)|{'metric_id'}:
        raise ValueError('fact_grain must explicitly contain every input dimension plus metric_id')
    version=required_text(config.get('definition_version'),'definition_version')
    if config.get('period_format')!='YYYY-MM':
        raise ValueError('Management period_format must be explicit YYYY-MM')
    missing=config.get('missing','propagate')
    if missing not in ('propagate','reject'):
        raise ValueError('missing must be propagate or reject; no zero fill or silent omit')
    precision=precision_option(config)
    scope=config.get('record_scope')
    if not isinstance(scope,dict) or set(scope)!={'column','value','report_level'}:
        raise ValueError('record_scope requires a source column, literal scope value and explicit report_level')
    required_text(scope['value'],'record_scope/value')
    report_level=required_text(scope['report_level'],'record_scope/report_level')
    labels=config.get('dimension_labels',{})
    if not isinstance(labels,dict) or set(labels)-set(dimensions) or any(not isinstance(value,str) or not value for value in labels.values()):
        raise ValueError('Invalid dimension_labels')
    hierarchy=None
    dropped=set(dimensions)-set(row_dimensions)
    if 'hierarchy' in config:
        from xlsx_hierarchy import Hierarchy
        hierarchy=Hierarchy(config['hierarchy'])
        if 'institution' not in row_dimensions or report_level!=hierarchy.levels[-1]:
            raise ValueError('Hierarchy reports retain institution and declare deepest source report_level')
        if any(level!=hierarchy.levels[-1] for level in hierarchy.report_levels):
            dropped.add('institution')
    definitions=metric_dictionary(config.get('metrics'),dimensions,dropped)
    targets=target_index(config.get('targets',[]),row_dimensions,definitions,version,report_level,hierarchy)
    targeted={identifier for _,identifier in targets}
    columns,_,rows,metadata=region(source,sheet,**options)
    schema=metadata.get('header_schema')
    dim_columns={name:resolve_column(selector,columns,schema) for name,selector in dimensions.items()}
    fields={name:resolve_column(config.get(name),columns,schema) for name in ('metric_column','value_column','unit_column','basis_column')}
    scope_column=resolve_column(scope['column'],columns,schema)
    if len(set(list(dim_columns.values())+list(fields.values())+[scope_column]))!=len(dim_columns)+len(fields)+1:
        raise ValueError('Dimension/metric/value/unit/basis/scope columns must be distinct')
    _,_,_,filters=validate_query({'filters':config.get('filters',[])},columns,schema)
    selected=[record for record in rows if matches(record[1],filters)]
    if hierarchy is not None and len(selected)*len(hierarchy.report_levels)>20000:
        raise ValueError('Hierarchy expansion exceeds 20000 lineage rows; narrow filters/levels')
    buckets={}; seen={}; lineage=[]
    for number,row in selected:
        if row[columns.index(scope_column)]!=('text',scope['value']):
            raise ValueError('Mixed/unknown record scope at row '+str(number)+'; filter explicitly to one trusted leaf management scope')
        dims={name:dimension_text(name,row[columns.index(column)],version,column+str(number)) for name,column in dim_columns.items()}
        metric=row[columns.index(fields['metric_column'])]
        if metric[0]!='text' or metric[1] not in definitions:
            raise ValueError('Unknown/nontext metric ID at row '+str(number))
        identifier=metric[1]; definition=definitions[identifier]
        if definition['kind']=='ratio':
            raise ValueError('Raw ratio facts are refused; supply numerator/denominator facts')
        if row[columns.index(fields['unit_column'])]!=('text',definition['unit']) or row[columns.index(fields['basis_column'])]!=('text',definition['time_basis']):
            raise ValueError('Fact unit/time basis mismatches metric dictionary at row '+str(number))
        marker=tuple(dims[name] if name!='metric_id' else identifier for name in grain)
        if marker in seen:
            raise ValueError('Duplicate fact grain at rows %d and %d' % (seen[marker],number))
        seen[marker]=number
        value=row[columns.index(fields['value_column'])]
        if is_blank(value):
            if missing=='reject':
                raise ValueError('Missing fact value at row '+str(number))
            numeric=None
        elif value[0]!='number':
            raise ValueError('Fact values must be numeric without text coercion at row '+str(number))
        else:
            numeric=value[1]
        nodes=hierarchy.ancestors(dims['institution'],dims['period']) if hierarchy is not None else [None]
        for node in nodes:
            report_dims=dict(dims)
            if node is not None:
                report_dims['institution']=node['id']
            key=tuple(report_dims[name] for name in row_dimensions)
            if key not in buckets:
                if len(buckets)>=10000:
                    raise ValueError('More than 10000 report groups')
                buckets[key]={'metrics':{},'sources':[],'metric_leaves':{}}
            bucket=buckets[key]
            bucket['metrics'].setdefault(identifier,[]).append(numeric)
            bucket['metric_leaves'].setdefault(identifier,set()).add(dims['institution'])
            bucket['sources'].append(number)
            lineage.append({'group':key,'source_file':str(source),'sheet':sheet,'source_row':number,'metric_id':identifier})
    output_columns=[]; generated=[]
    def add(parts,path):
        identifier=column_id(*parts)
        output_columns.append(identifier)
        generated.append({'column':identifier,'id':identifier,'path':[{'type':'text','value':value} for value in path],
                          'levels':[{'type':'text','value':value,'role':'configured_report'} for value in path]})
    for name in row_dimensions:
        add(('dimension',name),(labels.get(name,name),))
        if hierarchy is not None and name=='institution':
            add(('dimension','institution_name'),('机构名称',))
    add(('dimension','report_level'),('报告层级',))
    for identifier,definition in definitions.items():
        label=definition['label']+' ['+definition['unit']+']'
        add(('metric',identifier,'actual'),(label,'实际'))
        if identifier in targeted:
            for field,label_part in (('target','目标'),('achievement','完成率'),('gap','未达目标缺口')):
                add(('metric',identifier,field),(definition['label'] if field=='achievement' else label,
                                                          label_part+' [倍]' if field=='achievement' else label_part))
    add(('quality','flags'),('质量标记',))
    if len(output_columns)>16384 or len(buckets)*len(output_columns)>MAX_CELLS:
        raise ValueError('Wide report exceeds output column/cell budget')
    output=[]; matched_targets=set(); group_numbers={}
    for row_id,(key,bucket) in enumerate(buckets.items(),1):
        flags=[]; actual={}
        for identifier,definition in definitions.items():
            if definition['kind']=='ratio':
                continue
            values=bucket['metrics'].get(identifier,[])
            missing_leaves=set()
            if hierarchy is not None:
                ancestor=key[row_dimensions.index('institution')]
                missing_leaves=hierarchy.descendant_leaves[ancestor]-bucket['metric_leaves'].get(identifier,set())
                if missing_leaves:
                    flags.append(('incomplete_leaf_metric:' if hierarchy.coverage=='complete' else 'observed_leaf_metric:')+identifier)
            if not values or any(value is None for value in values) or missing_leaves and hierarchy.coverage=='complete':
                if missing=='reject':
                    raise ValueError('Missing metric/fact values for '+identifier+' in report group '+str(row_id))
                actual[identifier]=None
                flags.append(('missing_metric:' if not values else 'missing_values:')+identifier)
                continue
            method=definition['aggregation']
            actual[identifier]=exact(lambda:sum(values,Decimal(0))) if method=='SUM' else Decimal(len(values)) if method=='COUNT' else min(values) if method=='MIN' else max(values)
        for identifier,definition in definitions.items():
            if definition['kind']!='ratio':
                continue
            numerator=actual[definition['numerator']]; denominator=actual[definition['denominator']]
            value=None
            if numerator is None or denominator is None:
                flags.append('missing_components:'+identifier)
            elif denominator==0:
                flags.append('zero_denominator:'+identifier)
            elif denominator<0 and not definition.get('allow_negative_denominator',False):
                flags.append('negative_denominator:'+identifier)
            else:
                scaled=exact(lambda:numerator*Decimal(definition['scale']))
                ratio=rounded(lambda:scaled/denominator,precision)
                value=ratio['value']
                if ratio['rounded']:
                    flags.append('rounded_ratio:'+identifier)
                if numerator<0:
                    flags.append('negative_numerator:'+identifier)
                if denominator<0:
                    flags.append('negative_denominator_allowed:'+identifier)
            actual[identifier]=value
        current_level=report_level
        institution_node=None
        if hierarchy is not None:
            institution_node=hierarchy.nodes[key[row_dimensions.index('institution')]]
            current_level=institution_node['level']
        values=[]
        for name,value in zip(row_dimensions,key):
            values.append(('text',value))
            if institution_node is not None and name=='institution':
                values.append(('text',institution_node['name']))
        values.append(('text',current_level))
        for identifier,definition in definitions.items():
            value=actual[identifier]
            values.append(('blank',None) if value is None else ('number',value))
            if identifier not in targeted:
                continue
            target=targets.get((key,identifier)); attainment=gap=None
            target_value=None if target is None else target['value']
            if target is None:
                flags.append('missing_target:'+identifier)
            else:
                matched_targets.add((key,identifier))
                if value is None:
                    flags.append('missing_actual_for_target:'+identifier)
                else:
                    signed=exact(lambda:target_value-value if target['mode']=='at_least' else value-target_value)
                    gap=max(signed,Decimal(0))
                if target_value<=0:
                    flags.append(('zero_target:' if target_value==0 else 'negative_target:')+identifier)
                elif value is not None:
                    if target['mode']=='at_most' and value<=0:
                        flags.append('nonpositive_actual_for_inverse_achievement:'+identifier)
                    else:
                        completion=rounded(lambda:value/target_value if target['mode']=='at_least' else target_value/value,precision)
                        attainment=completion['value']
                        if completion['rounded']:
                            flags.append('rounded_achievement:'+identifier)
                        if value<0:
                            flags.append('negative_actual_for_target:'+identifier)
            values += [('blank',None) if item is None else ('number',item) for item in (target_value,attainment,gap)]
        values.append(('text',json.dumps(flags,ensure_ascii=False)))
        output.append({'row_id':row_id,'cells':[{'type':kind,'value':value} for kind,value in values],
                       'source_rows':bucket['sources'],'quality_flags':flags})
        group_numbers[key]=row_id
    for item in lineage:
        item['output_row']=group_numbers[item.pop('group')]
    orphan_targets=[target['target_index'] for marker,target in targets.items() if marker not in matched_targets]
    validation=[{'check':name,'status':'pass','detail':detail} for name,detail in (
        ('fact_grain_unique',str(len(seen))+' unique selected facts'),
        ('record_scope',scope['value']+'; upstream leaf-management marker trusted, no hierarchy inference'),
        ('protected_dimensions','period/currency/version retained; no cross-period or cross-currency sum'),
        ('definition_version',version),('unit_and_time_basis','fact metadata checked against dictionary'),
        ('aggregation_permissions','dropped dimensions explicitly allowed by each base metric'),
        ('ratios_recomputed','base components aggregated first; no averaging raw ratios'))]
    if orphan_targets:
        validation.append({'check':'unmatched_targets','status':'warning','detail':json.dumps(orphan_targets)})
    if any(row['quality_flags'] for row in output):
        validation.append({'check':'row_quality_flags','status':'warning','detail':str(sum(bool(row['quality_flags']) for row in output))+' rows have diagnostics'})
    metadata.pop('precision',None)
    return {'status':'management_wide','schema_version':'management_wide_v1',**metadata,
            'source_file':str(source),'requested_query':config,'columns':output_columns,
            'headers':[{'type':'text','value':' / '.join(part['value'] for part in item['path'])} for item in generated],
            'generated_columns':generated,'rows':output,'output_rows':len(output),'lineage':lineage,
            'metric_dictionary':list(definitions.values()),'definition_version':version,'report_level':report_level,
            'hierarchy_summary':None if hierarchy is None else hierarchy.summary(),
            'selected_facts':len(selected),'filtered_out_rows':len(rows)-len(selected),'validation':validation,
            'target_policy':{'at_least':'actual/positive target; gap=max(target-actual,0)',
                             'at_most':'positive target/positive actual; gap=max(actual-target,0)',
                             'nonpositive_targets':'achievement null with explicit zero/negative diagnostic; no success classification'},
            'arithmetic':{'exact_precision_limit':1000,'result_precision':precision,'intermediate_precision':precision+16,'rounding':'ROUND_HALF_EVEN'},
            'unsupported':(['automatic historical institution restatement'] if hierarchy is not None else ['institution hierarchy rollup'])+['customer distinct counts','daily or weighted averages','fiscal/YTD time accumulation','regulatory reporting or thresholds']}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input'); parser.add_argument('--sheet',required=True); parser.add_argument('--query',required=True)
    parser.add_argument('--xlsx',help='Optional new readable XLSX output')
    parser.add_argument('--range',dest='cell_range'); parser.add_argument('--header-rows')
    parser.add_argument('--no-header',action='store_true'); parser.add_argument('--header-exclude',nargs='*',type=int,default=[])
    parser.add_argument('--include-spanning-headings',action='store_true'); parser.add_argument('--formulas',choices=('reject','cached'),default='reject')
    args=parser.parse_args()
    try:
        result=build(args.input,args.sheet,read_query(args.query),cell_range=args.cell_range,header=not args.no_header,
                     header_rows=args.header_rows,header_exclude=args.header_exclude,
                     include_spanning_headings=args.include_spanning_headings,formulas=args.formulas)
        if args.xlsx:
            from xlsx_result_export import export_result
            result['export']=export_result(result,args.xlsx)
        print(json.dumps(result,ensure_ascii=False,default=str))
        return 0
    except (ValueError,OSError,LookupError,DecimalException,ET.ParseError,zipfile.BadZipFile) as exc:
        print(json.dumps({'status':'failed','message':str(exc),'output_written':False},ensure_ascii=False))
        return 1


if __name__=='__main__':
    sys.exit(main())
