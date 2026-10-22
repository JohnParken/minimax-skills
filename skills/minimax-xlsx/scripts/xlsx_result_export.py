#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Export known analysis-result schemas as data-only readable XLSX reports."""
import argparse
from decimal import Decimal, DecimalException
import json
from pathlib import Path
import re
import sys
from xml.etree import ElementTree as ET
import zipfile

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook, MAX_CELLS, read_spec
from xlsx_tabular import column_name

MAX_INPUT_BYTES = 16*1024*1024


class Report:
    def __init__(self):
        self.warnings=[]
        self.count=0

    def number(self,value,location):
        if value is None:
            return None
        if isinstance(value,bool) or not isinstance(value,(str,int,Decimal)) or not re.fullmatch(r'[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?',str(value)):
            raise ValueError('Expected an explicit numeric value at '+location)
        if len(str(value))>32767:
            raise ValueError('Numeric representation exceeds cell text limit at '+location)
        number=Decimal(value)
        if not number.is_finite():
            raise ValueError('Nonfinite number at '+location)
        digits=list(number.as_tuple().digits)
        while digits and digits[-1]==0:
            digits.pop()
        if len(digits)>15 or number and not -307<=number.adjusted()<=307:
            self.warnings.append([location,'Exact numeric value exported as text; Excel numeric precision/range limit',str(value)])
            return str(value)
        return number

    def typed(self,item,location):
        if not isinstance(item,dict) or set(item)!={'type','value'}:
            raise ValueError('Expected a typed {type,value} cell at '+location)
        kind,value=item['type'],item['value']
        if kind=='number':
            if value is None:
                raise ValueError('Numeric typed cells cannot be null at '+location)
            return self.number(value,location)
        if kind=='blank' and value is None:
            return None
        if kind=='boolean' and type(value) is bool:
            return value
        if kind in ('text','date_text') and isinstance(value,str):
            return value
        raise ValueError('Invalid typed cell at '+location)

    def sheet(self,name,rows):
        cells={}
        for row,values in enumerate(rows,1):
            if row>1048576 or len(values)>16384:
                raise ValueError('Report exceeds Excel row/column bounds')
            for col,value in enumerate(values,1):
                self.count+=1
                if self.count>MAX_CELLS:
                    raise ValueError('Report exceeds 200000 total cells including provenance/warnings')
                cells[column_name(col)+str(row)]={'value':value}
        return {'name':name,'cells':cells}


def count(value,location):
    if type(value) is not int or value<0:
        raise ValueError('Expected a nonnegative integer at '+location)
    return value


def source_columns(result):
    schema=result.get('header_schema')
    entries=[]
    if schema is None:
        headers=result.get('column_headers',[])
        if not isinstance(headers,list):
            raise ValueError('Invalid source column headers')
        for item in headers:
            if not isinstance(item,dict) or not isinstance(item.get('column'),str) or 'type' not in item or 'value' not in item:
                raise ValueError('Invalid source column header')
            entries.append({'column':item['column'],'id':'col:'+item['column'],
                            'path':[{'type':item['type'],'value':item['value']}],'levels':[item]})
    else:
        if not isinstance(schema,dict) or not isinstance(schema.get('columns'),list):
            raise ValueError('Invalid header_schema')
        entries.extend(schema['columns'])
    generated=result.get('generated_columns',[])
    if not isinstance(generated,list):
        raise ValueError('Invalid generated columns')
    entries.extend(generated)
    mapping={}
    for item in entries:
        if not isinstance(item,dict) or not isinstance(item.get('column'),str) or not isinstance(item.get('path'),list):
            raise ValueError('Invalid structured header column')
        if any(not isinstance(part,dict) or 'type' not in part or 'value' not in part for part in item['path']) or ('levels' in item and not isinstance(item['levels'],list)):
            raise ValueError('Invalid structured header path/levels')
        if item['column'] in mapping:
            raise ValueError('Duplicate structured header column')
        mapping[item['column']]=item
    return mapping


def columns_table(mapping,identifiers,headers,report):
    rows=[['Column ID','Source column','Level','Type','Label','Source coordinate','Merge','Role']]
    for index,identifier in enumerate(identifiers):
        item=mapping.get(identifier,{})
        levels=item.get('levels') or item.get('path') or ([headers[index]] if headers is not None else [])
        if not levels:
            levels=[{'type':'blank','value':None}]
        for level,part in enumerate(levels,1):
            if not isinstance(part,dict) or 'type' not in part or 'value' not in part:
                raise ValueError('Invalid header path component')
            value=report.typed({'type':part['type'],'value':part['value']},'header/'+identifier+'/'+str(level))
            rows.append([item.get('id','col:'+identifier),identifier,level,part['type'],value,
                         part.get('coordinate'),part.get('merge'),part.get('role','label')])
    return rows


def selected_table(result,report,mapping):
    columns=result.get('columns')
    rows=result.get('rows')
    headers=result.get('headers')
    if not isinstance(columns,list) or not columns or any(not isinstance(col,str) or not col for col in columns) or len(columns)!=len(set(columns)) or not isinstance(rows,list):
        raise ValueError('Invalid selected-table columns/rows')
    if headers is not None and (not isinstance(headers,list) or len(headers)!=len(columns)):
        raise ValueError('Selected header count differs from columns')
    if headers is not None and any(not isinstance(item,dict) or set(item)!={'type','value'} for item in headers):
        raise ValueError('Invalid selected header cells')
    if count(result.get('output_rows'),'output_rows')!=len(rows):
        raise ValueError('Selected output_rows differs from actual rows')
    headings=[]
    for index,col in enumerate(columns):
        path=mapping.get(col,{}).get('path')
        # Joined labels are display only; the structured path is retained separately.
        label=' / '.join(str(part.get('value','')) for part in path) if path else str(headers[index].get('value','')) if headers else col
        headings.append(col+' · '+label)
    output=[['Source row']+headings]
    for index,row in enumerate(rows):
        if not isinstance(row,dict) or set(row)!={'source_row','cells'} or not isinstance(row['cells'],list) or len(row['cells'])!=len(columns):
            raise ValueError('Invalid selected row schema')
        number=count(row['source_row'],'source_row')
        if not number:
            raise ValueError('Source rows must be positive')
        output.append([number]+[report.typed(cell,'rows/'+str(index)+'/cells/'+str(col)) for col,cell in enumerate(row['cells'])])
    return output,columns,headers


def summarized_table(result,report,mapping):
    groups=result.get('groups')
    keys=result.get('group_by')
    metric=result.get('measure')
    if not isinstance(groups,list) or not isinstance(keys,list) or any(not isinstance(key,str) for key in keys) or len(keys)!=len(set(keys)) or metric is not None and not isinstance(metric,str):
        raise ValueError('Invalid summarized schema')
    output=[[key+' · group key' for key in keys]+['Metric column','Group rows','Blank measure','Ignored nonnumeric','Numeric count','Sum','Min','Max','Mean','Mean rounded']]
    for index,group in enumerate(groups):
        if not isinstance(group,dict) or set(group)!={'keys','rows','blank_measure','ignored_non_numeric','numeric'} or not isinstance(group['keys'],list) or len(group['keys'])!=len(keys):
            raise ValueError('Invalid summary group')
        values=[]
        for key,item in zip(keys,group['keys']):
            if not isinstance(item,dict) or set(item)!={'column','type','value'} or item['column']!=key:
                raise ValueError('Group key order/schema mismatch')
            values.append(report.typed({'type':item['type'],'value':item['value']},'groups/'+str(index)+'/'+key))
        values += [metric,count(group['rows'],'group rows'),count(group['blank_measure'],'blank measure'),count(group['ignored_non_numeric'],'ignored nonnumeric')]
        numeric=group['numeric']
        if (numeric is None)!=(metric is None):
            raise ValueError('Summary measure/statistics mismatch')
        if numeric is None:
            if metric is not None:
                raise ValueError('Summary metric lacks numeric statistics')
            values += [None]*6
        else:
            if not isinstance(numeric,dict) or set(numeric)!={'count','sum','min','max','mean','mean_rounded'} or type(numeric['mean_rounded']) is not bool:
                raise ValueError('Invalid summary numeric statistics')
            values += [count(numeric['count'],'numeric count')]+[report.number(numeric[key],'groups/'+str(index)+'/'+key) for key in ('sum','min','max','mean')]+[numeric['mean_rounded']]
        output.append(values)
    report.warnings.append(['source lineage','Aggregated JSON does not contain contributing source-row IDs; source range/count and available query are retained',None])
    return output,keys+([metric] if metric is not None and metric not in keys else []),None


def described_table(result,report,mapping):
    columns=result.get('columns')
    if not isinstance(columns,list) or not columns:
        raise ValueError('Invalid described columns')
    output=[['Metric column','Statistic','Status','Value','Rounded','Probability','ddof','Numeric count','Missing omitted','Nonnumeric omitted']]
    identifiers=[]
    for column in columns:
        if not isinstance(column,dict) or not isinstance(column.get('column'),str) or column['column'] in identifiers:
            raise ValueError('Invalid described column entry')
        identifier=column['column']; identifiers.append(identifier)
        numeric_count=count(column.get('count'),'numeric count')
        missing=count(column.get('missing_omitted'),'missing omitted')
        ignored=count(column.get('non_numeric_omitted'),'nonnumeric omitted')
        def append(name,metric,status,probability=None,ddof=None):
            if metric is not None and (not isinstance(metric,dict) or set(metric)!={'value','rounded'} or type(metric['rounded']) is not bool):
                raise ValueError('Invalid descriptive metric object')
            if metric is not None and metric['value'] is None:
                raise ValueError('Descriptive numeric metric cannot be null inside a value object')
            output.append([identifier,name,status,None if metric is None else report.number(metric['value'],identifier+'/'+name),
                           None if metric is None else metric['rounded'],probability,ddof,numeric_count,missing,ignored])
        status=column.get('status')
        if status not in ('ok','insufficient_numeric_samples'):
            raise ValueError('Invalid descriptive status')
        for name in ('mean','median'):
            if name not in column:
                raise ValueError('Missing descriptive metric')
            append(name,column[name],status)
        quantiles=column.get('quantiles')
        variances=column.get('variance')
        if not isinstance(quantiles,list) or len(quantiles)>21 or not isinstance(variances,dict) or set(variances)-{'sample','population'}:
            raise ValueError('Invalid quantile/variance schema')
        for metric in quantiles:
            if not isinstance(metric,dict) or set(metric)!={'probability','value','rounded'}:
                raise ValueError('Invalid quantile record')
            append('quantile',{'value':metric['value'],'rounded':metric['rounded']},status,report.number(metric['probability'],'probability'))
        for mode,metrics in variances.items():
            if not isinstance(metrics,dict) or set(metrics)!={'status','ddof','variance','standard_deviation'} or metrics['status'] not in ('ok','insufficient_numeric_samples') or type(metrics['ddof']) is not int or metrics['ddof']!=(1 if mode=='sample' else 0):
                raise ValueError('Invalid variance record')
            for name in ('variance','standard_deviation'):
                append(mode+'_'+name,metrics[name],metrics['status'],ddof=metrics['ddof'])
    report.warnings.append(['source lineage','Descriptive JSON does not contain contributing source-row IDs; source range/count and available query are retained',None])
    return output,identifiers,None


def management_table(result,report,mapping):
    if result.get('schema_version')!='management_wide_v1' or not isinstance(result.get('rows'),list) or not isinstance(result.get('columns'),list):
        raise ValueError('Unknown management-wide schema version')
    for identifier in result.get('columns',[]):
        parts=json.loads(identifier) if isinstance(identifier,str) else None
        if not isinstance(parts,list) or not all(isinstance(part,str) for part in parts) or not (
            len(parts)==2 and parts[0] in ('dimension','quality') or
            len(parts)==3 and parts[0]=='metric' and parts[2] in ('actual','target','achievement','gap')):
            raise ValueError('Management column IDs must be reversible known role arrays')
    translated=[]
    for index,row in enumerate(result['rows'],1):
        if not isinstance(row,dict) or set(row)!={'row_id','cells','source_rows','quality_flags'} or row['row_id']!=index or not isinstance(row['source_rows'],list) or not isinstance(row['quality_flags'],list):
            raise ValueError('Invalid management-wide row')
        if any(type(number) is not int or number<1 for number in row['source_rows']) or any(not isinstance(flag,str) for flag in row['quality_flags']):
            raise ValueError('Invalid management-wide lineage/quality flags')
        translated.append({'source_row':index,'cells':row['cells']})
        if row['quality_flags']:
            report.warnings.append(['report row '+str(index),'Management quality diagnostics',json.dumps(row['quality_flags'],ensure_ascii=False)])
    proxy={**result,'rows':translated}
    table,identifiers,headers=selected_table(proxy,report,mapping)
    # Display clean management labels; reversible IDs remain in Columns/Metadata.
    table=[row[1:] for row in table]
    table[0]=[report.typed(header,'management header') for header in headers]
    return table,identifiers,headers


def daily_table(result,report,mapping):
    if result.get('schema_version')!='daily_summary_v1':
        raise ValueError('Unknown daily-summary schema')
    return management_table({**result,'schema_version':'management_wide_v1'},report,mapping)


def management_audit_tables(result,report,table_name="Wide"):
    dictionary=result.get('metric_dictionary'); validations=result.get('validation'); lineage=result.get('lineage')
    if not isinstance(dictionary,list) or not isinstance(validations,list) or not isinstance(lineage,list):
        raise ValueError('Management report needs dictionary, validation and lineage lists')
    definitions=[['Metric ID','Label','Unit','Kind','Time basis','Aggregation','Additive dimensions','Reduction dimensions','Numerator','Denominator','Scale','Definition version']]
    for item in dictionary:
        if not isinstance(item,dict) or any(not isinstance(item.get(field),str) for field in ('metric_id','label','unit','kind','time_basis','aggregation')):
            raise ValueError('Invalid management dictionary')
        definitions.append([item['metric_id'],item['label'],item['unit'],item['kind'],item['time_basis'],item['aggregation'],
                            json.dumps(item.get('additive_dimensions',[])),json.dumps(item.get('reduction_dimensions',[])),
                            item.get('numerator'),item.get('denominator'),report.number(item.get('scale'),'metric scale'),result.get('definition_version')])
    checks=[['Check','Status','Detail']]
    for item in validations:
        if not isinstance(item,dict) or set(item)!={'check','status','detail'} or item['status'] not in ('pass','warning') or not isinstance(item['check'],str) or not isinstance(item['detail'],str):
            raise ValueError('Invalid management validation record')
        checks.append([item['check'],item['status'],item['detail']])
    sources=[['Report row ID',table_name+' Excel row','Source file','Source sheet','Source row','Metric ID']]
    for item in lineage:
        if not isinstance(item,dict) or set(item)!={'output_row','source_file','sheet','source_row','metric_id'}:
            raise ValueError('Invalid management lineage record')
        row=count(item['output_row'],'lineage output_row'); source_row=count(item['source_row'],'lineage source_row')
        if not 1<=row<=result['output_rows'] or not source_row or any(not isinstance(item[field],str) for field in ('source_file','sheet','metric_id')):
            raise ValueError('Management lineage points outside output')
        sources.append([row,row+1,item['source_file'],item['sheet'],source_row,item['metric_id']])
    return [report.sheet('Definitions',definitions),report.sheet('Validation',checks),report.sheet('Lineage',sources)]


def export_result(result,output,readable=True):
    if not isinstance(result,dict) or result.get('status') not in ('selected','summarized','described','management_wide','daily_summary'):
        raise ValueError('Supported result statuses: selected, summarized, described, management_wide_v1, daily_summary_v1; unknown schemas are refused')
    if not isinstance(result.get('sheet'),str) or not isinstance(result.get('range'),str):
        raise ValueError('A native source sheet and range are required')
    report=Report()
    mapping=source_columns(result)
    convert={'selected':selected_table,'summarized':summarized_table,'described':described_table,'management_wide':management_table,'daily_summary':daily_table}[result['status']]
    data,identifiers,headers=convert(result,report,mapping)
    column_rows=columns_table(mapping,identifiers,headers,report)
    if result.get('unverified_formula_caches',0):
        report.warnings.append(['formula caches','Source used unverified formula caches; export does not recalculate them',str(result['unverified_formula_caches'])])
    if 'requested_query' not in result:
        report.warnings.append(['query provenance','Input JSON contains no complete requested query; none was reconstructed',None])
    snapshot=json.dumps(result,ensure_ascii=False,allow_nan=False,default=str)
    if len(snapshot.encode('utf-8'))>MAX_INPUT_BYTES:
        raise ValueError('Source result snapshot exceeds 16 MiB report budget')
    metadata=[['Section','Part','JSON/text'],['result_status',1,result['status']],
              ['numeric_policy',1,'Explicit numeric fields only; >15 significant digits/out-of-range values are exact text; text identifiers never coerced'],
              ['formula_policy',1,'All output cells are literal values; no input text is evaluated as a formula'],
              ['display_labels',1,'Joined path labels are display only; Columns and source JSON retain structured identity']]
    metadata += [['source_json',index+1,snapshot[start:start+30000]] for index,start in enumerate(range(0,len(snapshot),30000))]
    warning_rows=[['Location','Issue','Original value']]+report.warnings
    sheets=[report.sheet('Wide' if result['status']=='management_wide' else 'Daily' if result['status']=='daily_summary' else 'Results',data),report.sheet('Columns',column_rows),report.sheet('Metadata',metadata),report.sheet('Warnings',warning_rows)]
    if result['status'] in ('management_wide','daily_summary'):
        sheets += management_audit_tables(result,report,'Daily' if result['status']=='daily_summary' else 'Wide')
    preset=None
    if readable:
        preset={'preset':'readable','sheets':{sheet['name']:{'header_rows':'1','wrap_text':sheet['name'] in ('Metadata','Warnings'),
                    'column_widths':{'A':24}} for sheet in sheets}}
        if result['status']=='management_wide':
            preset['sheets']['Validation']['category_columns']=['B']
            widths={}
            for index,identifier in enumerate(result['columns'],1):
                parts=json.loads(identifier)
                widths[column_name(index)]=24 if parts[:2]==['dimension','institution'] else 32 if parts[0]=='quality' else 20
            preset['sheets']['Wide']['column_widths']=widths
            preset['sheets']['Metadata']['column_widths']={'A':24,'B':10,'C':80}
            preset['sheets']['Validation']['column_widths']={'A':30,'B':12,'C':80}
            preset['sheets']['Definitions']['column_widths']={'A':24,'B':30,'C':12,'D':12,'E':20,'F':16}
            preset['sheets']['Lineage']['column_widths']={'A':16,'B':16,'C':40,'D':20,'E':16,'F':24}
            row_dimensions=result.get('requested_query',{}).get('row_dimensions',[])
            if 'institution' in row_dimensions:
                preset['sheets']['Wide']['category_columns']=[column_name(row_dimensions.index('institution')+1)]
    book=Workbook.create({'sheets':sheets},appearance=preset)
    book.save(output)
    return {'status':'exported','source_status':result['status'],'output_written':True,'output':str(output),
            'data_rows':len(data)-1,'warning_count':len(report.warnings),'total_cells':report.count,
            'appearance':book.appearance_report,'formulas_created':0}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input')
    parser.add_argument('output')
    parser.add_argument('--plain',action='store_true')
    args=parser.parse_args()
    try:
        print(json.dumps(export_result(read_spec(args.input,max_bytes=MAX_INPUT_BYTES),args.output,not args.plain),ensure_ascii=False))
        return 0
    except (ValueError,OSError,LookupError,DecimalException,ET.ParseError,zipfile.BadZipFile) as exc:
        print(json.dumps({'status':'failed','output_written':False,'message':str(exc)},ensure_ascii=False))
        return 1


if __name__=='__main__':
    sys.exit(main())
