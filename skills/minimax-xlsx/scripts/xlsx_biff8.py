#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Fail-closed MS-XLS BIFF8 record subset, independent of ET/ETT compatibility claims."""
import argparse
from decimal import Decimal
import hashlib
import json
import math
from pathlib import Path
import struct
import sys
from types import SimpleNamespace

sys.path.insert(0,str(Path(__file__).resolve().parent))
from xlsx_cfb import CompoundFile,MAX_STREAM_BYTES
from xlsx_tabular import column_name
from stdlib_xlsx import address
from xlsx_headers import reconstruct

MAX_RECORDS=100000
MAX_CELLS=200000
MAX_MERGES=256
MAX_METADATA_BYTES=1024*1024
ERRORS={0:'#NULL!',7:'#DIV/0!',15:'#VALUE!',23:'#REF!',29:'#NAME?',36:'#NUM!',42:'#N/A',43:'#GETTING_DATA'}
GLOBAL_METADATA={0x0042:'CodePage',0x003D:'Window1',0x0031:'Font',0x0293:'Style'}
SHEET_METADATA={0x0200:'Dimensions',0x0208:'Row',0x020B:'Index',0x00D7:'DBCell',0x023E:'Window2'}


def require(condition,message):
    if not condition: raise ValueError(message)


def unicode_string(data,offset=0,short=False):
    prefix=2 if short else 3
    require(len(data)>=offset+prefix,'Truncated Unicode string header')
    count=data[offset] if short else struct.unpack_from('<H',data,offset)[0]
    flag=data[offset+prefix-1]
    require(count<=32767 and flag in (0,1),'Rich/phonetic/unknown string flags unsupported')
    start=offset+prefix; end=start+count*(2 if flag else 1)
    require(end<=len(data),'Truncated string; Continue records unsupported')
    try: value=data[start:end].decode('utf-16le' if flag else 'latin1')
    except UnicodeDecodeError as exc: raise ValueError('Malformed UTF-16 string') from exc
    return value,end


def number(raw):
    require(len(raw)==8,'Xnum must contain eight bytes')
    value=struct.unpack('<d',raw)[0]
    require(math.isfinite(value),'Nonfinite Xnum unsupported')
    return Decimal.from_float(value)


def rk_value(raw):
    require(len(raw)==4,'RK value must contain four bytes')
    bits=struct.unpack('<I',raw)[0]
    if bits&2:
        value=Decimal(struct.unpack('<i',raw)[0]>>2)
    else:
        value=number(struct.pack('<Q',(bits&0xFFFFFFFC)<<32))
    if bits&1:
        sign,digits,exponent=value.as_tuple(); value=Decimal((sign,digits,exponent-2))
    return value


def records(data):
    require(isinstance(data,bytes) and len(data)<=MAX_STREAM_BYTES,'BIFF stream exceeds 8 MiB or is not bytes')
    output=[]; offset=0
    while offset<len(data):
        require(len(output)<MAX_RECORDS,'BIFF record count limit exceeded')
        require(len(data)-offset>=4,'Truncated BIFF record header')
        kind,size=struct.unpack_from('<HH',data,offset)
        require(size<=8224 and offset+4+size<=len(data),'Unsupported/oversized/truncated BIFF record payload')
        if kind==0x002F: raise ValueError('Encrypted BIFF FILEPASS unsupported; no decryption attempted')
        output.append((offset,kind,data[offset+4:offset+4+size])); offset+=4+size
    require(output and output[0][1]==0x0809,'Stream must start with BIFF8 BOF, not a filename hint')
    return output


def bof(record,expected):
    offset,kind,data=record
    require(kind==0x0809 and len(data)==16,'Expected 16-byte BIFF8 BOF at '+str(offset))
    version,substream=struct.unpack_from('<HH',data)
    require(version==0x0600 and substream==expected,'Unsupported BIFF version/substream; only BIFF8 globals and worksheets')
    return data.hex()


class Reader:
    def __init__(self,data):
        self.data=data; self.records=records(data); self.styles=[]; self.formats={}; self.strings=[]
        self.metadata=[]; self.metadata_bytes=0; self.issues=[]; self.cell_count=0; self.merge_count=0

    def metadata_record(self,record,names,scope):
        offset,kind,payload=record
        self.metadata_bytes+=len(payload)
        require(self.metadata_bytes<=MAX_METADATA_BYTES and len(self.metadata)<4096,'Uninterpreted metadata budget exceeded')
        self.metadata.append({'scope':scope,'offset':offset,'record_id':hex(kind),'name':names[kind],'payload_hex':payload.hex()})

    def issue(self,issue):
        require(len(self.issues)<1000,'Diagnostic count exceeds 1000')
        self.issues.append(issue)

    def parse(self):
        bof(self.records[0],5)
        bounds=[]; names=set(); date_system=None; seen_sst=False; global_end=None
        for index,record in enumerate(self.records[1:],1):
            offset,kind,payload=record
            if kind==0x000A:
                require(not payload,'EOF payload must be empty'); global_end=index+1; break
            if kind==0x0085:
                require(len(payload)>=8,'Truncated BoundSheet8')
                position=struct.unpack_from('<I',payload)[0]; hidden=payload[4]&3; sheet_type=payload[5]
                require(hidden in (0,1,2) and sheet_type==0,'Macro/chart/VBA BoundSheet8 types unsupported')
                name,end=unicode_string(payload,6,True)
                require(end==len(payload) and 1<=payload[6]<=31 and not any(char in name for char in '\0\x03:\\*?/[]') and not name.startswith("'") and not name.endswith("'"),'Invalid worksheet name')
                require(name.casefold() not in names and len(bounds)<128,'Duplicate sheet names or sheet count exceeds 128')
                names.add(name.casefold()); bounds.append({'name':name,'offset':position,'hidden_state':hidden})
            elif kind==0x0022:
                require(len(payload)==2 and date_system is None,'Duplicate/malformed Date1904')
                flag=struct.unpack('<H',payload)[0]; require(flag in (0,1),'Invalid Date1904 value'); date_system='1904' if flag else '1900'
            elif kind==0x00E0:
                require(len(payload)==20 and len(self.styles)<4096,'Malformed XF or style count exceeds 4096')
                self.styles.append({'xf_index':len(self.styles),'number_format_id_raw':struct.unpack_from('<H',payload,2)[0],'raw_hex':payload.hex()})
            elif kind==0x041E:
                require(len(payload)>=5,'Truncated Format')
                identity=struct.unpack_from('<H',payload)[0]; value,end=unicode_string(payload,2)
                require(end==len(payload) and 1<=len(value)<=255 and identity not in self.formats,'Malformed/duplicate number format')
                self.formats[identity]=value
            elif kind==0x00FC:
                require(not seen_sst and len(payload)>=8,'Duplicate/truncated SST'); seen_sst=True
                total,unique=struct.unpack_from('<ii',payload)
                require(0<=unique<=total and unique<=10000,'Invalid SST reference/unique counts')
                position=8
                for _ in range(unique):
                    value,position=unicode_string(payload,position); self.strings.append(value)
                require(position==len(payload),'Unparsed SST tail; rich/continued strings unsupported')
            elif kind in GLOBAL_METADATA:
                if kind==0x0042: require(len(payload)==2 and struct.unpack('<H',payload)[0]==1200,'Only Unicode CodePage 1200 accepted')
                self.metadata_record(record,GLOBAL_METADATA,'globals')
            else: raise ValueError('Unsupported global record '+hex(kind)+' at '+str(offset))
        require(global_end is not None and bounds and self.styles and date_system is not None,'Require complete globals, at least one worksheet/XF and explicit Date1904')
        offset_index={record[0]:index for index,record in enumerate(self.records)}
        require(len({item['offset'] for item in bounds})==len(bounds),'Aliased worksheet substreams')
        physical=sorted(bounds,key=lambda item:item['offset']); position=global_end; sheets={}
        for bound in physical:
            require(bound['offset'] in offset_index and offset_index[bound['offset']]==position,'Worksheet offset gap/overlap or unbound records')
            sheet,position=self.sheet(position,bound); sheets[bound['name']]=sheet
        require(position==len(self.records),'Trailing/unbound BIFF records are not ignored')
        return {'status':'read_biff8_subset_with_diagnostics' if self.issues else 'read_biff8_subset',
                'schema_version':'biff8_readonly_v1','support_level':'biff8_strict_record_subset','biff_parsed':True,
                'native_et_ett_parsed':False,'wps_compatibility_verified':False,'formulas_executed':False,'external_links_followed':False,
                'date_system':date_system,'sheets':[sheets[item['name']] for item in bounds],
                'styles':self.styles,'custom_number_formats':{str(key):value for key,value in self.formats.items()},
                'uninterpreted_metadata':self.metadata,'issues':self.issues,'record_count':len(self.records),'cell_count':self.cell_count,
                'policies':{'numbers':'Exact Decimal representation of stored IEEE754; raw bytes retained; no text coercion',
                            'dates':'All stored numeric serials remain numbers; raw XF/format IDs and custom strings retained; no ISO conversion or effective-style inheritance inference',
                            'formulas':'Opaque tokens/ancillary bytes only, never decompiled or executed; every present cache is unverified and may be stale',
                            'unknown_records':'Fail closed; known metadata is explicitly retained, not silently discarded',
                            'scope':'Record-subset extraction, not full MS-XLS grammar conformance or real ET/ETT acceptance'}}

    def sheet(self,index,bound):
        bof_raw=bof(self.records[index],0x0010); index+=1; cells={}; merges=[]; ws_bool=False; pending=None
        def put(row,col,xf,kind,value,record,**extra):
            require(0<=row<65536 and 0<=col<256 and 0<=xf<len(self.styles),'Cell coordinate or XF index outside BIFF8 bounds')
            address=column_name(col+1)+str(row+1)
            require(address not in cells and self.cell_count<MAX_CELLS,'Duplicate cell address or cell limit exceeded')
            fmt=self.styles[xf]['number_format_id_raw']
            cells[address]={'type':kind,'value':value,'xf_index':xf,'number_format_id_raw':fmt,'custom_format_raw':self.formats.get(fmt),
                            'record_offset':record[0],'record_id':hex(record[1]),'cache_status':'not_applicable',**extra}
            self.cell_count+=1; return address
        while index<len(self.records):
            record=self.records[index]; offset,kind,payload=record; index+=1
            if pending is not None:
                if kind==0x0207:
                    value,end=unicode_string(payload); require(end==len(payload),'Unparsed cached String tail')
                    cells[pending].update(type='text',value=value,cache_status='unverified',cache_record_offset=offset); pending=None; continue
                cells[pending].update(type='blank',value=None,cache_status='missing')
                self.issue({'sheet':bound['name'],'cell':pending,'code':'missing_formula_string_cache','expected':'Immediate String record'}); pending=None
            if kind==0x000A:
                require(not payload and ws_bool,'Require empty EOF and exactly one non-dialog WsBool')
                return {**bound,'bof_raw_hex':bof_raw,'cells':cells,'merges':merges},index
            if kind==0x0081:
                require(not ws_bool and len(payload)==2 and not struct.unpack('<H',payload)[0]&0x0010,'Duplicate/malformed/dialog WsBool unsupported')
                ws_bool=True
            elif kind in (0x0203,0x027E,0x0205,0x0201,0x00FD,0x0006):
                require(len(payload)>=6,'Truncated cell record'); row,col,xf=struct.unpack_from('<3H',payload)
                if kind==0x0203:
                    require(len(payload)==14,'Number payload must be 14 bytes'); put(row,col,xf,'number',number(payload[6:]),record,raw_numeric_hex=payload[6:].hex())
                elif kind==0x027E:
                    require(len(payload)==10,'RK payload must be 10 bytes'); put(row,col,xf,'number',rk_value(payload[6:]),record,raw_numeric_hex=payload[6:].hex())
                elif kind==0x0201:
                    require(len(payload)==6,'Blank payload must be 6 bytes'); put(row,col,xf,'blank',None,record)
                elif kind==0x00FD:
                    require(len(payload)==10,'LabelSst payload must be 10 bytes'); string_index=struct.unpack_from('<I',payload,6)[0]
                    require(string_index<len(self.strings),'SST index outside table'); put(row,col,xf,'text',self.strings[string_index],record,sst_index=string_index)
                elif kind==0x0205:
                    require(len(payload)==8 and payload[7] in (0,1),'Malformed BoolErr')
                    if payload[7]:
                        require(payload[6] in ERRORS,'Unknown BErr code'); put(row,col,xf,'error',ERRORS[payload[6]],record,error_code=payload[6])
                    else:
                        require(payload[6] in (0,1),'Invalid boolean cell'); put(row,col,xf,'boolean',bool(payload[6]),record)
                else:
                    require(len(payload)>=23,'Truncated Formula')
                    flags=struct.unpack_from('<H',payload,14)[0]; count=struct.unpack_from('<H',payload,20)[0]
                    require(not flags&8 and 0<count<=len(payload)-22,'Shared or malformed formula tokens unsupported')
                    formula={'encoding':'opaque_biff8_ptgs','tokens_hex':payload[22:22+count].hex(),'ancillary_hex':payload[22+count:].hex(),
                             'flags_raw':flags,'needs_recalculation_flag':bool(flags&1),'cache_raw_hex':payload[6:14].hex(),'cache_may_be_stale':True}
                    raw=payload[6:14]; cache_status='unverified'; error_code=None
                    if raw[6:8]!=b'\xff\xff': value=number(raw); cell_type='number'
                    else:
                        cache_type=raw[0]
                        if cache_type==0: value=None; cell_type='blank'; cache_status='pending_string'
                        elif cache_type==1:
                            require(raw[2] in (0,1),'Invalid boolean formula cache'); value=bool(raw[2]); cell_type='boolean'
                        elif cache_type==2:
                            require(raw[2] in ERRORS,'Unknown cached error code'); value=ERRORS[raw[2]]; cell_type='error'; error_code=raw[2]; cache_status='error_unverified'
                        elif cache_type==3: value=''; cell_type='text'; cache_status='unverified_empty_string'
                        else: raise ValueError('Unknown special FormulaValue type')
                    address=put(row,col,xf,cell_type,value,record,formula=formula,error_code=error_code)
                    cells[address]['cache_status']=cache_status
                    if cache_status=='pending_string': pending=address
            elif kind in (0x00BD,0x00BE):
                require(len(payload)>=6,'Truncated multi-cell record'); row,first=struct.unpack_from('<HH',payload); last=struct.unpack_from('<H',payload,len(payload)-2)[0]
                width=6 if kind==0x00BD else 2
                require(first<last<256 and len(payload)==6+(last-first+1)*width,'Malformed MulRk/MulBlank bounds/length')
                for col in range(first,last+1):
                    start=4+(col-first)*width; xf=struct.unpack_from('<H',payload,start)[0]
                    raw=payload[start+2:start+6] if kind==0x00BD else None
                    put(row,col,xf,'number' if raw is not None else 'blank',rk_value(raw) if raw is not None else None,record,raw_numeric_hex=None if raw is None else raw.hex())
            elif kind==0x00E5:
                require(len(payload)>=2,'Truncated MergeCells'); count=struct.unpack_from('<H',payload)[0]
                require(count<=1026 and len(payload)==2+8*count and self.merge_count+count<=MAX_MERGES,'Malformed/oversized merge record')
                for start in range(2,len(payload),8):
                    r1,r2,c1,c2=struct.unpack_from('<4H',payload,start)
                    require(r1<=r2 and c1<=c2<256 and (r1,c1)!=(r2,c2),'Invalid merge bounds')
                    for other in merges:
                        a,b,c,d=other['bounds_zero_based']
                        require(r2<a or b<r1 or c2<c or d<c1,'Overlapping merges unsupported')
                    merges.append({'range':column_name(c1+1)+str(r1+1)+':'+column_name(c2+1)+str(r2+1),
                                   'bounds_zero_based':[r1,r2,c1,c2],'record_offset':offset})
                    self.merge_count+=1
            elif kind in SHEET_METADATA: self.metadata_record(record,SHEET_METADATA,bound['name'])
            else: raise ValueError('Unsupported worksheet record '+hex(kind)+' at '+str(offset))
        raise ValueError('Truncated worksheet without EOF')


def parse_stream(data):
    return Reader(data).parse()


def read_file(path):
    container=CompoundFile.load(path)
    require(not container.encrypted,'Encrypted compound workbook unsupported')
    require(not any(part.casefold() in ('vba','_vba_project_cur','macros') for stream in container.paths for part in stream),'Macro storage unsupported')
    candidates=[stream for stream in container.streams if len(stream)==1 and stream[0].casefold() in ('workbook','book')]
    require(len(candidates)==1,'Require exactly one root Workbook/Book stream; name is only a candidate until BOF validation')
    content=container.read_stream(candidates[0]); result=parse_stream(content)
    return {**result,'source':{'file':str(path),'extension':Path(path).suffix.lower(),'stream_path':list(candidates[0]),
                             'stream_sha256':hashlib.sha256(content).hexdigest(),'container_version':container.major},
            'other_container_streams_not_decoded':[item for item in container.report()['streams'] if item['path']!=list(candidates[0])]}



def reconstruct_headers(result,sheet,cell_range,header_rows,formulas='reject'):
    require(result.get('schema_version')=='biff8_readonly_v1','Require native BIFF8 subset result')
    require(formulas in ('reject','cached'),'Header formulas policy must be reject/cached')
    choices=[item for item in result['sheets'] if item['name']==sheet]
    require(len(choices)==1,'Unknown BIFF worksheet')
    require(isinstance(cell_range,str),'Explicit A1 range required for BIFF header adapter')
    ends=cell_range.split(':'); require(len(ends) in (1,2),'Invalid BIFF header range')
    first_row,first_col=address(ends[0]); last_row,last_col=address(ends[-1])
    require(1<=first_row<=last_row<=65536 and 1<=first_col<=last_col<=256 and (last_row-first_row+1)*(last_col-first_col+1)<=MAX_CELLS,'BIFF header range exceeds bounds/cell budget')
    data=choices[0]; cells=data['cells']
    def read_cell(coordinate):
        cell=cells.get(coordinate,{'type':'blank','value':None})
        if cell.get('formula') is not None:
            require(formulas=='cached','Formula header requires explicit cached opt-in; caches unverified')
            require(cell['cache_status'] not in ('missing','error_unverified'),'Missing/error formula header cache cannot form a path')
        require(cell['type']!='error','Error cell cannot form a header path')
        return cell['type'],cell['value']
    schema=reconstruct(SimpleNamespace(cells={sheet:cells}),sheet,(first_row,first_col,last_row,last_col),header_rows,read_cell,
                       merge_nodes=[{'ref':item['range']} for item in data['merges']])
    schema['source_format']='BIFF8 strict subset'; schema['formula_policy']=formulas
    schema['cache_warning']='Any formula cache used is unverified and may be stale' if formulas=='cached' else None
    return schema


def main():
    parser=argparse.ArgumentParser(description=__doc__); parser.add_argument('input')
    parser.add_argument('--sheet'); parser.add_argument('--range',dest='cell_range'); parser.add_argument('--header-rows')
    parser.add_argument('--header-formulas',choices=('reject','cached'),default='reject'); args=parser.parse_args()
    try:
        result=read_file(args.input)
        if args.header_rows is not None:
            require(args.sheet is not None and args.cell_range is not None,'Header reconstruction requires explicit sheet/range/header-rows')
            result['header_schema']=reconstruct_headers(result,args.sheet,args.cell_range,args.header_rows,args.header_formulas)
        elif args.sheet is not None or args.cell_range is not None or args.header_formulas!='reject':
            raise ValueError('Header options require explicit header-rows')
        encoded=json.dumps(result,ensure_ascii=False,default=str,allow_nan=False)
        require(len(encoded.encode('utf-8'))<=16*1024*1024,'Decoded JSON exceeds 16 MiB budget')
        print(encoded); return 0
    except (ValueError,OSError,struct.error,OverflowError) as exc:
        print(json.dumps({'status':'invalid_or_unsupported','message':str(exc),'biff_parsed':False,'native_et_ett_parsed':False},ensure_ascii=False)); return 1


if __name__=='__main__': sys.exit(main())
