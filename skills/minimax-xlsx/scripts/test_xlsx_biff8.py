#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Synthetic MS-XLS record fixtures, never real Excel/WPS acceptance samples."""
from decimal import Decimal
import json
from pathlib import Path
import random
import struct
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parent))
from test_xlsx_cfb import synthetic_cfb
from xlsx_biff8 import parse_stream,read_file,rk_value,reconstruct_headers


def rec(kind,payload=b''): return struct.pack('<HH',kind,len(payload))+payload

def bof(kind=5,version=0x0600): return rec(0x0809,struct.pack('<4H2I',version,kind,0x2013,1997,1,6))

def string(value,short=False,compressed=False):
    raw=value.encode('latin1' if compressed else 'utf-16le'); count=len(raw) if compressed else len(raw)//2
    return (struct.pack('<BB',count,0 if compressed else 1) if short else struct.pack('<HB',count,0 if compressed else 1))+raw

def cell(row,col,xf=0): return struct.pack('<3H',row,col,xf)

def special(kind,value=0): return bytes([kind,0,value,0,0,0,255,255])

def formula(row,col,cache,flags=0,tokens=b'\x1e\x2a\x00'):
    return rec(6,cell(row,col)+cache+struct.pack('<HIH',flags,0,len(tokens))+tokens)


def workbook(sheet_records=None,strings=None,date_system=0,extra_global=b'',names=None):
    strings=['部门','金额','00012345678901234567890'] if strings is None else strings
    if sheet_records is None:
        sheet_records=[rec(0x00FD,cell(0,0)+struct.pack('<I',0))+rec(0x00FD,cell(0,1)+struct.pack('<I',1))+
                       rec(0x00FD,cell(1,0)+struct.pack('<I',2))+rec(0x0203,cell(1,1)+struct.pack('<d',.1))+
                       rec(0x0203,cell(1,2,1)+struct.pack('<d',60.5))+rec(0x0205,cell(1,3)+bytes([1,0]))+
                       rec(0x0205,cell(1,4)+bytes([7,1]))+rec(0x0201,cell(1,5))+
                       rec(0x027E,cell(2,0)+struct.pack('<I',((-123<<2)|2)&0xFFFFFFFF))+
                       rec(0x00BD,struct.pack('<HHHIHIH',2,1,0,((12345<<2)|3),0,((5<<2)|2),2))+
                       rec(0x00BE,struct.pack('<5H',2,3,0,0,4))+
                       formula(3,1,struct.pack('<d',42),1)+formula(3,2,special(0))+rec(0x0207,string('缓存中文'))+
                       formula(3,3,special(2,7))+formula(3,4,special(3))+formula(3,5,special(1,1))+
                       rec(0x00E5,struct.pack('<5H',1,4,4,0,1)),b'']
    names=['中文表'+str(index+1) for index in range(len(sheet_records))] if names is None else names
    globals_part=bof()+rec(0x0022,struct.pack('<H',date_system))+rec(0x00E0,bytes(20))+rec(0x00E0,struct.pack('<HH',0,164)+bytes(16))+rec(0x041E,struct.pack('<H',164)+string('yyyy-mm-dd'))
    globals_part+=rec(0x00FC,struct.pack('<ii',len(strings),len(strings))+b''.join(string(value) for value in strings))+extra_global
    sheets=[bof(0x10)+rec(0x0081,bytes(2))+payload+rec(0x000A) for payload in sheet_records]
    def bound(index,offset): return rec(0x0085,struct.pack('<IBB',offset,0,0)+string(names[index],True))
    header=globals_part+b''.join(bound(index,0) for index in range(len(sheets)))+rec(0x000A)
    offset=len(header); bounds=[]
    for index,sheet in enumerate(sheets): bounds.append(bound(index,offset)); offset+=len(sheet)
    return globals_part+b''.join(bounds)+rec(0x000A)+b''.join(sheets)


class BiffTests(unittest.TestCase):
    def test_typed_cells_chinese_exact_binary_and_long_text_id(self):
        result=parse_stream(workbook()); cells=result['sheets'][0]['cells']
        self.assertEqual(cells['A1']['value'],'部门'); self.assertEqual(cells['A2']['value'],'00012345678901234567890')
        self.assertEqual(cells['B2']['value'],Decimal.from_float(.1)); self.assertEqual(cells['D2']['value'],True)
        self.assertEqual(cells['E2']['type'],'error'); self.assertEqual(cells['E2']['value'],'#DIV/0!')
        self.assertIsNone(cells['F2']['value']); self.assertEqual(cells['A3']['value'],-123)
        self.assertEqual((cells['B3']['value'],cells['C3']['value']),(Decimal('123.45'),Decimal(5)))
        self.assertIsNone(cells['D3']['value']); self.assertEqual(result['sheets'][1]['cells'],{})
        self.assertTrue(result['biff_parsed']); self.assertFalse(result['native_et_ett_parsed'])

    def test_date_system_raw_serial_and_format_preserved_no_fake_1900_date(self):
        for system in (0,1):
            result=parse_stream(workbook(date_system=system)); value=result['sheets'][0]['cells']['C2']
            self.assertEqual(result['date_system'],'1904' if system else '1900')
            self.assertEqual(value['value'],Decimal('60.5')); self.assertEqual(value['type'],'number')
            self.assertEqual(value['number_format_id_raw'],164); self.assertEqual(value['custom_format_raw'],'yyyy-mm-dd')
            self.assertEqual(result['styles'][1]['raw_hex'][4:8],'a400')

    def test_formula_numeric_text_boolean_error_and_empty_cache_never_executed(self):
        result=parse_stream(workbook()); cells=result['sheets'][0]['cells']
        self.assertEqual(cells['B4']['value'],42); self.assertTrue(cells['B4']['formula']['needs_recalculation_flag'])
        self.assertEqual(cells['C4']['value'],'缓存中文'); self.assertEqual(cells['C4']['cache_status'],'unverified')
        self.assertEqual(cells['D4']['cache_status'],'error_unverified'); self.assertEqual(cells['D4']['value'],'#DIV/0!')
        self.assertEqual(cells['E4']['value'],''); self.assertEqual(cells['E4']['cache_status'],'unverified_empty_string')
        self.assertEqual(cells['F4']['value'],True); self.assertFalse(result['formulas_executed'])
        self.assertTrue(all(cells[address]['formula']['cache_may_be_stale'] for address in ('B4','C4','D4','E4','F4')))

    def test_missing_string_cache_is_explicit_not_zero_or_empty(self):
        result=parse_stream(workbook([formula(0,0,special(0))],strings=[]))
        value=result['sheets'][0]['cells']['A1']; self.assertEqual(value['cache_status'],'missing'); self.assertIsNone(value['value'])
        self.assertEqual(result['status'],'read_biff8_subset_with_diagnostics')
        self.assertEqual(result['issues'][0]['code'],'missing_formula_string_cache')

    def test_merge_bounds_and_no_expansion_or_data_erasure(self):
        payload=rec(0x0203,cell(0,0)+struct.pack('<d',1))+rec(0x0203,cell(0,1)+struct.pack('<d',2))+rec(0x00E5,struct.pack('<5H',1,0,0,0,1))
        sheet=parse_stream(workbook([payload],strings=[]))['sheets'][0]
        self.assertEqual(sheet['merges'][0]['range'],'A1:B1'); self.assertEqual(sheet['cells']['B1']['value'],2)
        bad=rec(0x00E5,struct.pack('<9H',2,0,1,0,1,1,2,1,2))
        with self.assertRaisesRegex(ValueError,'Overlapping'): parse_stream(workbook([bad],strings=[]))

    def test_rk_float_integer_scaled_and_nonfinite(self):
        self.assertEqual(rk_value(struct.pack('<I',0x3FF00000)),1)
        self.assertEqual(rk_value(struct.pack('<I',0x3FF00001)),Decimal('.01'))
        with self.assertRaises(ValueError): rk_value(struct.pack('<I',0x7FF00000))
        bad=rec(0x0203,cell(0,0)+struct.pack('<d',float('nan')))
        with self.assertRaisesRegex(ValueError,'Nonfinite'): parse_stream(workbook([bad],strings=[]))

    def test_unknown_encrypted_continued_shared_and_external_records_fail(self):
        for kind in (0x7777,0x003C,0x002F,0x01AE,0x0017,0x04BC,0x0221):
            with self.assertRaises(ValueError): parse_stream(workbook([rec(kind)],strings=[]))
        with self.assertRaisesRegex(ValueError,'Shared'): parse_stream(workbook([formula(0,0,struct.pack('<d',1),8)],strings=[]))
        with self.assertRaises(ValueError): parse_stream(workbook([rec(0x0207,string('orphan'))],strings=[]))

    def test_signature_version_offsets_names_dialog_and_truncation_fail(self):
        original=workbook()
        with self.assertRaises(ValueError): parse_stream(b'not BIFF')
        for data in (original[:-1],original[:3],original+rec(0x000A)):
            with self.assertRaises(ValueError): parse_stream(data)
        data=bytearray(original); struct.pack_into('<H',data,4,0x0500)
        with self.assertRaisesRegex(ValueError,'version'): parse_stream(bytes(data))
        data=bytearray(original)
        # Locate by records rather than relying on a particular Unicode name length.
        position=0
        while struct.unpack_from('<H',data,position)[0]!=0x0085: position+=4+struct.unpack_from('<H',data,position+2)[0]
        struct.pack_into('<I',data,position+4,1)
        with self.assertRaisesRegex(ValueError,'offset'): parse_stream(bytes(data))
        with self.assertRaisesRegex(ValueError,'name'): parse_stream(workbook(names=['same','same']))
        with self.assertRaises(ValueError): parse_stream(workbook(names=['bad/name','okay']))
        with self.assertRaises(ValueError): parse_stream(workbook([rec(0x0081,struct.pack('<H',16))],strings=[]))

    def test_sst_compressed_unicode_bad_index_rich_and_style_refs(self):
        result=parse_stream(workbook([rec(0x00FD,cell(0,0)+struct.pack('<I',0))],strings=['é😀中文']))
        self.assertEqual(result['sheets'][0]['cells']['A1']['value'],'é😀中文')
        for payload in [rec(0x00FD,cell(0,0)+struct.pack('<I',999)),rec(0x0201,cell(0,0,999)),rec(0x0203,cell(0,256)+struct.pack('<d',1))]:
            with self.assertRaises(ValueError): parse_stream(workbook([payload]))
        with self.assertRaisesRegex(ValueError,'Unsupported global'): parse_stream(workbook(extra_global=rec(0x7777)))
        from xlsx_biff8 import unicode_string
        self.assertEqual(unicode_string(string('é',compressed=True))[0],'é')
        with self.assertRaisesRegex(ValueError,'Rich'): unicode_string(struct.pack('<HB',1,9)+b'x')

    def test_duplicate_cells_limits_and_known_metadata_retention(self):
        payload=rec(0x0201,cell(0,0))*2
        with self.assertRaisesRegex(ValueError,'Duplicate cell'): parse_stream(workbook([payload],strings=[]))
        for limit,value in [('MAX_RECORDS',2),('MAX_CELLS',1),('MAX_MERGES',0),('MAX_METADATA_BYTES',0)]:
            with patch('xlsx_biff8.'+limit,value):
                with self.assertRaises(ValueError): parse_stream(workbook(extra_global=rec(0x0031,b'raw metadata')))
        result=parse_stream(workbook(extra_global=rec(0x0031,b'raw metadata')))
        self.assertEqual(result['uninterpreted_metadata'][0]['payload_hex'],b'raw metadata'.hex())

    def test_cfb_cli_synthetic_only_and_no_write_or_external_tools(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'synthetic-biff8.cfb'; original=synthetic_cfb({'Workbook':workbook()}); path.write_bytes(original)
            result=read_file(path); self.assertTrue(result['biff_parsed']); self.assertFalse(result['native_et_ett_parsed'])
            run=subprocess.run([sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_biff8.py')),str(path)],capture_output=True,text=True,env={'PATH':''})
            self.assertEqual(run.returncode,0,run.stdout+run.stderr); self.assertFalse(json.loads(run.stdout)['wps_compatibility_verified'])
            self.assertEqual(path.read_bytes(),original); self.assertEqual(len(list(Path(folder).iterdir())),1)
            header_run=subprocess.run([sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_biff8.py')),str(path),'--sheet','中文表1','--range','A1:C3','--header-rows','1'],capture_output=True,text=True)
            self.assertEqual(header_run.returncode,0,header_run.stdout+header_run.stderr)
            self.assertEqual(len(json.loads(header_run.stdout)['header_schema']['columns']),3)
            path.write_bytes(synthetic_cfb({'Workbook':workbook(),'Book':workbook()}))
            with self.assertRaisesRegex(ValueError,'exactly one'): read_file(path)
            path.write_bytes(synthetic_cfb({'EncryptionInfo':b'encrypted','Workbook':workbook()}))
            with self.assertRaisesRegex(ValueError,'Encrypted'): read_file(path)

    def test_header_adapter_reuses_merge_only_paths_and_refuses_clipping(self):
        payload=rec(0x00FD,cell(0,0)+struct.pack('<I',0))+rec(0x00FD,cell(0,2)+struct.pack('<I',1))
        payload+=rec(0x00FD,cell(1,0)+struct.pack('<I',2))+rec(0x00FD,cell(1,1)+struct.pack('<I',3))+rec(0x00FD,cell(1,2)+struct.pack('<I',2))
        payload+=rec(0x00E5,struct.pack('<5H',1,0,0,0,1))
        result=parse_stream(workbook([payload],strings=['2025年','2024年','金额','户数']))
        schema=reconstruct_headers(result,'中文表1','A1:C3','1:2')
        self.assertEqual([part['value'] for part in schema['columns'][1]['path']],['2025年','户数'])
        self.assertEqual(schema['columns'][1]['levels'][0]['source'],'A1')
        with self.assertRaisesRegex(ValueError,'clips'): reconstruct_headers(result,'中文表1','B1:C3','1:2')

    def test_header_formula_caches_require_opt_in_and_missing_errors_fail(self):
        result=parse_stream(workbook([formula(0,0,struct.pack('<d',2025))],strings=[]))
        with self.assertRaisesRegex(ValueError,'cached opt-in'): reconstruct_headers(result,'中文表1','A1:B2','1')
        schema=reconstruct_headers(result,'中文表1','A1:B2','1','cached')
        self.assertIsNotNone(schema['cache_warning'])
        result=parse_stream(workbook([formula(0,0,special(0))],strings=[]))
        with self.assertRaisesRegex(ValueError,'Missing/error'): reconstruct_headers(result,'中文表1','A1:B2','1','cached')

    def test_seeded_mutations_fail_closed_or_retain_unverified_flags(self):
        original=workbook(); randomizer=random.Random(254)
        for _ in range(100):
            data=bytearray(original); offset=randomizer.randrange(len(data)); data[offset]=randomizer.randrange(256)
            try: result=parse_stream(bytes(data))
            except ValueError: continue
            self.assertFalse(result['native_et_ett_parsed']); self.assertFalse(result['formulas_executed'])


if __name__=='__main__': unittest.main()
