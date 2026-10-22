#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook
from test_xlsx_headers import chinese_book,merged_book
from xlsx_rows import select_rows
from xlsx_analysis import summarize
from xlsx_advanced import describe
from xlsx_result_export import export_result


class ExportTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.input=Path(self.temp.name)/'input.xlsx'
        self.output=Path(self.temp.name)/'report.xlsx'

    def snapshot(self,book):
        cells=book.cells['Metadata']
        parts=[]
        for ref,cell in cells.items():
            if ref.startswith('A') and cell['value']=='source_json':
                row=ref[1:]
                parts.append((int(cells['B'+row]['value']),cells['C'+row]['value']))
        return json.loads(''.join(text for _,text in sorted(parts)))

    def test_selected_end_to_end_paths_rows_query_and_literal_formulas(self):
        chinese_book().save(self.input)
        original=self.input.read_bytes()
        query={'sort':[{'column':{'path':['2024年','风控','金额']},'descending':True}]}
        result=select_rows(self.input,'报表',query,header_rows='3:5')
        export_result(json.loads(json.dumps(result,default=str)),self.output)
        book=Workbook.load(self.output)
        self.assertEqual(list(book.sheets),['Results','Columns','Metadata','Warnings'])
        self.assertEqual(book.cells['Results']['A2']['value'],8)
        self.assertEqual(book.cells['Results']['B2']['value'],'华西')
        self.assertEqual(self.snapshot(book)['requested_query'],query)
        self.assertEqual(self.snapshot(book)['header_schema']['columns'][1]['path'],result['header_schema']['columns'][1]['path'])
        self.assertTrue(all(cell['formula'] is None for cells in book.cells.values() for cell in cells.values()))
        self.assertEqual(self.input.read_bytes(),original)

    def test_precision_long_ids_and_text_formula_injection(self):
        result={'status':'selected','sheet':'Data','range':'A1:D2','columns':['A','B','C','D'],
                'headers':None,'output_rows':1,'rows':[{'source_row':2,'cells':[
                    {'type':'text','value':'0012345678901234567890'},
                    {'type':'number','value':'1234567890123456'},
                    {'type':'number','value':'19.90'},
                    {'type':'text','value':'=HYPERLINK("https://example.invalid")'}]}]}
        report=export_result(result,self.output)
        book=Workbook.load(self.output)
        self.assertEqual(book.cells['Results']['B2']['value'],'0012345678901234567890')
        self.assertEqual(book.cells['Results']['C2']['type'],'inlineStr')
        self.assertEqual(book.cells['Results']['C2']['value'],'1234567890123456')
        self.assertEqual(book.cells['Results']['D2']['type'],'n')
        self.assertIsNone(book.cells['Results']['E2']['formula'])
        self.assertGreaterEqual(report['warning_count'],2)

    def test_group_summary_single_header_and_aggregate_lineage_warning(self):
        merged_book({'A1':'部门','B1':'金额','A2':'甲','B2':2,'A3':'甲','B3':3},[]).save(self.input)
        result=summarize(self.input,'报表',{'group_by':['A'],'measure':'B'})
        export_result(json.loads(json.dumps(result,default=str)),self.output)
        book=Workbook.load(self.output)
        self.assertEqual(book.cells['Results']['A2']['value'],'甲')
        self.assertEqual(book.cells['Results']['G2']['value'],5)
        self.assertEqual(book.cells['Columns']['E2']['value'],'部门')
        self.assertTrue(any('contributing source-row' in str(cell['value']) for cell in book.cells['Warnings'].values()))

    def test_descriptive_stats_retain_rounding_definition_and_exact_long_results(self):
        merged_book({'A1':'指标','A2':1,'A3':2,'A4':2},[]).save(self.input)
        result=describe(self.input,'报表',{'columns':['A']})
        export_result(json.loads(json.dumps(result,default=str)),self.output)
        book=Workbook.load(self.output)
        self.assertEqual(book.cells['Results']['B2']['value'],'mean')
        self.assertEqual(book.cells['Results']['D2']['type'],'inlineStr')
        self.assertEqual(book.cells['Results']['D2']['value'],'1.666666666666666666666666666666667')
        self.assertEqual(self.snapshot(book)['quantile_algorithm'],result['quantile_algorithm'])
        self.assertEqual(self.snapshot(book)['arithmetic'],result['arithmetic'])

    def test_empty_selected_and_chunked_provenance(self):
        result={'status':'selected','sheet':'Data','range':'A1:A1','columns':['A'],
                'headers':[{'type':'text','value':'标题'}],'output_rows':0,'rows':[],
                'note':'备注'*20000}
        export_result(result,self.output)
        book=Workbook.load(self.output)
        self.assertEqual(self.snapshot(book)['note'],result['note'])
        self.assertTrue(all(len(cell['value'])<=32767 for cell in book.cells['Metadata'].values() if isinstance(cell['value'],str)))
        self.assertEqual(book.cells['Results']['B1']['value'],'A · 标题')

    def test_unknown_and_malformed_schemas_refused(self):
        cases=[{'status':'compared'}, {'status':'selected','sheet':'D','range':'A1','columns':['A'],'rows':[],'output_rows':1},
               {'status':'selected','sheet':'D','range':'A1','columns':['A'],'headers':[1],'rows':[],'output_rows':0},
               {'status':'selected','sheet':'D','range':'A1','columns':['A'],'headers':None,'output_rows':1,
                'rows':[{'source_row':1,'cells':[{'type':'number','value':None}]}]}]
        for result in cases:
            with self.subTest(result=result),self.assertRaises(ValueError):
                export_result(result,self.output)
            self.assertFalse(self.output.exists())

    def test_cell_budget_and_existing_output_protected(self):
        result={'status':'selected','sheet':'D','range':'A1','columns':['A'],'rows':[],'output_rows':0}
        with patch('xlsx_result_export.MAX_CELLS',5),self.assertRaises(ValueError):
            export_result(result,self.output)
        self.assertFalse(self.output.exists())
        self.output.write_bytes(b'existing')
        with self.assertRaises(ValueError):
            export_result(result,self.output)
        self.assertEqual(self.output.read_bytes(),b'existing')

    def test_cli_json_pipeline(self):
        chinese_book().save(self.input)
        result=select_rows(self.input,'报表',{},header_rows='3:5')
        source=self.input.with_suffix('.json')
        source.write_text(json.dumps(result,default=str))
        run=subprocess.run([sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_result_export.py')),
                            str(source),str(self.output)],capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stderr+run.stdout)
        self.assertEqual(json.loads(run.stdout)['formulas_created'],0)
        self.assertEqual(Workbook.load(self.output).cells['Results']['A2']['value'],6)


if __name__=='__main__':
    unittest.main()
