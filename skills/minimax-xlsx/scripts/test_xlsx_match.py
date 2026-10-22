#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parent))
from test_xlsx_headers import merged_book, chinese_book
from xlsx_match import match_tables


class MatchTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.left=Path(self.temp.name)/'left.xlsx'
        self.right=Path(self.temp.name)/'right.xlsx'
        self.config={'left':{'sheet':'报表','keys':['A']},'right':{'sheet':'报表','keys':['A']},'relationship':'one_to_one'}

    def save(self,left,right):
        def cells(rows):
            from xlsx_tabular import column_name
            return {column_name(col)+str(row):value for row,values in enumerate(rows,1) for col,value in enumerate(values,1)}
        merged_book(cells(left),[]).save(self.left)
        merged_book(cells(right),[]).save(self.right)

    def test_full_match_strict_identifier_types_and_unmatched_diagnostics(self):
        self.save([['ID'],['001'],[1],[True]],[['ID'],['001'],['1'],[True]])
        self.config['how']='full'
        before=self.left.read_bytes(),self.right.read_bytes()
        report=match_tables(self.left,self.right,self.config)
        self.assertEqual([(r['left_row'],r['right_row']) for r in report['rows']],[(2,2),(3,None),(4,4),(None,3)])
        self.assertEqual(report['left']['unmatched_rows'],1)
        self.assertEqual(report['right']['unmatched_rows'],1)
        self.assertEqual(before,(self.left.read_bytes(),self.right.read_bytes()))

    def test_one_to_many_and_many_to_one_relationship_validation(self):
        self.save([['ID'],['A']],[['ID'],['A'],['A']])
        self.assertEqual(match_tables(self.left,self.right,self.config)['status'],'failed')
        self.config['relationship']='one_to_many'
        report=match_tables(self.left,self.right,self.config)
        self.assertEqual(report['output_rows'],2)
        self.assertEqual(report['right']['duplicate_key_groups'],1)
        self.config['relationship']='many_to_one'
        report=match_tables(self.right,self.left,self.config)
        self.assertEqual(report['output_rows'],2)

    def test_null_keys_never_match_default_and_typed_opt_in(self):
        self.save([['ID'],[None],['']],[['ID'],[None],['']])
        report=match_tables(self.left,self.right,self.config)
        self.assertTrue(all(row['right_row'] is None for row in report['rows']))
        self.assertEqual(report['left']['null_key_rows'],2)
        self.config['nulls']='match'
        report=match_tables(self.left,self.right,self.config)
        self.assertEqual([(r['left_row'],r['right_row']) for r in report['rows']],[(2,2),(3,3)])

    def test_composite_keys_and_multi_level_paths(self):
        chinese_book().save(self.left)
        chinese_book().save(self.right)
        for side in ('left','right'):
            self.config[side].update({'header_rows':'3:5','keys':[{'path':['部门']},{'path':['2024年','风控','金额']}]})
        report=match_tables(self.left,self.right,self.config)
        self.assertEqual(report['output_rows'],3)
        self.assertEqual(report['left']['key_columns'],['A','D'])

    def test_output_limit_inner_and_invalid_configuration(self):
        self.save([['ID'],['A']],[['ID'],['A'],['A']])
        self.config.update({'relationship':'one_to_many','max_rows':1})
        with self.assertRaisesRegex(ValueError,'budget'):
            match_tables(self.left,self.right,self.config)
        self.config['relationship']='many_to_many'
        with self.assertRaises(ValueError):
            match_tables(self.left,self.right,self.config)
        self.config.update({'relationship':'one_to_many','max_rows':10,'how':'inner'})
        self.assertEqual(match_tables(self.left,self.right,self.config)['output_rows'],2)

    def test_formula_key_policy_and_unverified_caches(self):
        from stdlib_xlsx import Workbook
        book=Workbook.create({'sheets':[{'name':'报表','cells':{'A1':{'value':'ID'},'A2':{'formula':'1+1'}}}]})
        book.save(self.left)
        book.save(self.right)
        with self.assertRaises(ValueError):
            match_tables(self.left,self.right,self.config)
        self.config['formulas']='cached'
        with self.assertRaises(ValueError):
            match_tables(self.left,self.right,self.config)
        self.left.unlink()
        self.right.unlink()
        book.calculate()
        book.save(self.left)
        book.save(self.right)
        report=match_tables(self.left,self.right,self.config)
        self.assertEqual(report['output_rows'],1)
        self.assertEqual(report['left']['unverified_formula_caches'],1)
        self.assertEqual(report['right']['unverified_formula_caches'],1)

    def test_cli_json_duplicate_diagnostics(self):
        self.save([['ID'],['A']],[['ID'],['A'],['A']])
        query=self.left.with_suffix('.json')
        query.write_text(json.dumps(self.config))
        run=subprocess.run([sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_match.py')),str(self.left),str(self.right),'--query',str(query)],capture_output=True,text=True)
        self.assertEqual(run.returncode,1)
        report=json.loads(run.stdout)
        self.assertEqual(report['right']['duplicate_key_row_samples'],[[2,3]])
        self.assertEqual(report['rows'],[])


if __name__=='__main__':
    unittest.main()
