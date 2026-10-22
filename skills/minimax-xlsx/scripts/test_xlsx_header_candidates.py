#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parent))
from test_xlsx_headers import chinese_book, merged_book
from xlsx_header_candidates import suggest


class CandidateTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path=Path(self.temp.name)/'input.xlsx'

    def save(self,cells,merges=()):
        merged_book(cells,merges).save(self.path)

    def test_multilevel_candidate_excludes_title_and_note(self):
        chinese_book().save(self.path)
        before=self.path.read_bytes()
        result=suggest(self.path,'报表')
        self.assertIn('3:5',[item['header_rows'] for item in result['candidates']])
        self.assertFalse(any(item['header_rows'].startswith('1:') for item in result['candidates']))
        self.assertTrue(result['requires_confirmation'])
        self.assertFalse(result['automatic_boundary_applied'])
        self.assertEqual(self.path.read_bytes(),before)

    def test_headerless_numeric_numeric_header_and_text_only_are_ambiguous(self):
        for cells in ({'A1':1,'B1':2,'A2':3,'B2':4},
                      {'A1':2024,'B1':2025,'A2':10,'B2':20},
                      {'A1':'甲','B1':'乙','A2':'丙','B2':'丁'}):
            self.save(cells)
            result=suggest(self.path,'报表')
            self.assertEqual(result['status'],'ambiguous')
            self.assertTrue(any('insufficient' in text for text in result['ambiguities']))
            self.path.unlink()

    def test_repeated_headers_and_multiple_blocks_report_ambiguity(self):
        self.save({'A1':'部门','B1':'金额','A2':'甲','B2':1,
                   'A4':'部门','B4':'金额','A5':'乙','B5':2})
        result=suggest(self.path,'报表')
        self.assertEqual(len(result['blocks']),2)
        self.assertEqual(result['status'],'ambiguous')
        self.assertTrue(any('repeated' in text for text in result['ambiguities']))

    def test_bounds_empty_and_clipping(self):
        self.save({'A1':'名称','B1':'金额','A2':'甲','B2':1,'A65':'外部'})
        result=suggest(self.path,'报表',scan_rows=4)
        self.assertTrue(result['scan_clipped'])
        for options in ({'scan_rows':65},{'scan_columns':0},{'max_depth':5}):
            with self.assertRaises(ValueError):
                suggest(self.path,'报表',**options)
        self.path.unlink()
        self.save({})
        self.assertEqual(suggest(self.path,'报表')['candidates'],[])

    def test_cli_json(self):
        self.save({'A1':'名称','B1':'金额','A2':'甲','B2':1})
        run=subprocess.run([sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_header_candidates.py')),str(self.path),'--sheet','报表'],capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stderr)
        self.assertTrue(json.loads(run.stdout)['requires_confirmation'])


if __name__=='__main__':
    unittest.main()
