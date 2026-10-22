#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
import contextlib
import io
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook, read_spec, xml
from test_xlsx_headers import chinese_book
from xlsx_analysis import main as analysis_main, profile
from xlsx_header_candidates import suggest
from xlsx_rows import read_query


class BoundaryTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path=Path(self.temp.name)/'input.xlsx'
        Workbook.create({'sheets':[{'name':'Data','cells':{'A1':{'value':'ID'},'A2':{'value':1}}}]}).save(self.path)
        self.query=self.path.with_suffix('.json')
        self.scripts=Path(__file__).resolve().parent
        self.base=[sys.executable,'-I','-S','-B']

    def run_cli(self,arguments):
        run=subprocess.run(self.base+[str(self.scripts/arguments[0])]+arguments[1:],capture_output=True,text=True)
        self.assertEqual(run.returncode,1,run.stdout+run.stderr)
        self.assertNotIn('Traceback',run.stderr)
        result=json.loads(run.stdout)
        self.assertEqual(result['status'],'failed')
        return result

    def test_invalid_range_types_fail_with_json_instead_of_traceback(self):
        for value in (123,0,[],{},''):
            self.query.write_text(json.dumps({'left':{'sheet':'Data','keys':['A'],'range':value},
                                              'right':{'sheet':'Data','keys':['A']},'relationship':'one_to_one'}))
            result=self.run_cli(['xlsx_match.py',str(self.path),str(self.path),'--query',str(self.query)])
            self.assertIn('Range',result['message'])
            with self.assertRaises(ValueError):
                profile(self.path,'Data',cell_range=value)

    def test_deep_query_is_structured_error_across_all_query_clis(self):
        self.query.write_text('['*2000+'0'+']'*2000)
        common=[str(self.path),'--sheet','Data','--query',str(self.query)]
        arguments=[['xlsx_rows.py']+common,
                   ['xlsx_advanced.py','describe']+common,
                   ['xlsx_advanced.py','compare']+common,
                   ['xlsx_analysis.py','summarize']+common,
                   ['xlsx_match.py',str(self.path),str(self.path),'--query',str(self.query)]]
        for arguments in arguments:
            with self.subTest(command=arguments[0]):
                self.assertIn('nesting',self.run_cli(arguments)['message'])

    def test_deep_creation_spec_leaves_no_output(self):
        self.query.write_text('['*2000+'0'+']'*2000)
        output=self.path.with_name('output.xlsx')
        self.run_cli(['stdlib_xlsx.py','create',str(self.query),str(output)])
        self.assertFalse(output.exists())

    def test_query_budget_is_enforced_on_read_not_file_stat(self):
        self.query.write_text(' '*65537+'{"measure":"A"}')
        # Simulate metadata claiming a smaller size: actual bytes must still be bounded.
        class SmallStat:
            st_size=1
        with patch.object(Path,'stat',return_value=SmallStat()), patch.object(sys,'argv',
            ['xlsx_analysis.py','summarize',str(self.path),'--sheet','Data','--query',str(self.query)]):
            output=io.StringIO()
            with contextlib.redirect_stdout(output):
                self.assertEqual(analysis_main(),1)
            self.assertIn('byte limit',json.loads(output.getvalue())['message'])

    def test_bounded_json_preserves_bom_decimal_duplicate_and_nan_rules(self):
        self.query.write_bytes(b'\xef\xbb\xbf{"n":0.1,"id":"001"}')
        data=read_query(self.query)
        self.assertEqual(str(data['n']),'0.1')
        self.assertEqual(data['id'],'001')
        for text in ('{"a":1,"a":2}','{"n":NaN}'):
            self.query.write_text(text)
            with self.assertRaises(ValueError):
                read_query(self.query)
        self.query.write_text('123456789')
        with self.assertRaises(ValueError):
            read_spec(self.query,max_bytes=8)

    def test_header_candidate_structure_is_parsed_once_for_multiple_ranges(self):
        self.path.unlink()
        chinese_book().save(self.path)
        with patch('xlsx_header_candidates.xml',wraps=xml) as parsed, patch('xlsx_headers.xml',wraps=xml) as repeated:
            result=suggest(self.path,'报表')
        self.assertGreater(result['candidate_count'],1)
        self.assertEqual(parsed.call_count,1)
        self.assertEqual(repeated.call_count,0)
        self.assertIn('3:5',[candidate['header_rows'] for candidate in result['candidates']])


if __name__=='__main__':
    unittest.main()
