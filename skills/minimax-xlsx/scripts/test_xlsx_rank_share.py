#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
from decimal import Decimal
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parent))
from test_xlsx_headers import merged_book,chinese_book
from xlsx_rank_share import analyze
from xlsx_result_export import export_result
from stdlib_xlsx import Workbook


class RankShareTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.source=Path(self.temp.name)/'input.xlsx'

    def save(self,values,groups=None):
        cells={'A1':'组','B1':'值'}
        for row,value in enumerate(values,2):
            cells['A'+str(row)]='甲' if groups is None else groups[row-2]
            cells['B'+str(row)]=value
        merged_book(cells,[]).save(self.source)

    def ranks(self,result):
        return [row['cells'][-4]['value'] for row in result['rows']]

    def shares(self,result):
        return [row['cells'][-3]['value'] for row in result['rows']]

    def test_competition_dense_ordinal_and_direction(self):
        self.save([10,10,5])
        before=self.source.read_bytes()
        for method,expected in (('competition',[1,1,3]),('dense',[1,1,2]),('ordinal',[1,2,3])):
            result=analyze(self.source,'报表',{'measure':'B','rank':method})
            self.assertEqual(self.ranks(result),expected)
            self.assertEqual([row['source_row'] for row in result['rows']],[2,3,4])
        result=analyze(self.source,'报表',{'measure':'B','descending':False})
        self.assertEqual(self.ranks(result),[2,2,1])
        self.assertEqual(self.source.read_bytes(),before)

    def test_group_and_filter_define_denominator(self):
        self.save([10,30,5,15],['甲','甲','乙','乙'])
        result=analyze(self.source,'报表',{'group_by':['A'],'measure':'B'})
        self.assertEqual(self.shares(result),[Decimal('.25'),Decimal('.75'),Decimal('.25'),Decimal('.75')])
        result=analyze(self.source,'报表',{'group_by':['A'],'measure':'B','filters':[{'column':'B','op':'ge','value':15}]})
        self.assertEqual(self.shares(result),[1,1])
        self.assertEqual(result['filtered_out_rows'],2)
        self.assertEqual([group['numeric_total'] for group in result['group_statistics']],[30,15])

    def test_zero_and_signed_totals_are_not_absolute_contributions(self):
        self.save([1,-1])
        result=analyze(self.source,'报表',{'measure':'B'})
        self.assertEqual(self.ranks(result),[1,2])
        self.assertEqual(self.shares(result),[None,None])
        self.assertTrue(all(row['cells'][-2]['value']=='zero_total' for row in result['rows']))
        self.source.unlink(); self.save([10,-2])
        result=analyze(self.source,'报表',{'measure':'B'})
        self.assertEqual(self.shares(result),[Decimal('1.25'),Decimal('-.25')])
        self.assertEqual(result['group_statistics'][0]['negative_numeric_count'],1)
        self.assertTrue(result['share_definition']['not_growth_contribution'])

    def test_missing_and_nonnumeric_policy_retains_diagnostic_rows(self):
        self.save([1,None,'2',True])
        with self.assertRaisesRegex(ValueError,'B4'):
            analyze(self.source,'报表',{'measure':'B'})
        result=analyze(self.source,'报表',{'measure':'B','non_numeric':'omit'})
        self.assertEqual(self.ranks(result),[1,None,None,None])
        self.assertEqual(self.shares(result),[1,None,None,None])
        self.assertEqual([row['cells'][-2]['value'] for row in result['rows']],['ok','missing_omitted','non_numeric_omitted','non_numeric_omitted'])
        with self.assertRaisesRegex(ValueError,'B3'):
            analyze(self.source,'报表',{'measure':'B','missing':'reject','non_numeric':'omit'})

    def test_precision_and_entirely_missing_group(self):
        self.save([1,2,None],['甲','甲','乙'])
        result=analyze(self.source,'报表',{'group_by':['A'],'measure':'B'})
        self.assertEqual(self.shares(result)[0],Decimal('0.3333333333333333333333333333333333'))
        self.assertTrue(result['rows'][0]['cells'][-1]['value'])
        self.assertEqual(result['group_statistics'][1]['status'],'insufficient_numeric_samples')
        self.assertIsNone(result['group_statistics'][1]['numeric_total'])

    def test_multilevel_paths_and_readable_export(self):
        chinese_book().save(self.source)
        result=analyze(self.source,'报表',{'group_by':[{'path':['部门']}],
                                          'measure':{'path':['2024年','风控','金额']}},header_rows='3:5')
        self.assertEqual(self.ranks(result),[2,1,1])
        output=self.source.with_name('ranked.xlsx')
        export_result(json.loads(json.dumps(result,default=str)),output)
        book=Workbook.load(output)
        self.assertEqual(book.cells['Results']['G2']['value'],2)
        self.assertTrue(any(cell['value']=='calc:rank' for cell in book.cells['Columns'].values()))

    def test_invalid_definitions_and_empty_selection(self):
        self.save([1])
        for config in ({'measure':'B','rank':'random'},{'measure':'B','descending':'yes'},
                       {'measure':'B','denominator':'growth_contribution'},{'measure':'B','precision':2},{}):
            with self.subTest(config=config),self.assertRaises(ValueError):
                analyze(self.source,'报表',config)
        result=analyze(self.source,'报表',{'measure':'B','filters':[{'column':'B','op':'gt','value':99}]})
        self.assertEqual(result['rows'],[])
        self.assertEqual(result['group_statistics'],[])

    def test_cli_rank_then_export(self):
        self.save([10,30])
        query=self.source.with_suffix('.json'); query.write_text('{"group_by":["A"],"measure":"B"}')
        scripts=Path(__file__).resolve().parent
        run=subprocess.run([sys.executable,'-I','-S','-B',str(scripts/'xlsx_rank_share.py'),str(self.source),
                            '--sheet','报表','--query',str(query)],capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stderr+run.stdout)
        result=self.source.with_name('result.json'); result.write_text(run.stdout)
        output=self.source.with_name('result.xlsx')
        run=subprocess.run([sys.executable,'-I','-S','-B',str(scripts/'xlsx_result_export.py'),str(result),str(output)],capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stderr+run.stdout)
        book=Workbook.load(output)
        self.assertEqual(book.cells['Results']['D2']['value'],2)
        self.assertEqual(book.cells['Results']['E2']['value'],Decimal('.25'))


if __name__=='__main__':
    unittest.main()
