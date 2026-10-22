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
from stdlib_xlsx import Workbook
from test_xlsx_headers import merged_book,chinese_book
from xlsx_advanced import describe,compare
from xlsx_tabular import column_name


class AdvancedTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path=Path(self.temp.name)/'input.xlsx'

    def save(self,rows):
        cells={column_name(col)+str(row):value for row,values in enumerate(rows,1) for col,value in enumerate(values,1)}
        merged_book(cells,[]).save(self.path)

    def test_known_median_quantiles_sample_and_population(self):
        self.save([['Value'],[1],[2],[3],[4]])
        before=self.path.read_bytes()
        result=describe(self.path,'报表',{'columns':['A'],'quantiles':[0,Decimal('.25'),Decimal('.5'),1]})
        column=result['columns'][0]
        self.assertEqual(column['median']['value'],Decimal('2.5'))
        self.assertEqual([item['value'] for item in column['quantiles']],[1,Decimal('1.75'),Decimal('2.5'),4])
        self.assertEqual(column['variance']['population']['variance']['value'],Decimal('1.25'))
        self.assertEqual(column['variance']['sample']['variance']['value'],Decimal('1.666666666666666666666666666666667'))
        self.assertTrue(column['variance']['sample']['variance']['rounded'])
        self.assertEqual(self.path.read_bytes(),before)

    def test_zero_and_single_numeric_sample_sizes(self):
        self.save([['Value'],[None],['']])
        result=describe(self.path,'报表',{'columns':['A']})['columns'][0]
        self.assertEqual(result['count'],0)
        self.assertIsNone(result['median'])
        self.assertIsNone(result['variance']['population']['variance'])
        self.path.unlink()
        self.save([['Value'],[7]])
        result=describe(self.path,'报表',{'columns':['A']})['columns'][0]
        self.assertEqual(result['variance']['sample']['status'],'insufficient_numeric_samples')
        self.assertEqual(result['variance']['population']['standard_deviation']['value'],0)

    def test_missing_and_non_numeric_are_explicit(self):
        self.save([['V'],[1],['2'],[True],[None],['']])
        with self.assertRaisesRegex(ValueError,'Non-numeric value at A3'):
            describe(self.path,'报表',{'columns':['A']})
        result=describe(self.path,'报表',{'columns':['A'],'non_numeric':'omit'})['columns'][0]
        self.assertEqual((result['count'],result['non_numeric_omitted'],result['missing_omitted']),(1,2,2))
        with self.assertRaisesRegex(ValueError,'Missing value at A5'):
            describe(self.path,'报表',{'columns':['A'],'non_numeric':'omit','missing':'reject'})

    def test_decimal_precision_and_no_cancellation_in_variance(self):
        self.save([['V'],[Decimal('1000000000000000000000000000001')],
                   [Decimal('1000000000000000000000000000003')]])
        result=describe(self.path,'报表',{'columns':['A'],'precision':34})['columns'][0]
        self.assertEqual(result['variance']['sample']['variance']['value'],2)
        self.assertEqual(result['variance']['population']['variance']['value'],1)
        self.assertFalse(result['variance']['sample']['variance']['rounded'])

    def test_stddev_exact_and_inexact_flags(self):
        self.save([['V'],[0],[2]])
        result=describe(self.path,'报表',{'columns':['A']})['columns'][0]
        self.assertEqual(result['variance']['population']['standard_deviation'],{'value':Decimal(1),'rounded':False})
        sample=result['variance']['sample']['standard_deviation']
        self.assertTrue(sample['rounded'])
        self.assertTrue(Decimal('1.414213562373095')<sample['value']<Decimal('1.414213562373096'))

    def test_multilevel_metric_selection(self):
        chinese_book().save(self.path)
        result=describe(self.path,'报表',{'columns':[{'path':['2024年','销售','金额']}]},header_rows='3:5')
        self.assertEqual(result['columns'][0]['column'],'B')
        self.assertEqual(result['columns'][0]['median']['value'],10)
        self.assertIn('header_schema',result)

    def test_invalid_descriptive_policies_quantiles_and_precision(self):
        self.save([['V'],[1]])
        for config in ({'columns':['A'],'quantiles':[True]}, {'columns':['A'],'quantiles':['0.5']},
                       {'columns':['A'],'quantiles':[Decimal('1.1')]},{'columns':['A'],'quantiles':[0,0]},
                       {'columns':['A'],'precision':4},{'columns':['A'],'variance':'guess'},
                       {'columns':['A'],'missing':'zero'},{'columns':['A'],'unknown':1}):
            with self.subTest(config=config),self.assertRaises(ValueError):
                describe(self.path,'报表',config)

    def test_exact_arithmetic_limit_fails_instead_of_silent_rounding(self):
        self.save([['V'],[Decimal('1.'+'0'*1100+'1')],[2]])
        with self.assertRaisesRegex(ValueError,'1000-digit'):
            describe(self.path,'报表',{'columns':['A']})

    def period_config(self,**overrides):
        config={'entity_keys':['A'],'period':'B','metrics':['C'],'period_format':'YYYY-MM'}
        config.update(overrides)
        return config

    def test_mom_yoy_exact_calendar_alignment_and_input_order(self):
        self.save([['Entity','Month','Amount'],['001','2024-01',100],['001','2025-02',150],
                   ['001','2025-01',120],['002','2025-02',999]])
        before=self.path.read_bytes()
        result=compare(self.path,'报表',self.period_config())
        feb=[row for row in result['rows'] if row['source_row']==3]
        self.assertEqual(feb[0]['base_period'],'2025-01')
        self.assertEqual(feb[0]['base_source_row'],4)
        self.assertEqual(feb[0]['metrics'][0]['difference'],30)
        self.assertEqual(feb[0]['metrics'][0]['growth_rate']['value'],Decimal('.25'))
        self.assertEqual(feb[1]['metrics'][0]['status'],'missing_period')
        january=[row for row in result['rows'] if row['source_row']==4 and row['comparison']=='yoy'][0]
        self.assertEqual(january['metrics'][0]['growth_rate']['value'],Decimal('.2'))
        self.assertEqual(self.path.read_bytes(),before)

    def test_missing_month_does_not_use_previous_record(self):
        self.save([['E','P','V'],['A','2025-01',10],['A','2025-03',30]])
        result=compare(self.path,'报表',self.period_config(comparisons=['mom']))
        march=result['rows'][1]
        self.assertEqual(march['base_period'],'2025-02')
        self.assertIsNone(march['base_source_row'])
        self.assertEqual(march['metrics'][0]['status'],'missing_period')

    def test_zero_and_negative_bases_have_diagnostic_semantics(self):
        self.save([['E','P','V'],['A','2025-01',0],['A','2025-02',10],
                   ['B','2025-01',-100],['B','2025-02',-50]])
        result=compare(self.path,'报表',self.period_config(comparisons=['mom']))
        zero=result['rows'][1]['metrics'][0]
        self.assertEqual(zero['status'],'zero_base')
        self.assertEqual(zero['difference'],10)
        self.assertIsNone(zero['growth_rate'])
        negative=result['rows'][3]['metrics'][0]
        self.assertTrue(negative['negative_base'])
        self.assertEqual(negative['growth_rate']['value'],Decimal('-.5'))

    def test_duplicate_periods_invalid_months_and_missing_entities_fail(self):
        for rows in ([['E','P','V'],['A','2025-01',1],['A','2025-01',2]],
                     [['E','P','V'],['A','2025-13',1]], [['E','P','V'],['A','0000-01',1]],
                     [['E','P','V'],['A',202501,1]], [['E','P','V'],[None,'2025-01',1]]):
            self.save(rows)
            with self.subTest(rows=rows),self.assertRaises(ValueError):
                compare(self.path,'报表',self.period_config())
            self.path.unlink()

    def test_missing_and_text_metrics_are_diagnosed_without_filling(self):
        self.save([['E','P','V'],['A','2025-01',None],['A','2025-02',10],['A','2025-03','12']])
        with self.assertRaisesRegex(ValueError,'C4'):
            compare(self.path,'报表',self.period_config(comparisons=['mom']))
        result=compare(self.path,'报表',self.period_config(comparisons=['mom'],non_numeric='diagnose'))
        self.assertEqual([row['metrics'][0]['status'] for row in result['rows']],['current_missing','base_missing','current_non_numeric'])

    def test_period_comparison_multilevel_selectors(self):
        book=merged_book({'A1':'维度','A2':'实体','B2':'期间','C1':'指标',
                          'A3':'A','B3':'2025-01','C3':10,'A4':'A','B4':'2025-02','C4':15},['A1:B1','C1:C2'])
        book.save(self.path)
        config=self.period_config(entity_keys=[{'path':['维度','实体']}],period={'path':['维度','期间']},
                                  metrics=[{'path':['指标']}],comparisons=['mom'])
        result=compare(self.path,'报表',config,header_rows='1:2')
        self.assertEqual(result['rows'][1]['metrics'][0]['growth_rate']['value'],Decimal('.5'))

    def test_year_boundary_and_growth_rounding(self):
        self.save([['E','P','V'],['A','2024-12',3],['A','2025-01',4]])
        result=compare(self.path,'报表',self.period_config(comparisons=['mom']))
        january=result['rows'][1]
        self.assertEqual(january['base_period'],'2024-12')
        self.assertTrue(january['metrics'][0]['growth_rate']['rounded'])
        self.assertEqual(january['metrics'][0]['growth_rate']['value'],Decimal('0.3333333333333333333333333333333333'))

    def test_advanced_formula_cache_policy(self):
        book=Workbook.create({'sheets':[{'name':'报表','cells':{'A1':{'value':'期间'},'B1':{'value':'值'},
            'A2':{'value':'2025-01'},'B2':{'formula':'2+3'}}}]})
        book.save(self.path)
        for policy in ('reject','cached'):
            with self.assertRaises(ValueError):
                describe(self.path,'报表',{'columns':['B']},formulas=policy)
        book.calculate()
        cached=self.path.with_name('cached.xlsx')
        book.save(cached)
        stats=describe(cached,'报表',{'columns':['B']},formulas='cached')
        self.assertEqual(stats['unverified_formula_caches'],1)
        config={'entity_keys':[],'period':'A','metrics':['B'],'period_format':'YYYY-MM'}
        compared=compare(cached,'报表',config,formulas='cached')
        self.assertEqual(compared['unverified_formula_caches'],1)
        self.assertEqual(compared['rows'][0]['metrics'][0]['status'],'missing_period')

    def test_cli_describe_compare_and_controlled_failure(self):
        self.save([['E','P','V'],['A','2025-01',1],['A','2025-02',2]])
        query=self.path.with_suffix('.json')
        command=[sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_advanced.py'))]
        for operation,config,status in (('describe',{'columns':['C']},'described'),('compare',self.period_config(),'compared')):
            query.write_text(json.dumps(config))
            run=subprocess.run(command+[operation,str(self.path),'--sheet','报表','--query',str(query)],capture_output=True,text=True)
            self.assertEqual(run.returncode,0,run.stderr+run.stdout)
            self.assertEqual(json.loads(run.stdout)['status'],status)
        query.write_text('{"columns":["C"],"precision":1}')
        run=subprocess.run(command+['describe',str(self.path),'--sheet','报表','--query',str(query)],capture_output=True,text=True)
        self.assertEqual(run.returncode,1)
        self.assertEqual(json.loads(run.stdout)['status'],'failed')


if __name__=='__main__':
    unittest.main()
