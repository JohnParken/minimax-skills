#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""All management fixtures are fictional, not actual bank/customer information."""
from copy import deepcopy
from decimal import Decimal
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parent))
from test_xlsx_headers import merged_book
from xlsx_management import build,column_id
from xlsx_result_export import export_result
from stdlib_xlsx import Workbook
from xlsx_tabular import column_name


def config_example():
    dimensions=dict(zip(('period','institution','line','currency','version','product'),'ABCDEF'))
    def metric(identifier,label,kind='stock',basis='point_in_time'):
        return {'metric_id':identifier,'label':label,'unit':'万元','kind':kind,'time_basis':basis,
                'aggregation':'SUM','additive_dimensions':['product']}
    metrics=[metric('balance','虚构贷款余额'),metric('bad','虚构不良余额'),
             {'metric_id':'risk_rate','label':'虚构不良比率','unit':'%','kind':'ratio','time_basis':'ratio','aggregation':'RATIO',
              'additive_dimensions':[],'numerator':'bad','denominator':'balance','scale':100},
             metric('income','虚构当月收入','flow','period_flow')]
    return {'dimensions':dimensions,'dimension_labels':{'period':'期间','institution':'机构','line':'条线','currency':'币种','version':'口径版本','product':'产品'},
            'row_dimensions':['period','institution','line','currency','version'],
            'fact_grain':list(dimensions)+['metric_id'],'metric_column':'G','value_column':'H','unit_column':'I','basis_column':'K',
            'record_scope':{'column':'J','value':'leaf','report_level':'branch'},'definition_version':'v1','period_format':'YYYY-MM',
            'metrics':metrics,'missing':'propagate','precision':34,'targets':[]}


def fictional_facts():
    facts=[]
    for product,values in (('P1',{'balance':100,'bad':10,'income':3}),('P2',{'balance':300,'bad':0,'income':5})):
        for metric,value in values.items():
            facts.append(['2025-01','00012345678901234567890','retail','CNY','v1',product,metric,value,'万元','leaf',
                          'period_flow' if metric=='income' else 'point_in_time'])
    return facts


def spec_for(facts):
    rows=[['期间','机构','条线','币种','版本','产品','指标','值','单位','记录层级','时间口径']]+facts
    return {'sheets':[{'name':'Facts','cells':{column_name(col)+str(row):{'value':value}
                                            for row,values in enumerate(rows,1) for col,value in enumerate(values,1)}}]}


class ManagementTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.source=Path(self.temp.name)/'fictional.xlsx'
        self.config=config_example()

    def save(self,facts):
        Workbook.create(spec_for(facts)).save(self.source)

    def value(self,result,metric,field='actual',row=0):
        index=result['columns'].index(column_id('metric',metric,field))
        return result['rows'][row]['cells'][index]['value']

    def target(self,metric,value,mode='at_least',**changes):
        definition=next(item for item in self.config['metrics'] if item['metric_id']==metric)
        result={'dimensions':{'period':'2025-01','institution':'00012345678901234567890','line':'retail','currency':'CNY','version':'v1'},
                'metric_id':metric,'unit':definition['unit'],'kind':definition['kind'],'time_basis':definition['time_basis'],
                'definition_version':'v1','report_level':'branch','mode':mode,'value':value}
        result.update(changes)
        return result

    def test_weighted_ratio_flow_and_target_wide_end_to_end(self):
        self.save(fictional_facts())
        original=self.source.read_bytes()
        self.config['targets']=[self.target('balance',500),self.target('risk_rate',3,'at_most')]
        result=build(self.source,'Facts',self.config)
        self.assertEqual(result['output_rows'],1)
        self.assertEqual(self.value(result,'balance'),400)
        self.assertEqual(self.value(result,'risk_rate'),Decimal('2.5'))
        self.assertEqual(self.value(result,'income'),8)
        self.assertEqual(self.value(result,'balance','achievement'),Decimal('.8'))
        self.assertEqual(self.value(result,'balance','gap'),100)
        self.assertEqual(self.value(result,'risk_rate','achievement'),Decimal('1.2'))
        self.assertEqual(self.value(result,'risk_rate','gap'),0)
        output=self.source.with_name('wide.xlsx')
        export_result(json.loads(json.dumps(result,default=str)),output)
        book=Workbook.load(output)
        self.assertEqual(list(book.sheets),['Wide','Columns','Metadata','Warnings','Definitions','Validation','Lineage'])
        self.assertEqual(book.cells['Wide']['B2']['value'],'00012345678901234567890')
        self.assertEqual(book.cells['Lineage']['F2']['value'],'balance')
        self.assertTrue(all(cell['formula'] is None for cells in book.cells.values() for cell in cells.values()))
        self.assertEqual(self.source.read_bytes(),original)

    def test_periods_ytd_and_currency_remain_separate(self):
        facts=fictional_facts()
        for period,currency in (('2025-02','CNY'),('2025-01','USD')):
            facts += [[period]+row[1:3]+[currency]+row[4:] for row in fictional_facts()]
        self.save(facts)
        result=build(self.source,'Facts',self.config)
        self.assertEqual(result['output_rows'],3)
        self.assertEqual([self.value(result,'balance',row=index) for index in range(3)],[400,400,400])
        self.source.unlink()
        facts=fictional_facts()
        for row in facts:
            if row[6]=='income': row[10]='ytd'
        self.config['metrics'][-1]['time_basis']='ytd'
        self.save(facts)
        self.assertEqual(self.value(build(self.source,'Facts',self.config),'income'),8)

    def test_protected_dimensions_and_additivity_permissions(self):
        self.save(fictional_facts())
        for name in ('period','currency','version'):
            config=deepcopy(self.config); config['row_dimensions'].remove(name)
            with self.subTest(name=name),self.assertRaises(ValueError):
                build(self.source,'Facts',config)
        config=deepcopy(self.config); config['metrics'][0]['additive_dimensions']=[]
        with self.assertRaisesRegex(ValueError,'does not permit'):
            build(self.source,'Facts',config)

    def test_duplicate_grain_and_mixed_summary_scope_fail(self):
        facts=fictional_facts(); self.save(facts+[facts[0]])
        with self.assertRaisesRegex(ValueError,'Duplicate fact grain'):
            build(self.source,'Facts',self.config)
        self.source.unlink(); facts=fictional_facts(); facts[0][9]='summary'; self.save(facts)
        with self.assertRaisesRegex(ValueError,'scope'):
            build(self.source,'Facts',self.config)

    def test_unit_basis_version_and_numeric_text_mismatches(self):
        for index,value in ((8,'元'),(10,'ytd'),(4,'v2'),(7,'100'),(1,123)):
            facts=fictional_facts(); facts[0][index]=value; self.save(facts)
            with self.subTest(index=index),self.assertRaises(ValueError):
                build(self.source,'Facts',self.config)
            self.source.unlink()

    def test_missing_is_not_zero_and_zero_ratio_denominator_is_diagnostic(self):
        facts=fictional_facts(); facts[0][7]=None; self.save(facts)
        result=build(self.source,'Facts',self.config)
        self.assertIsNone(self.value(result,'balance'))
        self.assertIsNone(self.value(result,'risk_rate'))
        self.assertIn('missing_values:balance',result['rows'][0]['quality_flags'])
        with self.assertRaises(ValueError):
            build(self.source,'Facts',{**self.config,'missing':'reject'})
        self.source.unlink(); facts=fictional_facts()
        for row in facts:
            if row[6]=='balance': row[7]=0
        self.save(facts)
        result=build(self.source,'Facts',self.config)
        self.assertIsNone(self.value(result,'risk_rate'))
        self.assertIn('zero_denominator:risk_rate',result['rows'][0]['quality_flags'])

    def test_missing_metric_and_raw_ratio_facts_refused(self):
        facts=[row for row in fictional_facts() if row[6]!='income']; self.save(facts)
        result=build(self.source,'Facts',self.config)
        self.assertIsNone(self.value(result,'income'))
        self.assertIn('missing_metric:income',result['rows'][0]['quality_flags'])
        self.source.unlink(); facts=fictional_facts(); row=deepcopy(facts[0]); row[6]='risk_rate'; row[8]='%'; row[10]='ratio'; self.save(facts+[row])
        with self.assertRaisesRegex(ValueError,'Raw ratio'):
            build(self.source,'Facts',self.config)

    def test_targets_match_unit_type_version_and_grain(self):
        self.save(fictional_facts())
        for changes in ({'unit':'元'},{'kind':'flow'},{'time_basis':'ytd'},{'definition_version':'v2'},
                        {'report_level':'regional'},{'mode':'automatic'}):
            config=deepcopy(self.config); config['targets']=[self.target('balance',500,**changes)]
            with self.subTest(changes=changes),self.assertRaises(ValueError):
                build(self.source,'Facts',config)
        config=deepcopy(self.config); target=self.target('balance',500); config['targets']=[target,target]
        with self.assertRaisesRegex(ValueError,'Duplicate target'):
            build(self.source,'Facts',config)

    def test_zero_negative_targets_and_missing_targets_are_diagnostic(self):
        self.save(fictional_facts())
        for value,flag in ((0,'zero_target'),(-10,'negative_target')):
            config=deepcopy(self.config); config['targets']=[self.target('balance',value)]
            result=build(self.source,'Facts',config)
            self.assertIsNone(self.value(result,'balance','achievement'))
            self.assertIn(flag+':balance',result['rows'][0]['quality_flags'])
        target=self.target('balance',500); target['dimensions']['institution']='not-in-input'
        result=build(self.source,'Facts',{**self.config,'targets':[target]})
        self.assertIsNone(self.value(result,'balance','target'))
        self.assertTrue(any(item['check']=='unmatched_targets' for item in result['validation']))

    def test_min_max_count_have_declared_reduction_semantics(self):
        config=deepcopy(self.config)
        config['metrics']=[]; facts=[]
        for identifier,method,unit in (('low','MIN','万元'),('high','MAX','万元'),('observations','COUNT','records')):
            config['metrics'].append({'metric_id':identifier,'label':identifier,'unit':unit,'kind':'stock','time_basis':'point_in_time',
                                      'aggregation':method,'additive_dimensions':['product'] if method=='COUNT' else [],
                                      'reduction_dimensions':['product'] if method!='COUNT' else []})
            for product,value in (('P1',10),('P2',20)):
                row=deepcopy(fictional_facts()[0]); row[5]=product; row[6]=identifier; row[7]=value; row[8]=unit; facts.append(row)
        self.save(facts)
        result=build(self.source,'Facts',config)
        self.assertEqual(self.value(result,'low'),10)
        self.assertEqual(self.value(result,'high'),20)
        self.assertEqual(self.value(result,'observations'),2)

    def test_reversible_ids_and_unsupported_hierarchy_or_averages(self):
        self.assertEqual(json.loads(column_id('metric','风险/率:[]','actual')),['metric','风险/率:[]','actual'])
        self.save(fictional_facts())
        for extra in ({'hierarchy':{}},{'daily_average':True}):
            with self.assertRaises(ValueError):
                build(self.source,'Facts',{**self.config,**extra})
        config=deepcopy(self.config); config['metrics'][0]['aggregation']='AVERAGE'
        with self.assertRaises(ValueError):
            build(self.source,'Facts',config)

    def test_filtered_scope_and_formula_policy(self):
        self.save(fictional_facts())
        config=deepcopy(self.config); config['filters']=[{'column':'F','op':'eq','value':'P1'}]
        result=build(self.source,'Facts',config)
        self.assertEqual(result['selected_facts'],3)
        self.assertEqual(self.value(result,'balance'),100)
        self.assertEqual(len(result['lineage']),3)
        self.assertEqual(result['requested_query']['filters'],config['filters'])
        book=Workbook.load(self.source)
        book.patch({'sheets':[{'name':'Facts','cells':{'H2':{'formula':'50+50'}}}]})
        formula=self.source.with_name('formula.xlsx'); book.save(formula)
        with self.assertRaises(ValueError):
            build(formula,'Facts',self.config)

    def test_ytd_periods_never_accumulate_across_months(self):
        facts=fictional_facts()
        for row in facts:
            if row[6]=='income': row[10]='ytd'
        second=deepcopy(facts)
        for row in second:
            row[0]='2025-02'
            if row[6]=='income': row[7]=20 if row[5]=='P1' else 30
        self.config['metrics'][-1]['time_basis']='ytd'
        self.save(facts+second)
        result=build(self.source,'Facts',self.config)
        self.assertEqual([self.value(result,'income',row=index) for index in range(2)],[8,50])

    def test_signed_denominator_and_inverse_target_zero_actual(self):
        facts=fictional_facts()
        for row in facts:
            if row[6]=='balance': row[7]=-row[7]
        self.save(facts)
        result=build(self.source,'Facts',self.config)
        self.assertIsNone(self.value(result,'risk_rate'))
        self.config['metrics'][2]['allow_negative_denominator']=True
        result=build(self.source,'Facts',self.config)
        self.assertEqual(self.value(result,'risk_rate'),Decimal('-2.5'))
        self.assertIn('negative_denominator_allowed:risk_rate',result['rows'][0]['quality_flags'])
        self.source.unlink(); facts=fictional_facts()
        for row in facts:
            if row[6]=='bad': row[7]=0
        self.save(facts)
        self.config['targets']=[self.target('risk_rate',3,'at_most')]
        result=build(self.source,'Facts',self.config)
        self.assertEqual(self.value(result,'risk_rate'),0)
        self.assertIsNone(self.value(result,'risk_rate','achievement'))
        self.assertIn('nonpositive_actual_for_inverse_achievement:risk_rate',result['rows'][0]['quality_flags'])

    def test_explicit_scope_filter_and_mixed_basis_ratios(self):
        facts=fictional_facts(); summary=deepcopy(facts[0]); summary[9]='summary'
        self.save(facts+[summary])
        self.config['filters']=[{'column':'J','op':'eq','value':'leaf'}]
        result=build(self.source,'Facts',self.config)
        self.assertEqual(result['selected_facts'],6)
        self.assertEqual(result['filtered_out_rows'],1)
        config=deepcopy(self.config); config['metrics'][2]['numerator']='income'
        with self.assertRaisesRegex(ValueError,'Mixed-unit/basis'):
            build(self.source,'Facts',config)

    def test_malformed_wide_export_schema_is_rejected(self):
        output=self.source.with_name('invalid.xlsx')
        result={'status':'management_wide','schema_version':'management_wide_v1','sheet':'Facts','range':'A1',
                'columns':None,'rows':[],'output_rows':0}
        with self.assertRaises(ValueError):
            export_result(result,output)
        self.assertFalse(output.exists())

    def test_cli_wide_xlsx_and_empty_result(self):
        self.save(fictional_facts())
        query=self.source.with_suffix('.json'); query.write_text(json.dumps(self.config))
        output=self.source.with_name('management.xlsx')
        run=subprocess.run([sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_management.py')),
                            str(self.source),'--sheet','Facts','--query',str(query),'--xlsx',str(output)],capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stderr+run.stdout)
        self.assertTrue(json.loads(run.stdout)['export']['output_written'])
        self.assertEqual(Workbook.load(output).cells['Wide']['B2']['value'],'00012345678901234567890')
        config=deepcopy(self.config); config['filters']=[{'column':'F','op':'eq','value':'absent'}]
        result=build(self.source,'Facts',config)
        self.assertEqual(result['rows'],[])
        export_result(result,self.source.with_name('empty.xlsx'))


if __name__=='__main__':
    unittest.main()
