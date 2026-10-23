#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
from copy import deepcopy
from decimal import Decimal
from datetime import date,timedelta
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook
from test_xlsx_hierarchy import config_with_tree,two_branches
from test_xlsx_management import spec_for
from xlsx_management import build,column_id
from xlsx_management_policy import analyze,export_policy
from test_xlsx_daily import daily_config,daily_spec
from xlsx_daily import summarize


FIRST='00012345678901234567890'


def policy_config():
    return {'metric_id':'balance','scope':{'report_level':'branch','currency':'CNY','version':'v1','definition_version':'v1','unit':'万元','kind':'stock','time_basis':'point_in_time','period_granularity':'month'},
            'periods':['2025-01','2025-02','2025-03'],
            'institutions':[{'id':FIRST,'level':'branch','peer_group':'comparable-A'},{'id':'0002','level':'branch','peer_group':'comparable-A'}],
            'context':{'line':'retail'},'plan_definition':{'target_type':'balance','unit':'万元','kind':'stock','time_basis':'point_in_time','currency':'CNY','version':'v1','definition_version':'v1','plan_version':'fictional-plan-v1','direction':'at_least'},
            'plans':[{'institution':institution,'period':period,'applicability':'applicable','value':target} for institution,target in [(FIRST,500),('0002',300)] for period in ('2025-01','2025-02','2025-03')],
            'rank_method':'competition'}


class ManagementPolicyTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.folder=Path(self.temp.name); facts=[]
        for period in ('2025-01','2025-02','2025-03'):
            rows=deepcopy(two_branches())
            for row in rows: row[0]=period
            facts += [row for row in rows if row[6]=='balance']
        path=self.folder/'facts.xlsx'; Workbook.create(spec_for(facts)).save(path)
        config=config_with_tree(); config['metrics']=[metric for metric in config['metrics'] if metric['metric_id']=='balance']
        self.source=build(path,'Facts',config); self.config=policy_config()

    def test_explicit_progress_comparable_rank_and_continuity(self):
        result=analyze(self.source,self.config)
        first=result['rows'][0]
        self.assertEqual((first['actual'],first['planned_to_date'],first['favorable_deviation'],first['gap'],first['attainment']),(400,500,-100,100,Decimal('.8')))
        self.assertEqual((first['peer_rank'],first['peer_count']),(1,2))
        self.assertTrue(first['behind_plan'])
        self.assertTrue(all(row['trailing_behind_periods']==3 for row in result['continuity']))
        self.assertTrue(result['excluded_source_records']>0)

    def test_peer_group_isolation_and_duplicate_mapping_rejected(self):
        self.config['institutions'][1]['peer_group']='comparable-B'
        result=analyze(self.source,self.config)
        self.assertTrue(all(row['peer_count']==1 for row in result['rows']))
        self.config['institutions'].append(deepcopy(self.config['institutions'][0]))
        with self.assertRaisesRegex(ValueError,'Duplicate institution'): analyze(self.source,self.config)

    def test_missing_period_is_unknown_not_zero_and_no_linear_plan(self):
        index=self.source['columns'].index(column_id('dimension','period'))
        inst=self.source['columns'].index(column_id('dimension','institution'))
        value=self.source['columns'].index(column_id('metric','balance','actual'))
        row=next(row for row in self.source['rows'] if row['cells'][index]['value']=='2025-02' and row['cells'][inst]['value']==FIRST)
        row['cells'][value]={'type':'blank','value':None}
        result=analyze(self.source,self.config)
        self.assertIsNone(result['continuity'][0]['trailing_behind_periods'])
        self.assertEqual(result['rows'][1]['comparison_status'],'missing_actual')
        self.config['plans']=[plan for plan in self.config['plans'] if not (plan['institution']=='0002' and plan['period']=='2025-02')]
        result=analyze(self.source,self.config)
        self.assertEqual(result['rows'][4]['comparison_status'],'missing_plan')
        self.assertIsNone(result['rows'][4]['planned_to_date'])

    def test_not_applicable_zero_and_negative_plan_are_distinct(self):
        self.config['plans'][0].update(applicability='not_applicable',value=None)
        self.config['plans'][1]['value']=0; self.config['plans'][2]['value']=-1
        result=analyze(self.source,self.config)
        self.assertEqual([row['comparison_status'] for row in result['rows'][:3]],['not_applicable','nonpositive_plan','nonpositive_plan'])
        self.assertTrue(all(row['attainment'] is None for row in result['rows'][:3]))
        self.assertIsNone(result['continuity'][0]['trailing_behind_periods'])

    def test_actual_zero_and_unverified_cache(self):
        value=self.source['columns'].index(column_id('metric','balance','actual'))
        self.source['rows'][0]['cells'][value]['value']=Decimal(0)
        result=analyze(self.source,self.config)
        self.assertEqual(result['rows'][0]['observation_status'],'observed_zero')
        self.assertEqual(result['rows'][0]['attainment'],0)
        self.source['unverified_formula_caches']=1
        result=analyze(self.source,self.config)
        self.assertTrue(all(row['comparison_status']=='unverified_or_quality_actual' for row in result['rows']))
        self.assertTrue(all(row['peer_rank'] is None for row in result['rows']))

    def test_incompatible_types_versions_levels_and_periods_fail(self):
        for mutate in [lambda value:value['institutions'][0].update(level='region'),lambda value:value['plan_definition'].update(target_type='ytd'),lambda value:value['plan_definition'].update(unit='元'),lambda value:value['plan_definition'].update(version='v2'),lambda value:value.update(periods=['2025-01','2025-03']),lambda value:value.update(context={})]:
            config=deepcopy(self.config); mutate(config)
            with self.assertRaises(ValueError): analyze(self.source,config)
        source=deepcopy(self.source); source['definition_version']='v2'
        with self.assertRaisesRegex(ValueError,'definition version'): analyze(source,self.config)
        index=source['columns'].index(column_id('dimension','version'))
        source=deepcopy(self.source); source['rows'][0]['cells'][index]['value']='v2'
        with self.assertRaisesRegex(ValueError,'incompatible'): analyze(source,self.config)

    def test_at_most_zero_actual_does_not_invent_infinite_rank(self):
        self.config['plan_definition']['direction']='at_most'
        value=self.source['columns'].index(column_id('metric','balance','actual'))
        self.source['rows'][0]['cells'][value]['value']=Decimal(0)
        row=analyze(self.source,self.config)['rows'][0]
        self.assertFalse(row['behind_plan']); self.assertIsNone(row['attainment']); self.assertIsNone(row['peer_rank'])
        self.assertIn('undefined_attainment_nonpositive_actual',row['quality_flags'])

    def daily_case(self,missing=False):
        days=[date(2024,2,1)+timedelta(days=index) for index in range(29)]
        rows=[[FIRST,'balance',day.isoformat(),100,'万元','CNY','v1','daily_snapshot'] for day in days if not missing or day.day!=29]
        path=self.folder/('daily-missing.xlsx' if missing else 'daily.xlsx')
        Workbook.create(daily_spec(rows)).save(path)
        query=daily_config(); query.update(start='2024-02-01',end='2024-02-29')
        source=summarize(path,'DailyFacts',query)
        config=policy_config(); config['context']={}; config['periods']=['2024-02']; config['institutions']=config['institutions'][:1]
        config['scope']['time_basis']='daily_average'; config['plan_definition'].update(time_basis='daily_average',target_type='daily_average')
        config['plans']=[{'institution':FIRST,'period':'2024-02','applicability':'applicable','value':120}]
        config['daily_entity_role']='institution'
        return source,config

    def test_daily_mean_has_own_target_type_and_complete_calendar(self):
        source,config=self.daily_case()
        result=analyze(source,config)
        self.assertEqual(result['rows'][0]['actual'],100)
        self.assertEqual(result['rows'][0]['gap'],20)
        self.assertEqual(result['source_row_kind'],'daily_summary_row')
        export_policy(result,self.folder/'daily-policy.xlsx')
        config['plan_definition']['target_type']='balance'
        with self.assertRaisesRegex(ValueError,'target type'): analyze(source,config)

    def test_daily_missing_day_or_undeclared_institution_does_not_rank(self):
        source,config=self.daily_case(missing=True)
        row=analyze(source,config)['rows'][0]
        self.assertIsNone(row['actual']); self.assertIsNone(row['peer_rank'])
        self.assertEqual(row['comparison_status'],'unverified_or_quality_actual')
        del config['daily_entity_role']
        with self.assertRaisesRegex(ValueError,'institution entity role'): analyze(source,config)
        config['daily_entity_role']='institution'; source['requested_query']['start']='2024-02-02'
        with self.assertRaisesRegex(ValueError,'natural-month'): analyze(source,config)

    def test_contradictory_daily_coverage_is_not_a_clean_actual(self):
        source,config=self.daily_case()
        for metric,value in [('observed_days',28),('reported_days',28),('expected_days',30)]:
            report=deepcopy(source)
            index=report['columns'].index(column_id('metric',metric,'actual'))
            report['rows'][0]['cells'][index]['value']=value
            with self.assertRaisesRegex(ValueError,'coverage|Coverage'):
                analyze(report,config)
        report=deepcopy(source); report['source_metric_dictionary']=[None]
        with self.assertRaisesRegex(ValueError,'dictionary'): analyze(report,config)

    def test_cli_safe_export_dictionary_lineage_and_no_overwrite(self):
        source=self.folder/'wide.json'; source.write_text(json.dumps(self.source,default=str))
        query=self.folder/'policy.json'; query.write_text(json.dumps(self.config)); output=self.folder/'policy.xlsx'
        command=[sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_management_policy.py')),str(source),'--query',str(query),'--xlsx',str(output)]
        run=subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stdout+run.stderr)
        book=Workbook.load(output); self.assertEqual(len(book.sheets),8)
        self.assertEqual(book.cells['Progress']['A2']['value'],FIRST)
        self.assertTrue(all(not cell.get('formula') for cells in book.cells.values() for cell in cells.values()))
        original=output.read_bytes(); run=subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(run.returncode,1); self.assertEqual(output.read_bytes(),original)


if __name__=='__main__': unittest.main()
