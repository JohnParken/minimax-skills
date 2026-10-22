#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
from copy import deepcopy
from decimal import Decimal
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook
from xlsx_tabular import column_name
from xlsx_daily import summarize
from xlsx_management import column_id
from xlsx_result_export import export_result


def daily_config(operation='average'):
    return {'entity_keys':['A'],'columns':dict(zip(('metric','date','value','unit','currency','version','basis'),'BCDEFGH')),
            'metrics':[{'metric_id':'balance' if operation=='average' else 'income','label':'虚构每日余额' if operation=='average' else '虚构每日发生额',
                        'unit':'万元','kind':'stock' if operation=='average' else 'flow','time_basis':'daily_snapshot' if operation=='average' else 'daily_flow'}],
            'date_format':'YYYY-MM-DD','start':'2024-02-28' if operation=='average' else '2024-01-01',
            'end':'2024-03-01' if operation=='average' else '2024-01-03','definition_version':'v1','coverage':'complete'}


def daily_spec(records):
    rows=[['实体','指标','日期','值','单位','币种','口径','时间口径']]+records
    return {'sheets':[{'name':'DailyFacts','cells':{column_name(col)+str(row):{'value':value}
                                                 for row,items in enumerate(rows,1) for col,value in enumerate(items,1)}}]}


def daily_records(operation='average'):
    days=('2024-02-28','2024-02-29','2024-03-01') if operation=='average' else ('2024-01-01','2024-01-02','2024-01-03')
    return [['00012345678901234567890','balance' if operation=='average' else 'income',day,value,'万元','CNY','v1',
             'daily_snapshot' if operation=='average' else 'daily_flow']
            for day,value in zip(days,(90,120,150) if operation=='average' else (10,20,30))]


class DailyTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.source=Path(self.temp.name)/'daily.xlsx'

    def save(self,records): Workbook.create(daily_spec(records)).save(self.source)
    def value(self,result,metric,row=0):
        return result['rows'][row]['cells'][result['columns'].index(column_id('metric',metric,'actual'))]['value']

    def test_leap_day_cross_month_mean_and_xlsx_readback(self):
        self.save(daily_records()); before=self.source.read_bytes()
        result=summarize(self.source,'DailyFacts',daily_config())
        self.assertEqual(self.value(result,'daily_average'),120)
        self.assertEqual(self.value(result,'expected_days'),3)
        self.assertEqual(result['coverage_details'][0]['missing_dates'],[])
        output=self.source.with_name('average.xlsx'); export_result(result,output)
        book=Workbook.load(output)
        self.assertEqual(book.cells['Daily']['A2']['value'],'00012345678901234567890')
        self.assertIn('Lineage',book.sheets)
        self.assertEqual(self.source.read_bytes(),before)

    def test_missing_date_full_mean_null_and_explicit_observed_mean(self):
        records=daily_records(); self.save([records[0],records[2]])
        result=summarize(self.source,'DailyFacts',daily_config())
        self.assertIsNone(self.value(result,'daily_average'))
        self.assertEqual(self.value(result,'observed_daily_average'),120)
        self.assertEqual(result['coverage_details'][0]['missing_dates'],['2024-02-29'])
        config=daily_config(); config['coverage']='observed'
        result=summarize(self.source,'DailyFacts',config)
        self.assertEqual(self.value(result,'daily_average'),120)
        self.assertIn('observed_only_not_full_interval',result['rows'][0]['quality_flags'])

    def test_blank_values_are_missing_days_and_zeros_are_observations(self):
        records=daily_records(); records[1][3]=None; records[0][3]=0; self.save(records)
        result=summarize(self.source,'DailyFacts',daily_config())
        self.assertEqual(self.value(result,'reported_days'),3)
        self.assertEqual(self.value(result,'observed_days'),2)
        self.assertEqual(self.value(result,'observed_daily_average'),75)
        self.assertEqual(result['coverage_details'][0]['missing_dates'],['2024-02-29'])

    def test_duplicate_invalid_date_non_numeric_and_basis_errors(self):
        for operation in ('duplicate','invalid_date','text_value','ytd_basis'):
            records=daily_records()
            if operation=='duplicate': records.append(deepcopy(records[0]))
            if operation=='invalid_date': records[0][2]='2023-02-29'
            if operation=='text_value': records[0][3]='90'
            if operation=='ytd_basis': records[0][7]='ytd'
            self.save(records)
            with self.subTest(operation=operation),self.assertRaises(ValueError):
                summarize(self.source,'DailyFacts',daily_config())
            self.source.unlink()

    def test_currencies_are_separate_and_versions_or_units_mismatch_fail(self):
        records=daily_records(); usd=deepcopy(records)
        for row in usd: row[5]='USD'; row[3]*=2
        self.save(records+usd)
        result=summarize(self.source,'DailyFacts',daily_config())
        self.assertEqual(result['output_rows'],2)
        self.assertEqual([self.value(result,'daily_average',index) for index in range(2)],[120,240])
        self.source.unlink(); records=daily_records(); records[0][6]='v2'; self.save(records)
        with self.assertRaises(ValueError): summarize(self.source,'DailyFacts',daily_config())

    def test_daily_flow_month_and_ytd_do_not_sum_ytd_snapshots(self):
        self.save(daily_records('accumulate')); config=daily_config('accumulate')
        result=summarize(self.source,'DailyFacts',config,'accumulate')
        self.assertEqual(result['output_rows'],2)
        self.assertEqual([self.value(result,'period_flow_sum',index) for index in range(2)],[60,60])
        self.assertEqual(len(result['lineage']),6)
        config['metrics'][0]['time_basis']='ytd'
        with self.assertRaisesRegex(ValueError,'YTD snapshots'):
            summarize(self.source,'DailyFacts',config,'accumulate')

    def test_partial_month_windows_and_no_partial_year_relabelled_ytd(self):
        records=daily_records('accumulate')[:2]; records[0][2]='2024-01-31'; records[1][2]='2024-02-01'
        self.save(records); config=daily_config('accumulate'); config.update(start='2024-01-31',end='2024-02-01',periods=['month'])
        result=summarize(self.source,'DailyFacts',config,'accumulate')
        self.assertEqual([self.value(result,'period_flow_sum',index) for index in range(2)],[10,20])
        config['periods']=['ytd']
        with self.assertRaisesRegex(ValueError,'January 1'):
            summarize(self.source,'DailyFacts',config,'accumulate')

    def test_missing_flow_dates_never_assumed_zero(self):
        records=daily_records('accumulate'); self.save([records[0],records[2]])
        result=summarize(self.source,'DailyFacts',daily_config('accumulate'),'accumulate')
        self.assertTrue(all(self.value(result,'period_flow_sum',index) is None for index in range(2)))
        self.assertEqual(self.value(result,'observed_flow_sum'),40)

    def test_explicit_outside_policy_and_interval_limits(self):
        records=daily_records(); records.append([records[0][0],'balance','2024-03-02',999,'万元','CNY','v1','daily_snapshot']); self.save(records)
        with self.assertRaisesRegex(ValueError,'outside'):
            summarize(self.source,'DailyFacts',daily_config())
        config=daily_config(); config['outside']='exclude'
        result=summarize(self.source,'DailyFacts',config)
        self.assertEqual(result['outside_rows'],1)
        self.assertEqual(self.value(result,'daily_average'),120)
        config['end']='2026-01-01'
        with self.assertRaises(ValueError): summarize(self.source,'DailyFacts',config)

    def test_full_leap_year_expected_days_and_month_year_flow_totals(self):
        from datetime import date,timedelta
        records=[]
        start=date(2024,1,1)
        for offset in range(366):
            day=start+timedelta(days=offset)
            records.append(['0001','income',day.isoformat(),1,'万元','CNY','v1','daily_flow'])
        self.save(records)
        config=daily_config('accumulate'); config.update(start='2024-01-01',end='2024-12-31')
        result=summarize(self.source,'DailyFacts',config,'accumulate')
        self.assertEqual(result['output_rows'],13)
        self.assertEqual(self.value(result,'period_flow_sum',12),366)
        self.assertEqual(result['coverage_details'][1]['expected_days'],29)
        self.assertTrue(all(not item['missing_dates'] for item in result['coverage_details']))

    def test_cli_average_and_accumulation_xlsx(self):
        for operation in ('average','accumulate'):
            self.save(daily_records(operation)); query=self.source.with_suffix('.json'); query.write_text(json.dumps(daily_config(operation)))
            output=self.source.with_name(operation+'.xlsx')
            run=subprocess.run([sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_daily.py')),operation,str(self.source),
                                '--sheet','DailyFacts','--query',str(query),'--xlsx',str(output)],capture_output=True,text=True)
            self.assertEqual(run.returncode,0,run.stderr+run.stdout)
            self.assertTrue(json.loads(run.stdout)['export']['output_written'])
            self.assertEqual(len(Workbook.load(output).sheets),7)
            self.source.unlink()


if __name__=='__main__': unittest.main()
