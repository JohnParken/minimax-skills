#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
from copy import deepcopy
from decimal import Decimal
import json
from pathlib import Path
from unittest.mock import patch
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook
from xlsx_tabular import column_name
from xlsx_deposits import analyze,export_deposits
from test_xlsx_daily import daily_config,daily_records,daily_spec


def config_example():
    side={'sheet':'Snapshots','columns':dict(zip(('entity','institution','product','balance','currency','unit','version','date','level'),'ABCDEFGHI'))}
    return {'grain':'account','left':deepcopy(side),'right':deepcopy(side),
            'left_date':'2024-02-28','right_date':'2024-03-01',
            'scope':{'currency':'CNY','unit':'万元','version':'v1','level':'branch'}}


def rows_example():
    return ([['00012345678901234567890','A','活期',100,'CNY','万元','v1','2024-02-28','branch'],
             ['0002','A','定期',200,'CNY','万元','v1','2024-02-28','branch'],
             ['0003','B','活期',40,'CNY','万元','v1','2024-02-28','branch']],
            [['00012345678901234567890','B','定期',150,'CNY','万元','v1','2024-03-01','branch'],
             ['0002','A','定期',180,'CNY','万元','v1','2024-03-01','branch'],
             ['0004','B','活期',70,'CNY','万元','v1','2024-03-01','branch']])


def snapshot_spec(rows):
    values=[['脱敏账户','机构','产品','余额','币种','单位','口径','日期','层级']]+rows
    return {'sheets':[{'name':'Snapshots','cells':{column_name(col)+str(row):{'value':value} for row,items in enumerate(values,1) for col,value in enumerate(items,1)}}]}


class DepositTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.folder=Path(self.temp.name); self.left=self.folder/'left.xlsx'; self.right=self.folder/'right.xlsx'
        self.before,self.after=rows_example(); self.config=config_example()
        self.save()

    def save(self):
        for path,rows in ((self.left,self.before),(self.right,self.after)):
            if path.exists(): path.unlink()
            Workbook.create(snapshot_spec(rows)).save(path)

    def run_analysis(self): return analyze(self.left,self.right,self.config)

    def test_components_reconcile_and_transfer_not_entry_exit(self):
        result=self.run_analysis(); totals=result['totals']
        self.assertEqual((totals['before'],totals['balance'],totals['net_change'],totals['residual']),(340,400,60,0))
        self.assertEqual(totals['components'],{'continuing_increase':50,'continuing_decrease':-20,'only_current_observed':70,'only_previous_observed':-40})
        moved=result['records'][0]; self.assertEqual(moved['component'],'continuing_increase')
        self.assertTrue(moved['institution_changed'] and moved['product_changed'])
        self.assertIsNone(result['records'][2]['balance']); self.assertIsNone(result['records'][3]['before'])
        self.assertIn('observation_absence_not_lifecycle_evidence',result['records'][3]['quality_flags'])

    def add_daily(self,records=None):
        path=self.folder/'daily.xlsx'; Workbook.create(daily_spec(daily_records() if records is None else records)).save(path)
        self.config['daily']={'sheet':'DailyFacts','query':daily_config(),'metric_id':'balance','grain':'account'}
        return path

    def test_complete_daily_mean_and_end_difference_with_leap_day(self):
        daily=self.add_daily(); result=analyze(self.left,self.right,self.config,daily)
        row=result['records'][0]
        self.assertEqual((row['daily_average'],row['end_minus_daily_average']),(120,30))
        self.assertIn('entity_daily_mean_not_attributed_to_institution_or_product',row['quality_flags'])
        self.assertIn('daily_entity_not_observed',result['records'][1]['quality_flags'])
        self.assertTrue(any(item['side']=='daily' for item in result['lineage']))

    def test_incomplete_daily_not_called_full_average(self):
        records=daily_records(); daily=self.add_daily([records[0],records[2]])
        row=analyze(self.left,self.right,self.config,daily)['records'][0]
        self.assertIsNone(row['daily_average']); self.assertIsNone(row['end_minus_daily_average'])
        self.assertEqual(row['observed_daily_average'],120)
        self.assertIn('incomplete_calendar_coverage',row['quality_flags'])
        self.config['daily']['query']['coverage']='observed'
        with self.assertRaisesRegex(ValueError,'complete coverage'): analyze(self.left,self.right,self.config,daily)

    def test_targets_are_separate_by_measure_and_period(self):
        daily=self.add_daily(); base={'entity':'00012345678901234567890',**self.config['scope'],'grain':'account'}
        self.config['targets']=[{**base,'metric':metric,'period':period,'value':value} for metric,period,value in
            [('balance','2024-03-01',200),('net_change','2024-02-28..2024-03-01',100),('daily_average','2024-02-28..2024-03-01',150)]]
        result=analyze(self.left,self.right,self.config,daily); targets=result['records'][0]['targets']
        self.assertEqual([targets[metric]['gap'] for metric in ('balance','net_change','daily_average')],[50,50,30])
        self.config['targets'][0]['period']='2024-02-28..2024-03-01'
        with self.assertRaisesRegex(ValueError,'cannot be reused'): analyze(self.left,self.right,self.config,daily)

    def test_zero_negative_target_and_missing_actual(self):
        base={'entity':'0003',**self.config['scope'],'grain':'account','metric':'balance','period':'2024-03-01','value':100}
        self.config['targets']=[base]
        self.assertEqual(self.run_analysis()['records'][2]['targets']['balance']['status'],'missing_actual')
        for target in (0,-1):
            self.config['targets']=[{**base,'entity':'0002','value':target}]
            item=self.run_analysis()['records'][1]['targets']['balance']
            self.assertEqual(item['status'],'nonpositive_target'); self.assertIsNone(item['achievement'])

    def test_duplicates_metadata_dates_and_types_fail(self):
        original=deepcopy(self.after)
        for mutate in [lambda rows:rows.append(deepcopy(rows[0])),lambda rows:rows[0].__setitem__(3,'150'),lambda rows:rows[0].__setitem__(3,-1),lambda rows:rows[0].__setitem__(4,'USD'),lambda rows:rows[0].__setitem__(5,'元'),lambda rows:rows[0].__setitem__(6,'v2'),lambda rows:rows[0].__setitem__(7,'2024-03-02'),lambda rows:rows[0].__setitem__(8,'region')]:
            self.after=deepcopy(original); mutate(self.after); self.save()
            with self.assertRaises(ValueError): self.run_analysis()

    def test_grain_daily_unit_and_interval_fail(self):
        self.config['grain']='unknown'
        with self.assertRaises(ValueError): self.run_analysis()
        self.config=config_example(); daily=self.add_daily()
        self.config['daily']['grain']='customer'
        with self.assertRaises(ValueError): analyze(self.left,self.right,self.config,daily)
        self.config['daily']['grain']='account'; self.config['daily']['query']['end']='2024-03-02'
        with self.assertRaises(ValueError): analyze(self.left,self.right,self.config,daily)

    def test_negative_daily_and_unit_incompatibility_fail(self):
        records=daily_records(); records[0][3]=-1; daily=self.add_daily(records)
        with self.assertRaisesRegex(ValueError,'Negative daily'): analyze(self.left,self.right,self.config,daily)
        self.config['daily']['query']['metrics'][0]['unit']='元'
        with self.assertRaisesRegex(ValueError,'matching unit'): analyze(self.left,self.right,self.config,daily)

    def test_filtered_negative_rows_are_excluded_and_daily_source_parsed_once(self):
        rows=daily_records(); excluded=deepcopy(rows[0]); excluded[0]='unselected'; excluded[3]=-100
        daily=self.add_daily(rows+[excluded])
        self.config['daily']['query']['filters']=[{'column':'A','op':'eq','value':'00012345678901234567890'}]
        from xlsx_analysis import region as real_region
        reads=[]
        def observed_region(source,*args,**kwargs):
            reads.append(str(source)); return real_region(source,*args,**kwargs)
        with patch('xlsx_daily.region',side_effect=observed_region),patch('xlsx_deposits.region',side_effect=observed_region):
            result=analyze(self.left,self.right,self.config,daily)
        self.assertEqual(result['records'][0]['daily_average'],120)
        self.assertEqual(reads.count(str(daily)),1)

    def test_precision_zero_and_customer_grain(self):
        self.before[0][3]=Decimal('123456789012345678901234567890.001')
        self.after[0][3]=Decimal('123456789012345678901234567890.002')
        self.before[1][3]=0; self.after[1][3]=0; self.save(); self.config['grain']='customer'
        result=self.run_analysis()
        self.assertEqual(result['records'][0]['net_change'],Decimal('.001'))
        self.assertEqual(result['records'][1]['component'],'continuing_unchanged')
        output=self.folder/'precise.xlsx'; export_deposits(result,output)
        self.assertTrue(any(cell['value']=='123456789012345678901234567890.002' for cell in Workbook.load(output).cells['Deposits'].values()))

    def test_cli_xlsx_readback_and_source_preservation(self):
        query=self.folder/'query.json'; query.write_text(json.dumps(self.config)); output=self.folder/'deposits.xlsx'
        before=self.left.read_bytes(); after=self.right.read_bytes()
        command=[sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_deposits.py')),str(self.left),str(self.right),'--query',str(query),'--xlsx',str(output)]
        run=subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stdout+run.stderr)
        self.assertEqual(json.loads(run.stdout)['totals']['net_change'],'60')
        book=Workbook.load(output); self.assertEqual(len(book.sheets),6)
        self.assertEqual(book.cells['Deposits']['A2']['value'],'00012345678901234567890')
        self.assertTrue(all(not cell.get('formula') for cells in book.cells.values() for cell in cells.values()))
        self.assertEqual(self.left.read_bytes(),before); self.assertEqual(self.right.read_bytes(),after)
        original=output.read_bytes(); run=subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(run.returncode,1); self.assertEqual(output.read_bytes(),original)


if __name__=='__main__': unittest.main()
