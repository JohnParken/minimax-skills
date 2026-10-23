#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
from copy import deepcopy
from datetime import date,timedelta
from decimal import Decimal
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook
from xlsx_dashboard import build as dashboard,export_dashboard
from xlsx_management_policy import analyze as policy,export_policy
from xlsx_maturity import analyze as maturity
from xlsx_customers import analyze as customers
from xlsx_deposits import analyze as deposits,export_deposits
from xlsx_daily import summarize
from xlsx_result_export import export_result
from test_xlsx_maturity import config_example as maturity_config,fictional_rows,spec_for as maturity_spec
from test_xlsx_customers import config_example as customer_config,records,spec_for as customer_spec
from test_xlsx_deposits import config_example as deposit_config,rows_example,snapshot_spec
from test_xlsx_daily import daily_config,daily_spec


def policy_config(metric='eligible_principal'):
    return {'metric_id':metric,'scope':{'report_level':'branch','currency':'CNY','version':'v1','definition_version':'v1','unit':'元','kind':'stock','time_basis':'point_in_time','period_granularity':'month'},
            'periods':['2024-02'],'institutions':[{'id':'A','level':'branch','peer_group':'group-1'},{'id':'B','level':'branch','peer_group':'group-1'}],
            'context':{'observation_date':'2024-02-28'},'plan_definition':{'target_type':'balance','unit':'元','kind':'stock','time_basis':'point_in_time','currency':'CNY','version':'v1','definition_version':'v1','plan_version':'explicit-plan','direction':'at_least'},
            'plans':[{'institution':institution,'period':'2024-02','applicability':'applicable','value':100} for institution in ('A','B')],'rank_method':'competition'}


class ManagementIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.folder=Path(self.temp.name)

    def save(self,name,spec):
        path=self.folder/name; Workbook.create(spec).save(path); return path

    def test_maturity_wide_dashboard_policy_and_readback(self):
        source=self.save('maturity.xlsx',maturity_spec(fictional_rows()))
        wide=maturity(source,maturity_config()); export_result(wide,self.folder/'wide.xlsx')
        dashboard_result=dashboard(wide,{'overview_period':'2024-02','metrics':['eligible_principal'],'ranks':{'eligible_principal':{'method':'competition','descending':True}}})
        export_dashboard(dashboard_result,self.folder/'dashboard.xlsx')
        progress=policy(wide,policy_config()); export_policy(progress,self.folder/'policy.xlsx')
        self.assertEqual(progress['rows'][0]['actual'],90)
        self.assertEqual(progress['rows'][0]['gap'],10)
        self.assertEqual(progress['rows'][1]['comparison_status'],'unverified_or_quality_actual')
        for name in ('wide.xlsx','dashboard.xlsx','policy.xlsx'):
            book=Workbook.load(self.folder/name)
            self.assertTrue(book.cells)
            self.assertFalse(any(cell.get('formula') for sheet in book.cells.values() for cell in sheet.values()))
        with self.assertRaisesRegex(ValueError,'management_wide_v1'): dashboard(progress,{'overview_period':'2024-02','metrics':['eligible_principal']})

    def test_customer_wide_preserves_peer_group_for_dashboard(self):
        left,right=records()
        result=customers(self.save('customers-left.xlsx',customer_spec(left)),self.save('customers-right.xlsx',customer_spec(right)),customer_config())
        wide=result['management_wide']; export_result(wide,self.folder/'customers-wide.xlsx')
        report=dashboard(wide,{'overview_period':'2025-02','metrics':['observed_customers'],'ranks':{'observed_customers':{'method':'competition','descending':True}}})
        institution=[row for row in report['overview'] if row['dimensions']['report_level']=='institution']
        self.assertEqual(len(institution),2)
        self.assertTrue(all(row['peer_count']==1 for row in institution),'Customer wide loses configured peer groups and pools unlike institutions')
        export_dashboard(report,self.folder/'customers-dashboard.xlsx')
        self.assertIn('Overview',Workbook.load(self.folder/'customers-dashboard.xlsx').cells)

    def test_deposit_export_is_supported_but_raw_dashboard_schema_is_refused(self):
        left,right=rows_example()
        result=deposits(self.save('deposit-left.xlsx',snapshot_spec(left)),self.save('deposit-right.xlsx',snapshot_spec(right)),deposit_config())
        export_deposits(result,self.folder/'deposits.xlsx')
        self.assertEqual(result['totals']['residual'],0)
        self.assertIn('Checks',Workbook.load(self.folder/'deposits.xlsx').cells)
        with self.assertRaisesRegex(ValueError,'management_wide_v1'): dashboard(result,{'overview_period':'2024-03','metrics':['balance']})

    def test_complete_daily_policy_adapter_and_native_readback(self):
        days=[date(2024,2,1)+timedelta(days=index) for index in range(29)]
        rows=[['A','balance',day.isoformat(),10,'元','CNY','v1','daily_snapshot'] for day in days]
        source=self.save('daily.xlsx',daily_spec(rows)); query=daily_config()
        query.update(start='2024-02-01',end='2024-02-29'); query['metrics'][0]['unit']='元'
        daily=summarize(source,'DailyFacts',query)
        config=policy_config('balance'); config['scope']['time_basis']='daily_average'
        config['plan_definition'].update(time_basis='daily_average',target_type='daily_average')
        config.update(context={},daily_entity_role='institution'); config['institutions']=config['institutions'][:1]; config['plans']=config['plans'][:1]
        result=policy(daily,config); self.assertEqual(result['rows'][0]['actual'],10)
        export_policy(result,self.folder/'daily-policy.xlsx')
        self.assertIn('Dictionary',Workbook.load(self.folder/'daily-policy.xlsx').cells)


if __name__=='__main__': unittest.main()
