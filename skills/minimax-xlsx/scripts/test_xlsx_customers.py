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
from xlsx_customers import analyze,export_customers
from xlsx_result_export import export_result


def config_example():
    side={'sheet':'Observed','columns':dict(zip(('account','customer','institution','product','holding','period','version','level','scope','owner'),'ABCDEFGHIJ'))}
    return {'left':deepcopy(side),'right':deepcopy(side),'periods':{'left':'2025-01','right':'2025-02'},
            'scope':{'level':'branch','version':'v1','definition_version':'definition-v1','observation_scope':'fictional-selected-universe','completeness':'declared_complete'},
            'institutions':[{'id':'A','level':'branch','peer_group':'small'},{'id':'B','level':'branch','peer_group':'large'}],
            'ownership':'designated','categories':[{'id':'deposit','label':'存款产品'},{'id':'payments','label':'支付产品'}],
            'product_mapping':{'P1':['deposit','payments'],'P2':['deposit']},'coverage_denominator':'observed_customers','fact_grain':'account_product'}


def records():
    def row(account,customer,institution,product,holding,period):
        return [account,customer,institution,product,holding,period,'v1','branch','fictional-selected-universe',institution]
    return ([row('account-001','customer-00012345678901234567890','A','P1',True,'2025-01'),
             row('account-002','customer-00012345678901234567890','A','P2',True,'2025-01'),
             row('account-003','customer-002','A','P2',True,'2025-01'),
             row('account-004','customer-003','B','P1',False,'2025-01')],
            [row('account-001','customer-00012345678901234567890','B','P1',True,'2025-02'),
             row('account-002','customer-00012345678901234567890','B','P2',True,'2025-02'),
             row('account-003','customer-002','A','P2',False,'2025-02'),
             row('account-005','customer-004','B','P2',True,'2025-02')])


def spec_for(rows):
    values=[['脱敏账户','稳定脱敏客户','观察机构','产品','明确持有','期间','版本','层级','范围','指定归属']]+rows
    return {'sheets':[{'name':'Observed','cells':{column_name(col)+str(row):{'value':value} for row,items in enumerate(values,1) for col,value in enumerate(items,1)}}]}


class CustomerTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.folder=Path(self.temp.name); self.left=self.folder/'left.xlsx'; self.right=self.folder/'right.xlsx'
        self.before,self.after=records(); self.config=config_example(); self.save()

    def save(self):
        for path,rows in ((self.left,self.before),(self.right,self.after)):
            if path.exists(): path.unlink()
            Workbook.create(spec_for(rows)).save(path)

    def result(self): return analyze(self.left,self.right,self.config)

    def test_account_dedup_overlap_coverage_and_distribution(self):
        result=self.result()
        bank=next(row for row in result['summary'] if row['scope_kind']=='bank' and row['period']=='2025-01')
        self.assertEqual((bank['customer_count'],bank['holding_customer_count']),(3,2))
        coverage={row['category_id']:row for row in result['coverage'] if row['scope_kind']=='bank' and row['period']=='2025-01'}
        self.assertEqual(coverage['deposit']['customer_count'],2); self.assertEqual(coverage['payments']['customer_count'],1)
        distribution={row['held_category_count']:row['customer_count'] for row in result['product_count_distribution'] if row['scope_kind']=='bank' and row['period']=='2025-01'}
        self.assertEqual(distribution,{0:1,1:1,2:1})

    def test_migration_not_bank_entry_exit_and_counts_reconcile(self):
        result=self.result(); bank=result['changes'][0]
        self.assertEqual((bank['continuing_observed'],bank['only_current_observed'],bank['only_previous_observed'],bank['net_observed_change'],bank['residual']),(2,1,1,0,0))
        self.assertEqual(result['migration_count'],1)
        for period in ('2025-01','2025-02'):
            rows=[row for row in result['summary'] if row['period']==period]
            self.assertEqual(sum(row['customer_count'] for row in rows if row['scope_kind']=='institution'),3)
            self.assertEqual(sum(row['customer_count'] for row in rows if row['scope_kind']=='peer_group'),3)

    def test_multi_institution_designated_owner_and_exclusive_rejection(self):
        extra=deepcopy(self.after[0]); extra[0]='account-006'; extra[2]='A'; self.after.append(extra); self.save()
        self.assertEqual(self.result()['changes'][0]['current_customers'],3)
        self.config['ownership']='exclusive'
        for side in ('left','right'): self.config[side]['columns'].pop('owner')
        with self.assertRaisesRegex(ValueError,'multiple institution owners'): self.result()

    def test_ambiguous_designated_owner_and_unknown_institution_fail(self):
        original=deepcopy(self.after)
        for mutate in [lambda rows:rows[1].__setitem__(9,'A'),lambda rows:rows[0].__setitem__(9,'unknown'),lambda rows:rows[0].__setitem__(2,'unknown')]:
            self.after=deepcopy(original); mutate(self.after); self.save()
            with self.assertRaises(ValueError): self.result()

    def test_missing_identity_duplicate_grain_mixed_scope_unmapped_product(self):
        original=deepcopy(self.after)
        for mutate in [lambda rows:rows[0].__setitem__(1,None),lambda rows:rows[0].__setitem__(1,123),lambda rows:rows.append(deepcopy(rows[0])),lambda rows:rows[0].__setitem__(3,'unknown'),lambda rows:rows[0].__setitem__(5,'2025-03'),lambda rows:rows[0].__setitem__(6,'v2'),lambda rows:rows[0].__setitem__(8,'other'),lambda rows:rows[0].__setitem__(4,'true')]:
            self.after=deepcopy(original); mutate(self.after); self.save()
            with self.assertRaises(ValueError): self.result()

    def test_empty_period_zero_denominator_and_completeness_unknown(self):
        self.after=[]; self.save(); self.config['scope']['completeness']='unknown'
        result=self.result()
        bank=next(row for row in result['summary'] if row['scope_kind']=='bank' and row['period']=='2025-02')
        self.assertEqual(bank['customer_count'],0)
        self.assertIn('observation_completeness_unknown',bank['quality_flags'])
        self.assertTrue(all(row['coverage']['value'] is None for row in result['coverage'] if row['period']=='2025-02'))
        self.assertEqual(result['changes'][0]['only_previous_observed'],3)

    def test_holding_denominator_and_no_zero_balance_inference(self):
        self.config['coverage_denominator']='holding_customers'; result=self.result()
        coverage=next(row for row in result['coverage'] if row['scope_kind']=='bank' and row['period']=='2025-01' and row['category_id']=='deposit')
        self.assertEqual(coverage['coverage']['value'],1)
        self.assertEqual(result['changes'][0]['continuing_observed'],2)
        # Customer-002 remains observed despite no explicitly held product in current period.
        current=next(row for row in result['summary'] if row['scope_kind']=='institution' and row['scope_id']=='A' and row['period']=='2025-02')
        self.assertEqual((current['customer_count'],current['holding_customer_count']),(1,0))

    def test_account_customer_identity_change_and_bad_categories_fail(self):
        self.after[0][1]='changed-customer'; self.after[1][1]='changed-customer'; self.save()
        with self.assertRaisesRegex(ValueError,'different customer'): self.result()
        self.before,self.after=records(); self.save()
        self.config['product_mapping']['P1']=['deposit','deposit']
        with self.assertRaisesRegex(ValueError,'unique known'): self.result()

    def test_native_wide_safe_export_and_ratio_semantics(self):
        result=self.result(); wide=result['management_wide']
        self.assertEqual(wide['schema_version'],'management_wide_v1')
        self.assertTrue(all(item['aggregation']=='RATIO' for item in wide['metric_dictionary'] if item['kind']=='ratio'))
        output=self.folder/'wide.xlsx'; export_result(wide,output)
        book=Workbook.load(output)
        self.assertIn('Definitions',book.cells); self.assertIn('Lineage',book.cells)
        self.assertNotIn('customer-00012345678901234567890',json.dumps(wide,default=str))

    def test_cli_aggregate_privacy_readback_and_source_preservation(self):
        config=self.folder/'query.json'; config.write_text(json.dumps(self.config)); output=self.folder/'aggregate.xlsx'
        before=self.left.read_bytes(); after=self.right.read_bytes()
        command=[sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_customers.py')),str(self.left),str(self.right),'--query',str(config),'--xlsx',str(output)]
        run=subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stdout+run.stderr)
        for identifier in ('customer-00012345678901234567890','customer-002','account-001'):
            self.assertNotIn(identifier,run.stdout)
        book=Workbook.load(output); self.assertEqual(len(book.sheets),10)
        cells=[cell for sheet in book.cells.values() for cell in sheet.values()]
        self.assertFalse(any(cell.get('formula') for cell in cells))
        text=' '.join(str(cell['value']) for cell in cells)
        self.assertNotIn('customer-00012345678901234567890',text); self.assertNotIn('account-001',text)
        self.assertEqual(self.left.read_bytes(),before); self.assertEqual(self.right.read_bytes(),after)
        original=output.read_bytes(); run=subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(run.returncode,1); self.assertEqual(output.read_bytes(),original)


if __name__=='__main__': unittest.main()
