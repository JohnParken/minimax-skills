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
from test_xlsx_hierarchy import config_with_tree,two_branches
from test_xlsx_management import spec_for
from xlsx_management import build,column_id
from xlsx_reconcile import reconcile,export_reconciliation


def query():
    return {'left_period':'2025-01','right_period':'2025-02',
            'keys':['institution','line','currency','version','report_level'],
            'metrics':['balance'],'group_by':['institution'],
            'scope':{'currency':'CNY','version':'v1','report_level':'branch'},'tolerance':Decimal('0.01')}


class ReconciliationTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.folder=Path(self.temp.name)
        self.left=self.make('left.xlsx',two_branches())
        facts=two_branches()
        for row in facts:
            row[0]='2025-02'
            if row[6]=='balance' and row[5]=='P1': row[7]+=10
        self.right=self.make('right.xlsx',facts)

    def make(self,name,facts,config=None):
        path=self.folder/name; Workbook.create(spec_for(facts)).save(path)
        if config is None:
            config=config_with_tree()
            config['metrics']=[metric for metric in config['metrics'] if metric['metric_id']=='balance']
            path.unlink(); Workbook.create(spec_for([row for row in facts if row[6]=='balance'])).save(path)
        return build(path,'Facts',config)

    def test_changed_records_totals_and_contribution(self):
        result=reconcile(self.left,self.right,query())
        self.assertEqual(len(result['records']),2)
        self.assertTrue(all(row['status']=='changed' for row in result['records']))
        self.assertEqual(result['checks'][0]['net_change'],Decimal(20))
        self.assertEqual(result['checks'][0]['residual'],0)
        self.assertEqual([row['signed_share'] for row in result['contributions']],[Decimal('.5'),Decimal('.5')])
        self.assertEqual(result['records'][0]['left_wide_row'],1)

    def test_added_removed_and_null_not_imputed_observations(self):
        facts=two_branches()
        for row in facts:
            row[0]='2025-02'
            if row[1]=='0002': row[1]='0003'
        config=config_with_tree(); config['hierarchy']['nodes'].append({'id':'0003','parent':'R','level':'branch','name':'虚构丙支行'})
        config['hierarchy']['coverage']='observed'
        # Compare leaf rows with the same declared mapping; ancestor partial flags are excluded by scope.
        self.left=self.make('a.xlsx',two_branches(),config)
        self.right=self.make('b.xlsx',facts,config)
        result=reconcile(self.left,self.right,query())
        states={row['status']:row for row in result['records']}
        self.assertIsNone(states['added']['before']); self.assertIsNone(states['removed']['after'])
        self.assertEqual(result['checks'][0]['net_change'],0)
        self.assertTrue(all(row['signed_share'] is None for row in result['contributions']))

    def test_tolerance_retains_delta_and_negative_net(self):
        index=self.right['columns'].index(column_id('metric','balance','actual'))
        self.right['rows'][0]['cells'][index]['value']=Decimal('399.995')
        self.right['rows'][-1]['cells'][index]['value']=Decimal('190')
        result=reconcile(self.left,self.right,query())
        self.assertEqual(result['records'][0]['status'],'unchanged_within_tolerance')
        self.assertEqual(result['checks'][0]['net_change'],Decimal('-10.005'))
        self.assertEqual(result['checks'][0]['net_status'],'negative_net_change')
        self.assertTrue(all(row['share_status']=='signed_negative_net' for row in result['contributions']))

    def test_precision_exact_large_amounts(self):
        index=self.left['columns'].index(column_id('metric','balance','actual'))
        for report in (self.left,self.right): report['rows'][0]['cells'][index]['value']=Decimal('123456789012345678901234567890.123')
        self.right['rows'][0]['cells'][index]['value']=Decimal('123456789012345678901234567890.124')
        result=reconcile(self.left,self.right,query())
        self.assertEqual(result['records'][0]['delta'],Decimal('.001'))

    def test_reject_incompatible_schema_and_duplicates(self):
        for mutate in [lambda report:report.update(definition_version='v2'),lambda report:report['hierarchy_summary'].update(mapping_version='changed'),lambda report:report['metric_dictionary'][0].update(unit='元'),lambda report:report['rows'].append(deepcopy(report['rows'][0]))]:
            report=deepcopy(self.right); mutate(report)
            with self.assertRaises(ValueError): reconcile(self.left,report,query())
        config=query(); config['metrics']=['risk_rate']; config['right_period']='2025-01'
        report=self.make('ratio.xlsx',two_branches(),config_with_tree())
        with self.assertRaisesRegex(ValueError,'additive SUM'): reconcile(report,report,config)

    def test_reject_null_quality_and_unsafe_keys(self):
        index=self.right['columns'].index(column_id('metric','balance','actual'))
        report=deepcopy(self.right); report['rows'][0]['cells'][index]={'type':'blank','value':None}
        with self.assertRaisesRegex(ValueError,'Missing'): reconcile(self.left,report,query())
        report=deepcopy(self.right); report['unverified_formula_caches']=1
        with self.assertRaisesRegex(ValueError,'quality'): reconcile(self.left,report,query())
        for mutate in [lambda config:config['keys'].remove('currency'),lambda config:config['scope'].pop('version'),lambda config:config.update(tolerance=-1),lambda config:config.update(right_period='2025-03')]:
            config=query(); mutate(config)
            with self.assertRaises(ValueError): reconcile(self.left,self.right,config)

    def test_nonadditive_totals_and_incompatible_scope_fail(self):
        reports=[deepcopy(self.left),deepcopy(self.right)]
        for report in reports: report['metric_dictionary'][0]['additive_dimensions']=[]
        with self.assertRaisesRegex(ValueError,'not declared additive'):
            reconcile(*reports,query())
        config=query(); config['scope']['currency']='USD'
        with self.assertRaisesRegex(ValueError,'No observations'):
            reconcile(self.left,self.right,config)

    def test_cli_safe_export_and_metadata_lineage(self):
        left=self.folder/'left.json'; right=self.folder/'right.json'; config=self.folder/'query.json'
        left.write_text(json.dumps(self.left,default=str)); right.write_text(json.dumps(self.right,default=str))
        config.write_text(json.dumps(query(),default=float)); output=self.folder/'report.xlsx'
        command=[sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_reconcile.py')),str(left),str(right),'--query',str(config),'--xlsx',str(output)]
        run=subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stdout+run.stderr)
        result=json.loads(run.stdout); self.assertEqual(result['status'],'reconciled')
        book=Workbook.load(output); self.assertEqual(len(book.sheets),5)
        self.assertTrue(any(cell['value']=='00012345678901234567890' for cell in book.cells['Records'].values()))
        self.assertTrue(all(not cell.get('formula') for cells in book.cells.values() for cell in cells.values()))
        self.assertIn('lineage', ''.join(str(cell['value']) for cell in book.cells['Metadata'].values()))
        original=output.read_bytes(); run=subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(run.returncode,1); self.assertEqual(output.read_bytes(),original)


if __name__=='__main__': unittest.main()
