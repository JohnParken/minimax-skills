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
from test_xlsx_deposits import config_example,rows_example,snapshot_spec
from xlsx_deposit_structure import analyze,export_structure


def structure_config():
    config=config_example()
    return {'snapshot':config['right'],'date':config['right_date'],'scope':config['scope'],
            'concentration_grain':'account','top_n':[1,2,10],
            'classifications':[{'id':'tenor','field':'product','mapping':{'活期':'活期','定期':'定期'}}]}


class DepositStructureTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.folder=Path(self.temp.name); self.source=self.folder/'snapshot.xlsx'
        self.rows=rows_example()[1]; self.config=structure_config(); self.save()

    def save(self,customers=None):
        if self.source.exists(): self.source.unlink()
        spec=snapshot_spec(self.rows)
        if customers is not None:
            cells=spec['sheets'][0]['cells']; cells['J1']={'value':'脱敏客户'}
            for index,value in enumerate(customers,2): cells['J'+str(index)]={'value':value}
        Workbook.create(spec).save(self.source)

    def test_explicit_structure_and_account_top_n(self):
        result=analyze(self.source,self.config)
        self.assertEqual(result['total_balance'],400)
        self.assertEqual({row['category']:row['balance'] for row in result['structure']},{'定期':330,'活期':70})
        self.assertEqual(result['concentration'][0]['share']['value'],Decimal('.45'))
        self.assertEqual(result['concentration'][-1]['effective_n'],3)
        self.assertEqual(result['policy']['grain'],'account')

    def test_customer_aggregated_before_top_n(self):
        self.save(['customer-A','customer-A','customer-B'])
        self.config['concentration_grain']='customer'; self.config['customer_column']='J'
        self.config['classifications'].append({'id':'segment','field':'customer','mapping':{'customer-A':'对公','customer-B':'零售'}})
        result=analyze(self.source,self.config)
        self.assertEqual(result['concentration'][0]['balance'],330)
        self.assertEqual(result['concentration'][0]['share']['value'],Decimal('.825'))
        self.assertEqual(result['concentration'][0]['population'],2)
        self.assertEqual(result['ranked_entities'][0]['account_count'],2)
        for axis in ('tenor','segment'):
            self.assertEqual(sum(row['balance'] for row in result['structure'] if row['axis']==axis),400)

    def test_classification_crossing_and_identity_fail(self):
        for mutate in [lambda config:config['classifications'][0]['mapping'].pop('活期'),lambda config:config['classifications'][0]['mapping'].update({'活期':['活期','定期']}),lambda config:config['classifications'].append(deepcopy(config['classifications'][0])),lambda config:config.update(concentration_grain='customer'),lambda config:config.update(customer_column='J'),lambda config:config.update(top_n=[True])]:
            config=deepcopy(self.config); mutate(config)
            with self.assertRaises(ValueError): analyze(self.source,config)
        self.rows.append(deepcopy(self.rows[0])); self.save()
        with self.assertRaisesRegex(ValueError,'Duplicate stable'): analyze(self.source,self.config)

    def test_zero_total_negative_and_tie_order(self):
        for row in self.rows: row[3]=0
        self.save(); result=analyze(self.source,self.config)
        self.assertTrue(all(row['share']['value'] is None for row in result['structure']+result['concentration']))
        self.assertEqual(result['ranked_entities'][0]['entity'],self.rows[0][0])
        self.assertEqual(result['quality_flags'],['zero_total_balance'])
        self.rows[0][3]=-1; self.save()
        with self.assertRaises(ValueError): analyze(self.source,self.config)

    def test_customer_ids_no_guessing(self):
        self.config.update(concentration_grain='customer',customer_column='J')
        for customers in [['a',None,'b'],['a',123,'b']]:
            self.save(customers)
            with self.assertRaisesRegex(ValueError,'Customer IDs'): analyze(self.source,self.config)

    def test_cli_readback_precision_and_source_untouched(self):
        self.rows[0][3]=Decimal('12345678901234567890.01'); self.save()
        query=self.folder/'config.json'; query.write_text(json.dumps(self.config)); output=self.folder/'structure.xlsx'
        original=self.source.read_bytes()
        run=subprocess.run([sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_deposit_structure.py')),str(self.source),'--query',str(query),'--xlsx',str(output)],capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stdout+run.stderr)
        book=Workbook.load(output); self.assertEqual(len(book.sheets),6)
        self.assertTrue(any(cell['value']=='12345678901234567890.01' for cell in book.cells['Entities'].values()))
        self.assertTrue(all(not cell.get('formula') for cells in book.cells.values() for cell in cells.values()))
        self.assertEqual(self.source.read_bytes(),original)
        with self.assertRaisesRegex(ValueError,'already exists'): export_structure(analyze(self.source,self.config),output)


if __name__=='__main__': unittest.main()
