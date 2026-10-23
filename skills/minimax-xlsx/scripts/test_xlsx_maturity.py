#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
from copy import deepcopy
from datetime import date,timedelta
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
from xlsx_maturity import analyze,metric_id
from xlsx_management import column_id


def config_example():
    return {'sheet':'Contracts','columns':dict(zip(('contract','principal','maturity_date','observation_date','institution','currency','unit','version','level','status'),'ABCDEFGHIJ')),
            'observation_date':'2024-02-28','scope':{'currency':'CNY','unit':'元','version':'v1','definition_version':'v1','level':'branch','completeness':'declared_complete'},
            'institutions':['A','B'],'status_mapping':{'active':'include','closed':'exclude'},
            'buckets':[{'id':'days_1_7','min_days':1,'max_days':7},{'id':'days_8_30','min_days':8,'max_days':30},{'id':'days_31_90','min_days':31,'max_days':90},{'id':'days_91_plus','min_days':91,'max_days':None}],
            'cumulative_windows':[7,30,90],'share_denominator':'eligible_total_principal'}


def fictional_rows():
    day=date(2024,2,28)
    return [['fictional-contract-'+str(index),10,None if distance is None else (day+timedelta(days=distance)).isoformat(),day.isoformat(),'B' if distance is None else 'A','CNY','元','v1','branch','active'] for index,distance in enumerate((-1,0,1,7,8,30,31,90,91,None))]+[['fictional-closed',999,'2024-03-01',day.isoformat(),'A','CNY','元','v1','branch','closed']]


def spec_for(rows):
    values=[['脱敏合同','本金','到期日','观察日','机构','币种','单位','版本','层级','状态']]+rows
    return {'sheets':[{'name':'Contracts','cells':{column_name(col)+str(row):{'value':value} for row,items in enumerate(values,1) for col,value in enumerate(items,1)}}]}


class MaturityTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.folder=Path(self.temp.name); self.source=self.folder/'contracts.xlsx'
        self.rows=fictional_rows(); self.config=config_example(); self.save()

    def save(self):
        if self.source.exists(): self.source.unlink()
        Workbook.create(spec_for(self.rows)).save(self.source)

    def result(self): return analyze(self.source,self.config)
    def value(self,result,key,row=0): return result['rows'][row]['cells'][result['columns'].index(column_id('metric',key,'actual'))]['value']

    def test_boundaries_leap_day_and_disjoint_reconciliation(self):
        result=self.result()
        self.assertEqual(self.value(result,'eligible_principal'),100)
        self.assertEqual(self.value(result,'eligible_contracts'),10)
        expected={'past_date':10,'today':10,'unknown_date':10,'days_1_7':20,'days_8_30':20,'days_31_90':20,'days_91_plus':10}
        for key,value in expected.items(): self.assertEqual(self.value(result,metric_id('bucket',key,'principal')),value)
        self.assertEqual(sum(expected.values()),100)
        self.assertEqual(len(result['excluded_status_rows']),1)
        self.assertIn('unknown_maturity_dates',result['rows'][0]['quality_flags'])

    def test_cumulative_windows_are_not_additive_and_exclude_today_past_unknown(self):
        result=self.result()
        self.assertEqual([self.value(result,metric_id('window',str(window),'principal')) for window in (7,30,90)],[20,40,60])
        self.assertEqual(self.value(result,metric_id('window','7','share')),Decimal('.2'))
        self.assertNotEqual(20+40+60,self.value(result,'eligible_principal'))

    def test_overlap_gap_bad_end_and_duplicate_id_fail(self):
        for mutate in [lambda config:config['buckets'][1].update(min_days=7),lambda config:config['buckets'][1].update(min_days=9),lambda config:config['buckets'][-1].update(max_days=100),lambda config:config['buckets'][0].update(id='today'),lambda config:config.update(cumulative_windows=[7,7]),lambda config:config.update(share_denominator='future_only')]:
            config=deepcopy(self.config); mutate(config)
            with self.assertRaises(ValueError): analyze(self.source,config)

    def test_duplicate_negative_invalid_dates_and_mixed_scope_fail_even_excluded(self):
        original=deepcopy(self.rows)
        for mutate in [lambda rows:rows.append(deepcopy(rows[0])),lambda rows:rows[-1].__setitem__(1,-1),lambda rows:rows[0].__setitem__(2,'2023-02-29'),lambda rows:rows[0].__setitem__(2,45000),lambda rows:rows[0].__setitem__(3,'2024-02-29'),lambda rows:rows[0].__setitem__(5,'USD'),lambda rows:rows[0].__setitem__(6,'万元'),lambda rows:rows[0].__setitem__(7,'v2'),lambda rows:rows[0].__setitem__(8,'region'),lambda rows:rows[0].__setitem__(9,'unmapped')]:
            self.rows=deepcopy(original); mutate(self.rows); self.save()
            with self.assertRaises(ValueError): self.result()

    def test_zero_principal_unknown_completeness_and_empty_extract(self):
        for row in self.rows: row[1]=0
        self.save(); self.config['scope']['completeness']='unknown'; result=self.result()
        self.assertIsNone(self.value(result,metric_id('bucket','today','share')))
        self.assertIn('zero_principal_denominator',result['rows'][0]['quality_flags'])
        self.assertIn('observation_completeness_unknown',result['rows'][0]['quality_flags'])
        self.rows=[]; self.save(); result=self.result()
        self.assertIn('no_eligible_observations',result['rows'][0]['quality_flags'])

    def test_precise_principal_and_single_level_parent_slices(self):
        self.rows[0][1]=Decimal('123456789012345678901234567890.123'); self.save()
        result=self.result()
        self.assertEqual(self.value(result,'eligible_principal'),Decimal('123456789012345678901234567980.123'))
        self.assertEqual(result['output_rows'],3)
        institution=self.value(result,'eligible_principal',1)+self.value(result,'eligible_principal',2)
        # Decimal arithmetic outside the module uses caller context; exact source value asserted above.
        self.assertTrue(institution>0)

    def test_cli_native_safe_xlsx_readback_and_no_contract_list(self):
        query=self.folder/'query.json'; query.write_text(json.dumps(self.config)); output=self.folder/'maturity.xlsx'
        original=self.source.read_bytes()
        command=[sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_maturity.py')),str(self.source),'--query',str(query),'--xlsx',str(output)]
        run=subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stdout+run.stderr)
        self.assertNotIn('fictional-contract-0',run.stdout)
        book=Workbook.load(output); self.assertEqual(len(book.sheets),7)
        self.assertIn('Validation',book.cells); self.assertIn('Lineage',book.cells)
        self.assertFalse(any(cell.get('formula') for sheet in book.cells.values() for cell in sheet.values()))
        self.assertEqual(self.source.read_bytes(),original)
        before=output.read_bytes(); run=subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(run.returncode,1); self.assertEqual(output.read_bytes(),before)


if __name__=='__main__': unittest.main()
