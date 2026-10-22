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
import zipfile

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook
from test_xlsx_hierarchy import config_with_tree, two_branches
from test_xlsx_management import spec_for
from xlsx_management import build as wide, column_id
from xlsx_dashboard import build, export_dashboard


def query():
    return {'overview_period':'2025-01','metrics':['balance','risk_rate'],
            'ranks':{'balance':{'method':'competition','descending':True}},
            'thresholds':[{'metric_id':'risk_rate','field':'actual','op':'gt','value':5,
                           'unit':'%','severity':'management_attention',
                           'scope':{'currency':'CNY','version':'v1','report_level':'branch'}}]}


class DashboardTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.source=Path(self.temp.name)/'facts.xlsx'
        Workbook.create(spec_for(two_branches())).save(self.source)
        self.report=wide(self.source,'Facts',config_with_tree())

    def test_rank_level_isolation_alerts_and_no_aggregation(self):
        result=build(self.report,query())
        records=[row for row in result['overview'] if row['metric_id']=='balance']
        branch=[row for row in records if row['dimensions']['report_level']=='branch']
        self.assertEqual([(row['actual'],row['rank'],row['peer_count']) for row in branch],[(Decimal(400),1,2),(Decimal(200),2,2)])
        self.assertTrue(all(row['peer_count']==1 for row in records if row not in branch))
        self.assertEqual([row['status'] for row in result['alerts']],['within_configured_threshold','triggered'])
        self.assertTrue(all(row['target'] is None and row['gap'] is None for row in records))
        self.assertFalse(result['charts_created'])

    def test_missing_and_quality_never_classified_normal(self):
        index=self.report['columns'].index(column_id('metric','risk_rate','actual'))
        self.report['rows'][0]['cells'][index]={'type':'blank','value':None}
        result=build(self.report,query())
        self.assertEqual(result['alerts'][0]['status'],'not_evaluated_missing')
        row=self.report['rows'][-1]; row['quality_flags']=['incomplete']
        index=self.report['columns'].index(column_id('quality','flags'))
        row['cells'][index]['value']=json.dumps(row['quality_flags'])
        result=build(self.report,query())
        self.assertEqual(result['alerts'][-1]['status'],'not_evaluated_quality')
        self.assertEqual(result['overview'][-2]['rank_status'],'excluded_quality')

    def test_ties_and_direction(self):
        index=self.report['columns'].index(column_id('metric','balance','actual'))
        self.report['rows'][-1]['cells'][index]['value']=Decimal(400)
        for method,expected in [('competition',[1,1]),('dense',[1,1]),('ordinal',[1,2])]:
            config=query(); config['ranks']['balance']={'method':method,'descending':False}
            rows=build(self.report,config)['overview']
            self.assertEqual([row['rank'] for row in rows if row['metric_id']=='balance' and row['dimensions']['report_level']=='branch'],expected)

    def test_context_currency_version_and_period_separation(self):
        config=config_with_tree(); config['hierarchy']['report_levels']=['branch']
        facts=two_branches(); extra=deepcopy(facts)
        for row in extra: row[3]='USD'
        later=deepcopy(facts)
        for row in later: row[0]='2025-03'
        self.source.unlink(); Workbook.create(spec_for(facts+extra+later)).save(self.source)
        result=build(wide(self.source,'Facts',config),query())
        self.assertTrue(all(row['peer_count']==2 for row in result['overview'] if row['metric_id']=='balance'))
        self.assertEqual({row['dimensions']['period'] for row in result['trends']},{'2025-01','2025-03'})
        self.assertEqual(len(result['alerts']),2)

    def test_targets_preserved(self):
        config=config_with_tree(); config['targets']=[{'dimensions':{'period':'2025-01','institution':'R','line':'retail','currency':'CNY','version':'v1'},'metric_id':'balance','unit':'万元','kind':'stock','time_basis':'point_in_time','definition_version':'v1','report_level':'region','mode':'at_least','value':700}]
        result=build(wide(self.source,'Facts',config),query())
        row=next(row for row in result['overview'] if row['metric_id']=='balance' and row['dimensions']['institution']=='R')
        self.assertEqual((row['target'],row['gap']),(Decimal(700),Decimal(100)))

    def test_reject_bad_schema_units_scope_duplicates_and_numeric_text(self):
        for mutate in [lambda value:value.update(status='selected'),lambda value:value['rows'].append(deepcopy(value['rows'][0])),lambda value:value['rows'][0]['quality_flags'].append('hidden')]:
            report=deepcopy(self.report); mutate(report)
            with self.assertRaises(ValueError): build(report,query())
        for mutate in [lambda value:value.update(overview_period=None),lambda value:value.update(metrics=['unknown']),lambda value:value['thresholds'][0].update(unit='元'),lambda value:value['thresholds'][0]['scope'].pop('currency'),lambda value:value.update(ranks={'balance':{'method':'dense','descending':'true'}})]:
            config=query(); mutate(config)
            with self.assertRaises(ValueError): build(self.report,config)
        index=self.report['columns'].index(column_id('metric','balance','actual'))
        self.report['rows'][0]['cells'][index]={'type':'text','value':'400'}
        with self.assertRaisesRegex(ValueError,'typed numbers'): build(self.report,query())

    def test_no_rules_means_no_normal_and_absent_period_warns(self):
        config=query(); config['thresholds']=[]; config['overview_period']='2025-02'
        result=build(self.report,config)
        self.assertEqual(result['alerts'],[]); self.assertEqual(result['overview'],[])
        self.assertTrue(result['warnings']); self.assertEqual(len(result['trends']),8)

    def test_cli_safe_xlsx_readback_and_overwrite_refusal(self):
        source=Path(self.temp.name)/'wide.json'; source.write_text(json.dumps(self.report,default=str))
        config=source.with_name('query.json'); config.write_text(json.dumps(query()))
        output=source.with_name('dashboard.xlsx')
        command=[sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_dashboard.py')),str(source),'--query',str(config),'--xlsx',str(output)]
        run=subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stdout+run.stderr)
        book=Workbook.load(output)
        self.assertEqual(len(book.sheets),6)
        self.assertTrue(any(cell['value']=='00012345678901234567890' for cell in book.cells['Overview'].values()))
        self.assertTrue(all(not cell.get('formula') for cells in book.cells.values() for cell in cells.values()))
        with zipfile.ZipFile(output) as package:
            self.assertFalse(any('/charts/' in name for name in package.namelist()))
            self.assertIn(b'pane',package.read('xl/worksheets/sheet1.xml'))
        original=output.read_bytes(); run=subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(run.returncode,1); self.assertEqual(output.read_bytes(),original)

    def test_unverified_caches_block_clean_ranks_and_alerts(self):
        self.report['unverified_formula_caches']=1
        result=build(self.report,query())
        self.assertTrue(all(row['rank_status']=='excluded_quality' for row in result['overview'] if row['metric_id']=='balance'))
        self.assertTrue(all(row['status']=='not_evaluated_quality' for row in result['alerts']))
        self.assertTrue(result['warnings'])

    def test_malformed_operator_is_diagnostic(self):
        config=query(); config['thresholds'][0]['op']=[]
        with self.assertRaises(ValueError): build(self.report,config)

    def test_literal_formula_text_and_precision(self):
        index=self.report['columns'].index(column_id('dimension','institution'))
        self.report['rows'][0]['cells'][index]['value']='=1+2'
        index=self.report['columns'].index(column_id('metric','balance','actual'))
        self.report['rows'][0]['cells'][index]['value']=Decimal('12345678901234567890.1')
        output=Path(self.temp.name)/'safe.xlsx'; export_dashboard(build(self.report,query()),output)
        cells=Workbook.load(output).cells['Overview']
        self.assertTrue(any(cell['value']=='=1+2' and not cell.get('formula') for cell in cells.values()))
        self.assertTrue(any(cell['value']=='12345678901234567890.1' for cell in cells.values()))


if __name__=='__main__': unittest.main()
