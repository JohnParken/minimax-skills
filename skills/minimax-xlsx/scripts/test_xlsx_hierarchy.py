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
from test_xlsx_management import config_example,fictional_facts,spec_for
from xlsx_management import build,column_id
from xlsx_hierarchy import Hierarchy
from stdlib_xlsx import Workbook
from xlsx_result_export import export_result


def hierarchy_example():
    return {'mapping_version':'fictional-tree-v1','valid_from':'2025-01','valid_to':'2025-12',
            'levels':['bank','region','branch'],'report_levels':['bank','region','branch'],
            'nodes':[{'id':'ROOT','parent':None,'level':'bank','name':'虚构总行'},
                     {'id':'R','parent':'ROOT','level':'region','name':'虚构区域'},
                     {'id':'00012345678901234567890','parent':'R','level':'branch','name':'虚构甲支行'},
                     {'id':'0002','parent':'R','level':'branch','name':'虚构乙支行'}]}


def config_with_tree():
    config=config_example(); config['hierarchy']=hierarchy_example()
    for metric in config['metrics']:
        if metric['kind']!='ratio': metric['additive_dimensions'].append('institution')
    return config


def two_branches():
    first=fictional_facts(); second=deepcopy(first)
    for row in second:
        row[1]='0002'
        row[7]=50 if row[6]=='balance' and row[5]=='P1' else 150 if row[6]=='balance' else 20 if row[6]=='bad' and row[5]=='P1' else 0 if row[6]=='bad' else 2
    return first+second


class HierarchyTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.source=Path(self.temp.name)/'fictional.xlsx'
        self.config=config_with_tree()

    def save(self,facts):
        Workbook.create(spec_for(facts)).save(self.source)

    def rows_by_id(self,result):
        index=result['columns'].index(column_id('dimension','institution'))
        return {row['cells'][index]['value']:row for row in result['rows']}

    def test_leaf_once_per_ancestor_and_ratio_recomputed_at_each_level(self):
        self.save(two_branches()); before=self.source.read_bytes()
        result=build(self.source,'Facts',self.config); rows=self.rows_by_id(result)
        balance=result['columns'].index(column_id('metric','balance','actual'))
        ratio=result['columns'].index(column_id('metric','risk_rate','actual'))
        level=result['columns'].index(column_id('dimension','report_level'))
        self.assertEqual(rows['ROOT']['cells'][balance]['value'],600)
        self.assertEqual(rows['R']['cells'][ratio]['value'],5)
        self.assertEqual(rows['0002']['cells'][ratio]['value'],10)
        self.assertEqual(rows['ROOT']['cells'][level]['value'],'bank')
        self.assertEqual(len(result['lineage']),36)
        for identifier,row in rows.items():
            self.assertEqual(len(row['source_rows']),len(set(row['source_rows'])))
        self.assertFalse(result['hierarchy_summary']['historical_restatement'])
        self.assertEqual(self.source.read_bytes(),before)
        output=self.source.with_name('hierarchy.xlsx'); export_result(result,output)
        self.assertIn('Wide',Workbook.load(output).sheets)

    def test_cycles_duplicate_ids_multi_parent_and_unknown_parent_fail(self):
        cases=[]
        tree=hierarchy_example(); tree['nodes'][0]['parent']='0002'; cases.append(tree)
        tree=hierarchy_example(); tree['nodes'].append(deepcopy(tree['nodes'][-1])); cases.append(tree)
        tree=hierarchy_example(); tree['nodes'][-1]['parent']=['R','ROOT']; cases.append(tree)
        tree=hierarchy_example(); tree['nodes'][-1]['parent']='unknown'; cases.append(tree)
        for tree in cases:
            with self.subTest(tree=tree),self.assertRaises(ValueError): Hierarchy(tree)

    def test_parent_facts_unmapped_and_outside_validity_fail(self):
        for institution,period in (('ROOT','2025-01'),('unknown','2025-01'),('0002','2024-12')):
            facts=fictional_facts()
            for row in facts: row[1]=institution; row[0]=period
            self.save(facts)
            with self.subTest(institution=institution,period=period),self.assertRaises(ValueError):
                build(self.source,'Facts',self.config)
            self.source.unlink()

    def test_rollup_requires_institution_additive_permissions(self):
        self.save(two_branches()); config=config_example(); config['hierarchy']=hierarchy_example()
        with self.assertRaisesRegex(ValueError,'does not permit'):
            build(self.source,'Facts',config)
        config['hierarchy']['report_levels']=['branch']
        self.assertEqual(build(self.source,'Facts',config)['output_rows'],2)

    def test_target_matches_ancestor_level_and_currency_stays_separate(self):
        facts=two_branches(); usd=deepcopy(fictional_facts())
        for row in usd: row[3]='USD'
        self.save(facts+usd)
        self.config['hierarchy']['coverage']='observed'
        self.config['targets']=[{'dimensions':{'period':'2025-01','institution':'R','line':'retail','currency':'CNY','version':'v1'},
                                 'metric_id':'balance','unit':'万元','kind':'stock','time_basis':'point_in_time',
                                 'definition_version':'v1','report_level':'region','mode':'at_least','value':700}]
        result=build(self.source,'Facts',self.config)
        institution=result['columns'].index(column_id('dimension','institution')); currency=result['columns'].index(column_id('dimension','currency'))
        actual=result['columns'].index(column_id('metric','balance','actual'))
        root_values={row['cells'][currency]['value']:row['cells'][actual]['value'] for row in result['rows'] if row['cells'][institution]['value']=='ROOT'}
        self.assertEqual(root_values,{'CNY':Decimal(600),'USD':Decimal(400)})
        config=deepcopy(self.config); config['targets'][0]['report_level']='branch'
        with self.assertRaisesRegex(ValueError,'Target type'):
            build(self.source,'Facts',config)

    def test_missing_leaf_metric_is_not_silently_a_complete_parent_total(self):
        facts=two_branches()
        facts=[row for row in facts if not (row[1]=='0002' and row[6]=='balance')]
        self.save(facts)
        result=build(self.source,'Facts',self.config); rows=self.rows_by_id(result)
        index=result['columns'].index(column_id('metric','balance','actual'))
        self.assertIsNone(rows['ROOT']['cells'][index]['value'])
        self.assertIn('incomplete_leaf_metric:balance',rows['ROOT']['quality_flags'])
        self.config['hierarchy']['coverage']='observed'
        result=build(self.source,'Facts',self.config); rows=self.rows_by_id(result)
        self.assertEqual(rows['ROOT']['cells'][index]['value'],400)
        self.assertIn('observed_leaf_metric:balance',rows['ROOT']['quality_flags'])

    def test_cli_mapping_and_xlsx(self):
        self.save(two_branches()); query=self.source.with_suffix('.json'); query.write_text(json.dumps(self.config))
        output=self.source.with_name('wide.xlsx')
        run=subprocess.run([sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_management.py')),str(self.source),
                            '--sheet','Facts','--query',str(query),'--xlsx',str(output)],capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stderr+run.stdout)
        self.assertEqual(json.loads(run.stdout)['hierarchy_summary']['mapping_version'],'fictional-tree-v1')
        self.assertEqual(len(Workbook.load(output).sheets),7)


if __name__=='__main__': unittest.main()
