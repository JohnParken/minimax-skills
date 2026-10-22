#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook
from xlsx_analysis import typed_value
from xlsx_headers import reconstruct, resolve_column


def merged_book(cells, merges):
    book = Workbook.create({'sheets': [{'name': '报表', 'cells': {ref: {'value': value} for ref, value in cells.items()}}]})
    if merges:
        doc = book.document(book.sheets['报表'])
        container = book.element(doc, 'mergeCells')
        for ref in merges:
            merge = book.element(doc, 'mergeCell')
            merge.setAttribute('ref', ref)
            container.appendChild(merge)
        doc.documentElement.appendChild(container)
        book.flush()
    return book


def chinese_book():
    return merged_book({'A1': '经营统计表', 'A2': '单位：万元；内部使用', 'A3': '部门',
                        'B3': '2024年', 'B4': '销售', 'D4': '风控',
                        'B5': '金额', 'C5': '笔数', 'D5': '金额', 'E5': '笔数',
                        'A6': '华东', 'B6': 10, 'C6': 2, 'D6': 4, 'E6': 1,
                        'A7': '华东', 'B7': 10, 'C7': 2, 'D7': 6, 'E7': 1,
                        'A8': '华西', 'B8': 30, 'C8': 3, 'D8': 8, 'E8': 1},
                       ['A1:E1', 'A3:A5', 'B3:E3', 'B4:C4', 'D4:E4'])


class HeaderTests(unittest.TestCase):
    def structure(self, book, rows='3:5', bounds=(1, 1, 8, 5), **kwargs):
        return reconstruct(book, '报表', bounds, rows,
                           lambda ref: typed_value(book.cells['报表'].get(ref, {})), **kwargs)

    def test_chinese_three_levels_and_both_merge_directions(self):
        schema = self.structure(chinese_book())
        columns = schema['columns']
        self.assertEqual([part['value'] for part in columns[0]['path']], ['部门'])
        self.assertEqual([part['value'] for part in columns[1]['path']], ['2024年', '销售', '金额'])
        self.assertEqual([part['value'] for part in columns[3]['path']], ['2024年', '风控', '金额'])
        self.assertEqual(columns[2]['levels'][0]['source'], 'B3')
        self.assertEqual(columns[2]['levels'][0]['merge'], 'B3:E3')
        self.assertEqual(columns[2]['levels'][0]['coordinate'], 'C3')
        self.assertEqual(schema['ignored_leading_rows'], [1, 2])
        self.assertEqual(schema['data_first_row'], 6)
        self.assertEqual(columns[1]['id'], 'col:B')

    def test_same_department_and_indicator_across_years(self):
        book=merged_book({'A1':'部门','B1':'2024年','D1':'2025年',
                          'B2':'销售','D2':'销售','B3':'金额','C3':'笔数','D3':'金额','E3':'笔数'},
                         ['A1:A3','B1:C1','D1:E1','B2:C2','D2:E2'])
        schema=self.structure(book,'1:3',(1,1,4,5))
        self.assertEqual(resolve_column({'path':['2024年','销售','金额']},['A','B','C','D','E'],schema),'B')
        self.assertEqual(resolve_column({'path':['2025年','销售','金额']},['A','B','C','D','E'],schema),'D')
        self.assertEqual([part['value'] for part in schema['columns'][0]['path']],['部门'])

    def test_blanks_do_not_forward_fill_and_repeated_paths_are_ambiguous(self):
        book = merged_book({'A1': '标题', 'A2': '重复', 'B2': '重复', 'C3': '金额'}, ['A1:C1'])
        schema = self.structure(book, '1:3', (1, 1, 4, 3))
        self.assertEqual(schema['columns'][2]['path'], [{'type': 'blank', 'value': None}, {'type': 'text', 'value': '金额'}])
        self.assertEqual(schema['context_headings'][0]['source'], 'A1')
        self.assertFalse(schema['context_headings'][0]['included_in_path'])
        with self.assertRaisesRegex(ValueError, 'matched 2'):
            resolve_column({'path': ['重复', None]}, ['A','B','C'], schema)
        self.assertEqual(resolve_column({'id': 'col:B'}, ['A','B','C'], schema), 'B')
        included = self.structure(book, '1:3', (1,1,4,3), include_spanning=True)
        self.assertEqual(included['columns'][0]['path'][0]['value'], '标题')

    def test_exclusions_and_equal_labels_at_different_levels(self):
        book = merged_book({'A1': '同名', 'A2': '说明', 'A3': '同名', 'B1': '其他'}, [])
        schema = self.structure(book, '1:3', (1,1,4,2), exclude=[2])
        self.assertEqual([part['value'] for part in schema['columns'][0]['path']], ['同名', '同名'])
        self.assertEqual(schema['columns'][0]['levels'][1]['role'], 'excluded')
        with self.assertRaises(ValueError):
            self.structure(book, '1:3', (1,1,4,2), exclude=[1,2,3])

    def test_invalid_overlap_data_boundary_and_crop_refused(self):
        for merges in (['A3:B3','B3:C3'], ['A3:A6'], ['A2:A3'], ['A3:A3'], ['B3:A3']):
            book = merged_book({'A3': '组'}, merges)
            with self.subTest(merges=merges), self.assertRaises(ValueError):
                self.structure(book)
        with self.assertRaisesRegex(ValueError, 'clips'):
            self.structure(chinese_book(), bounds=(3,3,8,5))
        with self.assertRaisesRegex(ValueError, 'Non-anchor'):
            self.structure(merged_book({'A3': '组', 'B3': '隐藏数据'}, ['A3:B3']))

    def test_typed_paths_and_stable_absolute_ids(self):
        schema = self.structure(merged_book({'B3': 2024, 'C3': '2024'}, []), '3', (3,2,4,3))
        self.assertEqual(resolve_column({'path': [2024]}, ['B','C'], schema), 'B')
        self.assertEqual(resolve_column({'path': ['2024']}, ['B','C'], schema), 'C')
        self.assertEqual(resolve_column('col:C', ['B','C'], schema), 'C')
        self.assertEqual(resolve_column({'path': [{'type':'number','value':'2024'}]}, ['B','C'], schema), 'B')

    def test_profile_group_filter_and_dedup_sort_by_paths(self):
        from xlsx_analysis import profile, summarize
        from xlsx_rows import select_rows
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)/'table.xlsx'
            chinese_book().save(path)
            before = path.read_bytes()
            report = profile(path, '报表', header_rows='3:5')
            self.assertEqual(report['source_rows'], 3)
            self.assertEqual(report['columns'][1]['numeric']['sum'], 50)
            self.assertEqual(report['columns'][1]['column_id'], 'col:B')
            query = {'group_by': [{'path':['部门']}],
                     'measure': {'path':['2024年','风控','金额']},
                     'filters': [{'column': {'path':['2024年','销售','金额']}, 'op':'ge','value':10}]}
            report = summarize(path, '报表', query, header_rows='3:5')
            self.assertEqual([group['numeric']['sum'] for group in report['groups']], [10,8])
            config = {'deduplicate': {'columns':[{'path':['部门']},{'id':'col:B'}]},
                      'sort':[{'column':{'path':['2024年','销售','金额']},'descending':True}]}
            report = select_rows(path, '报表', config, header_rows='3:5')
            self.assertEqual([row['source_row'] for row in report['rows']], [8,6])
            self.assertEqual(report['removed_duplicate_rows'], 1)
            self.assertEqual(path.read_bytes(), before)

    def test_ambiguous_paths_refuse_all_operation_selectors(self):
        from xlsx_analysis import summarize
        from xlsx_rows import select_rows
        book = merged_book({'A1':'同名','B1':'同名','A2':1,'B2':2}, [])
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'duplicate.xlsx'
            book.save(path)
            with self.assertRaisesRegex(ValueError,'matched 2'):
                summarize(path,'报表',{'measure':{'path':['同名']}},header_rows='1')
            for query in ({'deduplicate':{'columns':[{'path':['同名']}]}},
                          {'sort':[{'column':{'path':['同名']}}]},
                          {'filters':[{'column':{'path':['同名']},'op':'eq','value':1}]}):
                with self.subTest(query=query),self.assertRaisesRegex(ValueError,'matched 2'):
                    select_rows(path,'报表',query,header_rows='1')
            report = summarize(path,'报表',{'measure':{'id':'col:B'}},header_rows='1')
            self.assertEqual(report['groups'][0]['numeric']['sum'],2)

    def test_formula_headers_require_cache_and_count_anchor_once(self):
        from xlsx_analysis import profile
        book = merged_book({'A3':'placeholder','A4':1,'B4':2},['A3:B3'])
        book.patch({'sheets':[{'name':'报表','cells':{'A3':{'formula':'"金额"'}}}]})
        with tempfile.TemporaryDirectory() as directory:
            original=Path(directory)/'uncached.xlsx'
            book.save(original)
            for policy in ('reject','cached'):
                with self.assertRaises(ValueError):
                    profile(original,'报表',header_rows='3',formulas=policy)
            book.calculate()
            cached=Path(directory)/'cached.xlsx'
            book.save(cached)
            result=profile(cached,'报表',header_rows='3',formulas='cached',include_spanning_headings=True)
            self.assertEqual(result['unverified_formula_caches'],1)
            self.assertEqual(result['header_schema']['columns'][1]['path'][0]['value'],'金额')

    def test_cli_reconstruction_and_row_processing(self):
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'report.xlsx'
            chinese_book().save(path)
            query=Path(directory)/'query.json'
            query.write_text(json.dumps({'sort':[{'column':{'path':['2024年','风控','金额']},'descending':True}]}))
            base=[sys.executable,'-I','-S','-B']
            scripts=Path(__file__).resolve().parent
            run=subprocess.run(base+[str(scripts/'xlsx_headers.py'),str(path),'--sheet','报表','--header-rows','3:5'],capture_output=True,text=True)
            self.assertEqual(run.returncode,0,run.stderr+run.stdout)
            self.assertEqual(json.loads(run.stdout)['header_schema']['columns'][3]['id'],'col:D')
            run=subprocess.run(base+[str(scripts/'xlsx_rows.py'),str(path),'--sheet','报表','--header-rows','3:5','--query',str(query)],capture_output=True,text=True)
            self.assertEqual(run.returncode,0,run.stderr+run.stdout)
            self.assertEqual([row['source_row'] for row in json.loads(run.stdout)['rows']],[8,7,6])
            query.write_text(json.dumps({'group_by':[{'path':['部门']}], 'measure':{'path':['2024年','销售','金额']}}))
            run=subprocess.run(base+[str(scripts/'xlsx_analysis.py'),'summarize',str(path),'--sheet','报表','--header-rows','3:5','--query',str(query)],capture_output=True,text=True)
            self.assertEqual(run.returncode,0,run.stderr+run.stdout)
            self.assertEqual([group['numeric']['sum'] for group in json.loads(run.stdout)['groups']],['20','30'])
            run=subprocess.run(base+[str(scripts/'xlsx_analysis.py'),'profile',str(path),'--sheet','报表','--header-rows','3:5'],capture_output=True,text=True)
            self.assertEqual(run.returncode,0,run.stderr+run.stdout)
            self.assertEqual(json.loads(run.stdout)['source_rows'],3)


    def test_explicit_header_options_bounds_and_legacy_compatibility(self):
        from xlsx_analysis import profile
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'table.xlsx'
            merged_book({'A1':'旧表头','A2':7},[]).save(path)
            self.assertEqual(profile(path,'报表')['columns'][0]['header']['value'],'旧表头')
            self.assertEqual(profile(path,'报表',header_rows='1')['source_rows'],1)
            for options in ({'header_rows':'2:1'},{'header_rows':'0:1'},{'header_rows':'1:3'},
                            {'header_rows':'1','header':False},{'header_exclude':[1]},
                            {'header_rows':'1','header_exclude':[2]}):
                with self.subTest(options=options),self.assertRaises(ValueError):
                    profile(path,'报表',**options)


if __name__ == '__main__':
    unittest.main()
