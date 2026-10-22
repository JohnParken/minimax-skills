#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
import json
from decimal import Decimal
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook
from xlsx_analysis import profile, statistics


class AnalysisTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'input.xlsx'

    def save(self, rows):
        from xlsx_tabular import column_name
        cells = {column_name(col)+str(row): {'value': value} for row, values in enumerate(rows, 1)
                 for col, value in enumerate(values, 1)}
        book = Workbook.create({'sheets': [{'name': 'Data', 'cells': cells}]})
        book.save(self.path)

    def test_types_blanks_duplicates_and_decimal_stats(self):
        self.save([['type', 'amount'], ['A', Decimal('0.1')], ['A', Decimal('0.2')],
                   ['', '12'], [None, True], ['A', Decimal('0.1')]])
        before = self.path.read_bytes()
        report = profile(self.path, 'Data')
        self.assertEqual(report['duplicate_rows'], 1)
        self.assertEqual(report['source_rows'], 5)
        left, right = report['columns']
        self.assertEqual(left['blank_count'], 2)
        self.assertEqual(left['missing_cells'], 1)
        self.assertEqual(left['empty_text'], 1)
        self.assertEqual(left['duplicate_nonblank'], 2)
        self.assertTrue(right['mixed_types'])
        self.assertEqual(right['numeric']['sum'], Decimal('0.4'))
        self.assertEqual(right['numeric']['count'], 3)
        self.assertTrue(right['numeric']['mean_rounded'])
        self.assertEqual(self.path.read_bytes(), before)

    def test_offset_region_and_headerless(self):
        self.save([['ignore', 'title'], ['x', 5], ['y', 7]])
        result = profile(self.path, 'Data', 'B2:C3', header=False)
        self.assertEqual([c['column'] for c in result['columns']], ['B', 'C'])
        self.assertEqual(result['columns'][0]['numeric']['sum'], Decimal(12))
        self.assertEqual(result['columns'][1]['blank_count'], 2)
        self.assertIsNone(result['columns'][0]['header'])

    def test_header_only_and_blank_numeric_stats(self):
        self.save([['title']])
        result = profile(self.path, 'Data')
        self.assertEqual(result['source_rows'], 0)
        self.assertIsNone(result['columns'][0]['numeric']['sum'])
        self.assertEqual(result['duplicate_rows'], 0)

    def test_numeric_boolean_text_keys_are_distinct(self):
        self.save([[1], [True], ['1'], [Decimal('1.0')]])
        result = profile(self.path, 'Data', header=False)
        self.assertEqual(result['duplicate_rows'], 1)
        self.assertEqual(result['columns'][0]['distinct_nonblank'], 3)

    def test_formula_cache_policy(self):
        book = Workbook.create({'sheets': [{'name': 'Data', 'cells': {'A1': {'formula': '1+2'}}}]})
        book.save(self.path)
        for mode in ('reject', 'cached'):
            with self.assertRaises(ValueError):
                profile(self.path, 'Data', header=False, formulas=mode)
        book.calculate()
        cached = self.path.with_name('cached.xlsx')
        book.save(cached)
        result = profile(cached, 'Data', header=False, formulas='cached')
        self.assertEqual(result['unverified_formula_caches'], 1)
        self.assertEqual(result['columns'][0]['numeric']['sum'], Decimal(3))

    def test_region_bounds_and_unknown_sheet(self):
        self.save([[1]])
        for region in ('B2:A1', 'A1:XFD1048576', 'XFE1', 'A0', 'A1:B2:C3'):
            with self.subTest(region=region), self.assertRaises(ValueError):
                profile(self.path, 'Data', region)
        with self.assertRaises(ValueError):
            profile(self.path, 'missing')

    def test_exact_sum_and_rounding_contract(self):
        result = statistics([('number', Decimal('1e307')), ('number', Decimal('1')), ('number', Decimal('-1e307'))])
        self.assertEqual(result['sum'], 1)
        self.assertTrue(result['mean_rounded'])
        with self.assertRaises(ValueError):
            statistics([('number', Decimal('1e2000')), ('number', Decimal(1))])

    def test_iso_dates_serials_and_error_cells(self):
        book = Workbook.create({'sheets': [{'name': 'Data', 'cells': {'A1': {'value': '2026-10-22'}, 'B1': {'value': 46317}}}]})
        _, cell = book.cell_node('Data', 'A1')
        cell.setAttribute('t', 'd')
        book.put_value(book.document(book.sheets['Data']), cell, '2026-10-22', cache=True)
        cell.setAttribute('t', 'd')
        book.save(self.path)
        result = profile(self.path, 'Data', header=False)
        self.assertEqual(result['columns'][0]['types'], {'date_text': 1})
        self.assertEqual(result['columns'][1]['numeric']['sum'], 46317)
        _, cell = book.cell_node('Data', 'B1')
        book.put_value(book.document(book.sheets['Data']), cell, '#DIV/0!', cache=True)
        cell.setAttribute('t', 'e')
        error_path = self.path.with_name('error.xlsx')
        book.save(error_path)
        with self.assertRaisesRegex(ValueError, 'Excel error at B1'):
            profile(error_path, 'Data', header=False)

    def test_cli_json_and_failure(self):
        self.save([['Amount'], [Decimal('0.1')], [Decimal('0.2')]])
        command = [sys.executable, '-I', '-S', '-B', str(Path(__file__).with_name('xlsx_analysis.py')), 'profile', str(self.path), '--sheet']
        run = subprocess.run(command+['Data'], capture_output=True, text=True)
        self.assertEqual(run.returncode, 0, run.stderr)
        self.assertEqual(json.loads(run.stdout)['columns'][0]['numeric']['sum'], '0.3')
        run = subprocess.run(command+['missing'], capture_output=True, text=True)
        self.assertEqual(run.returncode, 1)
        self.assertEqual(json.loads(run.stdout)['status'], 'failed')

    def test_grouped_sums_filters_and_excluded_rows(self):
        from xlsx_analysis import summarize
        self.save([['Region', 'Value'], ['East', 10], ['West', 4], ['East', 20], ['West', '30'], ['East', None]])
        result = summarize(self.path, 'Data', {'group_by': ['A'], 'measure': 'B',
                           'filters': [{'column': 'B', 'op': 'ge', 'value': 10}]})
        self.assertEqual(result['selected_rows'], 2)
        self.assertEqual(result['excluded_rows'], 3)
        self.assertEqual(result['groups'][0]['numeric']['sum'], 30)
        self.assertEqual(result['groups'][0]['keys'][0]['value'], 'East')

    def test_mixed_measure_reject_or_explicit_ignore(self):
        from xlsx_analysis import summarize
        self.save([['Key', 'Value'], ['A', 2], ['A', '3'], ['A', True], ['A', ''], ['A', None]])
        with self.assertRaisesRegex(ValueError, 'B3'):
            summarize(self.path, 'Data', {'measure': 'B'})
        result = summarize(self.path, 'Data', {'measure': 'B', 'mixed': 'ignore'})
        group = result['groups'][0]
        self.assertEqual(group['numeric']['sum'], 2)
        self.assertEqual(group['ignored_non_numeric'], 2)
        self.assertEqual(group['blank_measure'], 2)
        self.assertEqual(group['rows'], 5)

    def test_typed_group_keys_and_multicolumn_groups(self):
        from xlsx_analysis import summarize
        self.save([['K', 'V'], [True, 1], [1, 1], ['1', 1], [None, 1], ['', 1], [Decimal('1.0'), 1]])
        result = summarize(self.path, 'Data', {'group_by': ['A', 'B']})
        self.assertEqual(len(result['groups']), 5)
        self.assertEqual(result['groups'][1]['rows'], 2)
        self.assertIsNone(result['groups'][0]['numeric'])

    def test_invalid_queries_and_empty_selection(self):
        from xlsx_analysis import summarize
        self.save([['A'], [1]])
        for query in ({'unknown': 1}, {'group_by': ['A', 'A']}, {'measure': 'Z'},
                      {'mixed': 'coerce'}, {'filters': [{'column': 'A', 'op': 'gt', 'value': '1'}]},
                      {'filters': [{'column': 'A', 'op': 'eq', 'value': {}}]}):
            with self.subTest(query=query), self.assertRaises(ValueError):
                summarize(self.path, 'Data', query)
        result = summarize(self.path, 'Data', {'filters': [{'column': 'A', 'op': 'eq', 'value': '1'}]})
        self.assertEqual(result['selected_rows'], 0)
        self.assertEqual(result['groups'], [])

    def test_summary_cli_decimal_json(self):
        self.save([['K', 'V'], ['A', Decimal('0.1')], ['A', Decimal('0.2')]])
        query = self.path.with_name('query.json')
        query.write_text('{"group_by":["A"],"measure":"B","filters":[{"column":"B","op":"ge","value":0.1}]}')
        run = subprocess.run([sys.executable, '-I', '-S', '-B', str(Path(__file__).with_name('xlsx_analysis.py')),
                              'summarize', str(self.path), '--sheet', 'Data', '--query', str(query)], capture_output=True, text=True)
        self.assertEqual(run.returncode, 0, run.stderr + run.stdout)
        self.assertEqual(json.loads(run.stdout)['groups'][0]['numeric']['sum'], '0.3')


if __name__ == '__main__':
    unittest.main()
