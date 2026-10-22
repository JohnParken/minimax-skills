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
from xlsx_rows import select_rows, read_query
from xlsx_tabular import column_name


class RowTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'source.xlsx'

    def save(self, rows):
        cells = {column_name(col)+str(row): {'value': value} for row, values in enumerate(rows, 1)
                 for col, value in enumerate(values, 1)}
        Workbook.create({'sheets': [{'name': 'Data', 'cells': cells}]}).save(self.path)

    def test_keep_first_last_unique_preserve_source_order(self):
        self.save([['ID', 'V'], ['001', 10], ['002', 20], ['001', 30], ['003', 40]])
        before = self.path.read_bytes()
        for keep, expected in (('first', [2, 3, 5]), ('last', [3, 4, 5]), ('unique', [3, 5])):
            result = select_rows(self.path, 'Data', {'deduplicate': {'columns': ['A'], 'keep': keep}})
            self.assertEqual([row['source_row'] for row in result['rows']], expected)
            self.assertEqual(result['duplicate_key_groups'], 1)
            self.assertEqual(result['removed_duplicate_rows'], 4-len(expected))
        self.assertEqual(self.path.read_bytes(), before)

    def test_typed_keys_and_long_identifiers_preserved(self):
        self.save([[True], [1], ['1'], ['00012345678901234567890'], [None], [''], [Decimal('1.0')]])
        result = select_rows(self.path, 'Data', {'deduplicate': {'columns': ['A']}}, header=False)
        self.assertEqual(result['output_rows'], 6)
        self.assertEqual(result['rows'][3]['cells'][0]['value'], '00012345678901234567890')

    def test_filter_then_deduplicate_then_stable_sort(self):
        self.save([['ID', 'V'], ['A', 1], ['B', 3], ['A', 2], ['C', 3]])
        result = select_rows(self.path, 'Data', {'filters': [{'column': 'B', 'op': 'gt', 'value': 1}],
                             'deduplicate': {'columns': ['A']}, 'sort': [{'column': 'B', 'descending': True}]})
        self.assertEqual([row['source_row'] for row in result['rows']], [3, 5, 4])
        self.assertEqual(result['filtered_out_rows'], 1)

    def test_multicolumn_sort_and_blank_placement(self):
        self.save([['K', 'V'], ['A', 2], ['B', 1], ['A', 1], [None, 9], ['', 8]])
        result = select_rows(self.path, 'Data', {'sort': [{'column': 'A'}, {'column': 'B', 'descending': True}]})
        self.assertEqual([row['source_row'] for row in result['rows']], [2, 4, 3, 5, 6])
        result = select_rows(self.path, 'Data', {'sort': [{'column': 'A', 'descending': True, 'blanks': 'first'}]})
        self.assertEqual([row['source_row'] for row in result['rows']], [5, 6, 3, 2, 4])

    def test_mixed_sort_has_documented_type_order(self):
        self.save([['2'], [2], [False], ['10'], [1]])
        result = select_rows(self.path, 'Data', {'sort': [{'column': 'A'}]}, header=False)
        self.assertEqual([row['source_row'] for row in result['rows']], [3, 5, 2, 4, 1])

    def test_invalid_configs_and_empty_selection(self):
        self.save([['H'], [1]])
        for query in ({'unknown': 1}, {'deduplicate': {'columns': []}}, {'deduplicate': {'columns': ['Z']}},
                      {'deduplicate': {'columns': ['A'], 'keep': 'all'}}, {'sort': [{'column': 'A', 'descending': 1}]},
                      {'sort': [{'column': 'A'}, {'column': 'A'}]}):
            with self.subTest(query=query), self.assertRaises(ValueError):
                select_rows(self.path, 'Data', query)
        result = select_rows(self.path, 'Data', {'filters': [{'column': 'A', 'op': 'eq', 'value': '1'}]})
        self.assertEqual(result['rows'], [])

    def test_query_limits_duplicates_and_invalid_numbers(self):
        query = self.path.with_suffix('.json')
        for text in ('{"sort":[],"sort":[]}', '{"filters":[{"column":"A","op":"eq","value":NaN}]}', ' '*65537):
            query.write_text(text)
            with self.assertRaises(ValueError):
                read_query(query)

    def test_cli_typed_json_and_formula_policy(self):
        self.save([['ID'], ['001']])
        query = self.path.with_suffix('.json')
        query.write_text('{"sort":[{"column":"A"}]}')
        run = subprocess.run([sys.executable, '-I', '-S', '-B', str(Path(__file__).with_name('xlsx_rows.py')),
                              str(self.path), '--sheet', 'Data', '--query', str(query)], capture_output=True, text=True)
        self.assertEqual(run.returncode, 0, run.stderr)
        self.assertEqual(json.loads(run.stdout)['rows'][0]['cells'][0], {'type': 'text', 'value': '001'})
        formula = self.path.with_name('formula.xlsx')
        book = Workbook.create({'sheets': [{'name': 'Data', 'cells': {'A1': {'formula': '1+1'}}}]})
        book.save(formula)
        for policy in ('reject', 'cached'):
            with self.assertRaises(ValueError):
                select_rows(formula, 'Data', {}, header=False, formulas=policy)
        book.calculate()
        cached = self.path.with_name('cached.xlsx')
        book.save(cached)
        result = select_rows(cached, 'Data', {}, header=False, formulas='cached')
        self.assertEqual(result['unverified_formula_caches'], 1)
        self.assertEqual(result['rows'][0]['cells'][0]['value'], Decimal(2))


if __name__ == '__main__':
    unittest.main()
