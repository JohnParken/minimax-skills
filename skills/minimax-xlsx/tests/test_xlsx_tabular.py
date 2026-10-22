# SPDX-License-Identifier: MIT
import csv
import json
import subprocess
import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1] / 'scripts'
sys.path.insert(0, str(SCRIPTS))
from stdlib_xlsx import Workbook
from xlsx_tabular import import_table, export_table


class TabularTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / 'input.csv'
        self.book = self.root / 'book.xlsx'
        self.output = self.root / 'output.csv'

    def write(self, rows):
        with self.source.open('w', encoding='utf-8-sig', newline='') as handle:
            csv.writer(handle).writerows(rows)

    def test_literal_import_roundtrip_preserves_identifiers_and_multiline(self):
        rows = [['id', 'value', 'note'], ['0012', '12345678901234567890', '=1+2'],
                ['中文', 'line\n two', 'comma, quote" and &<>'], ['', '', '']]
        self.write(rows)
        original = self.source.read_bytes()
        report = import_table(self.source, self.book)
        self.assertEqual(report['formulas_created'], 0)
        self.assertEqual(self.source.read_bytes(), original)
        loaded = Workbook.load(self.book)
        self.assertEqual(loaded.cells['Sheet1']['A2']['value'], '0012')
        self.assertIsNone(loaded.cells['Sheet1']['C2']['formula'])
        export_table(self.book, self.output, 'Sheet1', text_mode='raw')
        with self.output.open(encoding='utf-8-sig', newline='') as handle:
            self.assertEqual(list(csv.reader(handle)), rows)

    def test_explicit_numeric_columns_keep_headers_and_blanks(self):
        self.write([['id', 'amount'], ['001', '19.90'], ['002', '']])
        import_table(self.source, self.book, numeric_columns=['b'])
        cells = Workbook.load(self.book).cells['Sheet1']
        self.assertEqual(cells['B1']['value'], 'amount')
        self.assertEqual(cells['B2']['value'], Decimal('19.90'))
        self.assertEqual(cells['B3']['value'], '')
        self.assertEqual(cells['A2']['value'], '001')

    def test_numeric_failures_create_no_output(self):
        for value in ('1,000', 'NaN', '=1+1', '1234567890123456', '1e309'):
            with self.subTest(value=value):
                self.write([['amount'], [value]])
                with self.assertRaises((ValueError, ArithmeticError)):
                    import_table(self.source, self.book, numeric_columns=['A'])
                self.assertFalse(self.book.exists())

    def test_ragged_empty_and_out_of_table_column_fail(self):
        for rows in ([], [['a', 'b'], ['c']], [[]]):
            self.write(rows)
            with self.assertRaises(ValueError):
                import_table(self.source, self.book)
            self.assertFalse(self.book.exists())
        self.write([['a']])
        with self.assertRaises(ValueError):
            import_table(self.source, self.book, numeric_columns=['B'])

    def test_tsv_encoding_and_headerless_numbers(self):
        self.source = self.root / 'input.tsv'
        self.source.write_bytes('中文\t2.5\n名称\t3\n'.encode('gbk'))
        import_table(self.source, self.book, encoding='gbk', numeric_columns=['B'], header_rows=0)
        self.output = self.root / 'output.tsv'
        export_table(self.book, self.output, 'Sheet1', 'B1:B2')
        self.assertEqual(self.output.read_text(encoding='utf-8-sig'), '2.5\n3\n')

    def test_existing_outputs_are_preserved(self):
        self.write([['a']])
        import_table(self.source, self.book)
        before = self.book.read_bytes()
        with self.assertRaises(ValueError):
            import_table(self.source, self.book)
        self.assertEqual(self.book.read_bytes(), before)
        self.output.write_text('do not replace')
        with self.assertRaises(FileExistsError):
            export_table(self.book, self.output, 'Sheet1')
        self.assertEqual(self.output.read_text(), 'do not replace')

    def test_formula_cache_explicit_opt_in_and_missing_cache(self):
        book = Workbook.create({'sheets': [{'name': 'S', 'cells': {'A1': {'formula': '2+3'}}}]})
        book.save(self.book)
        for mode in ('reject', 'cached'):
            with self.assertRaises(ValueError):
                export_table(self.book, self.output, 'S', formula_mode=mode)
            self.assertFalse(self.output.exists())
        book.calculate()
        second = self.root / 'cached.xlsx'
        book.save(second)
        report = export_table(second, self.output, 'S', formula_mode='cached')
        self.assertEqual(report['unverified_formula_caches'], 1)
        self.assertEqual(self.output.read_text(encoding='utf-8-sig').strip(), '5')

    def test_safe_text_export_and_sparse_range(self):
        book = Workbook.create({'sheets': [{'name': 'S', 'cells': {
            'B2': {'value': '=1+1'}, 'D2': {'value': '  @cmd'}, 'B3': {'value': -3},
            'C3': {'value': True}}}]})
        book.save(self.book)
        report = export_table(self.book, self.output, 'S', 'B2:D3')
        self.assertEqual(report['escaped_text_cells'], 2)
        with self.output.open(encoding='utf-8-sig', newline='') as handle:
            self.assertEqual(list(csv.reader(handle)), [["'=1+1", '', "'  @cmd"], ['-3', 'TRUE', '']])

    def test_range_bounds_and_unknown_sheet(self):
        self.write([['a']])
        import_table(self.source, self.book)
        for region in ('B2:A1', 'A1:XFD1048576', 'XFE1', 'A0', 'A1:B2:C3'):
            with self.subTest(region=region), self.assertRaises(ValueError):
                export_table(self.book, self.output, 'Sheet1', region)
            self.assertFalse(self.output.exists())
        with self.assertRaises(ValueError):
            export_table(self.book, self.output, 'missing')

    def test_cli_import_export_and_failure_diagnostics(self):
        self.write([['name', 'value'], ['中文', '2']])
        base = [sys.executable, '-I', '-S', '-B', str(SCRIPTS / 'xlsx_tabular.py')]
        for arguments in (['import', str(self.source), str(self.book), '--numeric-columns', 'B'],
                          ['export', str(self.book), str(self.output), '--sheet', 'Sheet1']):
            result = subprocess.run(base + arguments, capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
            self.assertTrue(json.loads(result.stdout)['output_written'])
        result = subprocess.run(base + ['import', str(self.source), str(self.book), '--encoding', 'bad-encoding'], capture_output=True, text=True)
        self.assertEqual(result.returncode, 1)
        self.assertFalse(json.loads(result.stdout)['output_written'])
        self.assertNotIn('Traceback', result.stderr)


if __name__ == '__main__':
    unittest.main()
