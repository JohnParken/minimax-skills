# SPDX-License-Identifier: MIT
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree as ET

SCRIPTS = Path(__file__).resolve().parents[1] / 'scripts'
sys.path.insert(0, str(SCRIPTS))
from stdlib_xlsx import Workbook, NS, read_spec


def spec(cells, name='Sales'):
    return {'sheets': [{'name': name, 'cells': cells}]}


class WorkbookTests(unittest.TestCase):
    def test_decimal_round_trip_cross_sheet_and_cache(self):
        book = Workbook.create({'sheets': [
            {'name': 'Sales', 'cells': {'A1': {'value': 12}, 'B1': {'value': Decimal('19.90')},
                                       'C1': {'value': Decimal('.95')}, 'D1': {'formula': 'ROUND(A1*B1*C1,2)', 'style': 6}}},
            {'name': 'Summary', 'cells': {'A1': {'formula': "Sales!D1+1"}}}]})
        report = book.calculate()
        self.assertEqual(report['calculated_count'], 2)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'out.xlsx'
            book.save(path)
            again = Workbook.load(path)
            self.assertEqual(again.cells['Sales']['D1']['value'], Decimal('226.86'))
            self.assertEqual(again.cells['Summary']['A1']['value'], Decimal('227.86'))
            self.assertEqual(again.cells['Sales']['D1']['formula'], 'ROUND(A1*B1*C1,2)')
            self.assertEqual(again.cells['Sales']['D1']['style'], '6')
            self.assertEqual(again.cells['Sales']['D1']['cache_status'], 'unverified')

    def test_cache_types_and_values_only(self):
        cells = {'A1': {'formula': 'IF(1=1," yes ","no")'},
                 'A2': {'formula': '1=2'}, 'A3': {'formula': 'B9'}}
        book = Workbook.create(spec(cells))
        self.assertFalse(book.calculate()['issues'])
        book._index()
        self.assertEqual(book.cells['Sales']['A1']['value'], ' yes ')
        self.assertIs(book.cells['Sales']['A2']['value'], False)
        self.assertEqual(book.cells['Sales']['A3']['value'], Decimal(0))
        book = Workbook.create(spec(cells))
        self.assertEqual(book.calculate(values_only=True)['status'], 'calculated_subset')
        book._index()
        self.assertIsNone(book.cells['Sales']['A1']['formula'])
        self.assertEqual(book.cells['Sales']['A1']['value'], ' yes ')

    def test_unsupported_propagates_and_clears_stale_cache(self):
        book = Workbook.create(spec({'A1': {'formula': 'XLOOKUP(1,B1,B2)'}, 'A2': {'formula': 'A1+1'}}))
        doc, node = book.cell_node('Sales', 'A1')
        book.put_value(doc, node, Decimal(999), cache=True)
        book.flush()
        book._index()
        report = book.calculate()
        self.assertEqual(report['status'], 'incomplete')
        self.assertEqual(len(report['issues']), 2)
        book._index()
        self.assertIsNone(book.cells['Sales']['A1']['value'])
        self.assertIsNone(book.cells['Sales']['A2']['value'])
        self.assertEqual(book.calculate(values_only=True)['status'], 'failed')

    def test_blank_formula_dependency_counts_as_zero(self):
        book = Workbook.create(spec({'A1': {'formula': 'B9'}, 'A2': {'formula': 'COUNT(A1:A1)'}}))
        self.assertEqual(book.calculate()['status'], 'calculated_subset')
        book._index()
        self.assertEqual(book.cells['Sales']['A2']['value'], Decimal(1))

    def test_precision_as_displayed_not_silently_calculated(self):
        book = Workbook.create(spec({'A1': {'formula': '1/3'}}))
        doc = book.document('xl/workbook.xml')
        calc = book.element(doc, 'calcPr')
        calc.setAttribute('fullPrecision', '0')
        doc.documentElement.appendChild(calc)
        book.flush()
        book._index()
        self.assertEqual(book.calculate()['status'], 'incomplete')

    def test_errors_and_cycles_block_calculation(self):
        for cells in ({'A1': {'formula': '1/0'}},
                      {'A1': {'formula': 'A2'}, 'A2': {'formula': 'A1'}},
                      {'A1': {'formula': 'Missing!B1'}}):
            with self.subTest(cells=cells):
                book = Workbook.create(spec(cells))
                self.assertEqual(book.calculate()['status'], 'failed')

    def test_edit_preserves_unrelated_members_and_source(self):
        book = Workbook.create({'sheets': [
            {'name': 'Sales', 'cells': {'A1': {'value': 2}, 'B1': {'formula': 'A1*3'}}},
            {'name': 'Other', 'cells': {'A1': {'value': 'untouched'}}}]})
        book.calculate()
        book.parts['custom/binary.bin'] = bytes(range(256))
        with tempfile.TemporaryDirectory() as directory:
            original = Path(directory) / 'original.xlsx'
            output = Path(directory) / 'edited.xlsx'
            book.save(original)
            before = original.read_bytes()
            edit = Workbook.load(original)
            edit.patch(spec({'A1': {'value': 4}}))
            report = edit.calculate()
            edit.save(output)
            self.assertEqual(original.read_bytes(), before)
            loaded = Workbook.load(output)
            self.assertEqual(loaded.cells['Sales']['B1']['value'], Decimal(12))
            with zipfile.ZipFile(original) as left, zipfile.ZipFile(output) as right:
                for member in left.namelist():
                    if member not in report['changed_parts']:
                        self.assertEqual(left.read(member), right.read(member), member)
            with self.assertRaises(ValueError):
                edit.save(original)

    def test_shared_formula_sheet_edit_refused_and_calc_unverified(self):
        book = Workbook.create(spec({'A1': {'formula': '1+1'}}))
        doc, node = book.cell_node('Sales', 'A1')
        node.getElementsByTagNameNS(NS[1:-1], 'f')[0].setAttribute('t', 'shared')
        book.flush()
        book._index()
        with self.assertRaises(ValueError):
            book.patch(spec({'B1': {'value': 3}}))
        self.assertEqual(book.calculate()['status'], 'incomplete')

    def test_bounds_json_and_literal_formula_injection(self):
        with self.assertRaises(ValueError):
            Workbook.create(spec({'XFE1': {'value': 1}}))
        book = Workbook.create(spec({'A1': {'value': '=SUM(B1:B9)'}, 'A2': {'value': ' 中文 & < > '}}))
        book._index()
        self.assertIsNone(book.cells['Sales']['A1']['formula'])
        self.assertEqual(book.cells['Sales']['A2']['value'], ' 中文 & < > ')
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'spec.json'
            path.write_text('{"sheets":[],"sheets":[]}', encoding='utf-8')
            with self.assertRaises(ValueError):
                read_spec(path)

    def test_cli_standard_library_only(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            source = directory / 'spec.json'
            output = directory / 'result.xlsx'
            source.write_text(json.dumps(spec({'A1': {'formula': 'ROUND(1.005,2)'}})), encoding='utf-8')
            run = subprocess.run([sys.executable, '-I', '-S', '-B', str(SCRIPTS / 'stdlib_xlsx.py'),
                                  'create', str(source), str(output)], capture_output=True, text=True)
            self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
            self.assertTrue(json.loads(run.stdout)['output_written'])
            run = subprocess.run([sys.executable, '-I', '-S', '-B', str(SCRIPTS / 'stdlib_xlsx.py'),
                                  'read', str(output)], capture_output=True, text=True)
            self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
            self.assertEqual(json.loads(run.stdout)['sheets'][0]['cells']['A1']['value'], '1.01')
            bad = directory / 'bad.xlsx'
            source.write_text(json.dumps(spec({'A1': {'formula': '1/0'}})), encoding='utf-8')
            run = subprocess.run([sys.executable, '-I', '-S', '-B', str(SCRIPTS / 'stdlib_xlsx.py'),
                                  'create', str(source), str(bad)], capture_output=True, text=True)
            self.assertEqual(run.returncode, 1, run.stdout + run.stderr)
            self.assertFalse(bad.exists())

    def test_calc_chain_removed_and_other_relations_retained(self):
        book = Workbook.create(spec({'A1': {'formula': '1+1'}}))
        rel_path = 'xl/_rels/workbook.xml.rels'
        relationships = ET.fromstring(book.parts[rel_path])
        pkg = '{http://schemas.openxmlformats.org/package/2006/relationships}'
        ET.SubElement(relationships, pkg + 'Relationship', {
            'Id': 'chain', 'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
            'Target': 'calcChain.xml'})
        book.parts[rel_path] = ET.tostring(relationships)
        book.parts['xl/calcChain.xml'] = ('<calcChain xmlns="' + NS[1:-1] + '"/>').encode()
        types = ET.fromstring(book.parts['[Content_Types].xml'])
        ET.SubElement(types, '{http://schemas.openxmlformats.org/package/2006/content-types}Override', {
            'PartName': '/xl/calcChain.xml', 'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'})
        book.parts['[Content_Types].xml'] = ET.tostring(types)
        self.assertEqual(book.calculate()['status'], 'calculated_subset')
        self.assertNotIn('xl/calcChain.xml', book.parts)
        self.assertNotIn(b'calcChain', book.parts[rel_path])
        self.assertNotIn(b'calcChain', book.parts['[Content_Types].xml'])
        self.assertIn(b'/styles', book.parts[rel_path])

    def test_macro_part_and_extension_preservation(self):
        book = Workbook.create(spec({'A1': {'value': 2}}))
        book.parts['[Content_Types].xml'] = book.parts['[Content_Types].xml'].replace(
            b'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
            b'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
        # Opaque test sentinel verifies preservation only, not executable VBA validity.
        book.parts['xl/vbaProject.bin'] = b'opaque-vba-preservation-sentinel'
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'macros.xlsm'
            with self.assertRaises(ValueError):
                book.save(Path(directory) / 'wrong.xlsx')
            book.save(path)
            edited = Workbook.load(path)
            edited.patch(spec({'A1': {'value': 4}}))
            edited.calculate()
            output = Path(directory) / 'edited.xlsm'
            edited.save(output)
            with zipfile.ZipFile(output) as archive:
                self.assertEqual(archive.read('xl/vbaProject.bin'), b'opaque-vba-preservation-sentinel')

    def test_cli_incomplete_and_csv_read(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            source = directory / 'spec.json'
            output = directory / 'incomplete.xlsx'
            source.write_text(json.dumps(spec({'A1': {'formula': 'UNKNOWN(1)'}})), encoding='utf-8')
            run = subprocess.run([sys.executable, '-I', '-S', '-B', str(SCRIPTS / 'stdlib_xlsx.py'),
                                  'create', str(source), str(output)], capture_output=True, text=True)
            self.assertEqual(run.returncode, 2, run.stdout + run.stderr)
            self.assertTrue(json.loads(run.stdout)['output_written'])
            self.assertIsNone(Workbook.load(output).cells['Sales']['A1']['value'])
            csv_path = directory / 'text.csv'
            csv_path.write_text('name,id\n"a,b",001\n', encoding='utf-8-sig')
            run = subprocess.run([sys.executable, '-I', '-S', '-B', str(SCRIPTS / 'stdlib_xlsx.py'),
                                  'read', str(csv_path)], capture_output=True, text=True)
            self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
            self.assertEqual(json.loads(run.stdout)['rows'][1], ['a,b', '001'])

    def test_merge_anchor_guard(self):
        book = Workbook.create(spec({'A1': {'value': 'Merged'}}))
        doc = book.document(book.sheets['Sales'])
        merges = book.element(doc, 'mergeCells')
        merge = book.element(doc, 'mergeCell')
        merge.setAttribute('ref', 'A1:B1')
        merges.appendChild(merge)
        doc.documentElement.appendChild(merges)
        book.flush()
        with self.assertRaises(ValueError):
            book.patch(spec({'B1': {'value': 2}}))


if __name__ == '__main__':
    unittest.main()
