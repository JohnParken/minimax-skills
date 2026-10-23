#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
import json
from pathlib import Path
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook
from xlsx_identify import identify, COMPOUND_SIGNATURE


class IdentifyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'sample.et'

    def test_actual_ooxml_content_with_et_suffix_is_not_native_et_support(self):
        source = self.path.with_suffix('.xlsx')
        Workbook.create({'sheets': [{'name': 'Data', 'cells': {'A1': {'value': 1}}}]}).save(source)
        self.assertTrue(identify(source)['parser_supported'])
        self.path.write_bytes(source.read_bytes())
        report = identify(self.path)
        self.assertEqual(report['container'], 'ooxml_spreadsheet')
        self.assertEqual(report['status'], 'extension_mismatch')
        self.assertEqual(report['content_parser_probe'], 'passed')
        self.assertFalse(report['parser_supported'])
        self.assertFalse(report['native_et_ett_parsed'])

    def test_compound_container_header_is_only_a_container_hint(self):
        data = bytearray(512)
        data[:8] = COMPOUND_SIGNATURE
        data[28:30] = b'\xfe\xff'
        struct.pack_into('<H', data, 26, 3)
        struct.pack_into('<H', data, 30, 9)
        self.path.write_bytes(data)
        report = identify(self.path)
        self.assertEqual(report['status'], 'invalid_or_unsupported')
        self.assertTrue(report['header_consistent'])
        self.assertEqual(report['container'], 'compound_binary')
        self.assertFalse(report['native_et_ett_parsed'])
        self.path.write_bytes(COMPOUND_SIGNATURE)
        self.assertEqual(identify(self.path)['status'], 'invalid')

    def test_synthetic_compound_probe_is_not_et_ett_acceptance(self):
        from test_xlsx_cfb import synthetic_cfb
        self.path.write_bytes(synthetic_cfb({'Workbook':b'not a real workbook'}))
        report=identify(self.path)
        self.assertEqual(report['status'],'conversion_required')
        self.assertEqual(report['container_status'],'container_readable')
        self.assertEqual(report['support_level'],'cfb_container_only')
        self.assertFalse(report['parser_supported']); self.assertFalse(report['biff_parsed'])
        self.assertFalse(report['native_et_ett_parsed'])

    def test_unknown_empty_damaged_and_generic_zip(self):
        for content, status in ((b'not a spreadsheet', 'unrecognized'), (b'', 'invalid'),
                                (b'PK\x03\x04broken', 'invalid_or_unsupported')):
            self.path.write_bytes(content)
            self.assertEqual(identify(self.path)['status'], status)
        with zipfile.ZipFile(self.path, 'w') as archive:
            archive.writestr('unrelated.txt', 'data')
        self.assertEqual(identify(self.path)['status'], 'unsupported')

    def test_bounded_input_and_package(self):
        self.path.write_bytes(b'123456789')
        with patch('xlsx_identify.MAX_FILE_BYTES', 8):
            self.assertEqual(identify(self.path)['status'], 'limit_exceeded')
        with zipfile.ZipFile(self.path, 'w') as archive:
            archive.writestr('one', 'one')
        with patch('xlsx_identify.MAX_BYTES', 1):
            self.assertEqual(identify(self.path)['status'], 'limit_exceeded')

    def test_cli_json_and_missing_input(self):
        self.path.write_bytes(b'unknown')
        command = [sys.executable, '-I', '-S', '-B', str(Path(__file__).with_name('xlsx_identify.py')), str(self.path)]
        run = subprocess.run(command, capture_output=True, text=True)
        self.assertEqual(run.returncode, 0)
        self.assertEqual(json.loads(run.stdout)['status'], 'unrecognized')
        self.path.unlink()
        run = subprocess.run(command, capture_output=True, text=True)
        self.assertEqual(run.returncode, 1)
        self.assertEqual(json.loads(run.stdout)['status'], 'failed')


if __name__ == '__main__':
    unittest.main()
