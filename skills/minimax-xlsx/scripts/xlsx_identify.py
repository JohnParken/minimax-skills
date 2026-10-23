#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Bounded content inspection; container recognition does not imply ET/ETT decoding."""
import argparse
from decimal import DecimalException
import io
import json
from pathlib import Path
import struct
import sys
import zipfile
import zlib
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook, MAX_BYTES, MAX_MEMBERS, xml
from xlsx_cfb import CompoundFile

MAX_FILE_BYTES = 32 * 1024 * 1024
COMPOUND_SIGNATURE = bytes.fromhex('D0CF11E0A1B11AE1')
GUIDANCE = 'Open a trusted copy locally in WPS Spreadsheets; Save As Excel Workbook (.xlsx) to a new file, then verify sheets, values, formulas and formatting. Keep the original ET/ETT template. Renaming is not conversion.'


def identify(path):
    with open(path, 'rb') as handle:
        data = handle.read(MAX_FILE_BYTES + 1)
    result = {'extension': Path(path).suffix.lower(), 'container': 'unknown',
              'status': 'unrecognized', 'parser_supported': False,
              'native_et_ett_parsed': False, 'guidance': GUIDANCE}
    if len(data) > MAX_FILE_BYTES:
        return {**result, 'status': 'limit_exceeded', 'message': 'Input exceeds 32 MiB inspection limit'}
    if not data:
        return {**result, 'status': 'invalid', 'message': 'Empty input'}
    if data.startswith(COMPOUND_SIGNATURE):
        valid = len(data) >= 512 and data[28:30] == b'\xfe\xff'
        if valid:
            major, shift = struct.unpack_from('<H', data, 26)[0], struct.unpack_from('<H', data, 30)[0]
            valid = (major, shift) in ((3, 9), (4, 12)) and len(data) >= 2**shift
        if not valid:
            return {**result,'container':'compound_binary','status':'invalid','header_consistent':False,
                    'support_level':'none','message':'Truncated/inconsistent compound header'}
        try:
            probe=CompoundFile(data).report()
            return {**result,**probe,'container_status':probe['status'],
                    'status':'encrypted_unsupported' if probe['status']=='encrypted_unsupported' else 'conversion_required',
                    'parser_supported':False,'header_consistent':True}
        except (ValueError,struct.error) as exc:
            return {**result,'container':'compound_binary','status':'invalid_or_unsupported',
                    'header_consistent':True,'support_level':'none','message':str(exc)}
    if not data.startswith((b'PK\x03\x04', b'PK\x05\x06', b'PK\x07\x08')):
        return result
    result['container'] = 'zip'
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            infos = archive.infolist()
            if len(infos) > MAX_MEMBERS or sum(info.file_size for info in infos) > MAX_BYTES:
                return {**result, 'status': 'limit_exceeded', 'message': 'Package exceeds member/uncompressed size limit'}
            names = [info.filename for info in infos]
            if len(names) != len(set(names)) or any(name.startswith('/') or '\\' in name or '..' in name.split('/') for name in names):
                raise ValueError('Duplicate or unsafe package names')
            if any(info.flag_bits & 1 for info in infos):
                return {**result, 'status': 'unsupported', 'message': 'Encrypted ZIP members are not supported'}
            if '[Content_Types].xml' not in names:
                return {**result, 'status': 'unsupported', 'message': 'ZIP has no OOXML content-type registry'}
            if archive.getinfo('[Content_Types].xml').file_size > 2 * 1024 * 1024:
                return {**result, 'status': 'limit_exceeded', 'message': 'Content-type registry exceeds 2 MiB'}
            registry = archive.read('[Content_Types].xml')
            if b'<!DOCTYPE' in registry.replace(b'\x00', b'').upper() or b'<!ENTITY' in registry.replace(b'\x00', b'').upper():
                raise ValueError('DTD/entity declarations are refused')
            types = xml(registry)
            namespace = '{http://schemas.openxmlformats.org/package/2006/content-types}'
            if types.tag != namespace + 'Types':
                raise ValueError('Invalid content-type registry root')
            allowed = ('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                       'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
            workbook_type = next((node.get('ContentType') for node in types if node.get('PartName') == '/xl/workbook.xml'), None)
            if workbook_type not in allowed:
                return {**result, 'status': 'unsupported', 'message': 'No supported XLSX/XLSM workbook declaration at xl/workbook.xml'}
            result['container'] = 'ooxml_spreadsheet'
            if any(name.startswith('_xmlsignatures/') for name in names):
                return {**result, 'status': 'unsupported', 'message': 'Signed package is outside the maintained workflow'}
            parts = {info.filename: archive.read(info) for info in infos}
            for name, content in parts.items():
                if name.endswith(('.xml', '.rels')):
                    flattened = content.replace(b'\x00', b'').upper()
                    if b'<!DOCTYPE' in flattened or b'<!ENTITY' in flattened:
                        raise ValueError('DTD/entity declarations are refused')
            book = Workbook(parts)
            extension = '.xlsm' if 'macroEnabled' in workbook_type else '.xlsx'
            accepted = result['extension'] == extension
            return {**result, 'status': 'supported_subset' if accepted else 'extension_mismatch',
                    'parser_supported': accepted, 'content_parser_probe': 'passed',
                    'expected_extension': extension, 'sheets': list(book.sheets),
                    'message': 'OOXML worksheet indexing passed; formulas, rendering and application compatibility are not verified. ET/ETT native decoding is not implemented.'}
    except (ValueError, DecimalException, KeyError, IndexError, ET.ParseError, zipfile.BadZipFile, zlib.error, RuntimeError, NotImplementedError, OverflowError) as exc:
        return {**result, 'status': 'invalid_or_unsupported', 'message': str(exc)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input')
    args = parser.parse_args()
    try:
        result = identify(args.input)
    except OSError as exc:
        print(json.dumps({'status': 'failed', 'message': str(exc)}))
        return 1
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
