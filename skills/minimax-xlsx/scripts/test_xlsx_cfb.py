#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Independent synthetic MS-CFB structures, NOT real WPS/ET acceptance samples."""
import json
import random
from pathlib import Path
import struct
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parent))
from xlsx_cfb import CompoundFile,SIGNATURE,FREE,END,FAT_SECTOR


def entry(name,kind,start=END,size=0,right=FREE,child=FREE):
    block=bytearray(128); name_bytes=name.encode('utf-16le')+bytes(2)
    block[:len(name_bytes)]=name_bytes; struct.pack_into('<HBB3I',block,64,len(name_bytes),kind,1,FREE,right,child)
    struct.pack_into('<IQ',block,116,start,size); return block


def synthetic_cfb(streams=None,major=3):
    # This builder implements only tiny test fixtures and is not a file writer feature.
    streams={'数据':'中文数据'.encode('utf-8'),'Workbook':b'synthetic container payload, not a BIFF workbook'*100} if streams is None else streams
    size=512 if major==3 else 4096; directory_count=(len(streams)+1+size//128-1)//(size//128)
    sectors=[bytearray(size) for _ in range(1+directory_count)]
    fat=[FREE]*(size//4); fat[0]=FAT_SECTOR
    for index in range(1,1+directory_count): fat[index]=index+1 if index<directory_count else END
    small={name:content for name,content in streams.items() if 0<len(content)<4096}
    mini_fat=[FREE]*(size//4); mini_blob=bytearray(); starts={}
    for name,content in small.items():
        start=len(mini_blob)//64; blocks=(len(content)+63)//64; starts[name]=start
        for index in range(start,start+blocks): mini_fat[index]=index+1 if index+1<start+blocks else END
        mini_blob.extend(content+bytes(blocks*64-len(content)))
    mini_id=END; root_start=END
    if small:
        mini_id=len(sectors); sectors.append(bytearray(struct.pack('<'+'I'*len(mini_fat),*mini_fat))); fat[mini_id]=END
        root_start=len(sectors); blocks=(len(mini_blob)+size-1)//size
        for offset in range(blocks):
            identifier=len(sectors); content=mini_blob[offset*size:(offset+1)*size]; sectors.append(bytearray(content+bytes(size-len(content))))
            fat[identifier]=identifier+1 if offset+1<blocks else END
    for name,content in streams.items():
        if len(content)>=4096:
            starts[name]=len(sectors); blocks=(len(content)+size-1)//size
            for offset in range(blocks):
                identifier=len(sectors); data=content[offset*size:(offset+1)*size]; sectors.append(bytearray(data+bytes(size-len(data))))
                fat[identifier]=identifier+1 if offset+1<blocks else END
        elif not content: starts[name]=END
    directory=entry('Root Entry',5,root_start,len(mini_blob),child=1 if streams else FREE)
    for index,(name,content) in enumerate(streams.items(),1):
        directory.extend(entry(name,2,starts[name],len(content),right=index+1 if index<len(streams) else FREE))
    directory.extend(bytes(directory_count*size-len(directory)))
    for index in range(directory_count): sectors[index+1]=directory[index*size:(index+1)*size]
    sectors[0]=bytearray(struct.pack('<'+'I'*len(fat),*fat))
    header=bytearray(size); header[:8]=SIGNATURE
    struct.pack_into('<5H',header,24,0x3E,major,0xFFFE,9 if major==3 else 12,6)
    struct.pack_into('<9I',header,40,directory_count if major==4 else 0,1,1,0,4096,mini_id,1 if small else 0,END,0)
    struct.pack_into('<109I',header,76,0,*([FREE]*108))
    return bytes(header+b''.join(sectors))


class CompoundTests(unittest.TestCase):
    def test_synthetic_v3_regular_mini_unicode_and_empty(self):
        streams={'数据':'中文内容'.encode('utf-8'),'Empty':b'','Workbook':b'abcd'*1024}
        book=CompoundFile(synthetic_cfb(streams))
        for name,content in streams.items(): self.assertEqual(book.read_stream([name]),content)
        report=book.report(); self.assertEqual(report['support_level'],'cfb_container_only')
        self.assertEqual(report['workbook_stream_candidates'],[['Workbook']])
        self.assertFalse(report['biff_parsed']); self.assertFalse(report['native_et_ett_parsed'])

    def test_synthetic_v4_and_empty_container(self):
        for major in (3,4):
            book=CompoundFile(synthetic_cfb({'中文':b'x'*150},major))
            self.assertEqual(book.read_stream(['中文']),b'x'*150)
            self.assertEqual(book.report()['sector_size'],512 if major==3 else 4096)
        self.assertEqual(CompoundFile(synthetic_cfb({})).report()['streams'],[])

    def test_regular_fragmented_sector_order(self):
        content=b''.join(bytes([index])*512 for index in range(8))
        data=bytearray(synthetic_cfb({'Workbook':content})); start=struct.unpack_from('<I',data,1024+128+116)[0]
        order=[start,start+2,start+1]+list(range(start+3,start+8))
        for left,right in zip(order,order[1:]+[END]): struct.pack_into('<I',data,512+left*4,right)
        expected=content[:512]+content[1024:1536]+content[512:1024]+content[1536:]
        self.assertEqual(CompoundFile(bytes(data)).read_stream(['Workbook']),expected)

    def test_nested_storage_and_no_external_link_following(self):
        data=bytearray(synthetic_cfb({'Folder':b'','资料':b'http://example.invalid/never-follow'}))
        data[1024+128+66]=1
        struct.pack_into('<3I',data,1024+128+68,FREE,FREE,2)
        struct.pack_into('<I',data,1024+128+116,0)
        book=CompoundFile(bytes(data))
        self.assertEqual(book.read_stream(['Folder','资料']),b'http://example.invalid/never-follow')
        self.assertFalse(book.report()['external_links_followed'])

    def test_extended_difat_and_cycle_rejection(self):
        data=bytearray(synthetic_cfb({'one':b'1'})); physical=len(data)//512-1
        added=list(range(physical,physical+109)); difat=physical+109
        data.extend(bytes([255])*512*109)
        data.extend(struct.pack('<128I',added[-1],*([FREE]*126),END))
        struct.pack_into('<I',data,44,110); struct.pack_into('<II',data,68,difat,1)
        struct.pack_into('<109I',data,76,0,*added[:108])
        for identifier in added: struct.pack_into('<I',data,512+identifier*4,FAT_SECTOR)
        struct.pack_into('<I',data,512+difat*4,0xFFFFFFFC)
        self.assertEqual(CompoundFile(bytes(data)).read_stream(['one']),b'1')
        struct.pack_into('<I',data,72,2); struct.pack_into('<I',data,(difat+1)*512+508,difat)
        with self.assertRaisesRegex(ValueError,'cycle|cross-link'): CompoundFile(bytes(data))

    def test_regular_chain_cycle_size_and_cross_links_rejected(self):
        original=synthetic_cfb({'Workbook':b'x'*4096,'Other':b'y'*4096})
        first=struct.unpack_from('<I',original,1024+128+116)[0]
        for mutate in [lambda data:struct.pack_into('<I',data,512+first*4,first),lambda data:struct.pack_into('<Q',data,1024+128+120,5000),lambda data:struct.pack_into('<I',data,1024+256+116,first),lambda data:struct.pack_into('<I',data,512+first*4,123456)]:
            data=bytearray(original); mutate(data)
            with self.assertRaises(ValueError): CompoundFile(bytes(data))

    def test_mini_chain_cycles_cross_links_and_bounds(self):
        original=synthetic_cfb({'one':b'x'*100,'two':b'y'*64})
        mini=struct.unpack_from('<I',original,60)[0]
        for mutate in [lambda data:struct.pack_into('<I',data,(mini+1)*512,0),lambda data:struct.pack_into('<I',data,1024+256+116,0),lambda data:struct.pack_into('<I',data,1024+128+116,12345),lambda data:struct.pack_into('<Q',data,1024+120,64)]:
            data=bytearray(original); mutate(data)
            with self.assertRaises(ValueError): CompoundFile(bytes(data))

    def test_directory_cycles_orphans_bad_names_and_depth(self):
        original=synthetic_cfb({'one':b'1','two':b'2'})
        for mutate in [lambda data:struct.pack_into('<I',data,1024+76,0),lambda data:struct.pack_into('<I',data,1024+128+72,1),lambda data:struct.pack_into('<I',data,1024+128+72,FREE),lambda data:struct.pack_into('<H',data,1024+128+64,65),lambda data:struct.pack_into('<H',data,1024+128,ord('/'))]:
            data=bytearray(original); mutate(data)
            with self.assertRaises(ValueError): CompoundFile(bytes(data))
        with patch('xlsx_cfb.MAX_DEPTH',0):
            with self.assertRaisesRegex(ValueError,'depth'): CompoundFile(original)

    def test_header_truncation_padding_fat_and_resource_limits(self):
        original=synthetic_cfb()
        for data in (b'',original[:512],original[:-1],b'bad'+original[3:]):
            with self.assertRaises(ValueError): CompoundFile(data)
        for offset,value in ((44,513),(72,17),(56,2048),(68,0),(76,999999)):
            data=bytearray(original); struct.pack_into('<I',data,offset,value)
            with self.assertRaises(ValueError): CompoundFile(bytes(data))
        data=bytearray(synthetic_cfb(major=4)); data[512]=1
        with self.assertRaisesRegex(ValueError,'padding'): CompoundFile(bytes(data))
        with patch('xlsx_cfb.MAX_FILE_BYTES',1024):
            with self.assertRaises(ValueError): CompoundFile(original)
        with patch('xlsx_cfb.MAX_STREAM_BYTES',4):
            with self.assertRaisesRegex(ValueError,'8 MiB'): CompoundFile(original)
        with patch('xlsx_cfb.MAX_DIRECTORY_ENTRIES',1):
            with self.assertRaises(ValueError): CompoundFile(original)

    def test_encrypted_stream_names_never_decrypt_or_decode(self):
        book=CompoundFile(synthetic_cfb({'EncryptionInfo':b'not actual encryption info','EncryptedPackage':b'ciphertext'}))
        self.assertEqual(book.report()['status'],'encrypted_unsupported')
        with self.assertRaisesRegex(ValueError,'Encrypted'): book.read_stream(['EncryptedPackage'])
        self.assertFalse(book.report()['cells_decoded'])

    def test_unallocated_and_aliased_streams_not_silently_ignored(self):
        data=bytearray(synthetic_cfb({'one':b'1','two':b'2'}))
        # Make two active directory entries use the same name; no lossy dictionary overwrite.
        data[1024+256:1024+256+64]=data[1024+128:1024+128+64]
        with self.assertRaisesRegex(ValueError,'Duplicate'): CompoundFile(bytes(data))
        data=bytearray(synthetic_cfb({})); data.extend(bytes(512)); struct.pack_into('<I',data,512+2*4,END)
        with self.assertRaisesRegex(ValueError,'Unreferenced'): CompoundFile(bytes(data))

    def test_seeded_structure_mutations_stay_bounded_and_never_claim_cells(self):
        randomizer=random.Random(20261022)
        original=synthetic_cfb()
        for _ in range(200):
            data=bytearray(original)
            for _ in range(randomizer.randint(1,4)):
                offset=randomizer.randrange(min(len(data),1536)); data[offset]=randomizer.randrange(256)
            try:
                report=CompoundFile(bytes(data)).report()
            except ValueError:
                continue
            self.assertFalse(report['biff_parsed']); self.assertFalse(report['native_et_ett_parsed'])
            self.assertFalse(report['cells_decoded'])

    def test_cli_is_synthetic_container_only_and_no_files_written(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'synthetic-container.cfb'; path.write_bytes(synthetic_cfb())
            command=[sys.executable,'-I','-S','-B',str(Path(__file__).with_name('xlsx_cfb.py')),str(path)]
            before=path.read_bytes(); run=subprocess.run(command,capture_output=True,text=True,env={'PATH':''})
            self.assertEqual(run.returncode,0,run.stdout+run.stderr)
            report=json.loads(run.stdout); self.assertEqual(report['status'],'container_readable')
            self.assertFalse(report['biff_parsed']); self.assertFalse(report['native_et_ett_parsed'])
            self.assertEqual(path.read_bytes(),before); self.assertEqual(len(list(Path(folder).iterdir())),1)
            path.write_bytes(b'broken'); run=subprocess.run(command,capture_output=True,text=True)
            self.assertEqual(run.returncode,1); self.assertEqual(json.loads(run.stdout)['status'],'invalid_or_unsupported')


if __name__=='__main__': unittest.main()
