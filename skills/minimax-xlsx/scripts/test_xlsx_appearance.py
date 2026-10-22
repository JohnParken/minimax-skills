#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
from decimal import Decimal
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parent))
from stdlib_xlsx import Workbook, NS, xml
from xlsx_appearance import apply_appearance, HEADER_COLOR, HEADER_TEXT_COLOR, TEXT_COLOR
from xlsx_tabular import import_table


def specification(cells,name='Data'):
    return {'sheets':[{'name':name,'cells':cells}]}


def appearance(**settings):
    return {'preset':'readable','sheets':{'Data':settings}}


def cell_style(book,ref):
    root=xml(book.parts['xl/styles.xml'])
    index=int(book.cells['Data'][ref]['style'])
    return root,root.find(NS+'cellXfs')[index]


def color(book,ref):
    root,style=cell_style(book,ref)
    fill=root.find(NS+'fills')[int(style.get('fillId'))]
    return fill.find(NS+'patternFill/'+NS+'fgColor').get('rgb')


class AppearanceTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name)

    def test_header_bands_fonts_freeze_width_and_number_format(self):
        book=Workbook.create(specification({'A1':{'value':'编号'},'B1':{'value':'金额'},
                                           'A2':{'value':'001'},'B2':{'value':Decimal('19.90')},
                                           'A3':{'value':'002'},'B3':{'value':2}}),
                             appearance=appearance(number_formats={'B':'#,##0.00'},column_widths={'A':18,'B':22}))
        self.assertEqual(color(book,'A1'),HEADER_COLOR)
        self.assertNotEqual(color(book,'A2'),color(book,'A3'))
        styles,header=cell_style(book,'A1')
        font=styles.find(NS+'fonts')[int(header.get('fontId'))]
        self.assertEqual(font.find(NS+'color').get('rgb'),HEADER_TEXT_COLOR)
        self.assertIsNotNone(font.find(NS+'b'))
        self.assertEqual(header.find(NS+'alignment').get('wrapText'),'1')
        styles,body=cell_style(book,'B2')
        fmt=next(item for item in styles.find(NS+'numFmts') if item.get('numFmtId')==body.get('numFmtId'))
        self.assertEqual(fmt.get('formatCode'),'#,##0.00')
        self.assertEqual(styles.find(NS+'fonts')[int(body.get('fontId'))].find(NS+'color').get('rgb'),TEXT_COLOR)
        sheet=xml(book.parts[book.sheets['Data']])
        self.assertEqual(sheet.find(NS+'sheetViews/'+NS+'sheetView/'+NS+'pane').get('topLeftCell'),'A2')
        self.assertEqual(sheet.find(NS+'cols/'+NS+'col').get('width'),'18')
        self.assertEqual(sheet.find(NS+'sheetData/'+NS+'row').get('ht'),'30')
        self.assertIsNotNone(xml(book.parts['xl/workbook.xml']).find(NS+'bookViews/'+NS+'workbookView'))

    def test_values_formulas_and_identifiers_unchanged_after_save(self):
        spec=specification({'A1':{'value':'ID'},'B1':{'value':'V'},'A2':{'value':'0012345678901234567890'},
                            'B2':{'formula':'ROUND(1.005,2)'},'A3':{'value':'=1+1'},'B3':{'value':Decimal('0.10')}})
        plain=Workbook.create(spec)
        styled=Workbook.create(spec,appearance=appearance())
        for book in (plain,styled):
            self.assertFalse(book.calculate()['issues'])
        left,right=self.root/'plain.xlsx',self.root/'styled.xlsx'
        plain.save(left); styled.save(right)
        left_book,right_book=Workbook.load(left),Workbook.load(right)
        for ref in spec['sheets'][0]['cells']:
            for key in ('value','formula','type'):
                self.assertEqual(left_book.cells['Data'][ref][key],right_book.cells['Data'][ref][key])

    def test_category_colors_override_bands_and_noncontiguous_keys_repeat(self):
        spec=specification({'A1':{'value':'类别'},'B1':{'value':'值'},'A2':{'value':'甲'},'B2':{'value':1},
                            'A3':{'value':'乙'},'B3':{'value':2},'A4':{'value':'甲'},'B4':{'value':3}})
        book=Workbook.create(spec,appearance=appearance(category_columns=['A'],alternate_rows=True))
        self.assertEqual(color(book,'B2'),color(book,'B4'))
        self.assertNotEqual(color(book,'B2'),color(book,'B3'))
        self.assertEqual(color(book,'A1'),HEADER_COLOR)
        self.assertEqual(book.appearance_report['sheets'][0]['category_key_count'],2)
        self.assertEqual([book.cells['Data'][ref]['value'] for ref in ('A2','A3','A4')],['甲','乙','甲'])

    def test_composite_typed_categories_are_not_coerced(self):
        cells={'A1':{'value':'K'},'B1':{'value':'K2'},'C1':{'value':'V'}}
        for row,key in enumerate((1,True,'1',1),2):
            cells.update({f'A{row}':{'value':key},f'B{row}':{'value':'组'},f'C{row}':{'value':row}})
        book=Workbook.create(specification(cells),appearance=appearance(category_columns=['A','B']))
        self.assertEqual(len({color(book,ref) for ref in ('C2','C3','C4')}),3)
        self.assertEqual(color(book,'C2'),color(book,'C5'))

    def test_explicit_cell_style_wins_entirely(self):
        book=Workbook.create(specification({'A1':{'value':'Explicit','style':0},'B1':{'value':'Preset'},
                                           'A2':{'value':12,'style':5},'B2':{'value':1}}),
                             appearance=appearance(number_formats={'A':'0.00'}))
        self.assertEqual(book.cells['Data']['A1']['style'],'0')
        self.assertEqual(book.cells['Data']['A2']['style'],'5')
        self.assertEqual(color(book,'B1'),HEADER_COLOR)
        self.assertEqual(book.appearance_report['sheets'][0]['explicit_style_cells'],2)

    def test_multilevel_merged_headers_keep_sources_and_data(self):
        spec=specification({'A1':{'value':'部门'},'B1':{'value':'2025年'},'B2':{'value':'金额'},
                            'C2':{'value':'笔数'},'A3':{'value':'甲'},'B3':{'value':10},'C3':{'value':2}})
        book=Workbook.create(spec,appearance=appearance(header_rows='1:2',merges=['A1:A2','B1:C1'],
                                                        category_columns=[{'path':['部门']}]))
        sheet=xml(book.parts[book.sheets['Data']])
        self.assertEqual([item.get('ref') for item in sheet.find(NS+'mergeCells')],['A1:A2','B1:C1'])
        self.assertEqual(color(book,'A2'),HEADER_COLOR)
        self.assertEqual(color(book,'C1'),HEADER_COLOR)
        self.assertEqual(sheet.find(NS+'sheetViews/'+NS+'sheetView/'+NS+'pane').get('ySplit'),'2')
        self.assertEqual(book.cells['Data']['B3']['value'],10)

    def test_empty_sheet_and_explicit_blank_header_rectangle(self):
        book=Workbook.create(specification({}),appearance=appearance())
        self.assertEqual(book.appearance_report['sheets'][0]['status'],'empty')
        self.assertEqual(book.appearance_report['added_shared_styles'],0)
        book=Workbook.create(specification({}),appearance=appearance(range='A1:C1',header_rows='1'))
        self.assertEqual(len(book.cells['Data']),3)
        self.assertTrue(all(cell['value'] is None for cell in book.cells['Data'].values()))

    def test_shared_style_count_stays_constant_as_row_count_grows(self):
        counts=[]
        for size in (8,1000):
            cells={'A1':{'value':'类别'},'B1':{'value':'值'}}
            for row in range(2,size+2):
                cells[f'A{row}']={'value':'G'+str(row%3)}
                cells[f'B{row}']={'value':row}
            book=Workbook.create(specification(cells),appearance=appearance(category_columns=['A']))
            styles=xml(book.parts['xl/styles.xml'])
            counts.append(len(styles.find(NS+'cellXfs')))
            for name in ('cellXfs','fills','fonts','borders','numFmts'):
                self.assertEqual(int(styles.find(NS+name).get('count')),len(styles.find(NS+name)))
            self.assertTrue(all(int(cell['style'])<counts[-1] for cell in book.cells['Data'].values()))
        self.assertEqual(counts[0],counts[1])
        self.assertLessEqual(counts[1],17)

    def test_invalid_geometry_format_and_formula_categories_refused(self):
        spec=specification({'A1':{'value':'A'},'B1':{'value':'B'},'A2':{'value':1},'B2':{'value':2}})
        for settings in ({'merges':['A1:B1']},{'merges':['A1:A2']},{'column_widths':{'A':100}},
                         {'range':'A1:XFD1048576'},{'number_formats':{'Z':'0.00'}},
                         {'alternate_rows':'yes'},{'header_rows':'2:1'}):
            with self.subTest(settings=settings),self.assertRaises(ValueError):
                Workbook.create(spec,appearance=appearance(**settings))
        with self.assertRaisesRegex(ValueError,'literal keys'):
            Workbook.create(specification({'A1':{'value':'K'},'A2':{'formula':'1+1'}}),
                            appearance=appearance(category_columns=['A']))

    def test_editing_loaded_book_preserves_styles_and_disallows_preset(self):
        spec=specification({'A1':{'value':'V'},'A2':{'value':1}})
        book=Workbook.create(spec,appearance=appearance())
        original=self.root/'original.xlsx'; book.save(original)
        loaded=Workbook.load(original)
        before=loaded.parts['xl/styles.xml']
        with self.assertRaisesRegex(ValueError,'newly created'):
            apply_appearance(loaded,appearance(),spec)
        loaded.patch(specification({'A2':{'value':2}}))
        self.assertEqual(loaded.parts['xl/styles.xml'],before)
        self.assertEqual(loaded.cells['Data']['A2']['style'],book.cells['Data']['A2']['style'])

    def test_csv_import_header_count_numeric_types_and_headerless_preset(self):
        source=self.root/'source.csv'; source.write_text('ID,Amount\n001,19.90\n002,2\n')
        output=self.root/'import.xlsx'
        report=import_table(source,output,numeric_columns=['B'],appearance={'preset':'readable','sheets':{'Sheet1':{}}})
        self.assertEqual(report['appearance']['sheets'][0]['header_rows'],'1:1')
        book=Workbook.load(output)
        self.assertEqual(book.cells['Sheet1']['A2']['value'],'001')
        self.assertEqual(book.cells['Sheet1']['B2']['value'],Decimal('19.90'))
        source.write_text('001,2\n002,3\n')
        output=self.root/'no-header.xlsx'
        report=import_table(source,output,header_rows=0,appearance={'preset':'readable','sheets':{'Sheet1':{}}})
        self.assertIsNone(report['appearance']['sheets'][0]['header_rows'])
        self.assertIsNone(xml(Workbook.load(output).parts['xl/worksheets/sheet1.xml']).find(NS+'sheetViews'))

    def test_cli_creation_and_import_appearance_json(self):
        spec=self.root/'spec.json'; config=self.root/'appearance.json'
        spec.write_text(json.dumps(specification({'A1':{'value':'V'},'A2':{'formula':'1+2'}})))
        config.write_text(json.dumps(appearance()))
        scripts=Path(__file__).resolve().parent
        base=[sys.executable,'-I','-S','-B']
        output=self.root/'result.xlsx'
        run=subprocess.run(base+[str(scripts/'stdlib_xlsx.py'),'create',str(spec),str(output),'--appearance',str(config)],capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stderr+run.stdout)
        self.assertEqual(json.loads(run.stdout)['appearance']['preset'],'readable')
        self.assertEqual(Workbook.load(output).cells['Data']['A2']['value'],3)
        source=self.root/'source.csv'; source.write_text('K,V\n001,2\n')
        config.write_text(json.dumps({'preset':'readable','sheets':{'Sheet1':{}}}))
        run=subprocess.run(base+[str(scripts/'xlsx_tabular.py'),'import',str(source),str(self.root/'import.xlsx'),'--appearance',str(config)],capture_output=True,text=True)
        self.assertEqual(run.returncode,0,run.stderr+run.stdout)
        self.assertTrue(json.loads(run.stdout)['output_written'])


if __name__=='__main__':
    unittest.main()
