#!/usr/bin/env python3
# SPDX-License-Identifier: MIT
"""Opt-in readable styling for newly created workbooks; never a round-trip editor."""
from copy import deepcopy
from decimal import Decimal
import re
from xml.etree import ElementTree as ET

from stdlib_xlsx import MAX_CELLS, NS, address, clean_text, encoded_xml, xml
from xlsx_headers import header_bounds, reconstruct, resolve_column
from xlsx_analysis import typed_value
from xlsx_tabular import column_name

HEADER_COLOR = 'FFDCEAF7'
TEXT_COLOR = 'FF1F2937'
HEADER_TEXT_COLOR = 'FF17365D'
ROW_COLORS = ('FFFFFFFF', 'FFF4F8FC')
CATEGORY_COLORS = ('FFF1F8FD', 'FFF1F7ED', 'FFFFF6E8', 'FFF6EFFA')
MAX_STYLE_COUNT = 4096


class StyleRegistry:
    def __init__(self, root):
        self.root = root
        self.fonts = root.find(NS+'fonts')
        self.fills = root.find(NS+'fills')
        self.borders = root.find(NS+'borders')
        self.formats = root.find(NS+'numFmts')
        self.styles = root.find(NS+'cellXfs')
        self.cache = {}

    @staticmethod
    def append_unique(parent, node):
        encoded = ET.tostring(node)
        for index, existing in enumerate(parent):
            if ET.tostring(existing) == encoded:
                return index
        parent.append(node)
        parent.set('count', str(len(parent)))
        return len(parent)-1

    def font(self, header):
        node = ET.Element(NS+'font')
        if header:
            ET.SubElement(node, NS+'b')
        ET.SubElement(node, NS+'sz', {'val':'11'})
        ET.SubElement(node, NS+'name', {'val':'Calibri'})
        ET.SubElement(node, NS+'color', {'rgb':HEADER_TEXT_COLOR if header else TEXT_COLOR})
        return self.append_unique(self.fonts, node)

    def fill(self, color):
        node = ET.Element(NS+'fill')
        pattern = ET.SubElement(node, NS+'patternFill', {'patternType':'solid'})
        ET.SubElement(pattern, NS+'fgColor', {'rgb':color})
        ET.SubElement(pattern, NS+'bgColor', {'indexed':'64'})
        return self.append_unique(self.fills, node)

    def border(self):
        node = ET.Element(NS+'border')
        for side in ('left','right','top'):
            ET.SubElement(node, NS+side)
        bottom = ET.SubElement(node, NS+'bottom', {'style':'thin'})
        ET.SubElement(bottom, NS+'color', {'rgb':'FFDEE5ED'})
        ET.SubElement(node, NS+'diagonal')
        return self.append_unique(self.borders, node)

    def number_format(self, code):
        for node in self.formats:
            if node.get('formatCode') == code:
                return node.get('numFmtId')
        number = max([163]+[int(node.get('numFmtId')) for node in self.formats])+1
        ET.SubElement(self.formats, NS+'numFmt', {'numFmtId':str(number),'formatCode':code})
        self.formats.set('count', str(len(self.formats)))
        return str(number)

    def style(self, color, header, config, number_format=None):
        alignment = config.get('header_alignment','center') if header else config.get('body_alignment','general')
        wrap = config.get('wrap_headers',True) if header else config.get('wrap_text',False)
        key = (color,header,alignment,wrap,config.get('borders',True),number_format)
        if key in self.cache:
            return self.cache[key]
        if len(self.styles) >= MAX_STYLE_COUNT:
            raise ValueError('Appearance exceeds 4096 shared styles')
        node = deepcopy(self.styles[0])
        node.set('fontId',str(self.font(header)))
        node.set('fillId',str(self.fill(color)))
        node.set('applyFont','1')
        node.set('applyFill','1')
        node.set('borderId',str(self.border()) if config.get('borders',True) else '0')
        node.set('applyBorder','1')
        if number_format is not None and not header:
            node.set('numFmtId',self.number_format(number_format))
            node.set('applyNumberFormat','1')
        for child in list(node):
            if child.tag == NS+'alignment':
                node.remove(child)
        attrs = {'vertical':'center','wrapText':'1' if wrap else '0'}
        if alignment != 'general':
            attrs['horizontal'] = alignment
        ET.SubElement(node, NS+'alignment', attrs)
        node.set('applyAlignment','1')
        index = self.append_unique(self.styles,node)
        self.cache[key] = index
        return index


def validate_settings(settings):
    allowed = {'range','header_rows','merges','alternate_rows','category_columns','freeze_headers',
               'borders','wrap_headers','wrap_text','header_height','header_alignment','body_alignment','number_formats','column_widths'}
    if not isinstance(settings,dict) or set(settings)-allowed:
        raise ValueError('Unknown sheet appearance settings')
    if settings.get('header_rows') is not None:
        header_bounds(settings['header_rows'],1,1048576)
    height=settings.get('header_height',30)
    if isinstance(height,bool) or not isinstance(height,(int,float,Decimal)) or not Decimal(str(height)).is_finite() or not 12<=height<=80:
        raise ValueError('header_height must be finite in 12..80 points')
    for key in ('alternate_rows','freeze_headers','borders','wrap_headers','wrap_text'):
        if key in settings and type(settings[key]) is not bool:
            raise ValueError(key+' must be boolean')
    for key,default in (('header_alignment','center'),('body_alignment','general')):
        if settings.get(key,default) not in ('left','center','right','general'):
            raise ValueError('Invalid alignment')
    for key in ('number_formats','column_widths'):
        if not isinstance(settings.get(key,{}),dict):
            raise ValueError(key+' must be a column-letter mapping')
    for column,code in settings.get('number_formats',{}).items():
        if not re.fullmatch(r'[A-Z]{1,3}',column) or not isinstance(code,str) or not 1<=len(code)<=256:
            raise ValueError('Number formats require column letters and 1..256 character format codes')
        clean_text(code)
    for column,width in settings.get('column_widths',{}).items():
        if not re.fullmatch(r'[A-Z]{1,3}',column) or isinstance(width,bool) or not isinstance(width,(int,float,Decimal)) or not Decimal(str(width)).is_finite() or not 3<=width<=80:
            raise ValueError('Column widths must be finite numbers in 3..80')
    if not isinstance(settings.get('merges',[]),list) or len(settings.get('merges',[]))>10000:
        raise ValueError('merges must be a list with at most 10000 ranges')
    if not isinstance(settings.get('category_columns',[]),list):
        raise ValueError('category_columns must be a selector list')


def style_sheet(book, sheet, settings, explicit, registry):
    validate_settings(settings)
    cells = book.cells[sheet]
    spec = settings.get('range')
    if spec is not None:
        if not isinstance(spec,str) or len(spec.split(':')) not in (1,2):
            raise ValueError('Appearance range must be A1 or A1:D20')
        ends = spec.split(':')
        r1,c1 = address(ends[0]); r2,c2 = address(ends[-1])
    elif cells:
        positions = [address(ref) for ref in cells]
        r1,c1 = 1,1
        r2,c2 = max(row for row,_ in positions),max(col for _,col in positions)
    else:
        if settings.get('merges'):
            raise ValueError('Empty sheets with header merges need an explicit range')
        return {'sheet':sheet,'status':'empty','styled_cells':0,'explicit_style_cells':0}
    if r2<r1 or c2<c1 or (r2-r1+1)*(c2-c1+1)>MAX_CELLS:
        raise ValueError('Appearance range reversed or exceeds 200000 cells')
    columns = [column_name(col) for col in range(c1,c2+1)]
    if any(col not in columns for col in set(settings.get('number_formats',{}))|set(settings.get('column_widths',{}))):
        raise ValueError('Appearance format/width column lies outside selected range')
    header_spec = settings.get('header_rows',str(r1))
    schema = None
    header_first = header_last = 0
    if header_spec is not None:
        header_first,header_last = header_bounds(header_spec,r1,r2)
    part = book.sheets[sheet]
    tree = xml(book.parts[part])
    merges = tree.find(NS+'mergeCells')
    if settings.get('merges'):
        if merges is None:
            merges = ET.SubElement(tree,NS+'mergeCells')
        for ref in settings['merges']:
            if not isinstance(ref,str):
                raise ValueError('Merge ranges must be strings')
            ET.SubElement(merges,NS+'mergeCell',{'ref':ref.upper()})
        merges.set('count',str(len(merges)))
    book.parts[part] = encoded_xml(tree)
    if merges is not None:
        for merge in merges:
            ends = merge.get('ref','').split(':')
            if len(ends)!=2:
                raise ValueError('Header merges need two-cell ranges')
            top,left = address(ends[0]); bottom,right = address(ends[1])
            if not header_first or not header_first<=top<=bottom<=header_last or not c1<=left<=right<=c2:
                raise ValueError('Only complete merges inside declared header rows are supported')
    if header_spec is not None:
        schema = reconstruct(book,sheet,(r1,c1,r2,c2),header_spec,lambda ref:typed_value(cells.get(ref,{})))
    elif merges is not None and len(merges):
        raise ValueError('Header merges require explicit header rows')
    category_columns = [resolve_column(selector,columns,schema) for selector in settings.get('category_columns',[])]
    if len(category_columns)!=len(set(category_columns)):
        raise ValueError('Category columns must be unique')
    # New workbook XML is safe to serialize; there is no source round-trip here.
    data = tree.find(NS+'sheetData')
    row_nodes = {int(node.get('r')):node for node in data}
    category_colors = {}
    styled = skipped = 0
    for row in range(r1,r2+1):
        if header_first and row<header_first:
            continue
        header = bool(header_first and header_first<=row<=header_last)
        color = HEADER_COLOR if header else ROW_COLORS[(row-(header_last+1 if header_first else r1))%2] if settings.get('alternate_rows',True) else ROW_COLORS[0]
        if not header and category_columns:
            if any(cells.get(column+str(row),{}).get('formula') is not None for column in category_columns):
                raise ValueError('Category colors require literal keys; formula category at row '+str(row))
            key = tuple(typed_value(cells.get(column+str(row),{})) for column in category_columns)
            if key not in category_colors:
                category_colors[key] = len(category_colors)%len(CATEGORY_COLORS)
            color = CATEGORY_COLORS[category_colors[key]]
        node = row_nodes.get(row)
        if node is None:
            node = ET.SubElement(data,NS+'row',{'r':str(row)})
            row_nodes[row] = node
        if header:
            node.set('ht',str(settings.get('header_height',30)))
            node.set('customHeight','1')
        nodes = {cell.get('r'):cell for cell in node}
        for column in columns:
            ref = column+str(row)
            if ref in explicit:
                skipped += 1
                continue
            cell = nodes.get(ref)
            if cell is None:
                cell = ET.SubElement(node,NS+'c',{'r':ref})
            cell.set('s',str(registry.style(color,header,settings,settings.get('number_formats',{}).get(column))))
            styled += 1
        node[:] = sorted(node,key=lambda cell:address(cell.get('r'))[1])
    data[:] = sorted(data,key=lambda node:int(node.get('r')))
    if settings.get('freeze_headers',True) and header_first:
        if header_last>=1048576:
            raise ValueError('Cannot freeze beyond the final Excel row')
        views = tree.find(NS+'sheetViews')
        if views is None:
            views = ET.Element(NS+'sheetViews')
            tree.insert(0,views)
        view = views.find(NS+'sheetView')
        if view is None:
            view = ET.SubElement(views,NS+'sheetView',{'workbookViewId':'0'})
        for pane in list(view.findall(NS+'pane')):
            view.remove(pane)
        view.insert(0,ET.Element(NS+'pane',{'ySplit':str(header_last),'topLeftCell':'A'+str(header_last+1),'activePane':'bottomLeft','state':'frozen'}))
    widths = settings.get('column_widths',{})
    if widths:
        existing = tree.find(NS+'cols')
        if existing is not None:
            tree.remove(existing)
        cols = ET.Element(NS+'cols')
        tree.insert(list(tree).index(data),cols)
        for column,width in sorted(widths.items(),key=lambda item:address(item[0]+'1')[1]):
            index = str(address(column+'1')[1])
            ET.SubElement(cols,NS+'col',{'min':index,'max':index,'width':str(width),'customWidth':'1'})
    book.parts[part] = encoded_xml(tree)
    return {'sheet':sheet,'status':'styled','styled_cells':styled,'explicit_style_cells':skipped,
            'category_key_count':len(category_colors),'category_palette_size':len(CATEGORY_COLORS),
            'header_rows':header_spec,'data_order_changed':False,
            'frozen_through_row':header_last if settings.get('freeze_headers',True) and header_first else None}


def apply_appearance(book, appearance, spec):
    if not getattr(book,'_created_here',False):
        raise ValueError('Appearance presets are restricted to newly created workbooks')
    if not isinstance(appearance,dict) or set(appearance)!={'preset','sheets'} or appearance.get('preset')!='readable' or not isinstance(appearance.get('sheets'),dict):
        raise ValueError('Appearance requires preset: readable and a sheets settings mapping')
    if any(sheet not in book.sheets for sheet in appearance['sheets']):
        raise ValueError('Appearance names an unknown sheet')
    book.flush()
    book.documents.clear()
    explicit = {sheet['name']:{ref.upper() for ref,entry in sheet['cells'].items() if 'style' in entry} for sheet in spec['sheets']}
    styles = xml(book.parts['xl/styles.xml'])
    registry = StyleRegistry(styles)
    before = len(registry.styles)
    reports = [style_sheet(book,sheet,settings,explicit[sheet],registry) for sheet,settings in appearance['sheets'].items()]
    book.parts['xl/styles.xml'] = encoded_xml(styles)
    if any(report.get('frozen_through_row') for report in reports):
        workbook = xml(book.parts['xl/workbook.xml'])
        if workbook.find(NS+'bookViews') is None:
            views = ET.Element(NS+'bookViews')
            ET.SubElement(views,NS+'workbookView',{'activeTab':'0'})
            workbook.insert(list(workbook).index(workbook.find(NS+'sheets')),views)
            book.parts['xl/workbook.xml'] = encoded_xml(workbook)
    book._index()
    return {'preset':'readable','sheets':reports,'added_shared_styles':len(registry.styles)-before,
            'priority':'explicit cell style > header > category > alternating rows',
            'category_rule':'first-seen typed key gets palette color; noncontiguous equal keys share color; palette cycles',
            'visual_application_verified':False}
