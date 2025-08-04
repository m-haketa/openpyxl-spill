# Copyright (c) 2010-2024 openpyxl

from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, SubElement, whitespace, XML_NS
from openpyxl import LXML
from openpyxl.utils.datetime import to_excel, to_ISO8601
from datetime import timedelta

from openpyxl.worksheet.formula import DataTableFormula, ArrayFormula
from openpyxl.cell.rich_text import CellRichText

# Excel 365の新しいスピル関数リスト
SPILL_FUNCTIONS = {
    'UNIQUE', 'SORT', 'SORTBY', 'FILTER', 'SEQUENCE', 
    'RANDARRAY', 'XLOOKUP', 'XMATCH'
}

def _prepare_spill_formula(formula_text, cell):
    """
    スピル数式を適切な形式に変換する
    
    Args:
        formula_text: 元の数式テキスト（"=UNIQUE(B2:B10)"）
        cell: Cellオブジェクト
    
    Returns:
        tuple: (処理済み数式, 属性辞書)
    """
    if not getattr(cell, "_is_spill", False):
        return formula_text, {}
    
    # =を削除
    if formula_text and formula_text.startswith('='):
        formula_text = formula_text[1:]
    
    # 関数名を抽出して_xlfn.プレフィックスを追加
    for func in SPILL_FUNCTIONS:
        if formula_text.upper().startswith(func):
            formula_text = '_xlfn.' + formula_text
            break
    # SORT関数とFILTER関数の特殊ケース（_xlwsプレフィックスが必要）
    if formula_text.startswith('_xlfn.SORT'):
        formula_text = formula_text.replace('_xlfn.SORT', '_xlfn._xlws.SORT')
    elif formula_text.startswith('_xlfn.FILTER'):
        formula_text = formula_text.replace('_xlfn.FILTER', '_xlfn._xlws.FILTER')
    
    # 属性を設定
    attrib = {
        't': 'array',
        'ref': getattr(cell, '_spill_range', None) or cell.coordinate
    }
    
    return formula_text, attrib

def _set_attributes(cell, styled=None):
    """
    Set coordinate and datatype
    """
    coordinate = cell.coordinate
    attrs = {'r': coordinate}
    if styled:
        attrs['s'] = f"{cell.style_id}"

    if cell.data_type == "s":
        attrs['t'] = "inlineStr"
    elif cell.data_type != 'f':
        attrs['t'] = cell.data_type

    # スピル数式の場合はcm属性を追加
    if cell.data_type == "f" and getattr(cell, "_is_spill", False):
        attrs['cm'] = "1"

    value = cell._value

    if cell.data_type == "d":
        if hasattr(value, "tzinfo") and value.tzinfo is not None:
            raise TypeError("Excel does not support timezones in datetimes. "
                    "The tzinfo in the datetime/time object must be set to None.")

        if cell.parent.parent.iso_dates and not isinstance(value, timedelta):
            value = to_ISO8601(value)
        else:
            attrs['t'] = "n"
            value = to_excel(value, cell.parent.parent.epoch)

    if cell.hyperlink:
        cell.parent._hyperlinks.append(cell.hyperlink)

    return value, attrs


def etree_write_cell(xf, worksheet, cell, styled=None):

    value, attributes = _set_attributes(cell, styled)

    el = Element("c", attributes)
    if value is None or value == "":
        xf.write(el)
        return

    if cell.data_type == 'f':
        attrib = {}
        
        # スピル数式の処理
        original_value = value
        if getattr(cell, "_is_spill", False):
            value, spill_attrib = _prepare_spill_formula(value, cell)
            attrib.update(spill_attrib)
        elif isinstance(value, ArrayFormula):
            attrib = dict(value)
            value = value.text
        elif isinstance(value, DataTableFormula):
            attrib = dict(value)
            value = None

        formula = SubElement(el, 'f', attrib)
        if value is not None and not attrib.get('t') == "dataTable":
            # スピル数式は既に処理済み、通常の数式は=を削除
            if not getattr(cell, "_is_spill", False):
                formula.text = value[1:] if value.startswith('=') else value
            else:
                formula.text = value
            # スピル数式の場合、v要素に初期値を設定（配列の最初の要素のインデックス）
            if getattr(cell, "_is_spill", False):
                value = "0"  # Excelはスピル数式の初期値として0または計算結果を期待
            else:
                value = None

    if cell.data_type == 's':
        if isinstance(value, CellRichText):
            el.append(value.to_tree())
        else:
            inline_string = Element("is")
            text = Element('t')
            text.text = value
            whitespace(text)
            inline_string.append(text)
            el.append(inline_string)

    else:
        cell_content = SubElement(el, 'v')
        if value is not None:
            cell_content.text = safe_string(value)

    xf.write(el)


def lxml_write_cell(xf, worksheet, cell, styled=False):
    value, attributes = _set_attributes(cell, styled)

    if value == '' or value is None:
        with xf.element("c", attributes):
            return

    with xf.element('c', attributes):
        if cell.data_type == 'f':
            attrib = {}
            
            # スピル数式の処理
            original_value = value
            if getattr(cell, "_is_spill", False):
                value, spill_attrib = _prepare_spill_formula(value, cell)
                attrib.update(spill_attrib)
            elif isinstance(value, ArrayFormula):
                attrib = dict(value)
                value = value.text
            elif isinstance(value, DataTableFormula):
                attrib = dict(value)
                value = None

            with xf.element('f', attrib):
                if value is not None and not attrib.get('t') == "dataTable":
                    # スピル数式は既に処理済み、通常の数式は=を削除
                    if not getattr(cell, "_is_spill", False):
                        xf.write(value[1:] if value.startswith('=') else value)
                    else:
                        xf.write(value)
                    # スピル数式の場合、v要素に初期値を設定
                    if getattr(cell, "_is_spill", False):
                        value = "0"  # Excelはスピル数式の初期値として0または計算結果を期待
                    else:
                        value = None

        if cell.data_type == 's':
            if isinstance(value, CellRichText):
                el = value.to_tree()
                xf.write(el)
            else:
                with xf.element("is"):
                    if isinstance(value, str):
                        attrs = {}
                        if value != value.strip():
                            attrs["{%s}space" % XML_NS] = "preserve"
                        el = Element("t", attrs) # lxml can't handle xml-ns
                        el.text = value
                        xf.write(el)

        else:
            with xf.element("v"):
                if value is not None:
                    xf.write(safe_string(value))


if LXML:
    write_cell = lxml_write_cell
else:
    write_cell = etree_write_cell
