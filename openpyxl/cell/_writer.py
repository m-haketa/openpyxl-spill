# Copyright (c) 2010-2024 openpyxl

from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, SubElement, whitespace, XML_NS
from openpyxl import LXML
from openpyxl.utils.datetime import to_excel, to_ISO8601
from datetime import timedelta

from openpyxl.worksheet.formula import DataTableFormula, ArrayFormula
from openpyxl.cell.rich_text import CellRichText
from .formula_utils import add_function_prefix as _add_function_prefix, prepare_spill_formula as _prepare_spill_formula, EXCEL_NEW_FUNCTIONS


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
        # LAMBDA/LET関数も配列式として処理
        elif isinstance(value, str) and ('LAMBDA' in value or 'LET' in value):
            # プレフィックスを追加
            value = _add_function_prefix(value)
            # =を削除
            if value.startswith('='):
                value = value[1:]
            # 配列式の属性を設定
            attrib = {
                't': 'array',
                'ref': cell.coordinate
            }

        formula = SubElement(el, 'f', attrib)
        if value is not None and not attrib.get('t') == "dataTable":
            # LAMBDA/LET関数は既に処理済み
            is_lambda_or_let = attrib.get('t') == 'array' and original_value and isinstance(original_value, str) and ('LAMBDA' in original_value or 'LET' in original_value)
            
            # スピル数式とLAMBDA/LET関数は既に処理済み、通常の数式は=を削除
            if not getattr(cell, "_is_spill", False) and not is_lambda_or_let:
                # 通常の数式でも新関数にプレフィックスを追加
                value = _add_function_prefix(value)
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
            # LAMBDA/LET関数も配列式として処理
            elif isinstance(value, str) and ('LAMBDA' in value or 'LET' in value):
                # プレフィックスを追加
                value = _add_function_prefix(value)
                # =を削除
                if value.startswith('='):
                    value = value[1:]
                # 配列式の属性を設定
                attrib = {
                    't': 'array',
                    'ref': cell.coordinate
                }

            with xf.element('f', attrib):
                if value is not None and not attrib.get('t') == "dataTable":
                    # LAMBDA/LET関数は既に処理済み
                    is_lambda_or_let = attrib.get('t') == 'array' and original_value and isinstance(original_value, str) and ('LAMBDA' in original_value or 'LET' in original_value)
                    
                    # スピル数式とLAMBDA/LET関数は既に処理済み、通常の数式は=を削除
                    if not getattr(cell, "_is_spill", False) and not is_lambda_or_let:
                        # 通常の数式でも新関数にプレフィックスを追加
                        value = _add_function_prefix(value)
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
