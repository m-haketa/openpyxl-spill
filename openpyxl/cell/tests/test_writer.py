# Copyright (c) 2010-2024 openpyxl

import datetime
import decimal
from io import BytesIO

import pytest

from openpyxl.xml.functions import xmlfile

from openpyxl.tests.helper import compare_xml
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900

from openpyxl import LXML

@pytest.fixture
def worksheet():
    from openpyxl import Workbook
    wb = Workbook()
    return wb.active


@pytest.fixture
def etree_write_cell():
    from .._writer import etree_write_cell
    return etree_write_cell


@pytest.fixture
def lxml_write_cell():
    from .._writer import lxml_write_cell
    return lxml_write_cell


@pytest.fixture(params=['etree', 'lxml'])
def write_cell_implementation(request, etree_write_cell, lxml_write_cell):
    if request.param == "lxml" and LXML:
        return lxml_write_cell
    return etree_write_cell


@pytest.mark.parametrize("value, expected",
                         [
                             (9781231231230, """<c t="n" r="A1"><v>9781231231230</v></c>"""),
                             (decimal.Decimal('3.14'), """<c t="n" r="A1"><v>3.14</v></c>"""),
                             (1234567890, """<c t="n" r="A1"><v>1234567890</v></c>"""),
                             ("=sum(1+1)", """<c r="A1"><f>sum(1+1)</f><v></v></c>"""),
                             (True, """<c t="b" r="A1"><v>1</v></c>"""),
                             ("Hello", """<c t="inlineStr" r="A1"><is><t>Hello</t></is></c>"""),
                             ("", """<c r="A1" t="inlineStr"></c>"""),
                             (None, """<c r="A1" t="n"></c>"""),
                         ])
def test_write_cell(worksheet, write_cell_implementation, value, expected):
    write_cell = write_cell_implementation

    ws = worksheet
    cell = ws['A1']
    cell.value = value

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell, cell.has_style)

    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.parametrize("value, iso_dates, expected,",
                         [
                             (datetime.date(2011, 12, 25), False, """<c r="A1" t="n" s="1"><v>40902</v></c>"""),
                             (datetime.date(2011, 12, 25), True, """<c r="A1" t="d" s="1"><v>2011-12-25</v></c>"""),
                             (datetime.datetime(2011, 12, 25, 14, 23, 55), False, """<c r="A1" t="n" s="1"><v>40902.59994212963</v></c>"""),
                             (datetime.datetime(2011, 12, 25, 14, 23, 55), True, """<c r="A1" t="d" s="1"><v>2011-12-25T14:23:55</v></c>"""),
                             (datetime.time(14, 15, 25), False, """<c r="A1" t="n" s="1"><v>0.5940393518518519</v></c>"""),
                             (datetime.time(14, 15, 25), True, """<c r="A1" t="d" s="1"><v>14:15:25</v></c>"""),
                             (datetime.timedelta(1, 3, 15), False, """<c r="A1" t="n" s="1"><v>1.000034722395833</v></c>"""),
                             (datetime.timedelta(1, 3, 15), True, """<c r="A1" t="n" s="1"><v>1.000034722395833</v></c>"""),
                         ]
                         )
def test_write_date(worksheet, write_cell_implementation, value, expected, iso_dates):
    write_cell = write_cell_implementation

    ws = worksheet
    cell = ws['A1']
    cell.value = value
    cell.parent.parent.iso_dates = iso_dates

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell, cell.has_style)

    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.mark.parametrize("value, iso_dates",
                         [
                             (datetime.datetime(2021, 3, 19, 23, tzinfo=datetime.timezone.utc), True),
                             (datetime.datetime(2021, 3, 19, 23, tzinfo=datetime.timezone.utc), False),
                             (datetime.time(23, 58, tzinfo=datetime.timezone.utc), True),
                             (datetime.time(23, 58, tzinfo=datetime.timezone.utc), False),
                         ]
                         )
def test_write_invalid_date(worksheet, write_cell_implementation, value, iso_dates):
    write_cell = write_cell_implementation

    ws = worksheet
    cell = ws['A1']
    cell.value = value
    cell.parent.parent.iso_dates = iso_dates

    out = BytesIO()
    with pytest.raises(TypeError):
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell, cell.has_style)


@pytest.mark.parametrize("value, expected, epoch",
                         [
                             (datetime.date(2011, 12, 25), """<c r="A1" t="n" s="1"><v>40902</v></c>""",
                              CALENDAR_WINDOWS_1900),
                             (datetime.date(2011, 12, 25), """<c r="A1" t="n" s="1"><v>39440</v></c>""",
                              CALENDAR_MAC_1904),
                         ]
                         )
def test_write_epoch(worksheet, write_cell_implementation, value, expected, epoch):
    write_cell = write_cell_implementation

    ws = worksheet
    ws.parent.epoch = epoch
    cell = ws['A1']
    cell.value = value

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell, cell.has_style)

    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_hyperlink(worksheet, write_cell_implementation):
    write_cell = write_cell_implementation

    ws = worksheet
    cell = ws['A1']
    cell.value = "test"
    cell.hyperlink = "http://www.test.com"

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell, cell.has_style)

    assert len(worksheet._hyperlinks) == 1


@pytest.mark.parametrize("value, result, attrs",
                         [
                             ("test", "test", {'r': 'A1', 't': 'inlineStr'}),
                             ("=SUM(A1:A2)", "=SUM(A1:A2)", {'r': 'A1'}),
                             (datetime.date(2018, 8, 25), 43337, {'r':'A1', 't':'n'}),
                         ]
                         )
def test_attributes(worksheet, value, result, attrs):
    from .._writer import _set_attributes

    ws = worksheet
    cell = ws['A1']
    cell.value = value

    assert(_set_attributes(cell)) == (result, attrs)


def test_whitespace(worksheet, write_cell_implementation):
    write_cell = write_cell_implementation
    ws = worksheet
    cell = ws['A1']
    cell.value = "  whitespace   "

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)

    expected = """
    <c t="inlineStr" r="A1">
      <is>
        <t xml:space="preserve">  whitespace   </t>
      </is>
    </c>"""
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


from openpyxl.worksheet.formula import DataTableFormula, ArrayFormula

def test_table_formula(worksheet, write_cell_implementation):
    write_cell = write_cell_implementation
    ws = worksheet
    cell = ws["A1"]
    cell.value =  DataTableFormula(ref="A1:B10")
    cell.data_type = "f"

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)

    expected = """
    <c r="A1">
      <f t="dataTable" ref="A1:B10" />
      <v>0</v>
    </c>"""
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_array_formula(worksheet, write_cell_implementation):
    write_cell = write_cell_implementation
    ws = worksheet

    cell = ws["E2"]
    cell.value = ArrayFormula(ref="E2:E11", text="=C2:C11*D2:D11")

    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)

    expected = """
    <c r="E2">
      <f t="array" ref="E2:E11">C2:C11*D2:D11</f>
      <v>0</v>
    </c>"""
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_rich_text(worksheet, write_cell_implementation):
    write_cell = write_cell_implementation
    ws = worksheet

    from ..rich_text import CellRichText, TextBlock, InlineFont

    red = InlineFont(color='FF0000')
    rich_string = CellRichText(
        [TextBlock(red, 'red'),
         ' is used, you can expect ',
         TextBlock(red, 'danger')]
    )
    cell = ws["A2"]
    cell.value = rich_string


    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)

    expected = """
    <c r="A2" t="inlineStr">
      <is>
        <r>
        <rPr>
          <color rgb="00FF0000" />
        </rPr>
        <t>red</t>
        </r>
        <r>
          <t xml:space="preserve"> is used, you can expect </t>
        </r>
        <r>
          <rPr>
            <color rgb="00FF0000" />
          </rPr>
          <t>danger</t>
        </r>
      </is>
    </c>"""
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_nested_spill_formulas(worksheet, write_cell_implementation):
    """Test nested spill formulas with proper _xlfn prefix handling"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # Test 1: SORT(UNIQUE(...))
    cell = ws["F2"]
    cell.value = '=SORT(UNIQUE(B2:B8),1,-1)'
    cell._is_spill = True
    cell._spill_range = 'F2:F6'
    
    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)
    
    expected = """
    <c r="F2" cm="1">
      <f t="array" ref="F2:F6">_xlfn._xlws.SORT(_xlfn.UNIQUE(B2:B8),1,-1)</f>
      <v>0</v>
    </c>"""
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_deeply_nested_spill_formulas(worksheet, write_cell_implementation):
    """Test deeply nested spill formulas"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # Test: FILTER(SORT(...),condition)
    cell = ws["F8"]
    cell.value = '=FILTER(SORT(B2:B8,1,-1),SORT(B2:B8,1,-1)>=2000)'
    cell._is_spill = True
    cell._spill_range = 'F8:F12'
    
    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)
    
    expected = """
    <c r="F8" cm="1">
      <f t="array" ref="F8:F12">_xlfn._xlws.FILTER(_xlfn._xlws.SORT(B2:B8,1,-1),_xlfn._xlws.SORT(B2:B8,1,-1)&gt;=2000)</f>
      <v>0</v>
    </c>"""
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_mixed_spill_formulas(worksheet, write_cell_implementation):
    """Test mixed spill formulas with regular functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # Test: UNIQUE inside SUM (SUM is not a spill function)
    cell = ws["G2"]
    cell.value = '=SUM(UNIQUE(B2:B8))'
    cell._is_spill = False  # SUM doesn't spill
    
    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)
    
    # _is_spill is False but UNIQUE still gets _xlfn prefix
    expected = """
    <c r="G2">
      <f>SUM(_xlfn.UNIQUE(B2:B8))</f>
      <v/>
    </c>"""
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_phase1_array_functions(worksheet, write_cell_implementation):
    """Test Phase 1 array manipulation functions with proper _xlfn prefix"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # テストケース: (セル, 数式, スピル範囲, 期待されるXML)
    test_cases = [
        # VSTACK
        ("A1", '=VSTACK(A2:B3,A5:B6)', 'A1:B4', """
    <c r="A1" cm="1">
      <f t="array" ref="A1:B4">_xlfn.VSTACK(A2:B3,A5:B6)</f>
      <v>0</v>
    </c>"""),
        
        # HSTACK
        ("B1", '=HSTACK(A1:A3,B1:B3)', 'B1:C3', """
    <c r="B1" cm="1">
      <f t="array" ref="B1:C3">_xlfn.HSTACK(A1:A3,B1:B3)</f>
      <v>0</v>
    </c>"""),
        
        # TAKE
        ("C1", '=TAKE(A1:C5,3)', 'C1:E3', """
    <c r="C1" cm="1">
      <f t="array" ref="C1:E3">_xlfn.TAKE(A1:C5,3)</f>
      <v>0</v>
    </c>"""),
        
        # DROP
        ("D1", '=DROP(A1:C5,1)', 'D1:F4', """
    <c r="D1" cm="1">
      <f t="array" ref="D1:F4">_xlfn.DROP(A1:C5,1)</f>
      <v>0</v>
    </c>"""),
        
        # CHOOSEROWS
        ("E1", '=CHOOSEROWS(A1:C5,1,3)', 'E1:G2', """
    <c r="E1" cm="1">
      <f t="array" ref="E1:G2">_xlfn.CHOOSEROWS(A1:C5,1,3)</f>
      <v>0</v>
    </c>"""),
        
        # CHOOSECOLS
        ("F1", '=CHOOSECOLS(A1:C5,1,3)', 'F1:G5', """
    <c r="F1" cm="1">
      <f t="array" ref="F1:G5">_xlfn.CHOOSECOLS(A1:C5,1,3)</f>
      <v>0</v>
    </c>"""),
        
        # EXPAND
        ("G1", '=EXPAND(A1:B3,5,4,"N/A")', 'G1:J5', """
    <c r="G1" cm="1">
      <f t="array" ref="G1:J5">_xlfn.EXPAND(A1:B3,5,4,"N/A")</f>
      <v>0</v>
    </c>"""),
        
        # TOCOL
        ("H1", '=TOCOL(A1:C3)', 'H1:H9', """
    <c r="H1" cm="1">
      <f t="array" ref="H1:H9">_xlfn.TOCOL(A1:C3)</f>
      <v>0</v>
    </c>"""),
        
        # TOROW
        ("I1", '=TOROW(A1:B4)', 'I1:P1', """
    <c r="I1" cm="1">
      <f t="array" ref="I1:P1">_xlfn.TOROW(A1:B4)</f>
      <v>0</v>
    </c>"""),
        
        # WRAPCOLS
        ("J1", '=WRAPCOLS(SEQUENCE(10),3)', 'J1:L4', """
    <c r="J1" cm="1">
      <f t="array" ref="J1:L4">_xlfn.WRAPCOLS(_xlfn.SEQUENCE(10),3)</f>
      <v>0</v>
    </c>"""),
        
        # WRAPROWS
        ("K1", '=WRAPROWS(SEQUENCE(10),3)', 'K1:N3', """
    <c r="K1" cm="1">
      <f t="array" ref="K1:N3">_xlfn.WRAPROWS(_xlfn.SEQUENCE(10),3)</f>
      <v>0</v>
    </c>"""),
    ]
    
    # 各テストケースを実行
    for cell_ref, formula, spill_range, expected in test_cases:
        cell = ws[cell_ref]
        cell.value = formula
        cell._is_spill = True
        cell._spill_range = spill_range
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, f"Failed for {formula}: {diff}"


def test_phase2_text_functions(worksheet, write_cell_implementation):
    """Test Phase 2 text processing functions with proper _xlfn prefix"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # テストケース: (セル, 数式, スピル範囲, 期待されるXML)
    test_cases = [
        # ARRAYTOTEXT
        ("A1", '=ARRAYTOTEXT(A2:B6)', 'A1', """
    <c r="A1" cm="1">
      <f t="array" ref="A1">_xlfn.ARRAYTOTEXT(A2:B6)</f>
      <v>0</v>
    </c>"""),
        
        # VALUETOTEXT
        ("B1", '=VALUETOTEXT(D2)', 'B1', """
    <c r="B1" cm="1">
      <f t="array" ref="B1">_xlfn.VALUETOTEXT(D2)</f>
      <v>0</v>
    </c>"""),
        
        # TEXTAFTER
        ("C1", '=TEXTAFTER(B2:B6,"@")', 'C1:C5', """
    <c r="C1" cm="1">
      <f t="array" ref="C1:C5">_xlfn.TEXTAFTER(B2:B6,"@")</f>
      <v>0</v>
    </c>"""),
        
        # TEXTBEFORE
        ("D1", '=TEXTBEFORE(B2:B6,"@")', 'D1:D5', """
    <c r="D1" cm="1">
      <f t="array" ref="D1:D5">_xlfn.TEXTBEFORE(B2:B6,"@")</f>
      <v>0</v>
    </c>"""),
        
        # TEXTSPLIT
        ("E1", '=TEXTSPLIT(C2,"-")', 'E1:G1', """
    <c r="E1" cm="1">
      <f t="array" ref="E1:G1">_xlfn.TEXTSPLIT(C2,"-")</f>
      <v>0</v>
    </c>"""),
        
        # REGEXEXTRACT
        ("F1", '=REGEXEXTRACT(C2:C6,"\d+")', 'F1:F5', """
    <c r="F1" cm="1">
      <f t="array" ref="F1:F5">_xlfn.REGEXEXTRACT(C2:C6,"\d+")</f>
      <v>0</v>
    </c>"""),
        
        # REGEXREPLACE
        ("G1", '=REGEXREPLACE(B2:B6,"@.*","@company.com")', 'G1:G5', """
    <c r="G1" cm="1">
      <f t="array" ref="G1:G5">_xlfn.REGEXREPLACE(B2:B6,"@.*","@company.com")</f>
      <v>0</v>
    </c>"""),
        
        # REGEXTEST
        ("H1", '=REGEXTEST(B2:B6,"\.com$")', 'H1:H5', """
    <c r="H1" cm="1">
      <f t="array" ref="H1:H5">_xlfn.REGEXTEST(B2:B6,"\.com$")</f>
      <v>0</v>
    </c>"""),
    ]
    
    # 各テストケースを実行
    for cell_ref, formula, spill_range, expected in test_cases:
        cell = ws[cell_ref]
        cell.value = formula
        cell._is_spill = True
        cell._spill_range = spill_range
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, f"Failed for {formula}: {diff}"


def test_new_functions_without_spill_array(worksheet, write_cell_implementation):
    """Test new functions in normal formulas (not spilling) also get _xlfn prefix"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # テストケース: (セル, 数式, 期待されるXML) - スピルなし
    test_cases = [
        # Phase 1 functions
        ("A1", '=VSTACK(A2:A3,B2:B3)', """
    <c r="A1">
      <f>_xlfn.VSTACK(A2:A3,B2:B3)</f>
      <v/>
    </c>"""),
        
        ("B1", '=HSTACK(A1,B1)', """
    <c r="B1">
      <f>_xlfn.HSTACK(A1,B1)</f>
      <v/>
    </c>"""),
        
        ("C1", '=TAKE(A1:C5,1)', """
    <c r="C1">
      <f>_xlfn.TAKE(A1:C5,1)</f>
      <v/>
    </c>"""),
        
        # Phase 2 functions
        ("D1", '=ARRAYTOTEXT(A1:B2)', """
    <c r="D1">
      <f>_xlfn.ARRAYTOTEXT(A1:B2)</f>
      <v/>
    </c>"""),
        
        ("E1", '=TEXTBEFORE(A1,"@")', """
    <c r="E1">
      <f>_xlfn.TEXTBEFORE(A1,"@")</f>
      <v/>
    </c>"""),
        
        ("F1", '=REGEXTEST(A1,"test")', """
    <c r="F1">
      <f>_xlfn.REGEXTEST(A1,"test")</f>
      <v/>
    </c>"""),
        
        # 既存のスピル関数
        ("G1", '=UNIQUE(B1:B10)', """
    <c r="G1">
      <f>_xlfn.UNIQUE(B1:B10)</f>
      <v/>
    </c>"""),
        
        ("H1", '=SORT(A1:A10)', """
    <c r="H1">
      <f>_xlfn._xlws.SORT(A1:A10)</f>
      <v/>
    </c>"""),
    ]
    
    # 各テストケースを実行
    for cell_ref, formula, expected in test_cases:
        cell = ws[cell_ref]
        cell.value = formula
        # _is_spillは設定しない（通常の数式）
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, f"Failed for {formula}: {diff}"

