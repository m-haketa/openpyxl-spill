"""
Test TRIMRANGE function and its integration with formula prefix handling
"""

import pytest
from io import BytesIO
from openpyxl.xml.functions import xmlfile
from openpyxl.cell._writer import (
    etree_write_cell,
    lxml_write_cell
)
from openpyxl.cell.formula_prefix import add_function_prefix
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def worksheet():
    from openpyxl import Workbook
    wb = Workbook()
    return wb.active


@pytest.fixture(params=['lxml', 'etree'])
def write_cell_implementation(request):
    if request.param == 'lxml':
        return lxml_write_cell
    else:
        return etree_write_cell


@pytest.mark.parametrize("formula,expected", [
    # Basic TRIMRANGE usage
    ("=TRIMRANGE(A1:B10)", "=_xlfn.TRIMRANGE(A1:B10)"),
    ("=TRIMRANGE(A1:B10,3,3)", "=_xlfn.TRIMRANGE(A1:B10,3,3)"),
    ("=TRIMRANGE(A1:B10,1,0)", "=_xlfn.TRIMRANGE(A1:B10,1,0)"),
    ("=TRIMRANGE(A1:B10,0,1)", "=_xlfn.TRIMRANGE(A1:B10,0,1)"),
    # TRIMRANGE with sheet reference
    ("=TRIMRANGE(Sheet1!A1:B10)", "=_xlfn.TRIMRANGE(Sheet1!A1:B10)"),
    ("=TRIMRANGE('My Sheet'!A1:B10,2,2)", "=_xlfn.TRIMRANGE('My Sheet'!A1:B10,2,2)"),
    # TRIMRANGE in functions
    ("=SUM(TRIMRANGE(A1:B10))", "=SUM(_xlfn.TRIMRANGE(A1:B10))"),
    ("=AVERAGE(TRIMRANGE(A1:B10,3,3))", "=AVERAGE(_xlfn.TRIMRANGE(A1:B10,3,3))"),
    ("=COUNTA(TRIMRANGE(A1:D10))", "=COUNTA(_xlfn.TRIMRANGE(A1:D10))"),
    # TRIMRANGE with other array functions
    ("=SORT(TRIMRANGE(A1:B10))", "=_xlfn._xlws.SORT(_xlfn.TRIMRANGE(A1:B10))"),
    ("=UNIQUE(TRIMRANGE(A1:B10,2,2))", "=_xlfn.UNIQUE(_xlfn.TRIMRANGE(A1:B10,2,2))"),
    # Complex formulas
    ("=SUMPRODUCT(TRIMRANGE(A1:B10)*{1;2;3})", "=SUMPRODUCT(_xlfn.TRIMRANGE(A1:B10)*{1;2;3})"),
    ("=MAX(TRIMRANGE(A1:B10))-MIN(TRIMRANGE(A1:B10))", 
     "=MAX(_xlfn.TRIMRANGE(A1:B10))-MIN(_xlfn.TRIMRANGE(A1:B10))"),
    # Multiple TRIMRANGE calls
    ("=VSTACK(TRIMRANGE(A1:B5),TRIMRANGE(C1:D5))", 
     "=_xlfn.VSTACK(_xlfn.TRIMRANGE(A1:B5),_xlfn.TRIMRANGE(C1:D5))"),
    ("=HSTACK(TRIMRANGE(A1:B5,1,1),TRIMRANGE(C1:D5,2,2))", 
     "=_xlfn.HSTACK(_xlfn.TRIMRANGE(A1:B5,1,1),_xlfn.TRIMRANGE(C1:D5,2,2))"),
])
def test_trimrange_prefix(formula, expected):
    assert add_function_prefix(formula) == expected


def test_trimrange_basic(worksheet, write_cell_implementation):
    """Test basic TRIMRANGE function in cell writing"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # Basic TRIMRANGE
        ("A1", '=TRIMRANGE(J5:L10)', """
    <c r="A1">
      <f>_xlfn.TRIMRANGE(J5:L10)</f>
      <v/>
    </c>"""),
        
        # TRIMRANGE with parameters
        ("A2", '=TRIMRANGE(J5:L10,3,3)', """
    <c r="A2">
      <f>_xlfn.TRIMRANGE(J5:L10,3,3)</f>
      <v/>
    </c>"""),
        
        # TRIMRANGE with row/column trimming
        ("A3", '=TRIMRANGE(J5:L10,1,0)', """
    <c r="A3">
      <f>_xlfn.TRIMRANGE(J5:L10,1,0)</f>
      <v/>
    </c>"""),
    ]
    
    for cell_ref, formula, expected in test_cases:
        cell = ws[cell_ref]
        cell.value = formula
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, f"Failed for {formula}: {diff}"


def test_trimrange_in_functions(worksheet, write_cell_implementation):
    """Test TRIMRANGE inside other functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # In SUM
        ("B1", '=SUM(TRIMRANGE(P:Q))', """
    <c r="B1">
      <f>SUM(_xlfn.TRIMRANGE(P:Q))</f>
      <v/>
    </c>"""),
        
        # In AVERAGE
        ("B2", '=AVERAGE(TRIMRANGE(A1:D10,2,2))', """
    <c r="B2">
      <f>AVERAGE(_xlfn.TRIMRANGE(A1:D10,2,2))</f>
      <v/>
    </c>"""),
        
        # In COUNTA
        ("B3", '=COUNTA(TRIMRANGE(A1:C10))', """
    <c r="B3">
      <f>COUNTA(_xlfn.TRIMRANGE(A1:C10))</f>
      <v/>
    </c>"""),
    ]
    
    for cell_ref, formula, expected in test_cases:
        cell = ws[cell_ref]
        cell.value = formula
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, f"Failed for {formula}: {diff}"


def test_trimrange_with_array_functions(worksheet, write_cell_implementation):
    """Test TRIMRANGE combined with other array functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # TRIMRANGE with SORT
        ("C1", '=SORT(TRIMRANGE(J22:L26))', """
    <c r="C1">
      <f>_xlfn._xlws.SORT(_xlfn.TRIMRANGE(J22:L26))</f>
      <v/>
    </c>"""),
        
        # TRIMRANGE with UNIQUE
        ("C2", '=UNIQUE(TRIMRANGE(J5:L10,2,2))', """
    <c r="C2">
      <f>_xlfn.UNIQUE(_xlfn.TRIMRANGE(J5:L10,2,2))</f>
      <v/>
    </c>"""),
        
        # TRIMRANGE with VSTACK
        ("C3", '=VSTACK(TRIMRANGE(A1:B5),TRIMRANGE(C1:D5))', """
    <c r="C3">
      <f>_xlfn.VSTACK(_xlfn.TRIMRANGE(A1:B5),_xlfn.TRIMRANGE(C1:D5))</f>
      <v/>
    </c>"""),
        
        # TRIMRANGE with HSTACK
        ("C4", '=HSTACK(TRIMRANGE(A1:B5,1,1),{"A";"B"})', """
    <c r="C4">
      <f>_xlfn.HSTACK(_xlfn.TRIMRANGE(A1:B5,1,1),{"A";"B"})</f>
      <v/>
    </c>"""),
    ]
    
    for cell_ref, formula, expected in test_cases:
        cell = ws[cell_ref]
        cell.value = formula
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, f"Failed for {formula}: {diff}"


def test_trimrange_complex_formulas(worksheet, write_cell_implementation):
    """Test TRIMRANGE in complex formulas"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # TRIMRANGE in SUMPRODUCT
        ("D1", '=SUMPRODUCT(TRIMRANGE(A1:C10,3,0)*{1;2;3;4;5;6})', """
    <c r="D1">
      <f>SUMPRODUCT(_xlfn.TRIMRANGE(A1:C10,3,0)*{1;2;3;4;5;6})</f>
      <v/>
    </c>"""),
        
        # Multiple TRIMRANGE calls
        ("D2", '=MAX(TRIMRANGE(A1:B10))+MIN(TRIMRANGE(C1:D10))', """
    <c r="D2">
      <f>MAX(_xlfn.TRIMRANGE(A1:B10))+MIN(_xlfn.TRIMRANGE(C1:D10))</f>
      <v/>
    </c>"""),
        
        # TRIMRANGE with CONCATENATE
        ("D3", '=CONCATENATE("Max:",MAX(TRIMRANGE(A1:B10))," Min:",MIN(TRIMRANGE(A1:B10)))', """
    <c r="D3">
      <f>CONCATENATE("Max:",MAX(_xlfn.TRIMRANGE(A1:B10))," Min:",MIN(_xlfn.TRIMRANGE(A1:B10)))</f>
      <v/>
    </c>"""),
    ]
    
    for cell_ref, formula, expected in test_cases:
        cell = ws[cell_ref]
        cell.value = formula
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, f"Failed for {formula}: {diff}"


def test_trimrange_spill_formula(worksheet, write_cell_implementation):
    """Test TRIMRANGE as a spill formula"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # TRIMRANGE as spill formula
    cell = ws["E1"]
    cell.value = '=TRIMRANGE(A1:C10)'
    cell._is_spill = True
    cell._spill_range = "E1:G7"
    
    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)
    
    expected = """
    <c r="E1" cm="1">
      <f t="array" ref="E1:G7">_xlfn.TRIMRANGE(A1:C10)</f>
      <v>0</v>
    </c>"""
    
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_trimrange_with_lambda(worksheet, write_cell_implementation):
    """Test TRIMRANGE with LAMBDA functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # LAMBDA using TRIMRANGE
    cell = ws["F1"]
    cell.value = '=LAMBDA(range,SUM(TRIMRANGE(range)))(A1:B10)'
    
    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)
    
    expected = """
    <c r="F1">
      <f t="array" ref="F1">_xlfn.LAMBDA(_xlpm.range,SUM(_xlfn.TRIMRANGE(_xlpm.range)))(A1:B10)</f>
      <v/>
    </c>"""
    
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_trimrange_with_let(worksheet, write_cell_implementation):
    """Test TRIMRANGE with LET function"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # LET with TRIMRANGE
    cell = ws["G1"]
    cell.value = '=LET(data,A1:B10,trimmed,TRIMRANGE(data),SUM(trimmed))'
    
    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)
    
    expected = """
    <c r="G1">
      <f t="array" ref="G1">_xlfn.LET(_xlpm.data,A1:B10,_xlpm.trimmed,_xlfn.TRIMRANGE(_xlpm.data),SUM(_xlpm.trimmed))</f>
      <v/>
    </c>"""
    
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff