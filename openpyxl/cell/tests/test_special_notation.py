# Copyright (c) 2010-2024 openpyxl

"""
Test special cell range notations (.:.、:.、.:) and their conversion to _TRO_* functions
"""

import pytest
from io import BytesIO
from openpyxl.xml.functions import xmlfile
from openpyxl.cell._writer import (
    etree_write_cell,
    lxml_write_cell
)
from openpyxl.cell.formula_utils import (
    prepare_spill_formula,
    _convert_tro_notations
)
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


class TestSpecialNotationConversion:
    """Test _convert_special_notation function"""
    
    @pytest.mark.parametrize("formula,expected", [
        # Basic cell ranges
        ("A1.:.B10", "_xlfn._TRO_ALL(A1:B10)"),
        ("A1:.B10", "_xlfn._TRO_TRAILING(A1:B10)"),
        ("A1.:B10", "_xlfn._TRO_LEADING(A1:B10)"),
        # Column only
        ("A.:.C", "_xlfn._TRO_ALL(A:C)"),
        ("A:.C", "_xlfn._TRO_TRAILING(A:C)"),
        ("A.:C", "_xlfn._TRO_LEADING(A:C)"),
        # Row only
        ("1.:.10", "_xlfn._TRO_ALL(1:10)"),
        ("5:.20", "_xlfn._TRO_TRAILING(5:20)"),
        ("100.:500", "_xlfn._TRO_LEADING(100:500)"),
        # Absolute references
        ("$A$1.:.$B$10", "_xlfn._TRO_ALL($A$1:$B$10)"),
        ("$A.:.$C", "_xlfn._TRO_ALL($A:$C)"),
        # Sheet references
        ("Sheet1!A1.:.B10", "_xlfn._TRO_ALL(Sheet1!A1:Sheet1!B10)"),
        ("Data!A:.C", "_xlfn._TRO_TRAILING(Data!A:Data!C)"),
        # Normal colon (no conversion)
        ("A1:B10", "A1:B10"),
        ("Sheet1!A:C", "Sheet1!A:C"),
        # Multiple notations
        ("A1.:.B10+C1.:.D10", "_xlfn._TRO_ALL(A1:B10)+_xlfn._TRO_ALL(C1:D10)"),
    ])
    def test_convert_special_notation(self, formula, expected):
        assert _convert_tro_notations(formula) == expected
    
    @pytest.mark.parametrize("formula,expected", [
        # Edge cases
        ("", ""),
        (".:.", ".:."),
        ("A1.:.", "A1.:."),
        (".:.B10", ".:.B10"),
        # Case sensitivity
        ("a1.:.b10", "a1.:.b10"),  # lowercase not converted
        ("AA.:.AC", "_xlfn._TRO_ALL(AA:AC)"),  # uppercase converted
    ])
    def test_edge_cases(self, formula, expected):
        assert _convert_tro_notations(formula) == expected


class TestAddFunctionPrefix:
    """Test _add_function_prefix with special notations"""
    
    @pytest.mark.parametrize("formula,expected", [
        # Basic conversions
        ("=A1.:.B10", "=_xlfn._TRO_ALL(A1:B10)"),
        ("=A1:.B10", "=_xlfn._TRO_TRAILING(A1:B10)"),
        ("=A1.:B10", "=_xlfn._TRO_LEADING(A1:B10)"),
        # In functions
        ("=SUM(A1.:.B10)", "=SUM(_xlfn._TRO_ALL(A1:B10))"),
        ("=AVERAGE(P:.Q)", "=AVERAGE(_xlfn._TRO_TRAILING(P:Q))"),
        ("=COUNT(5.:10)", "=COUNT(_xlfn._TRO_LEADING(5:10))"),
        # With other new functions
        ("=SORT(A1.:.B10)", "=_xlfn._xlws.SORT(_xlfn._TRO_ALL(A1:B10))"),
        ("=UNIQUE(A:.C)", "=_xlfn.UNIQUE(_xlfn._TRO_TRAILING(A:C))"),
        # Complex formulas
        ("=IF(SUM(A1.:.B10)>100,\"Large\",\"Small\")", 
         "=IF(SUM(_xlfn._TRO_ALL(A1:B10))>100,\"Large\",\"Small\")"),
        # LAMBDA with special notation
        ("=LAMBDA(x,SUM(x))(A1.:.B10)",
         "=_xlfn.LAMBDA(_xlpm.x,SUM(_xlpm.x))(_xlfn._TRO_ALL(A1:B10))"),
    ])
    def test_add_function_prefix(self, formula, expected):
        class MockCell:
            coordinate = "A1"
        result, _ = prepare_spill_formula(formula, MockCell())
        assert result == expected


def test_special_notation_basic(worksheet, write_cell_implementation):
    """Test basic special notation conversion in cell writing"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # .:. notation (all)
        ("A1", '=J5.:.L10', """
    <c r="A1">
      <f>_xlfn._TRO_ALL(J5:L10)</f>
      <v/>
    </c>"""),
        
        # :. notation (trailing)
        ("A2", '=J5:.L10', """
    <c r="A2">
      <f>_xlfn._TRO_TRAILING(J5:L10)</f>
      <v/>
    </c>"""),
        
        # .: notation (leading)
        ("A3", '=J5.:L10', """
    <c r="A3">
      <f>_xlfn._TRO_LEADING(J5:L10)</f>
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


def test_special_notation_column_row(worksheet, write_cell_implementation):
    """Test column and row only special notations"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # Column notation
        ("B1", '=P.:.Q', """
    <c r="B1">
      <f>_xlfn._TRO_ALL(P:Q)</f>
      <v/>
    </c>"""),
        
        # Row notation
        ("B2", '=11.:.11', """
    <c r="B2">
      <f>_xlfn._TRO_ALL(11:11)</f>
      <v/>
    </c>"""),
        
        # In SUM function
        ("B3", '=SUM(P:.Q)', """
    <c r="B3">
      <f>SUM(_xlfn._TRO_TRAILING(P:Q))</f>
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


def test_special_notation_complex(worksheet, write_cell_implementation):
    """Test complex formulas with special notations"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # SORT with special notation
        ("C1", '=SORT(J22.:.L26)', """
    <c r="C1">
      <f>_xlfn._xlws.SORT(_xlfn._TRO_ALL(J22:L26))</f>
      <v/>
    </c>"""),
        
        # UNIQUE with special notation
        ("C2", '=UNIQUE(J5:.L10)', """
    <c r="C2">
      <f>_xlfn.UNIQUE(_xlfn._TRO_TRAILING(J5:L10))</f>
      <v/>
    </c>"""),
        
        # VLOOKUP with special notation
        ("C3", '=VLOOKUP(70,J5.:L10,3,FALSE)', """
    <c r="C3">
      <f>VLOOKUP(70,_xlfn._TRO_LEADING(J5:L10),3,FALSE)</f>
      <v/>
    </c>"""),
        
        # Multiple special notations
        ("C4", '=J5.:.J10*N5.:.N10', """
    <c r="C4">
      <f>_xlfn._TRO_ALL(J5:J10)*_xlfn._TRO_ALL(N5:N10)</f>
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


def test_special_notation_with_lambda(worksheet, write_cell_implementation):
    """Test special notations combined with LAMBDA functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # LAMBDA with special notation as argument
    cell = ws["D1"]
    cell.value = '=LAMBDA(range,SUM(range))(A1.:.B10)'
    
    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)
    
    expected = """
    <c r="D1">
      <f t="array" ref="D1">_xlfn.LAMBDA(_xlpm.range,SUM(_xlpm.range))(_xlfn._TRO_ALL(A1:B10))</f>
      <v/>
    </c>"""
    
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_special_notation_let_function(worksheet, write_cell_implementation):
    """Test special notations in LET function"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # LET with special notation
    cell = ws["E1"]
    cell.value = '=LET(data,A1.:.B10,total,SUM(data),total*2)'
    
    out = BytesIO()
    with xmlfile(out) as xf:
        write_cell(xf, ws, cell)
    
    expected = """
    <c r="E1">
      <f t="array" ref="E1">_xlfn.LET(_xlpm.data,_xlfn._TRO_ALL(A1:B10),_xlpm.total,SUM(_xlpm.data),_xlpm.total*2)</f>
      <v/>
    </c>"""
    
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff