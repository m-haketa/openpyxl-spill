"""
Test aggregation functions (GROUPBY, PIVOTBY, PERCENTOF) with proper prefix handling
"""

import pytest
from io import BytesIO
from openpyxl.xml.functions import xmlfile
from openpyxl.cell._writer import (
    etree_write_cell,
    lxml_write_cell
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


def test_groupby_basic(worksheet, write_cell_implementation):
    """Test GROUPBY function with _xlfn prefix and _xleta for aggregate functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # Test cases for GROUPBY functions
    test_cases = [
        # GROUPBY with SUM (requires _xleta prefix)
        ("A1", '=GROUPBY(E2:E7, F2:F7, SUM)', """
    <c r="A1">
      <f t="array" ref="A1">_xlfn.GROUPBY(E2:E7, F2:F7, _xleta.SUM)</f>
      <v/>
    </c>"""),
        
        # GROUPBY with AVERAGE (requires _xleta prefix)
        ("A2", '=GROUPBY(E2:E7, F2:F7, AVERAGE)', """
    <c r="A2">
      <f t="array" ref="A2">_xlfn.GROUPBY(E2:E7, F2:F7, _xleta.AVERAGE)</f>
      <v/>
    </c>"""),
        
        # GROUPBY with LAMBDA (no _xleta prefix)
        ("A3", '=GROUPBY(E2:E7, F2:F7, LAMBDA(x, MAX(x)-MIN(x)))', """
    <c r="A3">
      <f t="array" ref="A3">_xlfn.GROUPBY(E2:E7, F2:F7, _xlfn.LAMBDA(_xlpm.x, MAX(_xlpm.x)-MIN(_xlpm.x)))</f>
      <v/>
    </c>"""),
        
        # GROUPBY with PERCENTOF (requires _xleta prefix)
        ("A4", '=GROUPBY(E2:E7, F2:F7, PERCENTOF)', """
    <c r="A4">
      <f t="array" ref="A4">_xlfn.GROUPBY(E2:E7, F2:F7, _xleta.PERCENTOF)</f>
      <v/>
    </c>"""),
    ]
    
    for coord, formula, expected in test_cases:
        ws[coord] = formula
        cell = ws[coord]
        cell._is_spill = True
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell, cell.has_style)
        
        xml = out.getvalue()
        compare_xml(xml, expected)


def test_pivotby_basic(worksheet, write_cell_implementation):
    """Test PIVOTBY function with _xlfn prefix and _xleta for aggregate functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # Test cases for PIVOTBY functions
    test_cases = [
        # PIVOTBY with SUM (requires _xleta prefix)
        ("B1", '=PIVOTBY(H2:H7, I2:I7, J2:J7, SUM)', """
    <c r="B1">
      <f t="array" ref="B1">_xlfn.PIVOTBY(H2:H7, I2:I7, J2:J7, _xleta.SUM)</f>
      <v/>
    </c>"""),
        
        # PIVOTBY with COUNT (requires _xleta prefix)
        ("B2", '=PIVOTBY(H2:H7, I2:I7, J2:J7, COUNT)', """
    <c r="B2">
      <f t="array" ref="B2">_xlfn.PIVOTBY(H2:H7, I2:I7, J2:J7, _xleta.COUNT)</f>
      <v/>
    </c>"""),
        
        # PIVOTBY with LAMBDA (no _xleta prefix)
        ("B3", '=PIVOTBY(H2:H7, I2:I7, J2:J7, LAMBDA(x, AVERAGE(x)))', """
    <c r="B3">
      <f t="array" ref="B3">_xlfn.PIVOTBY(H2:H7, I2:I7, J2:J7, _xlfn.LAMBDA(_xlpm.x, AVERAGE(_xlpm.x)))</f>
      <v/>
    </c>"""),
        
        # PIVOTBY with PERCENTOF (requires _xleta prefix)
        ("B4", '=PIVOTBY(H2:H7, I2:I7, J2:J7, PERCENTOF)', """
    <c r="B4">
      <f t="array" ref="B4">_xlfn.PIVOTBY(H2:H7, I2:I7, J2:J7, _xleta.PERCENTOF)</f>
      <v/>
    </c>"""),
    ]
    
    for coord, formula, expected in test_cases:
        ws[coord] = formula
        cell = ws[coord]
        cell._is_spill = True
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell, cell.has_style)
        
        xml = out.getvalue()
        compare_xml(xml, expected)


def test_percentof_basic(worksheet, write_cell_implementation):
    """Test PERCENTOF function with _xlfn prefix"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # Test cases for PERCENTOF functions
    test_cases = [
        # PERCENTOF basic usage (subset and whole)
        ("C1", '=PERCENTOF(F2:F4, F2:F7)', """
    <c r="C1">
      <f>_xlfn.PERCENTOF(F2:F4, F2:F7)</f>
      <v/>
    </c>"""),
        
        # PERCENTOF in LET function
        ("C2", '=LET(catA, SUMIF(E2:E7, "A", F2:F7), total, SUM(F2:F7), PERCENTOF(catA, total))', """
    <c r="C2">
      <f>_xlfn.LET(_xlpm.catA,SUMIF(E2:E7, "A", F2:F7),_xlpm.total,SUM(F2:F7),_xlfn.PERCENTOF(_xlpm.catA, _xlpm.total))</f>
      <v/>
    </c>"""),
    ]
    
    for coord, formula, expected in test_cases:
        ws[coord] = formula
        cell = ws[coord]
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell, cell.has_style)
        
        xml = out.getvalue()
        compare_xml(xml, expected)


def test_complex_combinations(worksheet, write_cell_implementation):
    """Test complex combinations of aggregation functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # Test cases for complex combinations
    test_cases = [
        # LET with GROUPBY and HSTACK
        ("D1", '=LET(data, E2:E7, values, F2:F7, totals, GROUPBY(data, values, SUM), HSTACK(totals, INDEX(totals,,2)/SUM(values)))', """
    <c r="D1">
      <f t="array" ref="D1">_xlfn.LET(_xlpm.data, E2:E7, _xlpm.values, F2:F7, _xlpm.totals, _xlfn.GROUPBY(_xlpm.data, _xlpm.values, _xleta.SUM), _xlfn.HSTACK(_xlpm.totals, INDEX(_xlpm.totals,,2)/SUM(_xlpm.values)))</f>
      <v/>
    </c>"""),
        
        # Nested functions with multiple aggregate functions
        ("D2", '=GROUPBY(A2:A10, B2:B10, LAMBDA(x, SUM(x)/COUNT(x)))', """
    <c r="D2">
      <f t="array" ref="D2">_xlfn.GROUPBY(A2:A10, B2:B10, _xlfn.LAMBDA(_xlpm.x, SUM(_xlpm.x)/COUNT(_xlpm.x)))</f>
      <v/>
    </c>"""),
    ]
    
    for coord, formula, expected in test_cases:
        ws[coord] = formula
        cell = ws[coord]
        cell._is_spill = True
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell, cell.has_style)
        
        xml = out.getvalue()
        compare_xml(xml, expected)


def test_edge_cases(worksheet, write_cell_implementation):
    """Test edge cases for aggregation functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # Test cases for edge cases
    test_cases = [
        # GROUPBY with nested GROUPBY (unlikely but possible)
        ("E1", '=GROUPBY(A1:A10, GROUPBY(B1:B10, C1:C10, SUM), AVERAGE)', """
    <c r="E1">
      <f t="array" ref="E1">_xlfn.GROUPBY(A1:A10, _xlfn.GROUPBY(B1:B10, C1:C10, _xleta.SUM), _xleta.AVERAGE)</f>
      <v/>
    </c>"""),
        
        # PIVOTBY with complex expression in data argument
        ("E2", '=PIVOTBY(H2:H7, I2:I7, J2:J7*1.1, SUM)', """
    <c r="E2">
      <f t="array" ref="E2">_xlfn.PIVOTBY(H2:H7, I2:I7, J2:J7*1.1, _xleta.SUM)</f>
      <v/>
    </c>"""),
    ]
    
    for coord, formula, expected in test_cases:
        ws[coord] = formula
        cell = ws[coord]
        cell._is_spill = True
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell, cell.has_style)
        
        xml = out.getvalue()
        compare_xml(xml, expected)


if __name__ == "__main__":
    # Run tests with pytest
    pytest.main([__file__, "-v"])