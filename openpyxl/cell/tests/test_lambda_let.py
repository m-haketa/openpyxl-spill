"""
Test LAMBDA and LET functions with proper prefix handling
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


def test_lambda_basic(worksheet, write_cell_implementation):
    """Test basic LAMBDA functions with _xlpm prefix for parameters"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # Test cases for basic LAMBDA functions
    test_cases = [
        # Simple LAMBDA with one parameter
        ("A1", '=LAMBDA(x,x*2)(5)', """
    <c r="A1">
      <f>_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(5)</f>
      <v/>
    </c>"""),
        
        # LAMBDA with two parameters
        ("A2", '=LAMBDA(x,y,x+y)(3,4)', """
    <c r="A2">
      <f>_xlfn.LAMBDA(_xlpm.x,_xlpm.y,_xlpm.x+_xlpm.y)(3,4)</f>
      <v/>
    </c>"""),
        
        # LAMBDA with three parameters
        ("A3", '=LAMBDA(a,b,c,a+b*c)(2,3,4)', """
    <c r="A3">
      <f>_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.c,_xlpm.a+_xlpm.b*_xlpm.c)(2,3,4)</f>
      <v/>
    </c>"""),
        
        # LAMBDA with string concatenation
        ("A4", '=LAMBDA(x,y,CONCATENATE(x," ",y))("Hello","World")', """
    <c r="A4">
      <f>_xlfn.LAMBDA(_xlpm.x,_xlpm.y,CONCATENATE(_xlpm.x," ",_xlpm.y))("Hello","World")</f>
      <v/>
    </c>"""),
    ]
    
    # Execute each test case
    for cell_ref, formula, expected in test_cases:
        cell = ws[cell_ref]
        cell.value = formula
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, f"Failed for {formula}: {diff}"


def test_lambda_nested(worksheet, write_cell_implementation):
    """Test nested LAMBDA functions (currying)"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # LAMBDA returning LAMBDA
        ("B1", '=LAMBDA(x,LAMBDA(y,x+y))(5)(3)', """
    <c r="B1">
      <f>_xlfn.LAMBDA(_xlpm.x,_xlfn.LAMBDA(_xlpm.y,_xlpm.x+_xlpm.y))(5)(3)</f>
      <v/>
    </c>"""),
        
        # Triple nested LAMBDA
        ("B2", '=LAMBDA(x,LAMBDA(y,LAMBDA(z,x+y+z)))(1)(2)(3)', """
    <c r="B2">
      <f>_xlfn.LAMBDA(_xlpm.x,_xlfn.LAMBDA(_xlpm.y,_xlfn.LAMBDA(_xlpm.z,_xlpm.x+_xlpm.y+_xlpm.z)))(1)(2)(3)</f>
      <v/>
    </c>"""),
        
        # Conditional LAMBDA selection
        ("B3", '=LAMBDA(x,IF(x>0,LAMBDA(y,x+y),LAMBDA(y,x-y)))(5)(3)', """
    <c r="B3">
      <f>_xlfn.LAMBDA(_xlpm.x,IF(_xlpm.x>0,_xlfn.LAMBDA(_xlpm.y,_xlpm.x+_xlpm.y),_xlfn.LAMBDA(_xlpm.y,_xlpm.x-_xlpm.y)))(5)(3)</f>
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


def test_lambda_with_array_functions(worksheet, write_cell_implementation):
    """Test LAMBDA with array functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # LAMBDA with SEQUENCE (spill)
        ("C1", '=LAMBDA(n,SEQUENCE(n))(5)', 'C1:C5', """
    <c r="C1" cm="1">
      <f t="array" ref="C1:C5">_xlfn.LAMBDA(_xlpm.n,_xlfn.SEQUENCE(_xlpm.n))(5)</f>
      <v>0</v>
    </c>"""),
        
        # LAMBDA with FILTER
        ("C2", '=LAMBDA(arr,limit,FILTER(arr,arr>limit))({1,2,3,4,5},3)', 'C2:C4', """
    <c r="C2" cm="1">
      <f t="array" ref="C2:C4">_xlfn.LAMBDA(_xlpm.arr,_xlpm.limit,_xlfn._xlws.FILTER(_xlpm.arr,_xlpm.arr>_xlpm.limit))({1,2,3,4,5},3)</f>
      <v>0</v>
    </c>"""),
        
        # LAMBDA with SORT
        ("C3", '=LAMBDA(arr,SORT(arr,1,-1))({5,2,8,1,9})', 'C3:C7', """
    <c r="C3" cm="1">
      <f t="array" ref="C3:C7">_xlfn.LAMBDA(_xlpm.arr,_xlfn._xlws.SORT(_xlpm.arr,1,-1))({5,2,8,1,9})</f>
      <v>0</v>
    </c>"""),
        
        # LAMBDA with UNIQUE
        ("C4", '=LAMBDA(arr,UNIQUE(arr))({1,2,2,3,3,3})', 'C4:C6', """
    <c r="C4" cm="1">
      <f t="array" ref="C4:C6">_xlfn.LAMBDA(_xlpm.arr,_xlfn.UNIQUE(_xlpm.arr))({1,2,2,3,3,3})</f>
      <v>0</v>
    </c>"""),
    ]
    
    for i, (cell_ref, formula, spill_range, expected) in enumerate(test_cases):
        cell = ws[cell_ref]
        cell.value = formula
        # First test case spills, others don't for simplicity
        if i < len(test_cases):
            cell._is_spill = True
            cell._spill_range = spill_range
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, f"Failed for {formula}: {diff}"


def test_let_basic(worksheet, write_cell_implementation):
    """Test basic LET functions with _xlpm prefix for variables"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # Single variable
        ("D1", '=LET(x,10,x*2)', """
    <c r="D1">
      <f>_xlfn.LET(_xlpm.x,10,_xlpm.x*2)</f>
      <v/>
    </c>"""),
        
        # Multiple variables
        ("D2", '=LET(x,5,y,10,x+y)', """
    <c r="D2">
      <f>_xlfn.LET(_xlpm.x,5,_xlpm.y,10,_xlpm.x+_xlpm.y)</f>
      <v/>
    </c>"""),
        
        # Variable dependencies
        ("D3", '=LET(x,5,y,x*2,z,y+3,x+y+z)', """
    <c r="D3">
      <f>_xlfn.LET(_xlpm.x,5,_xlpm.y,_xlpm.x*2,_xlpm.z,_xlpm.y+3,_xlpm.x+_xlpm.y+_xlpm.z)</f>
      <v/>
    </c>"""),
        
        # String variables
        ("D4", '=LET(prefix,"ID-",num,123,CONCATENATE(prefix,num))', """
    <c r="D4">
      <f>_xlfn.LET(_xlpm.prefix,"ID-",_xlpm.num,123,CONCATENATE(_xlpm.prefix,_xlpm.num))</f>
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


def test_let_with_lambda(worksheet, write_cell_implementation):
    """Test LET combined with LAMBDA functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # LAMBDA as a variable
        ("E1", '=LET(double,LAMBDA(x,x*2),double(15))', """
    <c r="E1">
      <f>_xlfn.LET(_xlpm.double,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2),_xlpm.double(15))</f>
      <v/>
    </c>"""),
        
        # Multiple LAMBDAs
        ("E2", '=LET(add,LAMBDA(x,y,x+y),mul,LAMBDA(x,y,x*y),add(3,mul(4,5)))', """
    <c r="E2">
      <f>_xlfn.LET(_xlpm.add,_xlfn.LAMBDA(_xlpm.x,_xlpm.y,_xlpm.x+_xlpm.y),_xlpm.mul,_xlfn.LAMBDA(_xlpm.x,_xlpm.y,_xlpm.x*_xlpm.y),_xlpm.add(3,_xlpm.mul(4,5)))</f>
      <v/>
    </c>"""),
        
        # Conditional LAMBDA
        ("E3", '=LET(check,LAMBDA(x,IF(x>0,"正","負")),check(5))', """
    <c r="E3">
      <f>_xlfn.LET(_xlpm.check,_xlfn.LAMBDA(_xlpm.x,IF(_xlpm.x>0,"正","負")),_xlpm.check(5))</f>
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


def test_let_with_array_functions(worksheet, write_cell_implementation):
    """Test LET with array functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # LET with SEQUENCE
        ("F1", '=LET(size,5,arr,SEQUENCE(size),SUM(arr))', """
    <c r="F1">
      <f>_xlfn.LET(_xlpm.size,5,_xlpm.arr,_xlfn.SEQUENCE(_xlpm.size),SUM(_xlpm.arr))</f>
      <v/>
    </c>"""),
        
        # LET with FILTER
        ("F2", '=LET(data,{1,2,3,4,5},filtered,FILTER(data,data>2),SUM(filtered))', """
    <c r="F2">
      <f>_xlfn.LET(_xlpm.data,{1,2,3,4,5},_xlpm.filtered,_xlfn._xlws.FILTER(_xlpm.data,_xlpm.data>2),SUM(_xlpm.filtered))</f>
      <v/>
    </c>"""),
        
        # LET with array operations (spill)
        ("F3", '=LET(vals,{10,20,30,40,50},threshold,25,FILTER(vals,vals>threshold))', 'F3:F5', """
    <c r="F3" cm="1">
      <f t="array" ref="F3:F5">_xlfn.LET(_xlpm.vals,{10,20,30,40,50},_xlpm.threshold,25,_xlfn._xlws.FILTER(_xlpm.vals,_xlpm.vals>_xlpm.threshold))</f>
      <v>0</v>
    </c>"""),
    ]
    
    for i, item in enumerate(test_cases):
        if len(item) == 4:  # Has spill range
            cell_ref, formula, spill_range, expected = item
            cell = ws[cell_ref]
            cell.value = formula
            cell._is_spill = True
            cell._spill_range = spill_range
        else:  # No spill
            cell_ref, formula, expected = item
            cell = ws[cell_ref]
            cell.value = formula
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, f"Failed for {formula}: {diff}"


def test_text_processing_with_lambda_let(worksheet, write_cell_implementation):
    """Test that string literals are not modified in LET/LAMBDA"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # String literal should not have _xlpm prefix
        ("G1", '=LET(text,"A,B,C",TEXTSPLIT(text,","))', 'G1:G3', """
    <c r="G1" cm="1">
      <f t="array" ref="G1:G3">_xlfn.LET(_xlpm.text,"A,B,C",_xlfn.TEXTSPLIT(_xlpm.text,","))</f>
      <v>0</v>
    </c>"""),
        
        # TEXTBEFORE in LET
        ("G2", '=LET(email,"user@example.com",TEXTBEFORE(email,"@"))', """
    <c r="G2">
      <f>_xlfn.LET(_xlpm.email,"user@example.com",_xlfn.TEXTBEFORE(_xlpm.email,"@"))</f>
      <v/>
    </c>"""),
        
        # LAMBDA with TEXTBEFORE
        ("G3", '=LET(getName,LAMBDA(email,TEXTBEFORE(email,"@")),getName("john@company.com"))', """
    <c r="G3">
      <f>_xlfn.LET(_xlpm.getName,_xlfn.LAMBDA(_xlpm.email,_xlfn.TEXTBEFORE(_xlpm.email,"@")),_xlpm.getName("john@company.com"))</f>
      <v/>
    </c>"""),
    ]
    
    for i, item in enumerate(test_cases):
        if len(item) == 4:  # Has spill range
            cell_ref, formula, spill_range, expected = item
            cell = ws[cell_ref]
            cell.value = formula
            cell._is_spill = True
            cell._spill_range = spill_range
        else:  # No spill
            cell_ref, formula, expected = item
            cell = ws[cell_ref]
            cell.value = formula
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, f"Failed for {formula}: {diff}"


def test_lambda_let_edge_cases(worksheet, write_cell_implementation):
    """Test edge cases for LAMBDA and LET functions"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # Empty array handling
        ("H1", '=LET(empty,FILTER({1,2,3},FALSE),IFERROR(SUM(empty),0))', """
    <c r="H1">
      <f>_xlfn.LET(_xlpm.empty,_xlfn._xlws.FILTER({1,2,3},FALSE),IFERROR(SUM(_xlpm.empty),0))</f>
      <v/>
    </c>"""),
        
        # Type conversion
        ("H2", '=LET(txt,"123",num,VALUE(txt),num*2)', """
    <c r="H2">
      <f>_xlfn.LET(_xlpm.txt,"123",_xlpm.num,VALUE(_xlpm.txt),_xlpm.num*2)</f>
      <v/>
    </c>"""),
        
        # Error handling in LAMBDA
        ("H3", '=LAMBDA(x,y,IFERROR(x/y,"Error"))(10,0)', """
    <c r="H3">
      <f>_xlfn.LAMBDA(_xlpm.x,_xlpm.y,IFERROR(_xlpm.x/_xlpm.y,"Error"))(10,0)</f>
      <v/>
    </c>"""),
        
        # Range checking LAMBDA
        ("H4", '=LET(checkRange,LAMBDA(x,min,max,AND(x>=min,x<=max)),checkRange(15,10,20))', """
    <c r="H4">
      <f>_xlfn.LET(_xlpm.checkRange,_xlfn.LAMBDA(_xlpm.x,_xlpm.min,_xlpm.max,AND(_xlpm.x>=_xlpm.min,_xlpm.x<=_xlpm.max)),_xlpm.checkRange(15,10,20))</f>
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