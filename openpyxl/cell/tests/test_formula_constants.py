"""
Test that formula function names like LAMBDA, LET, FILTER are treated as string constants
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


@pytest.fixture
def dummy_cell(worksheet):
    """Create a dummy cell for testing"""
    from openpyxl.cell import Cell
    ws = worksheet
    cell = Cell(ws, column=1, row=1)
    return cell


@pytest.fixture(params=['lxml', 'etree'])
def write_cell_implementation(request):
    if request.param == 'lxml':
        return lxml_write_cell
    else:
        return etree_write_cell


def test_formula_names_as_constants(dummy_cell):
    """Test that function names like LAMBDA, LET, FILTER are treated as string constants"""
    test_values = [
        "LAMBDA(x,y,x+y)",
        "LET(x,1,x+2)",
        "FILTER(A:B,B:B>0)",
        "LAMBDA(x,y,LET(sum,x+y,product,x*y,sum*product))",
        "FILTER(A:C,(A:A>0)*(B:B<100))",
        "LET(x,5,y,10,z,x+y,z*2)",
        "LAMBDA(arr,REDUCE(0,arr,LAMBDA(a,b,a+b)))",
        "UNIQUE(FILTER(A:A,A:A<>\"\"))"
    ]
    
    for value in test_values:
        cell = dummy_cell
        cell.value = value
        # These should be treated as string constants, not formulas
        assert cell.data_type == 's', f"Failed for {value}: expected 's' but got '{cell.data_type}'"
        assert cell.value == value


def test_formula_constants_xml_output(worksheet, write_cell_implementation):
    """Test XML output for formula names as constants"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    # Simple test cases without special characters
    test_cases = [
        # Basic LAMBDA as string constant
        ("A1", "LAMBDA(x,y,x+y)"),
        # Basic LET as string constant
        ("A2", "LET(x,1,x+2)"),
        # Basic FILTER as string constant (simplified without comparison operators)
        ("A3", "FILTER(A:B,B:B)"),
        # Complex nested formula as string
        ("A4", "LAMBDA(x,y,LET(sum,x+y,product,x*y,sum*product))"),
    ]
    
    for cell_ref, value in test_cases:
        cell = ws[cell_ref]
        cell.value = value
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        # Check that it's treated as inlineStr (string) not formula
        assert b't="inlineStr"' in xml or b"t='inlineStr'" in xml, f"Failed for {value}: not treated as string"
        # Check that the value is contained in the XML
        assert value.encode() in xml or value.replace("<", "&lt;").replace(">", "&gt;").encode() in xml, f"Failed for {value}: value not found in XML"


def test_formula_constants_edge_cases(dummy_cell):
    """Test edge cases for formula-like string constants"""
    edge_cases = [
        # Function names with spaces (should be string)
        "LAMBDA (x, y, x + y)",
        "LET (x, 1, x + 2)",
        
        # Function names in lowercase (should be string)
        "lambda(x,y,x+y)",
        "let(x,1,x+2)",
        "filter(A:B,B:B>0)",
        
        # Mixed case (should be string)
        "Lambda(x,y,x+y)",
        "Let(x,1,x+2)",
        "Filter(A:B,B:B>0)",
        
        # Function names with special characters
        "LAMBDA#(x,y,x+y)",
        "LET$(x,1,x+2)",
        "FILTER!(A:B,B:B>0)",
        
        # Partial function expressions
        "LAMBDA(",
        "LET(x",
        "FILTER(A:B",
        
        # Function names alone
        "LAMBDA",
        "LET",
        "FILTER",
        "MAP",
        "REDUCE",
        "SCAN",
        "BYROW",
        "BYCOL",
        "MAKEARRAY",
        "ISOMITTED",
    ]
    
    for value in edge_cases:
        cell = dummy_cell
        cell.value = value
        # All these should be treated as strings, not formulas
        assert cell.data_type == 's', f"Failed for '{value}': expected 's' but got '{cell.data_type}'"
        assert cell.value == value


def test_actual_formulas_with_equals(dummy_cell):
    """Test that actual formulas (starting with =) are still treated as formulas"""
    formula_cases = [
        "=LAMBDA(x,y,x+y)(3,4)",
        "=LET(x,1,x+2)",
        "=FILTER(A:B,B:B>0)",
    ]
    
    for formula in formula_cases:
        cell = dummy_cell
        cell.value = formula
        # These should be treated as formulas because they start with =
        assert cell.data_type == 'f', f"Failed for {formula}: expected 'f' but got '{cell.data_type}'"


def test_phase6_function_names_as_constants(dummy_cell):
    """Test Phase 6 LAMBDA-based function names as constants"""
    phase6_functions = [
        # Basic function names
        "MAP(array,LAMBDA(x,x*2))",
        "REDUCE(0,array,LAMBDA(acc,val,acc+val))",
        "SCAN(0,array,LAMBDA(acc,val,acc+val))",
        "BYROW(array,LAMBDA(row,SUM(row)))",
        "BYCOL(array,LAMBDA(col,AVERAGE(col)))",
        "MAKEARRAY(3,3,LAMBDA(r,c,r*c))",
        
        # Complex combinations
        "MAP(FILTER(A1:A10,A1:A10>0),LAMBDA(x,x^2))",
        "REDUCE(1,MAP(A1:A5,LAMBDA(x,x*2)),LAMBDA(acc,val,acc*val))",
        "SCAN(0,FILTER(B1:B10,B1:B10<>0),LAMBDA(acc,val,acc+val))",
        
        # With ISOMITTED
        "LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))",
        "LAMBDA(price,[tax],price*(1+IF(ISOMITTED(tax),0.1,tax)))",
    ]
    
    for value in phase6_functions:
        cell = dummy_cell
        cell.value = value
        # These should be treated as string constants
        assert cell.data_type == 's', f"Failed for {value}: expected 's' but got '{cell.data_type}'"
        assert cell.value == value


def test_aggregation_function_names_as_constants(dummy_cell):
    """Test Phase 7 aggregation function names as constants"""
    phase7_functions = [
        # GROUPBY variations
        "GROUPBY(A:A,B:B,SUM)",
        "GROUPBY(data,categories,AVERAGE,3)",
        
        # PIVOTBY variations
        "PIVOTBY(A:A,B:B,C:C,D:D,SUM)",
        "PIVOTBY(rows,cols,values,data,AVERAGE,3)",
        
        # PERCENTOF variations
        "PERCENTOF(A1:A10)",
        "PERCENTOF(data,SUM(data))",
        "PERCENTOF(values,total,1)",
    ]
    
    for value in phase7_functions:
        cell = dummy_cell
        cell.value = value
        # These should be treated as string constants
        assert cell.data_type == 's', f"Failed for {value}: expected 's' but got '{cell.data_type}'"
        assert cell.value == value


def test_mixed_content_with_formula_names(worksheet, write_cell_implementation):
    """Test cells with mixed content including formula names"""
    write_cell = write_cell_implementation
    ws = worksheet
    
    test_cases = [
        # Text description with formula names
        ("B1", "Use LAMBDA(x,y,x+y) for addition", """
    <c r="B1" t="inlineStr">
      <is>
        <t>Use LAMBDA(x,y,x+y) for addition</t>
      </is>
    </c>"""),
        
        # Documentation text
        ("B2", "The LET function syntax: LET(name,value,calculation)", """
    <c r="B2" t="inlineStr">
      <is>
        <t>The LET function syntax: LET(name,value,calculation)</t>
      </is>
    </c>"""),
        
        # Example text
        ("B3", "Example: FILTER(A:A,A:A>100) filters values greater than 100", """
    <c r="B3" t="inlineStr">
      <is>
        <t>Example: FILTER(A:A,A:A&gt;100) filters values greater than 100</t>
      </is>
    </c>"""),
    ]
    
    for cell_ref, value, expected in test_cases:
        cell = ws[cell_ref]
        cell.value = value
        
        out = BytesIO()
        with xmlfile(out) as xf:
            write_cell(xf, ws, cell)
        
        xml = out.getvalue()
        diff = compare_xml(xml, expected)
        assert diff is None, f"Failed for {value}: {diff}"


def test_japanese_text_with_formula_names(dummy_cell):
    """Test Japanese text containing formula names"""
    japanese_cases = [
        "LAMBDA関数の使い方",
        "LET関数で変数を定義",
        "FILTERでデータを絞り込む",
        "これはLAMBDA(x,y,x+y)の例です",
        "LET(x,1,x+2)を使って計算",
    ]
    
    for value in japanese_cases:
        cell = dummy_cell
        cell.value = value
        # These should be treated as string constants
        assert cell.data_type == 's', f"Failed for {value}: expected 's' but got '{cell.data_type}'"
        assert cell.value == value