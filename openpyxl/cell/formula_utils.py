"""
Excel 365 spill formula processing utilities

Provides functionality to process Excel 365's new dynamic array formulas (spill formulas)
and LAMBDA/LET functions, adding appropriate prefixes.
"""

import re
from typing import Optional, Tuple, Dict, List, Any
import uuid


def prepare_spill_formula(formula: str, cell: Any) -> Tuple[str, Dict[str, Any]]:
    """
    Process Excel 365 spill formulas and add appropriate prefixes
    
    Args:
        formula: Excel formula string (starting with '=')
        cell: Cell object where the formula will be set
    
    Returns:
        tuple: (converted formula string, attribute dictionary)
    """
    if not formula or not formula.startswith('='):
        return formula, {}
    
    try:
        # Get formula body (excluding '=')
        formula_body = formula[1:]
        
        # 1. Protect string literals
        protected_formula, string_map = _protect_string_literals(formula_body)
        
        # 2. Protect array literals
        protected_formula, array_map = _protect_array_literals(protected_formula)
        
        # 3. LAMBDA/LET unified processing (highest priority)
        processed_formula = _process_lambda_let_unified(protected_formula)
        
        # 4. GROUPBY/PIVOTBY argument processing
        processed_formula = _process_groupby_pivotby_args(processed_formula)
        
        # 5. Function name conversion
        processed_formula = _add_function_prefixes(processed_formula)
        
        # 6. Special notation conversion
        processed_formula = _convert_tro_notations(processed_formula)
        
        # 7. Restore array literals
        processed_formula = _restore_array_literals(processed_formula, array_map)
        
        # 8. Restore string literals
        final_formula = _restore_string_literals(processed_formula, string_map)
        
        # 9. Attribute setting
        attributes = _determine_formula_attributes(final_formula, cell)
        
        # Return with '=' prefix
        return '=' + final_formula, attributes
        
    except Exception:
        # Return original formula and empty attributes on error
        return formula, {}


def _protect_string_literals(formula: str) -> Tuple[str, Dict[str, str]]:
    """
    Temporarily replace string literals with placeholders
    
    Args:
        formula: Formula string to process
    
    Returns:
        (Formula with placeholders, mapping of placeholders to original strings)
    """
    string_map = {}
    
    def replace_string(match):
        # Generate unique placeholder
        placeholder = f"__STR_{uuid.uuid4().hex[:8]}__"
        string_map[placeholder] = match.group(0)
        return placeholder
    
    # Replace double-quoted strings
    protected = re.sub(r'"(?:[^"]|"")*"', replace_string, formula)
    
    return protected, string_map


def _restore_string_literals(formula: str, string_map: Dict[str, str]) -> str:
    """
    Restore placeholders to original string literals
    
    Args:
        formula: Formula with placeholders
        string_map: Mapping of placeholders to original strings
    
    Returns:
        Formula with restored string literals
    """
    result = formula
    for placeholder, original in string_map.items():
        result = result.replace(placeholder, original)
    return result

def _protect_array_literals(formula: str) -> Tuple[str, Dict[str, str]]:
    """
    Temporarily replace array literals with placeholders
    Array literals are not nested, so we can use a simpler approach
    
    Args:
        formula: Formula string to process
    
    Returns:
        (Formula with placeholders, mapping of placeholders to original arrays)
    """
    array_map = {}
    
    def replace_array(match):
        # Generate unique placeholder
        placeholder = f"__ARR_{uuid.uuid4().hex[:8]}__"
        array_map[placeholder] = match.group(0)
        return placeholder
    
    # Replace array literals (no nesting, so simple regex works)
    # Match { followed by any characters except { or } and then }
    protected = re.sub(r'\{[^{}]*\}', replace_array, formula)
    
    return protected, array_map


def _restore_array_literals(formula: str, array_map: Dict[str, str]) -> str:
    """
    Restore placeholders to original array literals
    
    Args:
        formula: Formula with placeholders
        array_map: Mapping of placeholders to original arrays
    
    Returns:
        Formula with restored array literals
    """
    result = formula
    for placeholder, original in array_map.items():
        result = result.replace(placeholder, original)
    return result


def _process_lambda_let_unified(formula: str, scope: Optional[Dict[str, str]] = None) -> str:
    """
    Unified processing of LAMBDA/LET function variables/parameters
    
    Args:
        formula: Formula string to process
        scope: Current scope (mapping of variable names to their prefixes)
    
    Returns:
        Formula with processed variables/parameters
    """
    if scope is None:
        scope = {}
    
    result = []
    i = 0
    
    while i < len(formula):
        # Detect LAMBDA function
        if _is_function_at(formula, i, 'LAMBDA'):
            result.append('_xlfn.LAMBDA')
            i += 6  # Length of 'LAMBDA'
            
            # Parse arguments
            if i < len(formula) and formula[i] == '(':
                result.append('(')
                i += 1
                
                # Parse and process LAMBDA content
                params, body, end_pos = _parse_lambda_content(formula[i:])
                
                # Create new scope
                new_scope = scope.copy()
                
                # Process parameters
                processed_params = []
                for param in params:
                    param = param.strip()
                    if param:
                        # Check if optional parameter (with brackets)
                        if param.startswith('[') and param.endswith(']'):
                            # Remove brackets for optional parameters
                            clean_param = param[1:-1]
                            # Use _xlop. prefix for parameter definition, but _xlpm. for scope references
                            new_scope[clean_param] = f'_xlpm.{clean_param}'
                            processed_params.append(f'_xlop.{clean_param}')
                        else:
                            # Regular parameter with _xlpm. prefix
                            new_scope[param] = f'_xlpm.{param}'
                            processed_params.append(f'_xlpm.{param}')
                
                # Add parameter part to result
                if processed_params:
                    result.append(','.join(processed_params))
                    if body:
                        result.append(',')
                
                # Process body with new scope
                if body:
                    processed_body = _process_lambda_let_unified(body, new_scope)
                    # Replace variables in scope
                    processed_body = _replace_variables_in_scope(processed_body, new_scope)
                    result.append(processed_body)
                
                result.append(')')
                i += end_pos + 1
                
        # Detect LET function
        elif _is_function_at(formula, i, 'LET'):
            result.append('_xlfn.LET')
            i += 3  # Length of 'LET'
            
            # Parse arguments
            if i < len(formula) and formula[i] == '(':
                result.append('(')
                i += 1
                
                # Parse and process LET content
                var_pairs, final_expr, end_pos = _parse_let_content(formula[i:])
                
                # Create new scope (built incrementally for forward reference constraint)
                new_scope = scope.copy()
                processed_parts = []
                
                # Process variable definitions in order
                for var_name, var_expr in var_pairs:
                    var_name = var_name.strip()
                    if var_name:
                        # Process variable expression with current scope
                        processed_expr = _process_lambda_let_unified(var_expr, new_scope)
                        processed_expr = _replace_variables_in_scope(processed_expr, new_scope)
                        
                        # Add prefix to variable name and add to scope
                        prefixed_name = f'_xlpm.{var_name}'
                        new_scope[var_name] = prefixed_name
                        
                        # Add processed variable definition
                        processed_parts.append(prefixed_name)
                        processed_parts.append(processed_expr)
                
                # Process final expression
                if final_expr:
                    # First process nested LAMBDA/LET functions
                    processed_final = _process_lambda_let_unified(final_expr, new_scope)
                    # Then replace variables in the processed result
                    # The processed_final might have introduced new LAMBDA/LET that need variable replacement
                    processed_final = _replace_variables_in_scope(processed_final, new_scope)
                    processed_parts.append(processed_final)
                
                # Combine results
                result.append(','.join(processed_parts))
                result.append(')')
                i += end_pos + 1
                
        else:
            # Add regular character
            result.append(formula[i])
            i += 1
    
    return ''.join(result)


def _is_function_at(formula: str, pos: int, func_name: str) -> bool:
    """
    Check if a specific function exists at the specified position
    
    Args:
        formula: Formula string
        pos: Check position
        func_name: Function name
    
    Returns:
        True if function exists
    """
    # Check if position is out of range
    if pos >= len(formula) or pos + len(func_name) > len(formula):
        return False
    
    # Check if function name matches
    if formula[pos:pos + len(func_name)].upper() != func_name.upper():
        return False
    
    # Check that previous character is not part of a word
    if pos > 0:
        prev_char = formula[pos - 1]
        if prev_char.isalnum() or prev_char == '_' or prev_char == '.':
            return False
    
    # Check that next character is '('
    next_pos = pos + len(func_name)
    if next_pos < len(formula) and formula[next_pos] == '(':
        return True
    
    return False


def _parse_lambda_content(formula: str) -> Tuple[List[str], str, int]:
    """
    Parse LAMBDA function arguments and body
    
    Args:
        formula: LAMBDA function content inside parentheses (excluding parentheses)
    
    Returns:
        (parameter list, body expression, end position)
    """
    args = []
    depth = 0
    current = []
    i = 0
    
    while i < len(formula):
        char = formula[i]
        
        if char == '(':
            depth += 1
            current.append(char)
        elif char == ')':
            if depth == 0:
                # Add last argument
                if current:
                    args.append(''.join(current).strip())
                break
            else:
                depth -= 1
                current.append(char)
        elif char == ',' and depth == 0:
            # Argument separator
            if current:
                args.append(''.join(current).strip())
            current = []
        else:
            current.append(char)
        
        i += 1
    
    # Split into parameters and body
    # The last argument is always the body expression
    if len(args) > 1:
        params = args[:-1]
        body = args[-1]
    elif len(args) == 1:
        # Single argument - it's the body with no parameters
        params = []
        body = args[0]
    else:
        params = []
        body = ''
    
    return params, body, i


def _looks_like_parameter(text: str) -> bool:
    """
    Check if text looks like a parameter name
    
    Args:
        text: Text to check
    
    Returns:
        True if it looks like a parameter name
    """
    # Parameters are usually simple identifiers
    # They don't contain function calls or operators
    if '(' in text[:20] or '+' in text[:20] or '-' in text[:20]:
        return False
    
    # Check if first part matches identifier pattern
    match = re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*', text)
    if match:
        # Check if comma or closing parenthesis follows the identifier
        end_pos = match.end()
        if end_pos < len(text):
            next_char = text[end_pos:].lstrip()
            if next_char and next_char[0] in ',)':
                return True
    
    return False


def _parse_let_content(formula: str) -> Tuple[List[Tuple[str, str]], str, int]:
    """
    Parse LET function variable definitions and final expression
    
    Args:
        formula: LET function content inside parentheses (excluding parentheses)
    
    Returns:
        (list of variable definitions [(name, value expression),...], final expression, end position)
    """
    var_pairs = []
    paren_depth = 0
    current = []
    i = 0
    args = []
    
    # Split all arguments
    while i < len(formula):
        char = formula[i]
        
        if char == '(':
            paren_depth += 1
            current.append(char)
        elif char == ')':
            if paren_depth == 0:
                # Add last argument
                if current:
                    args.append(''.join(current).strip())
                    current = []  # Clear current to avoid double-adding
                break
            else:
                paren_depth -= 1
                current.append(char)

        elif char == ',' and paren_depth == 0:
            # Argument separator (only when not inside parentheses)
            if current:
                args.append(''.join(current).strip())
            current = []
        else:
            current.append(char)
        
        i += 1
    
    # Add last argument if there's any remaining content
    if current:
        args.append(''.join(current).strip())
    
    # Split arguments into pairs (last one is final expression)
    # LET has an odd number of arguments: pairs of (name, value) plus final expression
    # So we pair up all arguments except the last one if there's an odd count
    num_args = len(args)
    if num_args % 2 == 1:
        # Odd number of arguments - last one is the final expression
        for j in range(0, num_args - 1, 2):
            var_pairs.append((args[j], args[j + 1]))
    
    # Final expression
    final_expr = args[-1] if len(args) % 2 == 1 else ''
    
    return var_pairs, final_expr, i


def _replace_variables_in_scope(expr: str, scope: Dict[str, str]) -> str:
    """
    Replace variable references in expression based on scope
    
    Args:
        expr: Expression to process
        scope: Current scope
    
    Returns:
        Expression with replaced variables
    """
    if not scope:
        return expr
    
    result = expr
    
    # Replace each variable in scope
    for var_name, prefixed_name in scope.items():
        # Skip if variable name is a pure number (e.g., "2", "4")
        # This can happen with malformed variable names
        if var_name.isdigit():
            continue
            
        # Use negative lookbehind to avoid replacing already prefixed variables
        # This pattern matches the variable name only when it's not preceded by '_xlpm.' or '_xlop.'
        pattern = r'(?<!_xlpm\.)(?<!_xlop\.)\b' + re.escape(var_name) + r'\b'
        result = re.sub(pattern, prefixed_name, result)
    
    return result


def _process_groupby_pivotby_args(formula: str) -> str:
    """
    Process GROUPBY/PIVOTBY aggregate function arguments
    
    Args:
        formula: Formula string to process
    
    Returns:
        Processed formula string
    """
    result = formula
    
    # Process GROUPBY function (3rd argument)
    result = _process_aggregate_function(result, 'GROUPBY', 2)
    
    # Process PIVOTBY function (4th argument)
    result = _process_aggregate_function(result, 'PIVOTBY', 3)
    
    return result


def _process_aggregate_function(formula: str, func_name: str, arg_index: int) -> str:
    """
    Process aggregate function argument of specified function
    
    Args:
        formula: Formula string to process
        func_name: Function name (GROUPBY or PIVOTBY)
        arg_index: Argument index (0-based)
    
    Returns:
        Processed formula string
    """
    pattern = r'\b' + func_name + r'\s*\('
    matches = list(re.finditer(pattern, formula, re.IGNORECASE))
    
    # Process from back to front (to avoid position shifts)
    for match in reversed(matches):
        start_pos = match.end()
        
        # Parse arguments
        args = _split_function_args(formula, start_pos)
        
        if len(args) > arg_index:
            arg_content = args[arg_index].strip()
            
            # Add _xleta. if not LAMBDA (check both original and processed forms)
            if not arg_content.upper().startswith('LAMBDA') and not arg_content.startswith('_xlfn.LAMBDA'):
                # Only add if prefix not already present
                if not arg_content.startswith('_xleta.'):
                    # Find argument position
                    arg_start = start_pos
                    for i in range(arg_index):
                        arg_start = formula.find(',', arg_start) + 1
                    
                    # Skip spaces
                    while arg_start < len(formula) and formula[arg_start].isspace():
                        arg_start += 1
                    
                    # Add prefix
                    formula = formula[:arg_start] + '_xleta.' + formula[arg_start:]
    
    return formula


def _split_function_args(formula: str, start_pos: int) -> List[str]:
    """
    Split function arguments
    
    Args:
        formula: Formula string
        start_pos: Argument start position (after opening parenthesis)
    
    Returns:
        List of arguments
    """
    args = []
    current = []
    depth = 0
    i = start_pos
    
    while i < len(formula):
        char = formula[i]
        
        if char == '(':
            depth += 1
            current.append(char)
        elif char == ')':
            if depth == 0:
                # Add last argument
                if current:
                    args.append(''.join(current))
                break
            else:
                depth -= 1
                current.append(char)
        elif char == ',' and depth == 0:
            # Argument separator
            args.append(''.join(current))
            current = []
        else:
            current.append(char)
        
        i += 1
    
    return args


def _add_function_prefixes(formula: str) -> str:
    """
    Add prefixes to Excel 365 new function names
    
    Args:
        formula: Formula string to process
    
    Returns:
        Formula with prefixes added
    """
    # Mapping of new functions and their prefixes
    function_map = _get_new_function_list()
    
    result = formula
    
    # Replace each function name
    for func_name, prefix in function_map.items():
        # Skip if already has the prefix for this specific function
        if prefix + func_name in result:
            continue
        
        # Use word boundaries to replace function names
        # Exclude identifiers with _xlpm./_xlop./_xleta.
        # Also exclude if already has _xlfn. prefix
        pattern = r'(?<!_xlpm\.)(?<!_xlop\.)(?<!_xleta\.)(?<!_xlfn\.)\b' + func_name + r'(?=\s*\()'
        replacement = prefix + func_name
        result = re.sub(pattern, replacement, result, flags=re.IGNORECASE)
    
    return result


def _get_new_function_list() -> Dict[str, str]:
    """
    Return mapping of Excel 365 new functions and their prefixes
    
    Returns:
        Dictionary of {function_name: prefix}
    """
    return {
        # Regular new functions (add _xlfn.)
        'UNIQUE': '_xlfn.',
        'SORTBY': '_xlfn.',
        'SEQUENCE': '_xlfn.',
        'RANDARRAY': '_xlfn.',
        'XLOOKUP': '_xlfn.',
        'XMATCH': '_xlfn.',
        'VSTACK': '_xlfn.',
        'HSTACK': '_xlfn.',
        'TAKE': '_xlfn.',
        'DROP': '_xlfn.',
        'CHOOSEROWS': '_xlfn.',
        'CHOOSECOLS': '_xlfn.',
        'EXPAND': '_xlfn.',
        'TOCOL': '_xlfn.',
        'TOROW': '_xlfn.',
        'WRAPCOLS': '_xlfn.',
        'WRAPROWS': '_xlfn.',
        'ARRAYTOTEXT': '_xlfn.',
        'VALUETOTEXT': '_xlfn.',
        'TEXTAFTER': '_xlfn.',
        'TEXTBEFORE': '_xlfn.',
        'TEXTSPLIT': '_xlfn.',
        'REGEXEXTRACT': '_xlfn.',
        'REGEXREPLACE': '_xlfn.',
        'REGEXTEST': '_xlfn.',
        'ISOMITTED': '_xlfn.',
        'MAP': '_xlfn.',
        'REDUCE': '_xlfn.',
        'SCAN': '_xlfn.',
        'BYCOL': '_xlfn.',
        'BYROW': '_xlfn.',
        'MAKEARRAY': '_xlfn.',
        'PERCENTOF': '_xlfn.',
        'TRIMRANGE': '_xlfn.',
        'LAMBDA': '_xlfn.',
        'LET': '_xlfn.',
        'GROUPBY': '_xlfn.',
        'PIVOTBY': '_xlfn.',
        
        # Special double prefix (add _xlfn._xlws.)
        'SORT': '_xlfn._xlws.',
        'FILTER': '_xlfn._xlws.',
    }


def _convert_tro_notations(formula: str) -> str:
    """
    Convert special cell range notations to corresponding function calls
    
    Args:
        formula: Formula string to process
    
    Returns:
        Converted formula string
    """
    result = formula
    
    # Cell reference pattern - support cells, columns, rows
    cell_pattern = r'(\$?[A-Z]+\$?\d+|\$?[A-Z]+|\$?\d+)'
    sheet_pattern = r'([A-Za-z_][\w\.]*!)?'
    
    # .:. -> _xlfn._TRO_ALL
    pattern_all = sheet_pattern + cell_pattern + r'\.:\.' + cell_pattern
    result = re.sub(pattern_all, lambda m: _convert_tro_match(m, '_xlfn._TRO_ALL'), result)
    
    # :. -> _xlfn._TRO_TRAILING  
    pattern_trailing = sheet_pattern + cell_pattern + r':\.' + cell_pattern
    result = re.sub(pattern_trailing, lambda m: _convert_tro_match(m, '_xlfn._TRO_TRAILING'), result)
    
    # .: -> _xlfn._TRO_LEADING
    pattern_leading = sheet_pattern + cell_pattern + r'\.:' + cell_pattern
    result = re.sub(pattern_leading, lambda m: _convert_tro_match(m, '_xlfn._TRO_LEADING'), result)
    
    return result


def _convert_tro_match(match: re.Match, func_name: str) -> str:
    """
    Convert TRO notation match to function call
    
    Args:
        match: Regular expression match object
        func_name: Target function name
    
    Returns:
        Converted function call string
    """
    groups = match.groups()
    
    # Get sheet name
    sheet = groups[0] if groups[0] else ''
    
    # Get cell references
    if len(groups) >= 3:
        cell1 = groups[1]
        cell2 = groups[2]
    else:
        # Return original string if pattern doesn't match
        return match.group(0)
    
    # Convert to function call format
    if sheet:
        # For sheet names, add sheet name to both cells
        sheet_name = sheet.rstrip('!')
        return f'{func_name}({sheet_name}!{cell1}:{sheet_name}!{cell2})'
    else:
        return f'{func_name}({cell1}:{cell2})'


def _determine_formula_attributes(formula: str, cell: Any) -> Dict[str, Any]:
    """
    Determine formula type and generate appropriate array formula attributes
    
    Args:
        formula: Processed formula string
        cell: Cell object where the formula will be set
    
    Returns:
        Attribute dictionary
    """
    # Determine if spill formula
    if hasattr(cell, '_is_spill') and cell._is_spill:
        # Get spill range
        spill_range = _get_spill_range(cell)
        return {'t': 'array', 'ref': spill_range}
    
    # Determine if contains LAMBDA/LET
    if '_xlfn.LAMBDA' in formula or '_xlfn.LET' in formula:
        # Get cell coordinate
        cell_ref = _get_cell_coordinate(cell)
        return {'t': 'array', 'ref': cell_ref}
    
    # Regular formula
    return {}


def _get_spill_range(cell: Any) -> str:
    """
    Get cell's spill range
    
    Args:
        cell: Cell containing spill formula
    
    Returns:
        String representation of spill range (e.g., "A1:C3")
    """
    # Get spill range from cell object
    if hasattr(cell, '_spill_range') and cell._spill_range is not None:
        return cell._spill_range
    
    # Default to cell's own coordinate
    return _get_cell_coordinate(cell)


def _get_cell_coordinate(cell: Any) -> str:
    """
    Get cell coordinate
    
    Args:
        cell: Cell object
    
    Returns:
        String representation of cell coordinate (e.g., "A1")
    """
    if hasattr(cell, 'coordinate'):
        return cell.coordinate
    
    # Default coordinate
    return "A1"