# Copyright (c) 2010-2024 openpyxl

"""
Excel 365新関数のプレフィックス処理

このモジュールは、Excel 365で導入された新関数（LAMBDA、LET等）に
必要なプレフィックス（_xlfn.、_xlpm.）を自動的に付与する機能を提供します。
"""

import re

# Excel 365以降の新関数（_xlfn.プレフィックスが必要）
EXCEL_NEW_FUNCTIONS = {
    'UNIQUE', 'SORT', 'SORTBY', 'FILTER', 'SEQUENCE', 
    'RANDARRAY', 'XLOOKUP', 'XMATCH',
    # フェーズ1: 基本的な配列操作関数（LETは別途対応予定）
    'VSTACK', 'HSTACK', 'TAKE', 'DROP',
    'CHOOSEROWS', 'CHOOSECOLS', 'EXPAND', 'TOCOL', 'TOROW',
    'WRAPCOLS', 'WRAPROWS',
    # フェーズ2: テキスト処理・正規表現関数
    'ARRAYTOTEXT', 'VALUETOTEXT', 'TEXTAFTER', 'TEXTBEFORE', 'TEXTSPLIT',
    'REGEXEXTRACT', 'REGEXREPLACE', 'REGEXTEST',
    # フェーズ4: LAMBDA関連
    'LAMBDA', 'LET',
    # フェーズ6: LAMBDAを使う各種関数
    'ISOMITTED', 'MAP', 'REDUCE', 'SCAN', 'BYCOL', 'BYROW', 'MAKEARRAY'
}


def add_function_prefix(formula_text):
    """
    通常の数式内の新関数に_xlfn.プレフィックスを追加する
    
    この関数は数式全体を処理し：
    1. LAMBDA/LET関数の特殊処理を_process_lambda_function関数に委譲
    2. その他のExcel 365新関数に_xlfn.プレフィックスを追加
    3. SORT/FILTER関数には追加で_xlws.プレフィックスも付与
    
    処理の順序が重要：
    - LAMBDA/LET関数を先に処理（引数に_xlpm.プレフィックスが必要なため）
    - その後、他の新関数を処理
    
    ## 処理ロジックの概略
    
    1. **関数名の検出**: 正規表現の単語境界（\b）を使用して関数名を正確に識別
    2. **プレフィックス追加**: 
       - 通常の新関数: 関数名の前に `_xlfn.` を追加
       - LAMBDA/LETのパラメータ: パラメータ名の前に `_xlpm.` を追加
       - SORT/FILTER: `_xlfn._xlws.` という二重プレフィックスを使用
    3. **括弧の対応処理**: カンマ区切りの引数解析時、括弧の入れ子を正しく処理
    4. **再帰的処理**: ネストしたLAMBDA関数は内側から外側へ処理
    
    ## 対応できていないケース
    
    1. **文字列内の関数名**:
       - 例: `=CONCATENATE("LAMBDA", "(x,x)")` → LAMBDAが誤って置換される
       - 文字列リテラル内の内容は考慮していない
    
    2. **コメント内の関数名**:
       - Excelの数式にコメントがある場合、その中の関数名も置換される可能性
    
    3. **極めて深いネスト**:
       - LAMBDA関数が10レベル以上ネストしている場合、処理が不完全になる可能性
    
    4. **不正な構文**:
       - 括弧の対応が取れていない数式では、予期しない動作をする可能性
       - 例: `=LAMBDA(x,x+1` （閉じ括弧なし）
    
    5. **動的な関数名**:
       - INDIRECT関数などで動的に生成される関数名は処理できない
       - 例: `=INDIRECT("LAM" & "BDA")`
    
    6. **名前付き範囲内の数式**:
       - ワークブックレベルの名前付き範囲に含まれる数式は、このモジュールでは処理されない
    
    呼び出しフロー:
    ```
    add_function_prefix (エントリーポイント)
      └── _process_lambda_function (LAMBDA/LET専用の処理)
            ├── _add_xlpm_to_lambda_params (LAMBDA引数の処理)
            └── _process_let_variables (LET変数の処理)
                  └── _add_xlpm_to_let_vars (LET引数の処理)
    ```
    
    外部から呼ばれる場所:
    - _prepare_spill_formula: スピル数式の処理時
    - etree_write_cell: 通常の数式書き込み時
    - lxml_write_cell: LXML使用時の数式書き込み時
    
    Args:
        formula_text: 数式テキスト（"=UPPER(TEXTBEFORE(B2,"@"))"）
    
    Returns:
        str: プレフィックスが追加された数式
    """
    if not formula_text or not formula_text.startswith('='):
        return formula_text
    
    # =を一時的に削除
    formula_without_eq = formula_text[1:]
    
    # 新関数にプレフィックスを追加
    
    # LAMBDA関数の特殊処理を先に行う
    formula_without_eq = _process_lambda_function(formula_without_eq)
    
    # その他の新関数にプレフィックスを追加
    for func in EXCEL_NEW_FUNCTIONS:
        if func in ['LAMBDA', 'LET']:  # LAMBDAとLETは既に処理済み
            continue
        # 単語境界を使って関数名を正確にマッチング
        pattern = r'\b' + func + r'\b'
        # _xlfn.または_xlws.が既に付いていない場合のみ追加
        if not re.search(r'(_xlfn\.|_xlws\.)' + func, formula_without_eq):
            formula_without_eq = re.sub(pattern, '_xlfn.' + func, formula_without_eq)
    
    # SORT関数とFILTER関数の特殊ケース（_xlwsプレフィックスが必要）
    # すでに_xlws.が付いている場合は追加しない
    if not re.search(r'_xlfn\._xlws\.SORT', formula_without_eq):
        formula_without_eq = re.sub(r'_xlfn\.SORT\b', '_xlfn._xlws.SORT', formula_without_eq)
    if not re.search(r'_xlfn\._xlws\.FILTER', formula_without_eq):
        formula_without_eq = re.sub(r'_xlfn\.FILTER\b', '_xlfn._xlws.FILTER', formula_without_eq)
    
    return '=' + formula_without_eq


def _process_lambda_function(formula_text):
    """
    LAMBDA関数とLET関数を処理し、適切なプレフィックスを追加する
    
    この関数は以下の処理を行う：
    1. LAMBDA → _xlfn.LAMBDA への変換
    2. LET → _xlfn.LET への変換
    3. 各LAMBDA関数の引数解析と_xlpm.プレフィックス付与
    4. LET関数の変数解析と_xlpm.プレフィックス付与
    
    Args:
        formula_text: 数式テキスト（'='なし）
    
    Returns:
        str: プレフィックスが追加された数式
    """
    # 処理中のテキスト
    processed_text = formula_text
    
    # LAMBDA関数を処理（最も内側から処理）
    # まず、すべてのLAMBDA関数の位置を見つける
    while True:
        lambda_positions = []
        # _xlfn.が付いていないLAMBDAのみを検索
        for match in re.finditer(r'(?<!_xlfn\.)LAMBDA\s*\(', processed_text):
            start = match.start()
            end = match.end()
            
            # 対応する閉じ括弧を見つける
            paren_count = 1
            pos = end
            in_string = False
            quote_char = None
            
            while pos < len(processed_text) and paren_count > 0:
                char = processed_text[pos]
                
                # 文字列リテラルの処理
                if char in ['"', "'"] and (pos == 0 or processed_text[pos-1] != '\\'):
                    if not in_string:
                        in_string = True
                        quote_char = char
                    elif char == quote_char:
                        in_string = False
                        quote_char = None
                
                # 文字列外での括弧のカウント
                if not in_string:
                    if char == '(':
                        paren_count += 1
                    elif char == ')':
                        paren_count -= 1
                pos += 1
            
            if paren_count == 0:
                lambda_positions.append((start, pos, pos - start))  # start, end, length
        
        if not lambda_positions:
            break
        
        # 長さでソート（短いものから処理 = 内側から処理）
        lambda_positions.sort(key=lambda x: x[2])
        
        # 最も内側のLAMBDAを処理
        start, end, _ = lambda_positions[0]
        lambda_expr = processed_text[start:end]
        processed_expr = _add_xlpm_to_lambda_params(lambda_expr)
        processed_text = processed_text[:start] + processed_expr + processed_text[end:]
    
    # LETを_xlfn.LETに置換
    processed_text = re.sub(r'\bLET\b', '_xlfn.LET', processed_text)
    
    # LET関数内の変数を処理
    processed_text = _process_let_variables(processed_text)
    
    return processed_text


def _protect_string_literals(text):
    """
    文字列リテラルと配列リテラル内の内容を保護するため、一時的に置換する
    
    Args:
        text: 処理対象のテキスト
    
    Returns:
        tuple: (保護済みテキスト, 保護したリテラルのリスト)
    """
    import re
    
    literals = []
    protected = text
    
    # 配列リテラル（波括弧内）を検出して保護
    array_pattern = r'\{[^\}]*\}'
    array_matches = list(re.finditer(array_pattern, text))
    
    # 後ろから置換（インデックスがずれないように）
    for match in reversed(array_matches):
        array_content = match.group(0)
        placeholder = f'__ARRAY_{len(literals)}__'
        literals.insert(0, array_content)
        protected = protected[:match.start()] + placeholder + protected[match.end():]
    
    # ダブルクォートで囲まれた文字列を検出
    string_pattern = r'"([^"]*)"'
    string_matches = list(re.finditer(string_pattern, protected))
    
    # 後ろから置換（インデックスがずれないように）
    for match in reversed(string_matches):
        string_content = match.group(0)
        placeholder = f'__STRING_{len(literals)}__'
        literals.insert(0, string_content)
        protected = protected[:match.start()] + placeholder + protected[match.end():]
    
    return protected, literals

def _restore_string_literals(text, literals):
    """
    保護した文字列リテラルと配列リテラルを元に戻す
    
    Args:
        text: 保護済みテキスト
        literals: 保護したリテラルのリスト
    
    Returns:
        str: 復元されたテキスト
    """
    restored = text
    for i, literal_content in enumerate(literals):
        # 文字列リテラルのプレースホルダ
        string_placeholder = f'__STRING_{i}__'
        restored = restored.replace(string_placeholder, literal_content)
        
        # 配列リテラルのプレースホルダ
        array_placeholder = f'__ARRAY_{i}__'
        restored = restored.replace(array_placeholder, literal_content)
    
    return restored

def _add_xlpm_to_lambda_params(lambda_expr):
    """
    LAMBDA式のパラメータに_xlpm.または_xlop.プレフィックスを追加
    
    この関数はLAMBDA関数の引数部分を解析し：
    1. カンマで区切られた引数リストを解析（括弧の入れ子を考慮）
    2. 最後の要素を式、それ以前をパラメータとして分離
    3. 各パラメータ名に適切なプレフィックスを追加
       - 通常のパラメータ: _xlpm.
       - オプショナルパラメータ（[]で囲まれた）: _xlop.
    4. 式内のパラメータ参照も同様に置換
    5. 式内の他の新関数（SEQUENCE等）にも適切なプレフィックスを付与
    
    Args:
        lambda_expr: "LAMBDA(x,[y],x+y)"のようなLAMBDA式
    
    Returns:
        str: "_xlfn.LAMBDA(_xlpm.x,_xlop.y,_xlpm.x+_xlpm.y)"
    """
    # LAMBDA(の後の内容を抽出
    match = re.match(r'(.*?LAMBDA\s*\()(.+)(\))', lambda_expr)
    if not match:
        return lambda_expr
    
    prefix = match.group(1)
    content = match.group(2)
    suffix = match.group(3)
    
    # カンマで分割（括弧、波括弧、文字列リテラル内のカンマは無視）
    parts = []
    current = ""
    paren_depth = 0
    brace_depth = 0
    in_string = False
    quote_char = None
    
    for i, char in enumerate(content):
        # 文字列リテラルの開始・終了を検出
        if char in ['"', "'"] and (i == 0 or content[i-1] != '\\'):
            if not in_string:
                in_string = True
                quote_char = char
            elif char == quote_char:
                in_string = False
                quote_char = None
        
        # 文字列リテラル内でない場合のみ括弧をカウント
        if not in_string:
            if char == '(':
                paren_depth += 1
            elif char == ')':
                paren_depth -= 1
            elif char == '{':
                brace_depth += 1
            elif char == '}':
                brace_depth -= 1
            elif char == ',' and paren_depth == 0 and brace_depth == 0:
                parts.append(current.strip())
                current = ""
                continue
        current += char
    
    if current:
        parts.append(current.strip())
    
    # LAMBDAは最後の要素が式、それ以前がパラメータ
    if len(parts) < 2:
        return lambda_expr
    
    params = parts[:-1]
    expression = parts[-1]
    
    # パラメータに適切なプレフィックスを追加
    processed_params = []
    param_names = []
    
    for param in params:
        # パラメータ名を抽出（空白を除去）
        param_str = param.strip()
        
        # オプショナルパラメータ（[]で囲まれた）かチェック
        if param_str.startswith('[') and param_str.endswith(']'):
            # オプショナルパラメータ
            param_name = param_str[1:-1].strip()
            processed_params.append('_xlop.' + param_name)
            param_names.append(param_name)
        else:
            # 通常のパラメータ
            param_name = param_str
            processed_params.append('_xlpm.' + param_name)
            param_names.append(param_name)
    
    # 式内のパラメータ参照も置換（文字列リテラル内は除外）
    processed_expr = _replace_var_refs_outside_strings(expression, param_names)
    
    # 式内の他の新関数にもプレフィックスを追加（既に_xlfn.が付いていない場合のみ）
    for func_name in EXCEL_NEW_FUNCTIONS:
        if func_name != 'LAMBDA' and func_name != 'LET':  # LAMBDAとLETは別処理
            processed_expr = re.sub(r'\b' + re.escape(func_name) + r'\(', '_xlfn.' + func_name + '(', processed_expr)
    
    # 再構築
    # prefixに既に_xlfn.が含まれているかチェック
    if '_xlfn.LAMBDA' in prefix:
        result = prefix + ','.join(processed_params + [processed_expr]) + suffix
    else:
        result = prefix.replace('LAMBDA', '_xlfn.LAMBDA') + ','.join(processed_params + [processed_expr]) + suffix
    return result


def _process_let_variables(formula_text):
    """
    LET関数内の変数定義と参照を処理
    
    この関数は数式内のすべてのLET関数を見つけて：
    1. LET関数の開始位置と対応する閉じ括弧を特定
    2. 各LET関数を_add_xlpm_to_let_vars関数で処理
    3. 処理済みのLET関数で元の式を置換
    
    括弧の入れ子を正しく処理し、ネストしたLET関数にも対応する。
    
    Args:
        formula_text: 数式テキスト
    
    Returns:
        str: 処理済みの数式
    """
    # LET関数を見つける
    let_pattern = r'_xlfn\.LET\s*\('
    
    # すべてのLET関数を処理
    result = formula_text
    start = 0
    
    while True:
        match = re.search(let_pattern, result[start:])
        if not match:
            break
        
        # LET関数の開始位置
        let_start = start + match.start()
        let_end = let_start + len('_xlfn.LET(')
        
        # 対応する閉じ括弧を見つける
        paren_count = 1
        pos = let_end
        
        while pos < len(result) and paren_count > 0:
            if result[pos] == '(':
                paren_count += 1
            elif result[pos] == ')':
                paren_count -= 1
            pos += 1
        
        if paren_count == 0:
            # LET関数全体を抽出
            let_expr = result[let_start:pos]
            # 処理
            processed = _add_xlpm_to_let_vars(let_expr)
            # 置換
            result = result[:let_start] + processed + result[pos:]
            start = let_start + len(processed)
        else:
            # 対応する括弧が見つからない場合はスキップ
            start = let_end
    
    return result


def _replace_var_refs_outside_strings(text, var_names):
    """
    文字列リテラル外の変数参照のみを置換する
    
    Args:
        text: 置換対象のテキスト
        var_names: 置換する変数名のリスト
    
    Returns:
        str: 変数参照が置換されたテキスト
    """
    if not var_names:
        return text
    
    result = []
    in_string = False
    quote_char = None
    i = 0
    
    while i < len(text):
        char = text[i]
        
        # 文字列リテラルの開始・終了を検出
        if char in ['"', "'"] and (i == 0 or text[i-1] != '\\'):
            if not in_string:
                in_string = True
                quote_char = char
                result.append(char)
                i += 1
                continue
            elif char == quote_char:
                in_string = False
                quote_char = None
                result.append(char)
                i += 1
                continue
        
        # 文字列リテラル内の場合はそのまま追加
        if in_string:
            result.append(char)
            i += 1
            continue
        
        # 文字列リテラル外で変数名をチェック
        matched = False
        for var_name in var_names:
            # 単語境界をチェック
            if text[i:i+len(var_name)] == var_name:
                # 前後が単語構成文字でないことを確認
                before_ok = i == 0 or not text[i-1].isalnum() and text[i-1] != '_'
                after_ok = i+len(var_name) >= len(text) or not text[i+len(var_name)].isalnum() and text[i+len(var_name)] != '_'
                
                if before_ok and after_ok:
                    result.append('_xlpm.' + var_name)
                    i += len(var_name)
                    matched = True
                    break
        
        if not matched:
            result.append(char)
            i += 1
    
    return ''.join(result)

def _add_xlpm_to_let_vars(let_expr):
    """
    LET式の変数に_xlpm.プレフィックスを追加
    
    この関数はLET関数の引数を解析し：
    1. カンマで区切られた引数リストを解析（括弧の入れ子を考慮）
    2. 変数名と値のペアを特定（最後の要素は最終式）
    3. 各変数名に_xlpm.プレフィックスを追加
    4. 変数の値内の既存変数参照も置換（前方参照のみ）
    5. 最終式内のすべての変数参照を置換
    
    例：LET(x,5,y,x+10,x+y) → LET(_xlpm.x,5,_xlpm.y,_xlpm.x+10,_xlpm.x+_xlpm.y)
    
    Args:
        let_expr: "_xlfn.LET(x,5,y,10,x+y)"のようなLET式
    
    Returns:
        str: "_xlfn.LET(_xlpm.x,5,_xlpm.y,10,_xlpm.x+_xlpm.y)"
    """
    # LET(の後の内容を抽出
    match = re.match(r'(_xlfn\.LET\s*\()(.+)(\)$)', let_expr)
    if not match:
        return let_expr
    
    prefix = match.group(1)
    content = match.group(2)
    suffix = match.group(3)
    
    # カンマで分割（括弧、波括弧、文字列リテラル内のカンマは無視）
    parts = []
    current = ""
    paren_depth = 0
    brace_depth = 0
    in_string = False
    quote_char = None
    
    for i, char in enumerate(content):
        # 文字列リテラルの開始・終了を検出
        if char in ['"', "'"] and (i == 0 or content[i-1] != '\\'):
            if not in_string:
                in_string = True
                quote_char = char
            elif char == quote_char:
                in_string = False
                quote_char = None
        
        # 文字列リテラル内でない場合のみ括弧をカウント
        if not in_string:
            if char == '(':
                paren_depth += 1
            elif char == ')':
                paren_depth -= 1
            elif char == '{':
                brace_depth += 1
            elif char == '}':
                brace_depth -= 1
            elif char == ',' and paren_depth == 0 and brace_depth == 0:
                parts.append(current.strip())
                current = ""
                continue
        current += char
    
    if current:
        parts.append(current.strip())
    
    # LETは変数名、値のペアが続き、最後に式
    if len(parts) < 3 or len(parts) % 2 == 0:
        return let_expr
    
    # 変数名を収集
    var_names = []
    processed_parts = []
    
    # 変数定義部分を処理（最後の要素以外）
    for i in range(0, len(parts) - 1, 2):
        var_name = parts[i].strip()
        var_value = parts[i + 1] if i + 1 < len(parts) - 1 else ""
        
        var_names.append(var_name)
        processed_parts.append('_xlpm.' + var_name)
        
        if var_value:
            # 値の中の既存の変数参照も置換（文字列リテラル内は除外）
            processed_value = _replace_var_refs_outside_strings(var_value, var_names[:-1])
            processed_parts.append(processed_value)
    
    # 最後の式を処理
    final_expr = parts[-1]
    final_expr = _replace_var_refs_outside_strings(final_expr, var_names)
    processed_parts.append(final_expr)
    
    # 再構築
    result = prefix + ','.join(processed_parts) + suffix
    return result