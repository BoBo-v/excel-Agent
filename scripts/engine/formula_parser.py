import re

def col_to_num(col_str: str) -> int:
    """Converts Excel column letters to 1-based index: A->1, Z->26, AA->27."""
    num = 0
    for c in col_str.upper():
        num = num * 26 + (ord(c) - ord('A') + 1)
    return num

def num_to_col(num: int) -> str:
    """Converts 1-based index to Excel column letters: 1->A, 26->Z, 27->AA."""
    col = ''
    while num > 0:
        num, r = divmod(num - 1, 26)
        col = chr(65 + r) + col
    return col

# Matches Excel cell references like A1, $B$2, AA10, $C4, etc.
CELL_REF_PATTERN = re.compile(r'(?<![A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])')

def shift_formula_references(formula_str: str, inserted_col_idx: int, num_cols: int = 1) -> str:
    """
    Intelligently shifts all column references in a formula string if the column >= inserted_col_idx.
    
    Example:
        shift_formula_references("=(L2-N2)/L2", inserted_col_idx=11, num_cols=1)
        -> "=(M2-O2)/M2"
    """
    if not formula_str or not str(formula_str).startswith('='):
        return formula_str

    def _replace(match):
        col_abs, col_letters, row_abs, row_num = match.groups()
        # Avoid false positives like R1C1 if applicable, but standard A1 notation is checked
        col_idx = col_to_num(col_letters)
        
        # If absolute reference $A1 or if >= inserted_col_idx, shift the column
        if col_idx >= inserted_col_idx:
            new_col = num_to_col(col_idx + num_cols)
            return f"{col_abs}{new_col}{row_abs}{row_num}"
        return match.group(0)

    return CELL_REF_PATTERN.sub(_replace, formula_str)
