# -*- coding: utf-8 -*-
"""
原生联动公式构建器 (Native Excel Formula Builder)
用于在表格操作中生成标准、联动、绝对引用锁定正确且兼具容错的原生 Excel 公式，
杜绝直接写入硬编码死数据。
"""

from typing import List, Tuple, Optional


def format_range_address(
    col_letter: str,
    start_row: int,
    end_row: int,
    sheet_name: Optional[str] = None,
    absolute: bool = True
) -> str:
    """
    格式化生成区域地址，如 '$B$2:$B$100' 或 'Sheet1!$B$2:$B$100'。
    """
    prefix = f"'{sheet_name}'!" if sheet_name and (" " in sheet_name or "-" in sheet_name) else (f"{sheet_name}!" if sheet_name else "")
    dollar = "$" if absolute else ""
    return f"{prefix}{dollar}{col_letter.upper()}{dollar}{start_row}:{dollar}{col_letter.upper()}{dollar}{end_row}"


def build_xlookup(
    lookup_val_cell: str,
    lookup_range: str,
    return_range: str,
    if_not_found: str = '""'
) -> str:
    """
    构建现代原生 XLOOKUP 公式。
    例如: =XLOOKUP(A2, '数据源'!$B$2:$B$100, '数据源'!$C$2:$C$100, "")
    """
    return f"=XLOOKUP({lookup_val_cell}, {lookup_range}, {return_range}, {if_not_found})"


def build_index_match(
    lookup_val_cell: str,
    lookup_range: str,
    return_range: str
) -> str:
    """
    构建全版本通用的 INDEX + MATCH 查找公式（兼容旧版 Excel 2010/2013/2016 及 WPS）。
    例如: =INDEX('数据源'!$C$2:$C$100, MATCH(A2, '数据源'!$B$2:$B$100, 0))
    """
    return f"=INDEX({return_range}, MATCH({lookup_val_cell}, {lookup_range}, 0))"


def build_sumifs(
    sum_range: str,
    criteria_pairs: List[Tuple[str, str]]
) -> str:
    """
    构建多条件求和 SUMIFS 公式。
    例如: =SUMIFS($C$2:$C$100, $A$2:$A$100, "华东", $B$2:$B$100, ">=100")
    
    :param sum_range: 求和数值区域，如 '$C$2:$C$100'
    :param criteria_pairs: 条件对列表 [(条件区域1, 匹配条件1), (条件区域2, 匹配条件2)]
    """
    if not criteria_pairs:
        return f"=SUM({sum_range})"

    parts = [sum_range]
    for c_range, c_val in criteria_pairs:
        parts.append(c_range)
        parts.append(str(c_val))

    args_str = ", ".join(parts)
    return f"=SUMIFS({args_str})"


def build_iferror(inner_formula: str, fallback_val: str = '""') -> str:
    """
    用 IFERROR 安全包裹已有公式，防出现 #N/A 或 #VALUE! 破坏展示。
    例如: 将 '=VLOOKUP(...)' 转为 '=IFERROR(VLOOKUP(...), "")'
    """
    clean_inner = inner_formula.strip()
    if clean_inner.startswith("="):
        clean_inner = clean_inner[1:].strip()
    return f"=IFERROR({clean_inner}, {fallback_val})"
