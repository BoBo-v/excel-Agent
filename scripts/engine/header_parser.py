# -*- coding: utf-8 -*-
"""
多级表头解析与扁平化引擎 (Multi-level Header Parser & Flattener)
用于解决复杂中国企业报表中：双层/三层表头、合并单元格表头导致的字段混乱，
自动向下对齐并拼接扁平化列名（如 '一季度_1月'），准确定位真实明细数据起始行。
"""

import re
from typing import List, Optional, Any


def clean_header_text(val: Any) -> str:
    """清理表头单个单元格文本，去除换行符、不可见全角空格及两端留白。"""
    if val is None:
        return ""
    s = str(val).replace('\xa0', ' ').replace('\u3000', ' ').replace('\r', '').replace('\n', ' ')
    return re.sub(r'\s+', ' ', s).strip()


def flatten_multi_level_headers(
    header_rows: List[List[Any]],
    sep: str = "_"
) -> List[str]:
    """
    将多行表头矩阵（如 2~3 行复合表头）向下向前填充合并单元格，
    并以连接符拼接为扁平化列名。
    
    例如：
    Row 1: ["基本信息", None,   "第一季度", None]
    Row 2: ["商品编码", "商品名称", "1月",   "2月"]
    => 输出: ["基本信息_商品编码", "基本信息_商品名称", "第一季度_1月", "第一季度_2月"]
    
    :param header_rows: 表头行矩阵（每一项为一个行列表）
    :param sep: 父子级表头连接符（默认 '_'）
    :return: 扁平化后的唯一列名列表
    """
    if not header_rows:
        return []

    num_rows = len(header_rows)
    num_cols = max(len(r) for r in header_rows)

    # 1. 规范化为等长二维矩阵并初步清理
    matrix: List[List[str]] = []
    for r in header_rows:
        row_clean = [clean_header_text(r[i]) if i < len(r) else "" for i in range(num_cols)]
        matrix.append(row_clean)

    # 2. 对每一行执行从左向右的向前填充（处理跨列合并单元格产生的空白）
    filled_matrix: List[List[str]] = []
    for r in matrix:
        filled_row: List[str] = []
        last_val = ""
        for cell in r:
            if cell:
                last_val = cell
                filled_row.append(cell)
            else:
                filled_row.append(last_val)
        filled_matrix.append(filled_row)

    # 3. 从上到下拼接各层表头，去重重复词汇
    flattened_headers: List[str] = []
    for col_idx in range(num_cols):
        parts: List[str] = []
        for row_idx in range(num_rows):
            val = filled_matrix[row_idx][col_idx]
            if val and (not parts or parts[-1] != val):
                parts.append(val)
        
        if not parts:
            final_name = f"未命名列_{col_idx + 1}"
        else:
            final_name = sep.join(parts)
            
        flattened_headers.append(final_name)

    # 4. 确保最终列名唯一性（若存在重名，追加序号）
    seen = {}
    unique_headers: List[str] = []
    for h in flattened_headers:
        if h not in seen:
            seen[h] = 1
            unique_headers.append(h)
        else:
            seen[h] += 1
            unique_headers.append(f"{h}_{seen[h]}")

    return unique_headers


def detect_header_depth(sample_rows: List[List[Any]], max_check_rows: int = 4) -> int:
    """
    启发式探测表格的表头层数（1行、2行还是3行）。
    通过检查前几行单元格中包含大量字符串而非数值、以及存在连续空白合并特征进行判定。
    """
    if not sample_rows:
        return 1

    check_limit = min(len(sample_rows), max_check_rows)
    header_depth = 1

    for r_idx in range(1, check_limit):
        row = sample_rows[r_idx]
        # 计算该行字符串单元格占比
        non_empty = [c for c in row if c is not None and str(c).strip() != '']
        if not non_empty:
            continue
        
        str_count = 0
        for c in non_empty:
            s = str(c).strip()
            # 排除纯数字或百分比
            try:
                float(s.replace('%', '').replace(',', ''))
            except ValueError:
                str_count += 1

        # 若当前行绝大多数（>= 70%）依然是文本标题，强烈提示是子表头
        if (str_count / len(non_empty)) >= 0.70:
            header_depth = r_idx + 1
        else:
            break

    return header_depth
