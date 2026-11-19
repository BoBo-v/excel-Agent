# -*- coding: utf-8 -*-
"""
复杂表格有效数据区边界探测器 (Table Boundary & Range Detector)
用于在包含大标题、副标题说明、空行及表尾附注的非标准 Sheet 中，
精准探测真实二维数据矩形区的主表头起始行、数据起始行与结束边界。
"""

from typing import List, Dict, Any, Tuple, Optional


def is_metadata_row(row: List[Any]) -> bool:
    """
    判断一行是否属于元数据/说明行（例如: '制表人: 张三', '导出日期: 2026-11-19'）。
    这类行往往夹在主标题与真实表头之间，单元格多含冒号或说明前缀。
    """
    non_empty = [c for c in row if c is not None and str(c).strip() != '']
    if not non_empty:
        return False

    colon_count = sum(1 for c in non_empty if (":" in str(c) or "：" in str(c)))
    # 若 40% 以上的非空单元格都带有冒号键值对特征，视为说明元数据行而非字段列名
    if colon_count / len(non_empty) >= 0.35:
        return True

    return False


def detect_data_boundary(
    rows: List[List[Any]],
    min_columns_threshold: int = 2
) -> Dict[str, Any]:
    """
    自动识别工作表中核心业务明细矩形区域的四界：
    - header_row: 真实主表头所在行号（从 1 起算）
    - data_start_row: 明细数据首行行号
    - data_end_row: 明细数据末行行号
    - start_col: 有效数据起始列标（从 1 起算）
    - end_col: 有效数据结束列标
    
    :param rows: 逐行单元格值列表
    :param min_columns_threshold: 判定为一个有效表格所需的最小连续非空列数
    """
    if not rows:
        return {
            "header_row": 1,
            "data_start_row": 2,
            "data_end_row": 1,
            "start_col": 1,
            "end_col": 1
        }

    # 1. 扫描每一行的非空列数与非空单元格分布
    row_stats = []
    max_non_empty_cols = 0

    for r_idx, row in enumerate(rows):
        non_empty_indices = [c_idx for c_idx, val in enumerate(row) if val is not None and str(val).strip() != '']
        col_count = len(non_empty_indices)
        if col_count > max_non_empty_cols:
            max_non_empty_cols = col_count
        row_stats.append({
            "row_num": r_idx + 1,
            "count": col_count,
            "indices": non_empty_indices,
            "raw_row": row
        })

    # 2. 定位真实主表头行：
    # 排除大标题（通常仅1个单元格）、说明元数据行（带冒号）以及空行，
    # 选取列数充沛且后续紧跟明细行的候选
    header_row = 1
    found_header = False

    for stat in row_stats:
        # 排除冒号键值对说明行
        if is_metadata_row(stat["raw_row"]):
            continue

        # 当列数达到最大列数的 75% 以上，且满足最小列数门槛
        if stat["count"] >= max(min_columns_threshold, int(max_non_empty_cols * 0.75)):
            header_row = stat["row_num"]
            found_header = True
            break

    if not found_header:
        header_row = 1

    # 3. 计算水平方向边界 (start_col, end_col)
    header_stat = row_stats[header_row - 1]
    if header_stat["indices"]:
        start_col = min(header_stat["indices"]) + 1
        end_col = max(header_stat["indices"]) + 1
    else:
        start_col = 1
        end_col = max(1, max_non_empty_cols)

    # 4. 计算数据终止行 (data_end_row)：
    # 从 header_row 往后遍历，遇到连续全空行或纯脚注行即为表尾
    data_start_row = header_row + 1
    data_end_row = data_start_row

    for stat in row_stats[header_row:]:
        # 若整行全空，判断可能已达表格底部
        if stat["count"] == 0:
            break
        data_end_row = stat["row_num"]

    if data_end_row < data_start_row:
        data_end_row = data_start_row

    return {
        "header_row": header_row,
        "data_start_row": data_start_row,
        "data_end_row": data_end_row,
        "start_col": start_col,
        "end_col": end_col
    }
