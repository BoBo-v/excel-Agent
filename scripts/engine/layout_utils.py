# -*- coding: utf-8 -*-
"""
单元格排版与自适应列宽工具 (Layout & Auto-fit Column Widths)
用于计算单元格中英文混合文本的最佳列宽，防止数字在 Excel 打开时显示为 '###' 或表头文字被截断。
"""

import unicodedata
from typing import List, Dict, Any, Optional


def get_string_display_width(text: Any) -> float:
    """
    计算字符串在 Excel 中的视觉显示宽度：
    - 半角英文字母、数字、符号计为 1.0 宽度
    - 全角字符、中文字符（East Asian Wide / Fullwidth）计为 2.0 宽度
    """
    if text is None:
        return 0.0

    s = str(text)
    # 取换行符前最长的一行
    lines = s.split('\n')
    max_line_width = 0.0

    for line in lines:
        line_w = 0.0
        for ch in line:
            # 判断字符是半角还是全角/宽字符
            east_asian = unicodedata.east_asian_width(ch)
            if east_asian in ('W', 'F'):
                line_w += 2.0
            else:
                line_w += 1.0
        if line_w > max_line_width:
            max_line_width = line_w

    return max_line_width


def calculate_auto_column_widths(
    headers: List[str],
    rows: List[List[Any]],
    min_width: float = 10.0,
    max_width: float = 65.0,
    padding: float = 3.0
) -> Dict[int, float]:
    """
    计算各列推荐列宽字典 (列索引从 1 开始 -> 推荐宽度)。
    
    :param headers: 表头文字列表
    :param rows: 数据行列表
    :param min_width: 最小列宽（默认 10.0，足够展示常规数值或日期）
    :param max_width: 最大列宽上限（默认 65.0，防止长段落拉垮表格）
    :param padding: 额外安全留白（默认 +3.0，为 Excel 筛选下拉箭头及边框留出空间）
    :return: {1: 15.5, 2: 24.0, ...}
    """
    col_widths: Dict[int, float] = {}
    num_cols = len(headers)

    for i in range(num_cols):
        col_idx = i + 1
        # 表头自身宽度
        header_text = headers[i] if i < len(headers) else ""
        max_w = get_string_display_width(header_text)

        # 采样数据行宽度（最多采样前 200 行以保证性能）
        sample_rows = rows[:200]
        for row in sample_rows:
            if i < len(row):
                val_w = get_string_display_width(row[i])
                if val_w > max_w:
                    max_w = val_w

        # 加上留白并应用上下限
        final_w = max(min_width, min(max_w + padding, max_width))
        col_widths[col_idx] = round(final_w, 2)

    return col_widths
