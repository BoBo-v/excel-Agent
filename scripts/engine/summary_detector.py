# -*- coding: utf-8 -*-
"""
汇总行与小计行识别引擎 (Summary & Subtotal Detector)
用于在行过滤、数据分析和聚合计算中，智能识别末尾或中间的“合计/小计/Total”行，
防止汇总行被误当成普通明细数据参与计算导致翻倍（+161% 典型错误），或在过滤时被错误剔除。
"""

import re
from typing import List, Optional, Set, Any

# 常见汇总与小计关键词集合（中文与英文）
SUMMARY_KEYWORDS = {
    "合计", "总计", "小计", "总和", "累计", "汇总", "求和",
    "平均", "均值", "共计", "总额", "结计", "本期合计",
    "total", "subtotal", "sum", "avg", "average", "grand total", "summary"
}

# 脚注或说明前缀（通常不属于业务明细）
FOOTNOTE_PREFIXES = (
    "注：", "注:", "备注", "说明：", "说明:", "数据来源", "来源：", "制表人"
)


def clean_cell_text(val: Any) -> str:
    """提取并清理单元格文本，去除空格与不可见字符。"""
    if val is None:
        return ""
    return str(val).replace('\xa0', ' ').replace('\u3000', ' ').strip()


def is_summary_text(text: str) -> bool:
    """
    判断文本是否包含汇总、小计特征。
    例如："华东小计", "1月合计", "总计", "Total Amount", "Grand Total"
    """
    clean = clean_cell_text(text)
    if not clean:
        return False

    clean_lower = clean.lower()

    # 1. 精确匹配
    if clean_lower in SUMMARY_KEYWORDS:
        return True

    # 2. 包含中文汇总关键词（长度通常较短，如“华东小计”、“合计(元)”）
    if len(clean) <= 12:
        for kw in ("合计", "总计", "小计", "总和", "累计", "汇总", "求和", "共计"):
            if kw in clean:
                return True

    # 3. 英文关键词匹配
    for en_kw in ("grand total", "subtotal", "total", "average"):
        if clean_lower.startswith(en_kw) or clean_lower.endswith(en_kw) or f" {en_kw} " in f" {clean_lower} ":
            return True

    return False


def is_footnote_text(text: str) -> bool:
    """判断是否为表尾说明或脚注行。"""
    clean = clean_cell_text(text)
    if not clean:
        return False
    return any(clean.startswith(prefix) for prefix in FOOTNOTE_PREFIXES)


def is_summary_row(row_values: List[Any], max_check_cols: int = 5) -> bool:
    """
    判断整行是否属于汇总/小计/说明行。
    通常检查前几个非空单元格中是否包含汇总标识词。
    
    :param row_values: 当前行各单元格的值列表
    :param max_check_cols: 重点检查的前 N 个关键列
    :return: True 如果判定为汇总/小计行
    """
    checked = 0
    for val in row_values:
        text = clean_cell_text(val)
        if not text:
            continue
        
        if is_summary_text(text) or is_footnote_text(text):
            return True
        
        checked += 1
        if checked >= max_check_cols:
            break
            
    return False
