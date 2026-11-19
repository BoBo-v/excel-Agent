# -*- coding: utf-8 -*-
"""
数据清洗与规范化引擎 (Data Cleaner & Normalizer)
用于解决业务表格中的常见脏数据缺陷：
1. 假数字转换（去除千分位逗号、货币符号、中文'元'、括号负数、百分号转小数）；
2. 混杂日期规范化（自动处理 Excel 序列号 45615、带'年月日'中文日期、斜杠与点号日期转 ISO 格式）；
3. 文本分列拆分（按逗号、分号、斜杠精准拆分清洗）。
"""

import re
from datetime import datetime, date, timedelta
from typing import Any, Optional, List, Union


# 常见缺失值/占位符集合
NA_VALUES = {"", "—", "-", "－", "无", "n/a", "na", "null", "none", "nan", "#n/a", "#value!"}


def clean_numeric_value(val: Any) -> Optional[float]:
    """
    将文本型假数字及脏数据清洗为纯浮点数值。
    
    支持清洗：
    - 带千分位逗号: "1,234,567.89" -> 1234567.89
    - 带货币/单位: "¥ 59.90 元", "$120.00", "RMB 88.5" -> 对应数值
    - 百分比: "18.5%" -> 0.185
    - 会计用括号负数: "(150.00)" -> -150.0
    - 常见占位符: "—", "N/A", "无" -> None
    """
    if val is None:
        return None

    if isinstance(val, (int, float)):
        import math
        return None if math.isnan(val) else float(val)

    s = str(val).strip()
    if not s or s.lower() in NA_VALUES:
        return None

    # 1. 处理括号负数，如 (123.45)
    is_negative = False
    if s.startswith('(') and s.endswith(')'):
        is_negative = True
        s = s[1:-1].strip()
    elif s.startswith('（') and s.endswith('）'):
        is_negative = True
        s = s[1:-1].strip()

    # 2. 检查是否为百分数
    is_percent = s.endswith('%') or s.endswith('％')
    if is_percent:
        s = s[:-1].strip()

    # 3. 移除货币符号与单位
    s = re.sub(r'[¥$€£\s,元RMBrmb]', '', s)

    try:
        num = float(s)
        if is_percent:
            num = num / 100.0
        if is_negative:
            num = -abs(num)
        return num
    except ValueError:
        return None


def normalize_date_value(val: Any) -> Optional[str]:
    """
    将多种混杂的日期表达形式统一标准化为 ISO 字符串 'YYYY-MM-DD'。
    
    支持处理：
    - 原生 datetime/date 对象
    - Excel 日期数字序列号 (如 45615 -> '2024-11-19')
    - 中文年月日格式 (如 '2026年11月19日', '2026年5月1日')
    - 常见分隔符 ('2026/11/19', '2026.11.19', '2026-11-19')
    - 带时间的字符串 (如 '2026-11-19 14:30:00' -> '2026-11-19')
    """
    if val is None:
        return None

    if isinstance(val, datetime):
        return val.strftime('%Y-%m-%d')
    if isinstance(val, date):
        return val.strftime('%Y-%m-%d')

    # 处理 Excel 日期序列号（通常为大于 1000 的数字）
    if isinstance(val, (int, float)):
        try:
            # Excel 基准日期 1899-12-30 (考虑 Lotus 1900 闰年 Bug)
            epoch = datetime(1899, 12, 30)
            target_date = epoch + timedelta(days=float(val))
            if 1970 <= target_date.year <= 2100:
                return target_date.strftime('%Y-%m-%d')
        except (ValueError, OverflowError):
            pass

    s = str(val).strip()
    if not s or s.lower() in NA_VALUES:
        return None

    # 如果是纯数字字符串且在序列号范围内
    if s.isdigit() and len(s) == 5:
        try:
            epoch = datetime(1899, 12, 30)
            target_date = epoch + timedelta(days=int(s))
            if 1970 <= target_date.year <= 2100:
                return target_date.strftime('%Y-%m-%d')
        except (ValueError, OverflowError):
            pass

    # 1. 替换中文年月日为短横线
    s_clean = s.replace('年', '-').replace('月', '-').replace('日', '').strip()
    # 替换斜杠和点号为短横线
    s_clean = s_clean.replace('/', '-').replace('.', '-')
    
    # 截取前面的年月日部分
    date_part = s_clean.split(' ')[0].strip()

    # 正则匹配 YYYY-MM-DD
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})$', date_part)
    if m:
        year, month, day = m.groups()
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"

    return None


def split_text_column(
    text: Any,
    delimiter: str = ",",
    max_split: int = -1
) -> List[str]:
    """
    将单元格文本按指定分隔符（逗号、分号、竖线等）精准分列，
    并自动剔除前后空白，支持全角/半角标点自适应。
    """
    if text is None:
        return []

    s = str(text).strip()
    if not s:
        return []

    # 自适应中英文常见分隔符
    if delimiter in (',', '，'):
        pattern = r'[,，]'
    elif delimiter in (';', '；'):
        pattern = r'[;；]'
    elif delimiter in ('|', '｜'):
        pattern = r'[|｜]'
    else:
        pattern = re.escape(delimiter)

    if max_split > 0:
        parts = re.split(pattern, s, maxsplit=max_split)
    else:
        parts = re.split(pattern, s)

    return [p.strip() for p in parts]
