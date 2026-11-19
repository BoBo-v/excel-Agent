# -*- coding: utf-8 -*-
"""
Excel 智能体安全与公式注入防护引擎 (Formula Injection & Prompt Injection Guard)
防御 CWE-1236 (CSV/Formula Injection) 及单元格内恶意 Prompt 越狱指令。
"""

import re
from typing import Tuple, Any

# 危险外部调用及远程命令执行函数名
DANGEROUS_FUNCTIONS = (
    "WEBSERVICE", "ENCODEURL", "DDE", "CALL", "EXEC", "SHELL", "HYPERLINK"
)

# 提示词注入特征（中英文）
PROMPT_INJECTION_PATTERNS = [
    re.compile(r"ignore\s+(all\s+)?(previous|prior)\s+instructions", re.IGNORECASE),
    re.compile(r"disregard\s+(all\s+)?(previous|prior)\s+instructions", re.IGNORECASE),
    re.compile(r"system\s+(override|prompt|instructions)", re.IGNORECASE),
    re.compile(r"(忽略|无视)(之前|上述|所有|全部)的?(指令|要求|提示)", re.IGNORECASE),
    re.compile(r"(删除|清空)(所有|全部)(数据|文件|表格)", re.IGNORECASE),
]


def is_dangerous_formula(formula_str: str) -> Tuple[bool, str]:
    """
    检查公式字符串是否包含恶意外部网络请求（如 WEBSERVICE 外带窃密）
    或系统命令调用（DDE/CALL/EXEC）。
    
    :return: (is_dangerous, reason)
    """
    if not isinstance(formula_str, str):
        return False, ""

    s = formula_str.strip().upper()
    if not s.startswith("="):
        return False, ""

    for fn in DANGEROUS_FUNCTIONS:
        # 匹配函数调用如 WEBSERVICE( 或 _xlfn.WEBSERVICE(
        pattern = re.compile(rf"(?:_xlfn\.)?\b{fn}\b\s*\(")
        if pattern.search(s):
            return True, f"公式包含高危外部交互函数: {fn}"

    return False, ""


def sanitize_plain_text(val: Any) -> Any:
    """
    对作为普通文本存储的单元格数据进行脱敏防注入处理。
    如果文本以 '=', '+', '-', '@', '\t', '\r' 开头，可能被 Excel 误当成公式执行，
    在最前部添加单引号 "'" 进行安全转义（CWE-1236 防护）。
    """
    if not isinstance(val, str):
        return val

    s = val.strip()
    if not s:
        return val

    # 如果以危险计算前缀开头且不是合法的正常数字（如 -123.45 是合法的负数）
    if s[0] in ('=', '+', '-', '@'):
        # 排除合法负数或正数
        try:
            float(s)
            return val
        except ValueError:
            # 是包含公式前缀的非数字文本，做单引号转义
            return f"'{val}"

    return val


def detect_prompt_injection(cell_content: Any) -> Tuple[bool, str]:
    """
    检测单元格数据中是否混入了试图破坏 Agent 边界的恶意提示词注入。
    
    :return: (is_suspicious, match_desc)
    """
    if not isinstance(cell_content, str):
        return False, ""

    for pattern in PROMPT_INJECTION_PATTERNS:
        m = pattern.search(cell_content)
        if m:
            return True, f"命中可疑提示词注入特征: '{m.group(0)}'"

    return False, ""
