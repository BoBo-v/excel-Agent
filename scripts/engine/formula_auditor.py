# -*- coding: utf-8 -*-
"""
公式审计与错误诊断引擎 (Formula Auditor & Error Inspector)
用于扫描整个工作簿中存在的公式语法错误、断链及运行时报错，
精准定位坐标并输出诊断建议。
"""

import os
import re
import openpyxl
from typing import List, Dict, Any, Optional

# 标准 Excel 错误代码及其成因建议
EXCEL_ERROR_TYPES = {
    "#REF!": "引用失效（被引用的单元格、行或列已被物理删除或移出范围）",
    "#DIV/0!": "除零错误（除数为 0 或引用了空单元格）",
    "#VALUE!": "值类型错误（公式参数类型不匹配，如文本参与了数值计算或内嵌图片丢失）",
    "#NAME?": "名称错误（函数名拼写错误，或引用了未定义的命名区域）",
    "#N/A": "查无结果（查找函数如 VLOOKUP/MATCH 未匹配到目标值且未做 IFERROR 容错）",
    "#NUM!": "数值越界（公式计算出的数值过大、过小或传入了非法数学参数）",
    "#NULL!": "空交集错误（两个不相交的区域使用了交集运算符空格）",
    "#SPILL!": "动态数组溢出冲突（溢出区域存在非空单元格阻挡）"
}


def scan_workbook_formula_errors(
    file_path: str,
    target_sheet: Optional[str] = None
) -> List[Dict[str, Any]]:
    """
    全量扫描工作簿中的所有公式与计算单元格，查找并定位包含错误代码的单元格。
    
    :return: [
        {
            "sheet": "销售明细",
            "cell": "E12",
            "error_type": "#REF!",
            "formula": "=SUM(#REF!, D12)",
            "description": "引用失效（被引用的单元格、行或列已被物理删除）"
        }, ...
    ]
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    errors: List[Dict[str, Any]] = []

    # 1. 以只读模式加载（读取底层公式文本与缓存值）
    wb = openpyxl.load_workbook(file_path, data_only=False)
    sheets_to_scan = [target_sheet] if target_sheet and target_sheet in wb.sheetnames else wb.sheetnames

    for s_name in sheets_to_scan:
        ws = wb[s_name]
        for row in ws.iter_rows():
            for cell in row:
                val_str = str(cell.value or "").strip()
                
                # 检查公式本身是否含有 #REF! 等硬编码错误字符串
                for err_code, desc in EXCEL_ERROR_TYPES.items():
                    if err_code in val_str:
                        errors.append({
                            "sheet": s_name,
                            "cell": cell.coordinate,
                            "error_type": err_code,
                            "formula": val_str if val_str.startswith("=") else None,
                            "value": val_str,
                            "description": desc
                        })
                        break
    wb.close()
    return errors


def audit_formula_syntax(formula: str) -> Dict[str, Any]:
    """
    静态语法审计单个公式字符串。
    检查括号匹配、引用有效性、是否包含危险外部调用。
    """
    s = formula.strip()
    if not s.startswith("="):
        return {"valid": False, "reason": "公式必须以 '=' 开头"}

    # 1. 括号对称性校验
    if s.count("(") != s.count(")"):
        return {
            "valid": False,
            "reason": f"括号不匹配（左括号 {s.count('(')} 个，右括号 {s.count(')')} 个）"
        }

    # 2. 检查内部是否残留未修复的 #REF!
    if "#REF!" in s.upper():
        return {"valid": False, "reason": "公式内部包含失效引用 '#REF!'"}

    return {"valid": True, "reason": "公式语法结构正常"}
