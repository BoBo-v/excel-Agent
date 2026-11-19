# -*- coding: utf-8 -*-
"""
工作表生命周期管理器 (Sheet Lifecycle Manager)
提供完整的 Excel 工作表日常基础增删改查、复制、重命名与隐藏/显示管理能力，
包含防损坏校验与自动快照保护。
"""

import os
import openpyxl
from typing import List, Dict, Any, Optional
from .safe_io import create_backup, is_file_locked


def list_sheets(file_path: str) -> List[Dict[str, Any]]:
    """
    获取工作簿中所有工作表的基础信息清单。
    
    :return: [{"index": 0, "name": "Sheet1", "state": "visible", "max_row": 100, "max_column": 15}, ...]
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    wb = openpyxl.load_workbook(file_path, read_only=True)
    results = []
    for idx, sheet_name in enumerate(wb.sheetnames):
        ws = wb[sheet_name]
        state = getattr(ws, "sheet_state", "visible")
        results.append({
            "index": idx,
            "name": sheet_name,
            "state": state,
            "max_row": ws.max_row,
            "max_column": ws.max_column
        })
    wb.close()
    return results


def create_sheet(
    file_path: str,
    sheet_name: str,
    index: Optional[int] = None,
    output_path: Optional[str] = None
) -> str:
    """
    在工作簿中新建一个指定名称的工作表。
    """
    if output_path is None:
        output_path = file_path
        create_backup(file_path)

    wb = openpyxl.load_workbook(file_path)
    if sheet_name in wb.sheetnames:
        wb.close()
        raise ValueError(f"工作表 '{sheet_name}' 已存在，无法重复创建")

    wb.create_sheet(title=sheet_name, index=index)
    wb.save(output_path)
    wb.close()
    return output_path


def rename_sheet(
    file_path: str,
    old_name: str,
    new_name: str,
    output_path: Optional[str] = None
) -> str:
    """
    安全重命名指定的工作表。
    """
    if output_path is None:
        output_path = file_path
        create_backup(file_path)

    wb = openpyxl.load_workbook(file_path)
    if old_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"原工作表 '{old_name}' 不存在")
    if new_name in wb.sheetnames and old_name != new_name:
        wb.close()
        raise ValueError(f"目标工作表名称 '{new_name}' 已存在")

    ws = wb[old_name]
    ws.title = new_name
    wb.save(output_path)
    wb.close()
    return output_path


def copy_sheet(
    file_path: str,
    source_name: str,
    new_name: str,
    output_path: Optional[str] = None
) -> str:
    """
    复制现有工作表及其全部单元格内容与格式到新工作表。
    """
    if output_path is None:
        output_path = file_path
        create_backup(file_path)

    wb = openpyxl.load_workbook(file_path)
    if source_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"来源工作表 '{source_name}' 不存在")
    if new_name in wb.sheetnames:
        wb.close()
        raise ValueError(f"目标复制表名 '{new_name}' 已存在")

    src_ws = wb[source_name]
    new_ws = wb.copy_worksheet(src_ws)
    new_ws.title = new_name
    wb.save(output_path)
    wb.close()
    return output_path


def delete_sheet(
    file_path: str,
    sheet_name: str,
    output_path: Optional[str] = None
) -> str:
    """
    安全删除指定的工作表（确保工作簿中至少保留一个可见工作表）。
    """
    if output_path is None:
        output_path = file_path
        create_backup(file_path)

    wb = openpyxl.load_workbook(file_path)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"工作表 '{sheet_name}' 不存在")

    if len(wb.sheetnames) <= 1:
        wb.close()
        raise ValueError("工作簿中必须保留至少一张工作表，禁止全部删除")

    ws = wb[sheet_name]
    wb.remove(ws)
    wb.save(output_path)
    wb.close()
    return output_path


def set_sheet_visibility(
    file_path: str,
    sheet_name: str,
    visible: bool = True,
    output_path: Optional[str] = None
) -> str:
    """
    设置工作表的可见性状态（显示 / 隐藏）。
    """
    if output_path is None:
        output_path = file_path
        create_backup(file_path)

    wb = openpyxl.load_workbook(file_path)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"工作表 '{sheet_name}' 不存在")

    ws = wb[sheet_name]
    ws.sheet_state = "visible" if visible else "hidden"
    wb.save(output_path)
    wb.close()
    return output_path
