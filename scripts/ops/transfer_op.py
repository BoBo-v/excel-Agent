import os
import openpyxl
import pandas as pd
from typing import List, Optional, Dict
from collections import defaultdict
from engine import (
    detect_excel_flavor,
    shift_formula_references,
    CascadingMatcher,
    auto_detect_key_columns,
    is_file_locked,
    create_backup,
    atomic_replace_file,
    WpsImageEngine,
    MsOfficeImageEngine,
    XmlStreamPatcher
)

def run_transfer(
    source_path: str,
    target_path: str,
    output_path: Optional[str] = None,
    sheets: Optional[List[str]] = None,
    cols: Optional[List[str]] = None,
    keys: Optional[List[str]] = None,
    insert_after: Optional[str] = None,
    branch_col: Optional[str] = None,
    row_num_col: Optional[str] = None,
    no_backup: bool = False
) -> Dict[str, any]:
    """
    Executes a generic column and status transfer from source_path to target_path.
    Supports both ultra-fast streaming XML patch mode (2-3s for 250MB+ workbooks)
    and deep DOM structural fallback.
    """
    if not os.path.exists(source_path):
        raise FileNotFoundError(f"Source file not found: {source_path}")
    if not os.path.exists(target_path):
        raise FileNotFoundError(f"Target file not found: {target_path}")

    if output_path is None:
        output_path = target_path

    if not no_backup and output_path == target_path:
        backup = create_backup(target_path)
        if backup:
            print(f"[备份] 已创建目标文件备份: {backup}")

    tgt_flavor = detect_excel_flavor(target_path)
    print(f"[探测] 目标文件格式方言: {tgt_flavor}")

    # Check if source is a single-sheet inspection/summary report with a branch column
    xl_src_file = pd.ExcelFile(source_path)
    is_branch_dispatch = False
    first_sheet_df = pd.read_excel(source_path, sheet_name=0, nrows=5)
    
    # Auto-detect branch column (e.g. '所属分表', '分表', '工作表', 'sheet')
    detected_branch_col = branch_col
    if not detected_branch_col:
        for cand in ['所属分表', '分表', '工作表', 'sheet', '所属工作表']:
            if cand in first_sheet_df.columns:
                detected_branch_col = cand
                break
                
    if detected_branch_col and len(xl_src_file.sheet_names) == 1:
        is_branch_dispatch = True
        print(f"[模式] 检测到巡检汇总表分表派发模式 (分表字段: '{detected_branch_col}')")

    # Ultra-Fast Stream Channel for Branch Dispatch Mode (100% Image Safe)
    if is_branch_dispatch:
        print(f"[极速通道] 启用 XmlStreamPatcher 零损流式引擎，极速批量同步 (100% 物理级保护所有嵌入图片与公式)")
        df_src_all = pd.read_excel(source_path, sheet_name=0)
        
        detected_row_col = row_num_col
        if not detected_row_col:
            for cand in ['原表行号', '行号', 'row_num', 'excel_row']:
                if cand in df_src_all.columns:
                    detected_row_col = cand
                    break

        branches = df_src_all[detected_branch_col].dropna().unique()
        src_cols_available = [c for c in df_src_all.columns if c not in [detected_branch_col, detected_row_col]]
        
        actual_transfer_cols = []
        if cols and cols != ['auto']:
            actual_transfer_cols = [c for c in cols if c in src_cols_available]
        else:
            for cand in ['是否已下架（1下架）', '是否已下架', '是否下架', '检测备注/异常状态', '检测时间']:
                if cand in src_cols_available and cand not in actual_transfer_cols:
                    actual_transfer_cols.append(cand)
            if not actual_transfer_cols:
                actual_transfer_cols = src_cols_available

        print(f"[同步字段] 待同步列: {actual_transfer_cols}")
        columns_by_sheet = defaultdict(list)
        
        for b_name in branches:
            df_b = df_src_all[df_src_all[detected_branch_col] == b_name]
            if len(df_b) == 0:
                continue
                
            for col_name in actual_transfer_cols:
                row_map = {}
                for _, r in df_b.iterrows():
                    val = r.get(col_name)
                    if pd.notna(val) and detected_row_col and pd.notna(r.get(detected_row_col)):
                        row_map[int(r[detected_row_col])] = val
                
                if row_map:
                    columns_by_sheet[b_name].append({
                        "header": col_name,
                        "data": row_map
                    })

        res = XmlStreamPatcher.append_or_patch_columns(
            target_path=target_path,
            columns_by_sheet=columns_by_sheet,
            output_path=output_path
        )
        print(f"[完成] 流式极速回写成功！总更新单元格数: {res.get('total_cells_updated', 0)}，目标文件: {res.get('output')}")
        return res

    # Standard Fallback via DOM for Complex Relocations
    wb_src = openpyxl.load_workbook(source_path, data_only=False)
    wb_tgt = openpyxl.load_workbook(target_path, data_only=False)

    total_cells_updated = 0
    sheet_reports = {}

    target_sheet_names = wb_tgt.sheetnames if sheets == ['auto'] else [s for s in sheets if s in wb_tgt.sheetnames]

    for s_name in target_sheet_names:
        if s_name not in wb_src.sheetnames:
            continue

        ws_src = wb_src[s_name]
        ws_tgt = wb_tgt[s_name]

        df_src_sheet = pd.read_excel(source_path, sheet_name=s_name)
        df_tgt_sheet = pd.read_excel(target_path, sheet_name=s_name)

        src_headers = [str(ws_src.cell(1, col).value or '').strip() for col in range(1, ws_src.max_column + 1)]
        tgt_headers = [str(ws_tgt.cell(1, col).value or '').strip() for col in range(1, ws_tgt.max_column + 1)]

        transfer_cols = []
        if cols == ['auto']:
            transfer_cols = [c for c in src_headers if c and c not in tgt_headers]
        else:
            transfer_cols = [c for c in cols if c in src_headers]

        if not transfer_cols:
            continue

        key_cols = keys
        if not key_cols or key_cols == ['auto']:
            key_cols = auto_detect_key_columns(df_src_sheet, df_tgt_sheet)

        matcher = CascadingMatcher(df_src_sheet, df_tgt_sheet, key_cols)
        matched_pairs = matcher.match()

        for col_name in transfer_cols:
            src_col_idx = src_headers.index(col_name) + 1
            if col_name in tgt_headers:
                tgt_col_idx = tgt_headers.index(col_name) + 1
            else:
                tgt_col_idx = ws_tgt.max_column + 1
                ws_tgt.cell(1, tgt_col_idx, col_name)

            col_updates = 0
            for src_row_idx, tgt_row_idx in matched_pairs.items():
                val = ws_src.cell(src_row_idx, src_col_idx).value
                if val is not None:
                    ws_tgt.cell(tgt_row_idx, tgt_col_idx, val)
                    col_updates += 1

            total_cells_updated += col_updates

    # Save target workbook
    temp_dir = os.path.dirname(os.path.abspath(output_path))
    temp_out = os.path.join(temp_dir, f"_openpyxl_tmp_{os.path.basename(output_path)}")
    wb_tgt.save(temp_out)
    final_output = atomic_replace_file(temp_out, output_path)

    wb_src.close()
    wb_tgt.close()

    return {
        'status': 'success',
        'output': final_output,
        'total_cells_updated': total_cells_updated
    }
