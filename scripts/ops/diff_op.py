import os
import pandas as pd
from typing import List, Optional, Dict
from engine import CascadingMatcher, auto_detect_key_columns

def run_diff(
    base_path: str,
    compare_path: str,
    keys: Optional[List[str]] = None,
    watch_cols: Optional[List[str]] = None,
    output_report: Optional[str] = None
) -> Dict[str, any]:
    """
    Compares two table versions and generates change summary / report.
    """
    if not os.path.exists(base_path):
        raise FileNotFoundError(f"Base file not found: {base_path}")
    if not os.path.exists(compare_path):
        raise FileNotFoundError(f"Compare file not found: {compare_path}")

    xl_b = pd.ExcelFile(base_path)
    xl_c = pd.ExcelFile(compare_path)
    common_sheets = [s for s in xl_b.sheet_names if s in xl_c.sheet_names]

    all_diffs = []
    print(f"[比对] 共有 {len(common_sheets)} 个同名工作表需要核对")

    for s in common_sheets:
        df_b = pd.read_excel(base_path, sheet_name=s)
        df_c = pd.read_excel(compare_path, sheet_name=s)

        actual_keys = keys
        if not actual_keys or actual_keys == ['auto']:
            actual_keys = auto_detect_key_columns(df_b, df_c)

        actual_watch = watch_cols
        if not actual_watch or actual_watch == ['auto']:
            actual_watch = [c for c in df_b.columns if c in df_c.columns and c not in actual_keys]

        matcher = CascadingMatcher(df_c, actual_keys)

        for idx, row in df_b.iterrows():
            matched_indices = matcher.match_row(row)
            if not matched_indices:
                all_diffs.append({
                    'sheet': s,
                    'type': '已删除/未匹配',
                    'identifier': str(row.get(actual_keys[0]) if actual_keys else idx),
                    'col': '全部',
                    'old_val': '存在',
                    'new_val': '缺失'
                })
            else:
                c_row = df_c.iloc[matched_indices[0]]
                for col in actual_watch:
                    v_old = row.get(col)
                    v_new = c_row.get(col)
                    if str(v_old).strip() != str(v_new).strip() and (pd.notna(v_old) or pd.notna(v_new)):
                        all_diffs.append({
                            'sheet': s,
                            'type': '字段值变动',
                            'identifier': str(row.get(actual_keys[0]) if actual_keys else idx),
                            'col': col,
                            'old_val': str(v_old),
                            'new_val': str(v_new)
                        })

    print(f"[结果] 比对完成，共发现 {len(all_diffs)} 处差异/变动")
    df_result = pd.DataFrame(all_diffs)

    if output_report and len(df_result) > 0:
        df_result.to_excel(output_report, index=False)
        print(f"[报表] 差异报表已生成: {output_report}")

    return {
        'status': 'success',
        'diff_count': len(all_diffs),
        'report': output_report
    }
