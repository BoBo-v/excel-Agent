import os
import shutil
import zipfile
import re
from typing import Dict, List, Set, Optional
import pandas as pd
from .matcher import CompositeSpecMatcher
from .safe_io import atomic_replace_file
from .formula_parser import num_to_col

class MsOfficeImageEngine:
    """
    Handles Microsoft Excel / Tencent Docs RichData (XLRICHVALUE) cell-embedded images
    and DrawingML parts.
    """

    @staticmethod
    def preserve_msoffice_parts(orig_xlsx: str, temp_xlsx: str, output_path: str):
        """Preserves any Microsoft Office richData or drawing parts during save."""
        with zipfile.ZipFile(orig_xlsx, 'r') as z_orig:
            orig_names = z_orig.namelist()
            rich_parts = {n: z_orig.read(n) for n in orig_names if n.startswith('xl/richData/')}
            drawing_parts = {n: z_orig.read(n) for n in orig_names if n.startswith('xl/drawings/')}
            media_parts = {n: z_orig.read(n) for n in orig_names if n.startswith('xl/media/')}

        with zipfile.ZipFile(temp_xlsx, 'r') as z_temp:
            temp_files = {n: z_temp.read(n) for n in z_temp.namelist()}

        # Merge in any richData and media parts that might have been skipped
        for k, v in {**rich_parts, **drawing_parts, **media_parts}.items():
            if k not in temp_files:
                temp_files[k] = v

        temp_out = f"{output_path}.atomic_ms"
        with zipfile.ZipFile(temp_out, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
            for fname, fbytes in temp_files.items():
                z_out.writestr(fname, fbytes)

        atomic_replace_file(temp_out, output_path)

    @staticmethod
    def supplement_rich_images(
        target_path: str,
        source_paths: List[str],
        output_path: Optional[str] = None,
        image_col: Optional[str] = None,
        keys: Optional[List[str]] = None
    ) -> int:
        """
        Supplements missing RichData (XLRICHVALUE) embedded images from source_paths into target_path.
        Supports multi-spec cascading matching. Returns number of enriched cells.
        """
        if output_path is None:
            output_path = target_path

        df_tgt = pd.read_excel(target_path)

        # Detect image column in target
        target_img_col_name = image_col
        if not target_img_col_name:
            for cand in ['商品图片', '图片', '主图', '产品图片', 'image', 'picture']:
                for c in df_tgt.columns:
                    if cand.lower() in str(c).strip().lower():
                        target_img_col_name = c
                        break
                if target_img_col_name:
                    break

        if not target_img_col_name:
            # Fallback to column index 7 (col G, common in catalog tables) if available
            if len(df_tgt.columns) >= 7:
                target_img_col_name = df_tgt.columns[6]
            else:
                raise ValueError("Could not automatically detect image column in target table.")

        target_img_col_idx = list(df_tgt.columns).index(target_img_col_name) + 1
        col_letter = num_to_col(target_img_col_idx)
        print(f"[探测] 目标表图片列: '{target_img_col_name}' (列号: {col_letter})")

        with zipfile.ZipFile(target_path, 'r') as z_tgt:
            tgt_files = {name: z_tgt.read(name) for name in z_tgt.namelist()}

        total_enriched = 0

        for src_path in source_paths:
            if not os.path.exists(src_path):
                print(f"[警告] 图源文件不存在: {src_path}")
                continue

            with zipfile.ZipFile(src_path, 'r') as z_src:
                src_namelist = z_src.namelist()
                has_rich = any(n.startswith('xl/richData/') for n in src_namelist)
                if not has_rich:
                    continue

                src_files = {name: z_src.read(name) for name in src_namelist}

            df_src = pd.read_excel(src_path)

            # Detect image column in source
            src_img_col_name = target_img_col_name if target_img_col_name in df_src.columns else None
            if not src_img_col_name:
                for cand in ['商品图片', '图片', '主图', '产品图片', 'image', 'picture']:
                    for c in df_src.columns:
                        if cand.lower() in str(c).strip().lower():
                            src_img_col_name = c
                            break
                    if src_img_col_name:
                        break
            if not src_img_col_name:
                src_img_col_name = df_src.columns[6] if len(df_src.columns) >= 7 else df_src.columns[0]

            src_col_idx = list(df_src.columns).index(src_img_col_name) + 1
            src_col_letter = num_to_col(src_col_idx)

            # Extract row -> vm mapping from source sheet1.xml
            s_sheet_name = 'xl/worksheets/sheet1.xml'
            if s_sheet_name not in src_files:
                sheet_cands = [n for n in src_files.keys() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')]
                s_sheet_name = sheet_cands[0] if sheet_cands else None

            if not s_sheet_name:
                continue

            s_sheet_xml = src_files[s_sheet_name].decode('utf-8')
            pattern = rf'<c r="{src_col_letter}(\d+)"[^>]*?vm="(\d+)"'
            g_cells = re.findall(pattern, s_sheet_xml)
            src_row_to_vm = {int(r): int(vm) for r, vm in g_cells}

            # Map DataFrame index (0-based) -> Excel row (index + 2) -> vm
            df_src_vm = {}
            for idx in df_src.index:
                excel_r = idx + 2
                if excel_r in src_row_to_vm:
                    df_src_vm[idx] = src_row_to_vm[excel_r]

            # Match target rows using CompositeSpecMatcher
            matcher = CompositeSpecMatcher(df_src, df_tgt)
            target_row_to_vm = {}

            for tgt_idx, tgt_row in df_tgt.iterrows():
                tgt_excel_r = tgt_idx + 2
                matched_src_idx = matcher.match_row(tgt_row)
                if matched_src_idx is not None and matched_src_idx in df_src_vm:
                    target_row_to_vm[tgt_excel_r] = df_src_vm[matched_src_idx]

            print(f"[匹配] 从 {os.path.basename(src_path)} 成功匹配图源行: {len(target_row_to_vm)} / {len(df_tgt)}")

            if not target_row_to_vm:
                continue

            # Merge media and richData parts from source
            for name, content in src_files.items():
                if name.startswith('xl/media/') or name.startswith('xl/richData/') or name in ('xl/metadata.xml', 'xl/cellimages.xml'):
                    tgt_files[name] = content

            if 'xl/_rels/workbook.xml.rels' in src_files:
                tgt_files['xl/_rels/workbook.xml.rels'] = src_files['xl/_rels/workbook.xml.rels']
            if '[Content_Types].xml' in src_files:
                tgt_files['[Content_Types].xml'] = src_files['[Content_Types].xml']

            # Update target sheet XML
            t_sheet_name = 'xl/worksheets/sheet1.xml'
            if t_sheet_name in tgt_files:
                t_sheet_text = tgt_files[t_sheet_name].decode('utf-8')

                def replace_cell_vm(match):
                    full_tag = match.group(0)
                    row_num = int(match.group(1))
                    if row_num in target_row_to_vm:
                        vm = target_row_to_vm[row_num]
                        tag = full_tag
                        if 'vm="' in tag:
                            tag = re.sub(r'vm="\d+"', f'vm="{vm}"', tag)
                        else:
                            tag = re.sub(rf'(r="{col_letter}\d+")', rf'\1 vm="{vm}"', tag)
                        if 't="inlineStr"' in tag and '#VALUE!' in tag:
                            tag = re.sub(r't="inlineStr"', 't="e"', tag)
                            tag = re.sub(r'<is><t>#VALUE!</t></is>', '<v>#VALUE!</v>', tag)
                        return tag
                    return full_tag

                cell_regex = rf'<c r="{col_letter}(\d+)"[^>]*?(?:/>|>.*?</c>)'
                updated_sheet_text = re.sub(cell_regex, replace_cell_vm, t_sheet_text)
                tgt_files[t_sheet_name] = updated_sheet_text.encode('utf-8')

            total_enriched += len(target_row_to_vm)

        # Write to temporary output and atomically replace
        temp_out = f"{output_path}.atomic_rich"
        with zipfile.ZipFile(temp_out, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
            for fname, fbytes in tgt_files.items():
                z_out.writestr(fname, fbytes)

        atomic_replace_file(temp_out, output_path)
        return total_enriched
