import os
import re
import html
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Any, Union, Tuple
from engine import (
    detect_excel_flavor,
    create_backup,
    atomic_replace_file,
    is_file_locked,
    XmlStreamPatcher,
    col_to_num,
    num_to_col,
    is_summary_row
)


def extract_row_text_sample(row_xml: str, shared_strings: List[str], max_cells: int = 5) -> List[str]:
    """提取行内前若干单元格的文本值，用于智能识别汇总/小计或说明行。"""
    texts: List[str] = []
    for c_match in re.finditer(r'<c\b([^/>]*)(?:/>|>(.*?)</c>)', row_xml, re.DOTALL):
        c_attrs = c_match.group(1)
        c_inner = c_match.group(2) or ""
        v_match = re.search(r'<v>(.*?)</v>', c_inner)
        t_match = re.search(r'<t[^>]*>(.*?)</t>', c_inner)
        
        cell_val = ""
        if 't="s"' in c_attrs:
            if v_match and v_match.group(1).isdigit():
                s_idx = int(v_match.group(1))
                if 0 <= s_idx < len(shared_strings):
                    cell_val = shared_strings[s_idx]
        elif t_match:
            cell_val = t_match.group(1)
        elif v_match:
            cell_val = v_match.group(1)
        
        if cell_val.strip():
            texts.append(html.unescape(cell_val).strip())
            if len(texts) >= max_cells:
                break
    return texts


def parse_shared_strings(zip_ref: zipfile.ZipFile) -> List[str]:
    """Extracts all strings from xl/sharedStrings.xml if present."""
    if 'xl/sharedStrings.xml' not in zip_ref.namelist():
        return []
    
    strings = []
    try:
        content = zip_ref.read('xl/sharedStrings.xml')
        root = ET.fromstring(content)
        ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        for si in root.findall('s:si', ns):
            t_elems = si.findall('.//s:t', ns)
            strings.append(''.join(t.text or '' for t in t_elems))
    except Exception as e:
        print(f"[警告] 解析 sharedStrings.xml 失败: {e}")
    return strings


def parse_filter_expr(expr_str: str) -> Tuple[str, str, str]:
    """
    Parses expressions like:
      - '毛利率 >= 15%'
      - '京东价格 < 50'
      - '品牌 == 半亩花田'
      - '一级分类 contains 家居'
    Returns: (col_name, op, val_str)
    """
    pattern = re.compile(r'^\s*([^\s>=<!]+(?:\s+[^\s>=<!]+)*)\s*(>=|<=|>|<|==|!=|contains|not_contains)\s*(.*?)\s*$')
    m = pattern.match(expr_str.strip())
    if not m:
        raise ValueError(f"无法解析过滤表达式: '{expr_str}'。格式示例: '毛利率 >= 15%' 或 '品牌 == 半亩花田'")
    col_name, op, val_str = m.groups()
    return col_name.strip(), op.strip(), val_str.strip()


def run_filter(
    file_path: str,
    col: Optional[str] = None,
    op: str = ">=",
    val: Optional[Union[str, float, int]] = None,
    expr: Optional[str] = None,
    keep_empty: bool = False,
    preserve_summary: bool = True,
    sheet: Optional[str] = None,
    output_path: Optional[str] = None,
    no_backup: bool = False
) -> Dict[str, Any]:
    """
    High-speed, 100% lossless row filtering engine for large Excel workbooks (.xlsx).
    Preserves all Office 365 RichData images, WPS DISPIMG cell images, hyperlinks,
    column widths, formatting, and XML structures.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    # Parse expression if provided
    if expr:
        parsed_col, parsed_op, parsed_val = parse_filter_expr(expr)
        col = parsed_col
        op = parsed_op
        val = parsed_val

    if not col:
        raise ValueError("必须指定待过滤的列名 (--col) 或过滤表达式 (--expr)")

    if val is None:
        raise ValueError("必须指定过滤阈值/目标值 (--val) 或过滤表达式 (--expr)")

    # Normalize val
    val_str = str(val).strip()
    is_percent_val = val_str.endswith('%')
    threshold_numeric: Optional[float] = None
    try:
        clean_num_str = val_str.replace('%', '')
        v_float = float(clean_num_str)
        threshold_numeric = v_float / 100.0 if is_percent_val else v_float
    except ValueError:
        threshold_numeric = None

    if output_path is None:
        output_path = file_path

    if not no_backup and output_path == file_path:
        backup = create_backup(file_path)
        if backup:
            print(f"[备份] 已创建原文件备份: {backup}")

    flavor = detect_excel_flavor(file_path)
    print(f"[探测] 文件格式方言: {flavor}")

    with zipfile.ZipFile(file_path, 'r') as z_in:
        namelist = z_in.namelist()
        shared_strings = parse_shared_strings(z_in)
        sheet_xml_map = XmlStreamPatcher.get_sheet_xml_map(file_path)
        
        target_xml_path = None
        if sheet:
            target_xml_path = sheet_xml_map.get(sheet)
            if not target_xml_path:
                for k, v in sheet_xml_map.items():
                    if sheet.lower() in k.lower():
                        target_xml_path = v
                        break
        if not target_xml_path:
            target_xml_path = list(sheet_xml_map.values())[0] if sheet_xml_map else 'xl/worksheets/sheet1.xml'

        if target_xml_path not in namelist:
            raise FileNotFoundError(f"未在工作簿中找到工作表文件: {target_xml_path}")

        print(f"[读取] 目标工作表: {target_xml_path}")
        sheet_xml = z_in.read(target_xml_path).decode('utf-8')
        all_files = {name: z_in.read(name) for name in namelist}

    # Step 1: Parse row 1 (headers) to find target column letter
    row_pattern = re.compile(r'<row r="(\d+)"[^>]*>(.*?)</row>', re.DOTALL)
    header_match = row_pattern.search(sheet_xml)
    if not header_match or header_match.group(1) != '1':
        raise ValueError("工作表中未能定位第 1 行表头")

    header_row_xml = header_match.group(0)
    # Parse header cells strictly inspecting attributes and inner content
    header_cells = re.findall(
        r'<c\b([^/>]*\br="([A-Z]+)1"[^/>]*)(?:/>|>(.*?)</c>)',
        header_row_xml,
        re.DOTALL
    )

    col_letter_map: Dict[str, str] = {} # clean_header_name -> col_letter
    letter_to_header: Dict[str, str] = {}

    for c_attrs, c_letter, c_body in header_cells:
        h_text = ""
        if c_body:
            v_match = re.search(r'<v>(.*?)</v>', c_body)
            t_match = re.search(r'<t[^>]*>(.*?)</t>', c_body)
            if 't="s"' in c_attrs:
                if v_match and v_match.group(1).isdigit():
                    s_idx = int(v_match.group(1))
                    if 0 <= s_idx < len(shared_strings):
                        h_text = shared_strings[s_idx]
            elif t_match:
                h_text = t_match.group(1)
            elif v_match:
                h_text = v_match.group(1)
        
        clean_name = html.unescape(h_text).replace('\n', '').strip()
        col_letter_map[clean_name] = c_letter
        letter_to_header[c_letter] = clean_name

    # Locate target col letter
    target_col_letter: Optional[str] = None
    clean_target_col = col.replace('\n', '').strip()

    if clean_target_col.upper() in letter_to_header:
        target_col_letter = clean_target_col.upper()
    elif clean_target_col in col_letter_map:
        target_col_letter = col_letter_map[clean_target_col]
    else:
        for h_name, ltr in col_letter_map.items():
            if clean_target_col in h_name or h_name in clean_target_col:
                target_col_letter = ltr
                break

    if not target_col_letter:
        raise ValueError(f"未在表头中找到匹配的列 '{col}'。当前可用表头列: {list(col_letter_map.keys())}")

    target_header_display = letter_to_header.get(target_col_letter, target_col_letter)
    print(f"[定位] 目标字段: '{target_header_display}' (列号: {target_col_letter})")
    print(f"[条件] 过滤规则: {target_header_display} {op} {val_str} (保留空值: {keep_empty})")

    # Step 2: Stream scan all rows and evaluate condition
    kept_rows: List[Tuple[int, str]] = []
    summary_rows: List[Tuple[int, str]] = []
    dropped_count = 0
    total_data_rows = 0

    cell_target_pattern = re.compile(
        r'<c\b([^/>]*\br="' + target_col_letter + r'(\d+)"[^/>]*)(?:/>|>(.*?)</c>)',
        re.DOTALL
    )

    for match in row_pattern.finditer(sheet_xml):
        old_r = int(match.group(1))
        if old_r == 1:
            continue
        total_data_rows += 1
        row_xml = match.group(0)

        # 智能识别并保护汇总/小计/说明行（避免汇总行因不满足明细条件被误杀或导致重复计算）
        if preserve_summary:
            row_texts = extract_row_text_sample(row_xml, shared_strings)
            if is_summary_row(row_texts):
                desc = ' '.join(row_texts[:2])
                print(f"[保护] 识别到第 {old_r} 行为汇总/小计行 ('{desc}')，自动保留置于表尾")
                summary_rows.append((old_r, row_xml))
                continue

        # Extract cell value
        cell_match = cell_target_pattern.search(row_xml)
        cell_val: Optional[Any] = None

        if cell_match and cell_match.group(3):
            c_attrs = cell_match.group(1)
            c_inner = cell_match.group(3)
            
            v_match = re.search(r'<v>(.*?)</v>', c_inner)
            t_match = re.search(r'<t[^>]*>(.*?)</t>', c_inner)

            if 't="s"' in c_attrs:
                if v_match and v_match.group(1).isdigit():
                    s_idx = int(v_match.group(1))
                    if 0 <= s_idx < len(shared_strings):
                        cell_val = shared_strings[s_idx]
            elif t_match:
                cell_val = t_match.group(1)
            elif v_match:
                raw_v = v_match.group(1)
                try:
                    cell_val = float(raw_v)
                except ValueError:
                    cell_val = raw_v

            if isinstance(cell_val, str):
                cell_val = html.unescape(cell_val)

        # Evaluate condition
        matched = False
        if cell_val is None or str(cell_val).strip() == '':
            matched = keep_empty
        else:
            # Try numeric comparison
            num_cell: Optional[float] = None
            if isinstance(cell_val, (int, float)):
                num_cell = float(cell_val)
            elif isinstance(cell_val, str):
                s = cell_val.strip()
                if s.endswith('%'):
                    try:
                        num_cell = float(s[:-1]) / 100.0
                    except ValueError:
                        pass
                else:
                    try:
                        num_cell = float(s)
                    except ValueError:
                        pass

            if threshold_numeric is not None and num_cell is not None:
                if op == '>=':
                    matched = (num_cell >= threshold_numeric)
                elif op == '<=':
                    matched = (num_cell <= threshold_numeric)
                elif op == '>':
                    matched = (num_cell > threshold_numeric)
                elif op == '<':
                    matched = (num_cell < threshold_numeric)
                elif op in ('==', '='):
                    matched = (abs(num_cell - threshold_numeric) < 1e-9)
                elif op == '!=':
                    matched = (abs(num_cell - threshold_numeric) >= 1e-9)
            else:
                # String comparison
                s_cell = str(cell_val).strip()
                s_target = val_str
                if op == 'contains':
                    matched = (s_target in s_cell)
                elif op == 'not_contains':
                    matched = (s_target not in s_cell)
                elif op in ('==', '='):
                    matched = (s_cell == s_target)
                elif op == '!=':
                    matched = (s_cell != s_target)

        if matched:
            kept_rows.append((old_r, row_xml))
        else:
            dropped_count += 1

    all_final_rows = kept_rows + summary_rows
    print(f"[统计] 总扫描数据行: {total_data_rows} 行 | 满足条件保留: {len(kept_rows)} 行 | 保护汇总行: {len(summary_rows)} 行 | 过滤剔除: {dropped_count} 行")

    # Step 3: Renumber rows and cell coordinates
    old_to_new: Dict[int, int] = {}
    new_rows_xml_list = [header_row_xml]
    cell_r_pattern = re.compile(r'(\br=")([A-Z]+)(\d+)(")')

    for idx, (old_r, row_xml) in enumerate(all_final_rows, start=2):
        new_r = idx
        old_to_new[old_r] = new_r
        
        updated_row = re.sub(r'^<row r="' + str(old_r) + r'"', f'<row r="{new_r}"', row_xml)
        
        def replace_coord(m):
            prefix, col_letters, row_num, suffix = m.groups()
            if int(row_num) == old_r:
                return f"{prefix}{col_letters}{new_r}{suffix}"
            return m.group(0)

        updated_row = cell_r_pattern.sub(replace_coord, updated_row)
        new_rows_xml_list.append(updated_row)

    max_new_row = len(new_rows_xml_list)

    # Step 4: Reassemble sheetData, dimension, autoFilter, and hyperlinks
    new_sheet_data = "<sheetData>" + "".join(new_rows_xml_list) + "</sheetData>"
    sheet_new_xml = re.sub(r'<sheetData>.*?</sheetData>', lambda m: new_sheet_data, sheet_xml, flags=re.DOTALL)

    # Update dimension
    sheet_new_xml = re.sub(r'<dimension ref="([A-Z0-9]+:)([A-Z]+)\d+"', rf'<dimension ref="\g<1>\g<2>{max_new_row}"', sheet_new_xml)

    # Update autoFilter
    sheet_new_xml = re.sub(r'<autoFilter ref="([A-Z0-9]+:)([A-Z]+)\d+"', rf'<autoFilter ref="\g<1>\g<2>{max_new_row}"', sheet_new_xml)

    # Update hyperlinks
    hyperlinks_match = re.search(r'<hyperlinks>(.*?)</hyperlinks>', sheet_new_xml, re.DOTALL)
    if hyperlinks_match:
        hyperlinks_body = hyperlinks_match.group(1)
        hl_pattern = re.compile(r'<hyperlink ref="([A-Z]+)(\d+)"([^>]*)/>')
        new_hl_list = []

        for hl in hl_pattern.finditer(hyperlinks_body):
            col_letters, row_str, rest = hl.groups()
            row_int = int(row_str)
            if row_int in old_to_new:
                new_r = old_to_new[row_int]
                new_hl_list.append(f'<hyperlink ref="{col_letters}{new_r}"{rest}/>')

        new_hl_tag = f"<hyperlinks>{''.join(new_hl_list)}</hyperlinks>"
        sheet_new_xml = sheet_new_xml[:hyperlinks_match.start()] + new_hl_tag + sheet_new_xml[hyperlinks_match.end():]

    all_files[target_xml_path] = sheet_new_xml.encode('utf-8')

    # Step 5: Safe atomic save
    temp_dir = os.path.dirname(os.path.abspath(output_path))
    temp_out = os.path.join(temp_dir, f"_filter_tmp_{os.path.basename(output_path)}")

    with zipfile.ZipFile(temp_out, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
        for name, content in all_files.items():
            z_out.writestr(name, content)

    final_output = atomic_replace_file(temp_out, output_path, fallback_suffix="_已筛选")
    print(f"[完成] 物理级零损过滤完成！输出文件: {final_output}")

    return {
        'status': 'success',
        'output': final_output,
        'total_rows': total_data_rows,
        'kept_rows': len(kept_rows),
        'dropped_rows': dropped_count
    }
