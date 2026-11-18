import os
from typing import List, Optional, Dict
from engine import detect_excel_flavor, WpsImageEngine, MsOfficeImageEngine, create_backup

def run_fill_images(
    target_path: str,
    source_paths: List[str],
    output_path: Optional[str] = None,
    image_col: Optional[str] = None,
    keys: Optional[List[str]] = None,
    no_backup: bool = False
) -> Dict[str, any]:
    """
    Supplements missing product images in target_path from multiple source_paths.
    Dual-engine support: WPS DISPIMG embedded images + Microsoft Office / Tencent Docs RichData.
    """
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

    src_flavors = {s: detect_excel_flavor(s) for s in source_paths if os.path.exists(s)}
    print(f"[探测] 图源文件格式方言: {src_flavors}")

    total_added = 0

    # 1. Check for Microsoft Office / Tencent Docs RichData images in sources
    has_rich_source = any(f == 'msoffice_rich' for f in src_flavors.values())
    if has_rich_source or tgt_flavor == 'msoffice_rich':
        print(f"[处理] 调用 Office/腾讯文档 RichData (XLRICHVALUE) 跨表补图引擎...")
        rich_sources = [s for s, f in src_flavors.items() if f == 'msoffice_rich']
        if not rich_sources:
            rich_sources = source_paths
        enriched = MsOfficeImageEngine.supplement_rich_images(
            target_path=target_path,
            source_paths=rich_sources,
            output_path=output_path,
            image_col=image_col,
            keys=keys
        )
        total_added += enriched
        print(f"[完成] RichData 图片补全成功: {enriched} 单元格")

    # 2. Check for WPS DISPIMG cellimages in sources
    has_wps_source = any(f == 'wps' for f in src_flavors.values())
    if has_wps_source or tgt_flavor == 'wps':
        print(f"[处理] 调用 WPS DISPIMG 嵌入图片合并引擎...")
        wps_sources = [s for s, f in src_flavors.items() if f == 'wps']
        if not wps_sources:
            wps_sources = source_paths
        wps_added = WpsImageEngine.supplement_missing_images(
            target_path=output_path if total_added > 0 else target_path,
            source_paths=wps_sources,
            output_path=output_path
        )
        total_added += wps_added
        print(f"[完成] WPS DISPIMG 图片合并成功: {wps_added} 张")

    if total_added == 0 and not has_rich_source and not has_wps_source:
        print("[提示] 来源文件中未检测到嵌入图片结构 (WPS DISPIMG 或 RichData)。")

    return {
        'status': 'success' if total_added > 0 else 'skipped',
        'added_count': total_added,
        'output': output_path
    }
