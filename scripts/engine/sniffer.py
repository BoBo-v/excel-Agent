import zipfile
import os

def detect_excel_flavor(xlsx_path: str) -> str:
    """
    Detects whether an .xlsx workbook uses WPS cellimages, MS Office RichData/Drawings, or standard.
    Returns:
        'wps': Contains WPS cellimages.xml or DISPIMG formulas
        'msoffice_rich': Contains Office 365 / Tencent Docs richData cell images
        'msoffice_drawing': Contains standard Office drawings/shapes
        'standard': Plain tabular data without embedded images
    """
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"File not found: {xlsx_path}")

    try:
        with zipfile.ZipFile(xlsx_path, 'r') as z:
            names = z.namelist()

            has_rich = any(n.startswith('xl/richData/') for n in names)
            has_cellimages = 'xl/cellimages.xml' in names or any('cellimage' in n.lower() for n in names)

            # Check if cellimages has actual images or is just an empty placeholder
            if has_cellimages:
                try:
                    ci_sample = z.read('xl/cellimages.xml').decode('utf-8', errors='ignore')
                    # If it has <etc:cellImage or <cellImage, it's real WPS cellimages
                    if '<etc:cellImage ' in ci_sample or '<cellImage ' in ci_sample or 'etc:cellImage>' in ci_sample:
                        return 'wps'
                except Exception:
                    pass

            if has_rich:
                return 'msoffice_rich'

            if has_cellimages:
                return 'wps'

            # Check if any worksheet XML contains DISPIMG
            for n in names:
                if n.startswith('xl/worksheets/sheet') and n.endswith('.xml'):
                    try:
                        content_sample = z.read(n)[:2000].decode('utf-8', errors='ignore')
                        if 'DISPIMG' in content_sample:
                            return 'wps'
                    except Exception:
                        pass

            if any(n.startswith('xl/drawings/') for n in names):
                return 'msoffice_drawing'

            return 'standard'
    except Exception:
        return 'standard'
