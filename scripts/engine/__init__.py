from .sniffer import detect_excel_flavor
from .formula_parser import shift_formula_references, col_to_num, num_to_col
from .matcher import CascadingMatcher, auto_detect_key_columns
from .safe_io import is_file_locked, create_backup, atomic_replace_file
from .wps_engine import WpsImageEngine
from .msoffice_engine import MsOfficeImageEngine
from .xml_patcher import XmlStreamPatcher
from .summary_detector import is_summary_text, is_summary_row, is_footnote_text
from .layout_utils import get_string_display_width, calculate_auto_column_widths
from .security import is_dangerous_formula, sanitize_plain_text, detect_prompt_injection
