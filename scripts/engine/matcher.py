from typing import List, Dict, Any, Optional, Tuple
import pandas as pd

COMMON_KEY_CANDIDATES = [
    '京东链接\n（必填）', '京东链接', '商品链接', '链接', 'url',
    '产品69码（必填）', '产品69码', '69码', '商品条码', '条形码', '条码', 'barcode',
    '供应商产品\n唯一识别码', '供应商产品唯一识别码', '商家编码', '商品编码', '编码', 'sku', 'code',
    '产品名称', '商品名称', '标题', '名称', 'name', 'title',
    '产品规格', '规格', '规格名称', 'spec'
]

def auto_detect_key_columns(df1: pd.DataFrame, df2: pd.DataFrame) -> List[str]:
    """Finds matching key columns between df1 and df2 ranked by uniqueness and reliability."""
    common_cols = [c for c in df1.columns if c in df2.columns]
    matched_keys = []

    # Priority: known candidates first
    for cand in COMMON_KEY_CANDIDATES:
        for c in common_cols:
            if cand.lower() == str(c).strip().lower() or cand in str(c):
                if c not in matched_keys:
                    matched_keys.append(c)

    # If none found, fallback to any common columns with high uniqueness
    if not matched_keys:
        for c in common_cols:
            if df1[c].nunique() > len(df1) * 0.5:
                matched_keys.append(c)

    return matched_keys


class CompositeSpecMatcher:
    """
    Advanced matcher handling multi-specification e-commerce tables.
    Matches rows across tables by cascading composite keys (e.g. Link + Spec, Code + Spec).
    """

    def __init__(self, src_df: pd.DataFrame, tgt_df: pd.DataFrame):
        self.src_df = src_df
        self.tgt_df = tgt_df

        # Detect column roles in src and tgt
        self.src_link = self._find_col(src_df, ['京东链接\n（必填）', '京东链接', '商品链接', '链接', 'url'])
        self.tgt_link = self._find_col(tgt_df, ['京东链接\n（必填）', '京东链接', '商品链接', '链接', 'url'])

        self.src_spec = self._find_col(src_df, ['产品规格', '规格名称', '规格', 'spec'])
        self.tgt_spec = self._find_col(tgt_df, ['产品规格', '规格名称', '规格', 'spec'])

        self.src_name = self._find_col(src_df, ['产品名称', '商品名称', '标题', '名称', 'name', 'title'])
        self.tgt_name = self._find_col(tgt_df, ['产品名称', '商品名称', '标题', '名称', 'name', 'title'])

        self.src_69 = self._find_col(src_df, ['产品69码（必填）', '产品69码', '69码', '条码', 'barcode'])
        self.tgt_69 = self._find_col(tgt_df, ['产品69码（必填）', '产品69码', '69码', '条码', 'barcode'])

        self.src_code = self._find_col(src_df, ['编码', '商品编码', '商家编码', 'sku', 'code'])
        self.tgt_code = self._find_col(tgt_df, ['编码', '商品编码', '商家编码', 'sku', 'code'])

    @staticmethod
    def _find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
        for cand in candidates:
            for col in df.columns:
                if cand.lower() == str(col).strip().lower() or cand in str(col):
                    return col
        return None

    def match_row(self, tgt_row: pd.Series) -> Optional[int]:
        """
        Finds matching source DataFrame index for the given target row.
        Cascades through multi-level spec and link priorities.
        """
        link = tgt_row.get(self.tgt_link) if self.tgt_link else None
        spec = str(tgt_row.get(self.tgt_spec)).strip() if self.tgt_spec and pd.notna(tgt_row.get(self.tgt_spec)) else None
        name = str(tgt_row.get(self.tgt_name)).strip() if self.tgt_name and pd.notna(tgt_row.get(self.tgt_name)) else None
        c69 = tgt_row.get(self.tgt_69) if self.tgt_69 and pd.notna(tgt_row.get(self.tgt_69)) else None
        code = str(tgt_row.get(self.tgt_code)).strip() if self.tgt_code and pd.notna(tgt_row.get(self.tgt_code)) else None

        # Level 1: (Link, Spec)
        if link and spec and self.src_link and self.src_spec:
            sub = self.src_df[(self.src_df[self.src_link] == link) & (self.src_df[self.src_spec].astype(str).str.strip() == spec)]
            if len(sub) > 0:
                return sub.index[0]

        # Level 2: (Link, Name)
        if link and name and self.src_link and self.src_name:
            sub = self.src_df[(self.src_df[self.src_link] == link) & (self.src_df[self.src_name].astype(str).str.strip() == name)]
            if len(sub) > 0:
                return sub.index[0]

        # Level 3: (69 Code, Spec)
        if c69 and spec and self.src_69 and self.src_spec:
            sub = self.src_df[(self.src_df[self.src_69] == c69) & (self.src_df[self.src_spec].astype(str).str.strip() == spec)]
            if len(sub) > 0:
                return sub.index[0]

        # Level 4: 69 Code only
        if c69 and self.src_69:
            sub = self.src_df[self.src_df[self.src_69] == c69]
            if len(sub) > 0:
                return sub.index[0]

        # Level 5: (Code, Spec)
        if code and spec and self.src_code and self.src_spec:
            sub = self.src_df[(self.src_df[self.src_code].astype(str).str.strip() == code) & (self.src_df[self.src_spec].astype(str).str.strip() == spec)]
            if len(sub) > 0:
                return sub.index[0]

        # Level 6: Link only
        if link and self.src_link:
            sub = self.src_df[self.src_df[self.src_link] == link]
            if len(sub) > 0:
                return sub.index[0]

        # Level 7: Name only
        if name and self.src_name:
            sub = self.src_df[self.src_df[self.src_name].astype(str).str.strip() == name]
            if len(sub) > 0:
                return sub.index[0]

        return None


class CascadingMatcher:
    """
    Builds fast lookup indices for multi-key cascade matching.
    Tries key 1; if missing/unmatched, falls back to key 2, then key 3, etc.
    """
    def __init__(self, target_df: pd.DataFrame, key_columns: List[str]):
        self.key_columns = key_columns
        self.indices = {}
        for col in key_columns:
            if col in target_df.columns:
                mapping = {}
                for idx, val in target_df[col].items():
                    if pd.notna(val):
                        s_val = str(val).strip()
                        if s_val and s_val not in ('', 'nan', 'none'):
                            if s_val not in mapping:
                                mapping[s_val] = []
                            mapping[s_val].append(idx)
                self.indices[col] = mapping

    def match_row(self, row: pd.Series) -> Optional[List[int]]:
        """Returns target row index(es) for the given source row, or None."""
        for col in self.key_columns:
            if col in self.indices and col in row.index:
                val = row[col]
                if pd.notna(val):
                    s_val = str(val).strip()
                    if s_val in self.indices[col]:
                        return self.indices[col][s_val]
        return None
