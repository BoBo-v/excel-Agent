---
name: wps-excel-toolkit
description: >-
  Universal toolkit for WPS Office and Microsoft Office Excel operations.
  Use this skill whenever the user asks to process, modify, compare, transfer
  columns or on/off-shelf status, filter rows (e.g. margin >= 15%), or fill/recover embedded product images
  (WPS DISPIMG or Office drawings/RichData) in Excel (.xlsx) files. Never write ad-hoc
  Python scratch scripts for these tasks; always invoke this toolkit directly via cli.py.
---

# Universal WPS & Excel Toolkit (电商表格自动化套件)

This skill provides a high-performance, lossless, dual-engine CLI and Python library specifically designed for e-commerce tables, supporting WPS Office (`DISPIMG` embedded images) and Microsoft Office / Tencent Docs (`Drawings` / `RichData`).

## Location of CLI Entry Point
- Tool Path: `C:\Users\q1894\.gemini\config\skills\wps-excel-toolkit\scripts\cli.py`

## Core Command Reference

### 1. Lossless Row Filtering (`filter`)
Use when filtering rows based on condition (e.g. `毛利率 >= 15%`, `京东价格 < 50`, `品牌 == 半亩花田`).
Guarantees **100% zero image loss** (Office RichData, WPS cellimages, and drawings) and executes on 100MB+ workbooks in ~3-5 seconds.

```powershell
# Filter by expression (strictly >= 15%)
python "C:\Users\q1894\.gemini\config\skills\wps-excel-toolkit\scripts\cli.py" filter `
    --file "path/to/table.xlsx" `
    --expr "毛利率 >= 15%" `
    --output "path/to/output.xlsx"

# Filter by expression and keep empty/None margin rows
python "C:\Users\q1894\.gemini\config\skills\wps-excel-toolkit\scripts\cli.py" filter `
    --file "path/to/table.xlsx" `
    --expr "毛利率 >= 15%" `
    --keep-empty `
    --output "path/to/output.xlsx"
```

### 2. Transfer Columns / Shelf Status (`transfer`)
Use when transferring processed status (e.g. `是否已下架`, `检测备注`, `实际售价`) from a processed source table to a target master table.
Automatically shifts right-hand formula references (`(L2-N2)/L2` -> `(M2-O2)/M2`) and preserves all 6,000+ embedded images.

```powershell
python "C:\Users\q1894\.gemini\config\skills\wps-excel-toolkit\scripts\cli.py" transfer `
    --source "path/to/source.xlsx" `
    --target "path/to/target.xlsx" `
    --sheets auto `
    --cols auto `
    --keys auto `
    --insert-after "供应商产品\n唯一识别码"
```
*Note: `--cols auto` automatically detects all columns present in source that are missing in target.*

### 3. Supplement Missing Images (`fill-images`)
Use when product images in the target table are missing or broken (`#VALUE!`), and need to be replenished from one or more source tables.
**Dual-Engine Support**: Seamlessly supports both **WPS `DISPIMG`** cell images and **Microsoft Office / Tencent Docs `RichData` (XLRICHVALUE)**.
**Multi-Spec Matching**: Automatically handles multi-specification products by cascading `(链接 + 规格)` -> `(链接 + 名称)` -> `(69码 + 规格)` -> `(编码 + 规格)` to prevent mismatched pictures.

```powershell
# Automatic dual-engine detection & multi-spec cascade matching:
python "C:\Users\q1894\.gemini\config\skills\wps-excel-toolkit\scripts\cli.py" fill-images `
    --target "path/to/target.xlsx" `
    --sources "path/to/source1.xlsx" "path/to/source2.xlsx"

# Specify custom image column or output path:
python "C:\Users\q1894\.gemini\config\skills\wps-excel-toolkit\scripts\cli.py" fill-images `
    --target "path/to/target.xlsx" `
    --sources "path/to/source.xlsx" `
    --image-col "商品图片" `
    --output "path/to/filled.xlsx"
```

### 4. Reconcile / Diff Tables (`diff`)
Use when comparing two versions of a product catalog (e.g. comparing this month vs last month for price changes, deleted items, or off-shelf products).

```powershell
python "C:\Users\q1894\.gemini\config\skills\wps-excel-toolkit\scripts\cli.py" diff `
    --base "path/to/old_table.xlsx" `
    --compare "path/to/new_table.xlsx" `
    --keys "京东链接" "产品69码" `
    --watch-cols "京东价格" "含税运进价" "是否已下架" `
    --out-report "path/to/diff_report.xlsx"
```

### 5. Detect Excel Flavor (`sniff`)
Checks if a workbook uses WPS DISPIMG, Microsoft RichData, Drawings, or plain data:

```powershell
python "C:\Users\q1894\.gemini\config\skills\wps-excel-toolkit\scripts\cli.py" sniff `
    --file "path/to/table.xlsx"
```

## Mandatory Rules for Agent Execution
1. Always run operations through `cli.py`. Do NOT write custom ad-hoc scripts.
2. The engine automatically handles WPS file locking: if the target is currently open in WPS/Excel, it will safely save to `_已同步.xlsx` instead of crashing.
