import sys
import os
import io
import argparse

# Set UTF-8 stdout
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Ensure script dir is on sys.path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from engine import (
    detect_excel_flavor,
    create_backup,
    list_sheets,
    create_sheet,
    rename_sheet,
    copy_sheet,
    delete_sheet,
    set_sheet_visibility,
    scan_workbook_formula_errors
)
from ops import run_transfer, run_fill_images, run_diff, run_filter


def main():
    parser = argparse.ArgumentParser(
        prog="wps-excel-toolkit",
        description="通用型 WPS / 微软 Office 电商表格自动化操作工具集 (Universal Excel Toolkit)"
    )
    subparsers = parser.add_subparsers(dest="subcommand", help="支持的操作命令")

    # 1. transfer
    p_transfer = subparsers.add_parser("transfer", help="跨表字段/状态同步与公式智能平移")
    p_transfer.add_argument("--source", "-s", required=True, help="来源表路径 (含更新数据)")
    p_transfer.add_argument("--target", "-t", required=True, help="目标表路径 (大底表)")
    p_transfer.add_argument("--output", "-o", default=None, help="输出文件路径 (默认覆写目标表并自动备份)")
    p_transfer.add_argument("--sheets", nargs="+", default=["auto"], help="指定工作表名 (默认 auto 自动匹配同名表)")
    p_transfer.add_argument("--cols", nargs="+", default=["auto"], help="待同步字段列表 (默认 auto 自动提取新增列)")
    p_transfer.add_argument("--keys", nargs="+", default=["auto"], help="匹配主键链 (默认 auto 级联探测链接/69码/编码/名称)")
    p_transfer.add_argument("--insert-after", default=None, help="新字段插入锚点 (列名称)")
    p_transfer.add_argument("--branch-col", default=None, help="分表派发字段名 (如: 所属分表)")
    p_transfer.add_argument("--row-num-col", default=None, help="原表行号字段名 (如: 原表行号)")
    p_transfer.add_argument("--no-backup", action="store_true", help="不生成备份文件")

    # 2. fill-images
    p_images = subparsers.add_parser("fill-images", help="跨表检索并补齐嵌入单元格图片 (双引擎: WPS DISPIMG + Office/腾讯文档 RichData)")
    p_images.add_argument("--target", "-t", required=True, help="缺失图片的目标表")
    p_images.add_argument("--sources", "-s", nargs="+", required=True, help="图源表路径列表")
    p_images.add_argument("--output", "-o", default=None, help="输出文件路径")
    p_images.add_argument("--image-col", default=None, help="图片所在列名称 (默认 auto 自动探测 '商品图片' 等)")
    p_images.add_argument("--keys", nargs="+", default=["auto"], help="指定主键/规格字段匹配链 (默认 auto 自动多规格级联)")
    p_images.add_argument("--no-backup", action="store_true", help="不生成备份文件")

    # 3. diff
    p_diff = subparsers.add_parser("diff", help="双表数据比对核对 (如价格变动、上下架变动)")
    p_diff.add_argument("--base", "-b", required=True, help="基准表路径")
    p_diff.add_argument("--compare", "-c", required=True, help="待比对新表路径")
    p_diff.add_argument("--keys", nargs="+", default=["auto"], help="主键列表")
    p_diff.add_argument("--watch-cols", nargs="+", default=["auto"], help="需要重点监控变化的列")
    p_diff.add_argument("--out-report", default=None, help="生成差异报告 Excel 文件路径")

    # 4. filter
    p_filter = subparsers.add_parser("filter", help="物理级零损行数据条件过滤 (完整保留Office/WPS所有嵌入图片与格式)")
    p_filter.add_argument("--file", "-f", required=True, help="待过滤的 Excel 表格路径")
    p_filter.add_argument("--expr", "-e", default=None, help="过滤表达式 (如: '毛利率 >= 0.15' 或 '品牌 == 半亩花田')")
    p_filter.add_argument("--col", "-c", default=None, help="过滤目标列名或列标 (如: 毛利率 或 O)")
    p_filter.add_argument("--op", default=">=", choices=[">=", "<=", ">", "<", "==", "!=", "contains", "not_contains"], help="比较操作符")
    p_filter.add_argument("--val", "-v", default=None, help="比较阈值或目标字符串 (如: 0.15)")
    p_filter.add_argument("--keep-empty", action="store_true", help="是否保留目标单元格为空的行 (默认剔除空值)")
    p_filter.add_argument("--preserve-summary", action="store_true", default=True, help="智能识别并保护表格末尾的合计/小计行 (默认开启)")
    p_filter.add_argument("--no-preserve-summary", dest="preserve_summary", action="store_false", help="关闭合计行保护，全量按条件过滤")
    p_filter.add_argument("--sheet", default=None, help="工作表名称 (默认第 1 个工作表)")
    p_filter.add_argument("--output", "-o", default=None, help="输出文件路径 (默认安全另存/覆盖并备份)")
    p_filter.add_argument("--no-backup", action="store_true", help="不生成备份文件")

    # 5. sniff
    p_sniff = subparsers.add_parser("sniff", help="探测表格是 WPS 还是微软 Office 格式")
    p_sniff.add_argument("--file", "-f", required=True, help="待检测的 Excel 文件路径")

    # 6. sheets
    p_sheets = subparsers.add_parser("sheets", help="工作表生命周期管理 (查看/新建/重命名/复制/删除/隐藏)")
    p_sheets.add_argument("--file", "-f", required=True, help="Excel 文件路径")
    p_sheets.add_argument("--action", "-a", required=True, choices=["list", "create", "rename", "copy", "delete", "hide", "unhide"], help="操作类型")
    p_sheets.add_argument("--name", "-n", default=None, help="工作表名称")
    p_sheets.add_argument("--new-name", default=None, help="新工作表名称 (重命名或复制时使用)")
    p_sheets.add_argument("--output", "-o", default=None, help="另存文件路径 (默认覆写原表)")

    # 7. audit
    p_audit = subparsers.add_parser("audit", help="全表公式语法审计与错误值扫描 (#REF!, #DIV/0!, #VALUE! 等)")
    p_audit.add_argument("--file", "-f", required=True, help="待审计的 Excel 文件路径")
    p_audit.add_argument("--sheet", "-s", default=None, help="指定审计的工作表名称 (默认全工作簿扫描)")

    args = parser.parse_args()

    if not args.subcommand:
        parser.print_help()
        sys.exit(1)

    if args.subcommand == "transfer":
        run_transfer(
            source_path=args.source,
            target_path=args.target,
            output_path=args.output,
            sheets=args.sheets,
            cols=args.cols,
            keys=args.keys,
            insert_after=args.insert_after,
            branch_col=args.branch_col,
            row_num_col=args.row_num_col,
            no_backup=args.no_backup
        )
    elif args.subcommand == "fill-images":
        run_fill_images(
            target_path=args.target,
            source_paths=args.sources,
            output_path=args.output,
            image_col=args.image_col,
            keys=args.keys,
            no_backup=args.no_backup
        )
    elif args.subcommand == "diff":
        run_diff(
            base_path=args.base,
            compare_path=args.compare,
            keys=args.keys,
            watch_cols=args.watch_cols,
            output_report=args.out_report
        )
    elif args.subcommand == "filter":
        run_filter(
            file_path=args.file,
            col=args.col,
            op=args.op,
            val=args.val,
            expr=args.expr,
            keep_empty=args.keep_empty,
            preserve_summary=args.preserve_summary,
            sheet=args.sheet,
            output_path=args.output,
            no_backup=args.no_backup
        )
    elif args.subcommand == "sniff":
        flavor = detect_excel_flavor(args.file)
        print(f"表格格式探测结果: {flavor}")
    elif args.subcommand == "sheets":
        act = args.action
        if act == "list":
            s_list = list_sheets(args.file)
            print(f"[工作表列表] 共 {len(s_list)} 个工作表:")
            for s in s_list:
                status_tag = f"[{s['state']}]" if s['state'] != 'visible' else ""
                print(f"  - [{s['index']}] '{s['name']}' {status_tag} (最大行: {s['max_row']}, 最大列: {s['max_column']})")
        elif act == "create":
            if not args.name:
                raise ValueError("新建工作表必须通过 --name 指定名称")
            out = create_sheet(args.file, args.name, output_path=args.output)
            print(f"[完成] 已新建工作表 '{args.name}' -> {out}")
        elif act == "rename":
            if not args.name or not args.new_name:
                raise ValueError("重命名工作表必须指定 --name (原名称) 和 --new-name (新名称)")
            out = rename_sheet(args.file, args.name, args.new_name, output_path=args.output)
            print(f"[完成] 已重命名工作表 '{args.name}' -> '{args.new_name}' ({out})")
        elif act == "copy":
            if not args.name or not args.new_name:
                raise ValueError("复制工作表必须指定 --name (原名称) 和 --new-name (副本名称)")
            out = copy_sheet(args.file, args.name, args.new_name, output_path=args.output)
            print(f"[完成] 已复制工作表 '{args.name}' -> '{args.new_name}' ({out})")
        elif act == "delete":
            if not args.name:
                raise ValueError("删除工作表必须通过 --name 指定名称")
            out = delete_sheet(args.file, args.name, output_path=args.output)
            print(f"[完成] 已删除工作表 '{args.name}' -> {out}")
        elif act in ("hide", "unhide"):
            if not args.name:
                raise ValueError("隐藏/显示工作表必须通过 --name 指定名称")
            visible = (act == "unhide")
            out = set_sheet_visibility(args.file, args.name, visible=visible, output_path=args.output)
            action_desc = "取消隐藏" if visible else "隐藏"
            print(f"[完成] 已{action_desc}工作表 '{args.name}' -> {out}")
    elif args.subcommand == "audit":
        errs = scan_workbook_formula_errors(args.file, target_sheet=args.sheet)
        if not errs:
            print("[审计结果] 未发现任何公式错误值或语法断链 (#REF!, #DIV/0!, #VALUE! 等)，表格健康！")
        else:
            print(f"[审计结果] 发现 {len(errs)} 处公式/计算异常:")
            for e in errs:
                f_info = f" | 公式: {e['formula']}" if e['formula'] else ""
                print(f"  - [{e['sheet']}] 单元格 {e['cell']}: 错误类型 {e['error_type']}{f_info} -> {e['description']}")


if __name__ == "__main__":
    main()
