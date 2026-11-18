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

from engine import detect_excel_flavor, create_backup
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
    p_filter.add_argument("--sheet", default=None, help="工作表名称 (默认第 1 个工作表)")
    p_filter.add_argument("--output", "-o", default=None, help="输出文件路径 (默认安全另存/覆盖并备份)")
    p_filter.add_argument("--no-backup", action="store_true", help="不生成备份文件")

    # 5. sniff
    p_sniff = subparsers.add_parser("sniff", help="探测表格是 WPS 还是微软 Office 格式")
    p_sniff.add_argument("--file", "-f", required=True, help="待检测的 Excel 文件路径")

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
            sheet=args.sheet,
            output_path=args.output,
            no_backup=args.no_backup
        )
    elif args.subcommand == "sniff":
        flavor = detect_excel_flavor(args.file)
        print(f"表格格式探测结果: {flavor}")


if __name__ == "__main__":
    main()
