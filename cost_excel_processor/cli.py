"""
命令行接口
支持命令行直接调用处理程序

用法（标准管线）：
    python -m cost_excel_processor.cli <输入Excel文件路径> [输出Excel文件路径]

用法（清单合并模式）：
    python -m cost_excel_processor.cli --merge <输入Excel1> [输入Excel2] ... [--merge-output PATH]
    python -m cost_excel_processor.cli --merge --folder <文件夹路径> [--merge-output PATH]

示例：
    python -m cost_excel_processor.cli ./data/清单.xlsx
    python -m cost_excel_processor.cli ./data/清单.xlsx ./output/结果.xlsx
    python -m cost_excel_processor.cli --merge ./1.1.xlsx ./1.2.xlsx
    python -m cost_excel_processor.cli --merge --folder ./data/
"""

import sys
import os
import argparse

# 将项目根目录加入搜索路径，确保绝对导入可用
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from cost_excel_processor.main_processor import CostExcelProcessor


def build_parser():
    """构建命令行参数解析器"""
    parser = argparse.ArgumentParser(
        description="造价Excel清单自动化处理程序 v0.1.1",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  标准管线：
    %(prog)s 输入.xlsx [输出.xlsx]
  清单合并：
    %(prog)s --merge 输入1.xlsx [输入2.xlsx ...]
    %(prog)s --merge --folder ./data/
        """,
    )

    # 位置参数：输入文件（兼容旧用法）
    parser.add_argument(
        "input_files", nargs="*", default=[],
        help="输入 Excel 文件路径（可指定多个，用于合并模式）",
    )

    # 可选参数
    parser.add_argument(
        "-o", "--output",
        dest="output_path",
        help="输出 Excel 文件路径",
    )

    # ---- 清单合并相关参数 ----
    parser.add_argument(
        "--merge",
        action="store_true",
        dest="enable_merge",
        help="启用清单合并模式（跳过标准管线，直接合并所有清单表）",
    )
    parser.add_argument(
        "--merge-output",
        dest="merge_output_path",
        help="清单合并输出文件路径（默认自动生成）",
    )
    parser.add_argument(
        "--folder",
        dest="folder_path",
        help="扫描指定文件夹下所有 Excel 文件用于合并（与 --merge 配合使用）",
    )
    parser.add_argument(
        "--file-mode",
        choices=["folder", "dialog", "auto"],
        default="auto",
        dest="file_mode",
        help="文件来源模式：folder=扫描文件夹 / dialog=手动选择 / auto=自动（默认）",
    )

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    # ============================================================
    # 模式 A：--merge 清单合并模式
    # ============================================================
    if args.enable_merge:
        from file_scanner import discover_list_files
        from audit_logger import AuditLogger
        from sheet_merger import merge_sheets

        print("=" * 60)
        print("  造价Excel清单自动化处理程序 v0.1.1 — 合并模式")
        print("=" * 60)

        # 确定输入文件列表
        input_files = []

        if args.folder_path:
            # 模式：指定文件夹扫描
            mode = "folder"
            try:
                input_files = discover_list_files(mode, folder_path=args.folder_path)
            except FileNotFoundError as e:
                print(f"错误：{e}")
                sys.exit(1)
        elif args.input_files:
            # 模式：命令行传入的文件列表
            for f in args.input_files:
                if not os.path.exists(f):
                    print(f"错误：文件不存在 - {f}")
                    sys.exit(1)
            input_files = list(args.input_files)
            mode = "list"
        else:
            # 模式：自动发现（先尝试当前目录）
            cwd = os.getcwd()
            mode = "auto"
            try:
                input_files = discover_list_files(mode, folder_path=cwd)
            except (FileNotFoundError, NotADirectoryError):
                pass

        if not input_files:
            print(f"未找到任何 Excel 文件。")
            print(f"提示：请用 --folder 指定文件夹，或直接传入文件路径。")
            sys.exit(1)

        print(f"\n  输入文件数：{len(input_files)}")

        # 执行合并
        logger = AuditLogger()
        result = merge_sheets(
            input_files=input_files,
            output_path=args.merge_output_path,
            audit_logger=logger,
        )

        if result.get("output_file"):
            print(f"\n{'='*60}")
            print(f"  处理完成！")
            print(f"  输出文件：{result['output_file']}")
            print(f"  共 {result.get('rows', 0)} 行 × {result.get('cols', 0)} 列")
            print(f"  清单表数：{len(result.get('list_sheets', []))}")
            print("=" * 60)
        else:
            print("\n[WARN] 未生成输出文件（可能未找到清单表）")

        return

    # ============================================================
    # 模式 B：标准处理管线（默认）
    # ============================================================
    if not args.input_files:
        parser.print_help()
        sys.exit(1)

    input_path = args.input_files[0]
    output_path = args.output_path

    if not os.path.exists(input_path):
        print(f"错误：文件不存在 - {input_path}")
        sys.exit(1)

    print("=" * 60)
    print("  造价Excel清单自动化处理程序 v0.1.1 — 标准管线")
    print("=" * 60)

    processor = CostExcelProcessor()

    # 如果传入了多个文件且启用了合并，配置合并参数
    if len(args.input_files) > 1:
        processor.merge_input_files = list(args.input_files)

    processor.load(input_path)
    processor.analyze()
    result_df = processor.process(interactive=False)
    saved_path = processor.save(output_path)

    # 如果启用合并，执行合并步骤
    if processor.enable_merge and len(args.input_files) > 1:
        print("\n--- 执行可选的清单合并步骤 ---")
        merge_result = processor.run_merge(args.merge_output_path)
        if merge_result.get("output_file"):
            print(f"  合并输出：{merge_result['output_file']}")

    print("\n" + "=" * 60)
    print("  处理完成！")
    print(f"  输出文件：{saved_path}")
    print(f"  共 {len(result_df)} 行 × {len(result_df.columns)} 列")
    print("=" * 60)

    print("\n--- 数据预览（前5行） ---")
    print(result_df.head(5).to_string(index=False))


if __name__ == "__main__":
    main()
