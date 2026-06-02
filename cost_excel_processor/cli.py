"""
命令行接口
支持命令行直接调用处理程序
用法：
    python -m cost_excel_processor.cli <输入Excel文件路径> [输出Excel文件路径]
"""

import sys
import os

# 将项目根目录加入搜索路径，确保绝对导入可用
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from cost_excel_processor.main_processor import CostExcelProcessor


def main():
    if len(sys.argv) < 2:
        print("用法：python -m cost_excel_processor.cli <输入Excel> [输出Excel]")
        print("\n示例：")
        print("  python -m cost_excel_processor.cli ./data/清单.xlsx")
        print("  python -m cost_excel_processor.cli ./data/清单.xlsx ./output/结果.xlsx")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(input_path):
        print(f"错误：文件不存在 - {input_path}")
        sys.exit(1)

    print("=" * 60)
    print("  造价Excel清单自动化处理程序 v0.1.0")
    print("=" * 60)

    processor = CostExcelProcessor()
    processor.load(input_path)
    processor.analyze()
    result_df = processor.process(interactive=False)
    saved_path = processor.save(output_path)

    print("\n" + "=" * 60)
    print("  处理完成！")
    print(f"  输出文件：{saved_path}")
    print(f"  共 {len(result_df)} 行 × {len(result_df.columns)} 列")
    print("=" * 60)

    print("\n--- 数据预览（前5行） ---")
    print(result_df.head(5).to_string(index=False))


if __name__ == "__main__":
    main()
