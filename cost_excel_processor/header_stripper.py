"""
表头拆解器（精简版）— 只拆解多级表头，输出为一行
不做字段校验、不做数据整合、不填充业态、不插入序号
"""

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
import pandas as pd
from excel_reader import ExcelReader
from header_parser import HeaderParser


def parse_headers_only(input_path: str, output_path: str = None):
    """
    纯表头拆解：读取 Excel，对每个工作表的多级表头拆成一行
    输出为 Excel 文件，每个工作表一行，列名为该工作表拆解后的表头
    """
    # 加载文件
    reader = ExcelReader(input_path)
    reader.load()
    parser = HeaderParser(reader.wb)

    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = base + "_表头拆解.xlsx"

    # 用于收集所有工作表的表头
    all_headers = {}  # {sheet_name: [header1, header2, ...]}
    max_cols = 0

    print(f"文件: {os.path.basename(input_path)}")
    print(f"工作表数: {len(reader.sheet_names)}")
    print(f"{'='*80}")

    for sheet_name in reader.sheet_names:
        ws = reader.wb[sheet_name]
        info = reader._analyze_sheet(ws)

        # 跳过空表
        if info["is_empty"]:
            print(f"  [{sheet_name}] → 空表，跳过")
            continue

        first_header_row = info.get("first_header_row", 1)
        header_rows = info.get("header_rows", 1)

        # 纯表头拆解（不做标准化、不校验、不填充）
        parsed_headers = parser.parse_sheet_headers(ws, header_rows, first_header_row)

        # 清理空列尾
        while parsed_headers and not parsed_headers[-1]:
            parsed_headers.pop()

        all_headers[sheet_name] = parsed_headers
        max_cols = max(max_cols, len(parsed_headers))

        print(f"  [{sheet_name}]")
        print(f"    表头起始行: {first_header_row}, 表头行数: {header_rows}")
        print(f"    拆解后列数: {len(parsed_headers)}")
        # 打印前8列名
        preview = [h[:30] if h else "(空)" for h in parsed_headers[:8]]
        print(f"    列名预览: {preview}")
        if len(parsed_headers) > 8:
            print(f"    ... 共 {len(parsed_headers)} 列")
        print()

    # 构建输出 DataFrame：每行一个工作表
    rows = []
    for sheet_name, headers in all_headers.items():
        row = [sheet_name]  # 第一列为工作表名
        # 补齐到 max_cols
        padded = headers + [""] * (max_cols - len(headers))
        row.extend(padded)
        rows.append(row)

    col_names = ["工作表名"] + [f"列{i}" for i in range(1, max_cols + 1)]
    df = pd.DataFrame(rows, columns=col_names)

    # 保存
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Sheet 1: 表头拆解结果
        df.to_excel(writer, sheet_name="表头拆解", index=False)
        ws_out = writer.sheets["表头拆解"]
        for col_idx, col in enumerate(ws_out.columns, 1):
            max_len = 0
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws_out.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

        # Sheet 2: 转置视图（便于查看长表头）
        df_t = df.set_index("工作表名").T
        df_t.to_excel(writer, sheet_name="转置视图")
        ws_t = writer.sheets["转置视图"]
        for col_idx, col in enumerate(ws_t.columns, 1):
            max_len = 0
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws_t.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

    print(f"输出: {output_path}")
    print(f"  Sheet1 '表头拆解': {len(rows)} 个工作表 × {max_cols + 1} 列")
    print(f"  Sheet2 '转置视图': {max_cols} 列名 × {len(rows)} 个工作表")
    return output_path


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: python header_stripper.py <Excel文件路径> [输出路径]")
    else:
        input_path = sys.argv[1]
        output_path = sys.argv[2] if len(sys.argv) > 2 else None
        parse_headers_only(input_path, output_path)
