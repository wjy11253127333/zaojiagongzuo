"""
----------------------------------------------
清单表合并脚本 v1.0
----------------------------------------------
功能：
  1. 扫描工作簿，识别清单表（表头命中 >= 4 个关键字）
  2. 为每个清单表添加"业态"列（值=工作表名简称）
  3. pd.concat 合并（同名纵向拼接，异名横向新增列）
  4. 序号统一重新编号 1-N
  5. Unnamed 列自动重命名为"扩展列_N"
  6. 生成"合并说明"Sheet（列审计：同名列/异名列/空列）
  7. 格式标准化（深蓝表头、微软雅黑10pt、全表细线边框）

用法：
    python sheet_merger.py "input.xlsx"

输出：
    在原文件新增「清单汇总」Sheet +「合并说明」Sheet
    若原文件被占用，另存为 _含清单汇总.xlsx
"""

import os
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from utils import normalize_header_text

# ===== 配置 =====
LIST_KEYWORDS    = ["项目特征", "项目名称", "项目编码", "计量单位", "工程量"]
MIN_HIT          = 4
OUTPUT_SHEET     = "清单汇总"
REPORT_SHEET     = "合并说明"

# 样式常量
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
HDR_FILL  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HDR_FONT  = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
DATA_FONT  = Font(name="微软雅黑", size=10)
LEFT_ALIGN = Alignment(horizontal="left",    vertical="center", wrap_text=True)
CTR_ALIGN  = Alignment(horizontal="center",  vertical="center")
NUM_ALIGN  = Alignment(horizontal="right",   vertical="center")


# ============================================================
# 工作表识别
# ============================================================

def is_list_sheet(ws, max_scan=10) -> bool:
    """判断工作表是否为清单表（表头含 >= MIN_HIT 个关键字）"""
    hdr_row = get_header_row(ws, max_scan)
    if hdr_row is None:
        return False
    headers = [ws.cell(row=hdr_row, column=c).value or ""
               for c in range(1, ws.max_column + 1)]
    hdr_text = "".join(normalize_header_text(str(h)) for h in headers if h)
    return sum(1 for kw in LIST_KEYWORDS if kw in hdr_text) >= MIN_HIT


def get_header_row(ws, max_scan=10) -> int | None:
    """自动检测表头行（跳过前置标题行）"""
    # 先跳过内容超长的行（通常是标题行，如 >35字）
    start_row = 1
    for row in range(1, min(ws.max_row + 1, max_scan + 1)):
        cells = [str(ws.cell(row=row, column=c).value or "").strip()
                 for c in range(1, min(ws.max_column + 1, 10))]
        if any(len(c) > 35 for c in cells):
            continue
        start_row = row
        break

    best_row, best_hits = start_row, 0
    for row in range(start_row, min(ws.max_row + 1, start_row + max_scan)):
        cells = [str(ws.cell(row=row, column=c).value or "")
                 for c in range(1, ws.max_column + 1)]
        hdr_text = "".join(normalize_header_text(c) for c in cells if c)
        hits = sum(1 for kw in LIST_KEYWORDS if kw in hdr_text)
        if hits > best_hits:
            best_hits, best_row = hits, row

    return best_row if best_hits >= MIN_HIT else None


def get_sheet_headers(ws):
    """读取工作表的表头行，返回 (header_row, headers_list)"""
    hdr_row = get_header_row(ws)
    if hdr_row is None:
        return None, []
    headers = []
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=hdr_row, column=c).value
        headers.append(str(val).strip() if val else "")
    # 去除尾部空列
    while headers and not headers[-1]:
        headers.pop()
    return hdr_row, headers


# ============================================================
# Unnamed 列重命名
# ============================================================

def _rename_unnamed_columns(df: pd.DataFrame) -> pd.DataFrame:
    """将 Unnamed 列动态重命名为 扩展列_N"""
    new_cols = []
    unnamed_idx = 0
    for col in df.columns:
        col_str = str(col).strip()
        if col_str.startswith("Unnamed:") or col_str == "" or col_str == "nan":
            unnamed_idx += 1
            new_cols.append(f"扩展列_{unnamed_idx}")
        else:
            new_cols.append(col_str)
    df.columns = new_cols
    return df


# ============================================================
# 表头重复过滤
# ============================================================

def _is_header_duplicate(df: pd.DataFrame, header_col: str, threshold: float = 0.7) -> bool:
    """判断DataFrame中某列是否有 >= threshold 的行与该列名重复"""
    if header_col not in df.columns:
        return False
    col_values = df[header_col].dropna().astype(str)
    if len(col_values) == 0:
        return False
    match_count = (col_values == header_col).sum()
    return match_count / len(col_values) >= threshold


# ============================================================
# 列审计 + 合并说明
# ============================================================

def _generate_merge_report(all_dfs: list, merged: pd.DataFrame,
                           sheet_names: list) -> pd.DataFrame:
    """逐列分析来源：同名列（>=2表）vs 异名列（仅1表）"""
    report_rows = []
    for col in merged.columns:
        source_sheets = []
        has_data_sheets = []
        for i, df in enumerate(all_dfs):
            if col in df.columns:
                source_sheets.append(sheet_names[i])
                # 检查是否有数据（非空值）
                non_null = df[col].dropna()
                if len(non_null) > 0:
                    has_data_sheets.append(sheet_names[i])

        source_count = len(source_sheets)
        col_type = "同名列" if source_count >= 2 else "异名列"
        coverage = ", ".join(has_data_sheets) if has_data_sheets else "无"

        # 检查是否为空列
        merged_non_null = merged[col].dropna() if col in merged.columns else pd.Series()
        is_empty = len(merged_non_null) == 0
        action = "物理删除" if is_empty else "保留"

        report_rows.append({
            "列名": col,
            "来源类型": col_type,
            "来源表": ", ".join(source_sheets) if source_sheets else "无",
            "有数据表": coverage,
            "行覆盖率": f"{len(merged_non_null)}/{len(merged)}" if not is_empty else "0",
            "处理": action,
        })

    return pd.DataFrame(report_rows)


# ============================================================
# 格式标准化
# ============================================================

def _apply_list_merge_formatting(ws):
    """为「清单汇总」Sheet 应用格式标准"""
    max_row = ws.max_row
    max_col = ws.max_column

    # 判断列类型
    numeric_keywords = ["工程量", "合价", "单价", "数量", "金额", "造价", "增减"]
    numeric_cols = set()
    center_cols = {1}  # 序号列居中
    for col_idx in range(1, max_col + 1):
        h = ws.cell(row=1, column=col_idx).value
        if h and any(kw in str(h) for kw in numeric_keywords):
            numeric_cols.add(col_idx)
        if h and "计量单位" in str(h):
            center_cols.add(col_idx)

    for row_idx in range(1, max_row + 1):
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = THIN_BORDER

            if row_idx == 1:
                # 表头行
                cell.font = HDR_FONT
                cell.fill = HDR_FILL
                cell.alignment = HDR_ALIGN
            else:
                # 数据行
                cell.font = DATA_FONT
                if col_idx in center_cols:
                    cell.alignment = CTR_ALIGN
                elif col_idx in numeric_cols:
                    cell.alignment = NUM_ALIGN
                else:
                    cell.alignment = LEFT_ALIGN

            # 行高
            if row_idx == 1:
                ws.row_dimensions[row_idx].height = None  # 自适应
            elif row_idx >= 2:
                ws.row_dimensions[row_idx].height = 20

    # 列宽自适应（上限50）
    for col_idx in range(1, max_col + 1):
        max_width = 8
        for row_idx in range(1, min(max_row + 1, 100)):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val:
                max_width = max(max_width, min(len(str(val)) * 1.2, 50))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_width

    # 冻结首行
    ws.freeze_panes = "A2"


def _apply_report_formatting(ws, report_df: pd.DataFrame):
    """为「合并说明」Sheet 应用格式"""
    # 表头
    for col_idx in range(1, len(report_df.columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = HDR_ALIGN
        cell.border = THIN_BORDER
        cell.value = report_df.columns[col_idx - 1]

    # 数据
    for row_idx in range(len(report_df)):
        for col_idx in range(len(report_df.columns)):
            cell = ws.cell(row=row_idx + 2, column=col_idx + 1)
            cell.border = THIN_BORDER
            cell.font = DATA_FONT
            val = report_df.iloc[row_idx, col_idx]
            if pd.notna(val):
                cell.value = val
            cell.alignment = CTR_ALIGN if col_idx in (1, 5) else LEFT_ALIGN

    # 列宽
    col_widths = [18, 12, 30, 25, 12, 12]
    for col_idx, w in enumerate(col_widths, 1):
        if col_idx <= len(report_df.columns):
            ws.column_dimensions[get_column_letter(col_idx)].width = w

    ws.freeze_panes = "A2"


# ============================================================
# 主处理函数
# ============================================================

def merge_sheets(input_path: str):
    """
    主处理函数：扫描清单表 → 加业态列 → 合并 → 格式标准化 → 写合并说明
    """
    print(f"\n{'='*60}")
    print(f"清单表合并 v1.0 — {os.path.basename(input_path)}")
    print(f"{'='*60}")

    # Step 1: 扫描工作表
    wb = openpyxl.load_workbook(input_path, data_only=True)
    all_sheets = wb.sheetnames
    list_sheets = []

    for sn in all_sheets:
        if sn in (OUTPUT_SHEET, REPORT_SHEET):
            continue
        ws = wb[sn]
        if is_list_sheet(ws):
            list_sheets.append(sn)

    if len(list_sheets) == 0:
        print(f"[WARN] 未找到清单表（需表头命中 >= {MIN_HIT} 个关键字）")
        wb.close()
        return

    print(f"\n识别到 {len(list_sheets)} 个清单表：")
    for sn in list_sheets:
        print(f"  - {sn}")

    # Step 2: 逐个读取清单表 DataFrame
    all_dfs = []
    sheet_short_names = []

    for sn in list_sheets:
        ws = wb[sn]
        hdr_row, headers = get_sheet_headers(ws)
        if hdr_row is None:
            continue

        # 读取数据（跳过表头行之前的所有行）
        df = pd.read_excel(
            input_path, sheet_name=sn,
            header=hdr_row - 1,  # pandas header 是0-based
        )
        # 只保留有效列名（非空列）
        valid_cols = [c for c in df.columns if str(c).strip() not in ("", "nan")]
        df = df[valid_cols].copy()

        # 重命名 Unnamed 列
        df = _rename_unnamed_columns(df)

        # 过滤重复表头行
        for col in list(df.columns):
            if _is_header_duplicate(df, col):
                df = df[df[col] != col]

        # 加业态列
        short = sn
        for prefix in ["表-08 分部分项工程和单价措施项目清单-", "表-08 ", "表"]:
            if short.startswith(prefix):
                short = short[len(prefix):]
                break
        df.insert(0, "业态", short)
        sheet_short_names.append(short)

        all_dfs.append(df)
        print(f"  [{sn}] {len(df)}行 x {len(df.columns)}列")

    wb.close()

    # Step 3: 合并
    merged = pd.concat(all_dfs, join="outer", ignore_index=True, sort=False)

    # 重新编号：先删源表自带的序号列，再统一插入1-based编号
    if "序号" in merged.columns:
        merged.drop(columns=["序号"], inplace=True)
    merged.insert(0, "序号", range(1, len(merged) + 1))

    print(f"\n合并结果：{merged.shape[0]}行 x {merged.shape[1]}列")

    # Step 4: 列审计
    report_df = _generate_merge_report(all_dfs, merged, list_sheets)

    # 删除空列
    cols_to_drop = report_df[report_df["处理"] == "物理删除"]["列名"].tolist()
    cols_to_drop = [c for c in cols_to_drop if c in merged.columns
                    and c not in ("序号", "业态")]
    if cols_to_drop:
        merged.drop(columns=cols_to_drop, inplace=True)
        print(f"物理删除 {len(cols_to_drop)} 个空列：{[c[:20] for c in cols_to_drop[:5]]}...")

    print(f"最终：{merged.shape[0]}行 x {merged.shape[1]}列")

    # Step 5: 写入 Excel
    output_path = input_path
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl",
                            mode="a", if_sheet_exists="replace") as writer:
            merged.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
            report_df.to_excel(writer, sheet_name=REPORT_SHEET, index=False)

        # 格式标准化
        wb = openpyxl.load_workbook(output_path)
        ws_output = wb[OUTPUT_SHEET]
        _apply_list_merge_formatting(ws_output)

        ws_report = wb[REPORT_SHEET]
        _apply_report_formatting(ws_report, report_df)

        wb.save(output_path)
        wb.close()
        print(f"\n[OK] 已保存到原文件：{os.path.basename(output_path)}")

    except PermissionError:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_含清单汇总{ext}"
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            merged.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
            report_df.to_excel(writer, sheet_name=REPORT_SHEET, index=False)

        wb = openpyxl.load_workbook(output_path)
        ws_output = wb[OUTPUT_SHEET]
        _apply_list_merge_formatting(ws_output)
        ws_report = wb[REPORT_SHEET]
        _apply_report_formatting(ws_report, report_df)
        wb.save(output_path)
        wb.close()
        print(f"\n[OK] 原文件被占用，已另存为：{os.path.basename(output_path)}")

    # Step 6: 汇总
    print(f"\n{'='*60}")
    print(f"处理完成！")
    print(f"  清单表数：{len(list_sheets)}")
    print(f"  输出：{os.path.basename(output_path)}")
    print(f"  Sheets：{OUTPUT_SHEET} ({merged.shape[0]}行x{merged.shape[1]}列)")
    print(f"         {REPORT_SHEET} ({len(report_df)}列说明)")
    print(f"{'='*60}\n")

    return {
        "list_sheets": list_sheets,
        "output_file": output_path,
        "rows": merged.shape[0],
        "cols": merged.shape[1],
        "dropped_cols": len(cols_to_drop),
    }


# ============================================================
# CLI 入口
# ============================================================

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python sheet_merger.py <输入Excel>")
    else:
        input_path = sys.argv[1]
        if not os.path.exists(input_path):
            print(f"[ERROR] 文件不存在：{input_path}")
            sys.exit(1)
        merge_sheets(input_path)
