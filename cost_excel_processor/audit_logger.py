"""
------------------------------------------------
全量审计日志模块 v1.0
------------------------------------------------
功能：
  1. AuditLogger 类 — 记录清单合并全流程操作日志
  2. 每条记录含：时间戳 / 操作类型 / 详情 / 文件名 / 工作表名 / 行数 / 列数 / 备注
  3. get_dataframe() — 输出为 DataFrame，直接写入"操作记录"Sheet

写入位置：清单合并输出文件的「操作记录」Sheet
"""

import pandas as pd
from datetime import datetime
from typing import List, Dict, Any, Optional
import os


# ============================================================
# AuditLogger 类
# ============================================================

class AuditLogger:
    """
    全量审计日志器

    用法：
        logger = AuditLogger()
        logger.log("文件扫描", "扫描到3个Excel文件", file="1.1综合楼.xlsx")
        df = logger.get_dataframe()
        df.to_excel(writer, sheet_name="操作记录", index=False)
    """

    # 类级别：操作记录表头（写入 Excel 时的列名）
    COLUMNS = ["时间戳", "操作类型", "详情", "文件名", "工作表名", "行数", "列数", "备注"]

    def __init__(self):
        """初始化审计日志器，records 为内存中的操作记录列表"""
        self.records: List[Dict[str, Any]] = []

    # --------------------------------------------------
    # 核心写入方法
    # --------------------------------------------------

    def log(
        self,
        action: str,
        detail: str = "",
        *,
        file: str = "",
        sheet: str = "",
        rows: Optional[int] = None,
        cols: Optional[int] = None,
        remark: str = "",
    ):
        """
        记录一条操作日志

        Args:
            action: 操作类型（如"文件扫描"、"清单表识别"、"表头对比"、"合并完成"、"写文件"）
            detail: 操作详细描述
            file:   相关文件名（短名，不含路径）
            sheet:  相关工作表名
            rows:   当前操作的行数（若有）
            cols:   当前操作的列数（若有）
            remark: 备注信息
        """
        record = {
            "时间戳": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "操作类型": action,
            "详情": detail,
            "文件名": file,
            "工作表名": sheet,
            "行数": "" if rows is None else rows,
            "列数": "" if cols is None else cols,
            "备注": remark,
        }
        self.records.append(record)

    # --------------------------------------------------
    # 专用快捷方法
    # --------------------------------------------------

    def log_file_scan(self, files: List[str], source: str = ""):
        """
        记录文件扫描结果

        Args:
            files:  扫描到的 Excel 文件路径列表
            source: 来源描述（如"文件夹扫描：C:/xxx"或"手动选择"）
        """
        detail = f"扫描到 {len(files)} 个 Excel 文件"
        if source:
            detail += f"（{source}）"
        self.log("文件扫描", detail, remark=source)

        # 逐文件记录
        for f in files:
            short = os.path.basename(f)
            self.log("文件扫描", f"发现文件：{short}", file=short, remark=f)

    def log_sheet_detect(
        self,
        file: str,
        sheet: str,
        headers: List[str],
        row_count: int,
        hit_keywords: List[str] | None = None,
    ):
        """
        记录清单表检测结果

        Args:
            file:         文件名（短名）
            sheet:        工作表全名
            headers:      表头列名列表
            row_count:    数据行数
            hit_keywords: 命中的关键词列表（可选）
        """
        col_count = len(headers) if headers else 0
        remark = f"表头：{col_count}列"
        if hit_keywords:
            remark += f" | 命中关键词：{', '.join(hit_keywords)}"
        self.log(
            "清单表识别",
            f"识别为清单表（{len(hit_keywords) if hit_keywords else '?'}个关键词命中）",
            file=file,
            sheet=sheet,
            rows=row_count,
            cols=col_count,
            remark=remark,
        )

    def log_non_list_sheet(self, file: str, sheet: str, reason: str = ""):
        """记录非清单表（跳过原因）"""
        self.log(
            "清单表识别",
            f"非清单表，已跳过",
            file=file,
            sheet=sheet,
            remark=reason or "未命中足够关键词",
        )

    def log_header_compare(self, all_headers: Dict[str, List[str]]):
        """
        记录表头对比结果（列级别差异）

        Args:
            all_headers: { "文件名_工作表名": [表头列名列表], ... }
        """
        if not all_headers:
            self.log("表头对比", "无清单表，跳过表头对比")
            return

        # 计算共同列
        all_col_sets = [set(h) for h in all_headers.values()]
        if all_col_sets:
            common = set.intersection(*all_col_sets)
            unique = set.union(*all_col_sets) - common
        else:
            common, unique = set(), set()

        detail = f"共同列：{len(common)}个"
        if unique:
            detail += f" | 差异列：{len(unique)}个"
        self.log("表头对比", detail, cols=len(common) + len(unique), remark=f"差异列：{', '.join(sorted(unique)) if unique else '无'}")

        # 逐列记录来源
        all_cols = sorted(set.union(*all_col_sets)) if all_col_sets else []
        for col in all_cols:
            source_sheets = [
                name for name, headers in all_headers.items()
                if col in headers
            ]
            col_type = "共同列" if len(source_sheets) >= 2 else "差异列"
            self.log(
                "列来源分析",
                f"列「{col}」→ {col_type}（来自 {len(source_sheets)} 个表）",
                cols=col,
                remark=f"来源：{', '.join(source_sheets[:3])}" + ("..." if len(source_sheets) > 3 else ""),
            )

    def log_merge_start(self, sheet_count: int):
        """记录合并开始"""
        self.log("合并开始", f"开始合并 {sheet_count} 个清单表")

    def log_merge_result(self, merged_df: pd.DataFrame, added_cols: List[str] | None = None):
        """
        记录合并结果

        Args:
            merged_df:   合并后的 DataFrame
            added_cols:   合并后新增的列名列表（差异列）
        """
        rows, cols = merged_df.shape
        remark = ""
        if added_cols:
            remark = f"新增差异列：{', '.join(added_cols)}"
        self.log(
            "合并完成",
            f"合并后数据：{rows}行 x {cols}列",
            rows=rows,
            cols=cols,
            remark=remark,
        )

    def log_output(self, output_path: str, sheet_names: List[str]):
        """记录写文件操作"""
        short = os.path.basename(output_path)
        self.log(
            "写文件",
            f"已保存到：{short}",
            remark=f"Sheets：{', '.join(sheet_names)} | 完整路径：{output_path}",
        )

    def log_format_apply(self, sheet_name: str):
        """记录格式标准化操作"""
        self.log("格式标准化", f"已应用格式：{sheet_name}", sheet=sheet_name)

    def log_error(self, action: str, error_msg: str, *, file: str = "", sheet: str = ""):
        """记录错误信息"""
        self.log("错误", error_msg, file=file, sheet=sheet, remark=f"错误发生在：{action}")

    # --------------------------------------------------
    # 输出方法
    # --------------------------------------------------

    def get_dataframe(self) -> pd.DataFrame:
        """
        将内存中的操作记录转换为 DataFrame
        用于直接写入 Excel「操作记录」Sheet

        Returns:
            操作记录 DataFrame，列顺序为 COLUMNS
        """
        if not self.records:
            # 返回空 DataFrame（含表头）
            return pd.DataFrame(columns=self.COLUMNS)

        df = pd.DataFrame(self.records)

        # 确保列顺序正确（补齐全列）
        for col in self.COLUMNS:
            if col not in df.columns:
                df[col] = ""
        return df[self.COLUMNS]

    def save_to_excel(self, writer: pd.ExcelWriter, sheet_name: str = "操作记录"):
        """
        直接将操作记录写入 Excel Writer（便捷方法）

        Args:
            writer:     pd.ExcelWriter 对象（已打开）
            sheet_name: 输出的 Sheet 名，默认"操作记录"
        """
        df = self.get_dataframe()
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    def clear(self):
        """清空所有记录（复用 logger 时调用）"""
        self.records.clear()


# ============================================================
# 独立测试
# ============================================================

if __name__ == "__main__":
    import os

    print("=" * 60)
    print("audit_logger.py — 测试")
    print("=" * 60)

    logger = AuditLogger()

    # 模拟一次完整合并流程的日志
    logger.log_file_scan(["C:/1.1综合楼.xlsx", "C:/1.2教学楼.xlsx"], source="指定文件夹")

    logger.log_sheet_detect(
        "1.1综合楼.xlsx", "表-08 综合楼",
        ["序号", "项目编码", "项目名称", "项目特征描述", "计量单位"],
        241, hit_keywords=["序号", "项目编码", "项目名称", "项目特征", "计量单位"],
    )
    logger.log_sheet_detect(
        "1.2教学楼.xlsx", "表-08 教学楼",
        ["序号", "项目编码", "项目名称", "项目特征", "计量单位"],
        193, hit_keywords=["序号", "项目编码", "项目名称", "项目特征", "计量单位"],
    )

    logger.log_header_compare({
        "1.1综合楼_综合楼": ["序号", "项目编码", "项目名称", "项目特征描述", "计量单位", "施工单位复核工程量"],
        "1.2教学楼_教学楼": ["序号", "项目编码", "项目名称", "项目特征", "计量单位"],
    })

    logger.log_merge_result(pd.DataFrame({"序号": [1, 2, 3]}), added_cols=["施工单位复核工程量"])

    logger.log_output("C:/清单合并.xlsx", ["清单合并", "操作记录"])
    logger.log_format_apply("清单合并")
    logger.log_format_apply("操作记录")

    # 输出
    df = logger.get_dataframe()
    print(f"\n操作记录预览（{len(df)} 条）：")
    print(df.to_string(index=False))

    print("\n[OK] audit_logger 测试通过")
