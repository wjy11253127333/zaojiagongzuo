"""
主处理流程 v0.1.1
串联所有模块，完成完整的Excel清单自动化处理

更新日志（2026-05-20）：
- 添加清单表智能过滤：只处理表头含"项目特征"的工作表
- 添加"序号"列：在"业态"前自动插入，按行编号
- 移除企业规则匹配步骤（代码保留，日后配置后再启用）
"""

import os
import sys
import pandas as pd
import openpyxl

# 确保绝对导入可用
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_reader import ExcelReader
from header_parser import HeaderParser
from data_integrator import DataIntegrator
from rule_matcher import RuleMatcher
from utils import normalize_header_text
from config import OUTPUT_SHEET_NAME, REQUIRED_HEADERS


class CostExcelProcessor:
    """
    造价Excel清单自动化处理主程序
    使用流程：
        1. 初始化：processor = CostExcelProcessor()
        2. 加载文件：processor.load(file_path)
        3. 分析结构：processor.analyze()
        4. 处理：result_df = processor.process()
        5. 保存：processor.save(output_path)

    支持指定工作表：processor.set_sheet_filter(["土建-学校地下", "土建-学校地上"])
    支持启用规则匹配：processor.enable_rule_matching = True
    """

    def __init__(self):
        self.reader = None
        self.parser = None
        self.integrator = None
        self.matcher = None
        self.file_path = None
        self.wb = None
        self.structure = None
        self.sheet_raw_data = []   # 原始读取的各sheet数据
        self.parsed_sheets = []    # 解析后的各sheet数据（标准化表头）
        self.result_df = None       # 最终结果DataFrame
        self.logs = []             # 处理日志

        # 新增配置
        self.manual_sheet_names = None      # 手动指定工作表名列表（None=自动过滤）
        self.enable_rule_matching = False   # 是否启用企业规则匹配（默认关闭，日后配置）
        self.skipped_sheets = []            # 被跳过的非清单表

    def log(self, msg: str):
        """记录处理日志"""
        self.logs.append(msg)
        print(f"[INFO] {msg}")

    def load(self, file_path: str):
        """加载Excel文件"""
        self.file_path = file_path
        self.reader = ExcelReader(file_path)
        self.reader.load()
        self.wb = self.reader.wb
        self.parser = HeaderParser(self.wb)
        self.integrator = DataIntegrator()
        self.matcher = RuleMatcher()
        self.log(f"已加载文件：{os.path.basename(file_path)}")
        return self

    def set_sheet_filter(self, sheet_names: list):
        """
        手动指定要处理的工作表名列表
        使用后自动过滤模式失效，只处理指定的工作表
        """
        self.manual_sheet_names = sheet_names

    def analyze(self) -> dict:
        """分析文件结构（不修改数据，仅分析）"""
        if self.reader is None:
            raise ValueError("请先调用 load() 方法加载文件")
        self.structure = self.reader.recognize_structure()
        self.log(f"识别到 {self.structure['sheet_count']} 个工作表")
        for name, info in self.structure["sheets"].items():
            tag = ""
            if not info["is_empty"]:
                raw_ws = self.reader.read_sheet_raw(name)
                if not self._is_list_sheet(raw_ws):
                    tag = " [非清单表]"
            self.log(f"  工作表「{name}」：{info['max_row']}行×{info['max_col']}列，"
                     f"表头{info['header_rows']}行，{'有' if info['has_merge'] else '无'}合并单元格{tag}")
        return self.structure

    def _is_list_sheet(self, raw_sheet: dict) -> bool:
        """
        判断工作表是否为清单表
        规则：表头区域含"项目特征"关键词
        """
        if self.manual_sheet_names is not None:
            return raw_sheet["sheet_name"] in self.manual_sheet_names

        raw_headers = raw_sheet.get("raw_headers", [])
        all_text = ""
        for row in raw_headers:
            for cell in row:
                if cell:
                    all_text += normalize_header_text(str(cell))
        return "项目特征" in all_text

    def _should_skip_sheet(self, raw_sheet: dict) -> bool:
        """判断是否应跳过该工作表"""
        if self.manual_sheet_names is not None:
            return raw_sheet["sheet_name"] not in self.manual_sheet_names
        return not self._is_list_sheet(raw_sheet)

    def process(self, interactive: bool = False) -> pd.DataFrame:
        """
        执行完整处理流程
        interactive: 是否启用交互模式（缺失字段时提示用户手动映射）
        """
        if self.reader is None:
            raise ValueError("请先调用 load() 方法加载文件")

        self.log("=== 开始处理 ===")
        self.skipped_sheets = []

        # Step 1: 读取所有工作表原始数据
        self.log("Step 1: 读取所有工作表...")
        self.sheet_raw_data = self.reader.get_all_sheets_raw()

        # Step 2: 智能过滤 + 解析每个sheet的表头
        self.log("Step 2: 智能过滤清单表 + 解析表头（处理多级表头/合并单元格）...")
        if self.manual_sheet_names:
            self.log(f"  手动指定工作表：{self.manual_sheet_names}")
        else:
            self.log("  自动过滤模式：表头含'项目特征'关键词的视为清单表")
        self.parsed_sheets = []

        for raw in self.sheet_raw_data:
            sheet_name = raw["sheet_name"]

            # ===== 智能过滤 =====
            if self._should_skip_sheet(raw):
                self.skipped_sheets.append(sheet_name)
                self.log(f"  工作表「{sheet_name}」→ 跳过（非清单表）")
                continue

            building_type = raw["building_type"]
            header_rows = raw["header_rows"]
            first_header_row = raw.get("first_header_row", 1)

            # 用openpyxl直接解析表头（处理合并单元格）
            ws = self.wb[sheet_name]
            std_headers = self.parser.parse_sheet_headers(ws, header_rows, first_header_row)

            # 标准化表头名称
            std_headers = self.parser.standardize_headers(std_headers)

            # 校验必填字段
            validation = self.parser.validate_headers(std_headers)
            if not validation["is_valid"]:
                self.log(f"  工作表「{sheet_name}」缺失必填字段：{validation['missing']}")
                if interactive:
                    std_headers = self._auto_fix_headers(std_headers, validation, sheet_name)
                else:
                    std_headers = self._auto_fix_headers(std_headers, validation, sheet_name)

            # 确保业态列存在
            std_headers, yt_added = self.parser.auto_fill_building_type(std_headers, building_type)

            # 如果业态列是新增的，数据行也需要同步位移（在首位插入None）
            sheet_data = list(raw["data"])
            if yt_added:
                sheet_data = [[None] + list(row) for row in sheet_data]

            # 确保备注在最后
            std_headers = self.parser.ensure_remark_last(std_headers)

            self.parsed_sheets.append({
                "sheet_name": sheet_name,
                "headers": std_headers,
                "data": sheet_data,
                "building_type": building_type,
                "header_rows": header_rows,
                "yt_added": yt_added,
            })
            self.log(f"  工作表「{sheet_name}」→ 标准化表头：{std_headers[:8]}...")

        if not self.parsed_sheets:
            self.log("警告：没有找到可处理的清单表！请检查手动指定工作表名或文件内容。")
            self.result_df = pd.DataFrame()
            return self.result_df

        self.log(f"  过滤后保留 {len(self.parsed_sheets)} 个清单表，跳过 {len(self.skipped_sheets)} 个非清单表")

        # Step 3: 多表整合
        self.log("Step 3: 多表整合（纵向合并 + 横向拼接 + 去重）...")
        merged_df = self.integrator.integrate(self.parsed_sheets)
        self.log(f"  整合后共 {len(merged_df)} 行数据")

        # Step 4: 企业规则匹配（暂不启用，日后配置）
        if self.enable_rule_matching:
            self.log("Step 4: 应用企业规则匹配（新增分类列）...")
            result_df = self.matcher.apply_rules(merged_df)
            self.log(f"  新增分类列：{len(self.matcher.get_rule_columns())} 个")
        else:
            self.log("Step 4: 企业规则匹配 → 已跳过（暂不启用，日后配置）")
            result_df = merged_df

        # Step 5: 数据校验与清理
        self.log("Step 5: 数据校验与清理...")
        result_df = self._clean_data(result_df)

        # Step 6: 插入"序号"列（放在最前，业态之前）
        self.log("Step 6: 插入序号列...")
        result_df.insert(0, "序号", range(1, len(result_df) + 1))

        self.result_df = result_df
        self.log(f"=== 处理完成！ 共 {len(result_df)} 行 × {len(result_df.columns)} 列 ===")
        self.log(f"  包含 {len(self.parsed_sheets)} 个清单表，跳过 {len(self.skipped_sheets)} 个非清单表")
        return result_df

    def _auto_fix_headers(self, headers: list, validation: dict, sheet_name: str) -> list:
        """
        自动修复缺失的必填字段
        v0.1.1修复：不再插入列到表头（会导致数据列错位），
        缺失列由 data_integrator.py 自然处理（预置空列）
        此处仅记录日志
        """
        result = list(headers)
        missing = list(validation["missing"])

        # 尝试模糊匹配修复（仅重命名，不插入新列）
        remaining_missing = []
        for miss in missing:
            fixed = False
            for i, h in enumerate(result):
                h_norm = normalize_header_text(h)
                miss_norm = normalize_header_text(miss)
                if miss_norm in h_norm or h_norm in miss_norm:
                    result[i] = miss
                    fixed = True
                    break
            if not fixed:
                remaining_missing.append(miss)

        # 对无法匹配修复的字段，记录日志但不插入列
        # data_integrator.py 的 _align_sheet 方法会预置所有标准列
        for miss in remaining_missing:
            self.log(f"  [提示] 工作表「{sheet_name}」缺失字段「{miss}」，将保持为空")

        return result

    def _get_insert_position(self, header_name: str, current_headers: list) -> int:
        """根据标准表头顺序，确定新字段应插入的位置"""
        try:
            target_idx = REQUIRED_HEADERS.index(header_name)
        except ValueError:
            return len(current_headers)

        positions = []
        for i, h in enumerate(current_headers):
            if h in REQUIRED_HEADERS:
                positions.append((REQUIRED_HEADERS.index(h), i))

        positions.sort()
        for std_idx, col_idx in positions:
            if std_idx > target_idx:
                return col_idx
        return len(current_headers)

    def _clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """数据清洗"""
        result = df.copy()

        # 去除全空行
        result = result.dropna(how="all")

        # 去除"项目名称"为空的行
        if "项目名称" in result.columns:
            result = result[result["项目名称"].notna() & (result["项目名称"] != "")]

        # 数值列格式化
        for col in ["工程量", "综合单价", "综合合价"]:
            if col in result.columns:
                result[col] = pd.to_numeric(result[col], errors="coerce")

        return result.reset_index(drop=True)

    def save(self, output_path: str = None) -> str:
        """保存处理结果到Excel"""
        if self.result_df is None:
            raise ValueError("请先调用 process() 方法处理数据")

        if output_path is None:
            base = os.path.splitext(self.file_path)[0]
            output_path = base + "_标准化清单.xlsx"

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            self.result_df.to_excel(writer, sheet_name=OUTPUT_SHEET_NAME, index=False)

            # 自适应列宽
            ws = writer.sheets[OUTPUT_SHEET_NAME]
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

        self.log(f"结果已保存至：{output_path}")
        return output_path

    def get_logs(self) -> list:
        """获取处理日志"""
        return self.logs

    def get_summary(self) -> dict:
        """获取处理摘要"""
        if self.result_df is None:
            return {"status": "未处理"}
        return {
            "status": "已完成",
            "input_file": self.file_path,
            "output_rows": len(self.result_df),
            "output_columns": len(self.result_df.columns),
            "columns": list(self.result_df.columns),
            "processed_sheets": len(self.parsed_sheets) if self.parsed_sheets else 0,
            "skipped_sheets": len(self.skipped_sheets),
            "enable_rule_matching": self.enable_rule_matching,
        }
