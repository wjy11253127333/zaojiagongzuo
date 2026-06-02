"""
多表整合模块
功能：按指定列纵向合并与横向拼接多表数据，处理表头对齐、缺失值填充及数据去重
整合为标准化单表，命名为"成本分析清单表1"
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
from config import REQUIRED_HEADERS, OUTPUT_SHEET_NAME
from utils import normalize_header_text


class DataIntegrator:
    """多表数据整合器"""

    def __init__(self):
        self.standard_headers = list(REQUIRED_HEADERS)

    def integrate(self, sheet_data_list: list) -> pd.DataFrame:
        """整合多表数据"""
        if not sheet_data_list:
            return pd.DataFrame(columns=self.standard_headers)

        aligned_dfs = []
        for sheet_info in sheet_data_list:
            df = self._align_sheet(sheet_info)
            aligned_dfs.append(df)

        if not aligned_dfs:
            return pd.DataFrame(columns=self.standard_headers)

        merged = pd.concat(aligned_dfs, ignore_index=True, sort=False)
        merged = self._reorder_columns(merged)
        merged = self._ensure_remark_last(merged)

        # 填充业态列
        if "业态" in merged.columns:
            merged["业态"] = merged["业态"].ffill().fillna("")

        # 去重
        merged = self._deduplicate(merged)
        merged = merged.reset_index(drop=True)

        return merged

    def _align_sheet(self, sheet_info: dict) -> pd.DataFrame:
        """将单个sheet的数据对齐到标准表头"""
        headers = sheet_info.get("headers", [])
        data = sheet_info.get("data", [])
        building_type = sheet_info.get("building_type", "")

        if not headers or not data:
            return pd.DataFrame(columns=self.standard_headers)

        # 建立原始表头位置 → 标准表头名 的映射
        header_map = {}
        norm_standard = [normalize_header_text(h) for h in self.standard_headers]

        for idx, h in enumerate(headers):
            if not h:
                continue
            h_norm = normalize_header_text(h)
            for i, ns in enumerate(norm_standard):
                if ns == h_norm or ns in h_norm or h_norm in ns:
                    header_map[idx] = self.standard_headers[i]
                    break
            else:
                header_map[idx] = h

        # 构建DataFrame
        rows = []
        for row in data:
            new_row = {h: "" for h in self.standard_headers}
            for col_idx, val in enumerate(row):
                if col_idx in header_map:
                    target_col = header_map[col_idx]
                    # 业态列不从原始数据映射，统一用工作表名填充
                    if target_col == "业态":
                        continue
                    if target_col in self.standard_headers:
                        new_row[target_col] = val if val is not None else ""
            # 业态列始终填充为工作表名（清单名称）
            new_row["业态"] = building_type
            rows.append(new_row)

        df = pd.DataFrame(rows, columns=self.standard_headers)
        return df

    def _reorder_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """重新排列列顺序"""
        ordered = []
        for h in self.standard_headers:
            if h in df.columns:
                ordered.append(h)
        for h in df.columns:
            if h not in ordered:
                ordered.append(h)
        return df[ordered]

    def _deduplicate(self, df: pd.DataFrame) -> pd.DataFrame:
        """去重"""
        if df.empty:
            return df

        subset = []
        if "项目名称" in df.columns:
            subset.append("项目名称")
        if "工程量" in df.columns:
            subset.append("工程量")
        if "计量单位" in df.columns:
            subset.append("计量单位")

        if subset:
            df = df.drop_duplicates(subset=subset, keep="first")
        return df

    def _ensure_remark_last(self, df: pd.DataFrame) -> pd.DataFrame:
        """确保"备注"列在最后"""
        cols = list(df.columns)
        if "备注" in cols and cols[-1] != "备注":
            cols.remove("备注")
            cols.append("备注")
            df = df[cols]
        elif "备注" not in cols:
            df["备注"] = ""
            cols = list(df.columns)
            cols.remove("备注")
            cols.append("备注")
            df = df[cols]
        return df
