"""
企业规则匹配模块
功能：在"备注"列后新增分类列，根据"项目特征"和"项目名称"内容进行关键词匹配
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
from config import RULE_KEY_WORDS
from utils import normalize_header_text


class RuleMatcher:
    """企业规则匹配器"""

    def __init__(self):
        self.rules = RULE_KEY_WORDS

    def apply_rules(self, df: pd.DataFrame) -> pd.DataFrame:
        """对DataFrame应用所有规则"""
        if df.empty:
            return df

        result = df.copy()

        # 找到"备注"列位置
        cols = list(result.columns)
        if "备注" in cols:
            remark_idx = cols.index("备注")
        else:
            remark_idx = len(cols) - 1

        # 按规则顺序插入新列
        new_cols = list(self.rules.keys())
        insert_idx = remark_idx + 1

        for col_name in new_cols:
            if col_name not in result.columns:
                result.insert(insert_idx, col_name, "")
                insert_idx += 1

        # 对每一行应用规则
        for idx, row in result.iterrows():
            self._match_row(row, result, idx)

        return result

    def _match_row(self, row, df: pd.DataFrame, idx: int):
        """对单行进行关键词匹配"""
        search_texts = []
        for col in ["项目特征", "项目名称"]:
            if col in row and pd.notna(row[col]):
                search_texts.append(str(row[col]))

        if not search_texts:
            return

        search_text = " ".join(search_texts)
        search_text_norm = normalize_header_text(search_text)

        for col_name, keywords in self.rules.items():
            if col_name not in df.columns:
                continue
            matched_kw = ""
            for kw in keywords:
                kw_norm = normalize_header_text(kw)
                if kw_norm in search_text_norm:
                    matched_kw = kw
                    break
            df.at[idx, col_name] = matched_kw

    def get_rule_columns(self) -> list:
        """返回所有规则列名"""
        return list(self.rules.keys())

    def add_custom_rule(self, col_name: str, keywords: list):
        """动态添加自定义规则"""
        self.rules[col_name] = keywords
