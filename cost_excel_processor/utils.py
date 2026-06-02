"""
工具函数模块
"""

import re
import sys
import os

# 确保可以导入同包模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def extract_building_type(sheet_name: str) -> str:
    """
    从工作表名中提取业态信息
    直接使用工作表全名作为业态值
    例如："土建-学校地上" → "土建-学校地上"
    """
    return sheet_name.strip()


def normalize_header_text(text) -> str:
    """标准化表头文本：去空格、换行、全角转半角"""
    if text is None:
        return ""
    s = str(text).strip()
    s = s.replace("\n", "").replace("\r", "").replace(" ", "").replace("　", "")
    return s


def find_header_match(headers: list, target: str) -> list:
    """在表头列表中模糊匹配目标表头，返回所有匹配到的索引列表"""
    matches = []
    target_norm = normalize_header_text(target)
    for i, h in enumerate(headers):
        h_norm = normalize_header_text(h)
        if h_norm and (target_norm in h_norm or h_norm in target_norm):
            matches.append(i)
    return matches


def is_numeric_column(col_name: str) -> bool:
    try:
        from config import NUMERIC_COLUMNS
    except ImportError:
        NUMERIC_COLUMNS = ["工程量", "综合单价", "综合合价"]
    norm = normalize_header_text(col_name)
    return any(normalize_header_text(c) in norm for c in NUMERIC_COLUMNS)


def format_cell_value(value, col_name: str = ""):
    """格式化单元格值"""
    if value is None:
        return ""
    return str(value).strip()
