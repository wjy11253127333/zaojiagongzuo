"""
------------------------------------------------
文件扫描与选择模块 v1.0
------------------------------------------------
功能：
  1. scan_folder(folder_path)       — 扫描文件夹下所有 .xlsx/.xls 文件
  2. select_files_dialog()          — 弹出 tk 文件选择对话框
  3. discover_list_files(mode, **kwargs) — 统一入口

依赖：tkinter（内置，无需额外安装）
"""

import os
import glob
import sys
import tkinter as tk
from tkinter import filedialog
from typing import List


# ============================================================
# 1. 文件夹扫描
# ============================================================

def scan_folder(folder_path: str, recursive: bool = False) -> List[str]:
    """
    扫描文件夹下所有 Excel 文件（.xlsx / .xls）
    
    Args:
        folder_path: 目标文件夹路径
        recursive:   是否递归扫描子文件夹（默认否）
    
    Returns:
        Excel 文件完整路径列表，按文件名排序
    """
    if not os.path.isdir(folder_path):
        raise NotADirectoryError(f"文件夹不存在：{folder_path}")

    patterns = ["*.xlsx", "*.xls"]
    files = []

    if recursive:
        for pattern in patterns:
            files.extend(glob.glob(os.path.join(folder_path, "**", pattern), recursive=True))
    else:
        for pattern in patterns:
            files.extend(glob.glob(os.path.join(folder_path, pattern)))

    # 去重 + 排序
    files = sorted(set(files))
    return files


# ============================================================
# 2. 文件选择对话框（tkinter）
# ============================================================

def select_files_dialog(title: str = "选择 Excel 文件（可多选）") -> List[str]:
    """
    弹出 Windows 原生文件选择对话框，用户可多选 Excel 文件
    
    Returns:
        用户选择的文件路径列表；若用户取消则返回空列表
    """
    # 隐藏 tk 主窗口
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)  # 置顶，避免被遮挡

    filetypes = [
        ("Excel 文件", "*.xlsx;*.xls"),
        ("所有文件", "*.*"),
    ]

    paths = filedialog.askopenfilenames(
        title=title,
        filetypes=filetypes,
    )

    root.destroy()
    return list(paths)


def select_folder_dialog(title: str = "选择文件夹") -> str:
    """
    弹出文件夹选择对话框
    
    Returns:
        用户选择的文件夹路径；若取消则返回空字符串
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    folder = filedialog.askdirectory(title=title)
    root.destroy()
    return folder or ""


# ============================================================
# 3. 统一入口
# ============================================================

def discover_list_files(mode: str = "auto", **kwargs) -> List[str]:
    """
    统一文件发现入口
    
    Args:
        mode: 文件来源模式
            "folder"  — 扫描指定文件夹（需传 folder_path=...）
            "dialog"  — 弹出文件选择对话框
            "auto"    — 先尝试 folder（传了 folder_path），若无文件则弹 dialog
        **kwargs:
            folder_path: 扫描目标文件夹（mode="folder" 或 "auto" 时使用）
            recursive:   是否递归扫描（默认 False）
            file_list:   直接传入文件列表（跳过扫描）
    
    Returns:
        Excel 文件路径列表
    
    Raises:
        FileNotFoundError: 未找到任何 Excel 文件且用户未选择文件
        NotADirectoryError: folder_path 不是有效文件夹（mode="folder" 时）
    """
    mode = mode.lower().strip()

    # ---- 模式：直接传入文件列表 ----
    if mode == "list" and "file_list" in kwargs:
        return [f for f in kwargs["file_list"] if os.path.isfile(f)]

    # ---- 模式：folder ----
    if mode in ("folder", "auto"):
        folder_path = kwargs.get("folder_path", "")
        recursive = kwargs.get("recursive", False)

        if folder_path and os.path.isdir(folder_path):
            files = scan_folder(folder_path, recursive=recursive)
            if files:
                print(f"[文件扫描] 文件夹「{folder_path}」找到 {len(files)} 个 Excel 文件")
                return files
            else:
                print(f"[WARN] 文件夹「{folder_path}」未找到 Excel 文件")
                if mode == "folder":
                    raise FileNotFoundError(f"文件夹下未找到 Excel 文件：{folder_path}")
                # mode == "auto" → 继续往下走 dialog

    # ---- 模式：dialog / auto 兜底 ----
    if mode in ("dialog", "auto"):
        print("[文件选择] 请手动选择 Excel 文件...")
        files = select_files_dialog()
        if files:
            print(f"[文件选择] 用户选择了 {len(files)} 个文件")
            return files
        else:
            raise FileNotFoundError("未选择任何文件，操作取消")

    # ---- 模式：folder 但 folder_path 无效 ----
    if mode == "folder":
        raise NotADirectoryError(f"无效的文件夹路径：{kwargs.get('folder_path', '')}")

    raise ValueError(f"不支持的 mode：{mode}，请使用 folder / dialog / auto / list")


# ============================================================
# CLI 测试入口
# ============================================================

if __name__ == "__main__":
    print("=" * 60)
    print("file_scanner.py — 测试")
    print("=" * 60)

    # 测试1：扫描当前目录
    cwd = os.getcwd()
    print(f"\n[测试1] 扫描当前目录：{cwd}")
    files = scan_folder(cwd)
    print(f"找到 {len(files)} 个 Excel 文件：")
    for f in files[:5]:
        print(f"  - {os.path.basename(f)}")
    if len(files) > 5:
        print(f"  ...（共 {len(files)} 个）")

    # 测试2：auto 模式（需要用户交互）
    print(f"\n[测试2] auto 模式（当前目录无文件时会弹对话框）...")
    try:
        result = discover_list_files("auto", folder_path=cwd)
        print(f"结果：{len(result)} 个文件")
    except FileNotFoundError as e:
        print(f"提示：{e}")
