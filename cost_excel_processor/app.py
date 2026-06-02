"""
Streamlit Web界面
启动方式：streamlit run cost_excel_processor/app.py
"""

import sys
import os

# 将项目根目录加入 sys.path，确保绝对导入可用
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

import streamlit as st
import pandas as pd
import tempfile
from io import BytesIO

from cost_excel_processor.main_processor import CostExcelProcessor


st.set_page_config(
    page_title="造价清单自动化处理系统",
    page_icon="📊",
    layout="wide",
)

st.title("📊 造价Excel清单自动化处理系统")
st.caption("建筑施工企业 · Excel清单标准化与成本分析底座建设工具  v0.1.0")

# 侧边栏：使用说明
with st.sidebar:
    st.header("📖 使用说明")
    st.markdown("""
    **处理流程：**
    1. 上传非标准Excel清单文件
    2. 系统自动解析多级表头
    3. 校验并标准化字段
    4. 多工作表自动整合
    5. 企业规则自动分类匹配
    6. 下载标准化结果

    **必填标准表头：**
    - 业态 / 项目名称 / 项目特征
    - 计量单位 / 工程量
    - 综合单价 / 综合合价 / 备注

    **输出：**
    - 标准化清单（成本分析清单表1）
    - 自动新增分类匹配列（40+列）
    """)

# 主界面
uploaded_file = st.file_uploader(
    "第一步：上传Excel清单文件",
    type=["xlsx", "xls"],
    help="支持 .xlsx 和 .xls 格式，支持多工作表"
)

if uploaded_file is not None:
    temp_dir = tempfile.mkdtemp()
    temp_path = os.path.join(temp_dir, uploaded_file.name)
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success(f"✅ 文件已上传：{uploaded_file.name}")

    st.subheader("⚙️ 处理选项")
    col1, col2 = st.columns(2)
    with col1:
        interactive_mode = st.checkbox(
            "交互模式（缺失字段时提示手动映射）",
            value=False,
        )
    with col2:
        fill_building_type = st.checkbox(
            "自动提取业态（从工作表名）",
            value=True,
        )

    if st.button("🚀 开始处理", type="primary", use_container_width=True):
        with st.spinner("正在处理，请稍候..."):
            try:
                processor = CostExcelProcessor()
                processor.load(temp_path)
                structure = processor.analyze()

                st.subheader("📋 文件结构分析")
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("工作表数量", structure["sheet_count"])
                with c2:
                    st.metric("处理模式", "多表整合" if structure["is_multi_sheet"] else "单表处理")
                with c3:
                    total_rows = sum(
                        max(0, info["max_row"] - info.get("header_rows", 1))
                        for info in structure["sheets"].values()
                    )
                    st.metric("预估数据行", total_rows)

                with st.expander("查看各工作表详细信息"):
                    for name, info in structure["sheets"].items():
                        st.text(f"工作表：{name}")
                        st.text(f"  数据范围：{info['max_row']}行 × {info['max_col']}列")
                        st.text(f"  表头行数：{info['header_rows']}  合并单元格：{'是' if info['has_merge'] else '否'}")
                        st.text(f"  业态识别：{info['building_type']}")
                        st.divider()

                result_df = processor.process(interactive=interactive_mode)

                st.subheader("✅ 处理完成")
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("输出行数", len(result_df))
                with c2:
                    st.metric("输出列数", len(result_df.columns))
                with c3:
                    rule_cols = [c for c in result_df.columns if c not in [
                        "业态", "项目名称", "项目特征", "计量单位",
                        "工程量", "综合单价", "综合合价", "备注"
                    ]]
                    st.metric("分类匹配列", len(rule_cols))

                with st.expander("查看处理日志"):
                    for log in processor.get_logs():
                        st.text(log)

                st.subheader("👀 数据预览")
                st.dataframe(result_df.head(50), use_container_width=True, height=400)

                # 下载结果
                st.subheader("⬇️ 下载结果")
                output_buffer = BytesIO()
                with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                    result_df.to_excel(writer, sheet_name="成本分析清单表1", index=False)
                output_buffer.seek(0)

                st.download_button(
                    label="📥 下载标准化清单（Excel）",
                    data=output_buffer,
                    file_name=uploaded_file.name.replace(".xlsx", "").replace(".xls", "") + "_标准化清单.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                st.session_state["result_df"] = result_df
                st.session_state["processor"] = processor

            except Exception as e:
                st.error(f"❌ 处理失败：{str(e)}")
                st.exception(e)

# 结果显示
if "result_df" in st.session_state:
    st.divider()
    st.subheader("📊 快速统计（基于处理结果）")
    df = st.session_state["result_df"]
    if "工程量" in df.columns and "综合单价" in df.columns:
        df_num = df.copy()
        df_num["工程量"] = pd.to_numeric(df_num["工程量"], errors="coerce")
        df_num["综合单价"] = pd.to_numeric(df_num["综合单价"], errors="coerce")
        df_num["行合价"] = df_num["工程量"] * df_num["综合单价"]

        c1, c2, c3 = st.columns(3)
        with c1:
            total_qty = df_num["工程量"].sum()
            st.metric("总工程量", f"{total_qty:,.2f}" if pd.notna(total_qty) else "N/A")
        with c2:
            total_price = df_num["行合价"].sum()
            st.metric("总合价估算", f"{total_price:,.2f}" if pd.notna(total_price) else "N/A")
        with c3:
            avg_price = df_num["综合单价"].mean()
            st.metric("平均综合单价", f"{avg_price:,.2f}" if pd.notna(avg_price) else "N/A")

    rule_cols = [c for c in df.columns if c not in [
        "业态", "项目名称", "项目特征", "计量单位",
        "工程量", "综合单价", "综合合价", "备注"
    ]]
    if rule_cols:
        with st.expander("查看分类匹配统计"):
            shown = 0
            for col in rule_cols:
                non_empty = (df[col] != "") & (df[col].notna())
                cnt = non_empty.sum()
                if cnt > 0:
                    st.text(f"{col}：{cnt} 行匹配")
                    shown += 1
            if shown == 0:
                st.text("暂无匹配结果（可检查关键词配置）")
            if len(rule_cols) > shown:
                st.text(f"... 还有 {len(rule_cols) - shown} 个分类列未显示")
