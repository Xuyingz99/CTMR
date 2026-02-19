import streamlit as st
import pandas as pd
import io
import copy
import math
import warnings
import re
from datetime import datetime, timedelta
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# === 导入新增的风险日报逻辑模块 ===
from utils.logic_risk_report import process_risk_report

# 忽略警告
warnings.filterwarnings('ignore')

# --- 页面基础配置 ---
st.set_page_config(
    page_title="Take It Easy - 智能办公助手",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 此处保留你原有的 CSS 注入与 display_pretty_report 函数，不要删除 ---
# st.markdown("""<style> ... </style>""", unsafe_allow_html=True)
# def display_pretty_report(title, report_text, bg_color="#eef5ff"): ...

# --- 此处保留你原有的 Part 1 (初始保证金) & Part 2 (追加保证金) 所有的业务逻辑函数 ---
# def read_excel_safe(file_stream): ...
# def process_margin_deposit_logic(current_file, prev_file): ...
# def process_additional_margin_logic(uploaded_file, region_filter): ...

def main():
    st.markdown("""
        <div class="header-container">
            <h1 class="main-title">Take It Easy</h1>
            <div class="sub-title">Crafted by Xuyingzhe</div>
        </div>
    """, unsafe_allow_html=True)

    col_l, col_center, col_r = st.columns([1, 6, 1])

    with col_center:
        st.markdown('<div class="greeting-text">您好，有什么可以帮到你？</div>', unsafe_allow_html=True)

        function_map = {
            "📈 初始保证金处理": "init_margin",
            "📉 追加保证金处理": "add_margin",
            "📊 信用风险管理日报": "risk_report", # [新增项]
            "📝 格式转换 (Demo)": "demo"
        }

        mode = st.radio("选择功能", list(function_map.keys()), horizontal=True, label_visibility="collapsed")
        
        # --- 模块 1: 初始保证金处理 (原有) ---
        if mode == "📈 初始保证金处理":
            # (原逻辑保持不变...)
            pass
        
        # --- 模块 2: 追加保证金处理 (原有) ---
        elif mode == "📉 追加保证金处理":
            # (原逻辑保持不变...)
            pass
            
        # --- 模块 3: 信用风险管理日报 (新增) ---
        elif mode == "📊 信用风险管理日报":
            st.markdown("""
            <div class="info-box">
                <div class="info-title">⚠️ 注意事项</div>
                <div style="margin-left: 2px;">
                    <div>请上传包含「信用风险管理日报」及相应通报 Sheet 的 Excel 文件</div>
                    <div style="margin-top: 4px;">系统将自动抓取逾期数据生成 Word 简报，并导出相关 Sheet</div>
                    <div style="margin-top: 4px;">由于跨平台特性，云端部署时 PDF 导出将降级为高清图片输出</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            uploaded_file = st.file_uploader("📂 上传【信用风险管理日报】Excel 表", type=['xlsx'])
            
            if st.button("🚀 生成报告与导出文件 / Generate"):
                if uploaded_file:
                    with st.spinner("🤖 正在解析 Excel 数据并渲染跨平台文件，请稍候..."):
                        # 调用新模块逻辑
                        word_bytes, word_text, export_files, logs, env_msg = process_risk_report(uploaded_file)
                        
                        # 环境提示
                        st.info(f"💡 {env_msg}")
                        
                        if word_bytes or export_files:
                            st.success("✅ 任务处理完成！")
                            
                            # 展示运行日志
                            with st.expander("查看运行日志 / View Logs"):
                                for log in logs:
                                    st.write(log)
                            
                            # 渲染 Word 内容预览 (完美复用现有 CSS 组件)
                            if word_text:
                                display_pretty_report("信用风险管理日报 - 网页预览", word_text, "#fcf8f2")
                            
                            st.markdown("### 📥 下载生成文件")
                            # 下载布局，根据生成的文件数量动态创建列
                            dl_cols = st.columns(1 + len(export_files))
                            
                            with dl_cols[0]:
                                if word_bytes:
                                    original_base = os.path.splitext(uploaded_file.name)[0]
                                    st.download_button(
                                        label="📄 下载 Word 报告",
                                        data=word_bytes,
                                        file_name=f"{original_base}.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        use_container_width=True
                                    )
                                    
                            for i, export_file in enumerate(export_files, 1):
                                with dl_cols[i]:
                                    label = "📉 下载高清图" if export_file["type"] == "png" else "📊 下载 PDF"
                                    mime = "image/png" if export_file["type"] == "png" else "application/pdf"
                                    st.download_button(
                                        label=f"{label} ({export_file['name']})",
                                        data=export_file["data"],
                                        file_name=export_file["name"],
                                        mime=mime,
                                        use_container_width=True
                                    )
                                    
                            # 如果生成了降级的图片，在底部直接预览
                            png_files = [f for f in export_files if f["type"] == "png"]
                            if png_files:
                                st.markdown("#### 👁️ 图片预览")
                                for p_f in png_files:
                                    st.image(p_f["data"], caption=p_f["name"], use_container_width=True)

                        else:
                            st.error("处理失败，未能提取到有效数据。")
                            for log in logs:
                                st.write(log)
                else:
                    st.warning("⚠️ 请先上传 Excel 文件！")
                    
        else:
            st.info("此功能暂未开放，敬请期待...")

    st.markdown("<div style='text-align:center; color:#ccc; margin-top:50px;'>© 2026 TakeItEasy Tool</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
