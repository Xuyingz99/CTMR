import streamlit as st
import pandas as pd
import io
import copy
import math
import warnings
import re
import os
from datetime import datetime, timedelta
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# === 导入新增的信用风险管理模块 ===
from utils.logic_credit import process_credit_report
# 修改点：此处确保导入了 logic_add 中的处理逻辑（根据你的文件结构，如果是在同一文件或外部 utils 中请确认路径）

# 忽略警告
warnings.filterwarnings('ignore')

# --- 页面基础配置 ---
st.set_page_config(
    page_title="Take It Easy - 智能办公助手",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 注入设计师级 CSS (UI 优化版) ---
st.markdown("""
<style>
    /* 1. 全局字体与配色 */
    html { font-size: 18px !important; }

    :root {
        --deepseek-blue: #4d6bfe;
        --deepseek-dark: #2b4cff;
        --btn-gradient: linear-gradient(90deg, #4d6bfe 0%, #2b4cff 100%);
        --bg-color: #f8f9fa;
        --text-main: #1f1f1f;
        --text-sub: #5f6368;
    }

    .stApp { background-color: var(--bg-color); }

    .header-container {
        text-align: center;
        padding: 3rem 0 1rem 0;
    }
    .main-title {
        font-size: 4.5rem !important;
        font-weight: 800;
        letter-spacing: -2px;
        margin: 0;
        background: linear-gradient(90deg, #4285f4, #9b72cb, #d96570);
        background-size: 200% auto;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        animation: shine 5s linear infinite;
    }
    @keyframes shine { to { background-position: 200% center; } }
    
    .sub-title {
        font-size: 1rem;
        color: var(--text-sub);
        letter-spacing: 2px;
        text-transform: uppercase;
        margin-top: 0.5rem;
    }

    .greeting-text {
        font-size: 2rem;
        font-weight: 300;
        color: var(--text-main);
        text-align: center;
        margin-bottom: 2rem;
    }

    div[role="radiogroup"] > label > div:first-child { display: none !important; }
    div[role="radiogroup"] {
        display: flex;
        justify-content: center;
        gap: 15px;
        width: 100%;
        margin-bottom: 25px;
    }
    div[role="radiogroup"] label {
        background: white;
        border: 1px solid #e0e0e0;
        border-radius: 12px;
        padding: 15px;
        text-align: center;
        box-shadow: 0 4px 10px rgba(0,0,0,0.05);
        cursor: pointer;
        flex: 1;
        transition: all 0.3s;
        min-height: 80px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 600;
        color: var(--text-sub);
    }
    div[role="radiogroup"] label[data-checked="true"] {
        border: 2px solid transparent !important;
        background: linear-gradient(white, white) padding-box, var(--btn-gradient) border-box !important;
        color: var(--deepseek-blue) !important;
        transform: translateY(-4px);
        box-shadow: 0 8px 20px rgba(77, 107, 254, 0.2);
    }

    .info-box {
        background: #ffffff;
        border-left: 4px solid var(--deepseek-blue);
        padding: 20px 25px;
        border-radius: 0 8px 8px 0;
        margin-bottom: 25px;
        color: #4a4a4a;
        font-size: 1rem;
        box-shadow: 0 2px 10px rgba(0,0,0,0.03);
        text-align: left;
        line-height: 1.8;
    }
    .info-title {
        font-weight: 700;
        color: #1f1f1f;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    [data-testid="stFileUploader"] section {
        border-radius: 12px;
        background-color: white;
        border: 2px dashed #dbe0ea;
        padding: 1.5rem;
    }
    [data-testid="stFileUploader"] section:hover { border-color: var(--deepseek-blue); }
    
    div.stButton > button {
        width: 100%;
        height: 60px;
        border-radius: 12px;
        font-size: 1.2rem;
        font-weight: 600;
        background: var(--btn-gradient);
        color: white;
        border: none;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(77, 107, 254, 0.3);
    }
    div.stButton > button:hover {
        transform: scale(1.02);
        box-shadow: 0 8px 25px rgba(77, 107, 254, 0.4);
        color: white;
    }

    #MainMenu, header, footer { visibility: hidden; }
            
    [data-testid="stPills"] {
        display: flex;
        gap: 12px;
        flex-wrap: wrap;
        margin-bottom: 15px;
    }
    
    [data-testid="stPills"] button {
        border-radius: 20px !important;
        border: 1px solid #e0e0e0 !important;
        background: white !important;
        color: #5f6368 !important;
        padding: 6px 20px !important;
        font-size: 0.95rem !important;
        transition: all 0.2s ease;
        min-height: 40px !important;
        height: auto !important;
    }
    
    [data-testid="stPills"] button[aria-selected="true"] {
        background: var(--btn-gradient) !important;
        color: white !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(77, 107, 254, 0.3);
        font-weight: 600 !important;
    }
    
    [data-testid="stPills"] button:hover {
        border-color: var(--deepseek-blue) !important;
        color: var(--deepseek-blue) !important;
        transform: translateY(-1px);
    }
    [data-testid="stPills"] button[aria-selected="true"]:hover {
        color: white !important;
        transform: translateY(-1px);
    }           
</style>
""", unsafe_allow_html=True)

# ... (此处省略 process_margin_deposit_logic 等保证金处理函数，保持原样) ...

# ==========================================
# 网页美化渲染函数 (全局通用) - 重点修改此处
# ==========================================

def display_pretty_report(title, report_text, bg_color="#eef5ff"):
    if not report_text: return
    
    # 逻辑修改：将拆分关键词扩展到“分大区”、“分经营部”以及通用的换行符处理
    parts = re.split(r'(分大区情况如下：|分经营部情况如下：)', report_text)
    header_text = parts[0]
    detail_text = ""
    if len(parts) > 1:
        detail_text = "".join(parts[1:])
    else:
        # 如果没有关键词（比如是分客户报告），则将头部和明细按第一个双换行拆分
        if "\n\n" in header_text:
            sub_parts = header_text.split("\n\n", 1)
            header_text = sub_parts[0]
            detail_text = sub_parts[1]
    
    st.markdown(f"""
    <div style="background-color: {bg_color}; padding: 15px; border-radius: 8px; border: 1px solid #d1e3ff; margin-bottom: 10px;">
        <h4 style="margin-top: 0; color: #1f1f1f;">{title}</h4>
        <div style="font-size: 1rem; color: #333; margin-bottom: 10px; line-height: 1.6;">{header_text}</div>
    </div>
    """, unsafe_allow_html=True)
    
    if detail_text:
        # 修改点：先按换行符拆分，然后过滤掉空行
        lines = [line.strip() for line in detail_text.split('\n') if line.strip()]
        list_html = ""
        for line in lines:
            # 识别带编号的行（如 1、2、...）或特定标题行
            if "情况如下：" in line:
                 list_html += f"<div style='font-weight: bold; margin-top: 8px; margin-bottom: 4px;'>{line}</div>"
            elif re.match(r'^\d+、', line):
                 # 修改点：为每一个带数字编号的客户明细增加独立的 margin-bottom 和段落间距
                 list_html += f"<div style='margin-left: 10px; margin-bottom: 12px; border-bottom: 1px solid #f0f0f0; padding-bottom: 8px;'>• {line}</div>"
            else:
                 list_html += f"<div style='margin-left: 10px; margin-bottom: 4px;'>• {line}</div>"
                 
        st.markdown(f"""
        <div style="background-color: #ffffff; padding: 15px; border-radius: 8px; border: 1px solid #eee;">
            {list_html}
        </div>
        """, unsafe_allow_html=True)

# ... (此处省略 main() 函数及其内部逻辑，保持原样即可) ...
