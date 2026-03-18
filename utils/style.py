# utils/style.py
import streamlit as st
import re

def apply_custom_css():
    """
    注入设计师级 CSS (UI 优化版)
    如果主程序中已经注入过，这里可作为备用或供其他独立页面调用。
    """
    st.markdown("""
    <style>
        /* 1. 全局字体与配色 */
        html { font-size: 18px !important; }

        :root {
            /* DeepSeek 风格蓝色渐变 */
            --deepseek-blue: #4d6bfe;
            --deepseek-dark: #2b4cff;
            --btn-gradient: linear-gradient(90deg, #4d6bfe 0%, #2b4cff 100%);
            --bg-color: #f8f9fa;
            --text-main: #1f1f1f;
            --text-sub: #5f6368;
        }

        .stApp { background-color: var(--bg-color); }

        /* 2. 标题流光效果 */
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

        /* 3. 说明框优化 (纯 HTML 左对齐) */
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
    </style>
    """, unsafe_allow_html=True)

def display_pretty_report(title, report_text, bg_color="#f8f9fa"):
    """
    前端渲染优化：支持多行文本自然换行渲染，去除多余字眼，呈现 DeepSeek 蓝简约美观排版。
    兼容原有的信用报告和新增的逾期销售催收提醒文本。
    """
    if not report_text: 
        return
        
    # 直接呈现干净的标题和内容，无多余提醒框嵌套
    html_content = f"""
    <div style="background-color: {bg_color}; border-left: 4px solid var(--deepseek-blue); padding: 20px; border-radius: 8px; margin-top: 10px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.04);">
        <h4 style="color: var(--deepseek-dark); margin-top: 0; margin-bottom: 15px; font-weight: 600; font-size: 1.2rem;">{title}</h4>
    """
    
    lines = [line.strip() for line in report_text.split('\n') if line.strip()]
    
    for line in lines:
        if line.endswith("：") or line.endswith(":"):
            # 识别为大区/分类小标题（如：中粮贸易：、分大区看：）
            html_content += f"<div style='font-weight: 700; margin-top: 15px; margin-bottom: 8px; color: var(--text-main); font-size: 1.05rem;'>{line}</div>"
        elif line.startswith("•") or re.match(r'^\d+、', line):
            # 识别为列表项或特殊标签客户（如：1、华北区... 或 • 某某客户...）
            # 对于包含“严重逾期”或“重点关注”的文字给予轻度标红强调
            if "严重逾期" in line or "重点关注" in line or "逾期60天以上" in line:
                line = line.replace("严重逾期", "<span style='color: #d9534f; font-weight: 600;'>严重逾期</span>")
                line = line.replace("重点关注", "<span style='color: #f0ad4e; font-weight: 600;'>重点关注</span>")
                line = line.replace("逾期60天以上", "<span style='color: #c9302c; font-weight: 600;'>逾期60天以上</span>")
            html_content += f"<div style='margin-left: 15px; margin-bottom: 6px; color: var(--text-sub); line-height: 1.6;'>{line}</div>"
        else:
            # 普通段落总括句
            html_content += f"<div style='margin-bottom: 8px; color: var(--text-sub); line-height: 1.6;'>{line}</div>"
            
    html_content += "</div>"
    
    st.markdown(html_content, unsafe_allow_html=True)
