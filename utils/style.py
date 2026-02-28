import streamlit as st
import re

def apply_custom_css():
    """注入设计师级 CSS (UI 优化版)"""
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

        /* 3. 问候语 */
        .greeting-text {
            font-size: 2rem;
            font-weight: 300;
            color: var(--text-main);
            text-align: center;
            margin-bottom: 2rem;
        }

        /* 4. 功能选择器 */
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

        /* 5. 说明框优化 (纯 HTML 左对齐) */
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

        /* 6. 上传与按钮美化 */
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
                
     /* 7. [新增] 大区筛选器 (Pills) 专项优化 */
        [data-testid="stPills"] {
            display: flex;
            gap: 12px; /* 按钮之间的间距 */
            flex-wrap: wrap;
            margin-bottom: 15px;
        }
        
        [data-testid="stPills"] button {
            border-radius: 20px !important; /* 圆角胶囊形状 */
            border: 1px solid #e0e0e0 !important;
            background: white !important;
            color: #5f6368 !important;
            padding: 6px 20px !important;
            font-size: 0.95rem !important;
            transition: all 0.2s ease;
            min-height: 40px !important; /* 强制高度，防止被压缩 */
            height: auto !important;
        }
        
        /* 选中状态：DeepSeek 蓝渐变 */
        [data-testid="stPills"] button[aria-selected="true"] {
            background: var(--btn-gradient) !important;
            color: white !important;
            border: none !important;
            box-shadow: 0 4px 12px rgba(77, 107, 254, 0.3);
            font-weight: 600 !important;
        }
        
        /* 悬停效果 */
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

def display_pretty_report(title, report_text, bg_color="#eef5ff"):
    """
    前端渲染优化：将报告文本拆分为“抬头”和“列表项”，美观展示
    """
    if not report_text: return
    
    # 尝试拆分
    parts = re.split(r'(分大区情况如下：|分经营部情况如下：|分客户情况如下：)', report_text)
    
    header_text = parts[0]
    detail_text = ""
    if len(parts) > 1:
        detail_text = "".join(parts[1:]) 
    
    st.markdown(f"""
    <div style="background-color: {bg_color}; padding: 15px; border-radius: 8px; border: 1px solid #d1e3ff; margin-bottom: 10px;">
        <h4 style="margin-top: 0; color: #1f1f1f;">{title}</h4>
        <div style="font-size: 1rem; color: #333; margin-bottom: 10px; line-height: 1.6;">{header_text}</div>
    </div>
    """, unsafe_allow_html=True)
    
    if detail_text:
        lines = [line.strip() for line in detail_text.split('\n') if line.strip()]
        
        list_html = ""
        for line in lines:
            if "情况如下：" in line:
                 list_html += f"<div style='font-weight: bold; margin-top: 8px; margin-bottom: 4px;'>{line}</div>"
            else:
                 list_html += f"<div style='margin-left: 10px; margin-bottom: 4px;'>• {line}</div>"
                 
        st.markdown(f"""
        <div style="background-color: #ffffff; padding: 15px; border-radius: 8px; border: 1px solid #eee;">
            {list_html}
        </div>
        """, unsafe_allow_html=True)
