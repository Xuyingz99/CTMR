import streamlit as st
import re

def apply_custom_css():
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

def display_pretty_report(title, report_text, bold_first_para=False):
    if not report_text: return

    # 从 session state 获取主题对应的卡片色板
    card_bg = st.session_state.get("card_bg", "rgb(247, 243, 223)")
    card_border = st.session_state.get("card_border", "#c4b89e")

    html_content = f"""
    <div style="background: {card_bg}; padding: 20px 25px; border-radius: 20px;
                border: 2.5px solid {card_border};
                box-shadow: 0 4px 10px rgba(107, 92, 67, 0.12); margin-top: 15px; margin-bottom: 25px;">
        <h4 style="color: #794f27; margin-top: 0; margin-bottom: 15px; font-weight: 800; font-size: 1.15rem;">{title}</h4>
    """

    lines = [line.strip() for line in report_text.split('\n') if line.strip()]

    is_first_para = True
    for line in lines:
        if line.endswith("：") or line.endswith(":"):
            html_content += f"<div style='font-weight: 700; margin-top: 15px; margin-bottom: 8px; color: #794f27; font-size: 1.05rem;'>{line}</div>"
            is_first_para = False
        elif line.startswith("•") or re.match(r'^\d+、', line):
            if "严重逾期" in line or "重点关注" in line or "逾期60天以上" in line:
                line = line.replace("严重逾期", "<span style='color: #e05a5a; font-weight: 600;'>严重逾期</span>")
                line = line.replace("重点关注", "<span style='color: #e59266; font-weight: 600;'>重点关注</span>")
                line = line.replace("逾期60天以上", "<span style='color: #e05a5a; font-weight: 600;'>逾期60天以上</span>")
            html_content += f"<div style='margin-left: 15px; margin-bottom: 6px; color: #725d42; line-height: 1.7;'>{line}</div>"
            is_first_para = False
        else:
            weight = "700" if (is_first_para and bold_first_para) else "500"
            color = "#794f27" if (is_first_para and bold_first_para) else "#725d42"
            html_content += f"<div style='margin-bottom: 10px; color: {color}; line-height: 1.7; font-weight: {weight};'>{line}</div>"
            is_first_para = False

    html_content += "</div>"

    st.markdown(html_content, unsafe_allow_html=True)
