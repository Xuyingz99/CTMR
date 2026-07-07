import streamlit as st
import pandas as pd
import warnings
import os
from datetime import datetime

from utils.logic_credit import process_credit_report
from utils.logic_XS import process_overdue_sales
from utils.style import display_pretty_report
from utils.logic_CG import process_overdue_purchase
from utils.logic_init import process_margin_deposit_logic
from utils.logic_add import process_additional_margin_logic

warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="Take It Easy - 智能办公助手",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# 🎨 动态主题配置引擎
# ==========================================
THEMES = {
    "🏝️ 狸克海岛": {
        "--ac-bg": "#F9F6ED",         "--ac-text": "#604F43",
        "--ac-green": "#59C19A",      "--ac-green-dark": "#439E7C",
        "--ac-yellow": "#F7D273",     "--ac-yellow-dark": "#D4A93E",
        "--ac-wood": "#D2BA99",       "--ac-card": "#FFFFFF",
        "--ac-info-bg": "#FFF8E6",
        "--ac-card-bg": "rgb(247, 243, 223)",  "--ac-card-border": "#c4b89e",
        "--ac-bg-img": "linear-gradient(rgba(249, 246, 237, 0.92), rgba(249, 246, 237, 0.92)), url('https://cdn.jsdelivr.net/gh/guokaigdg/animal-island-ui@main/demo/img/menu_bg.svg')",
        "--ac-bg-size": "220px auto",
        "--ac-bg-position": "top left"
    },
    "🌊 冲浪海滩": {
        "--ac-bg": "#F2F9FF",         "--ac-text": "#2C4C5E",
        "--ac-green": "#6AC4D9",      "--ac-green-dark": "#4A9EAF",
        "--ac-yellow": "#E6F4F1",     "--ac-yellow-dark": "#8BD3E6",
        "--ac-wood": "#B5C8D6",       "--ac-card": "#FFFFFF",
        "--ac-info-bg": "#EAF4FC",
        "--ac-card-bg": "#F5FAFE",    "--ac-card-border": "#B5C8D6",
        "--ac-bg-img": "radial-gradient(rgba(106, 196, 217, 0.2) 2px, transparent 2px), radial-gradient(rgba(106, 196, 217, 0.1) 2px, transparent 2px)",
        "--ac-bg-size": "30px 30px",
        "--ac-bg-position": "0 0, 15px 15px"
    }
}

# 初始化 Session State 中的当前主题（首次加载随机二选一，后续 rerun 保持稳定）
if "current_theme" not in st.session_state:
    themes = list(THEMES.keys())
    st.session_state.current_theme = themes[int(datetime.now().timestamp() * 1000) % len(themes)]

# 提取当前主题的 CSS 变量
current_theme_vars = THEMES[st.session_state.current_theme]
css_vars_string = "\n".join([f"        {k}: {v};" for k, v in current_theme_vars.items()])

# 动态注入 CSS
st.markdown(f"""
<style>
    /* 引入圆润可爱的字体 */
    @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;700;900&display=swap');

    html {{ font-size: 18px !important; }}

    /* ✨ 动态注入的主题色板 */
    :root {{
{css_vars_string}
    }}

    /* 全局背景与字体 (动态调用主题底纹与色盘) */
    .stApp {{
        background-color: var(--ac-bg) !important;
        background-image: var(--ac-bg-img) !important;
        background-size: var(--ac-bg-size) !important;
        background-position: var(--ac-bg-position) !important;
        background-attachment: fixed !important;
        font-family: 'Nunito', 'PingFang SC', 'Microsoft YaHei', sans-serif !important;
        color: var(--ac-text) !important;
        cursor: url('https://cdn.jsdelivr.net/gh/guokaigdg/animal-island-ui@main/src/assets/img/cursor/cursor-icon.png'), auto !important;
    }}

    /* 关键修复：强制让 Streamlit 的顶层容器与内部视图透明，确保底层底纹完全暴露 */
    [data-testid="stAppViewContainer"],
    [data-testid="stHeader"] {{
        background: transparent !important;
    }}

    /* 缩紧 Streamlit 默认的全局外围边距，让 UI 更紧凑 */
    [data-testid="block-container"] {{
        padding-top: 1.2rem !important;
        padding-bottom: 2rem !important;
        padding-left: 3rem !important;
        padding-right: 3rem !important;
        max-width: 96% !important;
    }}

    /* 强制所有交互元素及其内部文字使用动森光标 */
    button, button *, div[role="radiogroup"] label, div[role="radiogroup"] label *, a, a *, input, [data-testid="stFileUploader"] section, [data-testid="stFileUploader"] section * {{
        cursor: url('https://cdn.jsdelivr.net/gh/guokaigdg/animal-island-ui@main/src/assets/img/cursor/cursor-icon.png'), auto !important;
    }}

    /* 覆盖 Streamlit 默认标题颜色 */
    h1, h2, h3, h4, h5, h6, p, span {{
        color: var(--ac-text) !important;
    }}

    /* 1. 网页标题美化 */
    .header-container {{ text-align: center; padding: 2rem 0 1rem 0; }}
    .main-title {{
        font-size: 4rem !important; font-weight: 900; color: var(--ac-green);
        text-shadow: 3px 3px 0px #FFFFFF, 6px 6px 0px var(--ac-wood); letter-spacing: 2px; margin: 0;
    }}
    .sub-title {{ font-size: 1.1rem; color: var(--ac-wood) !important; font-weight: 700; letter-spacing: 2px; margin-top: 1rem; }}
    .greeting-text {{ font-size: 1.6rem; font-weight: 700; color: var(--ac-text); text-align: center; margin-bottom: 2rem; }}

    /* 2. 顶部功能切换卡片 (无损增量升级：彻底隐藏原生红点) */
    [data-testid="stRadio"] div[role="radiogroup"] > label > div:first-child {{
        display: none !important; /* 核心：彻底抹除左侧刺眼大红点 */
    }}

    [data-testid="stRadio"] div[role="radiogroup"] {{
        display: flex !important; flex-wrap: wrap !important; justify-content: flex-start !important;
        gap: 16px; width: 100% !important; max-width: 964px !important; margin: 0 auto 30px auto !important;
    }}

    /* 未选中时的卡片基础态 */
    [data-testid="stRadio"] div[role="radiogroup"] label {{
        background: var(--ac-card) !important;
        border: 2px solid #E5E7EB !important;
        border-radius: 14px !important;
        padding: 10px 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.03) !important;
        flex: 0 0 180px !important; width: 180px !important; height: 85px !important; box-sizing: border-box !important;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
        display: flex !important; align-items: center !important; justify-content: center !important; text-align: center !important;
        cursor: url('https://cdn.jsdelivr.net/gh/guokaigdg/animal-island-ui@main/src/assets/img/cursor/cursor-icon.png'), pointer !important;
    }}

    /* 未选中时的卡片文字 */
    [data-testid="stRadio"] div[role="radiogroup"] label p {{
        color: var(--ac-text) !important; font-weight: 600 !important; font-size: 0.95rem !important; line-height: 1.3 !important; margin: 0 !important;
    }}

    /* 选中时的卡片高亮态 (动态无缝兼容：海岛自动变主题黄，海滩自动变海盐蓝) */
    [data-testid="stRadio"] div[role="radiogroup"] > label[data-checked="true"] {{
        background: var(--ac-yellow) !important;
        border-color: var(--ac-yellow-dark) !important;
        box-shadow: 0 5px 0 0 var(--ac-wood), 0 4px 12px rgba(107, 92, 67, 0.15) !important;
        transform: translateY(-3px) !important;
    }}

    /* 选中卡片内部的文字加粗穿透 */
    [data-testid="stRadio"] div[role="radiogroup"] > label[data-checked="true"] p {{
        font-weight: 800 !important;
        color: #794f27 !important; /* 经典的动森深木色高亮文字 */
    }}

    /* 未选中卡片的 Hover 悬浮律动 */
    [data-testid="stRadio"] div[role="radiogroup"] label:hover:not([data-checked="true"]) {{
        transform: translateY(-2px) !important;
        border-color: var(--ac-wood) !important;
        box-shadow: 0 6px 12px rgba(107, 92, 67, 0.08) !important;
    }}

    /* 2.5 动森风格复选框 (专门针对逾期销售周报的 checkbox) */
    [data-testid="stCheckbox"] label > div:first-child {{
        background: var(--ac-card-bg) !important;
        border: 2px solid var(--ac-card-border) !important;
        border-radius: 8px !important;
        width: 22px !important; height: 22px !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }}
    /* 选中状态变成主题绿 */
    [data-testid="stCheckbox"] label[data-checked="true"] > div:first-child {{
        background: var(--ac-green) !important;
        border-color: var(--ac-green-dark) !important;
    }}
    /* 确保里面的原生对勾为白色 */
    [data-testid="stCheckbox"] label > div:first-child svg {{
        fill: white !important;
        stroke: white !important;
    }}

    /* 3. 注意事项/信息框 (有机形态与 NookPhone 色卡) */
    .info-box {{
        background: var(--ac-info-bg) !important;
        border: 2.5px solid var(--ac-wood) !important; border-left: none !important;
        padding: 20px 25px;
        border-radius: 40px 35px 45px 38px / 38px 45px 35px 40px !important;
        margin-bottom: 25px; color: var(--ac-text) !important;
        font-size: 1rem; font-weight: 500;
        box-shadow: 0 4px 10px rgba(107, 92, 67, 0.08) !important; line-height: 1.8;
    }}
    .info-title {{ font-weight: 800; color: var(--ac-green); font-size: 1.15rem; margin-bottom: 12px; display: flex; align-items: center; gap: 8px; }}

    /* 4. NookPhone 圆润文件上传区 */
    [data-testid="stFileUploader"] section {{
        border-radius: 20px !important; background-color: var(--ac-card) !important;
        border: 3px dashed var(--ac-green) !important;
        padding: 1rem 0.5rem !important; /* 收紧内边距，给双列留出空间 */
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }}
    [data-testid="stFileUploader"] section:hover {{
        background-color: var(--ac-info-bg) !important;
        border-color: var(--ac-green-dark) !important;
        transform: translateY(-4px) !important;
        box-shadow: 0 8px 24px rgba(114, 93, 66, 0.15) !important;
    }}

    /* 4.1 强制穿透修改上传区内部文字与图标大小 (防止溢出) */
    [data-testid="stFileUploader"] section * {{
        font-size: 0.85rem !important;
    }}
    [data-testid="stFileUploader"] section small {{
        font-size: 0.7rem !important;
    }}
    [data-testid="stFileUploader"] section svg {{
        width: 30px !important; height: 30px !important; margin-bottom: 5px !important;
    }}

    /* 4.5 NookPhone 胶囊化文本输入框 */
    [data-testid="stTextInput"] [data-baseweb="input"] {{
        border-radius: 50px !important;
        border: 2.5px solid var(--ac-wood) !important;
        background-color: var(--ac-card) !important;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
        overflow: hidden;
    }}
    [data-testid="stTextInput"] [data-baseweb="input"]:focus-within {{
        border-color: var(--ac-green) !important;
        box-shadow: 0 3px 0 0 var(--ac-green-dark), 0 0 0 3px rgba(89, 193, 154, 0.15) !important;
    }}
    [data-testid="stTextInput"] input {{ color: var(--ac-text) !important; font-weight: 600 !important; }}

    /* 5. 核心操作与下载按钮 */
    div.stButton > button, div[data-testid="stDownloadButton"] > button {{
        width: 100%; height: 55px; border-radius: 12px; font-size: 1.15rem; font-weight: 700;
        background-color: var(--ac-green); color: white !important; border: none;
        box-shadow: 0 6px 0 var(--ac-green-dark); transition: all 0.15s cubic-bezier(0.4, 0, 0.2, 1);
    }}
    div.stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover {{ background-color: var(--ac-green); transform: translateY(2px); box-shadow: 0 4px 0 var(--ac-green-dark); opacity: 0.9; }}
    div.stButton > button:active, div[data-testid="stDownloadButton"] > button:active {{ transform: translateY(6px); box-shadow: none; }}

    /* 6. Pills / SegmentedControl 标签 */
    [data-testid="stPills"],
    [data-testid="stSegmentedControl"] {{ display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 20px; }}
    [data-testid="stPills"] button {{
        border-radius: 12px !important; border: 1px solid var(--ac-wood) !important; background: var(--ac-card) !important;
        color: var(--ac-text) !important; padding: 6px 20px !important; font-size: 1rem !important; font-weight: 700 !important; transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
    }}
    [data-testid="stPills"] button[aria-selected="true"] {{ background: var(--ac-green) !important; color: white !important; border-color: var(--ac-green-dark) !important; box-shadow: 0 4px 0 var(--ac-green-dark); transform: translateY(-2px); }}
    /* SegmentedControl 卡片风格 (模块选择器) */
    [data-testid="stSegmentedControl"] button {{
        border-radius: 12px !important; border: 1px solid #E5E7EB !important; background: var(--ac-card) !important;
        color: var(--ac-text) !important; padding: 10px 12px !important; font-size: 0.95rem !important; font-weight: 700 !important;
        width: 180px !important; height: 85px !important; box-shadow: 0 4px 6px rgba(0,0,0,0.04) !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1); white-space: normal !important;
    }}
    [data-testid="stSegmentedControl"] button[aria-selected="true"] {{
        background: var(--ac-yellow) !important; border-color: var(--ac-yellow-dark) !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08) !important; transform: translateY(-2px);
    }}
    [data-testid="stSegmentedControl"] button:hover:not([aria-selected="true"]) {{ transform: translateY(-2px); box-shadow: 0 6px 12px rgba(0,0,0,0.08) !important; }}

    /* ========================================================= */
    /* 🍃 海岛专属加载动画 */
    /* ========================================================= */
    [data-testid="stSpinner"] > div > svg {{ display: none !important; }}
    [data-testid="stSpinner"] > div {{ display: flex; align-items: center; }}
    [data-testid="stSpinner"] > div::before {{
        content: '';
        display: inline-block; width: 32px; height: 32px;
        background-image: url('https://cdn.jsdelivr.net/gh/guokaigdg/animal-island-ui@main/src/assets/img/icons/icon-leaf.png');
        background-size: contain; background-repeat: no-repeat;
        animation: ac-spin 1.5s cubic-bezier(0.4, 0, 0.2, 1) infinite;
        margin-right: 15px;
    }}
    @keyframes ac-spin {{
        0% {{ transform: rotate(0deg) scale(1); }}
        50% {{ transform: rotate(180deg) scale(1.15); }}
        100% {{ transform: rotate(360deg) scale(1); }}
    }}
    [data-testid="stSpinner"] p {{
        color: var(--ac-green) !important; font-weight: 900 !important; font-size: 1.25rem !important; letter-spacing: 1.5px !important;
    }}

    #MainMenu, header, footer {{ visibility: hidden; }}

    /* 7. 动森风格原生弹窗 (Modal) 美化 */
    [data-testid="stModal"] > div {{
        background: var(--ac-card) !important;
        border: 3px solid var(--ac-green) !important;
        border-radius: 40px 35px 45px 38px / 38px 45px 35px 40px !important; /* 动森有机圆角 */
        padding: 10px !important;
        box-shadow: 0 8px 24px rgba(107, 92, 67, 0.2) !important;
    }}
    /* 修改弹窗标题栏颜色和字体 */
    [data-testid="stModal"] header {{
        background: transparent !important;
    }}
    [data-testid="stModal"] h2 {{
        color: var(--ac-green) !important;
        font-weight: 800 !important;
        font-size: 1.5rem !important;
        text-shadow: 2px 2px 0px #fff !important;
    }}
    /* 弹窗内的关闭按钮 */
    [data-testid="stModal"] header button {{
        color: var(--ac-wood) !important;
    }}
    [data-testid="stModal"] header button:hover {{
        background-color: var(--ac-yellow) !important;
        color: white !important;
    }}

    /* 8. 修复：精准锁定设置面板内的按钮 (Small 尺寸 + Ghost 幽灵形态) */
    [data-testid="stExpanderDetails"] button {{
        height: 36px !important; /* 稍微加到36px，保证中文和Emoji不被裁切 */
        padding: 0 16px !important;
        border-radius: 12px !important;
        background: transparent !important; /* Ghost 核心：透明底色 */
        border: 2px solid var(--ac-text) !important;
        box-shadow: 0 3px 0 0 var(--ac-wood) !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
        margin-top: 5px !important;
    }}

    /* 强力覆盖内部文字的样式，防止被 Streamlit 原生覆盖 */
    [data-testid="stExpanderDetails"] button p {{
        font-size: 13.5px !important;
        font-weight: 700 !important;
        color: var(--ac-text) !important;
        margin: 0 !important;
    }}

    /* Hover 状态：加深边框与文字，铺一层极浅的主题色背景 */
    [data-testid="stExpanderDetails"] button:hover {{
        background: var(--ac-info-bg) !important;
        border-color: var(--ac-text) !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 0 0 var(--ac-wood) !important;
    }}

    [data-testid="stExpanderDetails"] button:hover p {{
        color: var(--ac-text) !important;
    }}

    /* Active 点击状态：真实的按压反馈 */
    [data-testid="stExpanderDetails"] button:active {{
        transform: translateY(2px) !important;
        box-shadow: 0 1px 0 0 var(--ac-wood) !important;
    }}

    /* 9. 修复所有原生组件标题 (如 "选择报告生成范围" 等 Widget Label) 的颜色断层 */
    [data-testid="stWidgetLabel"] p {{
        color: var(--ac-text) !important;
        font-weight: 800 !important;
        font-size: 1.1rem !important;
        letter-spacing: 0.02em !important;
        margin-bottom: 8px !important;
        text-shadow: 1px 1px 0px var(--ac-bg) !important;
    }}

</style>
""", unsafe_allow_html=True)

# (方案 B 已启用，CSS 重塑红点为打钩方块，JS 隐藏逻辑暂注释)
# st.html("""
# <script>
# (function(){
#     function hideDots() {
#         document.querySelectorAll('div[role="radiogroup"] label').forEach(function(l){
#             var fd = l.querySelector(':scope > div:first-child');
#             if (fd && fd.offsetWidth < 30) { fd.style.display = 'none'; }
#         });
#     }
#     hideDots();
#     new MutationObserver(hideDots).observe(document.body, {childList:true, subtree:true});
# })();
# </script>
# """)

def format_html_content_for_credit(text):
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    list_html = ""
    for line in lines:
        if "情况如下：" in line or "【" in line:
             list_html += f"<div style='font-weight: 700; margin-top: 8px; margin-bottom: 4px; color: #794f27;'>{line.replace('**', '')}</div>"
        else:
             list_html += f"<div style='margin-left: 10px; margin-bottom: 4px; color: #725d42; line-height: 1.6;'>• {line}</div>"
    return list_html

# ==========================================
# 动森风格弹窗：核心数据更新 Modal
# ==========================================
@st.dialog("🔒 核心数据更新")
def core_data_update_modal():
    card_bg = st.session_state.get("card_bg", "#FFFFFF")
    card_border = st.session_state.get("card_border", "#c4b89e")

    st.markdown(f"""
        <div style="font-size: 0.85rem; color: #725d42; background: {card_bg}; padding: 16px 20px; border-radius: 20px; margin-bottom: 10px; line-height: 1.6; border: 2.5px solid {card_border}; box-shadow: 0 4px 10px rgba(107, 92, 67, 0.10);">
            <b>📌 上传规范说明：</b><br>
            1. <b>文件名称</b>：必须包含 <span style='color:#c9302c'>客户关系清单</span><br>
            2. <b>表单名称</b>：必须有 <span style='color:#4d6bfe'>总</span>、<span style='color:#4d6bfe'>内部</span> 两个Sheet<br>
            3. <b>[总]</b> 列名需含：客户名称、客户所属集团<br>
            4. <b>[内部]</b> 列名需含：客户名称、所属专业化公司
        </div>
    """, unsafe_allow_html=True)

    pwd = st.text_input("Admin", type="password", placeholder="请输入通行证以解锁上传区...", label_visibility="collapsed")

    if pwd == "xuyingzhe":
        st.success("✅ 身份验证通过")
        new_mapping_file = st.file_uploader("上传清单", type=['xlsx'], label_visibility="collapsed")

        if new_mapping_file:
            if st.button("☁️ 确认并云端同步", use_container_width=True):
                with st.spinner("校验并同步中..."):
                    try:
                        df_total = pd.read_excel(new_mapping_file, sheet_name='总')
                        df_internal = pd.read_excel(new_mapping_file, sheet_name='内部')
                        if '客户名称' not in df_total.columns or '客户所属集团' not in df_total.columns:
                            st.error("❌ 拦截：【总】表缺少关键列！")
                        elif '客户名称' not in df_internal.columns or '所属专业化公司' not in df_internal.columns:
                            st.error("❌ 拦截：【内部】表缺少关键列！")
                        else:
                            from github import Github
                            gh_token = st.secrets["GITHUB_TOKEN"]
                            g = Github(gh_token)
                            repo = g.get_repo("Xuyingz99/CTMR")

                            new_mapping_file.seek(0)
                            content_bytes = new_mapping_file.read()

                            file_path = "客户关系清单.xlsx"
                            try:
                                contents = repo.get_contents(file_path)
                                repo.update_file(contents.path, "Admin: 网页端在线热更新客户关系清单", content_bytes, contents.sha)
                            except Exception:
                                repo.create_file(file_path, "Admin: 网页端创建客户关系清单", content_bytes)

                            st.success("🎉 云端更新成功！网页即将自动刷新。")
                            import time
                            time.sleep(2)
                            st.rerun()
                    except Exception as e:
                        st.error(f"❌ 更新失败: {str(e)}")
    elif pwd != "":
        st.error("❌ 密码错误")

# ==========================================
# 动森风格弹窗：网页说明
# ==========================================
@st.dialog("📖 CTMR 智能风控终端使用手册", width="large")
def help_modal():
    st.markdown("""
<div style="font-size: 0.95rem; color: var(--ac-text); line-height: 1.6;">
<h3 style="color: var(--ac-green);">📊 关于 Take It Easy</h3>
<p>本终端（CTMR）是专为粮食贸易风控条线量身定制的自动化数据中枢。我们将繁杂、冗长的传统多源报表处理，重塑为极简的“一键式”智能流水线，为您提供高效、优雅、精准的风险洞察体验。</p>

<hr style="border: 0; border-top: 1px solid var(--ac-wood); margin: 15px 0;">

<h4 style="color: var(--ac-green);">🆚 破局传统：为什么 CTMR 是更优解？</h4>
<p>传统商业智能 (BI) 与 Power Query 等工具在泛用型数据处理上固然强大，但在应对风控条线高频、定制化的报表需求时，往往面临“大而全却不精”的局限。CTMR 致力于解决传统工具的以下痛点，提供不可替代的专属体验：</p>
<ul style="padding-left: 20px;">
<li><b>打通数据汇报的“最后一公里” (End-to-End Output)：</b> 传统 BI 工具的终点往往是“可视化看板”，用户看完数据后仍需手动截图拼凑 Word。而 CTMR 能够穿透底层数据，直接封装为<b>排版就绪、带标准通报话术的 Word 催收函与风控简报</b>，彻底消除“看数据”与“写报告”的割裂感。</li>
<li><b>零代码的专家级风控引擎 (Hyper-Customization)：</b> 使用 Power Query 往往要求人员具备编写 M 函数或 DAX 的能力。CTMR 将特化的复杂业务逻辑（如：智能感知 1-49 吨与大宗标的的单位跃升、特定的百分比映射规则）<b>黑盒化封装</b>。它不仅是清洗工具，更是开箱即用的“数字风控专家”。</li>
<li><b>极简架构与“阅后即焚”的合规保障 (Ephemeral Security)：</b> 传统工具往往需要安装沉重客户端，极易在本地留下数据缓存。CTMR 采用极轻量架构，即开即用；且计算全过程在内存中流转，<b>阅后即焚、零数据落盘</b>，完美规避客户数据泄露的合规风险。</li>
</ul>

<h4 style="color: var(--ac-green); margin-top:20px;">🛡️ 企业级数据安全与底层架构</h4>
<ul style="padding-left: 20px;">
<li><b>瞬态计算 (Ephemeral Computing)：</b> 生命周期与当前会话严格绑定，任务结束即刻熔断销毁，系统零落盘。</li>
<li><b>端到端加密 (E2EE)：</b> 平台部署依托 TLS/SSL 军工级加密通道，保障传输绝对安全。</li>
<li><b>最小特权矩阵 (Least Privilege)：</b> 核心 Token 均通过环境变量进行黑盒加密托管。</li>
<li><b>逻辑沙箱隔离 (Logical Sandboxing)：</b> 前端交互无法逆向击穿或篡改源数据底座。</li>
</ul>

<h4 style="color: var(--ac-green); margin-top:20px;">🚀 核心业务引擎与数智亮点</h4>
<table style="width: 100%; border-collapse: collapse; margin-top: 10px; border: 1px solid var(--ac-card-border);">
<tr style="background-color: var(--ac-info-bg);">
<th style="padding: 10px; border: 1px solid var(--ac-card-border); text-align: left; width: 30%;">核心功能引擎</th>
<th style="padding: 10px; border: 1px solid var(--ac-card-border); text-align: left;">数智处理亮点</th>
</tr>
<tr>
<td style="padding: 10px; border: 1px solid var(--ac-card-border);"><b>初始保证金智能对账</b><br><span style="font-size: 0.8rem; opacity: 0.8;">(logic_init)</span></td>
<td style="padding: 10px; border: 1px solid var(--ac-card-border);">支持多期报表智能化比对，自动完成清洗提纯，动态渲染 A 类逾期汇总并秒级输出通报文案。</td>
</tr>
<tr>
<td style="padding: 10px; border: 1px solid var(--ac-card-border);"><b>追加保证金动态回填</b><br><span style="font-size: 0.8rem; opacity: 0.8;">(logic_add)</span></td>
<td style="padding: 10px; border: 1px solid var(--ac-card-border);">多大区数据矩阵无缝穿透，自动匹配最新日度数据流，一键生成契合管理层决策视角的定制化报告。</td>
</tr>
<tr>
<td style="padding: 10px; border: 1px solid var(--ac-card-border);"><b>逾期销售风控中枢</b><br><span style="font-size: 0.8rem; opacity: 0.8;">(logic_XS)</span></td>
<td style="padding: 10px; border: 1px solid var(--ac-card-border);">多源复杂数据流的高速合并与多维账期穿透，智能封装周报文档，同步生成标准催收文本。</td>
</tr>
<tr>
<td style="padding: 10px; border: 1px solid var(--ac-card-border);"><b>逾期采购智能对齐</b><br><span style="font-size: 0.8rem; opacity: 0.8;">(logic_CG)</span></td>
<td style="padding: 10px; border: 1px solid var(--ac-card-border);">多维采购日报的智能合并与对齐，算法级消除多表冗余，自动化提炼并渲染全局采购情况摘要。</td>
</tr>
<tr>
<td style="padding: 10px; border: 1px solid var(--ac-card-border);"><b>信用全景透视日报</b><br><span style="font-size: 0.8rem; opacity: 0.8;">(logic_credit)</span></td>
<td style="padding: 10px; border: 1px solid var(--ac-card-border);">深度抓取全局风险核心特征，内置专属格式防火墙与高亮映射，支持 Word 简报与可视化高清图表一键导出。</td>
</tr>
</table>

<h4 style="color: var(--ac-green); margin-top:20px;">💡 美学与交互巧思：减压办公哲学</h4>
<ul style="padding-left: 20px;">
<li><b>沉浸式双模主题：</b> 内置灵活的双主题切换引擎（浅色清新 / 深色专注）。一键无缝切换，昼夜交替间始终保持最佳视觉舒适度。</li>
<li><b>“动森海岛”治愈视觉：</b> 秉承减压设计哲学，采用舒缓色域、有机圆角边界及拟真光标，驱散传统数据面板的压迫感。</li>
<li><b>绝对防篡改锚点：</b> 无论系统后台逻辑如何高频迭代，底层部署的“前端 UI 防护墙”都会为您守住当前最完美的交互界面。</li>
</ul>
<p style="text-align: right; font-weight: bold; margin-top: 20px; opacity: 0.7;">—— Take It Easy 自动化工程团队</p>
</div>
    """, unsafe_allow_html=True)


# 主界面逻辑
# ==========================================

def main():
    # 根据当前主题设置卡片色板（供全局 inline style 使用）
    t = THEMES[st.session_state.current_theme]
    st.session_state.card_bg = t["--ac-card-bg"]
    st.session_state.card_border = t["--ac-card-border"]

    col_l, col_center, col_admin = st.columns([1, 6, 1])
    
    with col_center:
        st.markdown("""
            <div class="header-container" style="padding-bottom: 0rem; padding-top: 1rem;">
                <h1 class="main-title">Take It Easy</h1>
                <div class="sub-title">Crafted by Xuyingzhe</div>
            </div>
        """, unsafe_allow_html=True)

    with col_admin:
        st.markdown("<div style='margin-top: 30px;'></div>", unsafe_allow_html=True)
        with st.expander("⚙️ 设置", expanded=False):

            # --- 🎨 开放功能：主题切换 ---
            st.markdown("<div style='font-size:0.85rem; font-weight:bold; color:var(--ac-text); margin-bottom:8px;'>🎨 界面主题</div>", unsafe_allow_html=True)

            theme_keys = list(THEMES.keys())
            current_index = theme_keys.index(st.session_state.current_theme)

            selected_theme = st.selectbox(
                "选择主题",
                theme_keys,
                index=current_index,
                label_visibility="collapsed"
            )

            if selected_theme != st.session_state.current_theme:
                st.session_state.current_theme = selected_theme
                st.rerun()

            st.markdown("<hr style='margin: 10px 0; border: 1px solid #eaeaea;'>", unsafe_allow_html=True)

            # --- 🔒 私密功能：核心数据更新 ---
            # (已移除多余的文本标题，仅保留操作按钮，使界面更清爽)
            if st.button("数据更新管理", use_container_width=True):
                core_data_update_modal()

            if st.button("网页说明", use_container_width=True):
                help_modal()

    col_space_l, col_center_main, col_space_r = st.columns([1, 6, 1])

    with col_center_main:
        st.markdown('<div class="greeting-text">有什么我能帮你的么？</div>', unsafe_allow_html=True)

        function_map = {
            "📈 初始保证金处理": "init_margin",
            "📉 追加保证金处理": "add_margin",
            "⏱️ 逾期销售处理": "overdue_sales",
            "🛒 逾期采购处理": "overdue_purchase",
            "📊 信用风险管理日报": "credit_report"
        }

        mode = st.radio("选择功能", list(function_map.keys()), horizontal=True, label_visibility="collapsed")
        
        # --- 模块 1: 初始保证金处理 ---
        if mode == "📈 初始保证金处理":
            st.markdown("""
            <div class="info-box">
                <div class="info-title">⚠️ 注意事项</div>
                <div style="margin-left: 2px;">
                    <div>请务必同时上传两个文件以便进行数据比对</div>
                    <div style="margin-top: 4px;">原始表单 Sheet 名称必须包含 WSBZJQKB</div>
                    <div style="margin-top: 4px;">生成结果将包含清洗后的明细表及 A 类逾期汇总</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            c1, c2 = st.columns(2)
            with c1:
                current_file = st.file_uploader("📂 1. 上传【今日】报表", type=['xlsx'])
            with c2:
                prev_file = st.file_uploader("📂 2. 上传【对照日】报表", type=['xlsx'])
            
            if st.button("🚀 开始处理 / Analyze"):
                if current_file and prev_file:
                    with st.spinner("正在进行数据比对与清洗，请稍候..."):
                        excel_data, report_logs = process_margin_deposit_logic(current_file, prev_file)
                        
                        if excel_data:
                            st.success("✅ 处理完成！")
                            
                            today_dt = datetime.now()
                            custom_filename = f"{today_dt.month}.{today_dt.day}(未收保证金情况表)--沿海大区.xlsx"
                            
                            st.download_button(
                                label=f"📥 下载处理后的报表 ({custom_filename})",
                                data=excel_data,
                                file_name=custom_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
                            st.markdown("### 📢 生成的通报文案")
                            for log in report_logs:
                                if str(log).startswith("<div"):
                                    st.markdown(log, unsafe_allow_html=True)
                                else:
                                    st.markdown(f"""
                                    <div style="background: {st.session_state.card_bg}; padding: 16px 22px; border-radius: 20px;
                                                border: 2.5px solid {st.session_state.card_border}; margin-bottom: 12px;
                                                box-shadow: 0 4px 10px rgba(107, 92, 67, 0.10);">
                                        <span style="color: #725d42; font-weight: 500; line-height: 1.7;">{log}</span>
                                    </div>
                                    """, unsafe_allow_html=True)
                                
                        else:
                            st.error("处理失败，请查看下方错误日志")
                            st.code(report_logs[-1])
                else:
                    st.warning("⚠️ 请确保两个文件都已上传！")
                    
        # --- 模块 2: 追加保证金处理 ---
        elif mode == "📉 追加保证金处理":
            st.markdown("""
            <div class="info-box">
                <div class="info-title">⚠️ 注意事项</div>
                <div style="margin-left: 2px;">
                    <div>请务必上传“追加保证金填报表”</div>
                    <div style="margin-top: 4px;">系统将自动进行筛选、数据清洗与报告生成</div>
                    <div style="margin-top: 4px;">下方选择相应大区，即可生成专属定制报告</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.markdown('<div style="margin-bottom: 8px; font-weight: 700; color: #794f27;">选择报告生成范围</div>', unsafe_allow_html=True)
            region_options = ["中粮贸易", "沿海大区", "沿江大区", "内陆大区", "东北大区"]
            
            selection = st.pills("选择报告生成范围", region_options, default="中粮贸易", label_visibility="collapsed")
            selected_region = selection if selection is not None else "中粮贸易"

            uploaded_file = st.file_uploader("📂 上传【追加保证金填报表】", type=['xlsx'])

            if st.button("🚀 生成报告 / Generate Report"):
                if uploaded_file:
                    with st.spinner(f"正在为【{selected_region}】生成专属报告..."):
                        output_file, logs, report_a, report_b, max_date_str = process_additional_margin_logic(uploaded_file, selected_region)
                        
                        if output_file:
                            st.success(f"✅ {selected_region}报告生成完成！")
                            
                            today_mmdd = datetime.now().strftime('%m%d')
                            file_prefix = "中粮贸易" if selected_region == "中粮贸易" else f"{selected_region}"
                            dl_filename = f"{file_prefix}追加保证金填报表{today_mmdd}-截至{max_date_str}数据.xlsx"
                            st.download_button(
                                label=f"📥 下载定制报告 ({dl_filename})",
                                data=output_file,
                                file_name=dl_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
                            c_a, c_b = st.columns(2)
                            with c_a:
                                display_pretty_report(f"业务单位报告 ({selected_region})", report_a, bold_first_para=False)
                            with c_b:
                                display_pretty_report(f"分客户报告 ({selected_region})", report_b, bold_first_para=False)
                        else:
                            st.error("处理失败")
                            for l in logs: st.write(l)
                else:
                    st.warning("⚠️ 请先上传文件！")

        # --- 模块 3: 逾期销售处理 ---
        elif mode == "⏱️ 逾期销售处理":
            st.markdown("""
            <div class="info-box">
                <div class="info-title">⚠️ 注意事项</div>
                <div style="margin-left: 2px;">
                    <div>请分别上传【逾期销售（分批次）】和【逾期销售（一次性）】的表格数据</div>
                    <div style="margin-top: 4px;">系统将自动整合数据、计算逾期金额、匹配客户信息</div>
                    <div style="margin-top: 4px;">勾选复选框，可同时生成周报 Word 文档及催收提醒文本</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            need_report = st.checkbox("📝 需要生成【逾期销售周报】(Word格式)", value=False)
            
            col1, col2 = st.columns(2)
            with col1:
                batch_files = st.file_uploader("📂 逾期销售（分批次） [最多6个]", type=["xlsx", "xls"], accept_multiple_files=True, key="batch_upload")
                if batch_files and len(batch_files) > 6:
                    st.warning("⚠️ 分批次文件最多只能上传6个，超出的部分将被忽略。")
                    batch_files = batch_files[:6]
                    
            with col2:
                once_files = st.file_uploader("📂 逾期销售（一次性） [最多6个]", type=["xlsx", "xls"], accept_multiple_files=True, key="once_upload")
                if once_files and len(once_files) > 6:
                    st.warning("⚠️ 一次性文件最多只能上传6个，超出的部分将被忽略。")
                    once_files = once_files[:6]
                    
            if st.button("🚀 开始处理逾期数据", key="btn_xs"):
                if not batch_files and not once_files:
                    st.warning("⚠️ 请至少在一个文件栏中上传数据文件！")
                else:
                    with st.spinner("正在高速运算并生成报告中..."):
                        excel_io, word_io, collection_text, logs = process_overdue_sales(batch_files, once_files, need_report)
                        
                        if excel_io:
                            st.success("✅ 逾期数据处理成功！")
                            
                            with st.expander("查看处理日志", expanded=False):
                                for log in logs:
                                    st.write(log)
                                    
                            st.markdown("### 📥 下载结果文件")
                            dl_col1, dl_col2 = st.columns(2)
                            
                            mmdd_str = datetime.now().strftime('%m%d')
                            yyyymmdd_str = datetime.now().strftime('%Y%m%d')
                            
                            with dl_col1:
                                st.download_button(
                                    label="📊 下载【逾期销售监控表】 (Excel)",
                                    data=excel_io,
                                    file_name=f"逾期销售监控表_{mmdd_str}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                                
                            if need_report and word_io:
                                with dl_col2:
                                    st.download_button(
                                        label="📝 下载【逾期销售周报】 (Word)",
                                        data=word_io,
                                        file_name=f"逾期销售周报_{yyyymmdd_str}.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        use_container_width=True
                                    )

                            if collection_text:
                                st.markdown("### 📢 生成的通报文案")
                                display_pretty_report("💬 催收提醒预览", collection_text, bold_first_para=True)
                        else:
                            st.error("❌ 处理失败，请检查文件格式是否符合要求。")

        # --- 模块 4: 逾期采购处理 ---
        elif mode == "🛒 逾期采购处理":
            st.markdown("""
            <div class="info-box">
                <div class="info-title">⚠️ 注意事项</div>
                <div style="margin-left: 2px;">
                    <div>请上传包含【逾期采购监控表-日报】数据的 Excel 文件（支持多选最多 6 个）。</div>
                    <div style="margin-top: 4px;">系统将自动对齐、合并和清洗数据，并为您生成对应完整的报告与监控台账。</div>
                    <div style="margin-top: 4px;">网页将直接预览逾期采购情况的总结文本。</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            cg_files = st.file_uploader("📂 上传【逾期采购数据】 [最多6个]", type=["xlsx", "xls"], accept_multiple_files=True, key="cg_upload")
            
            if cg_files and len(cg_files) > 6:
                st.warning("⚠️ 最多只能上传6个文件，超出的部分将被忽略。")
                cg_files = cg_files[:6]
                
            if st.button("🚀 开始处理逾期采购数据", key="btn_cg"):
                if not cg_files:
                    st.warning("⚠️ 请先上传数据文件！")
                else:
                    with st.spinner("正在智能比对与合并多表，生成极速简报中..."):
                        excel_io, doc_io, web_text, logs = process_overdue_purchase(cg_files)
                        
                        if excel_io:
                            st.success("✅ 逾期采购数据清洗成功，报表已生成！")
                            
                            # 展示日志（屏蔽掉纯提示性标志，以防重复）
                            for log in logs:
                                if log.startswith("✅"): continue
                                st.write(log)
                                
                            st.markdown("### 📥 下载专属文件")
                            dl_col1, dl_col2 = st.columns(2)
                            
                            mmdd_str = datetime.now().strftime('%m%d')
                            
                            with dl_col1:
                                st.download_button(
                                    label="📊 下载【逾期采购监控表】 (Excel)",
                                    data=excel_io,
                                    file_name=f"逾期采购监控表Z_{mmdd_str}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                                
                            if doc_io:
                                with dl_col2:
                                    st.download_button(
                                        label="📝 下载【逾期采购报告】 (Word)",
                                        data=doc_io,
                                        file_name=f"逾期采购报告_{mmdd_str}.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        use_container_width=True
                                    )
                                    
                            # 渲染展示在网页端的纯文本报告摘要，复用 style.py 内置的美化函数
                            if web_text:
                                st.markdown("### 📢 采购情况通报")
                                display_pretty_report("💬 逾期采购情况摘要", web_text, bold_first_para=True)
                        else:
                            st.error("❌ 处理失败。")
                            for log in logs: st.write(log)

        # --- 模块 5: 信用风险管理日报 ---
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
                    with st.spinner("正在解析 Excel 数据并渲染跨平台文件，请稍候..."):
                        word_bytes, word_text_dict, export_files, logs, env_msg = process_credit_report(uploaded_file)
                        
                        st.info(f"💡 {env_msg}")
                        
                        if word_bytes or export_files:
                            st.success("✅ 任务处理完成！")
                    
                            if word_text_dict:
                                st.markdown("<h3 style='margin-top: 10px; margin-bottom: 20px; color: #794f27;'>信用风险管理日报</h3>", unsafe_allow_html=True)
                                
                                center_themes = {
                                    "玉米": {"bg": "#eef5ff", "bd": "#d1e3ff", "bar": "#4d6bfe"},
                                    "粮谷": {"bg": "#ebf9f1", "bd": "#c3e8d1", "bar": "#28a745"},
                                    "大豆": {"bg": "#fff6e5", "bd": "#ffe2b3", "bar": "#fd7e14"} 
                                }
                                
                                for center_name, content in word_text_dict.items():
                                    theme = center_themes.get(center_name, {"bg": "#fcf8f2", "bd": "#f0e6d2", "bar": "#6c757d"})
                                    html_content = format_html_content_for_credit(content)
                                    
                                    st.markdown(f"""
                                    <div style="background-color: {theme['bg']}; padding: 20px 25px; border-radius: 20px; border: 2.5px solid {theme['bd']}; margin-bottom: 20px; box-shadow: 0 4px 10px rgba(107, 92, 67, 0.12);">
                                        {html_content}
                                    </div>
                                    """, unsafe_allow_html=True)
                            
                            st.markdown("### 📥 下载生成文件")
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
                        
                            png_files = [f for f in export_files if f["type"] == "png"]
                            if png_files:
                                st.markdown("#### 👁️ 图片预览")
                                for p_f in png_files:
                                    st.image(p_f["data"], caption=p_f["name"], use_container_width=True)

                        else:
                            st.error("处理失败，未能提取到有效数据。")
                else:
                    st.warning("⚠️ 请先上传 Excel 文件！")

    # 动森风格 Footer
    current_theme = st.session_state.get("current_theme", "🏝️ 狸克海岛")
    theme = THEMES[current_theme]
    text_color = theme["--ac-text"]
    if "狸克海岛" in current_theme:
        footer_bg = "https://cdn.jsdelivr.net/gh/guokaigdg/animal-island-ui@main/src/assets/img/footer/footer-tree.webp"
        footer_h = "120px"
        footer_props = "bottom center / cover no-repeat"
    else:
        footer_bg = "https://cdn.jsdelivr.net/gh/guokaigdg/animal-island-ui@main/src/assets/img/footer/footer-sea.svg"
        footer_h = "80px"
        footer_props = "center / contain no-repeat"

    st.markdown(f"""
    <div style="
        width: 100vw; margin-left: calc(-50vw + 50%); margin-top: 150px;
        height: {footer_h};
        background: url('{footer_bg}') {footer_props}; position: relative; display: flex;
        align-items: flex-end; justify-content: center; padding-bottom: 20px; overflow: hidden;
    ">
        <div style="color:{text_color}; font-weight: 800; font-size: 1.05rem; letter-spacing: 1px; text-shadow: 1px 1px 0px #fff;">
            &copy; 2026 Take It Easy
        </div>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
