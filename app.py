import streamlit as st
import warnings
from datetime import datetime

# --- 引入模块 ---
from utils import style, logic_init, logic_add

# 忽略警告
warnings.filterwarnings('ignore')

# --- 页面基础配置 ---
st.set_page_config(
    page_title="Take It Easy - 智能办公助手",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 1. 加载样式 ---
style.apply_custom_css()

# ==========================================
# 主界面逻辑
# ==========================================

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
            "📝 格式转换 (Demo)": "demo"
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
                    with st.spinner("🤖 正在进行数据比对与清洗，请稍候..."):
                        # 调用 utils/logic_init.py 中的函数
                        excel_data, report_logs = logic_init.process_margin_deposit_logic(current_file, prev_file)
                        
                        if excel_data:
                            st.success("✅ 处理完成！")
                            st.markdown("### 📢 生成的通报文案")
                            for log in report_logs:
                                st.info(log)
                                
                            st.download_button(
                                label=f"📥 下载处理后的报表 ({current_file.name})",
                                data=excel_data,
                                file_name=current_file.name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
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

            # 大区筛选器
            st.markdown('<div style="margin-bottom: 8px; font-weight: 600; color: #333;">选择报告生成范围</div>', unsafe_allow_html=True)
            region_options = ["中粮贸易", "沿海大区", "沿江大区", "内陆大区", "东北大区"]
            
            selection = st.pills(
                "选择报告生成范围", 
                region_options, 
                default="中粮贸易", 
                label_visibility="collapsed"
            )
            
            # 逻辑兜底，防止取消选择
            if selection is None:
                selected_region = "中粮贸易"
            else:
                selected_region = selection

            uploaded_file = st.file_uploader("📂 上传【追加保证金填报表】", type=['xlsx'])
            
            if st.button("🚀 生成报告 / Generate Report"):
                if uploaded_file:
                    with st.spinner(f"🤖 正在为【{selected_region}】生成专属报告..."):
                        # 调用 utils/logic_add.py 中的函数
                        output_file, logs, report_a, report_b = logic_add.process_additional_margin_logic(uploaded_file, selected_region)
                        
                        if output_file:
                            st.success(f"✅ {selected_region}报告生成完成！")
                            
                            c_a, c_b = st.columns(2)
                            with c_a:
                                # 调用 utils/style.py 中的函数
                                style.display_pretty_report(f"业务单位报告 ({selected_region})", report_a, "#eef5ff")
                            with c_b:
                                style.display_pretty_report(f"分客户报告 ({selected_region})", report_b, "#fff8e6")
                            
                            today_mmdd = datetime.now().strftime('%m%d')
                            file_prefix = "" if selected_region == "中粮贸易" else f"{selected_region}"
                            dl_filename = f"{file_prefix}追加保证金填报表{today_mmdd}.xlsx"
                            
                            st.download_button(
                                label=f"📥 下载定制报告 ({dl_filename})",
                                data=output_file,
                                file_name=dl_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.error("处理失败")
                            for l in logs: st.write(l)
                else:
                    st.warning("⚠️ 请先上传文件！")

        else:
            st.info("此功能暂未开放，敬请期待...")
            st.file_uploader("上传文件", disabled=True)
            st.button("Analyze", disabled=True)

    st.markdown("<div style='text-align:center; color:#ccc; margin-top:50px;'>© 2026 TakeItEasy Tool</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()