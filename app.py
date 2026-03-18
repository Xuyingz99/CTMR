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

# === 导入功能模块 ===
from utils.logic_credit import process_credit_report
from utils.logic_XS import process_overdue_sales  # 新增：引入刚才写好的逾期销售处理模块

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

# ============================================================================
# PART 1: 初始保证金处理逻辑 (XSchushi.txt / app.py 原有逻辑)
# ============================================================================

def read_excel_safe(file_stream):
    try:
        file_stream.seek(0)
        df = pd.read_excel(file_stream, sheet_name="WSBZJQKB", dtype={'合同编号': str})
        if '合同编号' not in df.columns:
            file_stream.seek(0)
            df_temp = pd.read_excel(file_stream, sheet_name="WSBZJQKB", header=None, nrows=200)
            header_idx = -1
            for idx, row in df_temp.iterrows():
                if "合同编号" in row.values:
                    header_idx = idx
                    break
            if header_idx != -1:
                file_stream.seek(0)
                df = pd.read_excel(file_stream, sheet_name="WSBZJQKB", header=header_idx, dtype={'合同编号': str})
            else:
                raise ValueError("在文件前200行中无法找到包含'合同编号'的标题行，请检查文件格式。")
        return df
    except Exception as e:
        raise e

def fill_original_sheet_columns(ws_original, df_data):
    try:
        col_reason = get_column_by_name(ws_original, "逾期具体原因")
        col_type = get_column_by_name(ws_original, "逾期原因分类")
        col_client = get_column_by_name(ws_original, "客户")
        left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))
        if col_reason and col_type and col_client:
             for i, row_cells in enumerate(ws_original.iter_rows(min_row=2), start=0):
                if i >= len(df_data): break
                cell_reason = row_cells[col_reason - 1]
                cell_type = row_cells[col_type - 1]
                cell_client = row_cells[col_client - 1]
                new_reason_val = df_data.iloc[i].get("逾期具体原因_新", "")
                new_type_val = df_data.iloc[i].get("逾期原因分类_新", "")
                if cell_reason.value is None or str(cell_reason.value).strip() == "":
                    cell_reason.value = new_reason_val
                    if cell_client.has_style:
                        cell_reason.font = copy.copy(cell_client.font)
                        cell_reason.fill = copy.copy(cell_client.fill)
                    cell_reason.alignment = left_align
                    cell_reason.border = thin_border
                if cell_type.value is None or str(cell_type.value).strip() == "":
                    cell_type.value = new_type_val
                    if cell_client.has_style:
                        cell_type.font = copy.copy(cell_client.font)
                        cell_type.fill = copy.copy(cell_client.fill)
                    cell_type.alignment = left_align
                    cell_type.border = thin_border
        for row in ws_original.iter_rows():
            ws_original.row_dimensions[row[0].row].height = 24.5
            for cell in row:
                cell.border = thin_border
                if cell.alignment:
                    new_align = copy.copy(cell.alignment)
                    new_align.vertical = 'center'
                    cell.alignment = new_align
                else:
                    cell.alignment = Alignment(vertical='center')
    except Exception as e: pass

def get_true_column_width(value):
    if value is None: return 0
    str_val = str(value)
    width = 0
    for char in str_val:
        if ord(char) > 255: width += 2.1
        elif char.isupper() or char.isdigit(): width += 1.2
        else: width += 1.0
    return width

def auto_fit_columns(worksheet, min_width=10, max_width=60):
    custom_widths = {
        "序号": 6, "业务部门": 14, "合同编号": 28, "客户": 35, "品种": 10,
        "合同数量": 14, "合同单价": 14, "合同金额": 16, "应收保证金日期": 18,
        "应收保证金比例": 16, "应收保证金金额": 18, "已收定金/预收款": 18,
        "逾期初始保证金金额": 22
    }
    for col in worksheet.columns:
        column_letter = get_column_letter(col[0].column)
        header_text = str(col[0].value).strip() if col[0].value else ""
        matched_width = None
        for key, width in custom_widths.items():
            if key in header_text:
                matched_width = width
                break
        if matched_width:
            worksheet.column_dimensions[column_letter].width = matched_width
            continue
        max_length = 0
        for cell in col:
            try:
                if cell.value:
                    cell_width = get_true_column_width(cell.value)
                    if cell_width > max_length: max_length = cell_width
            except: pass
        adjusted_width = min(max(max_length + 3, min_width), max_width)
        worksheet.column_dimensions[column_letter].width = adjusted_width

def find_header_row(worksheet):
    try:
        max_search_rows = min(200, worksheet.max_row)
        critical_field = "合同编号"
        for row_idx in range(1, max_search_rows + 1):
            row_values = []
            for col_idx in range(1, min(20, worksheet.max_column) + 1):
                cell_value = worksheet.cell(row_idx, col_idx).value
                if cell_value: row_values.append(str(cell_value).strip())
            for val in row_values:
                if critical_field in val: return row_idx
        for row_idx in range(1, max_search_rows + 1):
            for col_idx in range(1, min(20, worksheet.max_column) + 1):
                val = str(worksheet.cell(row_idx, col_idx).value or "")
                if "业务部门" in val: return row_idx
        return 1
    except: return 1

def remove_empty_rows(worksheet):
    try:
        header_row = find_header_row(worksheet)
        if header_row > 1:
            rows_to_delete = header_row - 1
            worksheet.delete_rows(1, rows_to_delete)
            return True
        return True
    except: return False

def get_column_by_name(worksheet, column_names):
    if isinstance(column_names, str): column_names = [column_names]
    for col in range(1, worksheet.max_column + 1):
        cell_value = worksheet.cell(row=1, column=col).value
        if cell_value:
            for col_name in column_names:
                if col_name in str(cell_value).strip(): return col
    return None
def beautify_sheet_common(ws, title_color="BDD7EE"):
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color=title_color, end_color=title_color, fill_type="solid")
    header_font = Font(color="000000", bold=True, size=11)
    light_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for row in range(2, ws.max_row + 1):
        row_bg_fill = white_fill if row % 2 == 0 else light_fill
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            current_fill = cell.fill
            is_yellow_cell = False
            if current_fill and current_fill.start_color and current_fill.start_color.rgb:
                if str(current_fill.start_color.rgb).endswith("FFFF00"): is_yellow_cell = True
            if not is_yellow_cell: cell.fill = row_bg_fill
            if not cell.font.color or cell.font.color.rgb == "00000000": cell.font = Font(size=10)
            cell.alignment = center_align
    ws.row_dimensions[1].height = 25
    for row in range(2, ws.max_row + 1): ws.row_dimensions[row].height = 22
    ws.freeze_panes = 'A2'

def clean_and_organize_A_sheet(ws_A):
    try:
        columns_to_delete = ["区域公司", "公司名称", "销售类型", "业务模式", "合同提交日期", "合同签订日期", "合同生效日期", "出库数量", "是否约定保证金条款", "合同约定几个工作日收取", "已收货款金额（不含保证金）", "逾期具体原因", "逾期原因分类", "逾期具体原因_新", "逾期原因分类_新"]
        cols_found = []
        for col in range(1, ws_A.max_column + 1):
            val = str(ws_A.cell(row=1, column=col).value)
            for target in columns_to_delete:
                if target in val:
                    cols_found.append(col)
                    break
        for col_idx in sorted(cols_found, reverse=True): ws_A.delete_cols(col_idx, 1)
        data = list(ws_A.values)
        if not data: return False
        headers = data[0]
        df = pd.DataFrame(data[1:], columns=headers)
        date_col = next((c for c in df.columns if "应收保证金日期" in str(c)), None)
        if date_col:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%Y-%m-%d')
            df = df.sort_values(by=date_col)
        dept_col = next((c for c in df.columns if "业务部门" in str(c)), None)
        if dept_col:
            replacements = ['沿海深圳', '食品原料部', '经营部', '中粮贸易（深圳）有限公司-', '（旧）']
            for r in replacements: df[dept_col] = df[dept_col].astype(str).str.replace(r, '', regex=False)
        ws_A.delete_rows(2, ws_A.max_row)
        for r_idx, row in enumerate(df.values, 2):
            for c_idx, val in enumerate(row, 1): ws_A.cell(row=r_idx, column=c_idx, value=val)
        serial_col = get_column_by_name(ws_A, "序号")
        contract_col = get_column_by_name(ws_A, "合同编号")
        if serial_col and contract_col:
            col_letter = get_column_letter(contract_col)
            for r in range(2, ws_A.max_row + 1): ws_A.cell(row=r, column=serial_col, value=f'=SUBTOTAL(103, ${col_letter}$2:{col_letter}{r})')
        numeric_cols = ["合同数量", "合同单价", "合同金额", "应收保证金金额", "已收定金", "逾期初始保证金"]
        for col_name in numeric_cols:
            col_idx = get_column_by_name(ws_A, col_name)
            if col_idx:
                for r in range(2, ws_A.max_row + 1):
                    cell = ws_A.cell(row=r, column=col_idx)
                    try:
                        if cell.value:
                            cell.value = float(cell.value)
                            cell.number_format = '0.00'
                    except: pass
        pct_col = get_column_by_name(ws_A, "应收保证金比例")
        if pct_col:
            for r in range(2, ws_A.max_row + 1):
                cell = ws_A.cell(row=r, column=pct_col)
                try:
                    if cell.value:
                        cell.value = float(cell.value)
                        cell.number_format = '0%'
                except: pass
        return True
    except: return False

def optimize_A_sheet_formatting(ws_A):
    try:
        today = datetime.now().date()
        date_column = get_column_by_name(ws_A, "应收保证金日期")
        if date_column:
            dark_red_font = Font(color="8B0000")
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            for row in range(2, ws_A.max_row + 1):
                cell = ws_A.cell(row=row, column=date_column)
                try:
                    if cell.value:
                        cell_date_str = str(cell.value)
                        cell_date = None
                        for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"]:
                            try:
                                cell_date = datetime.strptime(cell_date_str, fmt).date()
                                break
                            except: continue
                        if cell_date and cell_date <= today:
                            for col in range(1, ws_A.max_column + 1): ws_A.cell(row=row, column=col).font = dark_red_font
                            if cell_date < today: cell.fill = yellow_fill
                except: continue
        beautify_sheet_common(ws_A, title_color="BDD7EE")
        right_align_keywords = ["应收保证金日期", "应收保证金比例", "应收保证金金额", "已收定金/预收款", "逾期初始保证金金额"]
        right_align = Alignment(horizontal='right', vertical='center', wrap_text=True)
        for keyword in right_align_keywords:
            col_idx = get_column_by_name(ws_A, keyword)
            if col_idx:
                for row in range(2, ws_A.max_row + 1): ws_A.cell(row=row, column=col_idx).alignment = right_align
        auto_fit_columns(ws_A)
    except: pass

def create_A_summary_sheet(workbook, ws_A, today_date_str):
    try:
        if "A类逾期明细汇总" in workbook.sheetnames: del workbook["A类逾期明细汇总"]
        ws_summary = workbook.create_sheet("A类逾期明细汇总")
        ws_summary.append(["业务部门", "提醒内容"])
        today_date = datetime.strptime(today_date_str, "%Y.%m.%d")
        yesterday_str = (today_date - timedelta(days=1)).strftime("%m月%d日")
        business_dept_col = get_column_by_name(ws_A, "业务部门")
        date_col = get_column_by_name(ws_A, "应收保证金日期")
        if not business_dept_col or not date_col: return False, []
        dept_stats = {}
        for row in range(2, ws_A.max_row + 1):
            dept_name = ws_A.cell(row=row, column=business_dept_col).value
            if not dept_name: dept_name = "未知部门"
            if dept_name not in dept_stats: dept_stats[dept_name] = {'total': 0, 'yellow_cells': 0, 'non_yellow_cells': 0}
            dept_stats[dept_name]['total'] += 1
            cell_fill = ws_A.cell(row=row, column=date_col).fill
            is_yellow = False
            if cell_fill and cell_fill.start_color and cell_fill.start_color.rgb:
                if str(cell_fill.start_color.rgb).endswith("FFFF00"): is_yellow = True
            if is_yellow: dept_stats[dept_name]['yellow_cells'] += 1
            else: dept_stats[dept_name]['non_yellow_cells'] += 1
        logs = []
        row_idx = 2
        for dept_name, stats in dept_stats.items():
            if stats['yellow_cells'] > 0:
                reminder_text = f"【逾期初始保证金】各位领导同事，截至{yesterday_str}，{dept_name}经营部初始保证金{stats['yellow_cells']}笔逾期，{stats['non_yellow_cells']}笔即将到期，请核对并及时催收，谢谢！ @所有人"
            else:
                reminder_text = f"【逾期初始保证金】各位领导同事，截至{yesterday_str}，{dept_name}经营部初始保证金{stats['non_yellow_cells']}笔即将到期，请核对并及时催收，谢谢！ @所有人"
            ws_summary.cell(row=row_idx, column=1, value=dept_name)
            ws_summary.cell(row=row_idx, column=2, value=reminder_text)
            clean_log = reminder_text.replace('\n', '').replace('\r', '')
            logs.append(f"📌 {dept_name}: {clean_log}")
            row_idx += 1
        beautify_sheet_common(ws_summary, title_color="BDD7EE")
        dept_len = 0
        for cell in ws_summary['A']:
            val_len = get_true_column_width(cell.value)
            if val_len > dept_len: dept_len = val_len
        ws_summary.column_dimensions['A'].width = min(max(dept_len + 4, 15), 40)
        fixed_text_width = 90
        ws_summary.column_dimensions['B'].width = fixed_text_width
        left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
        for row in range(2, ws_summary.max_row + 1): ws_summary.cell(row=row, column=2).alignment = left_align
        for row in range(2, ws_summary.max_row + 1):
            cell_val = str(ws_summary.cell(row=row, column=2).value or "")
            text_width = get_true_column_width(cell_val)
            estimated_lines = math.ceil(text_width / (fixed_text_width - 5))
            if estimated_lines <= 1: row_height = 25
            else: row_height = 20 + (estimated_lines - 1) * 18
            ws_summary.row_dimensions[row].height = row_height
        return True, logs
    except: return False, []
# ============================================================================
# PART 2: Streamlit 主页面渲染与逻辑分发
# ============================================================================

def main():
    # 顶部标题区域
    st.markdown("""
        <div class="header-container">
            <h1 class="main-title">Take It Easy</h1>
            <div class="sub-title">Intelligent Office Assistant</div>
        </div>
        <div class="greeting-text">你好！今天需要处理什么？</div>
    """, unsafe_allow_html=True)

    # 核心菜单栏：原 "格式转换 (Demo)" 现替换为 "逾期销售处理"
    action = st.radio(
        "选择功能",
        ["初始保证金处理", "逾期销售处理", "信用风险管理"],
        horizontal=True,
        label_visibility="collapsed"
    )

    # ======================== 分支 1：初始保证金处理 ========================
    if action == "初始保证金处理":
        st.markdown("""
        <div class="info-box">
            <div class="info-title">📊 初始保证金处理</div>
            请上传从【NC系统】导出的未收保证金情况表。<br>
            系统将自动清理无效数据，标记超期预警，并按大区生成催收提醒内容。
        </div>
        """, unsafe_allow_html=True)

        uploaded_file = st.file_uploader("📂 拖拽或点击上传 Excel 文件", type=["xlsx", "xls"], key="margin_upload")
        
        if st.button("🚀 开始处理数据", key="btn_margin"):
            if uploaded_file is not None:
                with st.spinner("正在拼命处理中，请稍候..."):
                    try:
                        df_original = read_excel_safe(uploaded_file)
                        wb = openpyxl.load_workbook(uploaded_file)
                        ws_A = None
                        if "WSBZJQKB" in wb.sheetnames:
                            ws_original = wb["WSBZJQKB"]
                            ws_original.title = "未收保证金情况表"
                            fill_original_sheet_columns(ws_original, df_original)
                            ws_A = wb.copy_worksheet(ws_original)
                            ws_A.title = "A类逾期明细"
                        if ws_A and clean_and_organize_A_sheet(ws_A):
                            optimize_A_sheet_formatting(ws_A)
                            today_str = datetime.now().strftime("%Y.%m.%d")
                            success, logs = create_A_summary_sheet(wb, ws_A, today_str)
                            
                            output = io.BytesIO()
                            wb.save(output)
                            output.seek(0)
                            
                            st.success("✅ 处理完成！")
                            if logs:
                                st.markdown("### 💬 各大区催收提醒")
                                for log in logs:
                                    st.info(log)
                                    
                            st.download_button(
                                label="📥 下载处理后的 Excel 文件",
                                data=output,
                                file_name=f"初始保证金情况跟踪表_{datetime.now().strftime('%m%d')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        else:
                            st.error("处理工作表失败，请检查数据格式。")
                    except Exception as e:
                        st.error(f"处理发生错误: {str(e)}")
            else:
                st.warning("⚠️ 请先上传 Excel 文件！")

    # ======================== 分支 2：逾期销售处理 (新增) ========================
    elif action == "逾期销售处理":
        st.markdown("""
        <div class="info-box">
            <div class="info-title">⏱️ 逾期销售处理</div>
            请分别上传【逾期销售（分批次）】和【逾期销售（一次性）】的表格数据。<br>
            系统将自动整合数据、计算逾期金额、匹配客户信息，并生成监控总表及催收提示。
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
                        
                        # 展示处理日志
                        with st.expander("查看处理日志", expanded=False):
                            for log in logs:
                                st.write(log)
                                
                        # 重点展示催收提醒文本
                        if collection_text:
                            st.markdown("### 💬 催收提醒预览")
                            st.markdown(f"""
                            <div style="background-color: #f8f9fa; border-left: 4px solid var(--deepseek-blue); padding: 15px; border-radius: 5px; font-size: 0.95rem; line-height: 1.6; white-space: pre-wrap;">{collection_text}</div>
                            """, unsafe_allow_html=True)

                        # 下载按钮区域
                        st.markdown("### 📥 下载结果文件")
                        dl_col1, dl_col2 = st.columns(2)
                        
                        # 动态获取 MMDD 日期（去除前导零的要求已在 logic 层兼顾，这里仅作文件名）
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
                    else:
                        st.error("❌ 处理失败，请检查文件格式是否符合要求。")

    # ======================== 分支 3：信用风险管理 ========================
    elif action == "信用风险管理":
        st.markdown("""
        <div class="info-box">
            <div class="info-title">🛡️ 信用风险管理</div>
            处理信用风险日报数据，生成各大区业务开展情况的可视化图表及 PDF 报告。
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader("📂 拖拽或点击上传 信用风险底表", type=["xlsx", "xls"], key="credit_upload")
        
        if st.button("🚀 生成图表与报告", key="btn_credit"):
            if uploaded_file is not None:
                with st.spinner("正在生成各大区分析报告及图表..."):
                    export_files = process_credit_report(uploaded_file)
                    if export_files:
                        st.success("✅ 信用风险报告生成完毕！")
                        st.markdown("### 📥 下载生成的文件")
                        cols = st.columns(4)
                        for idx, export_file in enumerate(export_files):
                            col_idx = idx % 4
                            with cols[col_idx]:
                                label = "📉 下载高清图" if export_file["type"] == "png" else "📊 下载 PDF"
                                mime = "image/png" if export_file["type"] == "png" else "application/pdf"
                                st.download_button(
                                    label=f"{label} ({export_file['name']})",
                                    data=export_file["data"],
                                    file_name=export_file["name"],
                                    mime=mime,
                                    use_container_width=True
                                )
                    else:
                        st.error("处理失败，未能提取到有效数据。")
            else:
                st.warning("⚠️ 请先上传 Excel 文件！")

    st.markdown("<div style='text-align:center; color:#ccc; margin-top:50px;'>© 2026 TakeItEasy Tool</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()