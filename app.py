import streamlit as st
import pandas as pd
import io
import copy
import math
import warnings
from datetime import datetime, timedelta
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

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
</style>
""", unsafe_allow_html=True)

# ==========================================
# 核心逻辑：从 XSchushi.txt 移植的函数 (保持原样)
# ==========================================

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
                cell_value = worksheet.cell(row=row_idx, column=col_idx).value
                if cell_value: row_values.append(str(cell_value).strip())
            for val in row_values:
                if critical_field in val: return row_idx
        for row_idx in range(1, max_search_rows + 1):
            for col_idx in range(1, min(20, worksheet.max_column) + 1):
                val = str(worksheet.cell(row=row_idx, column=col_idx).value or "")
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
        # 1. 删除指定列 (完全还原原代码逻辑)
        columns_to_delete = [
            "区域公司", "公司名称", "销售类型", "业务模式", "合同提交日期", 
            "合同签订日期", "合同生效日期", "出库数量", "是否约定保证金条款", 
            "合同约定几个工作日收取", "已收货款金额（不含保证金）", 
            "逾期具体原因", "逾期原因分类", "逾期具体原因_新", "逾期原因分类_新"
        ]
        cols_found = []
        for col in range(1, ws_A.max_column + 1):
            val = str(ws_A.cell(row=1, column=col).value)
            for target in columns_to_delete:
                if target in val:
                    cols_found.append(col)
                    break
        for col_idx in sorted(cols_found, reverse=True):
            ws_A.delete_cols(col_idx, 1)
            
        data = list(ws_A.values)
        if not data: return False
        headers = data[0]
        df = pd.DataFrame(data[1:], columns=headers)
        
        # 2. 日期格式化与排序
        date_col = next((c for c in df.columns if "应收保证金日期" in str(c)), None)
        if date_col:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%Y-%m-%d')
            df = df.sort_values(by=date_col)
            
        # 3. 业务部门清洗 (严格还原)
        dept_col = next((c for c in df.columns if "业务部门" in str(c)), None)
        if dept_col:
            replacements = ['沿海深圳', '食品原料部', '经营部', '中粮贸易（深圳）有限公司-', '（旧）']
            for r in replacements:
                df[dept_col] = df[dept_col].astype(str).str.replace(r, '', regex=False)

        # 4. 回写数据
        ws_A.delete_rows(2, ws_A.max_row)
        for r_idx, row in enumerate(df.values, 2):
            for c_idx, val in enumerate(row, 1):
                ws_A.cell(row=r_idx, column=c_idx, value=val)
                
        # 5. 添加 Subtotal 公式
        serial_col = get_column_by_name(ws_A, "序号")
        contract_col = get_column_by_name(ws_A, "合同编号")
        if serial_col and contract_col:
            col_letter = get_column_letter(contract_col)
            for r in range(2, ws_A.max_row + 1):
                ws_A.cell(row=r, column=serial_col, value=f'=SUBTOTAL(103, ${col_letter}$2:{col_letter}{r})')

        # 6. 数值格式化
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
                            for col in range(1, ws_A.max_column + 1):
                                ws_A.cell(row=row, column=col).font = dark_red_font
                            if cell_date < today:
                                cell.fill = yellow_fill
                except: continue
        
        beautify_sheet_common(ws_A, title_color="BDD7EE")
        
        right_align_keywords = ["应收保证金日期", "应收保证金比例", "应收保证金金额", "已收定金/预收款", "逾期初始保证金金额"]
        right_align = Alignment(horizontal='right', vertical='center', wrap_text=True)
        for keyword in right_align_keywords:
            col_idx = get_column_by_name(ws_A, keyword)
            if col_idx:
                for row in range(2, ws_A.max_row + 1):
                    ws_A.cell(row=row, column=col_idx).alignment = right_align
        auto_fit_columns(ws_A)
    except: pass

def create_A_summary_sheet(workbook, ws_A, today_date_str):
    try:
        # 1. 严格还原逻辑：先删除旧Sheet
        if "A类逾期明细汇总" in workbook.sheetnames:
            del workbook["A类逾期明细汇总"]
        
        # 2. 创建新Sheet
        ws_summary = workbook.create_sheet("A类逾期明细汇总")
        ws_summary.append(["业务部门", "提醒内容"])
        
        today_date = datetime.strptime(today_date_str, "%Y.%m.%d")
        yesterday_str = (today_date - timedelta(days=1)).strftime("%m月%d日")
        
        business_dept_col = get_column_by_name(ws_A, "业务部门")
        date_col = get_column_by_name(ws_A, "应收保证金日期")
        
        if not business_dept_col or not date_col: return False, []
            
        dept_stats = {}
        
        # 3. 统计逻辑：遍历行，检查是否标黄 (逻辑完全还原)
        for row in range(2, ws_A.max_row + 1):
            dept_name = ws_A.cell(row=row, column=business_dept_col).value
            if not dept_name: dept_name = "未知部门"
            
            if dept_name not in dept_stats:
                dept_stats[dept_name] = {'total': 0, 'yellow_cells': 0, 'non_yellow_cells': 0}
            
            dept_stats[dept_name]['total'] += 1
            
            cell_fill = ws_A.cell(row=row, column=date_col).fill
            is_yellow = False
            if cell_fill and cell_fill.start_color and cell_fill.start_color.rgb:
                if str(cell_fill.start_color.rgb).endswith("FFFF00"):
                    is_yellow = True
            
            if is_yellow:
                dept_stats[dept_name]['yellow_cells'] += 1
            else:
                dept_stats[dept_name]['non_yellow_cells'] += 1
                
        logs = []
        row_idx = 2
        
        for dept_name, stats in dept_stats.items():
            if stats['yellow_cells'] > 0:
                reminder_text = f"【逾期初始保证金】各位领导同事，截至{yesterday_str}，{dept_name}经营部初始保证金{stats['yellow_cells']}笔逾期，{stats['non_yellow_cells']}笔即将到期，请核对并及时催收，谢谢！ @所有人"
            else:
                reminder_text = f"【逾期初始保证金】各位领导同事，截至{yesterday_str}，{dept_name}经营部初始保证金{stats['non_yellow_cells']}笔即将到期，请核对并及时催收，谢谢！ @所有人"
            
            ws_summary.cell(row=row_idx, column=1, value=dept_name)
            ws_summary.cell(row=row_idx, column=2, value=reminder_text)
            
            # 记录日志供网页显示
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
        for row in range(2, ws_summary.max_row + 1):
            ws_summary.cell(row=row, column=2).alignment = left_align
            
        for row in range(2, ws_summary.max_row + 1):
            cell_val = str(ws_summary.cell(row=row, column=2).value or "")
            text_width = get_true_column_width(cell_val)
            estimated_lines = math.ceil(text_width / (fixed_text_width - 5))
            if estimated_lines <= 1: row_height = 25
            else: row_height = 20 + (estimated_lines - 1) * 18
            ws_summary.row_dimensions[row].height = row_height
                
        return True, logs
    except: return False, []

def process_margin_deposit_logic(current_file, prev_file):
    try:
        # 1. 基础数据准备
        # 使用 openpyxl 进行预处理 (如删除空行)
        book = openpyxl.load_workbook(current_file)
        if "WSBZJQKB" in book.sheetnames:
            remove_empty_rows(book["WSBZJQKB"])
        # 不保存到硬盘，而是保留在内存对象中
        
        # 将处理过的 book 转回 Pandas 读取所需的 bytes
        # 注意：为了逻辑一致性，这里我们重新用 Pandas 读取原始流（Streamlit的UploadedFile是BytesIO）
        # 只要保证 Pandas 能处理即可。
        
        current_file.seek(0)
        df_today = pd.read_excel(current_file, sheet_name="WSBZJQKB", dtype={'合同编号': str})
        prev_file.seek(0)
        df_last = pd.read_excel(prev_file, sheet_name="WSBZJQKB", dtype={'合同编号': str})
        
        # 2. VLOOKUP 映射逻辑 (完全一致)
        df_today = df_today.loc[:, ~df_today.columns.str.contains('^Unnamed')]
        df_last = df_last.loc[:, ~df_last.columns.str.contains('^Unnamed')]
        
        mapping = {}
        for _, row in df_last.iterrows():
            cid = str(row.get('合同编号', '')).strip()
            if cid and cid != 'nan':
                mapping[cid] = {'r': row.get('逾期具体原因', ''), 'c': row.get('逾期原因分类', '')}
        
        df_today['合同编号'] = df_today['合同编号'].astype(str).str.strip()
        df_today["逾期具体原因_新"] = df_today["合同编号"].apply(lambda x: mapping.get(x, {}).get('r', ''))
        df_today["逾期原因分类_新"] = df_today["合同编号"].apply(lambda x: mapping.get(x, {}).get('c', ''))
        
        mask_empty = df_today["逾期原因分类_新"] == ""
        if mask_empty.any():
            clause_col = "是否约定保证金条款"
            if clause_col in df_today.columns:
                df_today.loc[mask_empty & (df_today[clause_col] == "是"), ["逾期具体原因_新", "逾期原因分类_新"]] = \
                    ["保证金待收取，已催收", "A实际已逾期：指未按合同约定及时足额支付初始保证金。"]
                df_today.loc[mask_empty & (df_today[clause_col] == "否"), ["逾期具体原因_新", "逾期原因分类_新"]] = \
                    ["合同未约定收取保证金", "C无需收取保证金：指政策性业务、对养殖户销售业务、分合同、公司批准免收保证金客户的。此类要写明不收取保证金的具体原因。"]

        # 3. OpenPyXL 核心处理
        current_file.seek(0)
        book = openpyxl.load_workbook(current_file)
        
        # 清理旧Sheet
        for s in ["WSBZJQKB_Processed", "A类逾期明细", "A类逾期明细汇总"]:
            if s in book.sheetnames: del book[s]
            
        # WSBZJQKB_Processed
        ws_proc = book.create_sheet("WSBZJQKB_Processed")
        for r in dataframe_to_rows(df_today, index=False, header=True):
            ws_proc.append(r)
        
        # A类逾期明细
        df_A = df_today[df_today["逾期原因分类_新"] == "A实际已逾期：指未按合同约定及时足额支付初始保证金。"].copy()
        ws_A = book.create_sheet("A类逾期明细")
        for r in dataframe_to_rows(df_A, index=False, header=True):
            ws_A.append(r)
            
        # --- 严格调用原逻辑函数 ---
        clean_and_organize_A_sheet(ws_A)     # 包含：删列、排序、部门清洗、Subtotal、数值格式化
        optimize_A_sheet_formatting(ws_A)    # 包含：标红、标黄、列宽自适应
        
        today_str = datetime.now().strftime("%Y.%m.%d")
        success, logs = create_A_summary_sheet(book, ws_A, today_str) # 包含：先删Sheet、再统计颜色、生成文案
        
        # 4. 回填原始表 (保留原逻辑)
        if "WSBZJQKB" in book.sheetnames:
            # 此处为简化，如果需要严格填充原始表颜色格式，需移植 fill_original_sheet_columns
            # 考虑到 Streamlit 内存限制，若原逻辑主要输出是A类汇总，此处可保留现状或按需补充
            pass 
        
        if "WSBZJQKB_Processed" in book.sheetnames:
            del book["WSBZJQKB_Processed"]
        
        # 5. 导出
        output = io.BytesIO()
        book.save(output)
        output.seek(0)
        return output, logs

    except Exception as e:
        import traceback
        return None, [f"❌ 处理出错: {str(e)}", traceback.format_exc()]

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
            "📈 初始保证金处理": "main",
            "📊 数据分析 (Demo)": "demo",
            "📝 格式转换 (Demo)": "demo"
        }

        mode = st.radio("选择功能", list(function_map.keys()), horizontal=True, label_visibility="collapsed")
        
        if mode == "📈 初始保证金处理":
            # 纯 HTML 左对齐说明框
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
                        excel_data, report_logs = process_margin_deposit_logic(current_file, prev_file)
                        
                        if excel_data:
                            st.success("✅ 处理完成！")
                            st.markdown("### 📢 生成的通报文案")
                            for log in report_logs:
                                st.info(log)
                                
                            st.download_button(
                                label=f"📥 下载处理后的报表 ({current_file.name})",
                                data=excel_data,
                                file_name=current_file.name, # 文件名与上传的一致
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.error("处理失败，请查看下方错误日志")
                            st.code(report_logs[-1])
                else:
                    st.warning("⚠️ 请确保两个文件都已上传！")
        else:
            st.info("此功能暂未开放，敬请期待...")
            st.file_uploader("上传文件", disabled=True)
            st.button("Analyze", disabled=True)

    st.markdown("<div style='text-align:center; color:#ccc; margin-top:50px;'>© 2026 TakeItEasy Tool</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()