import os
import io
import datetime
from datetime import timedelta
import tempfile
import platform
import warnings
import textwrap
import openpyxl
from openpyxl.utils import range_boundaries, get_column_letter

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 尝试导入绘图库作为 Linux 环境的降级图片生成方案
try:
    import matplotlib.pyplot as plt
    import matplotlib as mpl
    import matplotlib.patches as patches
    from matplotlib.font_manager import FontProperties
    mpl.rcParams['axes.unicode_minus'] = False
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

warnings.simplefilter("ignore", category=UserWarning)

# ==================== 基础辅助函数 ====================

def kill_excel_processes():
    if platform.system() == "Windows":
        try:
            os.system("taskkill /f /im excel.exe >nul 2>&1")
            os.system("taskkill /f /im et.exe >nul 2>&1")
        except:
            pass

def clean_value(val):
    if val is None: return ""
    return str(val).strip()

def clean_money(val):
    if val is None: return 0.0
    try:
        if isinstance(val, str):
            val = val.replace(',', '').strip()
            if not val: return 0.0
        return float(val)
    except:
        return 0.0

def get_cell_fill_color(cell):
    if cell.fill and cell.fill.start_color:
        color = cell.fill.start_color
        if not color.index or color.index == '00000000':
             return False
        return True
    return False

# ==================== Word 生成逻辑 (保持不变) ====================

def set_font_style(run, font_name='宋体', size=12, bold=False):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.font.bold = bold

def generate_word_in_memory(file_stream):
    logs = []
    report_text_dict = {} 
    try:
        file_stream.seek(0)
        wb = openpyxl.load_workbook(file_stream, data_only=True)
    except Exception as e:
        return None, {}, [f"❌ 读取 Excel 文件失败: {e}"]

    target_sheet_name = "每日-各品种线战略客户逾期通报"
    if target_sheet_name not in wb.sheetnames:
        return None, {}, [f"❌ 未找到关键工作表: {target_sheet_name}"]
    
    ws = wb[target_sheet_name]
    header_row_idx = None
    col_map = {}
    required_cols = ["品种线", "大区", "经营部", "客户名称", "合同号", "品种", "逾期天数", "逾期金额"]
    for r in range(1, 21):
        row_values = [clean_value(c.value) for c in ws[r]]
        matches = sum(1 for k in required_cols if any(k in v for v in row_values))
        if matches >= 4:
            header_row_idx = r
            for idx, val in enumerate(row_values):
                for k in required_cols:
                    if k in val: col_map[k] = idx
            break
            
    if not header_row_idx:
        return None, {}, ["❌ 未找到标题行。"]

    scope_ranges = {}
    target_keywords = ["玉米", "粮谷", "大豆"] 
    p_col_idx = col_map.get("品种线")
    for rng in ws.merged_cells.ranges:
        if rng.min_col <= (p_col_idx + 1) <= rng.max_col:
            top_val = clean_value(ws.cell(row=rng.min_row, column=rng.min_col).value)
            for kw in target_keywords:
                if kw in top_val:
                    start_r = max(rng.min_row, header_row_idx + 1)
                    end_r = rng.max_row
                    if start_r <= end_r:
                        scope_ranges[kw] = range(start_r, end_r + 1)
                    break

    data_store = {}
    exclude_regions = ["东北大区", "内陆大区", "沿江大区", "沿海大区", "东北", "内陆", "沿江", "沿海"]
    for kw, row_range in scope_ranges.items():
        if kw not in data_store: data_store[kw] = {}
        current_group = None
        if kw == "大豆":
            current_group = "ALL_SOYBEAN"
            data_store[kw][current_group] = []
        for r_idx in row_range:
            row = ws[r_idx]
            def get_val(key):
                idx = col_map.get(key)
                return row[idx].value if idx is not None else None

            region_str = clean_value(get_val("大区"))
            client_name = clean_value(get_val("客户名称"))
            is_group = False
            if kw != "大豆":
                cell_region = row[col_map.get("大区")]
                if region_str and (region_str not in exclude_regions) and get_cell_fill_color(cell_region):
                    is_group = True
            if is_group:
                current_group = region_str
                if current_group not in data_store[kw]:
                    data_store[kw][current_group] = []
            elif current_group:
                money_val = clean_money(get_val("逾期金额"))
                if client_name and money_val > 0:
                    data_store[kw][current_group].append({
                        '大区': region_str, '经营部': clean_value(get_val("经营部")),
                        '客户名称': client_name, '合同号': clean_value(get_val("合同号")),
                        '品种': clean_value(get_val("品种")),
                        '逾期天数': clean_value(get_val("逾期天数")).replace('.0', ''),
                        '逾期金额': money_val
                    })

    doc = Document()
    yesterday = datetime.datetime.now() - timedelta(days=1)
    date_str = f"{yesterday.year}年{yesterday.month}月{yesterday.day}日"
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font_style(p_title.add_run("信用风险管理日报"), font_name='黑体', size=16, bold=True)
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_font_style(p_sub.add_run(f"截至日期：{date_str}"), font_name='楷体', size=12)
    doc.add_paragraph()

    chinese_nums = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
    has_content = False
    for key in target_keywords:
        if key not in data_store or not data_store[key]: continue
        groups = data_store[key]
        total_count = 0
        total_money = 0
        valid_groups_list = []
        all_rows_flat = [] 
        for g_name, rows in groups.items():
            if len(rows) > 0:
                g_total = sum(r['逾期金额'] for r in rows)
                valid_groups_list.append({'name': g_name, 'rows': rows, 'total': g_total})
                all_rows_flat.extend(rows)
                total_count += len(rows)
                total_money += g_total
        if total_count == 0: continue
        has_content = True
        
        header_text = (f"按照公司领导要求，请{key}中心充分发挥与战略客户的良好沟通机制，"
                       f"协助大区督促客户及时回款，压降逾期赊销。截至{date_str}，"
                       f"{key}中心战略客户，共计逾期{total_count}笔，逾期金额{int(total_money)}万元。")
        p_h = doc.add_paragraph()
        set_font_style(p_h.add_run(f"【{key}中心】"), font_name='黑体', size=14, bold=True)
        p_b = doc.add_paragraph()
        p_b.paragraph_format.first_line_indent = Pt(24)
        set_font_style(p_b.add_run(header_text), font_name='宋体', size=12)
        center_text_block = f"【{key}中心】\n{header_text}\n"

        if key == "大豆":
            all_rows_flat.sort(key=lambda r: r['逾期金额'], reverse=True)
            for j, row in enumerate(all_rows_flat):
                line = (f"{j+1}、{row['大区']}，{row['经营部']}，{row['客户名称']}，"
                        f"{row['合同号']}，{row['品种']}，逾期{row['逾期天数']}天，"
                        f"{int(row['逾期金额'])}万元；")
                p = doc.add_paragraph()
                p.paragraph_format.first_line_indent = Pt(24)
                set_font_style(p.add_run(line), font_name='宋体', size=12)
                center_text_block += f"{line}\n"
        else:
            valid_groups_list.sort(key=lambda x: x['total'], reverse=True)
            for i, g_data in enumerate(valid_groups_list):
                idx_str = chinese_nums[i] if i < len(chinese_nums) else str(i+1)
                group_line = (f"{idx_str}、{g_data['name']}，共计逾期{len(g_data['rows'])}笔，"
                              f"逾期金额{int(g_data['total'])}万元。")
                p = doc.add_paragraph()
                p.paragraph_format.first_line_indent = Pt(24)
                set_font_style(p.add_run(group_line), font_name='黑体', size=12, bold=True)
                center_text_block += f"{group_line}\n"
                
                sorted_rows = sorted(g_data['rows'], key=lambda r: r['逾期金额'], reverse=True)
                for j, row in enumerate(sorted_rows):
                    line = (f"{j+1}、{row['大区']}，{row['经营部']}，{row['客户名称']}，"
                            f"{row['合同号']}，{row['品种']}，逾期{row['逾期天数']}天，"
                            f"{int(row['逾期金额'])}万元；")
                    p = doc.add_paragraph()
                    p.paragraph_format.first_line_indent = Pt(24)
                    set_font_style(p.add_run(line), font_name='宋体', size=12)
                    center_text_block += f"{line}\n"
        doc.add_paragraph()
        report_text_dict[key] = center_text_block

    if not has_content:
        return None, {}, ["⚠️ 未提取到逾期数据，无 Word 报告生成。"]

    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    logs.append("✅ Word 报告内存生成成功！")
    return out_stream, report_text_dict, logs


# ==================== 终极防蜷缩：100%纯物理镜像渲染引擎 ====================

def render_sheet_range_to_image_stream(ws, range_str):
    if not MATPLOTLIB_AVAILABLE:
        return None

    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 严格挂载字体库
    regular_path = os.path.join(current_dir, 'msyh.ttc')
    if not os.path.exists(regular_path): regular_path = os.path.join(current_dir, 'msyh.ttf')
    custom_font_regular = FontProperties(fname=regular_path) if os.path.exists(regular_path) else None

    bold_path = os.path.join(current_dir, 'msyhbd.ttc')
    if not os.path.exists(bold_path): bold_path = os.path.join(current_dir, 'msyhbd.ttf')
    custom_font_bold = FontProperties(fname=bold_path) if os.path.exists(bold_path) else custom_font_regular

    # 1. 初始框选范围
    range_str = range_str.replace('$', '')
    min_col, min_row, max_col, max_row = range_boundaries(range_str)

    # 2. 动态截断：切除底部“是否填报”
    actual_max_row = max_row
    for r in range(min_row, max_row + 1):
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(min_col, max_col + 1)]
        combined = "".join(row_vals)
        if "是否填报" in combined or "填报说明" in combined:
            actual_max_row = r - 1 
            break
            
    # 【修复 2】倒推切除所有的纯空白行，彻底消灭底部多余空框
    while actual_max_row >= min_row:
        row_has_data = False
        for c in range(min_col, max_col + 1):
            val = ws.cell(row=actual_max_row, column=c).value
            if val is not None and str(val).strip() != "":
                row_has_data = True
                break
        if row_has_data:
            break
        actual_max_row -= 1

    # 3. 动态去除完全空白的冗余列
    valid_cols_set = set()
    for r in range(min_row, actual_max_row + 1):
        for c in range(min_col, max_col + 1):
            val = ws.cell(row=r, column=c).value
            if val is not None and str(val).strip() != "":
                valid_cols_set.add(c)
                # 兼容合并单元格覆盖的列
                for mr in ws.merged_cells.ranges:
                    if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
                        for mc in range(mr.min_col, mr.max_col + 1):
                            valid_cols_set.add(mc)
    valid_cols = sorted(list(valid_cols_set))
    if not valid_cols: return None

    # 4. 映射合并单元格
    merged_dict = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_col <= max_col and mr.max_col >= min_col and mr.min_row <= actual_max_row and mr.max_row >= min_row:
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    merged_dict[(r, c)] = {
                        'top_left': (mr.min_row, mr.min_col),
                        'bottom_right': (mr.max_row, mr.max_col)
                    }

    # 【修复 1】精准定位特定的百分比列，杜绝 459100% 现象
    col_is_percent = {c: False for c in valid_cols}
    for c in valid_cols:
        col_header_text = ""
        for r in range(min_row, min_row + 6):
            # 获取该列头部的所有文字内容，包含被合并的表头
            check_r, check_c = r, c
            for mr in ws.merged_cells.ranges:
                if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
                    check_r, check_c = mr.min_row, mr.min_col
                    break
            col_header_text += str(ws.cell(row=check_r, column=check_c).value or "")
            
        # 只要这列的表头包含这几个词，才允许转百分比
        if "赊销余额/授信额度" in col_header_text or "出库通知单/授信额度" in col_header_text or "使用率" in col_header_text:
            col_is_percent[c] = True

    # 5. 计算物理最佳行列比例
    col_widths = {c: 4.0 for c in valid_cols}
    row_heights = {r: 2.5 for r in range(min_row, actual_max_row + 1)}
    
    for r in range(min_row, actual_max_row + 1):
        for c in valid_cols:
            is_spanned = False
            for mr in ws.merged_cells.ranges:
                if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
                    if mr.max_col > mr.min_col: is_spanned = True
            if is_spanned: continue 
            val = ws.cell(row=r, column=c).value
            if val:
                text_len = sum(1.8 if ord(ch) > 255 else 1.1 for ch in str(val))
                w = text_len * 0.9 + 1.8 
                if w > col_widths[c]: col_widths[c] = min(w, 25.0)
    col_widths[valid_cols[0]] = max(3.5, col_widths[valid_cols[0]])

    # 6. 建立【DPI 1000】究极超清坐标画板
    W_grid = sum(col_widths.values())
    H_grid = sum(row_heights.values())
    A4_W, A4_H = 8.27, 11.69
    margin_x, margin_y = 0.4, 0.4
    max_w_in = A4_W - 2 * margin_x
    S = max_w_in / W_grid 
    H_in = H_grid * S
    
    Final_H = max(A4_H, H_in + 1.0)
    fig = plt.figure(figsize=(A4_W, Final_H), dpi=1000) # 【修复 4】告别模糊
    fig.patch.set_facecolor('white')

    ax = fig.add_axes([margin_x / A4_W, (Final_H - H_in - margin_y) / Final_H, max_w_in / A4_W, H_in / Final_H])
    ax.set_xlim(0, W_grid)
    ax.set_ylim(H_grid, 0)
    ax.axis('off')

    base_fs = 2.5 * S * 72 * 0.45 

    # 7. 逐像素网格镜像渲染
    y_curr = 0
    for r in range(min_row, actual_max_row + 1):
        x_curr = 0
        rh = row_heights[r]
        
        for c in valid_cols:
            cw = col_widths[c]
            
            is_merged_top_left = True
            draw_w, draw_h = cw, rh
            
            if (r, c) in merged_dict:
                info = merged_dict[(r, c)]
                if (r, c) != info['top_left']:
                    is_merged_top_left = False 
                else:
                    draw_w = sum(col_widths.get(mc, 4.0) for mc in range(info['top_left'][1], info['bottom_right'][1] + 1) if mc in valid_cols)
                    draw_h = sum(row_heights.get(mr, 2.5) for mr in range(info['top_left'][0], info['bottom_right'][0] + 1))

            if is_merged_top_left:
                cell = ws.cell(row=r, column=c)
                val_str = str(cell.value or "").strip()
                
                # --- 核心机制：识别独立排版的无边框单元格 (标题/落款) ---
                is_title_meta = False
                if "汇总表" in val_str or "监控表" in val_str:
                    is_title_meta = True
                elif "制表单位" in val_str or "截止时间" in val_str or ("单位" in val_str and "万" in val_str):
                    is_title_meta = True

                # 底色与边框提取
                bg_color = '#FFFFFF'
                if not is_title_meta:
                    if cell.fill and cell.fill.patternType == 'solid' and cell.fill.start_color.rgb:
                        rgb = str(cell.fill.start_color.rgb)
                        if len(rgb) == 8 and rgb != '00000000': bg_color = '#' + rgb[2:]
                        elif len(rgb) == 6: bg_color = '#' + rgb
                
                # 强制要求：除表头外，序号列禁止染底色
                if c == valid_cols[0] and r > min_row + 3 and not is_title_meta:
                    if str(cell.value).isdigit() or str(cell.value).strip() == "":
                        bg_color = '#FFFFFF'

                lw = 0.0 if is_title_meta else 0.8
                rect = patches.Rectangle((x_curr, y_curr), draw_w, draw_h, facecolor=bg_color, edgecolor='#000000', linewidth=lw)
                ax.add_patch(rect)
                
                # --- 核心机制：数值与格式的究极还原 ---
                val = cell.value
                fmt = cell.number_format or "General"
                text = ""
                
                if val is not None and str(val).strip() != "":
                    if isinstance(val, (int, float)):
                        # 【修复 1】双保险锁死：只有在指定列，且数值在合理范围内(<=10)，才赋予 %
                        if ('%' in fmt) or (col_is_percent.get(c, False) and abs(val) <= 10):
                            if '.00' in fmt: text = f"{val:.2%}"
                            elif '.0' in fmt: text = f"{val:.1%
