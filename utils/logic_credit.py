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
    
    regular_path = os.path.join(current_dir, 'msyh.ttc')
    if not os.path.exists(regular_path): regular_path = os.path.join(current_dir, 'msyh.ttf')
    custom_font_regular = FontProperties(fname=regular_path) if os.path.exists(regular_path) else None

    bold_path = os.path.join(current_dir, 'msyhbd.ttc')
    if not os.path.exists(bold_path): bold_path = os.path.join(current_dir, 'msyhbd.ttf')
    custom_font_bold = FontProperties(fname=bold_path) if os.path.exists(bold_path) else custom_font_regular

    range_str = range_str.replace('$', '')
    min_col, min_row, max_col, max_row = range_boundaries(range_str)

    # 1. 动态截断：物理切除底部的所有冗余空白与说明
    actual_max_row = max_row
    for r in range(min_row, max_row + 1):
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(min_col, max_col + 1)]
        combined = "".join(row_vals)
        if "是否填报" in combined or "填报说明" in combined:
            actual_max_row = r - 1 
            break
            
    while actual_max_row >= min_row:
        row_has_data = False
        row_vals = [str(ws.cell(row=actual_max_row, column=c).value or "").strip() for c in range(min_col, max_col + 1)]
        combined = "".join(row_vals)
        if "填报" in combined or "说明" in combined:
            actual_max_row -= 1
            continue
        if any(row_vals): 
            break
        actual_max_row -= 1

    # 2. 动态剥离无数据列
    valid_cols_set = set()
    for r in range(min_row, actual_max_row + 1):
        for c in range(min_col, max_col + 1):
            val = ws.cell(row=r, column=c).value
            if val is not None and str(val).strip() != "":
                valid_cols_set.add(c)
                for mr in ws.merged_cells.ranges:
                    if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
                        for mc in range(mr.min_col, mr.max_col + 1):
                            valid_cols_set.add(mc)
    valid_cols = sorted(list(valid_cols_set))
    if not valid_cols: return None

    # 3. 构建合并单元格字典
    merged_dict = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_col <= max_col and mr.max_col >= min_col and mr.min_row <= actual_max_row and mr.max_row >= min_row:
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    merged_dict[(r, c)] = {
                        'top_left': (mr.min_row, mr.min_col),
                        'bottom_right': (mr.max_row, mr.max_col)
                    }

    header_start_row = min_row
    for r in range(min_row, actual_max_row + 1):
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in valid_cols]
        combined = "".join(row_vals)
        
        if "序号" in combined or "业务单位" in combined or "大区" in combined:
            header_start_row = r
            break

    # 【增量优化点】文本提取逻辑解耦，即使都在同一行也能精准提取不丢失
    title_text = ""
    author_text = ""
    date_text = ""
    unit_text = ""
    
    for r in range(min_row, actual_max_row + 1):
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(min_col, max_col + 1)]
        combined = "".join(row_vals)
        
        # 提取标题
        if "汇总表" in combined or "监控表" in combined:
            if not title_text: title_text = next((v for v in row_vals if "表" in v), combined)
            
        # 遍历单元格，解耦提取并分别赋值（彻底解决丢失问题）
        for v in row_vals:
            if "制表单位" in v and not author_text:
                author_text = v
            if "截止时间" in v and not date_text:
                date_text = v
            if "单位" in v and ("万元" in v or "万" in v) and not unit_text:
                unit_text = v

    header_end_row = header_start_row
    for r in range(header_start_row, actual_max_row + 1):
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in valid_cols]
        combined = "".join(row_vals)
        if "沿江大区" in combined or "华东经营部" in combined or row_vals[0] == "1":
            header_end_row = r - 1
            break

    # 4. 精准标记跳过行，消除多余框框（即使是空行也会被消灭）
    row_types = {}
    row_heights = {}
    for r in range(min_row, actual_max_row + 1):
        combined = "".join([str(ws.cell(row=r, column=c).value or "").strip() for c in valid_cols])
        
        if r < header_start_row:
            row_types[r] = 'skip'
            row_heights[r] = 0.0
        elif "制表单位" in combined or "截止时间" in combined or (("单位" in combined and "万" in combined) and "合计" not in combined):
            row_types[r] = 'skip'
            row_heights[r] = 0.0
        else:
            row_types[r] = 'grid' 
            row_heights[r] = 3.2 if r <= header_end_row else 2.6

    # 5. 计算列宽
    col_widths = {c: 4.0 for c in valid_cols}
    for r in range(header_start_row, actual_max_row + 1):
        if row_types[r] == 'skip': continue
        for c in valid_cols:
            is_spanned = False
            if (r, c) in merged_dict:
                info = merged_dict[(r, c)]
                if info['bottom_right'][1] > info['top_left'][1]: 
                    is_spanned = True
            if is_spanned: continue 

            val = ws.cell(row=r, column=c).value
            if val:
                text_len = sum(1.8 if ord(ch) > 255 else 1.1 for ch in str(val))
                w = text_len * 0.9 + 1.5 
                if w > col_widths[c]: col_widths[c] = min(w, 25.0)

    # 极度压缩第一列（序号）的宽度
    col_widths[valid_cols[0]] = 2.5 

    # 双重死锁百分比逻辑
    def get_merged_cell_text(r, c):
        if (r, c) in merged_dict:
            tl_r, tl_c = merged_dict[(r, c)]['top_left']
            return str(ws.cell(row=tl_r, column=tl_c).value or "").strip()
        return str(ws.cell(row=r, column=c).value or "").strip()

    col_is_percent = {c: False for c in valid_cols}
    for c in valid_cols:
        col_header_full_text = ""
        for r in range(header_start_row, header_end_row + 1):
            col_header_full_text += get_merged_cell_text(r, c)
        
        col_header_clean = col_header_full_text.replace("\n", "").replace(" ", "")
        
        if "赊销余额/授信额度" in col_header_clean or "出库通知单/授信额度" in col_header_clean or "使用率" in col_header_clean:
            col_is_percent[c] = True

    # 6. 创建物理级绝对坐标画布
    W_grid = sum(col_widths.values())
    H_grid = sum(row_heights.values())

    A4_W, A4_H = 8.27, 11.69
    margin_x, margin_y = 0.4, 0.4
    max_w_in = A4_W - 2 * margin_x
    S = max_w_in / W_grid 
    
    top_space = 10.0
    bottom_space = 12.0
    H_total_virtual = top_space + H_grid + bottom_space
    H_in = H_total_virtual * S
    
    Final_H = max(A4_H, H_in + 1.0)
    
    fig = plt.figure(figsize=(A4_W, Final_H), dpi=800) 
    fig.patch.set_facecolor('white')

    ax = fig.add_axes([margin_x / A4_W, (Final_H - H_in - 0.4) / Final_H, max_w_in / A4_W, H_in / Final_H])
    ax.set_xlim(0, W_grid)
    ax.set_ylim(H_total_virtual, 0) # 翻转Y轴
    ax.axis('off')

    base_fs = 2.5 * S * 72 * 0.42 

    # ==================== 绘制外围提取图层 (完美落位) ====================
    # 顶部：大标题 (居中, 大字号, 必加粗)
    if title_text:
        prop_title = custom_font_bold.copy() if custom_font_bold else custom_font_regular
        if prop_title: prop_title.set_size(base_fs * 1.6)
        ax.text(W_grid / 2, 3.5, title_text, ha='center', va='center', fontproperties=prop_title, weight='bold', fontsize=base_fs*1.6)
    
    # 【增量优化点】顶部左侧：截止时间落位
    if date_text:
        ax.text(0.5, 8.0, date_text, ha='left', va='center', fontproperties=custom_font_regular, fontsize=base_fs*0.95)

    # 【增量优化点】顶部右侧：单位：万元 落位
    if unit_text:
        ax.text(W_grid - 0.5, 8.0, unit_text, ha='right', va='center', fontproperties=custom_font_regular, fontsize=base_fs*0.95)

    # ==================== 绘制核心表格层 ====================
    y_curr = top_space 
    
    for r in range(header_start_row, actual_max_row + 1):
        if row_types[r] == 'skip': continue
        
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
                    draw_w = sum(col_widths.get(mc, 0) for mc in range(info['top_left'][1], info['bottom_right'][1] + 1) if mc in valid_cols)
                    draw_h = sum(row_heights.get(mr_i, 2.2) for mr_i in range(info['top_left'][0], info['bottom_right'][0] + 1) if row_types.get(mr_i) != 'skip')

            if is_merged_top_left:
                cell = ws.cell(row=r, column=c)
                
                bg_color = '#FFFFFF'
                if cell.fill and cell.fill.patternType == 'solid' and cell.fill.start_color.rgb:
                    rgb = str(cell.fill.start_color.rgb)
                    if len(rgb) == 8 and rgb != '00000000': bg_color = '#' + rgb[2:]
                    elif len(rgb) == 6: bg_color = '#' + rgb
                
                is_header_row = (r <= header_end_row)
                if c == valid_cols[0] and not is_header_row:
                    bg_color = '#FFFFFF'

                rect = patches.Rectangle((x_curr, y_curr), draw_w, draw_h, facecolor=bg_color, edgecolor='#000000', linewidth=0.8)
                ax.add_patch(rect)
                
                val = cell.value
                fmt = cell.number_format or "General"
                text = ""
                
                if val is not None and str(val).strip() != "":
                    if isinstance(val, (int, float)):
                        # 双重死锁验证
                        if col_is_percent.get(c, False) and abs(val) <= 10:
                            if '.00' in fmt: text = f"{val:.2%}"
                            elif '.0' in fmt: text = f"{val:.1%}"
                            else: text = f"{val:.0%}"
                        else:
                            if ',' in fmt or (isinstance(val, (int, float)) and (val >= 1000 or val <= -1000)):
                                if isinstance(val, float) and not val.is_integer():
                                    text = f"{val:,.2f}".rstrip('0').rstrip('.')
                                else:
                                    text = f"{val:,.0f}"
                            else:
                                if isinstance(val, float): text = f"{val:.2f}".rstrip('0').rstrip('.')
                                else: text = str(val)
                    elif isinstance(val, datetime.datetime):
                        if "年" in fmt: text = val.strftime('%Y年%m月%d日')
                        else: text = val.strftime('%Y-%m-%d')
                    else:
                        text = str(val).strip()
                
                is_bold = False
                if cell.font and cell.font.bold: is_bold = True
                
                text_color = '#000000'
                if cell.font and cell.font.color and hasattr(cell.font.color, 'rgb') and cell.font.color.rgb:
                    rgb_val = str(cell.font.color.rgb)
                    if len(rgb_val) == 8 and rgb_val != '00000000': text_color = '#' + rgb_val[2:]
                    elif len(rgb_val) == 6: text_color = '#' + rgb_val
                
                # 【增量优化点】表格区内容绝对强制居中！
                halign, valign = 'center', 'center'
                text_x = x_curr + draw_w / 2
                text_y = y_curr + draw_h / 2
                
                if isinstance(text, str) and len(text) > (draw_w / 1.1):
                    wrap_w = max(1, int(draw_w / 1.1))
                    text = '\n'.join(textwrap.wrap(text, width=wrap_w))
                    
                if text:
                    kwargs = {
                        'ha': halign,
                        'va': valign,
                        'color': text_color,
                        'clip_on': True
                    }
                    
                    if is_bold and custom_font_bold:
                        prop = custom_font_bold.copy()
                        prop.set_size(base_fs)
                        kwargs['fontproperties'] = prop
                    elif custom_font_regular:
                        prop = custom_font_regular.copy()
                        if is_bold: prop.set_weight('bold')
                        prop.set_size(base_fs)
                        kwargs['fontproperties'] = prop
                    else:
                        kwargs['weight'] = 'bold' if is_bold else 'normal'
                        kwargs['fontsize'] = base_fs
                        
                    ax.text(text_x, text_y, text, **kwargs)

            x_curr += cw
        y_curr += rh

    # ==================== 绘制底部落款 ====================
    # 【增量优化点】制表单位独立成为表格正下方的优雅落款
    if author_text:
        ax.text(W_grid - 0.5, y_curr + 4.0, author_text, ha='right', va='center', fontproperties=custom_font_regular, fontsize=base_fs*0.95)

    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', dpi=800, facecolor=fig.get_facecolor(), edgecolor='none', bbox_inches='tight')
    plt.close(fig)
    img_stream.seek(0)
    return img_stream

# ==================== 导出文件生成逻辑 ====================

def generate_export_files_in_memory(file_stream):
    results = []
    logs = []
    today_mmdd = datetime.datetime.now().strftime('%m%d')
    sys_name = platform.system()
    
    sheets_info = [
        {"name": "每日-中粮贸易外部赊销限额使用监控表", "range": "$A$1:$G$30", "base_title": "中粮贸易外部赊销限额使用监控表"}
    ]
    
    if True: 
        sheets_info.append({"name": "每周-正大额度使用情况", "range": "$A$1:$L$34", "base_title": "正大额度使用情况"})
        
    if sys_name == 'Windows':
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
            file_stream.seek(0)
            tmp_in.write(file_stream.read())
            temp_excel_path = tmp_in.name
            
        try:
            import win32com.client
            try:
                app = win32com.client.Dispatch("Excel.Application")
            except:
                app = win32com.client.Dispatch("Ket.Application") 
                
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(temp_excel_path, ReadOnly=True)
            
            for s_info in sheets_info:
                try:
                    ws = wb.Sheets(s_info['name'])
                    ws.PageSetup.PrintArea = s_info['range']
                    ws.PageSetup.Orientation = 1
                    ws.PageSetup.PaperSize = 9
                    ws.PageSetup.Zoom = False
                    ws.PageSetup.FitToPagesWide = 1
                    ws.PageSetup.FitToPagesTall = 1
                    ws.PageSetup.CenterHorizontally = True
                    
                    out_name = f"{s_info['base_title']}{today_mmdd}.pdf"
                    temp_pdf_path = os.path.join(tempfile.gettempdir(), out_name)
                    
                    ws.ExportAsFixedFormat(0, temp_pdf_path, IgnorePrintAreas=False)
                    
                    with open(temp_pdf_path, "rb") as f:
                        results.append({"name": out_name, "data": f.read(), "type": "pdf"})
                    os.remove(temp_pdf_path)
                    logs.append(f"   ✅ 成功生成 PDF: {out_name}")
                except Exception as e:
                    logs.append(f"   ⚠️ 跳过 {s_info['name']}: {str(e)}")
                    
            wb.Close(SaveChanges=False)
            app.Quit()
        except Exception as e:
            logs.append(f"❌ PDF 引擎调用失败: {str(e)}")
        finally:
            if os.path.exists(temp_excel_path):
                os.remove(temp_excel_path)
                
    else:
        file_stream.seek(0)
        try:
            wb = openpyxl.load_workbook(file_stream, data_only=True)
            for s_info in sheets_info:
                if s_info['name'] in wb.sheetnames:
                    img_stream = render_sheet_range_to_image_stream(wb[s_info['name']], s_info['range'])
                    if img_stream:
                        out_name = f"{s_info['base_title']}{today_mmdd}.png"
                        results.append({"name": out_name, "data": img_stream.read(), "type": "png"})
                        logs.append(f"   ✅ 成功生成像素级对齐图片: {out_name}")
        except Exception as e:
            logs.append(f"❌ 跨平台渲染引擎出错: {str(e)}")
            
    return results, logs

# ==================== 主控入口 ====================

def process_credit_report(uploaded_file):
    logs = []
    sys_name = platform.system()
    env_msg = f"当前环境: {sys_name} " + ("(原生支持 PDF 导出)" if sys_name == 'Windows' else "(云端环境，将生成高清预览图替代 PDF)")
    
    kill_excel_processes()
    file_stream = io.BytesIO(uploaded_file.getvalue())
    
    word_bytes, word_text_dict, word_logs = generate_word_in_memory(file_stream)
    logs.extend(word_logs)
    
    export_files, export_logs = generate_export_files_in_memory(file_stream)
    logs.extend(export_logs)
    
    kill_excel_processes()
    
    return word_bytes, word_text_dict, export_files, logs, env_msg
