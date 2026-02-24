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
    """仅在 Windows 环境下强制关闭后台滞留的 Excel/WPS 进程"""
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

# ==================== Word 生成逻辑 (内存级操作) ====================

def set_font_style(run, font_name='宋体', size=12, bold=False):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.font.bold = bold

def generate_word_in_memory(file_stream):
    """内存级生成 Word 报告，返回 Docx 字节流和分类别的字典用于页面展示"""
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
    
    # 1. 定位标题行
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

    # 2. 划定范围
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

    # 3. 数据采集
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

    # 4. 生成 Word Document
    doc = Document()
    yesterday = datetime.datetime.now() - timedelta(days=1)
    date_str = f"{yesterday.year}年{yesterday.month}月{yesterday.day}日"
    
    # 构建 Document
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
        
        # Word 写入
        p_h = doc.add_paragraph()
        set_font_style(p_h.add_run(f"【{key}中心】"), font_name='黑体', size=14, bold=True)
        p_b = doc.add_paragraph()
        p_b.paragraph_format.first_line_indent = Pt(24)
        set_font_style(p_b.add_run(header_text), font_name='宋体', size=12)
        
        # 网页展示收集 (保持干净)
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
                
                # 【修复项】：删除了硬编码的 "分大区情况如下："，还原清爽文本
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

# ==================== 终极优化：1:1 完美还原的图片渲染引擎 ====================

def render_sheet_range_to_image_stream(ws, range_str):
    """
    基于 patches 自定义绘图的终极 1:1 还原引擎
    新增支持独立的粗体文件加载，并强制定制大区行、合计行的背景底色。
    """
    if not MATPLOTLIB_AVAILABLE:
        return None

    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 1. 尝试加载常规字体
    regular_path = os.path.join(current_dir, 'msyh.ttc')
    if not os.path.exists(regular_path): regular_path = os.path.join(current_dir, 'msyh.ttf')
    custom_font_regular = FontProperties(fname=regular_path) if os.path.exists(regular_path) else None

    # 2. 尝试加载真正的【粗体字体】(终极解决字体无法加粗的核心)
    bold_path = os.path.join(current_dir, 'msyhbd.ttc')
    if not os.path.exists(bold_path): bold_path = os.path.join(current_dir, 'msyhbd.ttf')
    custom_font_bold = FontProperties(fname=bold_path) if os.path.exists(bold_path) else custom_font_regular

    # --- 1. 数据范围精准框选 ---
    range_str = range_str.replace('$', '')
    min_col, min_row, max_col, max_row = range_boundaries(range_str)

    # --- 2. 映射所有合并单元格 ---
    merged_dict = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_col <= max_col and mr.max_col >= min_col and mr.min_row <= max_row and mr.max_row >= min_row:
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    merged_dict[(r, c)] = {
                        'top_left': (mr.min_row, mr.min_col),
                        'bottom_right': (mr.max_row, mr.max_col)
                    }

    # --- 3. 提取每一行的“类型”，用于精准控制底色与粗细 ---
    row_types = {}
    for r in range(min_row, max_row + 1):
        t1 = str(ws.cell(row=r, column=min_col).value or "").strip()
        t2 = str(ws.cell(row=r, column=min_col+1).value or "").strip()
        combined = t1 + t2
        
        if r == min_row:
            row_types[r] = 'title'
        elif "单位" in combined and "万元" in combined:
            row_types[r] = 'unit'
        elif t1 == "序号" or t2 == "业务单位" or "赊销余额" in combined:
            row_types[r] = 'header'
        elif "合计" in combined:
            row_types[r] = 'total'
        elif "大区" in t1 or "大区" in t2:
            row_types[r] = 'region'
        else:
            row_types[r] = 'normal'

    # --- 4. 动态「自适应」计算行列宽度 ---
    col_widths = {c: 3.0 for c in range(min_col, max_col + 1)}
    row_heights = {r: 1.2 for r in range(min_row, max_row + 1)}
    
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            if (r, c) in merged_dict and merged_dict[(r, c)]['top_left'] != (r, c):
                continue
            val = ws.cell(row=r, column=c).value
            if val is not None:
                text_len = sum(2.0 if ord(ch) > 255 else 1.1 for ch in str(val))
                if text_len > col_widths[c]:
                    col_widths[c] = min(text_len + 1.5, 25.0)

    # 窄列特殊处理
    col_widths[min_col] = 4.0 
    if max_col - min_col >= 6: col_widths[max_col] = 6.0 

    total_width = sum(col_widths.values())
    total_height = sum(row_heights.values())

    # --- 5. 画布创建 (强制锁定 A4 尺寸 300DPI) ---
    fig, ax = plt.subplots(figsize=(8.27, 11.69), dpi=300)
    
    a4_ratio = 11.69 / 8.27
    table_ratio = total_height / total_width if total_width > 0 else 1
    
    if table_ratio > a4_ratio:
        target_width = total_height / a4_ratio
        x_pad = (target_width - total_width) / 2
        ax.set_xlim(-x_pad, total_width + x_pad)
        ax.set_ylim(total_height, 0)
    else:
        target_height = total_width * a4_ratio
        y_pad = (target_height - total_height) / 2
        ax.set_xlim(0, total_width)
        ax.set_ylim(total_height + y_pad, -y_pad)
        
    ax.axis('off')

    # --- 6. 逐个单元格精准绘制与样式复刻 ---
    y = 0
    for r in range(min_row, max_row + 1):
        x = 0
        rh = row_heights.get(r, 1.2)
        rtype = row_types.get(r, 'normal')
        
        for c in range(min_col, max_col + 1):
            cw = col_widths.get(c, 5.0)
            
            is_merged_top_left = True
            draw_w, draw_h = cw, rh
            
            if (r, c) in merged_dict:
                merge_info = merged_dict[(r, c)]
                if (r, c) != merge_info['top_left']:
                    is_merged_top_left = False 
                else:
                    mc_w = sum(col_widths.get(mc, 5.0) for mc in range(merge_info['top_left'][1], merge_info['bottom_right'][1] + 1))
                    mc_h = sum(row_heights.get(mr, 1.2) for mr in range(merge_info['top_left'][0], merge_info['bottom_right'][0] + 1))
                    draw_w, draw_h = mc_w, mc_h

            if is_merged_top_left:
                cell = ws.cell(row=r, column=c)
                
                # --- 背景色精准还原与层级色强行染色机制 ---
                bg_color = '#FFFFFF'
                if cell.fill and cell.fill.patternType == 'solid':
                    color = cell.fill.start_color
                    if color and hasattr(color, 'rgb') and color.rgb:
                        rgb = str(color.rgb)
                        if len(rgb) == 8 and rgb != '00000000':
                            bg_color = '#' + rgb[2:] 
                        elif len(rgb) == 6:
                            bg_color = '#' + rgb

                # 【终极保底设计】为体现层级，主动染色重要汇总行
                if rtype == 'region': bg_color = '#EAF2FA' # 浅蓝色突出大区汇总
                elif rtype == 'total': bg_color = '#FFF0E0' # 浅橙色突出合计
                elif rtype == 'header': bg_color = '#F2F2F2' # 浅灰色表头
                
                rect = patches.Rectangle((x, y), draw_w, draw_h, facecolor=bg_color, edgecolor='#000000', linewidth=0.5)
                ax.add_patch(rect)
                
                # --- 数值格式化精准还原 ---
                val = cell.value
                fmt = cell.number_format or "General"
                text = ""
                
                if val is not None and str(val).strip() != "":
                    if isinstance(val, (int, float)):
                        if '%' in fmt:
                            text = f"{val:.0%}" if ('0%' in fmt and '.00' not in fmt) else f"{val:.2%}"
                        elif ',' in fmt or val >= 1000 or val <= -1000:
                            if isinstance(val, float) and not val.is_integer():
                                text = f"{val:,.2f}".rstrip('0').rstrip('.')
                            else:
                                text = f"{val:,.0f}"
                        else:
                            if isinstance(val, float):
                                text = f"{val:.2f}".rstrip('0').rstrip('.')
                            else:
                                text = str(val)
                    elif isinstance(val, datetime.datetime):
                        text = val.strftime('%Y-%m-%d')
                    else:
                        text = str(val)
                
                # --- 字体加粗与字号层级规则 ---
                text_str = text.strip()
                is_bold = False
                font_size = 6
                
                if rtype == 'title':
                    is_bold = True
                    font_size = 9
                elif rtype == 'unit':
                    is_bold = False
                elif rtype in ['header', 'region', 'total']:
                    is_bold = True
                else:
                    is_bold = False
                    
                text_color = '#000000'
                if cell.font and cell.font.color and hasattr(cell.font.color, 'rgb') and cell.font.color.rgb:
                    rgb = str(cell.font.color.rgb)
                    extracted_color = '#' + rgb[2:] if len(rgb)==8 else '#000000'
                    if extracted_color != '#00000000': text_color = extracted_color

                # --- 对齐方式还原 ---
                halign = cell.alignment.horizontal if cell.alignment and cell.alignment.horizontal else 'center'
                if halign not in ['left', 'center', 'right']: halign = 'center'
                valign = cell.alignment.vertical if cell.alignment and cell.alignment.vertical else 'center'
                
                if halign == 'left': text_x = x + 0.5
                elif halign == 'right': text_x = x + draw_w - 0.5
                else: text_x = x + draw_w / 2
                
                if valign == 'top': text_y = y + 0.3
                elif valign == 'bottom': text_y = y + draw_h - 0.3
                else: text_y = y + draw_h / 2
                
                if isinstance(text, str) and len(text) > (draw_w / 1.5):
                    chars_per_line = max(1, int(draw_w / 1.5))
                    text = '\n'.join(textwrap.wrap(text, width=chars_per_line))
                    
                # --- 渲染最终文字 (调用专属粗体文件) ---
                if text:
                    kwargs = {
                        'ha': halign,
                        'va': 'center' if valign == 'center' else ('top' if valign == 'top' else 'bottom'),
                        'color': text_color,
                        'clip_on': True
                    }
                    
                    if is_bold and custom_font_bold:
                        # 使用专属粗体包
                        prop = custom_font_bold.copy()
                        prop.set_size(font_size)
                        kwargs['fontproperties'] = prop
                    elif custom_font_regular:
                        # 使用常规包
                        prop = custom_font_regular.copy()
                        prop.set_size(font_size)
                        kwargs['fontproperties'] = prop
                    else:
                        # 极端兜底
                        kwargs['weight'] = 'bold' if is_bold else 'normal'
                        kwargs['fontsize'] = font_size
                        
                    ax.text(text_x, text_y, text, **kwargs)

            x += cw
        y += rh

    plt.tight_layout(pad=1.0)
    img_stream = io.BytesIO()
    plt.savefig(img_stream, format='png', dpi=300, bbox_inches='tight')
    plt.close()
    img_stream.seek(0)
    return img_stream

# ==================== 导出文件生成逻辑 (PDF/Image) ====================

def generate_export_files_in_memory(file_stream):
    """根据操作系统，智能生成 PDF (Windows) 或 高清 PNG (Linux/云端)"""
    results = []
    logs = []
    today_mmdd = datetime.datetime.now().strftime('%m%d')
    sys_name = platform.system()
    
    # --- 严格指定的数据区域，彻底剔除无关信息框选 ---
    sheets_info = [
        {"name": "每日-中粮贸易外部赊销限额使用监控表", "range": "$A$1:$G$30", "base_title": "中粮贸易外部赊销限额使用监控表"}
    ]
    if datetime.datetime.now().weekday() == 3: # 周四特供
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
                app = win32com.client.Dispatch("Ket.Application") # 兼容 WPS
                
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
                    logs.append(f"   ⚠️ 跳过 {s_info['name']} (可能不存在或渲染失败): {str(e)}")
                    
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
                    # 使用 1:1 完美复刻新引擎
                    img_stream = render_sheet_range_to_image_stream(wb[s_info['name']], s_info['range'])
                    if img_stream:
                        out_name = f"{s_info['base_title']}{today_mmdd}.png"
                        results.append({"name": out_name, "data": img_stream.read(), "type": "png"})
                        logs.append(f"   ✅ 成功生成高清图片: {out_name}")
        except Exception as e:
            logs.append(f"❌ 跨平台渲染引擎出错: {str(e)}")
            
    return results, logs

# ==================== 主控入口 ====================

def process_credit_report(uploaded_file):
    """
    处理风险管理日报主入口。
    返回: (word_bytes, word_text_dict, export_files, logs, env_msg)
    """
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
