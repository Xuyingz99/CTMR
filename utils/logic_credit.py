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

# ==================== Word 生成逻辑 (保持不变) ====================

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


# ==================== 终极优化：1:1 绝对映射坐标防变形引擎 ====================

def render_sheet_range_to_image_stream(ws, range_str):
    """
    终极版图片渲染器：强制数学映射物理比例，根除任何表格变形与内容蜷缩！
    """
    if not MATPLOTLIB_AVAILABLE:
        return None

    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 1. 加载常规与粗体字体 (彻底解决部分环境字体无加粗效果的问题)
    regular_path = os.path.join(current_dir, 'msyh.ttc')
    if not os.path.exists(regular_path): regular_path = os.path.join(current_dir, 'msyh.ttf')
    custom_font_regular = FontProperties(fname=regular_path) if os.path.exists(regular_path) else None

    bold_path = os.path.join(current_dir, 'msyhbd.ttc')
    if not os.path.exists(bold_path): bold_path = os.path.join(current_dir, 'msyhbd.ttf')
    custom_font_bold = FontProperties(fname=bold_path) if os.path.exists(bold_path) else custom_font_regular

    # --- 1. 精确框选有效范围 (根除"是否填报"等冗余信息) ---
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

    # --- 3. 智能推断行类型 (为了强制施加你的样式规则) ---
    row_types = {}
    for r in range(min_row, max_row + 1):
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(min_col, max_col + 1)]
        combined = "".join(row_vals)
        
        if r == min_row:
            row_types[r] = 'title'
        elif "单位" in combined and "万元" in combined:
            row_types[r] = 'unit'
        elif "序号" in combined or "业务单位" in combined or "赊销余额" in combined:
            row_types[r] = 'header'
        elif "合计" in combined:
            row_types[r] = 'total'
        elif any("大区" in v for v in row_vals):
            row_types[r] = 'region'
        else:
            row_types[r] = 'normal'

    # --- 4. 智能推断哪几列是百分比 (解决小数失真问题) ---
    col_is_percent = {c: False for c in range(min_col, max_col + 1)}
    for c in range(min_col, max_col + 1):
        for r in range(min_row, min_row + 4): # 只扫描头部判断列性质
            v = str(ws.cell(row=r, column=c).value or "")
            if "率" in v or "占比" in v:
                col_is_percent[c] = True
                break

    # --- 5. 动态内容自适应计算宽高度 (防文字截断与留白) ---
    col_widths = {c: 4.0 for c in range(min_col, max_col + 1)}
    row_heights = {}
    
    # 定义基础行高比例
    for r in range(min_row, max_row + 1):
        if row_types[r] == 'title': row_heights[r] = 4.0
        elif row_types[r] == 'unit': row_heights[r] = 1.5
        elif row_types[r] == 'header': row_heights[r] = 3.0
        else: row_heights[r] = 2.5

    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            if (r, c) in merged_dict and merged_dict[(r, c)]['top_left'] != (r, c): continue
            val = ws.cell(row=r, column=c).value
            if val:
                # 汉字算 1.8 宽，英数算 1.0 宽
                text_len = sum(1.8 if ord(ch) > 255 else 1.0 for ch in str(val))
                w = text_len * 0.9 + 2.0 # 留出完美边距
                if w > col_widths[c]: col_widths[c] = min(w, 20.0)

    # 强制压窄特定列宽度使排版更紧凑美观
    col_widths[min_col] = 3.5 # 序号列

    # --- 6. 核心重构：绝对防变形画布映射引擎 ---
    # 计算表格的内部虚拟总宽与总高
    W = sum(col_widths.values())
    H = sum(row_heights.values())

    # A4 纸张物理尺寸定义 (英寸)
    A4_W, A4_H = 8.27, 11.69
    margin_x = 0.5  # 左右各留半寸
    margin_y_top = 1.0 # 顶部留一寸
    margin_y_bot = 0.5
    
    max_w_in = A4_W - 2 * margin_x
    max_h_in = A4_H - margin_y_top - margin_y_bot

    # 计算缩放系数 S (1虚拟数据单位 = 多少英寸)
    S = max_w_in / W
    if H * S > max_h_in: 
        S = max_h_in / H # 如果表格太长，以高度压缩为准

    W_in = W * S # 表格的物理绝对宽
    H_in = H * S # 表格的物理绝对高

    # 创建标准 A4 画布
    fig = plt.figure(figsize=(A4_W, A4_H), dpi=300)
    fig.patch.set_facecolor('white')

    # 将坐标系完美嵌入 A4 纸的中央靠上，且坐标比例 1:1 绝对对齐
    left = (A4_W - W_in) / 2 / A4_W
    bottom = (A4_H - margin_y_top - H_in) / A4_H
    width_frac = W_in / A4_W
    height_frac = H_in / A4_H

    ax = fig.add_axes([left, bottom, width_frac, height_frac])
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0) # 反转 Y 轴，让表格从上往下画
    ax.axis('off')

    # 计算完美的字体尺寸 (占行高的约 42%)
    base_font_size = 2.5 * S * 72 * 0.42 

    # --- 7. 逐个单元格精准绘制与样式复刻 ---
    y = 0
    for r in range(min_row, max_row + 1):
        x = 0
        rh = row_heights[r]
        rtype = row_types[r]
        
        for c in range(min_col, max_col + 1):
            cw = col_widths[c]
            
            is_merged_top_left = True
            draw_w, draw_h = cw, rh
            
            if (r, c) in merged_dict:
                info = merged_dict[(r, c)]
                if (r, c) != info['top_left']:
                    is_merged_top_left = False 
                else:
                    draw_w = sum(col_widths.get(mc, 4.0) for mc in range(info['top_left'][1], info['bottom_right'][1] + 1))
                    draw_h = sum(row_heights.get(mr, 2.5) for mr in range(info['top_left'][0], info['bottom_right'][0] + 1))

            if is_merged_top_left:
                cell = ws.cell(row=r, column=c)
                
                # ------ 颜色层级精准控制 ------
                bg_color = '#FFFFFF'
                if rtype == 'header': bg_color = '#E8EDF2' # 浅灰蓝色
                elif rtype == 'region': bg_color = '#D9E1F2' # 稍深灰蓝色(突出大区)
                elif rtype == 'total': bg_color = '#FCE4D6' # 浅橙色(突出合计)
                elif rtype in ['title', 'unit']: bg_color = '#FFFFFF'
                else:
                    # 读取 Excel 本身底色作为备用
                    if cell.fill and cell.fill.patternType == 'solid' and cell.fill.start_color.rgb:
                        rgb = str(cell.fill.start_color.rgb)
                        if len(rgb) == 8 and rgb != '00000000': bg_color = '#' + rgb[2:] 

                # 剔除顶部的线框，确保清爽
                lw = 0 if rtype in ['title', 'unit'] else 0.8
                rect = patches.Rectangle((x, y), draw_w, draw_h, facecolor=bg_color, edgecolor='#000000', linewidth=lw)
                ax.add_patch(rect)
                
                # ------ 数值格式化机制 ------
                val = cell.value
                fmt = cell.number_format or "General"
                text = ""
                
                if val is not None and str(val).strip() != "":
                    if isinstance(val, (int, float)):
                        # 判断如果是使用率，强制转成完美的 50%
                        if '%' in fmt or col_is_percent[c]:
                            text = f"{val:.0%}"
                        else:
                            # 强制添加千分位，并抹除末尾冗余的 .00
                            if isinstance(val, float) and not val.is_integer():
                                text = f"{val:,.2f}".rstrip('0').rstrip('.')
                            else:
                                text = f"{val:,.0f}"
                    elif isinstance(val, datetime.datetime):
                        text = val.strftime('%Y-%m-%d')
                    else:
                        text = str(val).strip()
                
                # ------ 字体加粗规则 (严格遵循你的 5 条铁律) ------
                # ①标题 ②表头 ③大区汇总 ④合计 -> 全程加粗；⑤其余 -> 降级不加粗
                is_bold = rtype in ['title', 'header', 'region', 'total']
                
                # ------ 对齐方式规则 ------
                halign = 'center'
                valign = 'center'
                if rtype == 'title': halign = 'center'
                elif rtype == 'unit': halign = 'right'
                elif rtype == 'region': halign = 'center' # 大区也强制居中更好看
                else:
                    excel_h = cell.alignment.horizontal if cell.alignment else None
                    if excel_h in ['left', 'right', 'center']: halign = excel_h
                    
                pad_x = 1.0
                if halign == 'left': text_x = x + pad_x
                elif halign == 'right': text_x = x + draw_w - pad_x
                else: text_x = x + draw_w / 2
                text_y = y + draw_h / 2
                
                # 动态字号
                fs = base_font_size
                if rtype == 'title': fs = base_font_size * 1.5
                elif rtype == 'unit': fs = base_font_size * 0.9

                # 防止溢出自动换行
                if len(text) > (draw_w / 1.1):
                    wrap_w = max(1, int(draw_w / 1.1))
                    text = '\n'.join(textwrap.wrap(text, width=wrap_w))
                    
                # ------ 最终渲染渲染文字 ------
                if text:
                    kwargs = {
                        'ha': halign,
                        'va': valign,
                        'color': '#000000',
                        'clip_on': True
                    }
                    
                    if is_bold and custom_font_bold:
                        prop = custom_font_bold.copy()
                        prop.set_size(fs)
                        kwargs['fontproperties'] = prop
                    elif custom_font_regular:
                        prop = custom_font_regular.copy()
                        if is_bold: prop.set_weight('bold')
                        prop.set_size(fs)
                        kwargs['fontproperties'] = prop
                    else:
                        # 极端防崩溃兜底
                        kwargs['weight'] = 'bold' if is_bold else 'normal'
                        kwargs['fontsize'] = fs
                        
                    ax.text(text_x, text_y, text, **kwargs)

            x += cw
        y += rh

    # 不再需要 tight_layout 改变坐标系，因为我们已经手动精准计算了！
    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close(fig)
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
                    # 使用 1:1 绝对防变形映射引擎
                    img_stream = render_sheet_range_to_image_stream(wb[s_info['name']], s_info['range'])
                    if img_stream:
                        out_name = f"{s_info['base_title']}{today_mmdd}.png"
                        results.append({"name": out_name, "data": img_stream.read(), "type": "png"})
                        logs.append(f"   ✅ 成功生成高级重构图片: {out_name}")
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
