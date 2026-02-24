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
        
        # 网页展示收集
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
                center_text_block += f"\n分大区情况如下：\n{group_line}\n"
                
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

# ==================== 核心升级：1:1 完美还原的图片渲染引擎 ====================

def render_sheet_range_to_image_stream(ws, range_str):
    """(完全重构) 基于 patches 自定义绘图的终极 1:1 还原引擎"""
    if not MATPLOTLIB_AVAILABLE:
        return None

    # --- 获取本地字体对象 (强制微软雅黑，解决所有乱码和粗体问题) ---
    current_dir = os.path.dirname(os.path.abspath(__file__))
    yahei_path = os.path.join(current_dir, 'msyh.ttc')
    if not os.path.exists(yahei_path):
        yahei_path = os.path.join(current_dir, 'msyh.ttf') 
    custom_font = FontProperties(fname=yahei_path) if os.path.exists(yahei_path) else None

    # --- 1. 范围解析 ---
    range_str = range_str.replace('$', '')
    min_col, min_row, max_col, max_row = range_boundaries(range_str)

    # --- 2. 映射所有合并单元格 (解决大标题和层级不对齐的问题) ---
    merged_dict = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_col <= max_col and mr.max_col >= min_col and mr.min_row <= max_row and mr.max_row >= min_row:
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    merged_dict[(r, c)] = {
                        'top_left': (mr.min_row, mr.min_col),
                        'bottom_right': (mr.max_row, mr.max_col)
                    }

    # --- 3. 提取行列真实宽度与高度 (解决列宽错乱、留白问题) ---
    col_widths = {}
    for c in range(min_col, max_col + 1):
        letter = get_column_letter(c)
        dim = ws.column_dimensions.get(letter)
        w = dim.width if (dim and dim.width) else 12.0
        col_widths[c] = w * 7.5  # 等比例放大宽度

    row_heights = {}
    for r in range(min_row, max_row + 1):
        dim = ws.row_dimensions.get(r)
        h = dim.height if (dim and dim.height) else 18.0
        row_heights[r] = h * 1.3

    total_width = sum(col_widths.values())
    total_height = sum(row_heights.values())

    # --- 4. 画布创建 (强制锁定 A4 尺寸) ---
    fig, ax = plt.subplots(figsize=(8.27, 11.69), dpi=300)
    
    # 根据 A4 比例动态缩放内容使其水平居中且不被裁切
    a4_ratio = 11.69 / 8.27
    table_ratio = total_height / total_width if total_width > 0 else 1
    
    if table_ratio > a4_ratio:
        target_width = total_height / a4_ratio
        x_pad = (target_width - total_width) / 2
        ax.set_xlim(-x_pad, total_width + x_pad)
        ax.set_ylim(total_height, 0) # 翻转 Y 轴，符合从上到下的阅读习惯
    else:
        target_height = total_width * a4_ratio
        y_pad = (target_height - total_height) / 2
        ax.set_xlim(0, total_width)
        ax.set_ylim(total_height + y_pad, -y_pad)
        
    ax.axis('off')

    # --- 5. 逐个单元格精准绘制 ---
    y = 0
    for r in range(min_row, max_row + 1):
        x = 0
        rh = row_heights.get(r, 20)
        for c in range(min_col, max_col + 1):
            cw = col_widths.get(c, 80)
            
            is_merged_top_left = True
            draw_w, draw_h = cw, rh
            
            # 判断是否为合并单元格
            if (r, c) in merged_dict:
                merge_info = merged_dict[(r, c)]
                if (r, c) != merge_info['top_left']:
                    is_merged_top_left = False # 被合并的附属单元格不渲染
                else:
                    # 累加合并区域的总宽和总高
                    mc_w = sum(col_widths.get(mc, 80) for mc in range(merge_info['top_left'][1], merge_info['bottom_right'][1] + 1))
                    mc_h = sum(row_heights.get(mr, 20) for mr in range(merge_info['top_left'][0], merge_info['bottom_right'][0] + 1))
                    draw_w, draw_h = mc_w, mc_h

            if is_merged_top_left:
                cell = ws.cell(row=r, column=c)
                
                # --- 背景色精准还原 (解决底色丢失) ---
                bg_color = '#FFFFFF'
                if cell.fill and cell.fill.patternType == 'solid':
                    color = cell.fill.start_color
                    if color.type == 'rgb' and color.rgb:
                        rgb = str(color.rgb)
                        if len(rgb) == 8 and rgb != '00000000':
                            bg_color = '#' + rgb[2:]
                        elif len(rgb) == 6:
                            bg_color = '#' + rgb
                
                # 画出外框与底色
                rect = patches.Rectangle((x, y), draw_w, draw_h, facecolor=bg_color, edgecolor='#000000', linewidth=0.5)
                ax.add_patch(rect)
                
                # --- 数值格式精准还原 (解决小数变长串、百分比消失) ---
                val = cell.value
                text = ""
                if val is not None and val != "":
                    fmt = cell.number_format or "General"
                    if isinstance(val, (int, float)):
                        if '%' in fmt:
                            # 百分比还原
                            text = f"{val:.0%}" if ('0%' in fmt and '.00' not in fmt) else f"{val:.2%}"
                        elif ',' in fmt or val >= 1000 or val <= -1000:
                            # 千分位格式还原
                            text = f"{val:,.2f}" if ('0.00' in fmt or '0.0' in fmt) else f"{val:,.0f}"
                        else:
                            if isinstance(val, float):
                                text = f"{val:.2f}".rstrip('0').rstrip('.')
                            else:
                                text = str(val)
                    elif isinstance(val, datetime.datetime):
                        text = val.strftime('%Y-%m-%d')
                    else:
                        text = str(val)
                        
                # --- 对齐方式还原 ---
                halign = cell.alignment.horizontal if cell.alignment and cell.alignment.horizontal else 'center'
                if halign not in ['left', 'center', 'right']: halign = 'center'
                valign = cell.alignment.vertical if cell.alignment and cell.alignment.vertical else 'center'
                
                if halign == 'left': text_x = x + 3
                elif halign == 'right': text_x = x + draw_w - 3
                else: text_x = x + draw_w / 2
                
                if valign == 'top': text_y = y + 3
                elif valign == 'bottom': text_y = y + draw_h - 3
                else: text_y = y + draw_h / 2
                
                # --- 字体加粗与颜色还原 ---
                is_bold = cell.font.bold if cell.font else False
                font_weight = 'bold' if is_bold else 'normal'
                
                # 字体大小适配
                font_size = 7
                if cell.font and cell.font.size:
                    font_size = max(5, min(cell.font.size * 0.65, 12))
                    
                text_color = '#000000'
                if cell.font and cell.font.color and cell.font.color.rgb:
                    rgb = str(cell.font.color.rgb)
                    extracted_color = '#' + rgb[2:] if len(rgb)==8 else '#000000'
                    if extracted_color != '#00000000': text_color = extracted_color

                # 长文本自动换行以防超出边框截断
                if isinstance(text, str) and len(text) > (draw_w / 6):
                    chars_per_line = max(1, int(draw_w / 6))
                    text = '\n'.join(textwrap.wrap(text, width=chars_per_line))
                    
                # 渲染文字
                if text:
                    kwargs = {
                        'ha': halign,
                        'va': 'center' if valign == 'center' else ('top' if valign == 'top' else 'bottom'),
                        'color': text_color,
                        'clip_on': True
                    }
                    if custom_font:
                        # 使用自定义字体对象并强制附带粗细设置
                        prop = custom_font.copy()
                        prop.set_weight('bold' if is_bold else 'normal')
                        prop.set_size(font_size)
                        kwargs['fontproperties'] = prop
                    else:
                        kwargs['weight'] = font_weight
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
                    # 使用精准框选的新引擎
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
