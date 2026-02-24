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


# ==================== 终极防蜷缩：像素级物理复刻引擎 ====================

def render_sheet_range_to_image_stream(ws, range_str):
    """
    100% 像素级物理对齐引擎：
    - 废除自主推断染色，全盘复刻 Excel 原生 HEX 色值。
    - 独立渲染标题与落款，根除排版挤压。
    - 智能剔除无内容的空白列。
    """
    if not MATPLOTLIB_AVAILABLE:
        return None

    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 1. 严格挂载常规与粗体库
    regular_path = os.path.join(current_dir, 'msyh.ttc')
    if not os.path.exists(regular_path): regular_path = os.path.join(current_dir, 'msyh.ttf')
    custom_font_regular = FontProperties(fname=regular_path) if os.path.exists(regular_path) else None

    bold_path = os.path.join(current_dir, 'msyhbd.ttc')
    if not os.path.exists(bold_path): bold_path = os.path.join(current_dir, 'msyhbd.ttf')
    custom_font_bold = FontProperties(fname=bold_path) if os.path.exists(bold_path) else custom_font_regular

    # 2. 初始范围锁定
    range_str = range_str.replace('$', '')
    min_col, min_row, max_col, max_row = range_boundaries(range_str)

    # 3. 动态截断：彻底剔除底部冗余内容（如"是否填报"）
    actual_max_row = max_row
    for r in range(min_row, max_row + 1):
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(min_col, max_col + 1)]
        combined = "".join(row_vals)
        if "是否填报" in combined or "填报说明" in combined:
            actual_max_row = r - 1 
            break
    
    # 4. 文本剥离术：找出独立排版元素，将其移出表格主体渲染逻辑
    title_text = ""
    author_text = ""
    date_text = ""
    unit_text = ""
    
    grid_start_row = None
    grid_end_row = actual_max_row

    for r in range(min_row, actual_max_row + 1):
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(min_col, max_col + 1)]
        combined = "".join(row_vals)
        
        # 捕获特殊文本
        if "汇总表" in combined or "监控表" in combined:
            if not title_text: title_text = next((v for v in row_vals if "表" in v), combined)
        elif "制表单位" in combined:
            author_text = next((v for v in row_vals if "制表单位" in v), combined)
        elif "截止时间" in combined:
            date_text = next((v for v in row_vals if "截止时间" in v), combined)
        elif "单位" in combined and ("万元" in combined or "万" in combined):
            unit_text = next((v for v in row_vals if "单位" in v), combined)
        
        # 寻找表格主体的起点 (表头起点)
        if grid_start_row is None and ("序号" in combined or "大区" in combined or "业务单位" in combined):
            grid_start_row = r

    if grid_start_row is None:
        grid_start_row = min_row # 兜底

    # 寻找表头的终点行（用于序号列底色规则控制）
    header_end_row = grid_start_row
    for r in range(grid_start_row, grid_end_row + 1):
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in range(min_col, max_col + 1)]
        combined = "".join(row_vals)
        if "沿江大区" in combined or "华东经营部" in combined or row_vals[0] == "1":
            header_end_row = r - 1
            break

    # 5. 动态有效列扫描 (彻底剥离右侧多余空白列)
    valid_cols_set = set()
    for r in range(grid_start_row, grid_end_row + 1):
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

    # 6. 计算行列自适应物理尺寸
    col_widths = {c: 4.0 for c in valid_cols}
    row_heights = {r: 2.2 for r in range(grid_start_row, grid_end_row + 1)}
    
    for r in range(grid_start_row, grid_end_row + 1):
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in valid_cols]
        combined = "".join(row_vals)
        if r <= header_end_row: row_heights[r] = 3.2
        elif "小计" in combined or "合计" in combined or "总计" in combined: row_heights[r] = 2.6
        else: row_heights[r] = 2.2

    # 根据数据长短拓宽有效列
    for r in range(grid_start_row, grid_end_row + 1):
        for c in valid_cols:
            is_spanned = False
            for mr in ws.merged_cells.ranges:
                if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
                    if mr.max_col > mr.min_col: is_spanned = True
            if is_spanned: continue # 跨列数据不作为单列宽度的评估标准，防止撑爆

            val = ws.cell(row=r, column=c).value
            if val:
                text_len = sum(1.8 if ord(ch) > 255 else 1.1 for ch in str(val))
                w = text_len * 0.9 + 1.5 
                if w > col_widths[c]: col_widths[c] = min(w, 22.0)

    # 极窄列处理
    col_widths[valid_cols[0]] = max(3.5, col_widths[valid_cols[0]])

    # 7. 组装绝对坐标画布 (A4 比例，无缩放形变)
    W_grid = sum(col_widths.values())
    H_grid = sum(row_heights.values())

    A4_W, A4_H = 8.27, 11.69
    margin_x = 0.4
    
    # 根据真实网格比例建立物理画板
    max_w_in = A4_W - 2 * margin_x
    S = max_w_in / W_grid 

    # 规划额外的 Y 轴空间用于独立绘制标题
    H_total_virtual = H_grid + 12.0 # 留足顶部和底部的虚拟高度空间
    H_in = H_total_virtual * S
    
    # 强制画布符合 A4 或自动延长适应极长表格
    Final_H = max(A4_H, H_in + 1.0)
    fig = plt.figure(figsize=(A4_W, Final_H), dpi=300)
    fig.patch.set_facecolor('white')

    left = margin_x / A4_W
    ax = fig.add_axes([left, 0.5 / Final_H, max_w_in / A4_W, H_in / Final_H])
    ax.set_xlim(0, W_grid)
    ax.set_ylim(H_total_virtual, 0)
    ax.axis('off')

    base_fs = 2.5 * S * 72 * 0.45 

    # ==================== 独立绘制外围文字 (精准排版要求) ====================
    # 制表单位：大标题上方，居左
    if author_text:
        ax.text(1.0, 1.5, author_text, ha='left', va='center', fontproperties=custom_font_regular, fontsize=base_fs*0.95)
    
    # 大标题：居中、加粗
    if title_text:
        prop_title = custom_font_bold.copy() if custom_font_bold else custom_font_regular
        if prop_title: prop_title.set_size(base_fs * 1.5)
        ax.text(W_grid / 2, 4.0, title_text, ha='center', va='center', fontproperties=prop_title, weight='bold', fontsize=base_fs*1.5)
    
    # 截止时间：标题下方，表格上方，居右
    if date_text:
        ax.text(W_grid - 1.0, 6.5, date_text, ha='right', va='center', fontproperties=custom_font_regular, fontsize=base_fs*0.95)

    # ==================== 绘制表格主体 ====================
    y_curr = 8.0 # 网格起始 Y 坐标
    
    for r in range(grid_start_row, grid_end_row + 1):
        x_curr = 0
        rh = row_heights[r]
        
        row_vals = [str(ws.cell(row=r, column=c).value or "").strip() for c in valid_cols]
        combined_row = "".join(row_vals)
        
        is_header_row = (r <= header_end_row)
        is_subtotal = "小计" in combined_row
        is_total = "合计" in combined_row or "总计" in combined_row
        is_zhengda_group = "正大集团合计" in combined_row
        is_region_row = (row_vals[0] == "" and len(row_vals) > 1 and "大区" in row_vals[1])
        
        # 加粗铁律判定
        is_bold = is_header_row or is_subtotal or is_total or is_zhengda_group

        for c in valid_cols:
            cw = col_widths[c]
            
            is_merged_top_left = True
            draw_w, draw_h = cw, rh
            
            # 合并单元格精准求和逻辑
            if ws.merged_cells:
                for mr in ws.merged_cells.ranges:
                    if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
                        if (r, c) != (mr.min_row, mr.min_col):
                            is_merged_top_left = False
                        else:
                            draw_w = sum(col_widths.get(mc, 0) for mc in range(mr.min_col, mr.max_col + 1) if mc in valid_cols)
                            draw_h = sum(row_heights.get(mr_i, 2.2) for mr_i in range(mr.min_row, mr.max_row + 1))
                        break

            if is_merged_top_left:
                cell = ws.cell(row=r, column=c)
                
                # --- 原生底色提取机制 (禁止自作主张染色) ---
                bg_color = '#FFFFFF'
                if cell.fill and cell.fill.patternType == 'solid' and cell.fill.start_color.rgb:
                    rgb = str(cell.fill.start_color.rgb)
                    if len(rgb) == 8 and rgb != '00000000': 
                        bg_color = '#' + rgb[2:]
                    elif len(rgb) == 6:
                        bg_color = '#' + rgb
                
                # 序号列底色修正铁律：非表头行强制剥夺底色
                if c == valid_cols[0] and not is_header_row:
                    bg_color = '#FFFFFF'

                # 绘制精准线框 (纯黑实线)
                rect = patches.Rectangle((x_curr, y_curr), draw_w, draw_h, facecolor=bg_color, edgecolor='#000000', linewidth=0.8)
                ax.add_patch(rect)
                
                # --- 数值格式完美对齐 (彻底服从原生 Format) ---
                val = cell.value
                fmt = cell.number_format or "General"
                text = ""
                
                if val is not None and str(val).strip() != "":
                    if isinstance(val, (int, float)):
                        # PDF 要求百分比形式的，必须且仅能由原生 % 号触发
                        if '%' in fmt:
                            if '.00' in fmt: text = f"{val:.2%}"
                            elif '.0' in fmt: text = f"{val:.1%}"
                            else: text = f"{val:.0%}"
                        elif ',' in fmt or (isinstance(val, (int, float)) and (val >= 1000 or val <= -1000)):
                            # 金额与大整数强加千分位
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
                        if "年" in fmt: text = val.strftime('%Y年%m月%d日')
                        else: text = val.strftime('%Y-%m-%d')
                    else:
                        text = str(val).strip()
                
                # --- 对齐方式与渲染 ---
                halign = 'center'
                valign = 'center'
                
                # 合并单元格一律居中
                excel_h = cell.alignment.horizontal if cell.alignment else None
                if excel_h in ['left', 'right', 'center']: halign = excel_h
                
                # 为了达到 PDF 美观标准，表头及合并内容强制居中
                if is_header_row or draw_w > cw + 1.0 or draw_h > rh + 1.0:
                    halign = 'center'
                    valign = 'center'
                    
                pad_x = 1.0
                if halign == 'left': text_x = x_curr + pad_x
                elif halign == 'right': text_x = x_curr + draw_w - pad_x
                else: text_x = x_curr + draw_w / 2
                
                text_y = y_curr + draw_h / 2
                
                if isinstance(text, str) and len(text) > (draw_w / 1.1):
                    wrap_w = max(1, int(draw_w / 1.1))
                    text = '\n'.join(textwrap.wrap(text, width=wrap_w))
                    
                if text:
                    kwargs = {
                        'ha': halign,
                        'va': valign,
                        'color': '#000000',
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

    # 绘制底部的单位
    if unit_text:
        ax.text(W_grid - 1.0, y_curr + 1.5, unit_text, ha='right', va='center', fontproperties=custom_font_regular, fontsize=base_fs*0.95)

    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', dpi=300, facecolor=fig.get_facecolor(), edgecolor='none', bbox_inches='tight')
    plt.close(fig)
    img_stream.seek(0)
    return img_stream

# ==================== 导出文件生成逻辑 ====================

def generate_export_files_in_memory(file_stream):
    """根据操作系统，智能生成 PDF 或完美防变形 PNG"""
    results = []
    logs = []
    today_mmdd = datetime.datetime.now().strftime('%m%d')
    sys_name = platform.system()
    
    sheets_info = [
        {"name": "每日-中粮贸易外部赊销限额使用监控表", "range": "$A$1:$G$30", "base_title": "中粮贸易外部赊销限额使用监控表"}
    ]
    
    # 试运行期间默认开启正大生成
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
                    # 调用纯净 1:1 底层渲染器
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
