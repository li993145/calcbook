# -*- coding: utf-8 -*-
"""Excel 导出：用 openpyxl 生成真正的 .xlsx，WPS 可直接打开。"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from .units import precision_for_unit, section_precision
from .model import cell_display_text

# 纸张名称 → openpyxl 的 paperSize 常量名
PAPERSIZE_MAP = {
    'A3': 'PAPERSIZE_A3', 'A4': 'PAPERSIZE_A4', 'A5': 'PAPERSIZE_A5',
    'A6': 'PAPERSIZE_A6', 'B4': 'PAPERSIZE_B4', 'B5': 'PAPERSIZE_B5',
    'B6': 'PAPERSIZE_B6', 'LETTER': 'PAPERSIZE_LETTER', 'LEGAL': 'PAPERSIZE_LEGAL',
    'TABLOID': 'PAPERSIZE_TABLOID', 'C5': 'PAPERSIZE_C5', 'DL': 'PAPERSIZE_DL',
    '16K': 'PAPERSIZE_A5', '8K': 'PAPERSIZE_A3',
}


def _paper_code(ws, report):
    """报表纸张 → openpyxl 常量；未知（含自定义）回退 A4。"""
    paper = report.get('paper') or {}
    name = paper.get('name', '') if isinstance(paper, dict) else str(paper or '')
    up = (name or '').upper().replace(' ', '')
    for key, attr in PAPERSIZE_MAP.items():
        if up.startswith(key):
            return getattr(ws, attr, ws.PAPERSIZE_A4)
    return getattr(ws, 'PAPERSIZE_A4', 9)


def _fill_of(row, col):
    """单元格底色：单元格 > 整行 > 整列。"""
    key = col['key']
    c = (getattr(row, 'cell_fill', None) or {}).get(key)
    if c:
        return c
    if getattr(row, 'fill', ''):
        return row.fill
    return col.get('fill', '') or ''


def export_xlsx(project, path, precision=3, report=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "工程量计算书"

    report = report or getattr(project, 'report', {}) or {}
    thin = Side(style='thin', color='B0B0B0')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill('solid', fgColor='D9E2F3')
    sec_fill = PatternFill('solid', fgColor='FFF2CC')
    sum_fill = PatternFill('solid', fgColor='E2EFDA')
    halign = report.get('halign', 'center')
    valign = report.get('valign', 'center')

    cols = project.columns
    ncol = len(cols)
    last_letter = get_column_letter(ncol)

    # 页眉
    if report.get('header'):
        ws.merge_cells(f'A1:{last_letter}1')
        c = ws['A1']
        c.value = report['header']
        c.font = Font(name='微软雅黑', size=9, color='7F7F7F')
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 16

    # 标题
    title = report.get('title') or (project.name + ' 工程量计算书')
    title_row = 2 if report.get('header') else 1
    ws.merge_cells(f'A{title_row}:{last_letter}{title_row}')
    c = ws.cell(title_row, 1, title)
    c.font = Font(name='微软雅黑', size=15, bold=True)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[title_row].height = 28

    cur = title_row + 1
    if project.info:
        info_text = '  '.join(f'{k}: {v}' for k, v in project.info.items() if v)
        ws.merge_cells(f'A{cur}:{last_letter}{cur}')
        c = ws.cell(cur, 1, info_text)
        c.font = Font(name='微软雅黑', size=9, color='595959')
        c.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[cur].height = 18
        cur += 1

    headers = [cd['title'] for cd in cols]
    hr = cur
    for j, h in enumerate(headers, 1):
        c = ws.cell(hr, j, h)
        c.font = Font(name='微软雅黑', size=10, bold=True)
        c.fill = header_fill
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = border
    ws.row_dimensions[hr].height = 22

    r = hr + 1

    def _fmt(digits):
        return '0.' + '0' * int(digits) if digits > 0 else '0'

    def write_section(sec, level):
        nonlocal r
        # 分部标题行
        c = ws.cell(r, 1, '  ' * level + sec.name)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        c.font = Font(name='微软雅黑', size=10, bold=True)
        c.fill = sec_fill
        c.border = border
        c.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[r].height = 20
        r += 1

        for row in sec.rows:
            vals = []
            # 该行工程量按单位决定小数位（m/m2/kg→2、t→3、计件→0）
            row_fmt = _fmt(precision_for_unit(getattr(row, 'unit', ''), precision))
            for cd in cols:
                key = cd['key']
                if key == 'no_count':
                    vals.append('是' if row.no_count else '')
                elif key == 'value':
                    vals.append(row.value)
                elif key == 'img':
                    vals.append('')
                else:
                    # V3.9：备注列 / 自定义列里写了公式 → 导出算出来的数值
                    cc = (getattr(row, 'cell_calc', None) or {}).get(key)
                    if cc and cc.get('value') is not None:
                        vals.append(cc['value'])       # 数值型，WPS 里还能继续参与计算
                    else:
                        vals.append(cell_display_text(
                            row, key, precision_for_unit(getattr(row, 'unit', ''),
                                                        precision)))
            for j, v in enumerate(vals, 1):
                c = ws.cell(r, j, v)
                c.font = Font(name='微软雅黑', size=9)
                c.border = border
                c.alignment = Alignment(horizontal=halign, vertical=valign,
                                        wrap_text=True)
                if cols[j-1]['key'] == 'value' and isinstance(v, float):
                    c.number_format = row_fmt
                bg = _fill_of(row, cols[j-1])
                if bg:
                    c.fill = PatternFill('solid', fgColor=bg.replace('#', '').upper())
                if row.error:
                    c.font = Font(name='微软雅黑', size=9, color='C00000')
            r += 1

        # 小计行
        total = sec.total()
        if total is not None and (sec.rows or sec.children):
            c = ws.cell(r, 1, '小计')
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max(1, ncol-1))
            c2 = ws.cell(r, ncol, total)
            for cc in (c, c2):
                cc.fill = sum_fill
                cc.font = Font(name='微软雅黑', size=9, bold=True)
                cc.border = border
            c2.number_format = _fmt(section_precision(sec, precision))
            c.alignment = Alignment(horizontal='right')
            r += 1

        for child in sec.children:
            write_section(child, level + 1)

    # 遍历所有计算书（兼容旧项目：无 books 时用 root）
    books = getattr(project, 'books', None) or [project.root]
    for book in books:
        write_section(book, 0)

    widths = [cd.get('width', 100) for cd in cols]
    for j, w in enumerate(widths, 1):
        # 界面列宽是像素，Excel 列宽按字符数计，约 7px≈1 字符
        ws.column_dimensions[get_column_letter(j)].width = max(8, round(w / 7))
    ws.freeze_panes = f'A{hr+1}'

    # 页脚左 / 中 / 右（兼容旧版单一 footer 字段）
    fl = str(report.get('footer_left', '') or '')
    fc = str(report.get('footer_center', '') or '')
    fr_txt = str(report.get('footer_right', '') or '')
    if not (fl or fc or fr_txt):
        fc = str(report.get('footer', '') or '')
    if fl or fc or fr_txt:
        fr = r + 1
        ws.merge_cells(f'A{fr}:{last_letter}{fr}')
        c = ws.cell(fr, 1, '   '.join(x for x in (fl, fc, fr_txt) if x))
        c.font = Font(name='微软雅黑', size=9, color='7F7F7F')
        c.alignment = Alignment(horizontal='center', vertical='center')

    # 页边距（单位 cm → 英寸）、纸张、页眉页脚
    margins = report.get('margins') or {}
    m0 = float(report.get('margin', 1.9))

    def _cm(key):
        try:
            return max(0.0, float(margins.get(key, m0))) / 2.54
        except (TypeError, ValueError):
            return m0 / 2.54

    try:
        ws.page_margins.left = _cm('left')
        ws.page_margins.right = _cm('right')
        ws.page_margins.top = _cm('top')
        ws.page_margins.bottom = _cm('bottom')
        ws.page_margins.header = 0.3
        ws.page_margins.footer = 0.3
        ws.page_setup.paperSize = _paper_code(ws, report)
        ws.page_setup.orientation = ('landscape' if report.get('orientation') == 'landscape'
                                     else 'portrait')
        ws.oddHeader.center.text = report.get('header', '')
        page_no = '第 &P 页，共 &N 页'
        ws.oddFooter.left.text = fl
        ws.oddFooter.center.text = (fc + '  ' + page_no) if fc else page_no
        ws.oddFooter.right.text = fr_txt
        ws.print_title_rows = f'{hr}:{hr}'      # 每页重复表头
    except Exception:
        pass

    # 常量页（V2.4 起带“说明”列）
    if project.constants:
        cs = wb.create_sheet('常量表')
        cs.append(['常量代码', '值', '说明'])
        for c in cs[1]:
            c.font = Font(bold=True)
            c.fill = header_fill
        notes = getattr(project, 'constant_notes', None) or {}
        for k, v in project.constants.items():
            cs.append([k, v, notes.get(k, '')])
        cs.column_dimensions['A'].width = 20
        cs.column_dimensions['B'].width = 16
        cs.column_dimensions['C'].width = 26

    wb.save(path)
    return path
