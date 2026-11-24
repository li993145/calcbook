# -*- coding: utf-8 -*-
"""打印预览组件：按纸张实际比例分页绘制，鼠标拖动虚线即可调整页边距（仿 WPS）。"""

from PyQt5.QtWidgets import (
    QWidget, QScrollArea, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
    QPushButton, QSizePolicy, QFrame, QDoubleSpinBox, QApplication
)
from PyQt5.QtCore import Qt, QSizeF, QRectF, QSize, pyqtSignal, QTimer
from PyQt5.QtGui import (
    QPixmap, QPainter, QColor, QPen, QTextDocument, QFont, QCursor, QBrush
)
from ui_ribbon import make_icon

# 1cm 对应的文档像素数：Qt 富文本按屏幕逻辑 DPI 排 px（150% 缩放时为 144dpi），
# 预览画布必须用同一基准，纸张与字号比例才和实际打印一致。
def _screen_dpi():
    try:
        scr = QApplication.primaryScreen()
        dpi = float(scr.logicalDotsPerInch()) if scr is not None else 96.0
        return dpi if dpi > 0 else 96.0
    except Exception:
        return 96.0


def px_per_cm():
    """1cm 对应的文档像素数：Qt 富文本按屏幕逻辑 DPI 排 px（150% 缩放时为 144dpi），
    预览画布必须用同一基准，纸张与字号比例才和实际打印一致。

    注意：必须在 QApplication 建好之后调用（导入时取不到真实 DPI）。
    """
    return _screen_dpi() / 2.54
# 默认纸张 A4（mm）；实际纸张可由页面设置选择（跟随系统打印机支持的纸型）
DEFAULT_PAPER = (210.0, 297.0)

MARGIN_MIN = 0.5
MARGIN_MAX = 8.0


def paper_mm(report):
    """从报表设置取纸张尺寸（mm），兼容旧工程（无 paper 字段）。"""
    p = (report or {}).get('paper')
    if isinstance(p, dict):
        try:
            return float(p.get('w') or DEFAULT_PAPER[0]), float(p.get('h') or DEFAULT_PAPER[1])
        except (TypeError, ValueError):
            return DEFAULT_PAPER
    if isinstance(p, str) and p.strip():
        table = {'A3': (297, 420), 'A4': (210, 297), 'A5': (148, 210),
                 'B4': (250, 353), 'B5': (176, 250), 'LETTER': (215.9, 279.4),
                 'LEGAL': (215.9, 355.6)}
        up = p.strip().upper()
        for k, v in table.items():
            if up.startswith(k):
                return v
    return DEFAULT_PAPER


class CanvasScrollArea(QScrollArea):
    """报表预览的滚动区。

    刻意不按纸张尺寸撑大：否则窗口最小宽度会被纸张绑架，
    而且“适应宽度”会自我放大（纸张越大 → 视口越宽 → 缩放越大 → 纸张又变大）。
    """

    def minimumSizeHint(self):
        return QSize(120, 120)

    def sizeHint(self):
        return QSize(360, 360)


class ReportRenderer:
    """报表渲染器：把报表拆成「每页重复的头部」和「自动分页的正文」。

    Qt 富文本不支持跨页重复表头（<thead> 只出现在表格顶部），所以这里：
    1) 头部单独做成一个文档（报表页眉 + 标题 + 表头行）；
    2) 正文手动分页：逐行测高、按每页可用高度装箱，每页一个独立文档
       —— Qt 对放不进当页剩余空间的表格行会直接从页边界切半
       （实测），只有手动装箱才能保证整行绝不跨页；
    3) 绘制每一页时先画头部，再画该页正文 —— 于是每页都有表头。

    兼容旧接口：body_rows=None 时仍按整份 body_html 交给 Qt 自动分页
    （该路径行会被切开，仅保留给老调用方）。
    """

    def __init__(self, head_html, body_html, content_w, content_h,
                 footer_left='', footer_center='', footer_right='', show_page_no=True,
                 body_rows=None, table_open=''):
        self.content_w = max(40.0, float(content_w))
        self.content_h = max(40.0, float(content_h))
        self.footer_left = footer_left or ''
        self.footer_center = footer_center or ''
        self.footer_right = footer_right or ''
        self.show_page_no = bool(show_page_no)
        # 页脚占用正文可用高度，保证打印与预览都不会压到表格内容
        self.footer_h = 30.0 if (self.footer_left or self.footer_center or
                                 self.footer_right or self.show_page_no) else 0.0
        # 正文每页的高度 = 可用高度 - 头部高度 - 页脚高度
        self.head = self._make_doc(head_html, QSizeF(self.content_w, max(self.content_h, 1000.0)))
        # 头部高度按 documentSize 量（含表格边框/内边距，块测量会少算几个像素）
        self.head_h = max(0.0, float(self.head.documentLayout().documentSize().height()))
        self.body_page_h = max(40.0, self.content_h - self.head_h - self.footer_h)
        if body_rows is not None:
            # 手动装箱分页：每页元素 = (文档, 该页在文档里的纵向起点)
            self._pages = self._paginate_rows(body_rows, table_open)
        else:
            # 旧路径：整份正文交给 Qt 自动分页
            self.body = self._make_doc(body_html, QSizeF(self.content_w, 1000.0))
            self.body.setPageSize(QSizeF(self.content_w, self.body_page_h))
            self._pages = [(self.body, i * self.body_page_h)
                           for i in range(max(1, self.body.documentLayout().pageCount()))]
        self.page_count = max(1, len(self._pages))

    # ---------- 手动分页 ----------
    # 行高按 documentSize 原值测（比真实高度多算一个隐式尾段 ≈ 5~8px，
    # 偏保守方向，不会把行挤出页外）；整页校验时给这个余量。
    _IMPLICIT_PARA = 8.0

    def _paginate_rows(self, rows, table_open):
        """逐行测高 + 贪心装箱 + 整页校验：整行绝不跨页。

        单行连一页都放不下时：该行单独成文档、按 body_page_h 由 Qt
        拆成若干页续页显示（极端长计算式，保内容完整）。
        """
        budget = self.body_page_h + self._IMPLICIT_PARA
        heights = [measure_table_h((table_open or '') + html + '</table>', self.content_w)
                   for html in rows]
        items = list(zip(rows, heights))
        pages = []
        i, n = 0, len(items)
        while i < n:
            cur, cur_h = [], 0.0
            while i < n:
                html, h = items[i]
                if h > budget:
                    break                    # 超页行单独处理
                if cur and cur_h + h > budget:
                    break                    # 这行放不下：整行挪到下一页
                cur.append(html)
                cur_h += h
                i += 1
            if cur:
                # 整页校验：按页文档实测高度核对，实测超页就往后退行
                while cur:
                    doc = self._page_doc(cur, table_open)
                    ph = float(doc.documentLayout().documentSize().height())
                    if ph - self._IMPLICIT_PARA <= self.body_page_h + 0.5:
                        pages.append((doc, 0.0))
                        break
                    if len(cur) == 1:
                        # 单行实测就超页（理论上上游已缩字号）：按页高拆续页
                        doc.setPageSize(QSizeF(self.content_w, self.body_page_h))
                        cnt = max(1, doc.documentLayout().pageCount())
                        for k in range(cnt):
                            pages.append((doc, k * self.body_page_h))
                        break
                    cur.pop()
            if i < n and items[i][1] > budget:
                # 整行超页：单独一个文档按页高分页，续页也进页序列
                doc = self._make_doc((table_open or '') + items[i][0] + '</table>',
                                     QSizeF(self.content_w, self.body_page_h))
                cnt = max(1, doc.documentLayout().pageCount())
                for k in range(cnt):
                    pages.append((doc, k * self.body_page_h))
                i += 1
        if not pages:
            pages.append((self._make_doc('', QSizeF(self.content_w, self.body_page_h)), 0.0))
        return pages

    def _page_doc(self, rows, table_open):
        return self._make_doc((table_open or '') + ''.join(rows) + '</table>',
                              QSizeF(self.content_w, 100000.0))

    @staticmethod
    def _make_doc(html, page_size=None):
        doc = QTextDocument()
        doc.setDefaultFont(QFont('微软雅黑', 10))
        doc.setDocumentMargin(0)
        doc.setHtml(html or '')
        if page_size is not None:
            doc.setPageSize(page_size)
        return doc

    @staticmethod
    def _measure(doc):
        """文档实际内容高度（取所有非空内容的下边界，去掉末尾那个空段落）。"""
        try:
            lay = doc.documentLayout()
            lay.documentSize()          # 触发排版
            last = doc.lastBlock()
            h = 0.0
            b = doc.begin()
            while b.isValid():
                if b != last or b.text().strip():
                    h = max(h, float(lay.blockBoundingRect(b).bottom()))
                b = b.next()
            return max(0.0, h)
        except Exception:
            return 0.0

    # ---------- 绘制 ----------
    def draw_page(self, painter, index, x=0.0, y=0.0):
        """把第 index 页画到 painter 的 (x, y)（逻辑像素）。"""
        index = max(0, min(index, self.page_count - 1))
        doc, top = self._pages[index]
        painter.save()
        painter.translate(x, y)
        painter.setClipRect(QRectF(0, 0, self.content_w, self.content_h))
        # 头部：每页都画一遍（表头因此每页都有）
        self._draw_slice(self.head, painter,
                         QRectF(0, 0, self.content_w, self.head_h), 0.0)
        # 正文：只画属于本页的那一段
        self._draw_slice(doc, painter,
                         QRectF(0, top, self.content_w, self.body_page_h), self.head_h)
        # 页脚：左 / 中 / 右三栏，中间固定带“第 X 页，共 Y 页”
        if self.footer_h:
            self._draw_footer(painter, index)
        painter.restore()

    def _draw_footer(self, painter, index):
        """页脚：左签字栏、中间页码、右签字栏。"""
        y = self.content_h - self.footer_h
        third = self.content_w / 3.0
        painter.save()
        painter.setClipRect(QRectF(0, y, self.content_w, self.footer_h))
        font = QFont('微软雅黑', 8)
        painter.setFont(font)
        painter.setPen(QColor('#666666'))
        painter.drawText(QRectF(6, y, third - 12, self.footer_h),
                         Qt.AlignLeft | Qt.AlignVCenter, self.footer_left)
        page_no = '第 %d 页，共 %d 页' % (index + 1, self.page_count)
        center_text = (self.footer_center + '  ' + page_no if self.footer_center else page_no)
        painter.drawText(QRectF(third, y, third, self.footer_h),
                         Qt.AlignHCenter | Qt.AlignVCenter, center_text)
        painter.drawText(QRectF(third * 2, y, third - 6, self.footer_h),
                         Qt.AlignRight | Qt.AlignVCenter, self.footer_right)
        painter.restore()

    @staticmethod
    def _draw_slice(doc, painter, rect, offset_y):
        painter.save()
        painter.translate(0, offset_y - rect.top())
        doc.drawContents(painter, rect)
        painter.restore()


def measure_html(html, width):
    """测量一段 HTML 在给定宽度下排出来的高度（文档像素）。

    用于打印前判断“某个表格行是不是高过整页”：
    Qt 富文本会把高于一页的表格行从中间切开，导致计算式印不全，
    所以要先量一量，太高就自动缩小该行字号。
    """
    try:
        doc = QTextDocument()
        doc.setDefaultFont(QFont('微软雅黑', 10))
        doc.setDocumentMargin(0)
        doc.setHtml(html or '')
        doc.setPageSize(QSizeF(max(40.0, float(width)), 100000.0))
        return float(ReportRenderer._measure(doc))
    except Exception:
        return 0.0


_ROW_H_CACHE = {}


def measure_table_h(html, width):
    """测量表格 HTML 的真实高度（含单元格内边距与边框）。

    实测：QTextDocument 的 documentSize 恒等于排版后内容的真实底边
    （表格的 padding / 边框都算在内），而按文字块底部测量的 _measure
    会把最后一个单元格的下内边距和边框漏掉，每行少算几个像素、
    整页累积起来就把最后一行挤出页外 —— 手动分页必须用 documentSize。
    """
    try:
        key = (round(float(width)), hash(html or ''))
        got = _ROW_H_CACHE.get(key)
        if got is not None:
            return got
        doc = QTextDocument()
        doc.setDefaultFont(QFont('微软雅黑', 10))
        doc.setDocumentMargin(0)
        doc.setHtml(html or '')
        doc.setPageSize(QSizeF(max(40.0, float(width)), 100000.0))
        h = max(0.0, float(doc.documentLayout().documentSize().height()))
        if len(_ROW_H_CACHE) > 4000:
            _ROW_H_CACHE.clear()
        _ROW_H_CACHE[key] = h
        return h
    except Exception:
        return 0.0


class PageCanvas(QWidget):
    """纸张画布：绘制分页内容 + 四条页边距虚线（可拖动）。

    marginsChanged(dict) 在拖动过程中持续发出，便于主窗口即时保存设置。
    """

    marginsChanged = pyqtSignal(dict)

    PAD = 22        # 画布内边距
    GAP = 16        # 页与页之间的间隔

    def __init__(self, parent=None):
        super().__init__(parent)
        self.head_html = ''
        self.body_html = ''
        self.body_rows = None
        self.table_open = ''
        self.footer_left = ''
        self.footer_center = ''
        self.footer_right = ''
        self.show_page_no = True
        self.margins = {'top': 1.9, 'bottom': 1.9, 'left': 1.9, 'right': 1.9}
        self.landscape = False
        self.paper = DEFAULT_PAPER      # (宽mm, 高mm)
        self.zoom = 1.0
        self._renderer = None
        self._content = (100.0, 100.0)
        self._page_count = 1
        self._drag = None
        self._hover = None
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._rebuild)
        self.setMouseTracking(True)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.setCursor(Qt.ArrowCursor)

    # ---------- 几何换算 ----------
    def paper_cm(self):
        """当前纸张（cm），已按纸张方向交换宽高。"""
        w_mm, h_mm = self.paper
        w_cm, h_cm = w_mm / 10.0, h_mm / 10.0
        return (h_cm, w_cm) if self.landscape else (w_cm, h_cm)

    def page_size_px(self):
        w_cm, h_cm = self.paper_cm()
        scale = self.zoom * px_per_cm()
        return w_cm * scale, h_cm * scale, scale

    def content_px(self):
        """可打印区尺寸（文档像素，与实际打印同一基准，不随显示缩放变化）。"""
        w_cm, h_cm = self.paper_cm()
        cw = max(40.0, (w_cm - self.margins['left'] - self.margins['right']) * px_per_cm())
        ch = max(40.0, (h_cm - self.margins['top'] - self.margins['bottom']) * px_per_cm())
        return cw, ch

    def page_top(self, index=0):
        _w, h, _s = self.page_size_px()
        return self.PAD + index * (h + self.GAP)

    # ---------- 数据 ----------
    def set_content(self, head_html, body_html, margins=None, landscape=False,
                    zoom=None, paper=None, footer_left='', footer_center='',
                    footer_right='', show_page_no=True,
                    body_rows=None, table_open=''):
        self.head_html = head_html or ''
        self.body_html = body_html or ''
        self.body_rows = body_rows
        self.table_open = table_open or ''
        self.footer_left = footer_left or ''
        self.footer_center = footer_center or ''
        self.footer_right = footer_right or ''
        self.show_page_no = bool(show_page_no)
        if margins:
            self.margins = {k: float(margins.get(k, 1.9)) for k in ('top', 'bottom', 'left', 'right')}
        self.landscape = bool(landscape)
        if paper:
            try:
                self.paper = (float(paper[0]), float(paper[1]))
            except (TypeError, ValueError, IndexError):
                self.paper = DEFAULT_PAPER
        if zoom is not None:
            self.zoom = float(zoom)
        self._rebuild()

    def set_zoom(self, zoom):
        self.zoom = max(0.15, min(3.0, float(zoom)))
        self._rebuild()

    def _rebuild(self):
        w, h, s = self.page_size_px()
        # 内容区按“屏幕逻辑 DPI”排版（不随显示缩放变化），绘制时再整体缩放，
        # 这样预览的分页 / 每页行数与实际打印完全一致
        w_cm, h_cm = self.paper_cm()
        cw = max(40.0, (w_cm - self.margins['left'] - self.margins['right']) * px_per_cm())
        ch = max(40.0, (h_cm - self.margins['top'] - self.margins['bottom']) * px_per_cm())
        self._renderer = ReportRenderer(
            self.head_html, self.body_html, cw, ch,
            self.footer_left, self.footer_center, self.footer_right, self.show_page_no,
            body_rows=self.body_rows, table_open=self.table_open)
        self._content = (cw * self.zoom, ch * self.zoom)      # 设备像素
        self._page_count = self._renderer.page_count
        total_h = self._page_count * h + (self._page_count - 1) * self.GAP + 2 * self.PAD
        self.setFixedSize(int(round(w + 2 * self.PAD)), int(round(total_h)))
        self.update()

    # ---------- 绘制 ----------
    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setRenderHint(QPainter.TextAntialiasing, True)
        p.fillRect(self.rect(), QColor('#e9ecf0'))
        if self._renderer is None:
            return
        w, h, s = self.page_size_px()
        ml = self.margins['left'] * s
        mt = self.margins['top'] * s
        mr = self.margins['right'] * s
        mb = self.margins['bottom'] * s
        cw, ch = self._content

        for i in range(self._page_count):
            top = self.page_top(i)
            page = QRectF(self.PAD, top, w, h)
            # 阴影
            p.fillRect(page.adjusted(2, 3, 2, 5), QColor(0, 0, 0, 34))
            # 纸张
            p.fillRect(page, QColor('#ffffff'))
            # 内容（每页先画表头，再画该页正文）
            p.save()
            p.setClipRect(QRectF(self.PAD + ml, top + mt, cw, ch))
            p.translate(self.PAD + ml, top + mt)
            p.scale(self.zoom, self.zoom)
            self._renderer.draw_page(p, i, 0.0, 0.0)
            p.restore()
            # 纸张边框
            p.setPen(QPen(QColor('#c2c8d0'), 1))
            p.setBrush(Qt.NoBrush)
            p.drawRect(page)
            # 页脚与页码已由 ReportRenderer 画在正文区底部，这里不再重复绘制

            # 页边距区域淡蓝底 + 虚线（仅首页显示手柄）
            p.fillRect(QRectF(self.PAD, top, w, mt), QColor(70, 130, 200, 16))
            p.fillRect(QRectF(self.PAD, top + h - mb, w, mb), QColor(70, 130, 200, 16))
            p.fillRect(QRectF(self.PAD, top + mt, ml, h - mt - mb), QColor(70, 130, 200, 16))
            p.fillRect(QRectF(self.PAD + w - mr, top + mt, mr, h - mt - mb), QColor(70, 130, 200, 16))

            p.setPen(QPen(QColor('#3d7cc9'), 1, Qt.DashLine))
            x1, x2 = self.PAD + ml, self.PAD + w - mr
            y1, y2 = top + mt, top + h - mb
            p.drawLine(int(x1), int(top), int(x1), int(top + h))
            p.drawLine(int(x2), int(top), int(x2), int(top + h))
            p.drawLine(int(self.PAD), int(y1), int(self.PAD + w), int(y1))
            p.drawLine(int(self.PAD), int(y2), int(self.PAD + w), int(y2))

            if i == 0:
                # 四个拖动把手（上下把手放在偏左位置，避免压住标题与页码）
                hx = x1 + min(40, (x2 - x1) / 3)
                for name, cx, cy in (('left', x1, (y1 + y2) / 2),
                                     ('right', x2, (y1 + y2) / 2),
                                     ('top', hx, y1),
                                     ('bottom', hx, y2)):
                    active = (self._drag == name) or (self._hover == name and not self._drag)
                    p.setPen(QPen(QColor('#2f6fb5'), 1))
                    p.setBrush(QBrush(QColor('#2f6fb5' if active else '#7fb0e0')))
                    if name in ('left', 'right'):
                        p.drawRoundedRect(QRectF(cx - 4, cy - 13, 8, 26), 3, 3)
                    else:
                        p.drawRoundedRect(QRectF(cx - 13, cy - 4, 26, 8), 3, 3)
                p.setBrush(Qt.NoBrush)

            # 拖动时在把手上方显示当前值
            if self._drag in ('left', 'right', 'top', 'bottom'):
                txt = '%s %0.1f cm' % ({'left': '左边距', 'right': '右边距',
                                        'top': '上边距', 'bottom': '下边距'}[self._drag],
                                       self.margins[self._drag])
                p.setPen(Qt.NoPen)
                p.setBrush(QColor(47, 111, 181, 235))
                fm = p.fontMetrics()
                bw = fm.horizontalAdvance(txt) + 16
                if self._drag == 'left':
                    br = QRectF(x1 + 8, y1 + 8, bw, 22)
                elif self._drag == 'right':
                    br = QRectF(x2 - bw - 8, y1 + 8, bw, 22)
                elif self._drag == 'top':
                    br = QRectF(x1 + 8, y1 + 8, bw, 22)
                else:
                    br = QRectF(x1 + 8, y2 - 30, bw, 22)
                p.drawRoundedRect(br, 4, 4)
                p.setPen(QColor('#ffffff'))
                p.drawText(br, Qt.AlignCenter, txt)

    # ---------- 交互 ----------
    def _hit(self, pos):
        _w, h, s = self.page_size_px()
        top = self.page_top(0)
        x1 = self.PAD + self.margins['left'] * s
        x2 = self.PAD + _w - self.margins['right'] * s
        y1 = top + self.margins['top'] * s
        y2 = top + h - self.margins['bottom'] * s
        tol = 6
        if abs(pos.x() - x1) <= tol and y1 - 8 <= pos.y() <= y2 + 8:
            return 'left'
        if abs(pos.x() - x2) <= tol and y1 - 8 <= pos.y() <= y2 + 8:
            return 'right'
        if abs(pos.y() - y1) <= tol and x1 - 8 <= pos.x() <= x2 + 8:
            return 'top'
        if abs(pos.y() - y2) <= tol and x1 - 8 <= pos.x() <= x2 + 8:
            return 'bottom'
        return None

    def _cursor_for(self, name):
        if name in ('left', 'right'):
            return Qt.SizeHorCursor
        if name in ('top', 'bottom'):
            return Qt.SizeVerCursor
        return Qt.ArrowCursor

    def mousePressEvent(self, event):
        if event.button() != Qt.LeftButton:
            return super().mousePressEvent(event)
        name = self._hit(event.pos())
        if name:
            self._drag = name
            self.setCursor(self._cursor_for(name))
            self.update()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        pos = event.pos()
        if self._drag:
            self._apply_drag(pos)
            return
        name = self._hit(pos)
        if name != self._hover:
            self._hover = name
            self.setCursor(self._cursor_for(name))
            self.update()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._drag:
            name = self._drag
            self._apply_drag(event.pos(), final=True)
            self._drag = None
            self.setCursor(self._cursor_for(self._hit(event.pos())))
            self.update()
            self.marginsChanged.emit(dict(self.margins))
            _ = name
            return
        super().mouseReleaseEvent(event)

    def leaveEvent(self, event):
        if not self._drag:
            self._hover = None
            self.setCursor(Qt.ArrowCursor)
            self.update()
        super().leaveEvent(event)

    def _apply_drag(self, pos, final=False):
        _w, h, s = self.page_size_px()
        top = self.page_top(0)
        m = dict(self.margins)
        name = self._drag
        page_w_cm = _w / s
        page_h_cm = h / s
        if name == 'left':
            v = (pos.x() - self.PAD) / s
            v = max(MARGIN_MIN, min(page_w_cm - self.margins['right'] - 5.0, v))
            m['left'] = round(v, 1)
        elif name == 'right':
            v = (self.PAD + _w - pos.x()) / s
            v = max(MARGIN_MIN, min(page_w_cm - self.margins['left'] - 5.0, v))
            m['right'] = round(v, 1)
        elif name == 'top':
            v = (pos.y() - top) / s
            v = max(MARGIN_MIN, min(page_h_cm - self.margins['bottom'] - 5.0, v))
            m['top'] = round(v, 1)
        elif name == 'bottom':
            v = (top + h - pos.y()) / s
            v = max(MARGIN_MIN, min(page_h_cm - self.margins['top'] - 5.0, v))
            m['bottom'] = round(v, 1)
        if m != self.margins:
            self.margins = m
            # 拖动时实时重排（节流，避免卡顿）
            if final:
                self._timer.stop()
                self._rebuild()
            elif not self._timer.isActive():
                self._timer.start(60)
            self.marginsChanged.emit(dict(m))
        self.update()


class PrintPreviewPane(QWidget):
    """报表页：上方工具条 + 下方可滚动纸张预览（鼠标拖虚线改页边距）。"""

    def __init__(self, win, parent=None):
        super().__init__(parent)
        self.win = win
        self._zoom_mode = 'fit'

        lay = QVBoxLayout(self)
        lay.setContentsMargins(4, 4, 4, 4)
        lay.setSpacing(3)

        bar = QHBoxLayout()
        bar.setSpacing(4)
        bar.addWidget(QLabel('页边距(cm)：'))
        self.spins = {}
        for key, text in (('top', '上'), ('bottom', '下'), ('left', '左'), ('right', '右')):
            bar.addWidget(QLabel(text))
            sp = QDoubleSpinBox()
            sp.setRange(0.0, 8.0)
            sp.setSingleStep(0.1)
            sp.setDecimals(1)
            sp.setValue(1.9)
            sp.setFixedWidth(66)
            sp.valueChanged.connect(self._on_spin_changed)
            bar.addWidget(sp)
            self.spins[key] = sp
        bar.addSpacing(8)
        bar.addWidget(QLabel('显示比例：'))
        self.cb_zoom = QComboBox()
        for text, val in (('适应宽度', 'fit'), ('50%', 0.5), ('75%', 0.75),
                          ('100%', 1.0), ('125%', 1.25), ('150%', 1.5)):
            self.cb_zoom.addItem(text, val)
        self.cb_zoom.setFixedWidth(96)
        self.cb_zoom.currentIndexChanged.connect(self._on_zoom_changed)
        bar.addWidget(self.cb_zoom)
        self.lbl_paper = QLabel('')
        self.lbl_paper.setStyleSheet('color:#555;font-size:9pt;')
        bar.addWidget(self.lbl_paper)

        btn_setup = QPushButton('页面设置…')
        btn_setup.clicked.connect(self.win.page_setup)
        bar.addWidget(btn_setup)
        btn_refresh = QPushButton('刷新')
        btn_refresh.clicked.connect(self.refresh)
        bar.addWidget(btn_refresh)
        btn_print = QPushButton('打印')
        btn_print.setIcon(make_icon(QApplication.instance(), 'print'))
        btn_print.clicked.connect(self.win.print_report)
        bar.addWidget(btn_print)
        bar.addStretch(1)
        lay.addLayout(bar)

        self.scroll = CanvasScrollArea()
        self.scroll.setWidgetResizable(False)
        self.scroll.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setStyleSheet('QScrollArea { background:#e9ecf0; }')
        self.canvas = PageCanvas()
        self.canvas.marginsChanged.connect(self._on_canvas_margins)
        self.scroll.setWidget(self.canvas)
        lay.addWidget(self.scroll, 1)
        hint = QLabel('提示：把鼠标移到页面上的蓝色虚线上（光标会变成左右/上下箭头），按住拖动即可调整页边距；'
                      '上方四个数值框也可直接输入。页边距同时用于预览、打印与导出。')
        hint.setStyleSheet('color:#7a828c;font-size:9pt;')
        hint.setWordWrap(True)
        lay.addWidget(hint)
        self._loading = False

    # ---------- 外部调用 ----------
    def refresh(self):
        rep = self.win.project.report or {}
        margins = rep.get('margins') or {}
        m0 = float(rep.get('margin', 1.9))
        vals = {k: float(margins.get(k, m0)) for k in ('top', 'bottom', 'left', 'right')}
        self._loading = True
        for k, sp in self.spins.items():
            sp.setValue(vals[k])
        self._loading = False
        w_mm, h_mm = paper_mm(rep)
        name = (rep.get('paper') or {}).get('name', 'A4') if isinstance(rep.get('paper'), dict) else 'A4'
        self.lbl_paper.setText('纸张：%s  %0.0f×%0.0fmm（%s）'
                               % (name, w_mm, h_mm,
                                  '横向' if rep.get('orientation') == 'landscape' else '纵向'))
        fl, fc, fr = self.win.footer_fields()
        # 先把纸张/方向/页边距同步到画布，再按同一基准算可打印区，
        # 正文按真实页面高度把“高过整页”的行缩字号，保证整行印得全
        self.canvas.paper = (w_mm, h_mm)
        self.canvas.landscape = rep.get('orientation') == 'landscape'
        self.canvas.margins = dict(vals)
        cw, ch = self.canvas.content_px()
        table_open = self.win.report_table_open_html()
        body_rows = self.win.report_body_rows(cw, ch)
        self.canvas.set_content(self.win.report_head_html(), '',
                                vals, rep.get('orientation') == 'landscape',
                                paper=(w_mm, h_mm),
                                footer_left=fl, footer_center=fc, footer_right=fr,
                                show_page_no=True,
                                body_rows=body_rows, table_open=table_open)
        self._apply_zoom()
        # 首次显示时 viewport 尺寸还没定，等布局完成再按宽度适配一次
        QTimer.singleShot(0, self._apply_zoom)

    def _apply_zoom(self):
        val = self.cb_zoom.currentData()
        if val == 'fit':
            w_cm, _h_cm = self.canvas.paper_cm()
            page_w = w_cm * px_per_cm()
            avail = max(200, self.scroll.viewport().width() - 2 * PageCanvas.PAD - 24)
            zoom = avail / page_w
        else:
            zoom = float(val)
        self.canvas.set_zoom(zoom)

    # ---------- 事件 ----------
    def _on_zoom_changed(self, _idx):
        self._apply_zoom()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.cb_zoom.currentData() == 'fit':
            self._apply_zoom()

    def _on_spin_changed(self, _v):
        if self._loading:
            return
        self.canvas.margins = {k: sp.value() for k, sp in self.spins.items()}
        self.canvas._rebuild()
        self._save_margins(self.canvas.margins)

    def _on_canvas_margins(self, margins):
        self._loading = True
        for k, sp in self.spins.items():
            sp.setValue(float(margins.get(k, 1.9)))
        self._loading = False
        self._save_margins(margins)

    def _save_margins(self, margins):
        rep = dict(self.win.project.report or {})
        rep['margins'] = {k: float(margins.get(k, 1.9)) for k in ('top', 'bottom', 'left', 'right')}
        self.win.project.report = rep
        self.win._dirty = True
        self.win.status.showMessage(
            '页边距：上 %.1f / 下 %.1f / 左 %.1f / 右 %.1f cm（已保存）'
            % (rep['margins']['top'], rep['margins']['bottom'],
               rep['margins']['left'], rep['margins']['right']), 2500)
