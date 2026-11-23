# -*- coding: utf-8 -*-
"""紧凑 Ribbon 工具栏。

小图标(16px) + 小字(9pt)，组内按钮按两行网格紧凑排布，
组标题在底部，整条工具栏很矮。不使用 Qt 高 DPI 自动缩放，
保证在不同缩放比例的屏幕上显示尺寸一致。
图标用 QStyle 标准图标或代码绘制字符，不依赖外部图片。
"""

import os
import sys

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QToolButton, QLabel, QFrame,
    QSizePolicy, QStyle, QButtonGroup, QPushButton, QApplication
)
from PyQt5.QtCore import Qt, QSize, pyqtSignal
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor, QFont, QPen

STYLE_ICONS = {
    'new': QStyle.SP_FileIcon,
    'open': QStyle.SP_DirOpenIcon,
    'save': QStyle.SP_DialogSaveButton,
    'copy': QStyle.SP_FileDialogDetailedView,
    'undo': QStyle.SP_ArrowBack,
    'redo': QStyle.SP_ArrowForward,
    'add': QStyle.SP_FileDialogNewFolder,
    'delete': QStyle.SP_TrashIcon,
    'up': QStyle.SP_ArrowUp,
    'down': QStyle.SP_ArrowDown,
    'calc': QStyle.SP_DialogApplyButton,
    'export': QStyle.SP_DriveCDIcon,
    'doc': QStyle.SP_FileDialogInfoView,
    'settings': QStyle.SP_FileDialogContentsView,
    'close': QStyle.SP_DialogCloseButton,
}


def make_printer_icon(size=20, color='#2f5597'):
    """手绘打印机图标（Qt 没有 SP_PrinterIcon，标准信息图标不像打印机）。"""
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    scale = size / 20.0
    p.scale(scale, scale)
    pen = QPen(QColor(color))
    pen.setWidthF(1.3)
    p.setPen(pen)
    p.setBrush(QColor('#ffffff'))
    p.drawRect(5, 2, 10, 6)      # 进纸口上方纸张
    p.drawRect(3, 8, 14, 6)      # 机身
    p.drawRect(6, 13, 8, 5)      # 出纸托盘
    p.setBrush(QColor(color))    # 状态指示灯
    p.drawEllipse(14, 10, 2, 2)
    p.end()
    return QIcon(pm)


def make_icon(app, name_or_char, color='#2f5597'):
    """生成小尺寸图标（16~20px）。"""
    if name_or_char == 'print':
        return make_printer_icon(20, color)
    if name_or_char in STYLE_ICONS:
        return app.style().standardIcon(STYLE_ICONS[name_or_char])
    pm = QPixmap(20, 20)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setFont(QFont('Segoe UI Symbol', 13))
    p.setPen(QColor(color))
    p.drawText(pm.rect(), Qt.AlignCenter, name_or_char)
    p.end()
    return QIcon(pm)


class RibbonGroup(QFrame):
    """一组按钮：默认单排放置（组标题在底部），由 RibbonPage 负责换行。"""

    MAX_ROWS = 1

    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.NoFrame)
        self._count = 0
        self._rows = 1
        self.buttons = []
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        self.grid = QGridLayout()
        self.grid.setContentsMargins(0, 1, 0, 0)
        self.grid.setHorizontalSpacing(1)
        self.grid.setVerticalSpacing(0)
        lay.addLayout(self.grid)
        lbl = QLabel(title)
        lbl.setAlignment(Qt.AlignHCenter)
        lbl.setStyleSheet('color:#666;font-size:8pt;')
        lay.addWidget(lbl)
        self.setStyleSheet('''
            RibbonGroup { border-right: 1px solid #dcdcdc; }
            QToolButton { border:1px solid transparent; border-radius:2px;
                          background:transparent; color:#222; font-size:9pt;
                          padding:1px 0px; }
            QToolButton:hover { background:#e8f0fa; border:1px solid #c5d8f0; }
            QToolButton:pressed { background:#d4e4f7; }
        ''')

    def set_rows(self, rows):
        """1=单排（默认）；2=组内两排紧凑模式（窗口太窄时用）。"""
        rows = max(1, int(rows))
        if rows == self._rows:
            return
        self._rows = rows
        for b in self.buttons:
            self.grid.removeWidget(b)
        for i, b in enumerate(self.buttons):
            self.grid.addWidget(b, i % rows, i // rows)
        self.updateGeometry()

    def add_button(self, app, text, icon_key, slot, shortcut='', big=True):
        btn = QToolButton(self)
        btn.setText(text)
        btn.setIcon(make_icon(app, icon_key))
        btn.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        btn.setIconSize(QSize(16, 16))
        # 按文字实际宽度定宽（比 Qt 默认留白更紧）。
        # 宽度 = 文字宽 + 余量，保证不出现省略号裁剪。
        btn.setFont(QFont('微软雅黑', 9))
        tw = btn.fontMetrics().horizontalAdvance(text)
        btn.setFixedWidth(max(24, tw + 10))
        btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        if shortcut:
            btn.setShortcut(shortcut)
            btn.setToolTip('%s (%s)' % (text, shortcut))
        else:
            btn.setToolTip(text)
        btn.clicked.connect(lambda checked=False, sl=slot: sl())
        r = self._count % self._rows
        c = self._count // self._rows
        self.grid.addWidget(btn, r, c)
        self._count += 1
        self.buttons.append(btn)
        return btn


class RibbonPage(QWidget):
    """一组组的按钮区：所有组默认排成一排；宽度不够时，整组换到下一排。"""

    def __init__(self, groups, parent=None):
        super().__init__(parent)
        self.groups = list(groups)
        for g in self.groups:
            g.setParent(self)
        self._laid_w = -1
        self._rows = []
        self._h = 64

    # 单排所需宽度（供 sizeHint 用）
    def _natural_width(self):
        return sum(g.sizeHint().width() for g in self.groups) + 8

    def sizeHint(self):
        return QSize(self._natural_width(), self._h)

    def minimumSizeHint(self):
        return QSize(200, self._h)

    def showEvent(self, ev):
        super().showEvent(ev)
        self._laid_w = -1
        self._relayout()
        # 布局收敛后再校一次（最大化/首次显示时父窗口宽度可能晚一步确定）
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(0, self._recheck)
        QTimer.singleShot(120, self._recheck)

    def _recheck(self):
        self._laid_w = -1
        self._relayout()

    def resizeEvent(self, ev):
        super().resizeEvent(ev)
        self._relayout()

    def _compute_rows(self, avail):
        """按宽度把组排成若干排，返回 [[组,...], ...]"""
        rows, cur, curw = [], [], 0
        for g in self.groups:
            gw = g.sizeHint().width()
            if cur and curw + gw > avail:      # 这一排放不下了 → 换排
                rows.append(cur)
                cur, curw = [], 0
            cur.append(g)
            curw += gw
        if cur:
            rows.append(cur)
        return rows

    def _avail_width(self):
        """可用宽度：取自身、父控件、顶层窗口三者最小，避免布局未收敛时按过宽排布被裁。"""
        w = self.width()
        pw = self.parentWidget()
        if pw is not None and pw.width() > 50:
            w = min(w, pw.width())
        top = self.window()
        if top is not None and top is not self and top.width() > 100:
            w = min(w, top.width() - 4)
        return max(160, w - 12)     # 留余量，避免最后一组贴边被裁

    def _relayout(self):
        if self.width() < 50:          # 还没布局出宽度，等 resizeEvent
            return
        avail = self._avail_width()
        if avail == self._laid_w:
            return
        self._laid_w = avail

        # 1) 首选：所有组单排（组内按钮也单排）
        for g in self.groups:
            g.set_rows(1)
        rows = self._compute_rows(avail)
        # 2) 单排放不下且超过两排 → 组内改双排紧凑模式，尽量压回一到两排
        if len(rows) > 2:
            for g in self.groups:
                g.set_rows(2)
            rows = self._compute_rows(avail)

        y = 1
        for row in rows:
            x = 2
            h = max(g.sizeHint().height() for g in row)
            for g in row:
                gw = g.sizeHint().width()
                g.setGeometry(x, y, gw, h)
                x += gw
            y += h
        self._rows = rows

        h = y + 1
        if h != self._h:
            self._h = h
            self.setFixedHeight(h)
            # 高度变了，通知上层布局（Ribbon → 主窗口）重新排版
            p = self.parentWidget()
            while p is not None:
                if p.layout() is not None:
                    p.layout().invalidate()
                p.updateGeometry()
                p = p.parentWidget()


class TitleStrip(QWidget):
    """窗口标题栏（自绘，取代 Windows 系统标题栏）：

    左端一块信息：``[橘黄底软件图标] 软件名 V版本 | 工程文件名称 | [文件存放详细地址]``；
    最右是最小化/最大化/关闭按钮。

    V3.4.1 起不再显示快捷函数按钮（构造参数 actions 仅为兼容旧调用保留）。
    """

    maximizeRequested = pyqtSignal()
    minimizeRequested = pyqtSignal()
    closeRequested = pyqtSignal()

    def __init__(self, actions=None, parent=None, icon_path=''):
        super().__init__(parent)
        self.setFixedHeight(30)
        self.setMouseTracking(True)
        self.setStyleSheet('''
            TitleStrip { background:#e9edf3; border-bottom:1px solid #c2c8d2; }
            QToolButton { border:1px solid transparent; border-radius:2px; background:transparent; }
            QToolButton:hover { background:#dbe4f0; border:1px solid #b9c8dc; }
            QToolButton:pressed { background:#cddcf0; }
        ''')
        lay = QHBoxLayout(self)
        lay.setContentsMargins(8, 1, 2, 1)
        lay.setSpacing(1)
        # 兼容旧调用：V3.4.1 起主程序不再传入 actions，这里不会显示任何快捷按钮
        self.qat = []
        app = QApplication.instance()
        for item in (actions or []):
            text = item[0]
            icon_key = item[1] if len(item) > 1 else ''
            slot = item[2] if len(item) > 2 else None
            b = QToolButton(self)
            b.setIcon(make_icon(app, icon_key) if icon_key else QIcon())
            b.setIconSize(QSize(16, 16))
            b.setAutoRaise(True)
            b.setToolTip(text)
            b.setFixedSize(26, 26)
            b.setFocusPolicy(Qt.NoFocus)
            if callable(slot):
                b.clicked.connect(lambda checked=False, sl=slot: sl())
            lay.addWidget(b)
            self.qat.append(b)
        if self.qat:
            sep = QFrame(self)
            sep.setFrameShape(QFrame.VLine)
            sep.setFixedHeight(16)
            sep.setStyleSheet('color:#c2c8d2;')
            lay.addSpacing(4)
            lay.addWidget(sep)
            lay.addSpacing(6)
        # 左端信息块：软件图标 → 软件名+版本 → 工程文件名称 → [文件存放详细地址]
        self._path = ''
        self._fname = ''
        self._path_max_w = 620
        self.lbl_icon = QLabel('')
        self.lbl_icon.setFixedSize(24, 24)
        self.lbl_icon.setAlignment(Qt.AlignCenter)
        # V3.4.1：橘黄底色 + 圆角，左上角软件图标更显眼
        self.lbl_icon.setStyleSheet(
            'background:#ff8c00; border:1px solid #d97700; border-radius:4px;')
        self.lbl_icon.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.lbl_icon.setToolTip('工程量计算书')
        self._app_pixmap = self._load_app_icon(icon_path)
        if self._app_pixmap is not None:
            self.lbl_icon.setPixmap(self._app_pixmap)
        lay.addWidget(self.lbl_icon)
        lay.addSpacing(4)

        self.lbl_app = QLabel('')
        self.lbl_app.setStyleSheet("color:#22303f; font-family:'微软雅黑'; font-size:9pt;")
        self.lbl_app.setContentsMargins(0, 0, 0, 0)
        self.lbl_app.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.lbl_sep = QLabel('|')
        self.lbl_sep.setStyleSheet("color:#b6bec9; font-family:'微软雅黑'; font-size:9pt;")
        self.lbl_sep.setContentsMargins(0, 0, 0, 0)
        self.lbl_sep.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.lbl_name = QLabel('')            # 工程文件名称（文件名）
        self.lbl_name.setStyleSheet("color:#22303f; font-family:'微软雅黑'; font-size:9pt;")
        self.lbl_name.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.lbl_name.setToolTip('当前工程文件名称')
        self.lbl_name.setContentsMargins(0, 0, 0, 0)
        self.lbl_sep2 = QLabel('|')
        self.lbl_sep2.setStyleSheet("color:#b6bec9; font-family:'微软雅黑'; font-size:9pt;")
        self.lbl_sep2.setContentsMargins(0, 0, 0, 0)
        self.lbl_sep2.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.lbl_path = QLabel('')
        self.lbl_path.setStyleSheet("color:#4a5665; font-family:'微软雅黑'; font-size:9pt;")
        self.lbl_path.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.lbl_path.setCursor(Qt.IBeamCursor)
        self.lbl_path.setContentsMargins(0, 0, 0, 0)
        for w in (self.lbl_app, self.lbl_sep, self.lbl_name, self.lbl_sep2):
            lay.addWidget(w)
            lay.addSpacing(6)
        lay.addWidget(self.lbl_path)
        lay.addStretch(1)

        # 最右：最小化 / 最大化 / 关闭（自绘标题栏必须自带窗口按钮）
        self.btn_min = self._make_win_button('min', '最小化')
        self.btn_max = self._make_win_button('max', '最大化')
        self.btn_close = self._make_win_button('close', '关闭')
        self.btn_min.clicked.connect(self.minimizeRequested)
        self.btn_max.clicked.connect(self.maximizeRequested)
        self.btn_close.clicked.connect(self.closeRequested)
        for b in (self.btn_min, self.btn_max, self.btn_close):
            lay.addWidget(b)

        # 拖动窗口
        self._drag_from = None
        self._drag_win_pos = None

    # ---------- 窗口按钮 ----------
    def _make_win_button(self, kind, tip):
        b = QToolButton(self)
        b.setFixedSize(30, 26)
        b.setAutoRaise(True)
        b.setToolTip(tip)
        b.setFocusPolicy(Qt.NoFocus)
        b._kind = kind
        b.setIcon(self._win_icon(kind))
        b.setStyleSheet('''
            QToolButton { border:1px solid transparent; background:transparent; }
            QToolButton:hover { background:#dbe4f0; }
        ''')
        return b

    @staticmethod
    def _win_icon(kind, color='#3b4552'):
        """画最小化/最大化/关闭的小图形（不依赖字体符号，避免缺字）。"""
        pm = QPixmap(16, 16)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        pen = QPen(QColor(color))
        pen.setWidth(1)
        p.setPen(pen)
        if kind == 'min':
            p.drawLine(3, 12, 12, 12)
        elif kind == 'max':
            p.drawRect(3, 4, 9, 8)
        elif kind == 'restore':
            p.drawRect(5, 3, 7, 7)
            p.fillRect(4, 5, 8, 7, QColor('#e9edf3'))
            p.drawRect(4, 5, 8, 7)
        elif kind == 'close':
            p.drawLine(4, 4, 11, 11)
            p.drawLine(11, 4, 4, 11)
        p.end()
        return QIcon(pm)

    def set_maximized(self, maximized):
        """最大化时把中间那个按钮换成“还原”图形。"""
        try:
            self.btn_max.setIcon(self._win_icon('restore' if maximized else 'max'))
            self.btn_max.setToolTip('还原' if maximized else '最大化')
        except Exception:
            pass

    # ---------- 工程名称（不再显示在标题栏里） ----------
    def set_project_name(self, name):
        """标题栏不再显示工程名，这里只记下来备用。"""
        self._proj_name = name or ''

    # ---------- 拖动窗口 ----------
    def mousePressEvent(self, ev):
        if ev.button() == Qt.LeftButton:
            self._drag_from = ev.globalPos()
            self._drag_win_pos = self.window().pos()
        super().mousePressEvent(ev)

    def mouseMoveEvent(self, ev):
        if self._drag_from is not None and self._drag_win_pos is not None:
            d = ev.globalPos() - self._drag_from
            self.window().move(self._drag_win_pos + d)
        super().mouseMoveEvent(ev)

    def mouseReleaseEvent(self, ev):
        self._drag_from = None
        self._drag_win_pos = None
        super().mouseReleaseEvent(ev)

    def mouseDoubleClickEvent(self, ev):
        self._drag_from = None
        self._drag_win_pos = None
        self.maximizeRequested.emit()
        super().mouseDoubleClickEvent(ev)

    # ---------- 软件图标 ----------
    @staticmethod
    def _load_app_icon(icon_path=''):
        """取软件图标（优先显式 ico/png 路径，QIcon 会挑最合适的尺寸）。"""
        cands = []
        if icon_path:
            cands.append(icon_path)
        base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        cands += [os.path.join(base, 'calcbook.ico'), os.path.join(base, 'calcbook.png')]
        for p in cands:
            if p and os.path.exists(p):
                try:
                    pm = QIcon(p).pixmap(20, 20)
                    if pm is not None and not pm.isNull():
                        return pm
                except Exception:
                    continue
        # 最后再退回程序自带的窗口图标
        app = QApplication.instance()
        if app is not None:
            ic = app.windowIcon()
            if ic is not None and not ic.isNull():
                pm = ic.pixmap(20, 20)
                if pm is not None and not pm.isNull():
                    return pm
        return None

    # ---------- 软件名/版本 + 工程文件名称 + 文件位置 ----------
    def set_file_info(self, title, path='', fname=''):
        """title 例如“工程量计算书 V2.5.0”；path 为当前工程文件全路径；fname 为文件名。"""
        self.lbl_app.setText(title or '')
        self.lbl_sep.setVisible(bool(title))
        self._path = path or ''
        if not fname and self._path:
            try:
                fname = os.path.basename(os.path.normpath(self._path))
            except Exception:
                fname = ''
        self._fname = fname or ''
        self.lbl_name.setText(self._fname)
        self.lbl_sep2.setVisible(bool(self._fname))
        self._refresh_file_label()

    def _path_available_w(self):
        """路径标签还能占多少宽度：信息块右边界 → 右侧窗口按钮之间。"""
        right_edge = 0
        for w in (self.lbl_icon, self.lbl_app, self.lbl_sep, self.lbl_name, self.lbl_sep2):
            right_edge = max(right_edge, w.x() + w.width())
        if right_edge <= 0:
            # 布局还没跑起来时，用各控件的期望宽度粗算
            right_edge = (self.lbl_icon.width() + 4
                          + self.lbl_app.sizeHint().width() + 6
                          + self.lbl_sep.sizeHint().width() + 6
                          + self.lbl_name.sizeHint().width() + 6
                          + self.lbl_sep2.sizeHint().width() + 6)
        # 右侧三个窗口按钮 30px + 右边距
        avail = self.width() - right_edge - 30 * 3 - 12
        return max(80, min(getattr(self, '_path_max_w', 620), avail))

    def _refresh_file_label(self):
        p = getattr(self, '_path', '') or ''
        avail = self._path_available_w()
        if abs(self.lbl_path.maximumWidth() - avail) >= 1:
            self.lbl_path.setMaximumWidth(avail)
        if not p:
            self.lbl_path.setText('[未保存]')
            self.lbl_path.setToolTip('当前工程还没有保存到磁盘；Ctrl+S 保存后这里显示文件位置')
            self.lbl_path.setStyleSheet("color:#9aa2ae; font-family:'微软雅黑'; font-size:9pt;")
            return
        self.lbl_path.setStyleSheet("color:#4a5665; font-family:'微软雅黑'; font-size:9pt;")
        fm = self.lbl_path.fontMetrics()
        try:                       # 与资源管理器一致：反斜杠
            shown_path = os.path.normpath(p)
        except Exception:
            shown_path = p
        # 方括号前后各留 16px，否则末尾两个字符会被裁掉
        shown = fm.elidedText(shown_path, Qt.ElideMiddle, max(40, avail - 16))
        self.lbl_path.setText('[%s]' % shown)
        tip = p
        if getattr(self, '_fname', ''):
            tip = '%s\n%s' % (self._fname, p)
        self.lbl_path.setToolTip(tip)

    def resizeEvent(self, ev):
        super().resizeEvent(ev)
        self._refresh_file_label()


class RibbonWidget(QWidget):
    """sections: {页签名: [(组名, [(文字,图标,回调,快捷键,big),...]),...]}"""

    def __init__(self, app, sections, parent=None):
        super().__init__(parent)
        self.app = app
        # 横向用 Ignored：宽度始终跟随父窗口，不会因为 sizeHint 比窗口宽而被裁切
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        self.setStyleSheet('''
            RibbonWidget { background:#fafbfc; border-bottom:1px solid #c9ced6; }
            QPushButton#ribbonTab { border:1px solid transparent; background:transparent;
                    padding:3px 12px; color:#222; font-size:10pt; }
            QPushButton#ribbonTab:hover { background:#eef3f8; }
            QPushButton#ribbonTab:checked { background:#ffffff;
                    border:1px solid #c9ced6; border-bottom:2px solid #2f5597; }
        ''')

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # 页签行：靠左排列（软件名/版本/文件位置在更上面那条标题条里）
        tab_bar = QHBoxLayout()
        tab_bar.setContentsMargins(8, 2, 8, 0)
        tab_bar.setSpacing(2)
        self.tab_group = QButtonGroup(self)
        self.tab_group.setExclusive(True)
        names = list(sections.keys())
        for i, tab_name in enumerate(names):
            tab = QPushButton(tab_name)
            tab.setObjectName('ribbonTab')
            tab.setCheckable(True)
            tab.setChecked(i == 0)
            self.tab_group.addButton(tab)
            tab_bar.addWidget(tab)
        tab_bar.addStretch(1)
        lay.addLayout(tab_bar)

        # 按钮页：所有页叠放，show_page 切换显示（隐藏控件不占布局空间）
        self.pages = {}
        for i, tab_name in enumerate(names):
            groups = []
            for gname, buttons in sections[tab_name]:
                g = RibbonGroup(gname)
                for text, icon, slot, sc, big in buttons:
                    g.add_button(app, text, icon, slot, sc, big)
                groups.append(g)
            page = RibbonPage(groups, self)
            page.setVisible(i == 0)
            lay.addWidget(page)
            self.pages[tab_name] = page
        self.current_page = names[0] if names else None
        self.updateGeometry()

    def show_page(self, name):
        for n, p in self.pages.items():
            p.setVisible(n == name)
        self.current_page = name
        self.updateGeometry()
