# -*- coding: utf-8 -*-
"""工程量计算书 —— 桌面主界面"""

import sys, os, json, time, datetime, uuid, ctypes, re, tempfile, traceback, shutil
from decimal import Decimal, ROUND_HALF_UP
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QSplitter,
    QTreeWidget, QTreeWidgetItem, QTableWidget, QTableWidgetItem, QTabWidget,
    QToolBar, QAction, QPushButton, QLineEdit, QLabel, QFileDialog,
    QMessageBox, QInputDialog, QMenu, QTextEdit, QDialog, QFormLayout,
    QDialogButtonBox, QHeaderView, QAbstractItemView, QStatusBar, QFrame,
    QDoubleSpinBox, QComboBox, QToolButton, QShortcut, QStyledItemDelegate,
    QGridLayout, QRadioButton, QButtonGroup, QColorDialog, QSizePolicy,
    QAbstractItemDelegate, QAbstractButton, QCheckBox, QStyle,
    QStyleOptionViewItem, QStyleOptionButton, QListWidget, QListWidgetItem,
    QScrollArea, QGroupBox, QSpinBox
)
from PyQt5.QtCore import (Qt, QSettings, QMimeData, QEvent, QTimer, QUrl, QRectF, QPointF,
                          QPoint, QSizeF, QRect, pyqtSignal, QObject, QItemSelectionModel,
                          QItemSelection)
from PyQt5.QtGui import (QFont, QFontMetrics, QColor, QKeySequence, QIcon, QTextCursor,
                         QDesktopServices, QPixmap, QPainter, QPen, QBrush, QPalette,
                         QCursor)

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from core.model import Project, Section, Row, DEFAULT_LEVELS
from core.exporter import export_xlsx
from core.units import (UNIT_LIST, UNIT_GROUPS, INTEGER_UNITS, REBAR_CONSTANTS,
                        precision_for_unit, section_precision, fmt_const)
from ui_ribbon import RibbonWidget, TitleStrip
from ui_printpreview import PrintPreviewPane
from ui_sortfilter import fill_series
from core.highlighter import FormulaHighlighter, compute_runs, FCOLOR

APP_NAME = "工程量计算书"
APP_VER = "3.1.0"
FILE_EXT = '.gcls'          # 工程文件后缀（内容仍是 JSON，便于排查与迁移）

# ---- 工程文件关联（让资源管理器把 .gcls 显示成软件图标）----
ASSOC_PROGID = 'CalcBook.Project'
ASSOC_DESC = '工程量计算书工程文件'

# ---- 错误日志：打包成 exe 后没有控制台，异常写到这里便于排查 ----
LOG_FILE = os.path.join(tempfile.gettempdir(), 'calcbook_error.log')

# ---- 对话框按钮中文化：Qt 未加载中文翻译时，标准按钮会显示英文 ----
_DIALOG_BUTTON_TEXT = {
    'ok': '确定', 'cancel': '取消', 'yes': '是', 'no': '否',
    'close': '关闭', 'apply': '应用', 'save': '保存',
    'discard': '放弃', 'reset': '重置', 'help': '帮助',
}


def localize_dialog_buttons(dialog):
    """把对话框里的 OK / Cancel 等标准按钮文字统一改成中文。"""
    try:
        for btn in dialog.findChildren(QAbstractButton):
            text = btn.text()
            key = text.replace('&', '').replace('...', '').replace('…', '').strip().lower()
            zh = _DIALOG_BUTTON_TEXT.get(key)
            if zh and zh != text:
                btn.setText(zh)
    except Exception:
        pass


class DialogButtonI18nFilter(QObject):
    """应用级过滤器：任何弹窗显示时都把英文标准按钮换成中文。"""

    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.Show and isinstance(obj, QDialog):
            localize_dialog_buttons(obj)
            # 少数对话框（QInputDialog 等）显示后还会补建按钮，下一轮事件再刷一次
            QTimer.singleShot(0, lambda dlg=obj: localize_dialog_buttons(dlg))
        return super().eventFilter(obj, ev)


def log_exception(tag=''):
    """把当前异常写入日志。用于“绝不能让异常冒到 Qt 事件循环”的槽函数。"""
    try:
        import traceback
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write('\n[%s] %s\n' % (time.strftime('%Y-%m-%d %H:%M:%S'), tag))
            traceback.print_exc(file=f)
    except Exception:
        pass


def current_exe():
    """打包后返回 exe 全路径；源码运行时返回空串（不写注册表）。"""
    if getattr(sys, 'frozen', False):
        return os.path.abspath(sys.executable)
    return ''


def register_file_assoc():
    """把 .gcls 关联到本程序：默认图标 = 独立文件图标，双击 = 用本程序打开。

    只写 HKCU，不需要管理员权限。返回 (是否成功, 附加信息)。
    """
    exe = current_exe()
    if not exe:
        return False, '当前是以源码方式运行，文件关联只在打包后的 exe 中生效。'
    try:
        import winreg

        def set_val(path, name, value):
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0,
                                    winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, name, 0, winreg.REG_SZ, value)

        set_val(r'Software\Classes\%s' % FILE_EXT, '', ASSOC_PROGID)
        set_val(r'Software\Classes\%s' % FILE_EXT, 'Content Type', 'application/x-calcbook')
        set_val(r'Software\Classes\%s' % ASSOC_PROGID, '', ASSOC_DESC)
        set_val(r'Software\Classes\%s\DefaultIcon' % ASSOC_PROGID, '', file_icon_path())
        set_val(r'Software\Classes\%s\shell\open\command' % ASSOC_PROGID, '',
                '"%s" "%%1"' % exe)
        # 通知资源管理器：关联变了，刷新图标
        try:
            ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
        except Exception:
            pass
        return True, exe
    except Exception as e:
        return False, str(e)


def file_icon_path():
    """工程文件（.gcls）专用图标的稳定路径。

    onefile exe 每次运行解压到临时目录，路径会变；把随包的 calcbook_file.ico
    复制到 %LOCALAPPDATA%\\CalcBook\\ 下做注册表 DefaultIcon 的稳定指向。
    """
    base = os.path.join(os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'),
                        'CalcBook')
    dst = os.path.join(base, 'calcbook_file.ico')
    src = os.path.join(getattr(sys, '_MEIPASS', HERE), 'calcbook_file.ico')
    try:
        if os.path.exists(src) and (not os.path.exists(dst) or
                                    os.path.getsize(dst) != os.path.getsize(src)):
            os.makedirs(base, exist_ok=True)
            shutil.copyfile(src, dst)
    except Exception:
        pass
    return dst


def assoc_needs_update():
    """注册表里的图标指向是否与当前 exe 不一致（换目录后需要重新注册）。"""
    exe = current_exe()
    if not exe:
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r'Software\Classes\%s\DefaultIcon' % ASSOC_PROGID) as k:
            val, _ = winreg.QueryValueEx(k, '')
        return os.path.normcase(val.split(',')[0]) != os.path.normcase(file_icon_path())
    except Exception:
        return True

DEFAULT_PAPER = {'name': 'A4', 'w': 210.0, 'h': 297.0}   # 单位 mm

# 行底色快捷色板（右键行号“本行底色”用）
ROW_FILL_PRESETS = [('黄色', '#FFF2A8'), ('绿色', '#DDF3D6'), ('蓝色', '#DCE9F7'),
                    ('粉色', '#FADEE6'), ('灰色', '#ECECEC')]


def make_settings():
    """用户设置对象。

    自动化测试 / 沙箱环境可用环境变量 CALCBOOK_SETTINGS_FILE 指定一个独立文件，
    免得脚本把用户真实的偏好（面板宽度、上次工程、最近文件列表）写脏。
    """
    path = (os.environ.get('CALCBOOK_SETTINGS_FILE') or '').strip()
    if path:
        try:
            return QSettings(path, QSettings.IniFormat)
        except Exception:
            pass
    return QSettings('CalcBook', APP_NAME)


def screen_dpi():
    """屏幕逻辑 DPI（Qt 富文本排版用的 px 基准），取不到时按 96 处理。"""
    try:
        scr = QApplication.primaryScreen()
        dpi = float(scr.logicalDotsPerInch()) if scr is not None else 96.0
        return dpi if dpi > 0 else 96.0
    except Exception:
        return 96.0


def paper_of(report):
    """从报表设置里取纸张（mm），兼容旧工程（无 paper 或为字符串）。"""
    p = (report or {}).get('paper')
    if isinstance(p, dict):
        try:
            return {'name': str(p.get('name') or 'A4'),
                    'w': float(p.get('w') or 210.0),
                    'h': float(p.get('h') or 297.0)}
        except (TypeError, ValueError):
            return dict(DEFAULT_PAPER)
    if isinstance(p, str) and p.strip():
        return paper_by_name(p.strip())
    return dict(DEFAULT_PAPER)


def paper_by_name(name):
    """按名称取纸张尺寸（mm），未知名称回退 A4。"""
    key = (name or '').upper().replace(' ', '')
    table = {
        'A3': (297, 420), 'A4': (210, 297), 'A5': (148, 210), 'A6': (105, 148),
        'B4': (250, 353), 'B5': (176, 250), 'B6': (125, 176),
        'LETTER': (215.9, 279.4), 'LEGAL': (215.9, 355.6),
        'TABLOID': (279.4, 431.8), '16K': (184, 260), '8K': (260, 368),
        'C5': (162, 229), 'DL': (110, 220),
    }
    for k, (w, h) in table.items():
        if key.startswith(k):
            return {'name': name, 'w': float(w), 'h': float(h)}
    return dict(DEFAULT_PAPER, name=name or 'A4')


def system_paper_list():
    """系统打印机支持的纸张列表：[(名称, 宽mm, 高mm), ...]，按默认打印机优先去重。"""
    items, seen = [], set()
    try:
        from PyQt5.QtPrintSupport import QPrinterInfo
        from PyQt5.QtGui import QPageSize
        printers = []
        d = QPrinterInfo.defaultPrinter()
        if d is not None and d.printerName():
            printers.append(d)
        printers.extend(QPrinterInfo.availablePrinters())
        for pi in printers:
            try:
                sizes = pi.supportedPageSizes()
            except Exception:
                sizes = []
            for ps in sizes:
                try:
                    sz = ps.size(QPageSize.Millimeter)
                    name = ps.name() or '%0.0f×%0.0f' % (sz.width(), sz.height())
                    w, h = round(sz.width(), 1), round(sz.height(), 1)
                except Exception:
                    continue
                if w <= 0 or h <= 0:
                    continue
                key = (w, h)
                if key in seen:
                    continue
                seen.add(key)
                items.append((name, w, h))
            if items:
                break       # 取默认打印机支持的纸张即可，避免列表过长
    except Exception:
        items = []
    if not items:
        items = [('A4', 210.0, 297.0), ('A3', 297.0, 420.0),
                 ('B5', 176.0, 250.0), ('Letter', 215.9, 279.4)]
    # 去掉驱动里的“自定义大小”等占位项，常用纸型排前面
    items = [x for x in items if '自定义' not in x[0] and 'custom' not in x[0].lower()]
    ORDER = ['A4', 'A3', 'A5', 'A6', 'B5', 'B4', '16K', '8K', 'LETTER', 'LEGAL']

    def rank(item):
        up = item[0].upper()
        for i, k in enumerate(ORDER):
            if up.startswith(k):
                return i
        return len(ORDER)
    items = sorted(items, key=rank)      # 稳定排序，同级别保持驱动原顺序
    return items

# 默认列（实际以 self.project.columns 为准，此处仅作参考/回退）
COLUMNS = [
    ('code', '编号', 120),
    ('name', '项目名称', 200),
    ('part', '部位', 130),
    ('formula', '计算公式', 480),
    ('unit', '单位', 60),
    ('value', '工程量', 100),
    ('remark', '备注', 160),
    ('no_count', '不计', 50),
]


def fmt_num(v, prec=3):
    """按位数格式化工程量：四舍五入（half-up）、带千分位；None 显示空。"""
    if v is None:
        return ''
    try:
        q = Decimal(1).scaleb(-int(prec))
        d = Decimal(str(float(v))).quantize(q, rounding=ROUND_HALF_UP)
        if d == 0:
            d = abs(d)          # 避免 -0.00
    except Exception:
        return str(v)
    return format(d, ',.%df' % int(prec))


# ===== 工程符号 =====
# 钢筋等级按建筑行业习惯写法：一级 Φ（HPB300，圆圈+竖线），
# 二级/三级/四级在 Φ 右上角标 2/3/4（Φ²、Φ³、Φ⁴），
# 图纸原符号是圆圈+竖线、按等级在下方加 1/2/3 条横线（面板里用图形显示）。
REBAR_GRADES = [
    ('Φ', '一级钢 HPB300（圆圈加一竖）', 1),
    ('Φ²', '二级钢 HRB335（Φ 右上角 2，图形为一竖下加一横）', 2),
    ('Φ³', '三级钢 HRB400（Φ 右上角 3，图形为一竖下加两横）', 3),
    ('Φ⁴', '四级钢 HRB500（Φ 右上角 4，图形为一竖下加三横）', 4),
]

ENG_SYMBOLS = [
    ('×', '乘号'), ('÷', '除号'), ('π', '圆周率 3.14159'), ('²', '平方'),
    ('³', '立方'), ('@', '间距/常量'), ('±', '正负'), ('㎡', '平方米'),
    ('≤', '小于等于'), ('≥', '大于等于'), ('°', '度'), ('′', '分'),
    ('″', '秒'), ('∠', '角度'), ('⊥', '垂直'),
    ('√', '开平方：√9、√(3+4) 可直接计算'),
    ('D', '直径 D'), ('R', '半径 R'),
]


def rebar_pixmap(level, size=28, color='#1f3864'):
    """画钢筋等级符号：圆圈 + 竖线，等级 2/3/4 在下方依次加 1/2/3 条横线。"""
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    pen = QPen(QColor(color))
    pen.setWidthF(max(1.2, size / 15.0))
    p.setPen(pen)
    d = size * 0.46
    rect = QRectF((size - d) / 2.0, size * 0.06, d, d)
    p.drawEllipse(rect)
    p.drawLine(QPointF(size / 2.0, rect.top()), QPointF(size / 2.0, rect.bottom()))
    n = max(0, level - 1)
    for i in range(n):
        y = rect.bottom() + size * 0.09 + i * size * 0.115
        p.drawLine(QPointF(size * 0.24, y), QPointF(size * 0.76, y))
    p.end()
    return pm


from PyQt5.QtCore import pyqtSignal

class FormulaEdit(QTextEdit):
    """公式输入框：文字实时同步到表格；Ctrl+Enter 写入当前行。"""
    text_changed = pyqtSignal(str)

    def __init__(self, on_apply=None):
        super().__init__()
        self.on_apply = on_apply
        self._ignore = False
        self.textChanged.connect(self._on_text_changed)

    def _on_text_changed(self):
        if not self._ignore:
            self.text_changed.emit(self.toPlainText())

    def setTextSilent(self, text):
        """静默设置文字（不触发回写信号）。"""
        self._ignore = True
        self.setPlainText(text)
        self._ignore = False

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.InsertParagraphSeparator) and (event.modifiers() & Qt.ControlModifier):
            if self.on_apply:
                self.on_apply()
            return
        if event.key() == Qt.Key_Return and (event.modifiers() & Qt.ControlModifier):
            if self.on_apply:
                self.on_apply()
            return
        super().keyPressEvent(event)


class PasteTableWidget(QTableWidget):
    """支持从 WPS/Excel 多行粘贴（TSV）的表格。

    V2.10.0 起还承担：填充柄拖动复制（自动序号）、格式刷鼠标交互。
    """

    def __init__(self, on_paste=None, on_insert=None, on_delete=None):
        super().__init__()
        self.on_paste = on_paste
        self.on_insert = on_insert
        self.on_delete = on_delete
        self.on_resize = None       # 窗口尺寸变化 → 主窗口重排列宽
        self.on_copy = None         # Ctrl+C → 主窗口按选中区域复制
        self.on_cut = None          # Ctrl+X → 主窗口剪切行
        self.on_escape = None       # Esc → 取消格式刷 / 填充预览
        self.on_fill_apply = None   # 填充柄松开 → (源rect, 目标rect)
        self.on_paint_apply = None  # 格式刷松开 → 目标rect
        # 填充柄 / 格式刷运行态
        self._fill_drag = None      # {'src':(r0,r1,c0,c1),'start':QPoint,'tgt':rect|None}
        self._paint_drag = None     # {'anchor':(r,c),'cur':(r,c)}
        self.paint_mode = False     # 格式刷已拾取、等待点击目标
        self.setMouseTracking(True)

    # ---------- 几何辅助 ----------
    def sel_rect(self):
        """当前选择的包围盒（连续时才有意义）：(r0,r1,c0,c1)；空返回 None。"""
        idxs = self.selectedIndexes()
        if not idxs:
            return None
        rows = {i.row() for i in idxs}
        cols = {i.column() for i in idxs}
        r0, r1 = min(rows), max(rows)
        c0, c1 = min(cols), max(cols)
        # 修正：c1 可能超出当前列数（全选角）——裁掉
        c1 = min(c1, self.columnCount() - 1)
        if len(idxs) != (r1 - r0 + 1) * (c1 - c0 + 1):
            return None             # 非矩形选择不支持填充
        return (r0, r1, c0, c1)

    def _handle_rect(self):
        """填充柄（选择区右下角 7×7px）。"""
        rect = self.sel_rect()
        if rect is None:
            return None
        r1, c1 = rect[1], rect[3]
        if r1 >= self.rowCount() or c1 >= self.columnCount():
            return None
        cell = self.visualRect(self.model().index(r1, c1))
        if cell.width() <= 0 or cell.height() <= 0:
            return None
        return QRect(cell.right() - 4, cell.bottom() - 4, 8, 8)

    def _near_handle(self, pos):
        hr = self._handle_rect()
        if hr is None:
            return False
        return QRect(hr.x() - 5, hr.y() - 5, hr.width() + 10,
                     hr.height() + 10).contains(pos)

    def _extend_target(self, pos):
        """根据拖动位置算填充目标 rect（不含源区）。方向按位移大小自动判。"""
        src = self._fill_drag['src']
        r0, r1, c0, c1 = src
        idx = self.indexAt(pos)
        if not idx.isValid():
            return None
        tr, tc = idx.row(), idx.column()
        sp = self._fill_drag['start']
        dx, dy = pos.x() - sp.x(), pos.y() - sp.y()
        if abs(dy) >= abs(dx):
            if dy >= 0:
                if tr <= r1:
                    return None
                return (r1 + 1, tr, c0, c1)          # 向下
            if tr >= r0:
                return None
            return (tr, r0 - 1, c0, c1)              # 向上
        else:
            if dx >= 0:
                if tc <= c1:
                    return None
                return (r0, r1, c1 + 1, tc)          # 向右
            if tc >= c0:
                return None
            return (r0, r1, tc, c0 - 1)              # 向左

    @staticmethod
    def _union(src, tgt):
        r0 = min(src[0], tgt[0]); r1 = max(src[1], tgt[1])
        c0 = min(src[2], tgt[2]); c1 = max(src[3], tgt[3])
        return (r0, r1, c0, c1)

    def _rect_region(self, rect):
        """rect → QItemSelectionModel 选择（填充/刷子预览用）。"""
        r0, r1, c0, c1 = rect
        tl = self.model().index(r0, c0)
        br = self.model().index(r1, c1)
        sel = QItemSelection(tl, br)
        self.selectionModel().select(sel, QItemSelectionModel.ClearAndSelect)

    # ---------- 事件 ----------
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            pos = event.pos()
            if self.paint_mode:
                idx = self.indexAt(pos)
                if idx.isValid():
                    self._paint_drag = {'anchor': (idx.row(), idx.column()),
                                        'cur': (idx.row(), idx.column())}
                    event.accept()
                    return
            elif self._near_handle(pos):
                rect = self.sel_rect()
                if rect is not None:
                    self._fill_drag = {'src': rect, 'start': pos, 'tgt': None}
                    event.accept()
                    return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        pos = event.pos()
        if self._fill_drag is not None:
            self._fill_drag['tgt'] = self._extend_target(pos)
            self.viewport().update()
            event.accept()
            return
        if self._paint_drag is not None:
            idx = self.indexAt(pos)
            if idx.isValid():
                self._paint_drag['cur'] = (idx.row(), idx.column())
                a, cur = self._paint_drag['anchor'], self._paint_drag['cur']
                self._rect_region((min(a[0], cur[0]), max(a[0], cur[0]),
                                   min(a[1], cur[1]), max(a[1], cur[1])))
            event.accept()
            return
        # 悬停光标
        if self.paint_mode or self._near_handle(pos):
            self.setCursor(Qt.CrossCursor)
        else:
            self.setCursor(Qt.ArrowCursor)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self._fill_drag is not None:
            drag, self._fill_drag = self._fill_drag, None
            self.viewport().update()
            tgt = drag.get('tgt')
            if tgt is not None and self.on_fill_apply:
                self.on_fill_apply(drag['src'], tgt)
            event.accept()
            return
        if event.button() == Qt.LeftButton and self._paint_drag is not None:
            drag, self._paint_drag = self._paint_drag, None
            a, cur = drag['anchor'], drag['cur']
            rect = (min(a[0], cur[0]), max(a[0], cur[0]),
                    min(a[1], cur[1]), max(a[1], cur[1]))
            if self.on_paint_apply:
                self.on_paint_apply(rect)
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)
        p = QPainter(self.viewport())
        # 填充拖动预览：源区 + 目标区整体虚线框
        if self._fill_drag is not None and self._fill_drag.get('tgt'):
            u = self._union(self._fill_drag['src'], self._fill_drag['tgt'])
            tl = self.visualRect(self.model().index(u[0], u[2]))
            br = self.visualRect(self.model().index(u[1], u[3]))
            if tl.width() > 0 and br.width() > 0:
                p.setPen(QPen(QColor('#2f7de0'), 1, Qt.DashLine))
                p.drawRect(tl.adjusted(0, 0, br.right() - tl.left(),
                                       br.bottom() - tl.top()))
            return
        # 平时画填充柄小方块
        hr = self._handle_rect()
        if hr is not None and not self.paint_mode and not self.state() & QAbstractItemView.EditingState:
            p.setPen(QPen(QColor('#2f7de0'), 1))
            p.setBrush(QBrush(QColor('#2f7de0')))
            p.drawRect(hr)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.on_resize is not None:
            self.on_resize()

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Paste) or (
                event.modifiers() & Qt.ControlModifier and event.key() == Qt.Key_V):
            text = QApplication.clipboard().text()
            # 所有粘贴（含单个单元格文本）都走自定义粘贴，锚点统一由主窗口控制，
            # 避免走 QTableWidget 默认粘贴而错位到编码列。编辑中的单元格不走这里。
            if text and self.on_paste:
                self.on_paste(text)
                return
        if event.matches(QKeySequence.Copy) or (
                event.modifiers() & Qt.ControlModifier and event.key() == Qt.Key_C):
            if self.on_copy and not self.state() & QAbstractItemView.EditingState:
                self.on_copy()
                return
        if event.matches(QKeySequence.Cut) or (
                event.modifiers() & Qt.ControlModifier and event.key() == Qt.Key_X):
            if self.on_cut and not self.state() & QAbstractItemView.EditingState:
                self.on_cut()
                return
        if event.key() == Qt.Key_Escape:
            if self.on_escape and (self.paint_mode or self._fill_drag):
                self.on_escape()
                return
        # Insert 在下方插入行，Delete 删除选中内容（整行/整列/单元格由主窗口判断）
        if event.key() == Qt.Key_Insert:
            if self.on_insert:
                self.on_insert()
            return
        if event.key() == Qt.Key_Delete and not (event.modifiers() & Qt.ControlModifier):
            if self.on_delete:
                self.on_delete()
            return
        super().keyPressEvent(event)


from PyQt5.QtGui import QKeySequence


class LiveSyncDelegate(QStyledItemDelegate):
    """单元格委托：

    * 公式列：进入编辑时把编辑器里的每一次按键实时同步到公式辅助框
      （QTableWidget 的 itemChanged 只在编辑结束后才发出，做不到"边输入边显示"）；
      同时单元格自身用与辅助框完全相同的配色绘制（数字红、符号黑、注释绿…），
      方便对照检查；
    * 单位列：给出常用工程计量单位下拉框（可编辑，允许自定义单位）；
    * 不计列：勾选框画在单元格正中间，点格内任意位置即可切换。
    """

    def __init__(self, win):
        super().__init__(win.table)
        self.win = win

    # ---------- 编辑 ----------
    def createEditor(self, parent, option, index):
        key = self.win.col_key(index.column())
        if key == 'no_count':
            return None                 # 勾选框不需要文本编辑器
        if key == 'unit':
            cb = QComboBox(parent)
            cb.setEditable(True)
            cb.setInsertPolicy(QComboBox.NoInsert)
            cb.addItems(UNIT_LIST)
            cb.setMaxVisibleItems(18)
            cb.view().setMinimumWidth(130)
            cb.setToolTip('选择常用单位，或直接输入自定义单位\n'
                          'm/m2/m3/kg 保留2位、t 保留3位、个/块/套/处/台/组 取整')
            # activated 在 popup 的鼠标按下处理过程中发出：若此时同步 closeEditor
            # 销毁弹窗，Qt 后续对 popup 的处理会落空，表现为“选项点不中”。
            # 延时到事件循环尾部再提交/关闭，popup 先正常收完这一次点击。
            state = {'done': False}

            def _delayed_commit(ed=cb, st=state):
                if st['done']:
                    return
                st['done'] = True
                try:
                    self._commit(ed)
                except RuntimeError:
                    pass    # 编辑器可能已被 QTableWidget 正常关闭

            cb.activated[int].connect(lambda _i: QTimer.singleShot(0, _delayed_commit))
            return cb
        editor = super().createEditor(parent, option, index)
        if isinstance(editor, QLineEdit) and key == 'formula':
            self.win._cell_editor = editor
            editor.textChanged.connect(self.win.on_cell_editor_text)
            editor.destroyed.connect(self.win._on_editor_destroyed)
        return editor

    def _commit(self, editor):
        """下拉框选中一项后立刻写回单元格并结束编辑。"""
        self.commitData.emit(editor)
        self.closeEditor.emit(editor, QAbstractItemDelegate.NoHint)

    def setEditorData(self, editor, index):
        if isinstance(editor, QComboBox):
            editor.setCurrentText(str(index.data(Qt.EditRole) or ''))
            # 不做 selectAll：全选状态下点击弹窗选项在部分 Qt/样式组合下会失效，
            # 让行编辑保持普通光标即可，弹窗点选与手动输入都正常。
            return
        super().setEditorData(editor, index)

    def setModelData(self, editor, model, index):
        if isinstance(editor, QComboBox):
            model.setData(index, editor.currentText().strip(), Qt.EditRole)
            return
        super().setModelData(editor, model, index)

    # ---------- 绘制 ----------
    def _hilite_color(self, index):
        """当前行 / 列高亮底色：返回 (QColor, alpha) 或 None。不改选择模式，只叠加淡色底纹。"""
        win = getattr(self, 'win', None)
        if win is None:
            return None
        table = getattr(win, 'table', None)
        if table is None or not index.isValid():
            return None
        # 选中态优先，避免与系统选中蓝底冲突
        try:
            if index in table.selectionModel().selectedIndexes():
                return None
        except Exception:
            return None
        try:
            cur = table.currentIndex()
            row_hit = bool(getattr(win, '_hilite_row', False)
                           and cur.isValid() and index.row() == cur.row())
            col_hit = bool(getattr(win, '_hilite_col', False)
                           and cur.isValid() and index.column() == cur.column())
        except Exception:
            return None
        if not (row_hit or col_hit):
            return None
        if row_hit and col_hit:
            return (QColor('#2f7de0'), 75)       # 行列交叉：稍深
        return (QColor('#2f7de0'), 40)           # 当前行 / 当前列：同一颜色

    def _overlay_hilite(self, painter, option, index):
        """在单元格绘制完成后叠加半透明高亮色块（不动 QStyle 内部状态，最安全）。"""
        h = self._hilite_color(index)
        if h is None:
            return
        color, alpha = h
        painter.save()
        try:
            c = QColor(color)
            c.setAlpha(alpha)
            painter.fillRect(option.rect, c)
        finally:
            painter.restore()

    def paint(self, painter, option, index):
        key = self.win.col_key(index.column())
        if key not in ('no_count', 'formula'):
            super().paint(painter, option, index)
            try:
                self._overlay_hilite(painter, option, index)
            except Exception:
                pass
            return
        try:
            opt = QStyleOptionViewItem(option)
            self.initStyleOption(opt, index)
            text = opt.text or ''
            opt.text = ''
            if key == 'no_count':
                opt.features &= ~QStyleOptionViewItem.HasCheckIndicator
            style = opt.widget.style() if opt.widget is not None else QApplication.style()
            # 先画背景 / 选中态 / 网格（不带文字与勾选框）
            painter.save()
            style.drawControl(QStyle.CE_ItemViewItem, opt, painter, opt.widget)
            painter.restore()
            if key == 'no_count':
                self._paint_check(painter, opt, index, style)
            else:
                self._paint_formula(painter, opt, text, index)
            self._overlay_hilite(painter, opt, index)
        except Exception:
            # 槽函数里未捕获异常会让 PyQt5 直接 qFatal（表现为 exe 闪退）
            try:
                log_exception('单元格绘制')
                super().paint(painter, option, index)
            except Exception:
                pass

    def sizeHint(self, option, index):
        """公式列按当前列宽折行后的真实高度，resizeRowsToContents() 用它撑高行。"""
        if self.win.col_key(index.column()) != 'formula':
            return super().sizeHint(option, index)
        base = super().sizeHint(option, index)
        try:
            opt = QStyleOptionViewItem(option)
            self.initStyleOption(opt, index)
            font = QFont(self.win.table.font())
            lines = self._wrap_formula(opt.text or '', max(40, option.rect.width() - 9), font)
            need = len(lines) * QFontMetrics(font).height() + 8
            base.setHeight(int(max(30, min(300, need))))
        except Exception:
            log_exception('公式行高计算')
        return base

    # ---------- 公式折行排版（颜色分词 + 自动换行） ----------
    @staticmethod
    def _char_kinds(text):
        """逐字颜色类别。compute_runs 会跳过空格/字母等无类别字符，这里补齐为 None。"""
        kinds = [None] * len(text)
        for start, length, kind in compute_runs(text):
            for i in range(start, min(len(text), start + length)):
                kinds[i] = kind
        return kinds

    @staticmethod
    def _formula_tokens(text):
        """切成 (片段, 颜色类别)：中文/符号可逐字断行，ASCII 单词尽量不断开。"""
        kinds = LiveSyncDelegate._char_kinds(text)
        tokens = []
        buf, cur = '', None

        def flush():
            nonlocal buf, cur
            if buf:
                tokens.append((buf, cur))
                buf, cur = '', None

        for ch, kind in zip(text, kinds):
            if ch == '\n':
                flush()
                tokens.append(('\n', kind))
                continue
            word_char = ch.isascii() and (ch.isalnum() or ch in '._')
            if word_char and (not buf or cur == kind):
                buf += ch
                cur = kind
            else:
                flush()
                if word_char:
                    buf, cur = ch, kind
                else:
                    tokens.append((ch, kind))
        flush()
        return tokens

    def _wrap_formula(self, text, max_width, base_font):
        """按当前列宽折成多行：[[(片段, 颜色类别), ...], ...]。"""
        fonts, fms = {}, {}

        def fms_for(kind):
            bold = kind in ('op', 'bracket', 'super')
            if bold not in fonts:
                f = QFont(base_font)
                f.setBold(bold)
                fonts[bold] = f
                fms[bold] = QFontMetrics(f)
            return fonts[bold], fms[bold]

        lines, used = [[]], 0.0
        for tok, kind in self._formula_tokens(text or ''):
            if tok == '\n':
                lines.append([])
                used = 0.0
                continue
            _f, fm = fms_for(kind)
            w = float(fm.horizontalAdvance(tok))
            if used + w > max_width and lines[-1]:
                lines.append([])
                used = 0.0
            if w > max_width and len(tok) > 1:
                for ch in tok:
                    cw = float(fm.horizontalAdvance(ch))
                    if used + cw > max_width and lines[-1]:
                        lines.append([])
                        used = 0.0
                    lines[-1].append((ch, kind))
                    used += cw
            else:
                lines[-1].append((tok, kind))
                used += w
        return lines

    def _paint_check(self, painter, opt, index, style):
        """把勾选框画在单元格正中央。"""
        state = index.data(Qt.CheckStateRole)
        cb = QStyleOptionButton()
        cb.state = QStyle.State_Enabled
        cb.state |= QStyle.State_On if state == Qt.Checked else QStyle.State_Off
        if opt.state & QStyle.State_MouseOver:
            cb.state |= QStyle.State_MouseOver
        size = style.pixelMetric(QStyle.PM_IndicatorWidth, None, opt.widget)
        cb.rect = QRect(0, 0, size, size)
        cb.rect.moveCenter(opt.rect.center())
        painter.save()
        style.drawPrimitive(QStyle.PE_IndicatorCheckBox, cb, painter, opt.widget)
        painter.restore()

    def _paint_formula(self, painter, opt, text, index):
        """按当前单元格宽高自动折行绘制，配色与公式辅助框完全一致。"""
        if not text:
            return
        base = QFont(self.win.table.font())
        fm0 = QFontMetrics(base)
        rect = opt.rect.adjusted(5, 1, -3, -1)
        if rect.width() <= 4 or rect.height() <= 4:
            return
        # 有错误时整行公式标红（与原来 setForeground 的行为一致）
        fg = index.data(Qt.ForegroundRole)
        err_color = None
        if isinstance(fg, QBrush) and fg.color().name().lower() == '#c00000':
            err_color = fg.color()
        lines = self._wrap_formula(text, rect.width(), base)
        line_h = fm0.height()
        total = len(lines) * line_h
        # 行高够就垂直居中；不够则从顶部显示（超出的由 clip 裁掉）
        y = rect.top() + (max(0, rect.height() - total) // 2) + fm0.ascent()
        painter.save()
        painter.setClipRect(rect)
        fonts = {}
        for line in lines:
            x = float(rect.left())
            for seg, kind in line:
                bold = kind in ('op', 'bracket', 'super')
                f = fonts.get(bold)
                if f is None:
                    f = QFont(base)
                    f.setBold(bold)
                    fonts[bold] = f
                painter.setFont(f)
                painter.setPen(err_color if err_color is not None
                               else QColor(FCOLOR.get(kind, '#000000')))
                w = QFontMetrics(f).horizontalAdvance(seg)
                if x + w > rect.right() + 1:
                    break
                painter.drawText(int(x), int(y), seg)
                x += w
            y += line_h
        painter.restore()

    def editorEvent(self, event, model, option, index):
        # 不计列：点格内任意处切换勾选；空格键也能切换
        if self.win.col_key(index.column()) == 'no_count':
            if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
                if option.rect.contains(event.pos()):
                    cur = index.data(Qt.CheckStateRole)
                    model.setData(index, Qt.Unchecked if cur == Qt.Checked else Qt.Checked,
                                  Qt.CheckStateRole)
                    return True
            if event.type() == QEvent.KeyPress and event.key() == Qt.Key_Space:
                cur = index.data(Qt.CheckStateRole)
                model.setData(index, Qt.Unchecked if cur == Qt.Checked else Qt.Checked,
                              Qt.CheckStateRole)
                return True
        return super().editorEvent(event, model, option, index)


class FitHeader(QHeaderView):
    """表头：用户手动拖动列宽时通知主窗口关闭“自动列宽”，避免又被自动排版抢回去。"""

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self.on_manual_resize = None

    def mousePressEvent(self, event):
        if self.on_manual_resize and event.button() == Qt.LeftButton:
            try:
                if self.cursor().shape() == Qt.SplitHCursor:
                    self.on_manual_resize()
            except Exception:
                pass
        super().mousePressEvent(event)


class QuickScroll(QScrollArea):
    """快捷链接行的横向滚动容器。

    作用：链接再多也不会撑大工作区的最小宽度（否则左右面板会被挤到最窄）。
    高度按内容自适应，并在出现横向滚动条时自动让出滚动条的空间，避免文字被裁。
    """

    def __init__(self, widget, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.NoFrame)
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setMinimumWidth(180)
        self.setStyleSheet('QScrollArea { background:#eef2f7; border:none; }'
                           ' QScrollBar:horizontal { height:8px; margin:0; background:transparent; }'
                           ' QScrollBar::handle:horizontal { background:#b9c6d4; border-radius:4px; min-width:24px; }'
                           ' QScrollBar::handle:horizontal:hover { background:#8fa7bd; }'
                           ' QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width:0; }'
                           ' QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background:transparent; }')
        self._content = widget
        self.setWidget(widget)

    def _fit_height(self):
        try:
            h = self._content.sizeHint().height()
        except Exception:
            h = 24
        try:
            bar = self.horizontalScrollBar()
            if bar is not None and bar.isVisible():
                h += bar.sizeHint().height()
            h = max(24, int(h))
            if self.height() != h:
                self.setFixedHeight(h)
        except Exception:
            pass

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._fit_height()

    def showEvent(self, event):
        super().showEvent(event)
        QTimer.singleShot(0, self._fit_height)


class ConstGrid(QTableWidget):
    """常量框：三列（常量代码 | 值 | 说明），双击就地改，可自由增删。

    框上方是“新建常量 / 删除 / 导出常量”；
    V2.4 起不再内置任何常量（钢筋单重只作模板），全部由使用者自己定义。
    """

    changed = pyqtSignal()          # 用户改动了常量（代码 / 值 / 说明）

    COLS = ['常量代码', '值', '说明']

    def __init__(self, on_batch_edit=None, parent=None):
        super().__init__(0, 3, parent)
        self._on_batch = on_batch_edit
        self._loading = False
        self.setEditTriggers(QAbstractItemView.DoubleClicked |
                             QAbstractItemView.SelectedClicked |
                             QAbstractItemView.EditKeyPressed |
                             QAbstractItemView.AnyKeyPressed)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setShowGrid(True)
        self.setWordWrap(False)
        self.setAlternatingRowColors(False)
        self.setFrameShape(QFrame.NoFrame)
        self.verticalHeader().setVisible(False)
        # 行高下限 24（取值见 fit_rows）：行间距够大，点按钮新增/编辑都舒服
        self.verticalHeader().setMinimumSectionSize(22)
        self.verticalHeader().setDefaultSectionSize(26)
        self.setHorizontalHeaderLabels(self.COLS)
        hh = self.horizontalHeader()
        hh.setVisible(True)
        hh.setFixedHeight(21)
        hh.setSectionResizeMode(0, QHeaderView.Interactive)
        hh.setSectionResizeMode(1, QHeaderView.Interactive)
        hh.setSectionResizeMode(2, QHeaderView.Stretch)
        hh.setMinimumSectionSize(36)
        self.setColumnWidth(0, 92)
        self.setColumnWidth(1, 56)
        self.setStyleSheet('QTableWidget { background:#fbfcfe; border:1px solid #dfe3e8; '
                           'gridline-color:#e6eaef; }'
                           'QTableWidget::item { padding:0 2px; }'
                           'QTableWidget::item:selected { background:#dce9f7; color:#222; }'
                           'QHeaderView::section { background:#f0e6c8; color:#4a4020; '
                           'font-weight:400; padding:1px 2px; border:none; '
                           'border-right:1px solid #e0d6b6; border-bottom:1px solid #d8cfae; }')
        f = QFont('微软雅黑')
        f.setPointSizeF(9)
        self.setFont(f)
        self._fitting = False
        self.itemChanged.connect(self._item_changed)

    # ---------- 数据 ----------
    def set_constants(self, consts, notes=None):
        """把常量显示到网格（代码列带 @ 前缀，说明来自 notes）。"""
        notes = notes or {}
        self._loading = True
        try:
            self.clearSpans()
            self.clearContents()
            items = list((consts or {}).items())
            if not items:
                self.setRowCount(1)
                it = QTableWidgetItem('（无常量：双击或点“新建常量”）')
                it.setFlags(Qt.ItemIsEnabled)
                it.setForeground(QColor('#9aa4ae'))
                self.setItem(0, 0, it)
                self.setSpan(0, 0, 1, 3)
            else:
                self.setRowCount(len(items))
                for r, (k, v) in enumerate(items):
                    self._fill_row(r, k, fmt_const(v), notes.get(k, ''))
        finally:
            self._loading = False
        self.fit_rows()

    def _fill_row(self, r, name, value, note=''):
        it_n = QTableWidgetItem(str(name))
        it_n.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
        it_n.setForeground(QColor('#24507a'))
        it_n.setToolTip('常量代码（公式里写 @代码 即可参与计算）')
        it_v = QTableWidgetItem(str(value))
        it_v.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
        it_v.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        it_v.setToolTip('常量数值（只能是数字）')
        it_t = QTableWidgetItem(str(note or ''))
        it_t.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
        it_t.setForeground(QColor('#666666'))
        it_t.setToolTip(str(note or '') or '说明（只作备注，不参与计算）')
        self.setItem(r, 0, it_n)
        self.setItem(r, 1, it_v)
        self.setItem(r, 2, it_t)

    def rows(self):
        """读出全部行：[(代码, 数值文本, 说明), ...]，保持界面顺序。"""
        out = []
        for r in range(self.rowCount()):
            n = self.item(r, 0)
            if n is None:
                continue
            name = (n.text() or '').strip()
            if not name or name.startswith('（'):
                continue
            if not name.startswith('@'):
                name = '@' + name
            v = self.item(r, 1)
            txt = ((v.text() if v else '') or '').strip()
            note = (self.item(r, 2).text() if self.item(r, 2) else '') or ''
            out.append((name, txt, note.strip()))
        return out

    def constants(self):
        """读出 {代码: 数值}（非法数值跳过并在单元格标红）。"""
        out = {}
        for r in range(self.rowCount()):
            n = self.item(r, 0)
            v = self.item(r, 1)
            if n is None or v is None:
                continue
            name = (n.text() or '').strip()
            if not name or name.startswith('（'):
                continue
            if not name.startswith('@'):
                name = '@' + name
            txt = (v.text() or '').strip()
            try:
                out[name] = float(txt)
            except (TypeError, ValueError):
                if txt:
                    v.setForeground(QColor('#c00000'))
                    v.setToolTip('“%s”不是数字，该常量暂不生效' % txt)
        return out

    def notes(self):
        """读出 {代码: 说明}。"""
        out = {}
        for name, _txt, note in self.rows():
            if note:
                out[name] = note
        return out

    def add_row(self, name='@', value=''):
        """在末尾加一行并直接进入代码编辑。"""
        if self.rowCount() == 1 and self.item(0, 0) is not None \
                and self.item(0, 0).text().startswith('（'):
            self.clearSpans()
            self.setRowCount(0)
        self._loading = True
        r = self.rowCount()
        self.setRowCount(r + 1)
        self._fill_row(r, name, value)
        self._loading = False
        self.fit_rows()
        self.setCurrentCell(r, 0)
        if self.item(r, 0) is not None:
            self.scrollToItem(self.item(r, 0), QAbstractItemView.PositionAtCenter)
        self.editItem(self.item(r, 0))
        return r

    def remove_current_row(self):
        r = self.currentRow()
        if r < 0:
            return False
        self._loading = True
        self.removeRow(r)
        self._loading = False
        self.fit_rows()
        self.changed.emit()
        return True

    # ---------- 交互 ----------
    def _item_changed(self, item):
        if self._loading:
            return
        col, text = item.column(), (item.text() or '').strip()
        if col == 0:
            # 代码：自动补 @ 前缀
            if text and not text.startswith('@') and not text.startswith('（'):
                self._loading = True
                item.setText('@' + text)
                self._loading = False
            item.setToolTip('常量代码（公式里写 @代码 即可参与计算）')
        elif col == 1:
            try:
                float(text)
                item.setForeground(QColor('#000000'))
                item.setToolTip('常量数值（只能是数字）')
            except (TypeError, ValueError):
                item.setForeground(QColor('#c00000'))
                item.setToolTip('只能填数字')
        self.changed.emit()

    def mouseDoubleClickEvent(self, event):
        idx = self.indexAt(event.pos())
        if idx.isValid():
            super().mouseDoubleClickEvent(event)
            return
        # 双击空白区域 → 批量编辑
        if self._on_batch is not None:
            self._on_batch()

    def contextMenuEvent(self, event):
        idx = self.indexAt(event.pos())
        m = QMenu(self)
        m.addAction('新建常量', lambda: self.add_row())
        if idx.isValid():
            self.setCurrentCell(idx.row(), idx.column())
            m.addAction('编辑本行', lambda: self.editItem(self.item(idx.row(), idx.column())))
            m.addAction('删除本行', self.remove_current_row)
        m.addSeparator()
        m.addAction('批量编辑 / 粘贴…', lambda: self._on_batch() if self._on_batch else None)
        m.exec_(event.globalPos())

    # ---------- 布局 ----------
    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.fit_rows()

    def fit_rows(self):
        """按框体尺寸分配行高（框体可拖动改变大小时不必重新排版）。

        行高下限 24px：以前为了“不出现滚动条”把行压到 15px，
        输入时常量行挤在一起很难点，V2.5 起宁可出竖向滚动条也保证每行够高。
        """
        rows = self.rowCount()
        if self._fitting:
            return
        self._fitting = True
        try:
            # 列宽：代码 45%、值 22%，其余给说明（窗口窄时也保证“值”不被截断）
            avail = self.viewport().width()
            if avail > 80:
                w0 = max(64, min(150, int(avail * 0.45)))
                w1 = max(48, min(90, int(avail * 0.22)))
                if self.columnWidth(0) != w0:
                    self.setColumnWidth(0, w0)
                if self.columnWidth(1) != w1:
                    self.setColumnWidth(1, w1)
            if rows <= 0:
                return
            h = self.viewport().height()
            row_h = max(24, min(34, (h - 2) // rows)) if h > 0 else 26
            for r in range(rows):
                if self.rowHeight(r) != row_h:
                    self.setRowHeight(r, row_h)
        finally:
            self._fitting = False


class FoldStrip(QFrame):
    """两块面板之间的窄条（充当分割条），条上放两个小箭头按钮：

    一个收起相邻面板，一个展开——点一下工作区就变大 / 复原。
    """

    def __init__(self, vertical=True, fold_text='◀', open_text='▶',
                 fold_tip='收起面板', open_tip='展开面板', parent=None):
        super().__init__(parent)
        self.vertical = vertical
        self.setObjectName('foldStrip')
        if vertical:
            self.setFixedWidth(20)
            lay = QVBoxLayout(self)
            self.setCursor(Qt.ArrowCursor)
        else:
            self.setFixedHeight(19)
            lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 2, 0, 2)
        lay.setSpacing(1)
        lay.addStretch(1)
        self.btn_fold = self._mk(fold_text, fold_tip)
        self.btn_open = self._mk(open_text, open_tip)
        lay.addWidget(self.btn_fold)
        lay.addWidget(self.btn_open)
        lay.addStretch(1)
        self.setStyleSheet('#foldStrip { background:#eef2f7; border:none; '
                           'border-left:1px solid #e0e6ee; }')
        self.set_folded(False)

    def _mk(self, text, tip):
        b = QToolButton(self)
        b.setText(text)
        b.setToolTip(tip)
        b.setCursor(Qt.PointingHandCursor)
        b.setAutoRaise(True)
        if self.vertical:
            b.setFixedSize(19, 24)
        else:
            b.setFixedSize(28, 16)
        b.setStyleSheet(
            'QToolButton { border:1px solid #c9d6e3; border-radius:2px; background:#ffffff;'
            ' color:#3c6b9c; font-size:9pt; padding:0; }'
            ' QToolButton:hover { background:#dceafa; border-color:#8fb3d6; }'
            ' QToolButton:pressed { background:#c9dcf2; }'
            ' QToolButton:disabled { background:#f2f5f8; border-color:#e3e9f0; color:#c2cbd5; }')
        return b

    def set_folded(self, folded):
        """收起状态：只亮“展开”按钮，避免误点。"""
        self.btn_fold.setEnabled(not folded)
        self.btn_open.setEnabled(bool(folded))


def natural_key(text):
    """编号排序用的自然序（2 排在 10 前面）。"""
    return [int(t) if t.isdigit() else t for t in re.split(r'(\d+)', str(text or ''))]


class SummaryPanel(QFrame):
    """右侧分类汇总面板。

    可按 编号 / 项目名称 / 部位 / 分部 / 单位 把计算数据分类合计，
    结果可直接导出 Excel 或复制到剪贴板。
    """

    MODES = [('code', '按编号'), ('name', '按项目名称'), ('part', '按部位'),
             ('section', '按分部'), ('unit', '按单位')]
    MODE_COL = {'code': '编号', 'name': '项目名称', 'part': '部位',
                'section': '分部', 'unit': '单位'}

    def __init__(self, win, parent=None):
        super().__init__(parent)
        self.win = win
        self.setObjectName('summaryPanel')
        self.setMinimumWidth(230)
        self.setStyleSheet('#summaryPanel { background:#fdfcf8; border-left:1px solid #dfe3e8; }'
                           '#summaryPanel QTableWidget { background:#ffffff; }')
        self._groups = []
        self._total = (0.0, 0)
        self._units = set()
        # 汇总表列宽按表头文字记住，刷新 / 重开后仍保持用户拖出来的宽度
        self._col_widths = dict(win.settings.value('summary_col_widths', {}, type=dict) or {})

        v = QVBoxLayout(self)
        v.setContentsMargins(5, 4, 5, 4)
        v.setSpacing(4)

        head = QHBoxLayout()
        head.setSpacing(4)
        lb = QLabel('分类汇总')
        lb.setStyleSheet('font-weight:600; color:#4a4020;')
        head.addWidget(lb)
        head.addStretch(1)
        self.lbl_info = QLabel('')
        self.lbl_info.setStyleSheet('color:#8a8a8a; font-size:9pt;')
        head.addWidget(self.lbl_info)
        # 分隔改成 1px 细线后，收起按钮挪到面板标题行
        self.btn_fold = QToolButton(self)
        self.btn_fold.setText('▶')
        self.btn_fold.setToolTip('收起分类汇总面板（工作区变宽）')
        self.btn_fold.setCursor(Qt.PointingHandCursor)
        self.btn_fold.setAutoRaise(True)
        self.btn_fold.setFixedSize(26, 26)
        self.btn_fold.setStyleSheet('QToolButton { border:1px solid #d5dbe2; border-radius:2px; '
                                    'background:#fff; color:#3c6b9c; font-size:12pt; padding:0; }'
                                    'QToolButton:hover { background:#e8f0fa; border-color:#9dbce0; }')
        self.btn_fold.clicked.connect(lambda: self.win.fold_summary(True))
        head.addWidget(self.btn_fold)
        # 放大汇总：隐藏主工作区，让汇总表占满窗口；再点还原
        self.btn_max = QToolButton(self)
        self.btn_max.setText('⤢')
        self.btn_max.setToolTip('放大分类汇总（隐藏主工作区）')
        self.btn_max.setCursor(Qt.PointingHandCursor)
        self.btn_max.setAutoRaise(True)
        self.btn_max.setFixedSize(26, 26)
        self.btn_max.setStyleSheet(self.btn_fold.styleSheet())
        self.btn_max.clicked.connect(lambda: self.win.toggle_summary_max())
        head.addWidget(self.btn_max)
        v.addLayout(head)

        mode_row = QHBoxLayout()
        mode_row.setSpacing(4)
        self.mode = QComboBox()
        for key, label in self.MODES:
            self.mode.addItem(label, key)
        self.mode.setToolTip('按哪个字段把计算数据分类汇总（单条件模式）')
        mode_row.addWidget(self.mode, 1)
        self.btn_multi = QPushButton('多条件…')
        self.btn_multi.setToolTip('多字段分组 + 多条筛选条件')
        self.btn_multi.setStyleSheet('QPushButton { padding:3px 6px; font-size:9pt; }')
        self.btn_multi.clicked.connect(self.open_multi_dialog)
        mode_row.addWidget(self.btn_multi)
        self.btn_clear_multi = QPushButton('单条件')
        self.btn_clear_multi.setToolTip('清除多条件，恢复按上方方式汇总')
        self.btn_clear_multi.setStyleSheet('QPushButton { padding:3px 6px; font-size:9pt; }')
        self.btn_clear_multi.setVisible(False)
        self.btn_clear_multi.clicked.connect(self.clear_multi)
        mode_row.addWidget(self.btn_clear_multi)
        v.addLayout(mode_row)
        self.multi = None                 # None=单条件；dict=多字段/多筛选

        self.only_cur = QCheckBox('仅当前分部（含下级）')
        self.only_cur.setToolTip('勾选后只汇总左侧选中的分部；不勾选汇总全部计算书')
        v.addWidget(self.only_cur)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(['分类', '单位', '合计', '行数'])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # 与工作区表格一致：可多选单元格、可复制、可调列宽行高
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        # 表格的最小宽度：不设的话 QTableWidget 会把“所有列宽之和”当最小宽度，
        # 从而把左右面板挤到最窄（列宽本来就是自适应的，滚动条即可）
        self.table.setMinimumWidth(240)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.setStyleSheet('QTableWidget::item { padding:2px 3px; }')
        hh = self.table.horizontalHeader()
        # 所有列都允许拖动宽度（原来第 0 列是 Stretch，项目名称列拖不动）
        for i, w in ((0, 160), (1, 42), (2, 84), (3, 40)):
            hh.setSectionResizeMode(i, QHeaderView.Interactive)
            title = self.table.horizontalHeaderItem(i).text()
            self.table.setColumnWidth(i, int(self._col_widths.get(title, w)))
        hh.sectionResized.connect(self._on_col_resized)
        self.table.verticalHeader().setDefaultSectionSize(24)
        self._setup_table_ops()
        v.addWidget(self.table, 1)

        bar = QHBoxLayout()
        bar.setSpacing(4)
        b_ref = QPushButton('刷新')
        b_ref.clicked.connect(self.refresh)
        b_exp = QPushButton('导出')
        b_exp.setToolTip('把当前汇总结果导出为 Excel 表格')
        b_exp.clicked.connect(self.export_excel)
        b_cp = QPushButton('复制')
        b_cp.setToolTip('复制汇总表（可直接粘贴到 WPS / Excel）')
        b_cp.clicked.connect(self.copy_to_clipboard)
        self.btn_export, self.btn_copy = b_exp, b_cp
        for b in (b_ref, b_exp, b_cp):
            b.setStyleSheet('QPushButton { padding:3px 6px; font-size:9pt; }')
            bar.addWidget(b)
        v.addLayout(bar)

        self.mode.currentIndexChanged.connect(lambda _i: self.clear_multi())
        self.only_cur.stateChanged.connect(lambda _s: self.refresh())

    # ---------------- 汇总表的行列操作（与工作区表格一致） ----------------
    def _setup_table_ops(self):
        t = self.table
        t.setContextMenuPolicy(Qt.CustomContextMenu)
        t.customContextMenuRequested.connect(self.table_menu)
        hh = t.horizontalHeader()
        hh.setContextMenuPolicy(Qt.CustomContextMenu)
        hh.customContextMenuRequested.connect(self.header_menu)
        # 双击列分隔线：该列列宽自适应当前内容
        hh.sectionHandleDoubleClicked.connect(lambda i: self.fit_columns([i]))
        corner = t.findChild(QAbstractButton)
        if corner is not None:
            corner.setToolTip('点击全选；右键可设置列宽 / 行高自适应、复制')
            corner.setContextMenuPolicy(Qt.CustomContextMenu)
            corner.customContextMenuRequested.connect(
                lambda p, w=corner: self.corner_menu(w.mapToGlobal(p)))
        # Ctrl+C 复制选中的单元格（与工作区一致）
        sc = QShortcut(QKeySequence.Copy, t)
        sc.setContext(Qt.WidgetShortcut)
        sc.activated.connect(self.copy_selection)

    def copy_selection(self):
        """复制选中区域为 TSV（可直接粘到 WPS/Excel）。"""
        idxs = self.table.selectedIndexes()
        if not idxs:
            self.win.status.showMessage('请先选中要复制的汇总单元格', 2500)
            return
        rows = sorted({i.row() for i in idxs})
        cols = sorted({i.column() for i in idxs})
        lines = []
        for r in rows:
            vals = []
            for c in cols:
                it = self.table.item(r, c)
                vals.append(it.text() if it is not None else '')
            lines.append('\t'.join(vals))
        QApplication.clipboard().setText('\n'.join(lines))
        self.win.status.showMessage('已复制汇总表 %d 行 × %d 列' % (len(rows), len(cols)), 2500)

    def _on_col_resized(self, col, old, new):
        """用户拖动列分隔线后记住宽度，按表头文字存，刷新后仍生效。"""
        hi = self.table.horizontalHeaderItem(col)
        if hi is None or new <= 0:
            return
        self._col_widths[hi.text()] = int(new)
        try:
            self.win.settings.setValue('summary_col_widths', self._col_widths)
        except Exception:
            pass

    def fit_columns(self, cols=None):
        """按内容自动列宽（与工作区“列宽自适应”同样的口径）。"""
        hh = self.table.horizontalHeader()
        fm = QFontMetrics(self.table.font())
        measure = getattr(fm, 'horizontalAdvance', None) or fm.width
        cols = list(range(self.table.columnCount())) if cols is None else list(cols)
        for c in cols:
            if not (0 <= c < self.table.columnCount()):
                continue
            w = measure(self.table.horizontalHeaderItem(c).text() if
                        self.table.horizontalHeaderItem(c) else '') + 26
            for r in range(self.table.rowCount()):
                it = self.table.item(r, c)
                if it is not None and it.text():
                    w = max(w, measure(it.text()) + 22)
            if c == 0:
                w = max(80, min(240, w))
            else:
                w = max(40, min(140, w))
            hh.setSectionResizeMode(c, QHeaderView.Interactive)
            self.table.setColumnWidth(c, w)

    def fit_rows(self):
        """行高自适应：按内容撑高，最小 22。"""
        self.table.resizeRowsToContents()
        for r in range(self.table.rowCount()):
            self.table.setRowHeight(r, max(22, min(80, self.table.rowHeight(r))))

    def set_col_width_dialog(self, col):
        """输入像素值设置汇总表某一列列宽。"""
        hi = self.table.horizontalHeaderItem(col)
        title = hi.text() if hi is not None else '本列'
        val, ok = QInputDialog.getInt(self, '设置列宽',
                                      '「%s」列宽（像素，30–600）：' % title,
                                      self.table.columnWidth(col), 30, 600, 1)
        if ok:
            self.table.setColumnWidth(col, val)

    def set_row_height_dialog(self, row):
        """输入像素值设置汇总表某一行行高。"""
        val, ok = QInputDialog.getInt(self, '设置行高',
                                      '行高（像素，18–120）：',
                                      self.table.rowHeight(row), 18, 120, 1)
        if ok:
            self.table.setRowHeight(row, val)

    def reset_rows(self):
        for r in range(self.table.rowCount()):
            self.table.setRowHeight(r, 24)

    def header_menu(self, pos):
        h = self.table.horizontalHeader()
        c = h.logicalIndexAt(pos)
        m = QMenu(self)
        if c >= 0:
            hi = self.table.horizontalHeaderItem(c)
            title = hi.text() if hi is not None else ''
            m.addAction('设置本列列宽…「%s」' % title, lambda cc=c: self.set_col_width_dialog(cc))
            m.addAction('本列列宽自适应「%s」' % title, lambda cc=c: self.fit_columns([cc]))
        m.addAction('全部列宽自适应', lambda: self.fit_columns())
        m.addAction('全部行高自适应', self.fit_rows)
        m.addAction('恢复默认行高', self.reset_rows)
        m.addSeparator()
        m.addAction('复制选中单元格  (Ctrl+C)', self.copy_selection)
        m.addAction('复制整张汇总表', self.select_all_copy)
        m.addAction('导出为 Excel…', self.export_excel)
        m.exec_(h.mapToGlobal(pos))

    def corner_menu(self, global_pos):
        m = QMenu(self)
        m.addAction('列宽自适应（按内容）', lambda: self.fit_columns())
        m.addAction('行高自适应（按内容）', self.fit_rows)
        m.addAction('恢复默认行高', self.reset_rows)
        m.addSeparator()
        m.addAction('全选表格', self.table.selectAll)
        m.addAction('复制整张汇总表', self.select_all_copy)
        m.exec_(global_pos)

    def table_menu(self, pos):
        r = self.table.rowAt(pos.y())
        c = self.table.columnAt(pos.x())
        if r >= 0:
            self.table.setCurrentCell(r, max(c, 0))
        m = QMenu(self)
        m.addAction('复制选中单元格  (Ctrl+C)', self.copy_selection)
        m.addAction('复制整张汇总表', self.select_all_copy)
        m.addSeparator()
        if r >= 0:
            m.addAction('设置本行行高…', lambda rr=r: self.set_row_height_dialog(rr))
        if c >= 0:
            m.addAction('设置本列列宽…', lambda cc=c: self.set_col_width_dialog(cc))
        m.addAction('列宽自适应', lambda: self.fit_columns())
        m.addAction('行高自适应', self.fit_rows)
        m.addAction('恢复默认行高', self.reset_rows)
        m.addSeparator()
        m.addAction('刷新汇总', self.refresh)
        m.addAction('导出为 Excel…', self.export_excel)
        m.exec_(self.table.mapToGlobal(pos))

    def select_all_copy(self):
        """全选后复制整张表（含表头与合计行）。"""
        self.table.selectAll()
        self.copy_selection()

    # ---------------- 汇总 ----------------
    def mode_key(self):
        return self.mode.currentData() or 'code'

    def mode_label(self):
        if self.multi:
            names = [self.MODE_COL.get(f, f) for f in self.multi['fields']]
            label = '多条件（%s）' % '+'.join(names)
            if self.multi.get('filters'):
                label += '·%d筛选' % len(self.multi['filters'])
            return label
        return self.mode.currentText()

    def fields(self):
        """当前分组字段列表：单条件 1 个，多条件可多个。"""
        if self.multi:
            return list(self.multi['fields'])
        return [self.mode_key()]

    def show_unit_col(self):
        return 'unit' not in self.fields()

    def headers(self):
        names = [self.MODE_COL.get(f, f) for f in self.fields()]
        if self.show_unit_col():
            names.append('单位')
        return names + ['合计', '行数']

    @staticmethod
    def row_field(row, section, key):
        """取一行在指定字段上的显示值（与单条件模式口径完全一致）。"""
        if key == 'section':
            return section.name or '（未命名分部）'
        v = getattr(row, key, '')
        text = str(v or '').strip()
        if text:
            return text
        return {'code': '（无编号）', 'name': '（未命名）',
                'part': '（无部位）', 'unit': '（无单位）'}.get(key, '（空）')

    def pass_filters(self, row, section):
        """多条件模式的筛选：每条条件按 字段/关系/值 判断，再按 全部/任一 合并。"""
        multi = self.multi
        if not multi or not multi.get('filters'):
            return True
        results = []
        for field, op, value in multi['filters']:
            text = self.row_field(row, section, field)
            val = (value or '').strip()
            if op == 'contains':
                ok = bool(val) and val in text
            elif op == 'not_contains':
                ok = not val or val not in text
            elif op == 'eq':
                ok = text == val
            elif op == 'ne':
                ok = text != val
            elif op == 'empty':
                ok = not text or text.startswith('（无')
            else:                       # not_empty
                ok = bool(text) and not text.startswith('（无')
            results.append(ok)
        return all(results) if multi.get('match_all', True) else any(results)

    def open_multi_dialog(self):
        dlg = MultiSummaryDialog(self, self.multi)
        if dlg.exec_() != QDialog.Accepted:
            return
        cfg = dlg.config()
        if cfg is None:
            return
        self.multi = cfg
        self.btn_clear_multi.setVisible(True)
        self.btn_multi.setText('多条件✓')
        self.refresh()

    def clear_multi(self, refresh=True):
        """恢复原有单条件汇总。"""
        if self.multi:
            self.multi = None
            self.btn_clear_multi.setVisible(False)
            self.btn_multi.setText('多条件…')
        if refresh:
            self.refresh()

    def collect(self):
        """按当前设置收集分组结果：[(分组值元组, 单位, 合计, 行数), ...]"""
        proj = getattr(self.win, 'project', None)
        if proj is None:
            return []
        sec = getattr(self.win, 'current_section', None)
        if self.only_cur.isChecked() and sec is not None:
            sections = list(sec.iter_self_and_desc())
        else:
            sections = proj.all_sections()
        fields = self.fields()
        show_unit = self.show_unit_col()
        groups = {}
        for s in sections:
            for r in s.rows:
                if getattr(r, 'no_count', False) or r.value is None:
                    continue
                if not self.pass_filters(r, s):
                    continue
                unit = (getattr(r, 'unit', '') or '').strip()
                keys = tuple(self.row_field(r, s, f) for f in fields)
                g = groups.setdefault((keys, unit if show_unit else ''), [0.0, 0])
                g[0] += float(r.value)
                g[1] += 1
        rows = [(k, u, v[0], v[1]) for (k, u), v in groups.items()]
        if fields == ['code'] and not self.multi:
            rows.sort(key=lambda t: natural_key(t[0][0]))
        else:
            rows.sort(key=lambda t: (-t[2], natural_key(t[0][0]) if t[0] else ''))
        return rows

    def refresh(self):
        rows = self.collect()
        was_empty = self.table.rowCount() == 0
        self._groups = rows
        total = sum(r[2] for r in rows)
        cnt = sum(r[3] for r in rows)
        self._total = (total, cnt)
        self._units = {r[1] for r in rows if r[1]}
        prec = getattr(self.win, 'precision', 3)
        heads = self.headers()
        ncol = len(heads)
        nkeys = len(self.fields())
        show_unit = self.show_unit_col()
        idx_unit = nkeys if show_unit else -1
        idx_total, idx_cnt = ncol - 2, ncol - 1
        self.table.setColumnCount(ncol)
        self.table.setHorizontalHeaderLabels(heads)
        hh = self.table.horizontalHeader()
        defaults = (160, 42, 84, 40, 90, 90)
        for c in range(ncol):
            hh.setSectionResizeMode(c, QHeaderView.Interactive)
            saved = self._col_widths.get(heads[c])
            if saved:
                self.table.setColumnWidth(c, int(saved))
            elif c < len(defaults):
                self.table.setColumnWidth(c, defaults[c])
        self.table.setRowCount(0)
        for keys, u, val, n in rows:
            r = self.table.rowCount()
            self.table.insertRow(r)
            cells = list(keys) + ([u] if show_unit else []) + [
                fmt_num(val, precision_for_unit(u, prec)), str(n)]
            for j, text in enumerate(cells):
                it = QTableWidgetItem(text)
                if j < nkeys:
                    it.setToolTip(text)
                if j == idx_total:
                    it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                elif j in (idx_cnt, idx_unit):
                    it.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(r, j, it)
        # 合计行
        if rows:
            r = self.table.rowCount()
            self.table.insertRow(r)
            unit = (list(self._units)[0] if len(self._units) == 1 else '')
            mixed = len(self._units) > 1
            cells = ['合计（%d 项）' % len(rows)] + [''] * (nkeys - 1)
            if show_unit:
                cells.append(unit)
            cells += ['—' if mixed else fmt_num(total, precision_for_unit(unit, prec)),
                      str(cnt)]
            for j, text in enumerate(cells):
                it = QTableWidgetItem(text)
                it.setFont(QFont('微软雅黑', 9, QFont.Bold))
                it.setBackground(QBrush(QColor('#f5e6a8')))
                if j == idx_total:
                    it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    it.setToolTip('多种单位不能直接相加，请按单位分别看各行' if mixed else '')
                elif j in (idx_cnt, idx_unit):
                    it.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(r, j, it)
        self.lbl_info.setText('%d 项 / %d 行' % (len(rows), cnt))
        # 行高默认自适应；列宽只在第一次有数据时自适应（之后保留用户手动拖的宽度）
        if rows:
            self.fit_rows()
            if was_empty:
                unsaved = [c for c, h in enumerate(heads) if h not in self._col_widths]
                if unsaved:
                    self.fit_columns(unsaved)
        self.btn_export.setEnabled(bool(rows))
        self.btn_copy.setEnabled(bool(rows))
        self.win.lbl_status.setText('分类汇总（%s）：%d 项，%d 行'
                                    % (self.mode_label(), len(rows), cnt))

    # ---------------- 导出 / 复制 ----------------
    def as_matrix(self):
        data = [self.headers()]
        nkeys = len(self.fields())
        show_unit = self.show_unit_col()
        prec = getattr(self.win, 'precision', 3)
        for keys, u, val, n in self._groups:
            data.append(list(keys) + ([u] if show_unit else []) +
                        [fmt_num(val, precision_for_unit(u, prec)), n])
        if self._groups:
            unit = (list(self._units)[0] if len(self._units) == 1 else '')
            mixed = len(self._units) > 1
            tail = ['—' if mixed else fmt_num(self._total[0],
                                             precision_for_unit(unit, prec)),
                    self._total[1]]
            cells = ['合计（%d 项）' % len(self._groups)] + [''] * (nkeys - 1)
            if show_unit:
                cells.append(unit)
            data.append(cells + tail)
        return data

    def copy_to_clipboard(self):
        if not self._groups:
            QMessageBox.information(self, APP_NAME, '当前没有可汇总的数据。')
            return
        text = '\n'.join('\t'.join(str(c) for c in row) for row in self.as_matrix())
        QApplication.clipboard().setText(text)
        self.win.status.showMessage('汇总表已复制到剪贴板', 3000)

    def export_excel(self):
        if not self._groups:
            QMessageBox.information(self, APP_NAME, '当前没有可汇总的数据。')
            return
        name = (getattr(self.win.project, 'name', '') or '工程量') + '-%s汇总.xlsx' % self.mode_label()
        path, _ = QFileDialog.getSaveFileName(self, '导出分类汇总', name,
                                              'Excel 文件 (*.xlsx);;所有文件 (*)')
        if not path:
            return
        if not path.lower().endswith('.xlsx'):
            path += '.xlsx'
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
            from openpyxl.utils import get_column_letter
            wb = Workbook()
            ws = wb.active
            ws.title = self.mode_label()[:31]
            data = self.as_matrix()
            ncol = len(data[0])
            nkeys = len(self.fields())
            thin = Side(style='thin', color='B0B0B0')
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            ws.append([(getattr(self.win.project, 'name', '') or '工程量计算书')
                       + ' — ' + self.mode_label() + '汇总'])
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
            ws.cell(row=1, column=1).font = Font(bold=True, size=13)
            ws.cell(row=1, column=1).alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 26
            for j, cell_text in enumerate(data[0], start=1):
                c = ws.cell(row=2, column=j, value=cell_text)
                c.font = Font(bold=True)
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.fill = PatternFill('solid', fgColor='F0E6C8')
                c.border = border
            for i, row in enumerate(data[1:], start=3):
                for j, cell_text in enumerate(row, start=1):
                    c = ws.cell(row=i, column=j, value=cell_text)
                    c.border = border
                    ha = ('right' if j == ncol - 1 else
                          ('center' if j == ncol or j == nkeys + 1 else 'left'))
                    c.alignment = Alignment(horizontal=ha, vertical='center')
                if i == 2 + len(self._groups):
                    for j in range(1, ncol + 1):
                        ws.cell(row=i, column=j).font = Font(bold=True)
                    ws.cell(row=i, column=1).fill = PatternFill('solid', fgColor='FFF2A8')
            for j in range(1, ncol + 1):
                letter = get_column_letter(j)
                ws.column_dimensions[letter].width = (30 if j <= nkeys else
                                                      (10 if j == nkeys + 1 else 14 if j == ncol - 1 else 8))
            ws.freeze_panes = 'A3'
            wb.save(path)
        except Exception as e:
            QMessageBox.critical(self, '导出失败', str(e))
            return
        self.win.status.showMessage('已导出：' + os.path.basename(path), 5000)
        if QMessageBox.question(self, '导出成功', '已生成：\n%s\n\n是否立即打开？' % path,
                                QMessageBox.Yes | QMessageBox.No,
                                QMessageBox.Yes) == QMessageBox.Yes:
            try:
                os.startfile(path)
            except Exception:
                pass




class MultiSummaryDialog(QDialog):
    """多条件汇总设置：多个分组字段 + 多条筛选条件（满足全部 / 满足任一）。"""

    OPS = [('contains', '包含'), ('not_contains', '不包含'), ('eq', '等于'),
           ('ne', '不等于'), ('empty', '为空'), ('not_empty', '不为空')]

    def __init__(self, panel, current=None, parent=None):
        super().__init__(parent or panel)
        self.setWindowTitle('多条件汇总设置')
        self.resize(520, 420)
        self._rows = []
        v = QVBoxLayout(self)

        gb = QGroupBox('分组字段（至少选一个，可多选）')
        gbl = QVBoxLayout(gb)
        self.checks = {}
        grid = QGridLayout()
        for i, (key, label) in enumerate(SummaryPanel.MODES):
            cb = QCheckBox(label)
            self.checks[key] = cb
            grid.addWidget(cb, i // 3, i % 3)
        gbl.addLayout(grid)
        v.addWidget(gb)

        gb2 = QGroupBox('筛选条件（可留空，表示不筛选）')
        g2 = QVBoxLayout(gb2)
        self.cond_host = QWidget()
        self.cond_layout = QVBoxLayout(self.cond_host)
        self.cond_layout.setContentsMargins(0, 0, 0, 0)
        self.cond_layout.setSpacing(3)
        g2.addWidget(self.cond_host)
        bar = QHBoxLayout()
        b_add = QPushButton('+ 添加条件')
        b_add.clicked.connect(self.add_condition)
        bar.addWidget(b_add)
        bar.addStretch(1)
        g2.addLayout(bar)
        v.addWidget(gb2)

        match = QHBoxLayout()
        self.rb_all = QRadioButton('满足全部条件（AND）')
        self.rb_any = QRadioButton('满足任一条件（OR）')
        bg = QButtonGroup(self)
        bg.addButton(self.rb_all)
        bg.addButton(self.rb_any)
        self.rb_all.setChecked(True)
        match.addWidget(self.rb_all)
        match.addWidget(self.rb_any)
        match.addStretch(1)
        v.addLayout(match)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

        current = current or {}
        for key in (current.get('fields') or [panel.mode_key()]):
            if key in self.checks:
                self.checks[key].setChecked(True)
        self.rb_all.setChecked(current.get('match_all', True))
        self.rb_any.setChecked(not current.get('match_all', True))
        for field, op, value in (current.get('filters') or []):
            self.add_condition(field, op, value)
        if not self._rows:
            self.add_condition()

    def add_condition(self, field='code', op='contains', value=''):
        """添加一行筛选条件；信号连 lambda 时前三个参数有默认值，可直接当槽用。"""
        row = QWidget()
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(4)
        cb_field = QComboBox()
        for key, label in SummaryPanel.MODES:
            cb_field.addItem(label, key)
        idx = cb_field.findData(field if isinstance(field, str) else 'code')
        cb_field.setCurrentIndex(max(0, idx))
        cb_op = QComboBox()
        for key, label in self.OPS:
            cb_op.addItem(label, key)
        idx = cb_op.findData(op if isinstance(op, str) else 'contains')
        cb_op.setCurrentIndex(max(0, idx))
        ed = QLineEdit(value if isinstance(value, str) else '')
        ed.setPlaceholderText('关键词（为空/不为空可留空）')
        btn = QToolButton()
        btn.setText('×')
        btn.setToolTip('删除此条件')
        btn.setAutoRaise(True)
        btn.clicked.connect(lambda: self.remove_condition(row))
        h.addWidget(cb_field)
        h.addWidget(cb_op)
        h.addWidget(ed, 1)
        h.addWidget(btn)
        self.cond_layout.addWidget(row)
        self._rows.append((row, cb_field, cb_op, ed))
        return row

    def remove_condition(self, row):
        for i, item in enumerate(self._rows):
            if item[0] is row:
                self._rows.pop(i)
                self.cond_layout.removeWidget(row)
                row.setParent(None)
                row.deleteLater()
                break

    def config(self):
        fields = [key for key, cb in self.checks.items() if cb.isChecked()]
        if not fields:
            QMessageBox.information(self, APP_NAME, '请至少选择一个分组字段。')
            return None
        filters = []
        for _row, cb_field, cb_op, ed in self._rows:
            val = ed.text().strip()
            op = cb_op.currentData()
            if op in ('empty', 'not_empty') or val:
                filters.append((cb_field.currentData(), op, val))
        return {'fields': fields, 'filters': filters,
                'match_all': self.rb_all.isChecked()}


class FindDialog(QDialog):
    """查找：可按 编号 / 项目名称 / 单位 / 部位 / 计算公式 / 备注 查找，
    也可以直接用“全部字段（自定义关键词）”搜任意文字。

    命中后自动切到对应分部、选中并滚动到该行，方便连续查找。
    """

    FIELDS = [('all', '全部字段（自定义关键词）'), ('code', '编号'), ('name', '项目名称'),
              ('part', '部位'), ('unit', '单位'), ('formula', '计算公式'),
              ('remark', '备注')]

    def __init__(self, win):
        super().__init__(win)
        self.win = win
        self.setWindowTitle('查找')
        self.setWindowFlag(Qt.Tool, True)
        self.setMinimumWidth(520)
        self._hits = []
        self._idx = -1
        self._sig = None

        v = QVBoxLayout(self)
        v.setContentsMargins(10, 8, 10, 8)
        v.setSpacing(6)

        row1 = QHBoxLayout()
        row1.setSpacing(6)
        row1.addWidget(QLabel('关键词：'))
        self.kw = QLineEdit()
        self.kw.setPlaceholderText('输入要查找的文字（如 C30、混凝土、m3、外墙…）')
        self.kw.setClearButtonEnabled(True)
        row1.addWidget(self.kw, 1)
        v.addLayout(row1)

        row2 = QHBoxLayout()
        row2.setSpacing(6)
        row2.addWidget(QLabel('查找范围：'))
        self.field = QComboBox()
        for k, label in self.FIELDS:
            self.field.addItem(label, k)
        self.field.setMinimumWidth(170)
        row2.addWidget(self.field)
        self.scope = QComboBox()
        self.scope.addItem('全部单项工程', 'all')
        self.scope.addItem('当前分部（含下级）', 'cur')
        row2.addWidget(self.scope)
        row2.addStretch(1)
        v.addLayout(row2)

        row3 = QHBoxLayout()
        row3.setSpacing(6)
        self.lbl = QLabel('输入关键词后按回车查找')
        self.lbl.setStyleSheet('color:#666;')
        row3.addWidget(self.lbl, 1)
        self.btn_prev = QPushButton('上一个')
        self.btn_next = QPushButton('下一个')
        self.btn_next.setDefault(True)
        self.btn_prev.clicked.connect(lambda: self.step(-1))
        self.btn_next.clicked.connect(lambda: self.step(1))
        row3.addWidget(self.btn_prev)
        row3.addWidget(self.btn_next)
        v.addLayout(row3)

        self.kw.textChanged.connect(self._invalidate)
        self.kw.returnPressed.connect(lambda: self.step(1))
        self.field.currentIndexChanged.connect(self._invalidate)
        self.scope.currentIndexChanged.connect(self._invalidate)
        self.btn_prev.setEnabled(False)
        self.btn_next.setEnabled(False)

    # ---------- 查找 ----------
    def _invalidate(self, *args):
        self._sig = None
        self._hits = []
        self._idx = -1
        kw = self.kw.text().strip()
        if not kw:
            self.lbl.setText('输入关键词后按回车查找')
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)
            return
        self._build()
        n = len(self._hits)
        self.lbl.setText('共 %d 个匹配，按回车逐个查看' % n if n else '没有找到「%s」' % kw)
        self.btn_prev.setEnabled(n > 0)
        self.btn_next.setEnabled(n > 0)

    def _sections(self):
        win = self.win
        if self.scope.currentData() == 'cur' and win.current_section is not None:
            return list(win.current_section.iter_self_and_desc())
        return win.project.all_sections()

    def _texts(self, row, key):
        """返回该行参与匹配的 (列key, 文本) 列表。"""
        if key == 'all':
            pairs = [('code', row.code), ('name', row.name), ('part', row.part),
                     ('unit', row.unit), ('formula', row.formula), ('remark', row.remark)]
            for k, v in (getattr(row, 'custom', None) or {}).items():
                pairs.append((k, v))
            return [(k, str(v or '')) for k, v in pairs]
        val = getattr(row, key, None)
        if val is None:
            val = (getattr(row, 'custom', None) or {}).get(key, '')
        return [(key, str(val or ''))]

    def _build(self):
        """按当前关键词 / 范围 / 字段重建命中列表。"""
        kw = self.kw.text().strip().lower()
        field = self.field.currentData() or 'all'
        sig = (kw, field, self.scope.currentData())
        if sig == self._sig:
            return
        self._sig = sig
        hits = []
        if kw:
            for sec in self._sections():
                for ri, row in enumerate(sec.rows):
                    for k, text in self._texts(row, field):
                        if kw in (text or '').lower():
                            hits.append((sec, ri, k))
                            break
        self._hits = hits
        self._idx = -1

    def step(self, delta):
        kw = self.kw.text().strip()
        if not kw:
            return
        self._build()
        n = len(self._hits)
        if not n:
            self.lbl.setText('没有找到「%s」' % kw)
            return
        if self._idx < 0:
            self._idx = 0 if delta >= 0 else n - 1
        else:
            self._idx = (self._idx + delta) % n
        sec, ri, key = self._hits[self._idx]
        self.win.goto_cell(sec, ri, key, kw)
        self.lbl.setText('第 %d / %d 个匹配' % (self._idx + 1, n))


class SymbolPicker(QDialog):
    """特殊符号面板：工程符号 + 钢筋等级 + 直径/半径，点一下插入公式框。

    点按钮后插入并保持面板打开，可连续插入；点面板外面自动关闭。
    """

    def __init__(self, insert_cb, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.Popup)
        self.setWindowTitle('特殊符号')
        self.insert_cb = insert_cb
        self.setStyleSheet('''
            QDialog { background:#fbfcfe; border:1px solid #9db4cc; }
            QToolButton { border:1px solid #d8dde3; border-radius:2px; background:#fff;
                          color:#222; font-size:12pt; }
            QToolButton:hover { background:#e8f0fa; border:1px solid #9dbce0; }
            QToolButton:pressed { background:#d4e4f7; }
            QPushButton#symClose { border:1px solid #cfd4da; border-radius:3px;
                          background:#fff; padding:2px 10px; font-size:9pt; }
            QPushButton#symClose:hover { background:#eef3f8; }
        ''')
        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 6, 8, 6)
        lay.setSpacing(3)

        def title(text):
            lb = QLabel(text)
            lb.setStyleSheet('color:#666;font-size:9pt;')
            lay.addWidget(lb)

        # —— 工程符号 ——
        title('工程符号（插入到公式框光标处）')
        g1 = QGridLayout()
        g1.setContentsMargins(0, 0, 0, 0)
        g1.setSpacing(2)
        for i, (ch, tip) in enumerate(ENG_SYMBOLS):
            b = QToolButton()
            b.setText(ch)
            b.setToolTip(tip)
            b.setFixedSize(34, 28)
            b.clicked.connect(lambda checked=False, c=ch: self.insert_cb(c))
            g1.addWidget(b, i // 9, i % 9)
        lay.addLayout(g1)

        # —— 钢筋等级（图形显示标准符号，插入行业通用写法）——
        title('钢筋等级（图形为标准图纸符号，插入 Φ、Φ²、Φ³、Φ⁴）')
        g2 = QHBoxLayout()
        g2.setContentsMargins(0, 0, 0, 0)
        g2.setSpacing(3)
        for ch, tip, level in REBAR_GRADES:
            b = QToolButton()
            b.setText({1: '一级', 2: '二级', 3: '三级', 4: '四级'}[level])
            b.setToolTip(tip)
            b.setIcon(QIcon(rebar_pixmap(level, 30)))
            b.setIconSize(QSizeF(30, 30).toSize())
            b.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
            b.setFixedSize(56, 62)
            b.clicked.connect(lambda checked=False, c=ch: self.insert_cb(c))
            g2.addWidget(b)
        g2.addStretch(1)
        lay.addLayout(g2)

        row = QHBoxLayout()
        row.addWidget(QLabel('插入位置：公式辅助框光标处，或当前单元格文本末尾'))
        cl = QPushButton('关闭')
        cl.setObjectName('symClose')
        cl.clicked.connect(self.close)
        row.addWidget(cl)
        lay.addLayout(row)


# 函数面板：分组、显示名、插入模板、插入后光标回退字符数
FUNCTION_GROUPS = [
    ('常用运算', [
        ('abs', '绝对值', 'abs()', 1, 'abs(-3) = 3'),
        ('round', '四舍五入', 'round()', 1, 'round(3.456,2) = 3.46'),
        ('ceil', '向上取整', 'ceil()', 1, 'ceil(2.1) = 3'),
        ('floor', '向下取整', 'floor()', 1, 'floor(2.9) = 2'),
        ('power', '幂运算', 'power()', 1, 'power(2,3) = 8（也可写 2^3）'),
        ('sqr', '平方', 'sqr()', 1, 'sqr(3) = 9'),
        ('sqrt', '开平方', 'sqrt()', 1, 'sqrt(16) = 4（也可写 √16）'),
        ('√', '开方符号', '√()', 1, '√9 = 3、√(3+4)'),
    ]),
    ('三角与角度', [
        ('sin', '正弦', 'sin()', 1, 'sin(radians(30)) = 0.5'),
        ('cos', '余弦', 'cos()', 1, 'cos(0) = 1'),
        ('tan', '正切', 'tan()', 1, 'tan(radians(45)) = 1'),
        ('radians', '角度→弧度', 'radians()', 1, 'radians(180) = 3.1416'),
        ('du', '弧度→角度', 'du()', 1, 'du(3.1416) = 180'),
    ]),
    ('统计', [
        ('max', '最大值', 'max()', 1, 'max(1,2,3) = 3'),
        ('min', '最小值', 'min()', 1, 'min(1,2,3) = 1'),
        ('sum', '求和', 'sum()', 1, 'sum(1,2,3) = 6'),
        ('avg', '求平均值', 'avg()', 1, 'avg(1,2,3) = 2'),
    ]),
    ('常量', [
        ('π', '圆周率', 'π', 0, 'π = 3.14159'),
    ]),
]


class FunctionPicker(QDialog):
    """插入函数面板：点函数名即插入“函数名()”，光标停在括号内。"""

    def __init__(self, insert_cb, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.Popup)
        self.setWindowTitle('插入函数')
        self.insert_cb = insert_cb
        self.setStyleSheet('''
            QDialog { background:#fbfcfe; border:1px solid #9db4cc; }
            QToolButton { border:1px solid #d8dde3; border-radius:2px; background:#fff;
                          color:#222; font-size:9pt; text-align:left; padding:3px 6px; }
            QToolButton:hover { background:#e8f0fa; border:1px solid #9dbce0; }
            QToolButton:pressed { background:#d4e4f7; }
            QPushButton#fnClose { border:1px solid #cfd4da; border-radius:3px;
                          background:#fff; padding:2px 10px; font-size:9pt; }
            QPushButton#fnClose:hover { background:#eef3f8; }
        ''')
        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 6, 8, 6)
        lay.setSpacing(4)

        tip = QLabel('选择函数插入到公式框光标处（括号内输入参数）')
        tip.setStyleSheet('color:#666;font-size:9pt;')
        lay.addWidget(tip)

        for gname, items in FUNCTION_GROUPS:
            lb = QLabel(gname)
            lb.setStyleSheet('color:#24507a;font-size:9pt;font-weight:bold;')
            lay.addWidget(lb)
            grid = QGridLayout()
            grid.setContentsMargins(0, 0, 0, 0)
            grid.setHorizontalSpacing(3)
            grid.setVerticalSpacing(2)
            for i, (name, label, tpl, back, hint) in enumerate(items):
                b = QToolButton()
                b.setText('%s  %s' % (name, label))
                b.setToolTip('%s  ——  %s' % (label, hint))
                b.setFont(QFont('微软雅黑', 9))
                tw = b.fontMetrics().horizontalAdvance(b.text())
                b.setFixedWidth(max(96, tw + 26))
                b.setFixedHeight(26)
                b.clicked.connect(
                    lambda checked=False, t=tpl, k=back: self.insert_cb(t, k))
                grid.addWidget(b, i // 4, i % 4)
            for c in range(4):
                grid.setColumnStretch(c, 1)
            lay.addLayout(grid)

        row = QHBoxLayout()
        note = QLabel('汉字、[] 内的内容为备注不参与计算；Ctrl+Enter 写入单元格')
        note.setStyleSheet('color:#888;font-size:9pt;')
        row.addWidget(note)
        row.addStretch(1)
        cl = QPushButton('关闭')
        cl.setObjectName('fnClose')
        cl.clicked.connect(self.close)
        row.addWidget(cl)
        lay.addLayout(row)


class ProjectInfoDialog(QDialog):
    """工程/项目信息（工程名称、建设单位、施工单位、设计单位、编制日期、备注）。"""

    def __init__(self, project, parent=None, for_new=False):
        super().__init__(parent)
        self.setWindowTitle('新建工程' if for_new else '工程信息设置')
        self.setMinimumWidth(420)
        self.project = project
        self.for_new = for_new
        layout = QFormLayout(self)
        self.edits = {}
        fields = [('proj_name', '工程名称'), ('owner', '建设单位'), ('builder', '施工单位'),
                  ('designer', '设计单位'), ('date', '编制日期'), ('note', '备注')]
        defaults = {'proj_name': project.name, 'date': datetime.date.today().isoformat()}
        for key, label in fields:
            ed = QLineEdit(project.info.get(key, defaults.get(key, '')))
            self.edits[key] = ed
            layout.addRow(label, ed)
        tip = '填写后将写入工程信息，并用于报表表头（工程名称）'
        lbl = QLabel(tip if not for_new else '填写工程信息后创建新工程；只有工程名称必填')
        lbl.setStyleSheet('color:#666;')
        lbl.setWordWrap(True)
        layout.addRow(lbl)
        if for_new:
            self.edits['proj_name'].selectAll()
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

    def apply(self):
        self.project.name = self.edits['proj_name'].text().strip() or self.project.name
        for k, ed in self.edits.items():
            self.project.info[k] = ed.text().strip()


class ReportSettingsDialog(QDialog):
    """页面设置：表头标题、页眉页脚、上下左右页边距、纸张方向、水平/垂直对齐。"""

    def __init__(self, report, parent=None):
        super().__init__(parent)
        self.setWindowTitle('页面设置（打印 / 导出）')
        self.setMinimumWidth(560)
        self.report = dict(report)
        form = QFormLayout(self)
        self.ed_title = QLineEdit(self.report.get('title', ''))
        self.ed_title.setPlaceholderText('留空则使用"工程名 工程量计算书"')
        self.ed_header = QLineEdit(self.report.get('header', ''))
        self.ed_header.setPlaceholderText('每页顶部文字，如：XX项目 造价咨询有限公司')
        self.ed_footer_left = QLineEdit(self.report.get('footer_left', ''))
        self.ed_footer_left.setPlaceholderText('例如：编制人：__________')
        self.ed_footer_center = QLineEdit(self.report.get('footer_center', ''))
        self.ed_footer_center.setPlaceholderText('中间文字（页码自动加在后面）')
        self.ed_footer_right = QLineEdit(self.report.get('footer_right', ''))
        self.ed_footer_right.setPlaceholderText('例如：复核人：__________')
        if not (self.ed_footer_left.text() or self.ed_footer_center.text()
                or self.ed_footer_right.text()):
            # 兼容旧版单一页脚字段
            self.ed_footer_center.setText(self.report.get('footer', ''))
        form.addRow('表头标题：', self.ed_title)
        form.addRow('页眉：', self.ed_header)
        form.addRow('页脚左：', self.ed_footer_left)
        form.addRow('页脚中：', self.ed_footer_center)
        form.addRow('页脚右：', self.ed_footer_right)

        # —— 页边距：上下左右分别设置 ——
        margins = self.report.get('margins') or {}
        base = float(margins.get('left', self.report.get('margin', 1.9)))
        self.sp_top = self._spin(margins.get('top', base))
        self.sp_bottom = self._spin(margins.get('bottom', base))
        self.sp_left = self._spin(margins.get('left', base))
        self.sp_right = self._spin(margins.get('right', base))
        margin_box = QWidget()
        mg = QHBoxLayout(margin_box)
        mg.setContentsMargins(0, 0, 0, 0)
        for text, sp in (('上', self.sp_top), ('下', self.sp_bottom),
                         ('左', self.sp_left), ('右', self.sp_right)):
            mg.addWidget(QLabel(text))
            mg.addWidget(sp)
        form.addRow('页边距(cm)：', margin_box)

        self.cb_orient = QComboBox()
        self.cb_orient.addItems(['纵向', '横向'])
        self.cb_orient.setCurrentText('横向' if self.report.get('orientation') == 'landscape' else '纵向')
        form.addRow('纸张方向：', self.cb_orient)

        # —— 纸张大小：取系统打印机支持的纸张，也可自定义宽高 ——
        paper = paper_of(self.report)
        self.cb_paper = QComboBox()
        self._papers = system_paper_list()
        for name, w, h in self._papers:
            self.cb_paper.addItem('%s（%0.0f×%0.0f mm）' % (name, w, h), (name, w, h))
        self.cb_paper.addItem('自定义', None)
        self.sp_pw = QDoubleSpinBox()
        self.sp_ph = QDoubleSpinBox()
        for sp, val in ((self.sp_pw, paper['w']), (self.sp_ph, paper['h'])):
            sp.setRange(50.0, 2000.0)
            sp.setDecimals(1)
            sp.setSingleStep(1.0)
            sp.setSuffix(' mm')
            sp.setValue(val)
        paper_row = QHBoxLayout()
        paper_row.setContentsMargins(0, 0, 0, 0)
        paper_row.addWidget(self.cb_paper, 2)
        paper_row.addWidget(QLabel('宽'))
        paper_row.addWidget(self.sp_pw)
        paper_row.addWidget(QLabel('高'))
        paper_row.addWidget(self.sp_ph)
        form.addRow('纸张大小：', paper_row)
        # 选中与当前尺寸匹配的纸张
        idx = self.cb_paper.count() - 1
        for i, (_n, w, h) in enumerate(self._papers):
            if abs(w - paper['w']) < 0.6 and abs(h - paper['h']) < 0.6:
                idx = i
                break
        self.cb_paper.setCurrentIndex(idx)
        self._on_paper_changed()
        self.cb_paper.currentIndexChanged.connect(self._on_paper_changed)

        self.cb_halign = QComboBox()
        for text, val in (('居中', 'center'), ('靠左', 'left'), ('靠右', 'right')):
            self.cb_halign.addItem(text, val)
        self.cb_halign.setCurrentIndex(max(0, self.cb_halign.findData(self.report.get('halign', 'left'))))
        form.addRow('水平对齐：', self.cb_halign)
        self.cb_valign = QComboBox()
        for text, val in (('居中', 'center'), ('置顶', 'top'), ('置底', 'bottom')):
            self.cb_valign.addItem(text, val)
        self.cb_valign.setCurrentIndex(max(0, self.cb_valign.findData(self.report.get('valign', 'center'))))
        form.addRow('垂直对齐：', self.cb_valign)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        form.addRow(bb)

    @staticmethod
    def _spin(value):
        sp = QDoubleSpinBox()
        sp.setRange(0.0, 10.0)
        sp.setSingleStep(0.1)
        sp.setDecimals(1)
        sp.setSuffix(' cm')
        sp.setValue(float(value))
        return sp

    def _on_paper_changed(self, *args):
        """选了预设纸张 → 宽高跟着变且锁定；选“自定义” → 可自由输入。"""
        data = self.cb_paper.currentData()
        custom = data is None
        self.sp_pw.setEnabled(custom)
        self.sp_ph.setEnabled(custom)
        if not custom:
            _name, w, h = data
            self.sp_pw.setValue(w)
            self.sp_ph.setValue(h)

    def current_paper(self):
        data = self.cb_paper.currentData()
        if data is None:
            return {'name': '自定义', 'w': round(self.sp_pw.value(), 1),
                    'h': round(self.sp_ph.value(), 1)}
        return {'name': data[0], 'w': float(data[1]), 'h': float(data[2])}

    def apply(self):
        paper = self.current_paper()
        orient = 'landscape' if self.cb_orient.currentText() == '横向' else 'portrait'
        # 换纸张/方向后若页边距超出可用区域，自动收进合理范围
        pw, ph = paper['w'], paper['h']
        if orient == 'landscape':
            pw, ph = ph, pw
        max_lr = max(0.5, pw / 10.0 - 5.0)
        max_tb = max(0.5, ph / 10.0 - 5.0)
        self.report.update({
            'title': self.ed_title.text().strip(),
            'header': self.ed_header.text().strip(),
            'footer_left': self.ed_footer_left.text().strip(),
            'footer_center': self.ed_footer_center.text().strip(),
            'footer_right': self.ed_footer_right.text().strip(),
            'margins': {
                'top': min(self.sp_top.value(), max_tb),
                'bottom': min(self.sp_bottom.value(), max_tb),
                'left': min(self.sp_left.value(), max_lr),
                'right': min(self.sp_right.value(), max_lr),
            },
            'margin': self.sp_left.value(),   # 兼容旧字段
            'orientation': orient,
            'paper': paper,
            'halign': self.cb_halign.currentData() or 'left',
            'valign': self.cb_valign.currentData() or 'center',
        })


class FillColorDialog(QDialog):
    """行/列/单元格填充颜色：选颜色 + 选范围（整行 / 整列 / 仅当前单元格）。"""

    PALETTE = [
        ('#FFF2CC', '浅黄'), ('#FFE699', '黄'), ('#FFD966', '金黄'), ('#FFC000', '橙黄'),
        ('#F4B183', '橙'), ('#F8CBAD', '浅橙红'), ('#FF9999', '红'), ('#FF7C80', '深红'),
        ('#E2EFDA', '浅绿'), ('#C6E0B4', '绿'), ('#A9D08E', '深绿'), ('#70AD47', '草绿'),
        ('#DDEBF7', '浅蓝'), ('#BDD7EE', '蓝'), ('#9DC3E6', '中蓝'), ('#B4C6E7', '灰蓝'),
        ('#E4DFEC', '浅紫'), ('#D9D2E9', '紫'), ('#DAEEF3', '青'), ('#F2F2F2', '浅灰'),
        ('#E7E6E6', '灰'), ('#D0CECE', '中灰'), ('#D6DCE4', '灰蓝2'), ('#FFFFFF', '白色'),
    ]

    def __init__(self, parent=None, scope='cell', color=''):
        super().__init__(parent)
        self.setWindowTitle('填充颜色')
        self.setMinimumWidth(430)
        self.color = color or ''
        self.scope = scope
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(6)

        tip = QLabel('选择颜色后点“确定”应用到选中的区域；点“清除填充”去掉颜色。')
        tip.setStyleSheet('color:#666;font-size:9pt;')
        lay.addWidget(tip)

        # 应用范围
        scope_row = QHBoxLayout()
        scope_row.addWidget(QLabel('应用范围：'))
        self.scope_group = QButtonGroup(self)
        for i, (text, val) in enumerate((('整行', 'row'), ('整列', 'col'), ('仅选中单元格', 'cell'))):
            rb = QRadioButton(text)
            rb.setProperty('scope', val)
            if val == self.scope:
                rb.setChecked(True)
            self.scope_group.addButton(rb, i)
            scope_row.addWidget(rb)
        scope_row.addStretch(1)
        lay.addLayout(scope_row)

        # 色板
        grid = QGridLayout()
        grid.setSpacing(3)
        self._btns = []
        for i, (hexv, name) in enumerate(self.PALETTE):
            b = QToolButton()
            b.setFixedSize(44, 26)
            b.setToolTip(name)
            b.setStyleSheet('QToolButton { background:%s; border:1px solid #b6bcc4; }'
                            'QToolButton:hover { border:2px solid #2f5597; }' % hexv)
            b.clicked.connect(lambda checked=False, c=hexv: self._pick(c))
            grid.addWidget(b, i // 8, i % 8)
            self._btns.append((b, hexv))
        lay.addLayout(grid)

        # 当前颜色 + 自定义
        cur_row = QHBoxLayout()
        cur_row.addWidget(QLabel('当前颜色：'))
        self.lbl_cur = QLabel()
        self.lbl_cur.setFixedSize(56, 24)
        self.lbl_cur.setStyleSheet('border:1px solid #b6bcc4; background:%s;'
                                   % (self.color or '#ffffff'))
        cur_row.addWidget(self.lbl_cur)
        self.lbl_hex = QLabel(self.color or '无填充')
        self.lbl_hex.setStyleSheet('color:#555;font-size:9pt;')
        cur_row.addWidget(self.lbl_hex)
        btn_custom = QPushButton('自定义颜色…')
        btn_custom.clicked.connect(self._custom)
        cur_row.addWidget(btn_custom)
        cur_row.addStretch(1)
        lay.addLayout(cur_row)
        self._highlight_current()

        bb = QHBoxLayout()
        btn_clear = QPushButton('清除填充')
        btn_clear.clicked.connect(self._clear)
        bb.addWidget(btn_clear)
        bb.addStretch(1)
        box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        box.button(QDialogButtonBox.Ok).setText('确定')
        box.button(QDialogButtonBox.Cancel).setText('取消')
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        bb.addWidget(box)
        lay.addLayout(bb)

    def _pick(self, hexv):
        self.color = hexv
        self.lbl_cur.setStyleSheet('border:1px solid #b6bcc4; background:%s;' % hexv)
        self.lbl_hex.setText(hexv.upper())
        self._highlight_current()

    def _highlight_current(self):
        for b, hexv in self._btns:
            base = 'QToolButton { background:%s; border:1px solid #b6bcc4; }' % hexv
            if hexv.lower() == (self.color or '').lower():
                b.setStyleSheet('QToolButton { background:%s; border:2px solid #2f5597; }' % hexv)
            else:
                b.setStyleSheet(base + 'QToolButton:hover { border:2px solid #2f5597; }')

    def _custom(self):
        c = QColorDialog.getColor(QColor(self.color or '#FFF2CC'), self, '选择填充颜色')
        if c.isValid():
            self._pick(c.name())

    def _clear(self):
        self.color = ''
        self.lbl_cur.setStyleSheet('border:1px solid #b6bcc4; background:#ffffff;')
        self.lbl_hex.setText('无填充')
        self._highlight_current()
        self.accept()

    def current_scope(self):
        for rb in self.scope_group.buttons():
            if rb.isChecked():
                return rb.property('scope')
        return 'cell'


class MainWindow(QMainWindow):
    # 去掉 Windows 系统标题栏，改用自绘标题栏（信息块放在真正的窗口右上角）
    FRAMELESS = True
    RESIZE_MARGIN = 6          # 边缘拖拽调整大小的感应宽度（px）

    def __init__(self, app):
        super().__init__()
        self.app = app
        # 自绘标题栏相关状态
        self._maximized = False
        self._normal_geom = None
        self._rs_dir = ''          # 正在拖拽的边：l/r/t/b 组合
        self._rs_gpos = None
        self._rs_geom = None
        self._cursor_overridden = False
        if self.FRAMELESS:
            self.setWindowFlag(Qt.FramelessWindowHint, True)
            try:
                app.installEventFilter(self)
            except Exception:
                pass
        self.settings = make_settings()
        self.precision = int(self.settings.value('precision', 3))
        self.project = Project('未命名工程')
        self.current_section = self.project.root
        self.current_file = None
        self.recent = self.settings.value('recent', [], type=list) or []
        self._loading = False
        self._formula_row = None   # 辅助框当前关联的行索引
        self._print_area = None    # 选定打印区域：{'section', 'rows', 'keys'}
        self._syncing = False      # 双向同步防循环标志
        self._cell_editor = None   # 当前正在编辑的表格单元格编辑器
        # 自动保存设置
        self.autosave_enabled = self.settings.value('autosave_enabled', False, type=bool)
        self.autosave_min = int(self.settings.value('autosave_min', 5))
        self.autosave_dir = self.settings.value('autosave_dir', '')
        self._dirty = False
        # 行高 / 列宽默认自适应（V2.5）：不用每次手动点“自适应”
        self._row_autofit = self.settings.value('row_autofit', True, type=bool)
        self._col_autofit = self.settings.value('col_autofit', True, type=bool)
        # 当前行 / 列高亮（只加淡色底纹，不改变选择模式，避免影响框选删除、粘贴锚点）
        self._hilite_row = self.settings.value('hilite_row', False, type=bool)
        # V2.10.0：格式刷拾取的格式、生效中的筛选条件
        self._paint_fmt = None
        self._filter = None
        self._hilite_col = self.settings.value('hilite_col', False, type=bool)
        # 表格字体（类型取系统字库，大小 6-24；空 = 跟随软件默认微软雅黑）
        self._font_family = self.settings.value('table_font_family', '') or ''
        self._font_size = int(self.settings.value('table_font_size', 0) or 0)
        # 撤销 / 恢复：栈里存的是“改动前”的工程快照
        self._undo_stack = []        # [(标签, 快照)]
        self._redo_stack = []
        self._pending_undo = None    # 正在连续编辑时的改动前快照
        self._pending_label = ''
        self.last_file = self.settings.value('last_file', '') or ''
        self.autosave_timer = QTimer(self)
        self.autosave_timer.timeout.connect(self.autosave_tick)
        self.reset_autosave_timer()

        self.update_title()
        self.resize(1500, 900)
        self.setMinimumWidth(1100)
        def res_path(name):
            base = getattr(sys, '_MEIPASS', HERE)
            return os.path.join(base, name)
        icon = res_path('calcbook.ico')
        if not os.path.exists(icon):
            icon = res_path('rabbit-green.ico')
        self.app_icon_path = icon
        if os.path.exists(icon):
            self.setWindowIcon(QIcon(icon))
            app.setWindowIcon(QIcon(icon))
        # 打包运行时自动维护 .gcls 关联（换过目录/首次运行都会重新指向当前 exe）
        if getattr(sys, 'frozen', False) and assoc_needs_update():
            register_file_assoc()
        self.build_ui()
        self._apply_table_font(refit=False)   # 恢复上次设置的表格字体
        # 启动时恢复上次打开的工程（工程名直接显示在标题栏）
        # 测试/调试时可设 CALCBOOK_NO_RESTORE=1 跳过，避免读到用户上次的工程
        restored = False
        if not os.environ.get('CALCBOOK_NO_RESTORE'):
            restored = self.restore_last_project()
        if not restored:
            self.refresh_tree()
            self.refresh_table()
        self.update_recent_menu()
        # 诊断开关（发版时用 set CALCBOOK_DUMP=1 启动，会把面板几何写到 dump 文件）
        if os.environ.get('CALCBOOK_DUMP'):
            QTimer.singleShot(3000, self._dump_geometry)

    def _dump_geometry(self):
        try:
            p = os.path.join(os.path.expanduser('~'), 'calcbook_dump.txt')
            with open(p, 'w', encoding='utf-8') as f:
                f.write('win %dx%d\n' % (self.width(), self.height()))
                f.write('main sizes %s width %s\n' % (self.main_split.sizes(), self.main_split.width()))
                f.write('left %s sum %s work %s\n' % (
                    self.left_frame.width(), self.summary_panel.width(), self.work_pane.width()))
                f.write('left_split %s h %s\n' % (self.left_split.sizes(), self.left_split.height()))
                f.write('const %s\n' % self.const_box.height())
                f.write('folded nav=%s sum=%s const=%s\n' % (
                    self._nav_folded, self._sum_folded, self._const_folded))
                f.write('const rows %s cols %s\n' % (
                    self.const_list.rowCount(),
                    [self.const_list.columnWidth(i) for i in range(3)]))
                f.write('summary rows %s\n' % self.summary_panel.table.rowCount())
                f.write('section %s level %s rows %d table %d\n' % (
                    self.current_section.name, self.project.level_of(self.current_section),
                    len(self.current_section.rows), self.table.rowCount()))
                f.write('autofit col=%s row=%s\n' % (self._col_autofit, self._row_autofit))
        except Exception as e:
            pass

    # ---------------- 标题栏 / 上次工程 ----------------
    def update_title(self):
        """自绘标题栏：左上角放“软件图标 + 软件名版本 + 工程文件名称 + 文件位置”，
        工程名（“某某项目”）只留在窗口标题（任务栏）里。"""
        name = (self.project.name or '').strip() or '未命名工程'
        self.setWindowTitle(name)
        try:
            if hasattr(self, 'title_strip'):
                self.title_strip.set_project_name(name)
        except Exception:
            pass
        self.update_file_info()

    def update_file_info(self):
        """把“软件图标 + 软件名 版本 | 工程文件名称 | <文件存放详细地址>”刷到标题栏左端。"""
        if not hasattr(self, 'title_strip'):
            return
        path = getattr(self, 'current_file', '') or ''
        try:
            fname = os.path.basename(os.path.normpath(path)) if path else ''
        except Exception:
            fname = ''
        try:
            self.title_strip.set_file_info(f'{APP_NAME} V{APP_VER}', path, fname)
        except Exception:
            pass

    # ---------------- 自绘标题栏：窗口控制 + 边缘缩放 ----------------
    def _ensure_taskbar_entry(self):
        """自绘标题栏（无边框）默认不会出现在任务栏，补上 WS_EX_APPWINDOW。"""
        try:
            if not self.FRAMELESS:
                return
            hwnd = ctypes.c_void_p(int(self.winId()))
            GWL_EXSTYLE = -20
            WS_EX_APPWINDOW = 0x00040000
            user32 = ctypes.windll.user32
            user32.GetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int]
            user32.GetWindowLongW.restype = ctypes.c_long
            user32.SetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_long]
            user32.SetWindowLongW.restype = ctypes.c_long
            ex = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            if not (ex & WS_EX_APPWINDOW):
                user32.SetWindowLongW(hwnd, GWL_EXSTYLE, ex | WS_EX_APPWINDOW)
        except Exception:
            pass

    def _work_area(self):
        """当前屏幕的可用区域（去掉任务栏）。"""
        try:
            scr = self.screen() or QApplication.primaryScreen()
            return scr.availableGeometry()
        except Exception:
            return QRect(0, 0, 1500, 900)

    def toggle_maximize(self):
        """最大化 / 还原（自绘标题栏没有系统按钮，只能自己实现）。"""
        try:
            if self._maximized:
                if self._normal_geom is not None:
                    self.setGeometry(self._normal_geom)
                self._maximized = False
            else:
                self._normal_geom = QRect(self.geometry())
                self.setGeometry(self._work_area())
                self._maximized = True
            try:
                self.title_strip.set_maximized(self._maximized)
            except Exception:
                pass
        except Exception:
            pass

    def _win_minimize(self):
        try:
            self.showMinimized()
        except Exception:
            pass

    def _win_close(self):
        try:
            self.close()
        except Exception:
            pass

    def _hit_edge(self, gpos):
        """判断鼠标是否在窗口边缘，返回方向串（l/r/t/b 组合），不在边缘返回空串。"""
        try:
            r = self.frameGeometry()
            m = self.RESIZE_MARGIN
            x, y = gpos.x(), gpos.y()
            left = abs(x - r.left()) <= m
            right = abs(x - r.right()) <= m
            top = abs(y - r.top()) <= m
            bottom = abs(y - r.bottom()) <= m
            if not (left or right or top or bottom):
                return ''
            if top and left:
                return 'tl'
            if top and right:
                return 'tr'
            if bottom and left:
                return 'bl'
            if bottom and right:
                return 'br'
            if left:
                return 'l'
            if right:
                return 'r'
            if top:
                return 't'
            return 'b'
        except Exception:
            return ''

    _EDGE_CURSORS = {
        'l': Qt.SizeHorCursor, 'r': Qt.SizeHorCursor,
        't': Qt.SizeVerCursor, 'b': Qt.SizeVerCursor,
        'tl': Qt.SizeFDiagCursor, 'br': Qt.SizeFDiagCursor,
        'tr': Qt.SizeBDiagCursor, 'bl': Qt.SizeBDiagCursor,
    }

    def _set_edge_cursor(self, direction):
        """边缘显示调整大小的鼠标指针。"""
        try:
            if direction:
                if not self._cursor_overridden:
                    QApplication.setOverrideCursor(QCursor(self._EDGE_CURSORS.get(direction, Qt.ArrowCursor)))
                    self._cursor_overridden = True
            elif self._cursor_overridden:
                QApplication.restoreOverrideCursor()
                self._cursor_overridden = False
        except Exception:
            pass

    def _apply_resize(self, gpos):
        """按鼠标位移调整窗口大小。"""
        try:
            g = QRect(self._rs_geom)
            dx = gpos.x() - self._rs_gpos.x()
            dy = gpos.y() - self._rs_gpos.y()
            d = self._rs_dir
            if 'l' in d:
                if g.width() - dx >= self.minimumWidth():
                    g.setLeft(g.left() + dx)
            if 'r' in d:
                g.setRight(g.right() + dx)
            if 't' in d:
                if g.height() - dy >= self.minimumHeight():
                    g.setTop(g.top() + dy)
            if 'b' in d:
                g.setBottom(g.bottom() + dy)
            self.setGeometry(g)
        except Exception:
            pass

    def eventFilter(self, obj, ev):
        """应用级事件过滤：窗口边缘的拖拽缩放 + 指针形状。"""
        try:
            # 独立弹窗（单位下拉框、菜单、提示气泡等）不能参与主窗口的边缘缩放：
            # 弹窗是独立顶层窗口，其选项若恰好落在主窗口边缘 6px 内，点击会被
            # 下面的边缘检测 return True 吃掉，表现为“下拉选项点不中”。
            # 主窗口自身及其子控件（标题栏等）的 window() 就是 self，照常处理。
            if isinstance(obj, QWidget):
                top = obj.window()
                if top is not None and top is not self:
                    return False
            et = ev.type()
            # 双击单元格进入编辑态后，Ctrl+V 事件被单元格编辑器（QLineEdit）消费，
            # 到不了表格的 keyPressEvent，于是多行数据只塞进当前单元格、行数也不会
            # 自动扩展。这里接管：若焦点在表格编辑器里且剪贴板是多行/多列(TSV)，
            # 先关掉编辑器再走自定义粘贴，由 paste_text 负责自动加行。
            if (et == QEvent.KeyPress and ev.key() == Qt.Key_V
                    and (ev.modifiers() & Qt.ControlModifier)
                    and isinstance(obj, QLineEdit)):
                p = obj
                in_table = False
                while p is not None:
                    if p is self.table:
                        in_table = True
                        break
                    p = p.parent()
                if in_table:
                    clip = QApplication.clipboard().text()
                    if clip and ('\n' in clip or '\r' in clip or '\t' in clip):
                        try:
                            self.table.closeEditor(obj, QAbstractItemDelegate.NoHint)
                        except Exception:
                            pass
                        self.paste_text(clip)
                        return True
            if et == QEvent.MouseButtonPress and ev.button() == Qt.LeftButton:
                if not self._maximized:
                    d = self._hit_edge(ev.globalPos())
                    if d:
                        self._rs_dir = d
                        self._rs_gpos = ev.globalPos()
                        self._rs_geom = QRect(self.geometry())
                        return True          # 吃掉这次点击，避免影响控件
            elif et == QEvent.MouseMove:
                if self._rs_dir:
                    self._apply_resize(ev.globalPos())
                    return True
                if not self._maximized:
                    self._set_edge_cursor(self._hit_edge(ev.globalPos()))
                else:
                    self._set_edge_cursor('')
            elif et == QEvent.MouseButtonRelease:
                if self._rs_dir:
                    self._rs_dir = ''
                    self._rs_gpos = None
                    self._rs_geom = None
                    return True
            elif et == QEvent.Leave:
                self._set_edge_cursor('')
        except Exception:
            pass
        return super(MainWindow, self).eventFilter(obj, ev)

    def remember_file(self, path):
        """记住当前工程文件，下次启动自动打开。"""
        self.last_file = path or ''
        self._set_setting('last_file', self.last_file)

    @staticmethod
    def name_from_path(path):
        return os.path.splitext(os.path.basename(path))[0]

    def restore_last_project(self):
        """启动时自动恢复上次保存/打开的工程，避免标题一直是“未命名工程”。"""
        path = self.last_file
        if not path or not os.path.isfile(path):
            return False
        try:
            self.project = Project.load(path)
            if (self.project.name or '').strip() in ('', '未命名工程'):
                self.project.name = self.name_from_path(path)
            self.current_file = path
            self.current_section = self.project.root
            # 上次选中的节点（例如数据所在的分项工程），恢复后直接打开它
            sid = self.settings.value('last_section', '') or ''
            sec = self.find_section(sid) if sid else None
            if sec is not None and self._section_has_data(sec):
                self.current_section = sec
            else:
                # 没记节点（或记的节点是空壳）时，自动落到内容最多的分部/分项，
                # 否则一打开就是空表，用户会以为数据丢了
                self.current_section = (self._pick_section_with_rows(None)
                                       or self.project.root)
            self._dirty = False
            self.refresh_tree()
            self.refresh_table()
            self.update_title()
            self.status.showMessage('已恢复上次打开的工程：%s' % os.path.basename(path), 5000)
            return True
        except Exception:
            return False

    # ---------------- UI ----------------
    @staticmethod
    def _section_has_data(sec):
        """节点里是否有“真正有内容”的行（不是空占位行）。"""
        for r in (getattr(sec, 'rows', None) or []):
            if any(str(getattr(r, k, '') or '').strip()
                   for k in ('code', 'name', 'formula', 'part', 'remark')):
                return True
        return False

    def _pick_section_with_rows(self, sec=None):
        """当前节点没有内容时，自动落到内容最多的分部/分项。

        打开工程后如果停在空的根节点（或只有空占位行）上，工作区会是一片空白，
        用户容易以为数据丢了，所以统一跳到真正有数据的节点（行最多的那个）。
        """
        if sec is not None and self._section_has_data(sec):
            return sec
        best = None
        try:
            for s in self.project.all_sections():
                if not self._section_has_data(s):
                    continue
                if best is None or len(s.rows) > len(best.rows):
                    best = s
        except Exception:
            pass
        return best or sec

    def _main_qss(self, table_pt=None):
        """主窗口样式表。表格字号参数化：表格字体设置（V2.8.2）动态重建，
        否则 QSS 里写死的 font-size 会覆盖 setFont。"""
        if table_pt is None:
            table_pt = getattr(self, '_font_size', 0) or 10
        return """
            QWidget { font-size: 10pt; }
            QMainWindow { background: #f5f6f8; }
            QTreeWidget { background: #ffffff; alternate-background-color: #f7f9fc;
                                        font-size: 10pt; }
            QTableWidget { background: #ffffff; alternate-background-color: #f7f9fc;
                                        font-size: %dpt; }
            QHeaderView::section { background: #f0e6c8; padding: 4px 5px;
                                   border: 1px solid #d8cfae; font-weight: 500;
                                   font-size: 10pt; color:#4a4020; }
            QTableWidget::item { padding: 3px 4px; }
            QTableWidget::item:selected { background: #f5e6a8; color: #222; }
            QTableWidget::item:hover { background: #faf6e2; }
            QLineEdit, QTextEdit { border: 1px solid #cfd4da; border-radius: 2px;
                                   padding: 3px 5px; font-size: 10pt; }
            QPushButton { border: 1px solid #cfd4da; border-radius: 3px;
                          padding: 3px 9px; background:#fff; font-size: 10pt; }
            QPushButton:hover { background:#eef3f8; }
            QLabel { font-size: 10pt; }
            QMenu { font-size: 10pt; }
            QMenu::item { padding: 4px 18px; }
        """ % int(max(6, min(24, table_pt)))

    def build_ui(self):
        self.setStyleSheet(self._main_qss())

        central = QWidget()
        central.setMouseTracking(True)     # 边缘拖拽缩放需要持续收到鼠标移动
        cv = QVBoxLayout(central)
        cv.setContentsMargins(0, 0, 0, 0)
        cv.setSpacing(0)

        # ===== 最上面一条：自绘标题栏 =====
        # 左：快捷图标 + 工程名；右：图标 软件名版本 | 工程文件名称 | <文件位置> + 窗口按钮
        self.title_strip = TitleStrip([
            ('新建工程', 'new', self.new_project),
            ('打开工程', 'open', self.open_project),
            ('保存工程', 'save', self.save_project),
            ('打印', 'print', self.print_report),
        ], self, icon_path=getattr(self, 'app_icon_path', ''))
        try:
            self.title_strip.minimizeRequested.connect(self._win_minimize)
            self.title_strip.maximizeRequested.connect(self.toggle_maximize)
            self.title_strip.closeRequested.connect(self._win_close)
        except Exception:
            pass
        cv.addWidget(self.title_strip, 0)
        self.update_title()

        # ===== Ribbon 工具栏 =====
        sections = {
            '计算书编辑': [
                ('文件', [
                    ('新建', 'new', self.new_project, 'Ctrl+N', True),
                    ('打开', 'open', self.open_project, 'Ctrl+O', True),
                    ('保存', 'save', self.save_project, 'Ctrl+S', True),
                    ('另存为', 'copy', self.save_as, '', False),
                    ('最近打开', 'open', self.show_recent, '', False),
                ]),
                ('编辑', [
                    ('插入计算式', 'add', self.insert_row_below, 'Ctrl+I', True),
                    ('删除', 'delete', self.delete_rows, 'Ctrl+Delete', True),
                    ('剪切', 'copy', self.cut_rows, 'Ctrl+X', False),
                    ('复制', 'copy', self.copy_rows, 'Ctrl+C', False),
                    ('粘贴', 'copy', self.paste_from_clipboard, 'Ctrl+V', True),
                ]),
                ('编辑2', [
                    ('上移', 'up', self.move_section_up, '', False),
                    ('下移', 'down', self.move_section_down, '', False),
                    ('升级', 'up', self.promote_section, '', False),
                    ('降级', 'down', self.demote_section, '', False),
                    ('层级设置', 'settings', self.level_setup, '', False),
                    ('特殊符号', 'settings', self.show_symbol_picker, '', False),
                ]),
                ('编辑3', [
                    ('恢复', 'redo', self.redo, 'Ctrl+Y', False),
                    ('撤销', 'undo', self.undo, 'Ctrl+Z', False),
                    ('重命名', 'settings', self.rename_section, 'F2', False),
                    ('查找', 'settings', self.find_replace, 'Ctrl+F', False),
                ]),
                ('格式', [
                    ('填充颜色', '▤', self.fill_color_dialog, '', False),
                    ('清除填充', '▢', self.clear_fill, '', False),
                    ('表格字体', 'settings', self.font_setup, '', False),
                ]),
                ('计算书', [
                    ('导入表格', 'open', self.import_table, '', True),
                    ('分类汇总', 'calc', self.data_summary, '', True),
                    ('格式刷', 'copy', self.start_format_painter, '', False),
                    ('排序', 'settings', self.sort_dialog, '', False),
                    ('筛选', 'settings', self.filter_dialog, '', False),
                    ('数据列定义', 'settings', self.column_setup, '', False),
                    ('选项配置', 'settings', self.autosave_settings, '', False),
                ]),
            ],
            '公式编辑': [
                ('公式', [
                    ('插入函数', 'settings', self.show_functions, '', True),
                    ('查看公式', 'doc', self.view_formula, '', False),
                    ('重算', 'calc', self.recalc_all, 'F9', True),
                ]),
                ('括号符号', [
                    ('F4 前置()', '( )', lambda: self.insert_bracket('pre'), '', False),
                    ('F5 插入()', '( )', lambda: self.insert_bracket('cur'), '', False),
                    ('F6 插入[]', '[ ]', lambda: self.insert_bracket('cur_sq'), '', False),
                    ('F7 末尾[]', '[ ]', lambda: self.insert_bracket('end_sq'), '', False),
                ]),
            ],
            '数据输出': [
                ('导出', [
                    ('导出Excel表格', 'export', self.export_excel, '', True),
                    ('导出WPS表格', 'export', self.export_wps, '', True),
                ]),
                ('打印', [
                    ('打印', 'print', self.print_report, 'Ctrl+P', True),
                    ('打印预览', 'print', self.print_preview, '', True),
                    ('页面设置', 'settings', self.page_setup, '', True),
                ]),
            ],
            '关于产品': [
                ('关于', [
                    ('使用说明', 'doc', self.show_formula_help, '', True),
                    ('关于产品', 'settings', self.show_about, '', False),
                ]),
            ],
        }
        self.ribbon = RibbonWidget(self.app, sections, self)
        cv.addWidget(self.ribbon, 0, Qt.AlignTop)

        # 新建按钮：点击即弹“新建工程”（同时填写工程信息），不再挂下拉菜单
        for g in self.ribbon.pages['计算书编辑'].groups:
            for b in g.buttons:
                if b.text() == '新建':
                    self.btn_new = b
                    b.setToolTip('新建工程（Ctrl+N）—— 可填工程名称、建设单位、施工单位、'
                                 '设计单位、编制日期')
        # 说明：撤销/恢复/查找的快捷键由工具条按钮自己注册（Ctrl+Z / Ctrl+Y / Ctrl+F），
        # 这里是窗口级生效的，不需要再额外注册 QShortcut（重复注册会互相冲突）。

        # F4~F7 括号快捷键：注册为窗口级快捷键，任何页签、任何输入位置都生效
        self.bracket_shortcuts = []
        for key, mode, label in (('F4', 'pre', '公式前插入 ()'),
                                 ('F5', 'cur', '插入点插入 ()'),
                                 ('F6', 'cur_sq', '插入点插入 []'),
                                 ('F7', 'end_sq', '末尾追加 []')):
            sc = QShortcut(QKeySequence(key), self)
            sc.setContext(Qt.WindowShortcut)
            sc.activated.connect(lambda m=mode: self.insert_bracket(m))
            sc.setWhatsThis(label)
            self.bracket_shortcuts.append(sc)


        # 主工作区（三个页签内容，但工具条常驻）
        self.tabs = QTabWidget()
        self.tabs.tabBar().hide()  # 页签由 Ribbon 控制
        cv.addWidget(self.tabs, 1)
        self.setCentralWidget(central)

        # ===== 编辑页 =====
        page1 = QWidget()
        v1 = QVBoxLayout(page1)
        v1.setContentsMargins(2, 2, 2, 2)
        self.tabs.addTab(page1, '计算书编辑')

        # 左导航 + 右（公式栏 + 表格）左右分栏：
        # 导航栏从工作区顶部开始，公式辅助栏的左边与表格左边严格对齐
        main_split = QSplitter(Qt.Horizontal)
        # 三块之间只用一条 1px 细线分隔（省出横向空间给工作区）
        main_split.setHandleWidth(1)
        main_split.setChildrenCollapsible(False)
        main_split.setStyleSheet('QSplitter::handle { background:#dde3ea; }')
        v1.addWidget(main_split, 1)

        # 公式辅助输入栏（表格上方；可拖动分隔条调节高度）
        form_frame = QFrame()
        form_frame.setStyleSheet('background:#f0f0f0; border:1px solid #d8d8d8;')
        form_v = QVBoxLayout(form_frame)
        form_v.setContentsMargins(4, 2, 4, 2)
        form_v.setSpacing(2)

        # —— 快捷操作链接行 ——
        quick_bar = QFrame()
        quick_bar.setStyleSheet('QFrame { background:#eef2f7; border-bottom:1px solid #d6dbe1; }')
        qb = QHBoxLayout(quick_bar)
        qb.setContentsMargins(3, 0, 3, 0)
        qb.setSpacing(1)
        # 左侧导航收起后，这里出现“展开导航”按钮
        self.btn_nav_open = QPushButton('▶ 导航')
        self.btn_nav_open.setObjectName('quickLink')
        self.btn_nav_open.setFlat(True)
        self.btn_nav_open.setStyleSheet('QPushButton#quickLink { border:1px solid #c9d6e3;'
                                        ' border-radius:2px; background:#ffffff;'
                                        ' color:#24507a; font-size:8.5pt; padding:0px 4px; }'
                                        'QPushButton#quickLink:hover { background:#dce8f5; }')
        self.btn_nav_open.setToolTip('展开左侧导航栏')
        self.btn_nav_open.setVisible(False)
        self.btn_nav_open.clicked.connect(lambda: self.fold_nav(False))
        qb.addWidget(self.btn_nav_open)
        qb.addSpacing(4)
        quick_actions = [
            ('分类汇总', self.data_summary),
            ('打印选中区域', self.set_print_area),
            ('插入函数', self.show_function_picker),
            ('树形显示', self.toggle_tree_expand),
            ('选中行高亮', self.toggle_row_highlight, self._hilite_row),
            ('选中列高亮', self.toggle_col_highlight, self._hilite_col),
            ('F4 前置()', lambda: self.insert_bracket('pre')),
            ('F5 插入()', lambda: self.insert_bracket('cur')),
            ('F6 插入[]', lambda: self.insert_bracket('cur_sq')),
            ('F7 末尾[]', lambda: self.insert_bracket('end_sq')),
            ('数据列定义', self.column_setup),
            ('列宽自适应', self.fit_column_widths),
            ('行高自适应', self.fit_row_heights),
        ]
        self._hilite_buttons = {}
        for item in quick_actions:
            text, slot = item[0], item[1]
            b = QPushButton(text)
            b.setObjectName('quickLink')
            b.setFlat(True)
            b.setStyleSheet('QPushButton#quickLink { border:1px solid transparent; background:transparent;'
                            ' color:#24507a; font-size:8.5pt; padding:0px 3px; text-align:left; }'
                            ' QPushButton#quickLink:hover { background:#dce8f5; border:1px solid #b8cde3; }'
                            ' QPushButton#quickLink:checked { background:#cfe0f2; border:1px solid #7ba3cc; }')
            b.setMinimumWidth(0)
            try:
                b.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
            except Exception:
                pass
            if len(item) >= 3:
                # 可勾选开关（列高亮）：用 toggled 传勾选态，按钮外观即启用状态
                b.setCheckable(True)
                b.setChecked(bool(item[2]))
                self._hilite_buttons[text] = b
                b.toggled.connect(lambda checked=False, sl=slot: sl(checked))
            else:
                b.clicked.connect(lambda checked=False, sl=slot: sl())
            qb.addWidget(b)
        qb.addStretch(1)
        # 右侧分类汇总收起后，这里出现“展开汇总”按钮
        self.btn_sum_open = QPushButton('汇总 ▶')
        self.btn_sum_open.setObjectName('quickLink')
        self.btn_sum_open.setFlat(True)
        self.btn_sum_open.setStyleSheet('QPushButton#quickLink { border:1px solid #c9d6e3;'
                                        ' border-radius:2px; background:#ffffff;'
                                        ' color:#24507a; font-size:8.5pt; padding:0px 4px; }'
                                        'QPushButton#quickLink:hover { background:#dce8f5; }')
        self.btn_sum_open.setToolTip('展开右侧分类汇总面板')
        self.btn_sum_open.setVisible(False)
        self.btn_sum_open.clicked.connect(lambda: self.fold_summary(False))
        qb.addWidget(self.btn_sum_open)
        # 快捷链接行放进横向滚动区：链接再多也不会把整块工作区的最小宽度撑大
        # （否则左右面板会被挤到最窄）
        quick_scroll = QuickScroll(quick_bar)
        form_v.addWidget(quick_scroll)

        fh = QHBoxLayout()
        fh.setContentsMargins(2, 3, 2, 2)
        fh.setSpacing(4)
        lb_formula = QLabel('公式辅助输入（仅公式列）：')
        lb_formula.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        lb_formula.setContentsMargins(0, 4, 0, 0)
        fh.addWidget(lb_formula)
        self.formula_help = FormulaEdit(self.apply_formula_help)
        self.formula_help.setPlaceholderText('选中"计算公式"列的单元格，这里会同步显示并着色；也可直接在此输入')
        self.formula_help.setStyleSheet('background:#f7f7f7; color:#888;')
        self.formula_help.setMaximumHeight(64)
        self.formula_help.setReadOnly(True)
        self.formula_help.text_changed.connect(self.on_formula_help_text)
        self.formula_highlighter = FormulaHighlighter(self.formula_help.document())
        fh.addWidget(self.formula_help, 1)
        btn_write = QPushButton('写入')
        btn_write.clicked.connect(self.apply_formula_help)
        fh.addWidget(btn_write)
        form_v.addLayout(fh)
        self.formula_dock = form_frame

        # ===== 左：项目结构（导航栏，顶到工作区最上沿）=====
        left = QFrame()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(0, 0, 0, 0)
        # 标题行：左边“项目结构”，右边一个收起按钮（分隔改成 1px 细线后按钮挪到这里）
        lhead = QHBoxLayout()
        lhead.setContentsMargins(4, 1, 3, 1)
        lhead.setSpacing(4)
        lb_tree = QLabel('项目结构')
        lhead.addWidget(lb_tree)
        lhead.addStretch(1)
        self.btn_nav_fold = self._mini_btn('◀', '收起左侧导航（工作区变宽）')
        self.btn_nav_fold.clicked.connect(lambda: self.fold_nav(True))
        lhead.addWidget(self.btn_nav_fold)
        lv.addLayout(lhead)
        self.tree = QTreeWidget()
        # 两列：名称 + 层级（单项工程 / 单位工程 / 分部工程 / 分项工程，可在“层级设置”里改名）
        self.tree.setColumnCount(2)
        self.tree.setHeaderLabels(['名称', '层级'])
        self.tree.setMinimumWidth(200)
        th = self.tree.header()
        # 分隔线可手动拖动：名称列 Interactive，层级列吃掉剩余宽度
        th.setStretchLastSection(True)
        th.setSectionResizeMode(0, QHeaderView.Interactive)
        th.setSectionResizeMode(1, QHeaderView.Interactive)
        self.tree.setColumnWidth(0, 150)
        self.tree.setColumnWidth(1, 84)
        # 导航栏（含表头）统一常规字重，不用加粗
        self.tree.setStyleSheet('QHeaderView::section { font-weight: 400; }')
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.tree_menu)
        self.tree.currentItemChanged.connect(self.on_tree_select)
        self.tree.itemDoubleClicked.connect(self._on_tree_double)
        # 导航树与常量框之间放可拖动分隔条：常量框高度可自由拉伸
        left_split = QSplitter(Qt.Vertical)
        left_split.setChildrenCollapsible(False)
        tree_box = QWidget()
        tb = QVBoxLayout(tree_box)
        tb.setContentsMargins(0, 0, 0, 0)
        tb.addWidget(self.tree, 1)
        left_split.addWidget(tree_box)

        # 导航树 / 常量框之间的折叠条：条上两个小按钮，可收起 / 展开常量框
        self.const_strip = FoldStrip(vertical=False, fold_text='▾', open_text='▴',
                                     fold_tip='收起常量框（导航栏变高）',
                                     open_tip='展开常量框')
        self.const_strip.btn_fold.clicked.connect(lambda: self.fold_const(True))
        self.const_strip.btn_open.clicked.connect(lambda: self.fold_const(False))
        left_split.addWidget(self.const_strip)

        const_box = QWidget()
        cbv = QVBoxLayout(const_box)
        cbv.setContentsMargins(0, 0, 0, 0)
        cbv.setSpacing(1)
        # 常量框上方一排“新建常量 / 删除 / 导出常量”
        chead = QHBoxLayout()
        chead.setContentsMargins(4, 1, 2, 1)
        chead.setSpacing(4)
        self.btn_const_add = self._link_btn('新建常量', '新建一条常量（也可双击表格空白处）',
                                            self.add_constant, 'ab')
        self.btn_const_del = self._link_btn('删除', '删除选中的常量行', self.del_constant, '✕')
        self.btn_const_exp = self._link_btn('导出常量', '把常量导出成 CSV（WPS/Excel 可打开）',
                                            self.export_constants, '⇩')
        for b in (self.btn_const_add, self.btn_const_del, self.btn_const_exp):
            chead.addWidget(b)
        chead.addStretch(1)
        self.btn_const_more = self._mini_btn('▾', '更多：批量编辑 / 插入钢筋单重模板 / 清空')
        self.btn_const_more.setPopupMode(QToolButton.InstantPopup)
        self.btn_const_more.setMenu(self.const_menu())
        chead.addWidget(self.btn_const_more)
        cbv.addLayout(chead)
        self.const_list = ConstGrid(self.edit_constants)
        self.const_list.setMinimumHeight(48)
        self.const_list.setToolTip('常量全部由自己定义：双击单元格改代码 / 值 / 说明\n'
                                   '右键：新建、删除、批量编辑\n'
                                   '公式用法：@层高×4、@Φ8×12')
        self.const_list.changed.connect(self.on_constants_changed)
        cbv.addWidget(self.const_list, 1)
        left_split.addWidget(const_box)
        left_split.setStretchFactor(0, 1)
        left_split.setStretchFactor(2, 0)
        left_split.setHandleWidth(5)
        sizes = self.settings.value('left_split_sizes')
        try:
            sizes = [int(x) for x in (sizes or [])]
        except (TypeError, ValueError):
            sizes = []
        if len(sizes) != 3 or min(sizes) <= 0:
            sizes = [520, 19, 180]
        left_split.setSizes(sizes)
        self._const_h = sizes[2]          # 常量框高度（收起 / 展开时复原用）
        self.left_split = left_split
        self.const_box = const_box
        lv.addWidget(left_split, 1)
        left.setMinimumWidth(220)
        self.left_frame = left
        main_split.addWidget(left)

        # ===== 右：公式辅助栏 + 表格（上下分栏，公式栏左边与表格对齐）=====
        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(0)
        right_split = QSplitter(Qt.Vertical)
        rv.addWidget(right_split, 1)
        right_split.addWidget(form_frame)

        table_host = QWidget()
        th = QVBoxLayout(table_host)
        th.setContentsMargins(0, 0, 0, 0)
        th.setSpacing(0)
        self.table = PasteTableWidget(self.paste_text, self.insert_row_below, self.delete_selection)
        self.table.on_copy = self.copy_cells
        self.table.on_cut = self.cut_rows
        self.table.on_escape = self.cancel_painter
        self.table.on_fill_apply = self.apply_fill_drag
        self.table.on_paint_apply = self.painter_apply
        # 自定义表头：用户一旦手动拖列宽，就自动关闭“列宽自适应”，避免布局抢回去
        fhdr = FitHeader(Qt.Horizontal, self.table)
        fhdr.on_manual_resize = self._col_autofit_off
        self.table.setHorizontalHeader(fhdr)
        self.table.on_resize = self._on_table_resized
        self.table.setColumnCount(len(self.project.columns))
        self.table.setHorizontalHeaderLabels([c['title'] for c in self.project.columns])
        self.table.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.horizontalHeader().customContextMenuRequested.connect(self.header_menu)
        # 行号列右键：插入/删除行、行底色、行高自适应
        self.table.verticalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.verticalHeader().customContextMenuRequested.connect(self.row_menu)
        # 表格左上角（全选角）右键：删除全部数据 / 列宽行高自适应
        corner = self.table.findChild(QAbstractButton)
        if corner is not None:
            corner.setToolTip('点击全选；右键可“删除全部数据”、设置列宽/行高自适应')
            corner.setContextMenuPolicy(Qt.CustomContextMenu)
            corner.customContextMenuRequested.connect(
                lambda p, w=corner: self.corner_menu(w.mapToGlobal(p)))
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        # 单击只选中、不进编辑：Ctrl+C 直接复制选中区域（WPS/Excel 同款）。
        # 双击 / F2 才进入单元格编辑。
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked |
                                   QAbstractItemView.EditKeyPressed)
        # 实时双向同步：公式列单元格编辑时逐字同步到辅助框
        self.table.setItemDelegate(LiveSyncDelegate(self))
        self.table.itemChanged.connect(self.on_cell_changed)
        self.table.currentCellChanged.connect(self.on_cell_cursor)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.table_menu)
        # 列可拖动
        self.table.horizontalHeader().setSectionsMovable(True)
        self.table.horizontalHeader().setStretchLastSection(False)
        # 仿 WPS：双击列/行分隔线即自适应该列列宽 / 该行行高
        self.table.horizontalHeader().sectionHandleDoubleClicked.connect(
            lambda idx: self.fit_column_widths([idx]))
        self.table.verticalHeader().sectionHandleDoubleClicked.connect(
            lambda idx: self.fit_row_heights([idx]))
        # 列宽变化（手动拖动 / 自动排版）后，自动行高模式下按新宽度重新折行
        self.table.horizontalHeader().sectionResized.connect(self._on_col_width_changed)

        # 表格直接占满右侧（备注就在表格“备注”列里编辑）
        th.addWidget(self.table, 1)
        right_split.addWidget(table_host)
        right_split.setStretchFactor(0, 0)
        right_split.setStretchFactor(1, 1)
        right_split.setSizes([86, 700])
        self.right_split = right_split
        self.work_pane = right
        main_split.addWidget(right)

        # ===== 最右：分类汇总面板 =====
        self.summary_panel = SummaryPanel(self)
        main_split.addWidget(self.summary_panel)

        main_split.setStretchFactor(0, 0)
        main_split.setStretchFactor(1, 1)
        main_split.setStretchFactor(2, 0)
        # 记住用户拖出来的宽度，收起 / 展开时复原
        self._nav_folded = False
        self._sum_folded = False
        self._sum_max = False
        self._const_folded = False
        self._left_width = int(self.settings.value('main_left_w', 300) or 300)
        self._sum_width = int(self.settings.value('main_sum_w', 300) or 300)
        if os.environ.get('CALCBOOK_NO_RESTORE'):
            # 自动化测试时不读用户的面板宽度，避免被历史脏数据影响
            self._left_width = 300
            self._sum_width = 300
        self._applying_sizes = True
        try:
            main_split.setSizes([self._left_width, 1000, self._sum_width])
        finally:
            self._applying_sizes = False
        main_split.splitterMoved.connect(self._remember_split_widths)
        self.main_split = main_split

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.lbl_status = QLabel('就绪')
        # 放到状态栏右侧（永久部件）：临时提示在左侧显示，两者不会叠在一起
        self.lbl_status.setStyleSheet('color:#555; padding-right:6px;')
        self.status.addPermanentWidget(self.lbl_status)

        # 报表页：工具条 + 可拖动页边距的纸张预览
        page2 = QWidget()
        v2 = QVBoxLayout(page2)
        v2.setContentsMargins(0, 0, 0, 0)
        self.preview_pane = PrintPreviewPane(self)
        self.report_preview = self.preview_pane   # 兼容旧引用
        v2.addWidget(self.preview_pane)
        self.tabs.addTab(page2, '报表导出')

        # ===== 使用说明页（内嵌帮助） =====
        page3 = QWidget()
        v3 = QVBoxLayout(page3)
        self.help_view = QTextEdit()
        self.help_view.setReadOnly(True)
        self.help_view.setHtml(self.help_html())
        v3.addWidget(self.help_view)
        self.tabs.addTab(page3, '使用说明')

        # Ribbon 页签切换联动主工作区
        self.ribbon.tab_group.buttonClicked.connect(self.on_ribbon_tab)

    # ---------------- 面板收起 / 展开（分割条上的小按钮）----------------
    def _mini_btn(self, text, tip, slot=None):
        """左侧常量框标题行上的小方块按钮。"""
        b = QToolButton(self)
        b.setText(text)
        b.setToolTip(tip)
        b.setCursor(Qt.PointingHandCursor)
        b.setAutoRaise(True)
        b.setFixedSize(20, 20)
        b.setStyleSheet('QToolButton { border:1px solid #d5dbe2; border-radius:2px; '
                        'background:#fff; color:#3c6b9c; font-size:10pt; padding:0; }'
                        'QToolButton:hover { background:#e8f0fa; border-color:#9dbce0; }')
        if slot is not None:
            b.clicked.connect(lambda checked=False: slot())
        return b

    def _link_btn(self, text, tip, slot, icon=''):
        """文字按钮（常量框上方“新建常量 / 删除 / 导出常量”）。"""
        b = QToolButton(self)
        b.setText(('%s %s' % (icon, text)) if icon else text)
        b.setToolTip(tip)
        b.setCursor(Qt.PointingHandCursor)
        b.setAutoRaise(True)
        b.setStyleSheet('QToolButton { border:none; background:transparent; color:#24507a;'
                        ' font-size:9pt; padding:1px 3px; }'
                        'QToolButton:hover { background:#dce8f5; border-radius:2px; }')
        b.clicked.connect(lambda checked=False: slot())
        return b

    def const_menu(self):
        """常量框“⋯”菜单：批量编辑 / 钢筋单重模板 / 清空。"""
        m = QMenu(self)
        m.addAction('批量编辑 / 粘贴常量…', self.edit_constants)
        m.addAction('插入常用钢筋单重（14 条模板）', self.insert_rebar_template)
        m.addSeparator()
        m.addAction('清空全部常量', self.clear_constants)
        return m

    def showEvent(self, event):
        super().showEvent(event)
        # 自绘标题栏（无边框）默认不在任务栏显示，这里补上 WS_EX_APPWINDOW；
        # showEvent 之后 Qt 还会再套一次窗口样式，所以延后一拍再补一次
        self._ensure_taskbar_entry()
        try:
            QTimer.singleShot(0, self._ensure_taskbar_entry)
        except Exception:
            pass
        # 首次显示后按真实窗口宽度重排一次（build_ui 时窗口还没定尺寸，
        # 直接 setSizes 会被按比例压缩，导致左侧面板偏窄）
        if not getattr(self, '_sizes_applied', False):
            self._sizes_applied = True
            QTimer.singleShot(0, self._apply_initial_sizes)

    def _apply_initial_sizes(self, tries=0):
        """窗口真正布局好后再排一次分栏尺寸。

        坑：build_ui 里窗口还没 lay out，此时 setSizes 会被按比例压缩成很小的值，
        之后窗口变大只会把多出来的空间分给 stretch 部件（工作区），面板就保持极窄。
        所以这里要等 main_split 有真实宽度后再设置；不行就延时重试。
        """
        try:
            ok_w = self.main_split.width() >= max(600, int(self.width() * 0.7))
            ok_h = self.left_split.height() >= max(260, int(self.height() * 0.45))
            if not (ok_w and ok_h):
                if tries < 25:
                    QTimer.singleShot(60, lambda: self._apply_initial_sizes(tries + 1))
                return
            self._sync_main_sizes()
            self._sync_left_sizes()
        except Exception:
            log_exception('分栏初始尺寸')

    def _set_setting(self, key, value):
        """写设置。测试模式（CALCBOOK_NO_RESTORE=1）下不写，
        免得自动化测试把用户的窗口尺寸 / 面板宽度 / 自适应开关改掉。"""
        if os.environ.get('CALCBOOK_NO_RESTORE'):
            return
        try:
            self.settings.setValue(key, value)
        except Exception:
            pass

    def _remember_split_widths(self, pos, index):
        """记住用户拖出来的面板宽度 / 高度，收起后再展开能复原。

        注意：程序里调 setSizes() 时也可能触发 splitterMoved（Qt 实现细节），
        那样会把中间状态的宽度当成“用户拖出来的宽度”，所以只在真的按住鼠标拖动时记录。
        """
        if getattr(self, '_applying_sizes', False):
            return
        if not (QApplication.mouseButtons() & Qt.LeftButton):
            return
        try:
            s = self.main_split.sizes()
            if len(s) == 3:
                if s[0] >= 120:
                    self._left_width = s[0]
                if s[2] >= 120:
                    self._sum_width = s[2]
        except Exception:
            pass

    def _fold(self, which, hide):
        """统一收起 / 展开；用“折叠状态”标志判断，避免窗口未显示时 isVisible() 恒 False。"""
        attr = '_%s_folded' % which
        if hide == getattr(self, attr):
            return False
        setattr(self, attr, hide)
        if which == 'nav':
            self.left_frame.setVisible(not hide)
            self.btn_nav_open.setVisible(hide)      # 收起后工作区左上角出现“展开导航”
            self._sync_main_sizes()
        elif which == 'sum':
            self.summary_panel.setVisible(not hide)
            self.btn_sum_open.setVisible(hide)      # 收起后工作区右上角出现“展开汇总”
            self._sync_main_sizes()
            if not hide:
                self.summary_panel.refresh()
        else:
            self.const_box.setVisible(not hide)
            self.const_strip.set_folded(hide)
            self._sync_left_sizes()
        return True

    def fold_nav(self, hide=None):
        """收起 / 展开左侧导航栏。"""
        if hide is None:
            hide = not self._nav_folded
        if hide and not self._nav_folded and self.left_frame.width() > 60:
            self._left_width = self.left_frame.width()
        self._fold('nav', hide)

    def fold_summary(self, hide=None):
        """收起 / 展开右侧分类汇总面板。"""
        if hide is None:
            hide = not self._sum_folded
        if hide and getattr(self, '_sum_max', False):
            self.toggle_summary_max(False)
        if hide and not self._sum_folded and self.summary_panel.width() > 60:
            self._sum_width = self.summary_panel.width()
        self._fold('sum', hide)

    def toggle_summary_max(self, on=None):
        """放大 / 还原分类汇总：放大时隐藏左侧导航与工作区，汇总表占满窗口。"""
        if on is None:
            on = not getattr(self, '_sum_max', False)
        on = bool(on)
        if on == getattr(self, '_sum_max', False):
            return
        self._sum_max = on
        if on:
            if self._sum_folded:               # 放大前先确保汇总面板可见
                self._sum_folded = False
                self.summary_panel.setVisible(True)
                self.btn_sum_open.setVisible(False)
            if self.summary_panel.width() > 60:
                self._sum_width = self.summary_panel.width()
            self.left_frame.setVisible(False)
            self.work_pane.setVisible(False)
            self.btn_nav_open.setVisible(False)
            self.btn_sum_open.setVisible(False)
            self.summary_panel.btn_max.setText('⤡')
            self.summary_panel.btn_max.setToolTip('还原工作区')
            self._sync_main_sizes()
            self.summary_panel.refresh()
            self.status.showMessage('分类汇总已放大：主工作区暂时隐藏', 3000)
        else:
            self.left_frame.setVisible(not self._nav_folded)
            self.work_pane.setVisible(True)
            self.btn_nav_open.setVisible(self._nav_folded)
            self.btn_sum_open.setVisible(False)
            self.summary_panel.btn_max.setText('⤢')
            self.summary_panel.btn_max.setToolTip('放大分类汇总（隐藏主工作区）')
            self._sync_main_sizes()
            self.status.showMessage('已还原工作区', 2000)

    def fold_const(self, hide=None):
        """收起 / 展开左下常量框。"""
        if hide is None:
            hide = not self._const_folded
        if hide and not self._const_folded and self.const_box.height() > 60:
            self._const_h = self.const_box.height()
        self._fold('const', hide)

    def _sync_main_sizes(self):
        """左导航 / 工作区 / 汇总三块按记录宽度分配（窗口太窄时按比例让位）。"""
        total = self.main_split.width()
        if total < 480:
            return
        if getattr(self, '_sum_max', False):
            self._applying_sizes = True
            try:
                self.main_split.setSizes([0, 0, total])
            finally:
                self._applying_sizes = False
            return
        HANDLE = 1 * 2          # 两个 1px 细线分隔
        lw = self._left_width if not self._nav_folded else 0
        sw = self._sum_width if not self._sum_folded else 0
        room = total - HANDLE - 260           # 留给工作区的最小空间
        if room > 0 and lw + sw > room:
            k = room / float(lw + sw)
            lw, sw = int(lw * k), int(sw * k)
        ww = max(160, total - lw - sw - HANDLE)
        self._applying_sizes = True
        try:
            self.main_split.setSizes([lw, ww, sw])
        finally:
            self._applying_sizes = False

    def _sync_left_sizes(self):
        """导航树 / 常量框按记录高度分配。"""
        h = self.left_split.height()
        if h < 120:
            return
        ch = self._const_h if not self._const_folded else 0
        ch = max(0, min(ch, h - 120 - 19))
        th = max(60, h - ch - 19 - 2)
        self._applying_sizes = True
        try:
            self.left_split.setSizes([th, 19, ch])
        finally:
            self._applying_sizes = False

    # ---------------- 常量：自定义 / 增删 / 模板 ----------------
    def add_constant(self):
        self.const_list.add_row()

    def del_constant(self):
        if not self.const_list.remove_current_row():
            self.status.showMessage('请先在常量框里选中要删除的行', 3000)

    def on_constants_changed(self):
        """常量框里改了代码/数值/说明 → 写回工程并重算。"""
        if self._loading:
            return
        self.push_undo_coalesce('改常量')
        rows = self.const_list.rows()
        consts, notes = {}, {}
        for name, txt, note in rows:
            try:
                consts[name] = float(txt)
            except (TypeError, ValueError):
                continue
            if note:
                notes[name] = note
        self.project.constants = consts
        self.project.constant_notes = notes
        self.project.recalc()
        self._update_value_cells()
        self._dirty = True
        self.schedule_summary_refresh()
        self.status.showMessage('常量已更新（%d 条）' % len(consts), 2500)

    def insert_rebar_template(self):
        """插入常用钢筋单重模板（只是模板，用户可随意改 / 删）。"""
        consts = dict(self.project.constants or {})
        notes = dict(self.project.constant_notes or {})
        added = 0
        for k, v in REBAR_CONSTANTS:
            if k not in consts:
                consts[k] = v
                notes.setdefault(k, '钢筋单重 kg/m')
                added += 1
        if not added:
            self.status.showMessage('常用钢筋单重已经在常量里了', 3000)
            return
        self.push_undo('插入钢筋单重')
        self.project.constants = consts
        self.project.constant_notes = notes
        self.update_constants_display()
        self.project.recalc()
        self._update_value_cells()
        self._dirty = True
        self.status.showMessage('已插入 %d 条钢筋单重常量（可自行修改或删除）' % added, 4000)

    def export_constants(self):
        """导出常量（CSV，可用 WPS/Excel 打开，也可再粘回来）。"""
        if not self.project.constants:
            QMessageBox.information(self, APP_NAME, '还没有常量可导出。')
            return
        name = (self.project.name or '工程量') + '-常量.csv'
        path, _ = QFileDialog.getSaveFileName(self, '导出常量', name,
                                              'CSV 文件 (*.csv);;所有文件 (*)')
        if not path:
            return
        if not path.lower().endswith('.csv'):
            path += '.csv'
        try:
            import csv
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                w = csv.writer(f)
                w.writerow(['常量代码', '值', '说明'])
                for k, v in self.project.constants.items():
                    w.writerow([k, fmt_const(v),
                                (self.project.constant_notes or {}).get(k, '')])
        except Exception as e:
            QMessageBox.critical(self, '导出失败', str(e))
            return
        self.status.showMessage('常量已导出：' + os.path.basename(path), 5000)

    def clear_constants(self):
        if not self.project.constants:
            return
        if QMessageBox.question(self, '清空常量',
                                '确定要清空全部 %d 条常量吗？\n（公式里引用的 @代码 会红色提示）'
                                % len(self.project.constants),
                                QMessageBox.Yes | QMessageBox.No,
                                QMessageBox.No) != QMessageBox.Yes:
            return
        self.push_undo('清空常量')
        self.project.constants = {}
        self.project.constant_notes = {}
        self.update_constants_display()
        self.project.recalc()
        self._update_value_cells()
        self._dirty = True

    def on_ribbon_tab(self, btn):
        name = btn.text()
        self.ribbon.show_page(name)
        mapping = {'计算书编辑': 0, '公式编辑': 0, '数据输出': 1, '关于产品': 2}
        idx = mapping.get(name, 0)
        self.tabs.setCurrentIndex(idx)
        if idx == 1:
            self.print_preview()
        elif idx == 2:
            self.show_formula_help()

    def help_html(self):
        """软件内嵌「使用说明」页：按当前功能完整重写。"""
        return ('<h2 style="color:#2f5597;">工程量计算书 V%s · 使用说明</h2>' % APP_VER) + """
        <p>本软件用于工程量清单计算、分类汇总、报表打印与 Excel 导出，纯本地运行、无需联网。以下内容按当前界面功能整理。</p>

        <h3 style="color:#2f5597;">一、快速开始</h3>
        <ul>
        <li>双击 <code>工程量计算书.exe</code> 直接运行，免安装、无需 Python 环境。</li>
        <li>点工具栏「新建」，填写工程名称（必填）、建设单位、施工单位、设计单位、编制日期等，确定后即创建工程。</li>
        <li>在左侧树右键「新建分部 / 新建子分部」建立层级，再在表格逐行填写编号、项目名称、部位、计算公式、单位、备注。</li>
        <li>公式写完自动算出工程量，每个分部自动汇总「小计」；勾选「不计」的行不参与汇总。</li>
        <li>点「保存」生成 <code>.gcls</code> 工程文件；点「导出Excel表格」生成 xlsx，WPS / Excel 都能打开。</li>
        </ul>

        <h3 style="color:#2f5597;">二、界面布局与标题栏</h3>
        <ul>
        <li>最上面是本程序自绘标题栏：左端一排小图标为「新建 / 打开 / 保存 / 打印」，紧跟着是 软件图标 → 软件名与版本 → 工程文件名称 → <code>&lt;文件存放详细地址&gt;</code>；最右端是最小化 / 最大化 / 关闭。未保存的工程显示 <code>&lt;未保存&gt;</code>。</li>
        <li>标题栏空白处按住可拖动窗口，双击最大化 / 还原；鼠标移到窗口边缘可拖拽调整大小。路径过长时中间省略，鼠标悬停可看全路径。</li>
        <li>Ribbon 四个页签：<b>计算书编辑</b>（文件、编辑、格式、计算书等）、<b>公式编辑</b>（插入函数、查看公式、重算、括号符号）、<b>数据输出</b>（导出、打印、打印预览、页面设置）、<b>关于产品</b>（使用说明、关于产品）。</li>
        <li>左侧为项目结构树与常量框，中间为公式辅助框、快捷操作行和工作区表格，右侧为分类汇总面板。各面板分割条上有小箭头按钮可收起 / 展开（宽度会被记住）；汇总标题行的「⤢」可隐藏主工作区、放大查看汇总，再点「⤡」还原。</li>
        </ul>

        <h3 style="color:#2f5597;">三、项目结构与层级</h3>
        <ul>
        <li>导航树有两列「名称 / 层级」。第一级固定是<b>单项工程</b>（默认“单项工程1”，不可删除），往下依次为<b>单位工程 / 分部工程 / 分项工程</b>，可继续往下加，超过四级沿用最后一级名称。</li>
        <li>工具栏「层级设置」可改名、增删层级（例如改成 标段 / 单位工程 / 分部 / 分项）；新建节点自动使用下一级名称，如“分部工程2”。</li>
        <li>树右键：新建单项工程 / 新建下一级 / 新建同级 / 重命名 / 上移 / 下移 / 升级 / 降级 / 删除 / 层级设置 / 工程信息。右键会先选中该节点，不可用的项变灰；删除后可用 Ctrl+Z 撤销。</li>
        </ul>

        <h3 style="color:#2f5597;">四、表格录入与编辑</h3>
        <ul>
        <li><b>公式辅助框</b>：只在「计算公式」列生效。选中公式列单元格时公式同步显示在上方辅助框并着色；在表格里逐字输入会立即同步到辅助框，在辅助框输入也会实时更新到单元格和工程量。按 Ctrl+Enter 确认。</li>
        <li><b>单位与小数位</b>：单位列是下拉框（可滚动选择，也可直接输入自定义单位），内置 m、m2、m3、kg、t、个、块、套、处、台、组、根、樘、座、件、项、台班、延长米、千块等。工程量小数位按单位自动决定：m / m2 / m3 / kg 保留 2 位，t 保留 3 位，个 / 块 / 套 / 处 / 台 / 组 取整，其他单位默认 3 位；表格、打印、Excel 导出口径一致。</li>
        <li><b>粘贴</b>：从 WPS / Excel 复制多行多列，选中表格起始格按 Ctrl+V，起点取选中区左上角；数据行数超过现有行数时会自动在末尾补足空行。</li>
        <li><b>清除内容 / 删除</b>：选中单元格（可多选）后按 <b>Delete</b> 或右键「清除单元格内容」，只清空内容、保留行列结构，可撤销；整行 / 整列选中时按 Delete 才删除整行 / 整列，结构删除也保留在专门的右键菜单里。</li>
        <li><b>右键菜单</b>：单元格右键可插入行、清除内容、设置 / 清除打印区域、调整列宽行高、标记为不计等；行号右键可插入 / 删除行、整行上色、行高自适应；表头右键可添加 / 删除 / 重命名列、整列着色、列宽自适应；左上角角标右键可删除当前单项工程全部数据 / 删除全部数据、自动行列宽高开关、全选。</li>
        <li><b>填充颜色</b>：选中单元格或整行 / 整列后，点「格式」组「填充颜色」，选颜色和应用范围（整行 / 整列 / 仅选中单元格）。优先级为 单元格 &gt; 整行 &gt; 整列；「清除填充」按当前选择自动判断范围。颜色会带到打印、预览和 Excel。</li>
        <li><b>列宽 / 行高</b>：<b>默认自适应</b>——列宽按内容计算并铺满窗口，行高按内容撑高（长公式 / 长备注自动换行），无需手点。手动拖过列宽后自动列宽会关闭，可在左上角角标右键重新勾选「自动列宽 / 自动行高」恢复；把鼠标移到两列（或两行）表头之间的分隔线上<b>双击</b>可只自适应该列 / 该行。</li>
        <li><b>撤销 / 恢复、查找</b>：Ctrl+Z 撤销、Ctrl+Y 恢复，逐字输入公式只算一步；Ctrl+F 打开查找，可按编号 / 项目名称 / 单位 / 部位 / 计算公式 / 备注或「全部字段」搜索，回车逐个定位到对应行。</li>
        </ul>

        <h3 style="color:#2f5597;">五、公式与函数</h3>
        <ul>
        <li><b>基本运算</b>：+ - * / 与 × ÷ 均可，幂用 ^，如 <code>2^3</code>、<code>4×9.6×0.1</code>；支持上标 ² ³ 直接当乘方，如 <code>3³=27</code>。</li>
        <li><b>注释</b>：公式中的汉字、<code>[方括号]</code>、<code>{花括号}</code>内容自动忽略，如 <code>4×9.6[柱基预埋]</code> 只计算 4×9.6；方括号常用于写备注。</li>
        <li><b>常量</b>：@加名称，如 <code>@层高=3.6</code>。左下常量框为「常量代码 / 值 / 说明」三列，双击即可修改，不写 @ 会自动补上。常量全部自定义；需要钢筋单重时点「▾」→「插入常用钢筋单重（14 条模板）」，插入后可随意改删。公式里写 <code>@层高×4</code>、<code>@Φ8×12</code> 即可参与计算；引用未定义的常量，该行红色提示“未定义的常量：@xxx”。</li>
        <li><b>函数</b>：ceil 向上取整、floor 向下取整、sqr 平方、sqrt 开平方、round 四舍五入、power(x,n) 幂、pi / π 圆周率、radians 角度转弧度、du 弧度转角度、abs 绝对值、max、min、sum、<b>avg 平均值</b>、sin、cos、tan。点「插入函数」弹出函数面板，点函数名即插入 <code>函数名()</code>，光标停在括号内，每个函数都有写法提示。</li>
        <li><b>括号快捷键</b>：F4 公式前插入 <code>()</code>、F5 插入点插入 <code>()</code>、F6 插入点插入 <code>[]</code>、F7 末尾追加 <code>[]</code>；在单元格里直接打字时也生效。</li>
        <li><b>特殊写法</b>：度分秒如 34°56′30″ 可直接计算；<code>√9=3</code>、<code>√(3+4)</code>、<code>2×√9</code> 均可；钢筋等级一级 Φ、二级 Φ²、三级 Φ³、四级 Φ⁴（按行业习惯在 Φ 右上角标级别），写在公式开头作规格说明时不参与计算，如 <code>Φ³8@200 4×5</code>=20。点「特殊符号」可看图选符号。</li>
        <li><b>汇总与配色</b>：每个分部自动「小计」（XJ 子分部汇总函数已取消，旧工程中的 XJ() 公式仍可正常计算）。公式列与辅助框用同一套配色：数字红、运算符黑、<code>[备注]</code>深绿、汉字淡绿、@常量紫、上标蓝，便于核对。</li>
        </ul>

        <h3 style="color:#2f5597;">六、打印与导出</h3>
        <ul>
        <li><b>打印 / 预览</b>：Ctrl+P 调出系统打印对话框，可选打印机、份数、页码范围，也可选“Microsoft Print to PDF”存成 PDF。「打印预览」按纸张实际比例分页显示，鼠标拖动页面上的蓝色虚线即可调整页边距，上方数值框同步更新；预览与实际打印的分页完全一致，<b>每页都会重复报表标题和表头行</b>。</li>
        <li><b>页面设置</b>：可设表头标题、页眉、<b>页脚左 / 中 / 右三栏</b>（一般放编制人、复核人、日期等签字信息）、上 / 下 / 左 / 右页边距、纸张方向、纸张大小（下拉列出系统打印机支持的纸型，选“自定义”可填宽高 mm）、水平 / 垂直对齐。打印和预览每页底部自动显示「第 X 页，共 Y 页」，无需手动填页码。</li>
        <li><b>打印区域</b>：在表格里框选单元格后<b>右键 →「设置打印区域」</b>，报表只输出选中的行列，小计也只算选中行；右键「清除打印区域」恢复打印整张计算书。快捷操作行的「打印选中区域」也可直接使用。</li>
        <li><b>导出</b>：「导出Excel表格 / 导出WPS表格」生成 .xlsx，包含各计算书与常量表；页面设置中的页脚三栏和自动页码也会写入导出文件。</li>
        </ul>

        <h3 style="color:#2f5597;">七、分类汇总</h3>
        <ul>
        <li>右侧面板可按 <b>编号 / 项目名称 / 部位 / 分部 / 单位</b> 把计算数据分类合计，同名同单位的多行自动累加；勾「仅当前分部（含下级）」只统计左侧选中分部，点「刷新」重算，「导出」存 Excel，「复制」可直接贴到 WPS。</li>
        <li>点「多条件…」可用多字段分组（如「部位 + 单位」组合）和多条筛选（字段 + 包含 / 不包含 / 等于 / 不等于 / 为空 / 不为空 + 关键词），条件之间可选满足全部（AND）或满足任一（OR）；点「单条件」切回。</li>
        <li>汇总表单元格可多选并按 Ctrl+C 复制；所有列宽均可直接拖动（含项目名称列）并被记住，表头 / 角标 / 单元格右键可设列宽、行高、自适应、复制整表、导出；标题行「⤢ / ⤡」可放大汇总或还原。</li>
        </ul>

        <h3 style="color:#2f5597;">八、其他功能与快捷键</h3>
        <ul>
        <li><b>自动保存与备份</b>：「计算书」组「选项配置」里可设自动保存开关、间隔（1–120 分钟）和备份目录；到点时若工程有改动会写回主文件，并在备份目录生成带时间戳的副本。</li>
        <li><b>工程文件</b>：保存为 <code>.gcls</code>（内容仍是 JSON 文本，便于排查），包含所有计算书、分部结构、公式、备注、常量、列定义、填充色与页面设置；旧版 .json 工程也能打开。软件启动会自动打开上次使用的工程。</li>
        <li><b>常用快捷键</b>：Ctrl+N 新建、Ctrl+O 打开、Ctrl+S 保存、Ctrl+P 打印、Ctrl+F 查找、Ctrl+C / X / V 复制 / 剪切 / 粘贴、Delete 清除单元格内容、Insert 插入行、Ctrl+Z / Ctrl+Y 撤销 / 恢复、F4–F7 插入括号、F9 重算、Ctrl+Enter 确认公式。</li>
        <li><b>行 / 列高亮</b>：表格上方快捷操作行的「选中行高亮」「选中列高亮」两个开关——当前行、当前列加同色淡蓝底纹，可同时开（交叉处稍深）；只加底纹，不影响框选、删除与粘贴。</li>
        <li><b>表格字体</b>：功能区「计算书编辑 → 格式 → 表格字体」，字体类型来自电脑系统字库（宋体、黑体、仿宋、楷体等），字号 6–24；确定后表格行高 / 列宽按新字体自动重排，设置会记住，下次启动自动恢复。</li>
        </ul>
        """

    # ---------------- 树 / 表 ----------------
    def refresh_tree(self):
        self._loading = True
        self.tree.clear()
        proj = self.project

        def add(sec, parent_item, is_book=False):
            it = QTreeWidgetItem([sec.name, proj.level_of(sec)])
            it.setData(0, Qt.UserRole, sec.id)
            it.setData(0, Qt.UserRole + 1, 'book' if is_book else 'section')
            it.setExpanded(sec.expanded)
            it.setToolTip(0, sec.name)
            it.setToolTip(1, proj.level_of(sec))
            # 层级列灰一点，不抢名称的注意力
            it.setForeground(1, QColor('#7a7a7a'))
            if is_book:
                # 导航栏统一用常规字体（不加粗），单项工程用 ▣ 前缀区分
                it.setText(0, '▣ ' + sec.name)
            if parent_item:
                parent_item.addChild(it)
            else:
                self.tree.addTopLevelItem(it)
            for c in sec.children:
                add(c, it)
            return it

        # 顶级显示所有单项工程
        first_item = None
        for book in proj.books:
            bi = add(book, None, is_book=True)
            if first_item is None:
                first_item = bi
        # 恢复当前选中项（按 id 找回）
        target = self._find_tree_item(self.current_section.id) if getattr(self, 'current_section', None) else None
        if target is None:
            target = first_item
            if proj.books:
                self.current_section = proj.books[0]
        if target is not None:
            p = target.parent()
            while p is not None:
                p.setExpanded(True)
                p = p.parent()
            self.tree.setCurrentItem(target)
        self.update_constants_display()
        self._loading = False

    def update_constants_display(self):
        """常量框：三列（常量代码 | 值 | 说明），双击即可就地修改。"""
        self.const_list.set_constants(self.project.constants or {},
                                      getattr(self.project, 'constant_notes', None))

    def _find_tree_item(self, sid, parent=None):
        """在树中按 id 查找 QTreeWidgetItem。"""
        def walk(item):
            for i in range(item.childCount()):
                ch = item.child(i)
                if ch.data(0, Qt.UserRole) == sid:
                    return ch
                r = walk(ch)
                if r:
                    return r
            return None
        if parent is None:
            for i in range(self.tree.topLevelItemCount()):
                top = self.tree.topLevelItem(i)
                if top.data(0, Qt.UserRole) == sid:
                    return top
                r = walk(top)
                if r:
                    return r
            return None
        return walk(parent)

    def find_section(self, sid, sec=None):
        """跨所有计算书按 id 查找分部。"""
        if sec is not None:
            roots = [sec]
        elif sid:
            roots = self.project.books
        else:
            roots = self.project.books
        for root in roots:
            for s in root.iter_self_and_desc():
                if s.id == sid:
                    return s
        return None

    def on_tree_select(self, current, previous):
        if self._loading or current is None:
            return
        sid = current.data(0, Qt.UserRole)
        sec = self.find_section(sid)
        if sec:
            self.current_section = sec
            self.refresh_table()

    def row_precision(self, row):
        """每行工程量的小数位：由单位决定（m/m2/kg→2、t→3、计件→0）。"""
        return precision_for_unit(getattr(row, 'unit', ''), self.precision)

    def cell_text_for(self, row, key):
        """按列 key 取单元格显示文本。"""
        if key == 'no_count':
            return ''
        if key == 'value':
            return fmt_num(row.value, self.row_precision(row))
        if hasattr(row, key):
            return str(getattr(row, key))
        return str(row.custom.get(key, ''))

    def row_fill_for(self, row_obj, col_def):
        """单元格实际底色：单元格 > 整行 > 整列（空串表示无填充）。"""
        key = col_def['key']
        if row_obj is not None:
            cell = (getattr(row_obj, 'cell_fill', None) or {}).get(key)
            if cell:
                return cell
            if getattr(row_obj, 'fill', ''):
                return row_obj.fill
        return col_def.get('fill', '') or ''

    def refresh_table(self):
        self._loading = True
        sec = self.current_section
        cols = self.project.columns
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels([c['title'] for c in cols])
        for j, cd in enumerate(cols):
            hi = self.table.horizontalHeaderItem(j)
            if hi is not None and cd['key'] == 'unit':
                hi.setToolTip('双击（或选中后再点一下）弹出常用单位下拉框，可直接选择\n'
                              'm/m2/m3/kg 保留2位、t 保留3位、个/块/套/处/台/组 取整')
        # 应用列宽（默认比例）。setSectionsMovable 下按逻辑索引设置宽度
        hh = self.table.horizontalHeader()
        for j, cd in enumerate(cols):
            w = int(cd.get('width', 100))
            # 逻辑列 j 对应的视觉段：用 setColumnWidth 作用于逻辑列
            self.table.setColumnWidth(j, w)
        self.table.setRowCount(max(len(sec.rows), 1))
        for i in range(self.table.rowCount()):
            for j, cd in enumerate(cols):
                key = cd['key']
                if i < len(sec.rows):
                    row = sec.rows[i]
                    val = self.cell_text_for(row, key)
                    if key == 'no_count':
                        it = QTableWidgetItem()
                        it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable)
                        it.setCheckState(Qt.Checked if row.no_count else Qt.Unchecked)
                        it.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                    else:
                        it = QTableWidgetItem(val)
                        if key == 'value' and row.value is not None:
                            it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                        if row.error and key in ('formula', 'value'):
                            it.setForeground(QColor('C00000'))
                            it.setToolTip(row.error)
                        self._apply_item_font(it, row, key)
                        self._apply_item_align(it, row, key)
                    bg = self.row_fill_for(row, cd)
                    if bg:
                        it.setBackground(QBrush(QColor(bg)))
                else:
                    it = QTableWidgetItem()
                    if key == 'no_count':
                        it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable)
                        it.setCheckState(Qt.Unchecked)
                    if cd.get('fill'):
                        it.setBackground(QBrush(QColor(cd['fill'])))
                self.table.setItem(i, j, it)
            self.table.setRowHeight(i, 34)
        self._apply_row_height_mode()
        # 列宽：默认按内容自适应并铺满窗口（用户手动拖过列宽后自动关闭）
        self._auto_size_columns()
        # 记住界面上当前各行的真实高度，打印时按同一行高输出
        heights = {}
        for i, row in enumerate(sec.rows):
            if i < self.table.rowCount():
                heights[id(row)] = self.table.rowHeight(i)
        self._row_height_cache = heights
        vals = list(heights.values())
        self._row_height_default = (max(set(vals), key=vals.count) if vals else 34)
        self._loading = False
        self._apply_filter()          # 筛选生效时保持隐藏状态
        self.lbl_status.setText('%s：%d 行，小计 %s'
                                % (sec.name, len(sec.rows),
                                   fmt_num(sec.total(), section_precision(sec, self.precision))))
        self.schedule_summary_refresh()

    def col_key(self, col):
        """逻辑列索引 -> 列 key。

        item.column() / currentColumn() / QModelIndex.column() 返回的都是模型逻辑列；
        表头拖动只改变视觉顺序，不改变 item 的逻辑列，故此处直接按逻辑列取值。
        """
        cols = self.project.columns
        if 0 <= col < len(cols):
            return cols[col]['key']
        return ''

    def cell_value(self, row, col):
        item = self.table.item(row, col)
        return item.text() if item else ''

    def set_cell_value(self, row, col, value):
        item = self.table.item(row, col)
        if item is None:
            item = QTableWidgetItem('')
            self.table.setItem(row, col, item)
        item.setText(str(value))

    def on_cell_changed(self, item):
        if self._loading or self._syncing:
            return
        row_i, col = item.row(), item.column()
        if row_i >= len(self.current_section.rows):
            # 在空行输入内容 → 自动追加行
            entered = item.text()
            if entered.strip():
                self.push_undo()
                self.current_section.rows.append(Row())
                self.refresh_table()
                # 追加后把已输入的值写回新行（用保存的文本，item 已被重建销毁）
                if row_i < len(self.current_section.rows):
                    self.set_row_value(row_i, col, entered)
            return
        row = self.current_section.rows[row_i]
        key = self.col_key(col)
        text = item.text()
        if key == 'value':
            return
        if key == 'formula':
            # 公式是边打字边写进模型的：改动前快照在 on_cell_editor_text 里已记下
            self.commit_edit()
        else:
            self.begin_edit('编辑单元格')
        if key == 'no_count':
            row.no_count = (item.checkState() == Qt.Checked)
        elif hasattr(row, key):
            setattr(row, key, text)
        else:
            row.custom[key] = text
        if key != 'formula':
            self.commit_edit()
        # 表格里改公式时，同步辅助框（着色显示）
        if key == 'formula':
            self._formula_row = row_i
            if self.formula_help.isReadOnly():
                self.formula_help.setReadOnly(False)
                self.formula_help.setStyleSheet('background:#ffffff;')
            self.formula_help.setTextSilent(text)
        # 表格里改备注时直接写在“备注”列，不再同步右侧面板
        if key in ('formula', 'no_count'):
            self.project.recalc()
            # 只刷新工程量列与小计，不整表重建，避免打断正在进行的单元格输入
            self._update_value_cells()
            self._dirty = True
        elif key == 'unit':
            # 单位变了 → 该行工程量的小数位跟着变，只刷显示不重建表格
            self._update_value_cells()
            self._dirty = True

    def set_row_value(self, row_i, col, text):
        """把值写入指定行（供空行追加时用）。"""
        if row_i >= len(self.current_section.rows):
            return
        row = self.current_section.rows[row_i]
        key = self.col_key(col)
        if key == 'no_count':
            row.no_count = (text == '是')
        elif key and hasattr(row, key) and key != 'value':
            setattr(row, key, text)
        elif key:
            row.custom[key] = text
        if key == 'formula':
            self._formula_row = row_i
            self.formula_help.setTextSilent(text)
        self.project.recalc()
        self.refresh_table()

    def on_cell_cursor(self, row, col, prev_row, prev_col):
        """选中单元格：公式列联动辅助框，双向同步；行 / 列高亮随光标重绘。"""
        if self._loading:
            return
        # 行 / 列高亮需要跟着当前单元格重绘（委托在 paint 时读取 currentIndex）
        if getattr(self, '_hilite_row', False) or getattr(self, '_hilite_col', False):
            self.table.viewport().update()
        is_formula_col = (self.col_key(col) == 'formula')
        if is_formula_col and 0 <= row < len(self.current_section.rows):
            text = self.current_section.rows[row].formula
            self._formula_row = row
            self.formula_help.setReadOnly(False)
            self.formula_help.setStyleSheet('background:#ffffff;')
            self.formula_help.setTextSilent(text)
        else:
            self.formula_help.setReadOnly(True)
            self.formula_help.setStyleSheet('background:#f7f7f7; color:#888;')
            self.formula_help.setPlaceholderText('请选中"计算公式"列的单元格，公式将在此同步显示并着色')
            self.formula_help.setTextSilent('')
            self._formula_row = None

    def on_note_changed(self):
        """（已取消右侧备注面板，保留空实现兼容旧调用。）"""
        return

    def _on_editor_destroyed(self, *args):
        self._cell_editor = None
        # 单元格编辑结束：把本次连续输入记成一步可撤销操作
        self.commit_edit()

    def on_cell_editor_text(self, text):
        """表格公式单元格【正在编辑】→ 逐字实时同步到辅助框、模型与工程量。"""
        if self._loading or self._syncing:
            return
        row = self.table.currentRow()
        col = self.table.currentColumn()
        if col < 0 or self.col_key(col) != 'formula':
            return
        if row < 0 or row >= len(self.current_section.rows):
            return
        self._syncing = True
        try:
            self.begin_edit('改公式')
            self._formula_row = row
            # 辅助框保持可编辑，并显示同样的内容（着色）
            if self.formula_help.isReadOnly():
                self.formula_help.setReadOnly(False)
                self.formula_help.setStyleSheet('background:#ffffff;')
            if self.formula_help.toPlainText() != text:
                self.formula_help.setTextSilent(text)
            self.current_section.rows[row].formula = text
            self.project.recalc()
            self._update_value_cells()
            self._dirty = True
        finally:
            self._syncing = False

    def on_formula_help_text(self, text):
        """辅助框内输入 → 实时写回表格公式单元格（及正在编辑的编辑器）并着色。"""
        if self._loading or self._syncing or self.formula_help.isReadOnly():
            return
        row = self._formula_row
        if row is None or row < 0 or row >= len(self.current_section.rows):
            # 若当前选中的是公式列，用当前行
            cur = self.table.currentRow()
            c = self.table.currentColumn()
            if c >= 0 and self.col_key(c) == 'formula' and 0 <= cur < len(self.current_section.rows):
                row = cur
                self._formula_row = row
            else:
                return
        self._syncing = True
        try:
            self.begin_edit('改公式')
            self.current_section.rows[row].formula = text
            # 同步表格单元格文本（静默，避免触发 itemChanged 递归）
            for c in range(self.table.columnCount()):
                if self.col_key(c) == 'formula':
                    it = self.table.item(row, c)
                    if it is not None and it.text() != text:
                        self._loading = True
                        it.setText(text)
                        self._loading = False
                    break
            # 若该单元格正处于编辑状态，编辑器内容也要同步，
            # 否则结束编辑时旧内容会覆盖辅助框，破坏双向同步。
            ed = self._cell_editor
            if ed is not None:
                try:
                    if ed.text() != text:
                        ed.blockSignals(True)
                        ed.setText(text)
                        ed.blockSignals(False)
                except RuntimeError:
                    self._cell_editor = None
            self.project.recalc()
            # 只更新工程量列显示，避免整表刷新打断输入
            self._update_value_cells()
            self._dirty = True
        finally:
            self._syncing = False

    def _update_value_cells(self):
        """只刷新工程量列与小计状态，不重建整个表格。"""
        for r in range(len(self.current_section.rows)):
            for c in range(self.table.columnCount()):
                if self.col_key(c) == 'value':
                    it = self.table.item(r, c)
                    if it is not None:
                        row_obj = self.current_section.rows[r]
                        self._loading = True
                        it.setText(fmt_num(row_obj.value, self.row_precision(row_obj)))
                        it.setForeground(QColor('C00000') if row_obj.error else QColor('000000'))
                        self._loading = False
                    break
        sec = self.current_section
        self.lbl_status.setText('%s：%d 行，小计 %s'
                                % (sec.name, len(sec.rows),
                                   fmt_num(sec.total(), section_precision(sec, self.precision))))
        self.schedule_summary_refresh()

    def apply_formula_help(self):
        """把辅助框内容写入当前行的公式列（Ctrl+Enter / 写入按钮）。"""
        if self.formula_help.isReadOnly():
            return
        row = self._formula_row
        if row is None or row < 0 or row >= len(self.current_section.rows):
            return
        text = self.formula_help.toPlainText()
        self.current_section.rows[row].formula = text
        for c in range(self.table.columnCount()):
            if self.col_key(c) == 'formula':
                it = self.table.item(row, c)
                if it is not None and it.text() != text:
                    self._loading = True
                    it.setText(text)
                    self._loading = False
                break
        self.project.recalc()
        self._update_value_cells()
        self._dirty = True
        self.commit_edit()

    # ---------------- 粘贴 ----------------
    def paste_from_clipboard(self):
        self.paste_text(QApplication.clipboard().text())

    def paste_text(self, text):
        if not text:
            return
        lines = text.replace('\r\n', '\n').replace('\r', '\n').rstrip('\n').split('\n')
        if not lines:
            return
        rows = [ln.split('\t') for ln in lines]
        # 粘贴锚点：优先取选中单元格区域的左上角（用户框选备注列再粘贴时，
        # currentColumn() 可能仍停在编码列或 -1，导致贴错列）。
        sel = self.table.selectionModel()
        sel_indexes = sel.selectedIndexes() if sel is not None else []
        if sel_indexes:
            start_row = min(i.row() for i in sel_indexes)
            start_col = min(i.column() for i in sel_indexes)
        else:
            start_row = max(self.table.currentRow(), 0)
            start_col = max(self.table.currentColumn(), 0)
        sec = self.current_section
        need = start_row + len(rows)
        before_n = len(sec.rows)
        self.push_undo('粘贴')
        # 粘贴行数超过现有行数时，自动在末尾补足空行，数据才能完整落位
        while len(sec.rows) < need:
            sec.rows.append(Row())
        added = len(sec.rows) - before_n
        cols = self.project.columns
        self._loading = True
        for i, cells in enumerate(rows):
            r = start_row + i
            for j, val in enumerate(cells):
                c = start_col + j
                if c < len(cols):
                    row = sec.rows[r]
                    key = cols[c]['key']
                    if key == 'value':
                        continue
                    if key == 'no_count':
                        row.no_count = val.strip() == '是'
                    elif hasattr(row, key):
                        setattr(row, key, val.rstrip())
                    else:
                        row.custom[key] = val.rstrip()
        self._loading = False
        self.project.recalc()
        self.refresh_table()
        msg = f'已粘贴 {len(rows)} 行 × {max(len(x) for x in rows)} 列'
        if added:
            msg += f'，已自动新增 {added} 行'
        self.status.showMessage(msg, 4000)

    # ---------------- 导入表格文件（xlsx / csv / txt） ----------------
    def import_table(self):
        from ui_import import ImportTableDialog, read_table_file, trim_data
        path, _ = QFileDialog.getOpenFileName(
            self, '导入表格文件', '',
            '表格文件 (*.xlsx *.csv *.txt);;Excel 工作簿 (*.xlsx);;'
            'CSV 文件 (*.csv);;文本文件 (*.txt);;所有文件 (*)')
        if not path:
            return
        sec = self.current_section
        dlg = ImportTableDialog(self, self.project.columns,
                                default_row=len(sec.rows) + 1, path=path)
        if dlg.exec_() != QDialog.Accepted:
            return
        pairs = dlg.map_pairs()
        src_rows = dlg.data_from()
        tgt0 = dlg.target_row0()          # None=追加到末尾
        if tgt0 is None:
            tgt0 = len(sec.rows)
        need = tgt0 + len(src_rows)
        before_n = len(sec.rows)
        self.push_undo('导入表格')
        # 行数不够时自动在末尾补足空行
        while len(sec.rows) < need:
            sec.rows.append(Row())
        added = len(sec.rows) - before_n
        cols = self.project.columns
        self._loading = True
        n_cells = 0
        for i, srow in enumerate(src_rows):
            row = sec.rows[tgt0 + i]
            for si, key in pairs:
                val = srow[si].strip() if si < len(srow) else ''
                if key == 'no_count':
                    row.no_count = val in ('是', '√', 'Y', 'y', '1', 'true', 'True')
                    n_cells += 1
                    continue
                if not any(c['key'] == key for c in cols):
                    continue
                if hasattr(row, key) and key != 'value':
                    setattr(row, key, val)
                else:
                    row.custom[key] = val
                n_cells += 1
        self._loading = False
        self.project.recalc()
        self.refresh_table()
        self._dirty = True
        msg = f'已导入 {len(src_rows)} 行 × {len(pairs)} 列（共 {n_cells} 个单元格）'
        if added:
            msg += f'，已自动新增 {added} 行'
        if tgt0:
            msg += f'，写入位置：第 {tgt0 + 1} 行起'
        self.status.showMessage(msg, 6000)

    # ---------------- 树节点操作（单项工程 / 单位工程 / 分部 / 分项 …） ----------------
    def _tree_section(self):
        """当前操作对象：优先取导航树里选中的节点，避免界面选中与内部状态不一致。"""
        it = self.tree.currentItem()
        if it is not None:
            sec = self.find_section(it.data(0, Qt.UserRole))
            if sec is not None:
                self.current_section = sec
                return sec
        return self.current_section

    def add_section(self):
        """新建同级节点（与当前节点同一层级）。"""
        sec = self._tree_section()
        if sec is None:
            return
        if self.project.is_book(sec):
            return self.add_book()
        parent = sec.parent
        if parent is None:
            return
        name = self.project.new_child_name(parent)
        self.push_undo('新建' + self.project.next_level_of(parent))
        new = Section(name, parent)
        try:
            idx = parent.children.index(sec) + 1
        except ValueError:
            idx = len(parent.children)
        parent.children.insert(idx, new)
        self.refresh_tree()
        self.select_section(new)
        self._dirty = True
        self.status.showMessage('已新建%s「%s」' % (self.project.level_of(new), new.name), 3000)

    def add_child_section(self):
        """新建下一级节点（当前节点的子节点）。"""
        sec = self._tree_section()
        if sec is None:
            return
        name = self.project.new_child_name(sec)
        label = self.project.next_level_of(sec)
        self.push_undo('新建' + label)
        new = sec.add_child(name)
        sec.expanded = True
        self.refresh_tree()
        self.select_section(new)
        self._dirty = True
        self.status.showMessage('已新建%s「%s」' % (label, new.name), 3000)

    def select_section(self, sec):
        if sec is None:
            return
        it = self._find_tree_item(sec.id)
        if it is not None:
            p = it.parent()
            while p is not None:
                p.setExpanded(True)
                p = p.parent()
            self.tree.setCurrentItem(it)
            self.tree.scrollToItem(it)
        self.current_section = sec
        self.refresh_table()

    def _on_tree_double(self):
        self.rename_section()

    def rename_section(self):
        sec = self._tree_section()
        if sec is None:
            return
        if self.project.is_book(sec):
            return self.rename_book()
        label = self.project.level_of(sec)
        name, ok = QInputDialog.getText(self, '重命名' + label, label + '名称：',
                                        QLineEdit.Normal, sec.name)
        if ok and name.strip():
            self.push_undo('重命名')
            sec.name = name.strip()
            self.refresh_tree()
            self.select_section(sec)
            self._dirty = True

    def delete_section(self):
        sec = self._tree_section()
        if sec is None:
            return
        if self.project.is_book(sec):
            return self.delete_book()
        label = self.project.level_of(sec)
        cnt_sec = len(list(sec.iter_self_and_desc())) - 1
        cnt_row = sum(len(s.rows) for s in sec.iter_self_and_desc())
        msg = '确定删除%s「%s」？' % (label, sec.name)
        if cnt_sec or cnt_row:
            msg += '\n（包含 %d 个下级节点、%d 行数据）' % (cnt_sec, cnt_row)
        msg += '\n\n只删除该节点及其下级，不影响同级或上级。'
        if QMessageBox.question(self, '确认删除', msg,
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.push_undo('删除' + label)
            parent = sec.parent
            if parent:
                parent.children.remove(sec)
            self.current_section = parent if parent else self.project.root
            self.refresh_tree()
            self.select_section(self.current_section)
            self._dirty = True

    def move_section_up(self):
        self._move_section(-1)

    def move_section_down(self):
        self._move_section(1)

    def _move_section(self, delta):
        """上移 / 下移：同一层内换位置（单项工程也支持调整顺序）。"""
        sec = self._tree_section()
        if sec is None:
            return
        if self.project.is_book(sec):
            lst = self.project.books
        elif sec.parent is not None:
            lst = sec.parent.children
        else:
            lst = None
        if not lst or sec not in lst:
            self.status.showMessage('当前节点无法移动', 2500)
            return
        i = lst.index(sec)
        j = i + delta
        if not (0 <= j < len(lst)):
            self.status.showMessage('已到%s，不能再%s'
                                    % ('最前' if delta < 0 else '最后',
                                       '上移' if delta < 0 else '下移'), 3000)
            return
        self.push_undo('上移' if delta < 0 else '下移')
        lst[i], lst[j] = lst[j], lst[i]
        self.refresh_tree()
        self.select_section(sec)
        self._dirty = True
        self.status.showMessage('已%s「%s」' % ('上移' if delta < 0 else '下移', sec.name), 2500)

    def promote_section(self):
        """升级：往上提一层，成为原父节点的兄弟。"""
        sec = self._tree_section()
        if sec is None:
            return
        if self.project.is_book(sec):
            self.status.showMessage('「%s」已是最顶层（%s），不能再升级'
                                    % (sec.name, self.project.level_of(sec)), 3000)
            return
        old_parent = sec.parent
        if old_parent is None or self.project.is_book(old_parent):
            self.status.showMessage('「%s」已是一级%s，不能再升级'
                                    % (sec.name, self.project.level_of(sec)), 3000)
            return
        grand = old_parent.parent
        if grand is None:
            return
        self.push_undo('升级')
        old_parent.children.remove(sec)
        sec.parent = grand
        idx = grand.children.index(old_parent) + 1
        grand.children.insert(idx, sec)
        self.refresh_tree()
        self.select_section(sec)
        self._dirty = True
        self.status.showMessage('已升级：现在是%s「%s」'
                                % (self.project.level_of(sec), sec.name), 3000)

    def demote_section(self):
        """降级：挂到同级前一个节点下面（下移一层）。"""
        sec = self._tree_section()
        if sec is None:
            return
        if self.project.is_book(sec):
            self.status.showMessage('「%s」是最顶层（%s），不能降级'
                                    % (sec.name, self.project.level_of(sec)), 3000)
            return
        siblings = sec.parent.children if sec.parent else []
        if sec not in siblings:
            return
        i = siblings.index(sec)
        if i == 0:
            self.status.showMessage('前面没有同级节点可以挂靠，无法降级', 3000)
            return
        prev = siblings[i - 1]
        self.push_undo('降级')
        siblings.remove(sec)
        sec.parent = prev
        prev.children.append(sec)
        prev.expanded = True
        self.refresh_tree()
        self.select_section(sec)
        self._dirty = True
        self.status.showMessage('已降级：现在是%s「%s」的下级'
                                % (self.project.level_of(prev), sec.name), 3000)

    def level_setup(self):
        """层级设置：自定义 单项工程 / 单位工程 / 分部工程 / 分项工程 的名字与层数。"""
        dlg = QDialog(self)
        dlg.setWindowTitle('层级设置')
        dlg.resize(420, 360)
        v = QVBoxLayout(dlg)
        tip = QLabel('按从上到下的顺序填写每一级的名称（第一级就是导航栏最顶层）。\n'
                     '可以增删层级、改名，例如改成 标段 / 单位工程 / 分部 / 分项。\n'
                     '新建节点会自动使用下一级的名称。')
        tip.setWordWrap(True)
        tip.setStyleSheet('color:#555;')
        v.addWidget(tip)
        lst = QListWidget()
        lst.setDragDropMode(QAbstractItemView.InternalMove)
        for name in (self.project.levels or DEFAULT_LEVELS):
            it = QListWidgetItem(name)
            it.setFlags(it.flags() | Qt.ItemIsEditable)
            lst.addItem(it)
        v.addWidget(lst, 1)
        bar = QHBoxLayout()
        b_add = QPushButton('添加一级')
        b_del = QPushButton('删除选中')
        b_up = QPushButton('上移')
        b_down = QPushButton('下移')
        b_def = QPushButton('恢复默认')
        for b in (b_add, b_del, b_up, b_down, b_def):
            bar.addWidget(b)
        bar.addStretch(1)
        v.addLayout(bar)

        def add_level():
            it = QListWidgetItem('新层级%d' % (lst.count() + 1))
            it.setFlags(it.flags() | Qt.ItemIsEditable)
            lst.addItem(it)
            lst.setCurrentItem(it)
            lst.editItem(it)
        def del_level():
            r = lst.currentRow()
            if r >= 0 and lst.count() > 1:
                lst.takeItem(r)
        def move(delta):
            r = lst.currentRow()
            j = r + delta
            if r >= 0 and 0 <= j < lst.count():
                it = lst.takeItem(r)
                lst.insertItem(j, it)
                lst.setCurrentRow(j)
        b_add.clicked.connect(add_level)
        b_del.clicked.connect(del_level)
        b_up.clicked.connect(lambda: move(-1))
        b_down.clicked.connect(lambda: move(1))
        b_def.clicked.connect(lambda: (lst.clear(), [lst.addItem(QListWidgetItem(x))
                                                     for x in DEFAULT_LEVELS]))
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        v.addWidget(bb)
        if dlg.exec_() != QDialog.Accepted:
            return
        names = [(lst.item(i).text() or '').strip() for i in range(lst.count())]
        names = [n for n in names if n]
        if not names:
            QMessageBox.information(self, APP_NAME, '至少要保留一级名称。')
            return
        self.push_undo('层级设置')
        self.project.levels = names
        self.refresh_tree()
        self._dirty = True
        self.status.showMessage('层级已更新：' + ' > '.join(names), 4000)

    # ---------------- 单项工程（多文档）操作 ----------------
    def add_book(self):
        top = self.project.level_name(0)
        name, ok = QInputDialog.getText(self, '新建' + top, top + '名称：',
                                        QLineEdit.Normal, self.project.new_book_name())
        if not ok or not name.strip():
            return
        self.push_undo('新建' + top)
        book = self.project.add_book(name.strip())
        self.project.recalc()
        self.refresh_tree()
        self.select_section(book)
        self._dirty = True
        self.status.showMessage('已新建%s「%s」' % (top, book.name), 3000)

    def rename_book(self):
        book = self._tree_section()
        if not self.project.is_book(book):
            return
        top = self.project.level_name(0)
        name, ok = QInputDialog.getText(self, '重命名' + top, top + '名称：',
                                        QLineEdit.Normal, book.name)
        if ok and name.strip():
            self.push_undo('重命名')
            book.name = name.strip()
            self.refresh_tree()
            self.select_section(book)
            self._dirty = True

    def delete_book(self):
        book = self._tree_section()
        if not self.project.is_book(book):
            return
        top = self.project.level_name(0)
        if self.project.is_first_book(book):
            QMessageBox.information(self, APP_NAME,
                                    '「%s」是工程的主%s，不可删除。\n'
                                    '如需清空内容，可右键工作区左上角“删除全部数据”，'
                                    '或在上面新建别的%s。' % (book.name, top, top))
            return
        cnt_sec = len(list(book.iter_self_and_desc())) - 1
        cnt_row = sum(len(s.rows) for s in book.iter_self_and_desc())
        msg = '确定删除%s「%s」？' % (top, book.name)
        if cnt_sec or cnt_row:
            msg += '\n（包含 %d 个下级节点、%d 行数据）' % (cnt_sec, cnt_row)
        if QMessageBox.question(self, '确认删除', msg,
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.push_undo('删除' + top)
            self.project.remove_book(book)
            self.current_section = self.project.books[0]
            self.project.recalc()
            self.refresh_tree()
            self.refresh_table()
            self._dirty = True

    def tree_menu(self, pos):
        # 右键时先选中鼠标指向的项，确保操作针对该项本身
        item = self.tree.itemAt(pos)
        sec = None
        if item:
            self.tree.setCurrentItem(item)
            sid = item.data(0, Qt.UserRole)
            sec = self.find_section(sid)
            if sec:
                self.current_section = sec
        if sec is None:
            sec = self._tree_section()
        proj = self.project
        top = proj.level_name(0)
        is_book_target = proj.is_book(sec)
        level = proj.level_of(sec)
        next_level = proj.next_level_of(sec)
        menu = QMenu(self)
        menu.addAction('新建%s' % top, self.add_book)
        menu.addSeparator()
        if is_book_target:
            menu.addAction('重命名 F2', self.rename_section)
            act = menu.addAction('删除%s' % top, self.delete_book)
            act.setEnabled(not proj.is_first_book(sec))
            if proj.is_first_book(sec):
                act.setToolTip('第一个%s不可删除' % top)
            menu.addSeparator()
            bi = proj.books.index(sec) if sec in proj.books else -1
            a_up = menu.addAction('上移（调整%s顺序）' % top, self.move_section_up)
            a_up.setEnabled(bi > 0)
            a_dn = menu.addAction('下移（调整%s顺序）' % top, self.move_section_down)
            a_dn.setEnabled(0 <= bi < len(proj.books) - 1)
            menu.addSeparator()
            menu.addAction('新建下一级（%s）' % next_level, self.add_child_section)
        else:
            siblings = sec.parent.children if sec.parent is not None else []
            i = siblings.index(sec) if sec in siblings else -1
            menu.addAction('新建下一级（%s）' % next_level, self.add_child_section)
            menu.addAction('新建同级（%s）' % level, self.add_section)
            menu.addAction('重命名 F2', self.rename_section)
            menu.addSeparator()
            a_up = menu.addAction('上移', self.move_section_up)
            a_up.setEnabled(i > 0)
            a_dn = menu.addAction('下移', self.move_section_down)
            a_dn.setEnabled(0 <= i < len(siblings) - 1)
            a_pro = menu.addAction('升级（提上一层）', self.promote_section)
            a_pro.setEnabled(sec.parent is not None and not proj.is_book(sec.parent))
            a_dem = menu.addAction('降级（挂到上一个下面）', self.demote_section)
            a_dem.setEnabled(i > 0)
            menu.addSeparator()
            menu.addAction('删除%s' % level, self.delete_section)
        # 工具栏不再放“工程信息”按钮，改到树右键里，需要改时随时能改
        menu.addSeparator()
        menu.addAction('层级设置（自定义各层名称）…', self.level_setup)
        menu.addAction('工程信息设置…', self.edit_info)
        menu.exec_(self.tree.mapToGlobal(pos))

    # ---------------- 行操作 ----------------
    def insert_row(self, *args):
        """插入空行。工具栏/菜单传入的 bool 信号参数会被忽略；
        若显式传入整数行号，则在该行位置插入。"""
        row = None
        for a in args:
            if isinstance(a, int) and not isinstance(a, bool):
                row = a
                break
        if row is None:
            row = self.table.currentRow()
        if row < 0 or row > len(self.current_section.rows):
            row = len(self.current_section.rows)
        self.push_undo('插入行')
        self.current_section.rows.insert(row, Row())
        self.project.recalc()
        self.refresh_table()
        self._dirty = True
        self.table.setCurrentCell(min(row, self.table.rowCount() - 1), 0)

    def insert_row_above(self, *args):
        self.insert_row(max(self.table.currentRow(), 0))

    def insert_row_below(self, *args):
        r = self.table.currentRow()
        self.insert_row(r + 1 if r >= 0 else len(self.current_section.rows))

    def delete_rows(self, *args):
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        if not rows and self.table.currentRow() >= 0:
            rows = [self.table.currentRow()]
        if not rows:
            return
        if QMessageBox.question(self, '确认', f'删除选中的 {len(rows)} 行？',
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.push_undo('删除行')
            for i in rows:
                if i < len(self.current_section.rows):
                    del self.current_section.rows[i]
            self.project.recalc()
            self.refresh_table()
            self._dirty = True

    def delete_cells_data(self, *args):
        """清除选中单元格的数据（保留行列结构），支持单选 / 框选多个单元格。

        这里只清内容，不因为整行 / 整列选中而删除行列；结构删除请用专门的“删除行”菜单。
        """
        sel = self.table.selectionModel()
        indexes = sel.selectedIndexes() if sel is not None else []
        if not indexes:
            r, c = self.table.currentRow(), self.table.currentColumn()
            if r < 0 or c < 0:
                return
            indexes = [self.table.model().index(r, c)]
        cols = self.project.columns
        sec = self.current_section
        touched = 0
        self.push_undo('删除单元格数据')
        self._loading = True
        try:
            for idx in indexes:
                r, c = idx.row(), idx.column()
                if r < 0 or r >= len(sec.rows) or c < 0 or c >= len(cols):
                    continue
                row = sec.rows[r]
                key = cols[c]['key']
                if key == 'value':      # 工程量是计算值，随公式重算，不单独清空
                    continue
                if key == 'no_count':
                    row.no_count = False
                elif hasattr(row, key):
                    setattr(row, key, '')
                else:
                    row.custom.pop(key, None)
                it = self.table.item(r, c)
                if it is not None:
                    if key == 'no_count':
                        it.setCheckState(Qt.Unchecked)
                    else:
                        it.setText('')
                touched += 1
        finally:
            self._loading = False
        if touched:
            self.project.recalc()
            self._update_value_cells()
            self._dirty = True

    def delete_selection(self, *args):
        """Delete 快捷键 / 右键菜单：只清除选中单元格内容，不删除行列。"""
        self.delete_cells_data()

    def table_menu(self, pos):
        row = self.table.rowAt(pos.y())
        col = self.table.columnAt(pos.x())
        # 若右键落在已选中的区域内，保留用户框选（便于批量删除单元格数据）；
        # 否则把当前单元格定位到鼠标位置，保证插入/删除针对鼠标所指。
        hit = self.table.indexAt(pos)
        in_selection = hit.isValid() and hit in self.table.selectionModel().selectedIndexes()
        if not in_selection:
            if row >= 0 and col >= 0:
                hit_idx = self.table.model().index(row, col)
                self.table.selectionModel().select(hit_idx, QItemSelectionModel.ClearAndSelect)
                self.table.setCurrentIndex(hit_idx)
        menu = QMenu(self)
        menu.addAction('复制选中单元格  (Ctrl+C)', self.copy_cells)
        menu.addAction('全选表格', self.table.selectAll)
        menu.addSeparator()
        menu.addAction('在上方插入行', self.insert_row_above)
        menu.addAction('在下方插入行', self.insert_row_below)
        menu.addSeparator()
        menu.addAction('清除单元格内容  (Delete)', self.delete_cells_data)
        menu.addAction('删除选中行', self.delete_rows)
        menu.addAction('粘贴 WPS 数据  (Ctrl+V)', self.paste_from_clipboard)
        menu.addSeparator()
        menu.addAction('格式刷（复制当前单元格格式，再点目标）', self.start_format_painter)
        menu.addAction('排序…', self.sort_dialog)
        menu.addAction('筛选…', self.filter_dialog)
        act_cf = menu.addAction('清除筛选', self.clear_filter)
        act_cf.setEnabled(self._filter is not None)
        menu.addSeparator()
        menu.addAction('设置选中行行高…', self.set_row_height_dialog)
        menu.addAction('设置选中列列宽…', self.set_column_width_dialog)
        menu.addAction('列宽自适应（按内容）', lambda: self.fit_column_widths())
        menu.addAction('行高自适应（按内容）', lambda: self.fit_row_heights())
        menu.addAction('恢复默认行高', self.reset_row_heights)
        menu.addSeparator()
        menu.addAction('填充颜色…', self.fill_color_dialog)
        menu.addAction('清除填充', self.clear_fill)
        menu.addSeparator()
        menu.addAction('单元格字体…', self.cell_font_setup)
        menu.addAction('单元格字体颜色…', self.cell_font_color)
        menu.addAction('清除单元格字体', self.cell_font_clear)
        align_menu = menu.addMenu('单元格对齐')
        align_menu.addAction('靠左对齐', lambda: self.cell_align_set('left'))
        align_menu.addAction('居中对齐', lambda: self.cell_align_set('center'))
        align_menu.addAction('靠右对齐', lambda: self.cell_align_set('right'))
        align_menu.addSeparator()
        align_menu.addAction('恢复默认对齐', lambda: self.cell_align_set('default'))
        menu.addAction('选中单元格转半角（全角→半角）', self.narrow_cells)
        menu.addSeparator()
        menu.addAction('标记为不计', self.toggle_no_count)
        menu.addSeparator()
        menu.addAction('设置打印区域（仅打印选中部分）', self.set_print_area)
        menu.addAction('清除打印区域（恢复打印全部）', self.clear_print_area)
        menu.exec_(self.table.mapToGlobal(pos))

    def row_menu(self, pos):
        """行号（表格最左侧）右键：插入/删除行、整行底色、行高自适应。"""
        vh = self.table.verticalHeader()
        row = vh.logicalIndexAt(pos)
        if row < 0:
            return
        self.table.setCurrentCell(row, 0)
        self.table.selectRow(row)
        menu = QMenu(self)
        menu.addAction('在上方插入行', lambda r=row: self.insert_row(r))
        menu.addAction('在下方插入行', lambda r=row: self.insert_row(r + 1))
        menu.addAction('在末尾追加行', lambda: self.insert_row(self.table.rowCount()))
        menu.addSeparator()
        menu.addAction('删除本行', self.delete_rows)
        color_menu = menu.addMenu('本行底色')
        for name, color in ROW_FILL_PRESETS:
            act = color_menu.addAction(name)
            act.setData(color)
            act.triggered.connect(lambda checked=False, c=color: self.apply_fill(c, 'row'))
        color_menu.addSeparator()
        color_menu.addAction('其他颜色…', lambda: self.fill_color_dialog('row'))
        color_menu.addAction('清除本行底色', lambda: self.apply_fill('', 'row'))
        menu.addSeparator()
        menu.addAction('设置本行行高…', lambda r=row: self.set_row_height_dialog([r]))
        menu.addAction('本行行高自适应', lambda: self.fit_row_heights([row]))
        menu.addAction('全部行高自适应', lambda: self.fit_row_heights())
        menu.addAction('全部列宽自适应', lambda: self.fit_column_widths())
        menu.exec_(vh.mapToGlobal(pos))

    def corner_menu(self, global_pos):
        """表格左上角（全选钮）右键：删除全部数据 + 整表列宽 / 行高自适应。"""
        menu = QMenu(self)
        sec = self.current_section
        menu.addAction('删除当前%s的全部数据（%d 行）' % (self.project.level_of(sec), len(sec.rows)),
                       self.clear_current_rows)
        total = sum(len(s.rows) for s in self.project.all_sections())
        menu.addAction('删除全部数据（所有%s/分部，共 %d 行）'
                       % (self.project.level_name(0), total), self.clear_all_rows)
        menu.addSeparator()
        act_row = menu.addAction('自动行高（随内容撑高）')
        act_row.setCheckable(True)
        act_row.setChecked(bool(self._row_autofit))
        act_row.triggered.connect(self.toggle_row_autofit)
        act_col = menu.addAction('自动列宽（随窗口铺满）')
        act_col.setCheckable(True)
        act_col.setChecked(bool(self._col_autofit))
        act_col.triggered.connect(self.toggle_col_autofit)
        menu.addSeparator()
        menu.addAction('列宽自适应（按内容铺满窗口）', lambda: self.fit_column_widths())
        menu.addAction('行高自适应（按内容）', lambda: self.fit_row_heights())
        menu.addAction('列宽 + 行高 全部自适应', self.fit_all)
        menu.addAction('恢复默认行高', self.reset_row_heights)
        menu.addSeparator()
        menu.addAction('全选表格', self.table.selectAll)
        menu.exec_(global_pos)

    # ---------------- 删除全部数据 ----------------
    def clear_current_rows(self):
        """清空当前节点下的所有计算行（保留节点本身与下级节点）。"""
        sec = self.current_section
        n = len(sec.rows)
        if n == 0:
            QMessageBox.information(self, APP_NAME, '当前%s没有数据行。'
                                    % self.project.level_of(sec))
            return
        if QMessageBox.question(self, '确认删除',
                                '确定删除「%s」下的全部 %d 行计算数据？\n\n'
                                '（只删数据行，保留节点与下级；可用 Ctrl+Z 撤销）'
                                % (sec.name, n),
                                QMessageBox.Yes | QMessageBox.No,
                                QMessageBox.No) != QMessageBox.Yes:
            return
        self.push_undo('删除全部数据')
        sec.rows.clear()
        self.project.recalc()
        self.refresh_table()
        self._dirty = True
        self.status.showMessage('已删除「%s」的 %d 行数据' % (sec.name, n), 3000)

    def clear_all_rows(self):
        """清空整个工程的计算数据（所有单项工程/分部），节点结构保留。"""
        sections = self.project.all_sections()
        total = sum(len(s.rows) for s in sections)
        if total == 0:
            QMessageBox.information(self, APP_NAME, '当前工程没有数据行。')
            return
        if QMessageBox.question(self, '确认删除全部数据',
                                '确定删除整个工程的 %d 行计算数据？\n\n'
                                '（保留所有%s与分部结构，只清空数据行；可用 Ctrl+Z 撤销）'
                                % (total, self.project.level_name(0)),
                                QMessageBox.Yes | QMessageBox.No,
                                QMessageBox.No) != QMessageBox.Yes:
            return
        self.push_undo('删除全部数据')
        for s in sections:
            s.rows.clear()
        self.project.recalc()
        self.refresh_table()
        self._dirty = True
        self.status.showMessage('已删除全部数据（%d 行）' % total, 3000)

    # ---------------- 行高 / 列宽 默认自适应（V2.5） ----------------
    def _on_table_resized(self):
        """表格尺寸变化（拖动分栏 / 最大化）后重排列宽，始终铺满可用宽度。"""
        if getattr(self, '_col_autofit', False):
            QTimer.singleShot(0, self._safe_auto_size)

    def _safe_auto_size(self, tries=0):
        # 延时执行时窗口可能已经关了（控件被销毁），这里兜一下。
        # 关键：PyQt5 里槽函数抛出未捕获异常会直接 qFatal 退出（用户看到的就是“闪退”），
        # 所以这里必须吞掉所有异常，只记日志。
        try:
            ok = self._auto_size_columns()
        except Exception:
            log_exception('自动列宽')
            return
        # 表格宽度还没就绪（刚打开工程 / 窗口还在布局）时，过一会儿再排一次
        if (not ok) and tries < 12 and getattr(self, '_col_autofit', False):
            QTimer.singleShot(120, lambda: self._safe_auto_size(tries + 1))

    def _col_autofit_off(self):
        if not getattr(self, '_col_autofit', False):
            return
        self._col_autofit = False
        self._set_setting('col_autofit', False)
        self.status.showMessage('已切换为手动列宽（左上角右键可重新开启“自动列宽”）', 3500)

    def toggle_col_autofit(self, checked=None):
        self._col_autofit = bool(checked)
        self._set_setting('col_autofit', self._col_autofit)
        if self._col_autofit:
            self._auto_size_columns()
            self.status.showMessage('已开启自动列宽（随窗口宽度铺满）', 3000)
        else:
            self.status.showMessage('已关闭自动列宽，可自由拖动列宽', 3000)

    def toggle_row_autofit(self, checked=None):
        self._row_autofit = bool(checked)
        self._set_setting('row_autofit', self._row_autofit)
        self._apply_row_height_mode()
        self.status.showMessage('已开启自动行高（按内容撑高）' if self._row_autofit
                                else '已关闭自动行高（统一默认行高）', 3000)

    def _auto_size_columns(self):
        """按内容算列宽，并把剩余宽度分给“计算公式”列，让表格铺满窗口。

        返回 True = 已按当前宽度排好；False = 宽度还没就绪，需要稍后重排。
        """
        self._last_autosize_ok = False
        if not getattr(self, '_col_autofit', False):
            return False
        if not hasattr(self, 'table'):
            return False
        cols = self.project.columns
        if not cols or self.table.columnCount() != len(cols):
            return False
        # 预留一条竖向滚动条的宽度：否则“填满”之后滚动条一出现就多出横向滚动条
        avail = self.table.viewport().width() - 26
        if avail < 240:
            return False                        # 窗口还没布局好，等 resize 事件再来
        fm = QFontMetrics(self.table.font())
        measure = getattr(fm, 'horizontalAdvance', None) or fm.width
        # 各列的“不低于”宽度：再挤也不能窄到看不清表头与数字
        MINW = {'code': 72, 'name': 96, 'part': 76, 'unit': 58, 'value': 86,
                'remark': 80, 'no_count': 56}
        widths = []
        nrows = min(self.table.rowCount(), 300)
        for j, cd in enumerate(cols):
            key = cd['key']
            if key == 'no_count':
                widths.append(MINW['no_count'])
                continue
            w = measure(cd['title']) + 26
            for i in range(nrows):
                it = self.table.item(i, j)
                if it is None:
                    continue
                for line in (it.text() or '').split('\n'):
                    if line:
                        w = max(w, measure(line) + 20)
            lo, hi = (220, 640) if key == 'formula' else (MINW.get(key, 72), 340)
            if key == 'unit':
                hi = 110
            widths.append(int(max(lo, min(hi, w))))
        total = sum(widths)
        if total < avail - 4:
            # 宽度有富余：全部给“计算公式”列，表格铺满窗口
            extra = avail - 4 - total
            k = len(cols) - 1
            for i, c in enumerate(cols):
                if c['key'] == 'formula':
                    k = i
                    break
            widths[k] += extra
        elif total > avail - 4:
            # 内容比窗口宽：按优先级压缩（名称 → 备注 → 部位 → 公式 → 编号），
            # 尽量不出横向滚动条；压到最小值还不够，就按比例再缩一档
            MINS = {'formula': 220, 'name': 90, 'part': 70, 'remark': 66,
                    'code': 72, 'unit': 54, 'value': 80, 'no_count': 52}
            target = avail - 4
            overflow = total - target
            order = ['name', 'remark', 'part', 'formula', 'code', 'value', 'unit']
            idx = {c['key']: i for i, c in enumerate(cols)}
            for key in order:
                if overflow <= 0:
                    break
                i = idx.get(key)
                if i is None:
                    continue
                lo = MINS.get(key, 72)
                can = widths[i] - lo
                if can <= 0:
                    continue
                cut = min(can, overflow)
                widths[i] -= int(cut)
                overflow -= cut
            if overflow > 0:
                # 还差一口气：所有列按“可缩余量”等比例再压一次（不压到 0）
                # 注意：这里必须保持整数，setColumnWidth 只接受 int，
                # 传 float 会抛 TypeError，而槽函数里的未捕获异常会让 PyQt5 直接 qFatal 退出（表现为闪退）
                squeezable = {i: int(widths[i] - max(36, int(MINS.get(cols[i]['key'], 48) * 0.7)))
                              for i in range(len(widths))}
                can_total = sum(v for v in squeezable.values() if v > 0)
                if can_total > 0:
                    for i, can in squeezable.items():
                        if can <= 0:
                            continue
                        cut = int(overflow * (can / can_total))
                        widths[i] = int(widths[i] - min(can, cut))
        # 兜底：全部收成合法整数，避免任何 float 漏到 setColumnWidth / setRowHeight
        widths = [max(56, int(round(w))) for w in widths]
        for j, w in enumerate(widths):
            if self.table.columnWidth(j) != w:
                self.table.setColumnWidth(j, w)
        self._last_autosize_ok = True
        # 布局还没稳定时算出的宽度可能多出几像素，等一拍后兜底收一次，避免横向滚动条
        QTimer.singleShot(0, self._fit_hbar)
        return True

    def _fit_hbar(self):
        """兜底：若仍出现横向滚动条，按比例把列宽收到刚好铺满（自动列宽模式下）。"""
        try:
            if not getattr(self, '_col_autofit', False):
                return
            if not hasattr(self, 'table') or self.table.columnCount() == 0:
                return
            if not self.table.horizontalScrollBar().isVisible():
                return
            cols = list(range(self.table.columnCount()))
            for _ in range(4):
                vp = self.table.viewport().width()
                over = sum(self.table.columnWidth(j) for j in cols) - vp + 2
                if over <= 0:
                    return
                idxs = [j for j in cols if self.table.columnWidth(j) > 60]
                if not idxs:
                    return
                per = max(1, over // len(idxs))
                for j in idxs:
                    self.table.setColumnWidth(j, int(max(56, self.table.columnWidth(j) - per)))
        except Exception:
            log_exception('列宽兜底')

    # ---------------- 列宽 / 行高自适应 ----------------
    def fit_column_widths(self, cols=None):
        """按内容自动列宽（表头与单元格取最长文本），并把宽度写回列定义。

        V2.5：自适应是默认状态，手动点一次“列宽自适应”只是按内容重排一次，
        不会再像旧版那样把自动列宽关掉（否则用户会觉得“点了反而不管用了”）。
        """
        if cols is None:
            # 全部列：直接走“按内容算 + 余量给公式列 + 铺满窗口”的自动算法
            if not getattr(self, '_col_autofit', False):
                self._col_autofit = True
                self._set_setting('col_autofit', True)
            self._auto_size_columns()
            if getattr(self, '_row_autofit', False):
                self._apply_row_height_mode()
            self.status.showMessage('列宽已自适应（%d 列，随窗口宽度铺满）'
                                    % self.table.columnCount(), 3000)
            return
        hh = self.table.horizontalHeader()
        fm = QFontMetrics(self.table.font())
        measure = getattr(fm, 'horizontalAdvance', None) or fm.width
        cols = list(cols)
        done = 0
        for j in cols:
            if not (0 <= j < len(self.project.columns)):
                continue
            cd = self.project.columns[j]
            w = measure(cd['title']) + 30
            for i in range(self.table.rowCount()):
                it = self.table.item(i, j)
                if it is None:
                    continue
                for line in (it.text() or '').split('\n'):
                    if line:
                        w = max(w, measure(line) + 24)
            w = max(56, min(460, w))
            hh.setSectionResizeMode(j, QHeaderView.Interactive)
            self.table.setColumnWidth(j, w)
            cd['width'] = w
            done += 1
        self._dirty = True
        if getattr(self, '_row_autofit', False):
            self._apply_row_height_mode()
        self.status.showMessage('列宽已自适应（%d 列，已切换为手动列宽）' % done, 3000)

    def fit_row_heights(self, rows=None):
        """按内容自动撑高行，长公式/长备注换行完整显示；rows=None 表示全部行。"""
        self.table.setWordWrap(True)
        if rows is None:
            self._row_autofit = True
            self._set_setting('row_autofit', True)
            self.table.resizeRowsToContents()
            rows = list(range(self.table.rowCount()))
        else:
            rows = [r for r in rows if 0 <= r < self.table.rowCount()]
            for r in rows:
                self.table.resizeRowToContents(r)
        for r in rows:
            self.table.setRowHeight(r, max(30, min(200, self.table.rowHeight(r))))
        self.status.showMessage('行高已自适应（%d 行）' % len(rows), 3000)

    def reset_row_heights(self):
        """取消自动换行，恢复统一默认行高。"""
        self._row_autofit = False
        self._set_setting('row_autofit', False)
        self.table.setWordWrap(False)
        for i in range(self.table.rowCount()):
            self.table.setRowHeight(i, 34)
        self.status.showMessage('已恢复默认行高', 3000)

    def set_row_height_dialog(self, rows=None):
        """输入具体像素值设置行高（手动设定会关闭自动行高，避免被自适应覆盖）。"""
        if rows is None:
            sel = sorted({i.row() for i in self.table.selectionModel().selectedIndexes()})
            cur = self.table.currentRow()
            rows = sel if sel else ([cur] if cur >= 0 else [])
        rows = [r for r in rows if 0 <= r < self.table.rowCount()]
        if not rows:
            QMessageBox.information(self, APP_NAME, '请先选中要设置行高的行。')
            return
        sample = self.table.rowHeight(rows[0]) or 34
        val, ok = QInputDialog.getInt(self, '设置行高', '行高（像素，18–300）：',
                                      sample, 18, 300, 1)
        if not ok:
            return
        self.push_undo('设置行高')
        self._row_autofit = False
        self._set_setting('row_autofit', False)
        self.table.setWordWrap(False)
        for r in rows:
            self.table.setRowHeight(r, val)
        self._dirty = True
        self.status.showMessage('已设置 %d 行的行高为 %d 像素' % (len(rows), val), 3000)

    def set_column_width_dialog(self, cols=None):
        """输入具体像素值设置列宽，并写回列定义（打印/导出同步）。"""
        if cols is None:
            sel = sorted({i.column() for i in self.table.selectionModel().selectedIndexes()})
            cur = self.table.currentColumn()
            cols = sel if sel else ([cur] if cur >= 0 else [])
        cols = [c for c in cols if 0 <= c < self.table.columnCount()]
        if not cols:
            QMessageBox.information(self, APP_NAME, '请先选中要设置列宽的列。')
            return
        sample = self.table.columnWidth(cols[0]) or 100
        val, ok = QInputDialog.getInt(self, '设置列宽', '列宽（像素，30–600）：',
                                      sample, 30, 600, 1)
        if not ok:
            return
        self.push_undo('设置列宽')
        # 手动设定后关闭自动列宽，否则刷新/缩放时会被“自动铺满”覆盖
        self._col_autofit = False
        self._set_setting('col_autofit', False)
        for c in cols:
            self.table.horizontalHeader().setSectionResizeMode(c, QHeaderView.Interactive)
            self.table.setColumnWidth(c, val)
        try:
            self._sync_col_widths_to_model()
        except Exception:
            log_exception('列宽写回列定义')
        self._dirty = True
        self.status.showMessage('已设置 %d 列的列宽为 %d 像素' % (len(cols), val), 3000)

    def fit_all(self):
        self.fit_column_widths()
        self.fit_row_heights()

    def _sync_col_widths_to_model(self):
        """把界面当前列宽写回列定义：打印 / 预览 / 导出 Excel 都用同一套宽度。"""
        table = getattr(self, 'table', None)
        if table is None or not hasattr(self, 'project'):
            return
        for j, cd in enumerate(self.project.columns):
            if j >= table.columnCount():
                break
            w = int(table.columnWidth(j))
            if w > 0 and cd.get('width') != w:
                cd['width'] = w

    def _on_col_width_changed(self, *_args):
        """列宽变化后：写回列定义供打印使用，并按需重排自动行高。"""
        try:
            self._sync_col_widths_to_model()
        except Exception:
            log_exception('列宽写回列定义')
        if not getattr(self, '_row_autofit', False):
            return
        if getattr(self, '_row_refit_pending', False):
            return
        self._row_refit_pending = True
        QTimer.singleShot(90, self._refit_rows_delayed)

    def _refit_rows_delayed(self):
        self._row_refit_pending = False
        try:
            if getattr(self, '_row_autofit', False) and hasattr(self, 'table'):
                self._apply_row_height_mode()
        except Exception:
            log_exception('列宽变化后重排行高')

    def _apply_row_height_mode(self):
        """刷新表格时按当前行高模式还原（自适应 / 默认 34px）。"""
        if getattr(self, '_row_autofit', False):
            self.table.setWordWrap(True)
            self.table.resizeRowsToContents()
            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, max(30, min(200, self.table.rowHeight(i))))
        else:
            self.table.setWordWrap(False)
            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, 34)

    def header_menu(self, pos):
        """表头右键：列的增删。"""
        h = self.table.horizontalHeader()
        col = h.logicalIndexAt(pos)
        menu = QMenu(self)
        menu.addAction('在左侧添加列', lambda: self.add_column(col))
        menu.addAction('在右侧添加列', lambda: self.add_column(col + 1))
        if col >= 0:
            menu.addAction('设置本列列宽…', lambda c=col: self.set_column_width_dialog([c]))
            cd = self.project.columns[col]
            menu.addAction(f'整列填充颜色「{cd["title"]}」…',
                           lambda c=col: self.fill_color_dialog('col', c))
            menu.addAction(f'清除「{cd["title"]}」整列填充',
                           lambda c=col: self.apply_fill('', 'col', c))
            menu.addAction(f'整列转半角「{cd["title"]}」（全角→半角）',
                           lambda c=col: self.narrow_column(c))
            menu.addAction(f'本列列宽自适应「{cd["title"]}」',
                           lambda c=col: self.fit_column_widths([c]))
            menu.addSeparator()
            menu.addAction(f'重命名列「{cd["title"]}」', lambda: self.rename_column(col))
            menu.addSeparator()
            menu.addAction(f'删除列「{cd["title"]}」', lambda: self.delete_column(col))
        menu.exec_(h.mapToGlobal(pos))

    def narrow_column(self, col):
        """表头右键“整列转半角”：该列（全工程所有分部）全角数字/字母/标点
        转半角并合并连续空格。公式/不计/工程量列跳过——公式里的全角括号
        是注释语法，转半角会改变语义。"""
        from ui_import import narrow_width
        if not (0 <= col < len(self.project.columns)):
            return
        cd = self.project.columns[col]
        key = cd['key']
        if key in ('no_count', 'formula', 'value'):
            QMessageBox.information(self, APP_NAME,
                                    '「%s」列不能转半角：公式/工程量里的全角括号是注释语法，'
                                    '转半角会改变计算结果。' % cd['title'])
            return
        self.push_undo('整列转半角')
        n = 0
        for sec in self.project.all_sections():
            for row in sec.rows:
                if hasattr(row, key):
                    old = str(getattr(row, key, '') or '')
                else:
                    old = str((row.custom or {}).get(key, ''))
                new = narrow_width(old)
                if new != old:
                    if hasattr(row, key):
                        setattr(row, key, new)
                    else:
                        row.custom[key] = new
                    n += 1
        self.refresh_table()
        self._dirty = True
        self.status.showMessage('已将「%s」列 %d 个单元格转为半角（Ctrl+Z 可撤销）' % (cd['title'], n), 5000)

    def add_column(self, index):
        name, ok = QInputDialog.getText(self, '添加列', '列名称：', QLineEdit.Normal, '新列')
        if not ok or not name.strip():
            return
        name = name.strip()
        key = 'c_' + uuid.uuid4().hex[:8]
        self.push_undo('添加列')
        self.project.columns.insert(index, {'key': key, 'title': name, 'width': 120, 'builtin': False})
        self.refresh_table()
        self._dirty = True
        self.status.showMessage(f'已添加列「{name}」', 3000)

    def rename_column(self, col):
        cd = self.project.columns[col]
        name, ok = QInputDialog.getText(self, '重命名列', '列名称：', QLineEdit.Normal, cd['title'])
        if ok and name.strip():
            self.push_undo('重命名列')
            cd['title'] = name.strip()
            self.refresh_table()
            self._dirty = True

    def delete_column(self, col):
        cd = self.project.columns[col]
        if QMessageBox.question(self, '确认删除', f'确定删除列「{cd["title"]}」？该列数据也会删除。',
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        key = cd['key']
        self.push_undo('删除列')
        for sec in self.project.all_sections():
            for row in sec.rows:
                row.custom.pop(key, None)
        self.project.columns.pop(col)
        self.refresh_table()
        self._dirty = True

    def toggle_no_count(self):
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            sel = [self.table.model().index(self.table.currentRow(), 0)] if self.table.currentRow() >= 0 else []
        if not sel:
            self.status.showMessage('请先选中要标记的行', 2500)
            return
        self.push_undo('标记不计')
        for idx in sel:
            if idx.row() < len(self.current_section.rows):
                r = self.current_section.rows[idx.row()]
                r.no_count = not r.no_count
        self.project.recalc()
        self.refresh_table()
        self._dirty = True

    def find_replace(self):
        """查找（Ctrl+F）：按编号/项目名称/单位/部位等定位到具体行。"""
        dlg = getattr(self, '_find_dlg', None)
        if dlg is None:
            dlg = FindDialog(self)
            self._find_dlg = dlg
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
        if dlg.kw.text().strip():
            dlg._invalidate()
        dlg.kw.setFocus()
        dlg.kw.selectAll()

    def goto_cell(self, sec, row_index, key, kw=''):
        """把工作区定位到某个分部的某一行某一列（查找/定位共用）。"""
        if sec is None:
            return
        if sec is not self.current_section:
            self.select_section(sec)
        else:
            self.refresh_table()
        if not (0 <= row_index < self.table.rowCount()):
            return
        col = 0
        for j, cd in enumerate(self.project.columns):
            if cd['key'] == key:
                col = j
                break
        self.table.setCurrentCell(row_index, col)
        it = self.table.item(row_index, col)
        if it is not None:
            self.table.scrollToItem(it, QAbstractItemView.PositionAtCenter)
        self.table.setFocus()
        self.status.showMessage('已定位到「%s」第 %d 行%s'
                                % (sec.name, row_index + 1,
                                   ('：' + kw) if kw else ''), 4000)

    # ---------------- 填充颜色（整行 / 整列 / 单元格） ----------------
    def fill_color_dialog(self, scope=None, col=None):
        """给选中的整行、整列或单元格设置底色。col 指定整列时的列号（表头右键用）。"""
        if col is None and self.table.currentRow() < 0:
            QMessageBox.information(self, APP_NAME, '请先在表格里选中要填充的行、列或单元格。')
            return
        if scope is None:
            idxs = self.table.selectedIndexes()
            ncol = max(1, self.table.columnCount())
            rows = {i.row() for i in idxs}
            scope = 'row' if (idxs and len(rows) == 1 and len(idxs) >= ncol) else 'cell'
        dlg = FillColorDialog(self, scope=scope, color=getattr(self, '_last_fill', ''))
        if dlg.exec_() != QDialog.Accepted:
            return
        self._last_fill = dlg.color
        self.apply_fill(dlg.color, dlg.current_scope(), col=col)

    def apply_fill(self, color, scope, col=None):
        """按范围写入填充色；color 为空串表示清除。"""
        sel = self.table.selectedIndexes()
        rows = sorted({i.row() for i in sel})
        if not rows:
            rows = [self.table.currentRow()]
        sec = self.current_section
        self.push_undo()
        cnt = 0
        if scope == 'row':
            for r in rows:
                if 0 <= r < len(sec.rows):
                    sec.rows[r].fill = color or ''
                    if color:
                        sec.rows[r].cell_fill.clear()   # 整行统一着色时清掉零散单元格色
                    cnt += 1
        elif scope == 'col':
            c = col if col is not None else self.table.currentColumn()
            if 0 <= c < len(self.project.columns):
                cd = self.project.columns[c]
                cd['fill'] = color or ''
                for s in self.project.all_sections():
                    for r in s.rows:
                        (r.cell_fill or {}).pop(cd['key'], None)
                cnt = 1
        else:
            cols = sorted({i.column() for i in sel}) or [self.table.currentColumn()]
            for r in rows:
                if not (0 <= r < len(sec.rows)):
                    continue
                for c in cols:
                    key = self.col_key(c)
                    if not key:
                        continue
                    if color:
                        sec.rows[r].cell_fill[key] = color
                    else:
                        sec.rows[r].cell_fill.pop(key, None)
                    cnt += 1
        self.refresh_table()
        self._dirty = True
        where = {'row': '整行', 'col': '整列', 'cell': '单元格'}.get(scope, '区域')
        self.status.showMessage(('已填充%s %d 处：%s' % (where, cnt, color))
                                if color else ('已清除%s填充' % where), 3000)

    def clear_fill(self):
        """按当前选择自动判断清除范围：整行 → 整列 → 单元格。"""
        idxs = self.table.selectedIndexes()
        ncol = max(1, self.table.columnCount())
        rows = {i.row() for i in idxs}
        cols = {i.column() for i in idxs}
        if idxs and len(cols) >= ncol:
            self.apply_fill('', 'row')
        elif idxs and len(cols) == 1 and len(rows) > 1:
            self.apply_fill('', 'col')
        else:
            self.apply_fill('', 'cell')

    # ---------------- 单元格字体（V2.9.1） ----------------
    def _apply_item_font(self, item, row, key):
        """把 row.cell_font[key] 里的字体属性应用到表格项。"""
        spec = (getattr(row, 'cell_font', None) or {}).get(key)
        if not spec:
            return
        font = QFont(spec.get('family') or self.table.font().family())
        ps = int(spec.get('size') or 0)
        if ps > 0:
            font.setPointSize(ps)
        font.setBold(bool(spec.get('bold')))
        font.setItalic(bool(spec.get('italic')))
        font.setUnderline(bool(spec.get('underline')))
        font.setStrikeOut(bool(spec.get('strikeout')))
        item.setFont(font)
        color = spec.get('color')
        if color:
            item.setForeground(QColor(color))

    def _apply_item_align(self, item, row, key):
        """把 row.cell_align[key] 里的对齐方式应用到表格项（V3.1）。"""
        a = (getattr(row, 'cell_align', None) or {}).get(key)
        if not a:
            return
        h = {'left': Qt.AlignLeft, 'center': Qt.AlignHCenter,
             'right': Qt.AlignRight}.get(a)
        if h is not None:
            item.setTextAlignment(h | Qt.AlignVCenter)

    def _font_target_cells(self):
        """当前选中的目标格子：[(行对象, 列key, 行号, 列号)]，跳过“不计”列。"""
        idxs = self.table.selectedIndexes()
        sec = self.current_section
        cells = []
        for i in idxs:
            r, c = i.row(), i.column()
            key = self.col_key(c)
            if not key or key == 'no_count':
                continue
            if 0 <= r < len(sec.rows):
                cells.append((sec.rows[r], key, r, c))
        return cells

    def cell_font_setup(self):
        """右键“单元格字体…”：给选中单元格设字体、字号、加粗、斜体、下划线、删除线。"""
        cells = self._font_target_cells()
        if not cells:
            QMessageBox.information(self, APP_NAME, '请先选中要设置字体的单元格。')
            return
        spec0 = ((cells[0][0].cell_font or {}).get(cells[0][1])) or {}
        base = self.table.font()
        dlg = QDialog(self)
        dlg.setWindowTitle('单元格字体（选中 %d 格）' % len(cells))
        v = QVBoxLayout(dlg)
        form = QFormLayout()
        fam = QComboBox()
        from PyQt5.QtGui import QFontDatabase
        families = QFontDatabase().families()
        fam.addItems(families)
        fam.setEditable(True)
        fam.setInsertPolicy(QComboBox.NoInsert)
        fam.setCurrentText(spec0.get('family') or base.family())
        size = QSpinBox()
        size.setRange(6, 24)
        size.setValue(int(spec0.get('size') or (base.pointSize() if base.pointSize() > 0 else 10)))
        form.addRow('字体：', fam)
        form.addRow('大小：', size)
        checks = QHBoxLayout()
        ck_bold = QCheckBox('加粗')
        ck_italic = QCheckBox('斜体')
        ck_under = QCheckBox('下划线')
        ck_strike = QCheckBox('删除线')
        for ck, k in ((ck_bold, 'bold'), (ck_italic, 'italic'),
                      (ck_under, 'underline'), (ck_strike, 'strikeout')):
            ck.setChecked(bool(spec0.get(k)))
            checks.addWidget(ck)
        checks.addStretch(1)
        form.addRow(checks)
        v.addLayout(form)
        preview = QLabel('预览  123.45  挖基坑  A-001')
        preview.setAlignment(Qt.AlignCenter)
        preview.setStyleSheet('padding:8px; background:#fff; border:1px solid #cfd4da;')

        def do_preview():
            f = QFont(fam.currentText())
            f.setPointSize(size.value())
            f.setBold(ck_bold.isChecked())
            f.setItalic(ck_italic.isChecked())
            f.setUnderline(ck_under.isChecked())
            f.setStrikeOut(ck_strike.isChecked())
            preview.setFont(f)

        for w in (fam, size, ck_bold, ck_italic, ck_under, ck_strike):
            if hasattr(w, 'currentIndexChanged'):
                try:
                    w.currentIndexChanged.connect(do_preview)
                except TypeError:
                    pass
            if hasattr(w, 'valueChanged'):
                w.valueChanged.connect(do_preview)
            if hasattr(w, 'stateChanged'):
                w.stateChanged.connect(do_preview)
        fam.lineEdit().textChanged.connect(do_preview)
        do_preview()
        v.addWidget(preview)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        v.addWidget(bb)
        if dlg.exec_() != QDialog.Accepted:
            return
        family = fam.currentText().strip()
        if not family:
            return
        self.push_undo('单元格字体')
        n = 0
        for row, key, r, c in cells:
            row.cell_font[key] = {
                'family': family, 'size': int(size.value()),
                'bold': ck_bold.isChecked(), 'italic': ck_italic.isChecked(),
                'underline': ck_under.isChecked(), 'strikeout': ck_strike.isChecked(),
                'color': (row.cell_font.get(key) or {}).get('color', ''),
            }
            n += 1
        self.refresh_table()
        self._dirty = True
        self.status.showMessage('已设置 %d 个单元格的字体：%s %d 号'
                                % (n, family, int(size.value())), 4000)

    def cell_font_color(self):
        """右键“单元格字体颜色…”：给选中单元格的文字设颜色。"""
        cells = self._font_target_cells()
        if not cells:
            QMessageBox.information(self, APP_NAME, '请先选中要设置字体颜色的单元格。')
            return
        spec0 = ((cells[0][0].cell_font or {}).get(cells[0][1])) or {}
        init = QColor(spec0.get('color') or '#000000')
        color = QColorDialog.getColor(init, self, '单元格字体颜色')
        if not color.isValid():
            return
        self.push_undo('单元格字体颜色')
        n = 0
        for row, key, r, c in cells:
            spec = row.cell_font.setdefault(key, {})
            spec['color'] = color.name()
            n += 1
        self.refresh_table()
        self._dirty = True
        self.status.showMessage('已设置 %d 个单元格的字体颜色：%s' % (n, color.name()), 4000)

    def cell_font_clear(self):
        """右键“清除单元格字体”：恢复默认字体与颜色。"""
        cells = self._font_target_cells()
        if not cells:
            QMessageBox.information(self, APP_NAME, '请先选中要清除字体的单元格。')
            return
        self.push_undo('清除单元格字体')
        n = 0
        for row, key, r, c in cells:
            if (row.cell_font or {}).pop(key, None) is not None:
                n += 1
        self.refresh_table()
        self._dirty = True
        self.status.showMessage('已清除 %d 个单元格的字体设置' % n, 4000)

    # ---------------- 单元格对齐（V3.1） ----------------
    def cell_align_set(self, align):
        """右键“单元格对齐”：'left'/'center'/'right'，'default' 表示恢复默认。"""
        cells = self._font_target_cells()
        if not cells:
            QMessageBox.information(self, APP_NAME, '请先选中要设置对齐的单元格。')
            return
        name = {'left': '靠左', 'center': '居中', 'right': '靠右',
                'default': '默认'}.get(align, align)
        self.push_undo('单元格对齐')
        n = 0
        for row, key, r, c in cells:
            if getattr(row, 'cell_align', None) is None:
                row.cell_align = {}
            if align == 'default':
                (row.cell_align or {}).pop(key, None)
            else:
                row.cell_align[key] = align
            n += 1
        self.refresh_table()
        self._dirty = True
        self.status.showMessage('已设置 %d 个单元格对齐方式：%s' % (n, name), 4000)

    def narrow_cells(self):
        """右键“选中单元格转半角”：全角数字/字母/标点转半角（公式列除外，
        公式里的全角括号是注释语法，转半角会改变语义）。"""
        from ui_import import narrow_width
        idxs = self.table.selectedIndexes()
        sec = self.current_section
        cols = self.project.columns
        targets = []
        for i in idxs:
            r, c = i.row(), i.column()
            if not (0 <= r < len(sec.rows)) or not (0 <= c < len(cols)):
                continue
            key = cols[c]['key']
            if key in ('no_count', 'formula', 'value'):
                continue
            targets.append((sec.rows[r], key))
        if not targets:
            QMessageBox.information(self, APP_NAME, '请先选中要转半角的单元格。')
            return
        self.push_undo('单元格转半角')
        n = 0
        for row, key in targets:
            if hasattr(row, key):
                old = str(getattr(row, key) or '')
            else:
                old = str((row.custom or {}).get(key, ''))
            new = narrow_width(old)
            if new != old:
                if hasattr(row, key):
                    setattr(row, key, new)
                else:
                    row.custom[key] = new
                n += 1
        if n:
            self.refresh_table()
            self._dirty = True
        self.status.showMessage('已把 %d 个单元格的文字转为半角' % n, 4000)

    # ---------------- 工具条补充功能 ----------------
    def show_recent(self):
        """显示最近打开的工程列表（底部带“清除最近打开列表”）。"""
        if not self.recent:
            QMessageBox.information(self, APP_NAME, '暂无最近打开的工程。')
            return
        menu = QMenu(self)
        for path in self.recent:
            menu.addAction(os.path.basename(path), lambda p=path: self.open_project(p))
        menu.addSeparator()
        menu.addAction('清除最近打开列表', self.clear_recent)
        menu.exec_(self.cursor().pos())

    def clear_recent(self):
        """清空最近打开列表并持久化。"""
        n = len(self.recent)
        if n == 0:
            return
        self.recent = []
        self._set_setting('recent', [])
        self.update_recent_menu()
        self.status.showMessage('已清除最近打开列表（%d 条）' % n, 3000)

    def choose_backup_dir(self):
        """选择自动保存备份目录。"""
        d = QFileDialog.getExistingDirectory(self, '选择备份目录', self.autosave_dir or '')
        if d:
            self.autosave_dir = d
            self.settings.setValue('autosave_dir', d)
            self.reset_autosave_timer()
            self.status.showMessage('备份目录已设为：' + d, 4000)

    def copy_rows(self, *args):
        """把选中行按 TSV 复制到剪贴板。"""
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()})
        if not rows:
            return
        cols = self.project.columns
        lines = []
        for r in rows:
            if r < len(self.current_section.rows):
                row = self.current_section.rows[r]
                vals = [self.cell_text_for(row, cd['key']) for cd in cols]
                lines.append('\t'.join(vals))
        QApplication.clipboard().setText('\n'.join(lines))
        self.status.showMessage(f'已复制 {len(lines)} 行', 2500)

    def cut_rows(self, *args):
        """复制并删除选中行。"""
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        if not rows:
            return
        self.copy_rows()
        self.push_undo('剪切行')
        for r in rows:
            if r < len(self.current_section.rows):
                del self.current_section.rows[r]
        self.project.recalc()
        self.refresh_table()
        self._dirty = True
        self.status.showMessage(f'已剪切 {len(rows)} 行', 2500)

    # ---------------- 复制 / 拖动填充 / 格式刷 / 排序 / 筛选（V2.10.0） ----------------

    def copy_cells(self, *args):
        """按选中区域复制（矩形区域按区域，非矩形退回整行复制）。"""
        idxs = self.table.selectedIndexes()
        if not idxs:
            return
        rect = self.table.sel_rect()
        sec = self.current_section
        cols = self.project.columns
        if rect is None:
            self.copy_rows()
            return
        r0, r1, c0, c1 = rect
        lines = []
        for r in range(r0, r1 + 1):
            if r < len(sec.rows):
                row = sec.rows[r]
                vals = [self.cell_text_for(row, cols[c]['key'])
                        for c in range(c0, c1 + 1)]
            else:
                vals = [''] * (c1 - c0 + 1)
            lines.append('\t'.join(vals))
        QApplication.clipboard().setText('\n'.join(lines))
        self.status.showMessage('已复制 %d 行 × %d 列' % (r1 - r0 + 1, c1 - c0 + 1), 2500)

    # ---- 拖动填充 ----

    def apply_fill_drag(self, src, tgt):
        """填充柄松开：把源区值按 WPS 规则填充到目标区（自动序号）。"""
        sec = self.current_section
        cols = self.project.columns
        r0, r1, c0, c1 = src
        tr0, tr1, tc0, tc1 = tgt
        vertical = (tc0, tc1) == (c0, c1)       # 纵向填充
        self.push_undo('拖动填充')
        n = 0
        if vertical:
            offsets = list(range(tr0 - r0, tr1 - r0 + 1))
            for c in range(c0, c1 + 1):
                key = cols[c]['key'] if c < len(cols) else ''
                if not key or key == 'value' or r0 >= len(sec.rows):
                    continue
                svals = [self.cell_text_for(sec.rows[r], key)
                         for r in range(r0, min(r1 + 1, len(sec.rows)))]
                if key == 'formula':
                    # 公式列只整体复制——尾号递增会悄悄改掉公式里的常数
                    newvals = [svals[i % max(1, len(svals))] for i in offsets]
                else:
                    newvals = fill_series(svals, offsets)
                for i, t in enumerate(range(tr0, tr1 + 1)):
                    if t >= len(sec.rows):
                        break
                    self._fill_write(sec.rows[t], key, newvals[i],
                                     sec.rows[r0 + (i % max(1, len(svals)))])
                    n += 1
        else:
            offsets = list(range(tc0 - c0, tc1 - c0 + 1))
            for r in range(r0, r1 + 1):
                if r >= len(sec.rows):
                    break
                row = sec.rows[r]
                svals = [self.cell_text_for(row, cols[c]['key'])
                         for c in range(c0, min(c1 + 1, len(cols)))]
                newvals = fill_series(svals, offsets)   # 横向一般不涉及公式列
                for i, t in enumerate(range(tc0, tc1 + 1)):
                    if t >= len(cols):
                        break
                    self._fill_write(row, cols[t]['key'], newvals[i], row)
                    n += 1
        self.project.recalc()
        self.refresh_table()
        self._dirty = True
        self.status.showMessage('已填充 %d 个单元格（Ctrl+Z 可撤销）' % n, 4000)

    def _fill_write(self, row, key, val, src_row):
        """把填充值写进单元格（no_count 用布尔，其它用文本）。"""
        if key == 'no_count':
            row.no_count = bool(getattr(src_row, 'no_count', False))
            return
        if hasattr(row, key):
            setattr(row, key, str(val))
        else:
            row.custom[key] = str(val)

    # ---- 格式刷 ----

    def start_format_painter(self):
        """拾取当前单元格格式（字体 + 填充色），等待点击目标区域。"""
        r, c = self.table.currentRow(), self.table.currentColumn()
        sec = self.current_section
        cols = self.project.columns
        if not (0 <= r < len(sec.rows)) or not (0 <= c < len(cols)):
            QMessageBox.information(self, APP_NAME, '请先选中要复制格式的单元格。')
            return
        key = cols[c]['key']
        row = sec.rows[r]
        font = dict((row.cell_font or {}).get(key) or {})
        fill = self.row_fill_for(row, cols[c])
        align = str((getattr(row, 'cell_align', None) or {}).get(key) or '')
        self._paint_fmt = {'font': font, 'fill': fill, 'align': align}
        self.table.paint_mode = True
        self.table.setCursor(Qt.CrossCursor)
        tip = []
        if font:
            tip.append('字体')
        if fill:
            tip.append('底色')
        if align:
            tip.append('对齐')
        self.status.showMessage('格式刷已拾取（%s）：点击或拖选目标单元格，'
                                'Esc 取消' % ('+'.join(tip) or '默认格式'), 6000)

    def cancel_painter(self):
        """Esc：取消格式刷 / 清除填充预览。"""
        if self._paint_fmt is not None:
            self._paint_fmt = None
            self.table.paint_mode = False
            self.table.setCursor(Qt.ArrowCursor)
            self.status.showMessage('已取消格式刷', 2000)

    def painter_apply(self, rect):
        """把拾取的格式刷到目标区域。"""
        fmt = self._paint_fmt
        self.table.paint_mode = False
        self.table.setCursor(Qt.ArrowCursor)
        self._paint_fmt = None
        if not fmt:
            return
        sec = self.current_section
        cols = self.project.columns
        r0, r1, c0, c1 = rect
        self.push_undo('格式刷')
        n = 0
        for r in range(r0, r1 + 1):
            if r >= len(sec.rows):
                break
            row = sec.rows[r]
            for c in range(c0, c1 + 1):
                if c >= len(cols):
                    break
                key = cols[c]['key']
                if key in ('value', 'no_count'):
                    continue
                if fmt['font']:
                    row.cell_font[key] = dict(fmt['font'])
                else:
                    (row.cell_font or {}).pop(key, None)
                if fmt['fill']:
                    row.cell_fill[key] = fmt['fill']
                else:
                    (row.cell_fill or {}).pop(key, None)
                if getattr(row, 'cell_align', None) is None:
                    row.cell_align = {}
                if fmt.get('align'):
                    row.cell_align[key] = fmt['align']
                else:
                    (row.cell_align or {}).pop(key, None)
                n += 1
        self.refresh_table()
        self._dirty = True
        self.status.showMessage('格式刷已应用 %d 个单元格（Ctrl+Z 可撤销）' % n, 4000)

    # ---- 排序 ----

    def sort_dialog(self):
        from ui_sortfilter import SortDialog, sort_rows_with_colors
        cols = self.project.columns
        sec = self.current_section
        col = self.table.currentColumn()
        if not (0 <= col < len(cols)):
            col = 0
        res = SortDialog.exec(self, [c['title'] for c in cols], col)
        if not res:
            return
        cd = cols[res['col']]
        key = cd['key']
        items = [(self.cell_text_for(row, key), self.row_fill_for(row, cd))
                 for row in sec.rows]
        order = sort_rows_with_colors(items, res['mode'], res['pattern'],
                                      res['reverse'], res.get('regex_num', True))
        if order == list(range(len(sec.rows))):
            self.status.showMessage('顺序未变化', 2000)
            return
        self.push_undo('排序')
        sec.rows = [sec.rows[i] for i in order]
        self.project.recalc()
        self.refresh_table()
        self._dirty = True
        self.status.showMessage('已按「%s」%s排序（%s，Ctrl+Z 可撤销）'
                                % (cd['title'], res['mode'],
                                   '降序' if res['reverse'] else '升序'), 4000)

    # ---- 筛选 ----

    def filter_dialog(self):
        from ui_sortfilter import FilterDialog, filter_match
        cols = self.project.columns
        sec = self.current_section
        col = self.table.currentColumn()
        if not (0 <= col < len(cols)):
            col = 0
        cd = cols[col]
        # 收集该列现有填充色供"按颜色"筛选选择
        seen = []
        for row in sec.rows:
            colr = self.row_fill_for(row, cd)
            if colr and colr not in seen:
                seen.append(colr)
        colors = [('', '无填充')] + [(cv, cv) for cv in sorted(seen)]
        res = FilterDialog.exec(self, [c['title'] for c in cols], col, colors)
        if not res:
            return
        self._filter = {'key': cd['key'], 'mode': res['mode'],
                        'value': res['value'], 'color_value': res['color_value']}
        self._apply_filter()

    def _apply_filter(self):
        """按生效的筛选条件隐藏当前分部不符合的行（只隐藏，不删数据）。"""
        f = self._filter
        sec = self.current_section
        cols = self.project.columns
        if f and not any(c['key'] == f['key'] for c in cols):
            self._filter = None                     # 列被删了，筛选失效
            f = None
        if not f:
            for i in range(self.table.rowCount()):
                self.table.setRowHidden(i, False)
            return
        from ui_sortfilter import filter_match
        cd = next(c for c in cols if c['key'] == f['key'])
        shown = 0
        for i, row in enumerate(sec.rows):
            ok = filter_match(self.cell_text_for(row, f['key']),
                              self.row_fill_for(row, cd),
                              f['mode'], f['value'], f.get('color_value'))
            self.table.setRowHidden(i, not ok)
            shown += 1 if ok else 0
        # 筛选生效时把末尾的空占位行也藏起来
        for i in range(len(sec.rows), self.table.rowCount()):
            self.table.setRowHidden(i, True)
        self.status.showMessage('筛选中：「%s」%s%s，显示 %d/%d 行（右键→清除筛选）'
                                % (cd['title'], f['mode'],
                                   f['value'] or f.get('color_value') or '',
                                   shown, len(sec.rows)), 6000)

    def clear_filter(self):
        if self._filter:
            self._filter = None
            self._apply_filter()
            self.status.showMessage('已清除筛选', 2500)

    def show_symbol_picker(self, btn=None):
        """弹出特殊符号面板（工程符号 + 钢筋等级 + 直径/半径），点选即插入。"""
        if btn is None:
            s = self.sender()
            btn = s if isinstance(s, QWidget) else None
        dlg = SymbolPicker(self.insert_symbol, self)
        dlg.adjustSize()
        self._place_popup(dlg, btn)
        dlg.show()
        self._symbol_picker = dlg   # 保持引用，防止被回收

    def focus_symbols(self):
        """旧的“定位符号栏”入口，保留为打开符号面板。"""
        self.show_symbol_picker()

    def toggle_tree_expand(self):
        """展开/折叠全部计算书与分部。"""
        # 依据当前是否已全部展开来决定动作
        any_collapsed = False
        def walk(item):
            nonlocal any_collapsed
            if item.childCount() and not item.isExpanded():
                any_collapsed = True
            for i in range(item.childCount()):
                walk(item.child(i))
        for i in range(self.tree.topLevelItemCount()):
            walk(self.tree.topLevelItem(i))
        def apply(item, expand):
            item.setExpanded(bool(expand and item.childCount()))
            for i in range(item.childCount()):
                apply(item.child(i), expand)
        for i in range(self.tree.topLevelItemCount()):
            apply(self.tree.topLevelItem(i), any_collapsed)
        self.status.showMessage('已' + ('展开' if any_collapsed else '折叠') + '全部节点', 2000)

    def _set_hilite(self, which, on):
        """统一处理当前行 / 列高亮开关：状态、设置、按钮外观、重绘。"""
        on = bool(on)
        if which == 'row':
            self._hilite_row = on
            self._set_setting('hilite_row', on)
        else:
            self._hilite_col = on
            self._set_setting('hilite_col', on)
        btn = (getattr(self, '_hilite_buttons', {}) or {}).get(
            '选中行高亮' if which == 'row' else '选中列高亮')
        if btn is not None and btn.isChecked() != on:
            btn.blockSignals(True)
            btn.setChecked(on)
            btn.blockSignals(False)
        self.table.viewport().update()
        label = '当前行' if which == 'row' else '当前列'
        self.status.showMessage(('已开启' if on else '已关闭') + label + '高亮', 2000)

    def toggle_row_highlight(self, checked=None):
        """开关当前行高亮（淡色底纹，不改选择模式）。"""
        self._set_hilite('row', (not self._hilite_row) if checked is None else checked)

    def toggle_col_highlight(self, checked=None):
        """开关当前列高亮（淡色底纹，不改选择模式）。"""
        self._set_hilite('col', (not self._hilite_col) if checked is None else checked)

    def _apply_table_font(self, family=None, size=None, refit=True):
        """把表格字体应用到工作区（单元格委托、行高自适应都以 table.font() 为基准）。

        注意：主窗口 QSS 里 QTableWidget 的 font-size 需要同步重建，
        否则样式表会覆盖 setFont 的字号（V2.8.2 踩坑）。
        """
        if family is None:
            family = getattr(self, '_font_family', '')
            size = getattr(self, '_font_size', 0)
        try:
            size = int(size)
        except Exception:
            size = 0
        table = getattr(self, 'table', None)
        if table is None:
            return
        if not family:
            # 没设过字体：恢复软件默认（微软雅黑 10pt）
            family = '微软雅黑'
            size = 10
        size = max(6, min(24, size))
        table.setFont(QFont(family, size))
        # 主窗口样式表里 QTableWidget 的 font-size 跟着改
        try:
            self.setStyleSheet(self._main_qss(size))
        except Exception:
            log_exception('表格字体样式表更新')
        if refit:
            try:
                if getattr(self, '_col_autofit', False):
                    self._auto_size_columns()
                if getattr(self, '_row_autofit', False):
                    self._apply_row_height_mode()
                table.viewport().update()
            except Exception:
                log_exception('表格字体应用后重排')

    def font_setup(self):
        """设置表格字体：字体类型来自系统字库（宋体/黑体/仿宋等），大小 6-24。"""
        from PyQt5.QtGui import QFontDatabase
        cur = self.table.font()
        dlg = QDialog(self)
        dlg.setWindowTitle('表格字体')
        v = QVBoxLayout(dlg)
        form = QFormLayout()
        fam = QComboBox()
        families = QFontDatabase().families()
        fam.addItems(families)
        fam.setEditable(True)
        fam.setInsertPolicy(QComboBox.NoInsert)
        if cur.family() in families:
            fam.setCurrentText(cur.family())
        else:
            fam.setEditText(cur.family())
        size = QSpinBox()
        size.setRange(6, 24)
        size.setValue(max(6, cur.pointSize() if cur.pointSize() > 0 else 10))
        form.addRow('字体：', fam)
        form.addRow('大小：', size)
        v.addLayout(form)
        tip = QLabel('更改应用到工作区表格，行高 / 列宽会按新字体自动重排')
        tip.setStyleSheet('color:#889; font-size:8.5pt;')
        v.addWidget(tip)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        v.addWidget(bb)
        if dlg.exec_() != QDialog.Accepted:
            return
        family = fam.currentText().strip()
        pt = int(size.value())
        if not family:
            return
        self._font_family = family
        self._font_size = pt
        self._set_setting('table_font_family', family)
        self._set_setting('table_font_size', pt)
        self._apply_table_font(family, pt)
        self.status.showMessage('表格字体已设为 %s %d 号' % (family, pt), 3000)

    def precision_setup(self):
        """自由定制小数精度（工程量显示位数）。"""
        val, ok = QInputDialog.getInt(self, '定制精度', '工程量小数位数（0-6）：',
                                      self.precision, 0, 6, 1)
        if ok:
            self.precision = val
            self._set_setting('precision', val)
            self.refresh_table()
            self.status.showMessage('工程量精度已设为 %d 位小数' % val, 3000)

    def insert_note_tag(self):
        """在当前行备注追加标签（注释标注）。"""
        row = self.table.currentRow()
        if row < 0 or row >= len(self.current_section.rows):
            QMessageBox.information(self, APP_NAME, '请先选中一行。')
            return
        text, ok = QInputDialog.getText(self, '注释标签', '标注内容：', QLineEdit.Normal, '')
        if ok and text.strip():
            self.push_undo('注释标签')
            r = self.current_section.rows[row]
            tag = '[%s]' % text.strip()
            r.remark = (r.remark + ' ' if r.remark else '') + tag
            self.refresh_table()
            self._dirty = True

    def column_setup(self):
        """数据列定义：集中管理列的显示与宽度。"""
        from PyQt5.QtWidgets import QListWidget, QListWidgetItem
        dlg = QDialog(self)
        dlg.setWindowTitle('数据列定义')
        dlg.resize(520, 460)
        lay = QVBoxLayout(dlg)
        lay.addWidget(QLabel('双击列名可重命名；右侧可调整宽度（像素），可新增/删除自定义列。'))
        lst = QListWidget()
        for cd in self.project.columns:
            it = QListWidgetItem(f"{cd['title']}    宽度：{cd.get('width',100)}")
            it.setData(Qt.UserRole, cd['key'])
            it.setFlags(it.flags() | Qt.ItemIsEditable)
            lst.addItem(it)
        lay.addWidget(lst, 1)
        bar = QHBoxLayout()
        btn_editw = QPushButton('设置宽度')
        btn_add = QPushButton('新增列')
        btn_del = QPushButton('删除列')
        bar.addWidget(btn_editw); bar.addWidget(btn_add); bar.addWidget(btn_del)
        bar.addStretch(1)
        lay.addLayout(bar)

        def cur_cd():
            it = lst.currentItem()
            if not it:
                return None, it
            key = it.data(Qt.UserRole)
            for cd in self.project.columns:
                if cd['key'] == key:
                    return cd, it
            return None, it

        def rename(it):
            cd, _ = cur_cd()
            if cd and it.text().strip():
                cd['title'] = it.text().split('    宽度')[0].strip()
                it.setText(f"{cd['title']}    宽度：{cd.get('width',100)}")
                self.refresh_table()
                self._dirty = True
        lst.itemChanged.connect(rename)

        def set_width():
            cd, it = cur_cd()
            if not cd:
                return
            w, ok = QInputDialog.getInt(self, '列宽', f"「{cd['title']}」宽度（像素）：",
                                        cd.get('width', 100), 40, 1000, 10)
            if ok:
                cd['width'] = w
                it.setText(f"{cd['title']}    宽度：{w}")
                self.refresh_table()
                self._dirty = True
        btn_editw.clicked.connect(set_width)

        def add_col():
            name, ok = QInputDialog.getText(self, '新增列', '列名称：', QLineEdit.Normal, '新列')
            if ok and name.strip():
                key = 'c_' + uuid.uuid4().hex[:8]
                self.project.columns.append({'key': key, 'title': name.strip(), 'width': 120, 'builtin': False})
                it = QListWidgetItem(f"{name.strip()}    宽度：120")
                it.setData(Qt.UserRole, key)
                it.setFlags(it.flags() | Qt.ItemIsEditable)
                lst.addItem(it)
                self.refresh_table()
                self._dirty = True
        btn_add.clicked.connect(add_col)

        def del_col():
            cd, it = cur_cd()
            if not cd:
                return
            if QMessageBox.question(self, '确认', f"删除列「{cd['title']}」？该列数据也会删除。",
                                    QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
                return
            key = cd['key']
            for sec in self.project.all_sections():
                for row in sec.rows:
                    row.custom.pop(key, None)
            self.project.columns.remove(cd)
            lst.takeItem(lst.row(it))
            self.refresh_table()
            self._dirty = True
        btn_del.clicked.connect(del_col)

        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(dlg.reject)
        bb.accepted.connect(dlg.accept)
        lay.addWidget(bb)
        dlg.exec_()

    # ---------------- 工程 ----------------
    def new_project(self):
        if not self.confirm_discard():
            return
        # 新建时一并填写工程（项目）信息：名称、建设单位、施工单位、设计单位、日期、备注
        proj = Project('未命名工程')
        dlg = ProjectInfoDialog(proj, self, for_new=True)
        if dlg.exec_() != QDialog.Accepted:
            return
        dlg.apply()
        proj.name = proj.info.get('proj_name', '').strip() or '未命名工程'
        self.project = proj
        self.current_file = None
        self.current_section = self.project.root
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._pending_undo = None
        self.remember_file('')
        self.update_title()
        self.refresh_tree()
        self.refresh_table()
        self.status.showMessage('已新建工程：' + self.project.name, 4000)

    def confirm_discard(self):
        return QMessageBox.question(self, APP_NAME, '当前工程未保存的修改将丢失，继续？',
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes

    def save_project(self):
        if not self.current_file:
            return self.save_as()
        try:
            self.project.save(self.current_file)
            self._dirty = False
            self.remember_file(self.current_file)
            self.update_title()
            self.reset_autosave_timer()
            self.status.showMessage('已保存：' + os.path.basename(self.current_file), 3000)
        except Exception as e:
            QMessageBox.critical(self, '保存失败', str(e))

    def save_as(self):
        path, _ = QFileDialog.getSaveFileName(self, '另存工程',
                                              self.project.name + FILE_EXT,
                                              '工程量计算书工程 (%s)' % FILE_EXT + ';;所有文件 (*)')
        if path:
            if not path.lower().endswith(FILE_EXT):
                path += FILE_EXT
            self.current_file = path
            # 工程名还是“未命名工程”时，用文件名当工程名，标题栏不再显示“未命名工程”
            base = self.name_from_path(path)
            if (self.project.name or '').strip() in ('', '未命名工程'):
                self.project.name = base
                self.project.info['proj_name'] = base
                self.refresh_tree()
            self.save_project()
            self.add_recent(path)
            self.update_title()

    def add_recent(self, path):
        if path in self.recent:
            self.recent.remove(path)
        self.recent.insert(0, path)
        self.recent = self.recent[:10]
        self._set_setting('recent', self.recent)

    def update_recent_menu(self):
        pass

    def edit_info(self):
        dlg = ProjectInfoDialog(self.project, self)
        if dlg.exec_() == QDialog.Accepted:
            dlg.apply()
            self.update_title()
            self.refresh_tree()
            self.status.showMessage('工程信息已更新', 3000)

    def edit_constants(self):
        """批量编辑 / 粘贴常量（每行一个 @名称=数值）。"""
        text = '\n'.join(f'{k}={fmt_const(v)}' for k, v in self.project.constants.items())
        hint = ('每行一个：@名称=数值（常量全部自己定义）\n'
                '例如：@层高=3.6、@墙厚=0.24、@Φ8=0.395\n'
                '用法：公式里写 @层高×4、@Φ8×12\n'
                '需要钢筋单重可从“⋯”菜单一键插入模板')
        new, ok = QInputDialog.getMultiLineText(self, '批量编辑常量', hint, text)
        if ok:
            consts = {}
            notes = dict(getattr(self.project, 'constant_notes', None) or {})
            bad = []
            for line in new.splitlines():
                if '=' in line:
                    k, v = line.split('=', 1)
                    k, v = k.strip(), v.strip()
                    if not k:
                        continue
                    if not k.startswith('@'):
                        k = '@' + k
                    try:
                        consts[k] = float(v)
                    except ValueError:
                        bad.append(v)
            if bad:
                QMessageBox.warning(self, '常量错误',
                                    '以下数值不是数字，已跳过：\n' + '、'.join(bad[:8]))
            self.push_undo('批量改常量')
            self.project.constants = consts
            self.project.constant_notes = {k: v for k, v in notes.items() if k in consts}
            self.project.recalc()
            self.refresh_table()
            self.update_constants_display()
            self._dirty = True

    def recalc_all(self):
        self.project.recalc()
        self.refresh_table()
        self.status.showMessage('已重算全部公式', 3000)

    # ---------------- 报表 ----------------
    def export_excel(self):
        self.project.recalc()
        path, _ = QFileDialog.getSaveFileName(self, '导出 Excel', self.project.name + '-工程量计算书.xlsx',
                                              'Excel 文件 (*.xlsx);;所有文件 (*)')
        if not path:
            return
        if not path.lower().endswith('.xlsx'):
            path += '.xlsx'
        try:
            export_xlsx(self.project, path, self.precision, self.project.report)
            QMessageBox.information(self, '导出成功', f'已导出：\n{path}\n\n可直接用 WPS 表格或 Excel 打开。')
        except Exception as e:
            QMessageBox.critical(self, '导出失败', str(e))

    def export_wps(self):
        """导出 WPS 表格（.xlsx，WPS 表格原生支持，可再另存为 .et）。"""
        self.project.recalc()
        path, _ = QFileDialog.getSaveFileName(
            self, '导出 WPS 表格', self.project.name + '-工程量计算书.xlsx',
            'WPS 表格 (*.xlsx *.et);;所有文件 (*)')
        if not path:
            return
        low = path.lower()
        if low.endswith('.et'):
            path = path[:-3] + '.xlsx'      # .et 为 WPS 私有格式，改用 xlsx 生成
        elif not low.endswith('.xlsx'):
            path += '.xlsx'
        try:
            export_xlsx(self.project, path, self.precision, self.project.report)
        except Exception as e:
            QMessageBox.critical(self, '导出失败', str(e))
            return
        self.status.showMessage('已导出 WPS 表格：' + os.path.basename(path), 4000)
        if QMessageBox.question(self, '导出成功',
                                f'已生成：\n{path}\n\n是否立即用 WPS 表格打开？',
                                QMessageBox.Yes | QMessageBox.No,
                                QMessageBox.Yes) == QMessageBox.Yes:
            try:
                os.startfile(path)          # 交给系统默认程序（通常是 WPS 表格）
            except Exception as e:
                QMessageBox.warning(self, '打开失败',
                                    f'请手动打开该文件：\n{path}\n\n{e}')

    def page_setup(self):
        """页面设置：页边距（上下左右）、页眉页脚、纸张方向、对齐方式。"""
        dlg = ReportSettingsDialog(self.project.report, self)
        if dlg.exec_() == QDialog.Accepted:
            dlg.apply()
            self.project.report = dlg.report
            self.print_preview()
            self.status.showMessage('页面设置已保存（打印预览与导出时生效）', 3000)

    def report_settings(self):
        self.page_setup()

    def _report_style(self):
        """报表的单元格对齐方式。"""
        rep = self.project.report or {}
        va = {'center': 'middle', 'top': 'top', 'bottom': 'bottom'}.get(rep.get('valign', 'center'), 'middle')
        ha = rep.get('halign', 'left')
        return va, ha

    def _active_columns(self):
        """当前打印区域涉及的列；未设置打印区域时返回全部列。"""
        area = getattr(self, '_print_area', None)
        if area and area.get('keys'):
            keys = set(area['keys'])
            return [c for c in self.project.columns if c.get('key') in keys]
        return self.project.columns

    def _report_col_pct(self):
        """各列宽度百分比：优先取界面上当前的真实列宽（含手动拖动后的宽度）。"""
        table = getattr(self, 'table', None)
        cols = self._active_columns()
        key_index = {c.get('key'): i for i, c in enumerate(self.project.columns)}
        cw = []
        for i, c in enumerate(cols):
            w = 0.0
            src = key_index.get(c.get('key'), i)
            if table is not None and src < table.columnCount():
                try:
                    w = float(table.columnWidth(src))
                except Exception:
                    w = 0.0
            if w <= 0:
                try:
                    w = float(c.get('width', 100) or 100)
                except Exception:
                    w = 100.0
            if c.get('key') == 'no_count':
                w = max(w, 72.0)
            cw.append(w)
        tot = sum(cw) or 1.0
        pct = [w / tot * 100.0 for w in cw]
        # 窄列（编号/单位/工程量/备注/不计）保证最小占比，避免表头与数字被挤成两行
        MINPCT = {'code': 12.0, 'unit': 6.5, 'value': 9.5, 'no_count': 6.0, 'remark': 10.0}
        for _ in range(2):
            pct = [max(pct[i], MINPCT.get(c['key'], 0.0))
                   for i, c in enumerate(cols)]
            _sum = sum(pct) or 1.0
            pct = [x / _sum * 100.0 for x in pct]
        return pct

    def _report_row_pad(self, row=None):
        """按界面当前行高生成单元格上下内边距（Qt 富文本只认 padding，不认 height）。"""
        h = self._row_height_default if hasattr(self, '_row_height_default') else 34
        if row is not None:
            h = getattr(self, '_row_height_cache', {}).get(id(row), h)
        try:
            h = float(h)
        except Exception:
            h = 34.0
        base = QFontMetrics(QFont('微软雅黑', 9)).height() + 6
        pad = 3 + max(0, int(h - base)) // 2
        return int(max(3, min(60, pad)))

    def _report_table_open(self, pct):
        """表格开始标签（含列宽比例）。"""
        va, ha = self._report_style()
        return ('<table border="1" cellspacing="0" cellpadding="3" width="100%" '
                f'style="font-size:9pt;vertical-align:{va};text-align:{ha}">'
                '<colgroup>' + ''.join('<col width="%.1f%%">' % x for x in pct)
                + '</colgroup>')

    def footer_fields(self):
        """页脚左 / 中 / 右文本；旧版单一 footer 字段作为中间文字兼容。"""
        rep = self.project.report or {}
        left = str(rep.get('footer_left', '') or '')
        center = str(rep.get('footer_center', '') or '')
        right = str(rep.get('footer_right', '') or '')
        if not (left or center or right):
            center = str(rep.get('footer', '') or '')
        return left, center, right

    def report_head_html(self):
        """报表「头部」：报表页眉 + 标题 + 表头行 —— 每一页都会重复打印。"""
        rep = self.project.report or {}
        cols = self._active_columns()
        headers = [c['title'] for c in cols]
        pct = self._report_col_pct()
        parts = []
        if rep.get('header'):
            parts.append(f'<div style="text-align:center;font-size:9pt;color:#888;'
                         f'border-bottom:1px solid #ccc;padding:3px">{rep["header"]}</div>')
        title = rep.get('title') or (self.project.name + ' 工程量计算书')
        parts.append(f'<h2 style="text-align:center;font-size:15pt">{title}</h2>')
        parts.append(self._report_table_open(pct))
        parts.append('<tr>' + ''.join(
            '<th width="%.1f%%" style="background:#D9E2F3">%s</th>' % (pct[i], h)
            for i, h in enumerate(headers)) + '</tr>')
        parts.append('</table>')
        return ''.join(parts)

    def report_body_html(self):
        """报表「正文」：数据表格（不含表头行）+ 页脚。"""
        va, ha = self._report_style()
        area = getattr(self, '_print_area', None)
        cols = self._active_columns()
        ncol = len(cols)
        pct = self._report_col_pct()
        parts = [self._report_table_open(pct)]

        def walk(sec, level=0):
            parts.append(f'<tr><td colspan="{ncol}" style="background:#fff2cc"><b>{"　"*level}{sec.name}</b></td></tr>')
            for ri, r in enumerate(sec.rows):
                if area and (sec is not area.get('section') or ri not in area.get('rows', ())):
                    continue
                tds = []
                for k, cd in enumerate(cols):
                    key = cd['key']
                    if key == 'no_count':
                        text = '是' if r.no_count else ''
                    elif key == 'value':
                        text = fmt_num(r.value, precision_for_unit(r.unit, self.precision))
                    elif hasattr(r, key):
                        text = str(getattr(r, key))
                    else:
                        text = str(r.custom.get(key, ''))
                    # 填充色优先级：单元格 > 整行 > 整列
                    bg = (getattr(r, 'cell_fill', None) or {}).get(key) or \
                        getattr(r, 'fill', '') or cd.get('fill', '') or ''
                    pad = self._report_row_pad(r)
                    style = ('vertical-align:%s;padding-top:%dpx;padding-bottom:%dpx;'
                             % (va, pad, pad)) + ('background:%s;' % bg if bg else '')
                    tds.append('<td width="%.1f%%" style="%s">%s</td>'
                               % (pct[k] if k < len(pct) else 0, style, text))
                parts.append('<tr>' + ''.join(tds) + '</tr>')
            if area and sec is area.get('section'):
                vals = [r.value for ri, r in enumerate(sec.rows)
                        if ri in area.get('rows', ()) and not getattr(r, 'no_count', False)
                        and r.value is not None]
                t = sum(float(v) for v in vals) if vals else None
            else:
                t = sec.total() if not area else None
            if t is not None:
                # 小计值放在"工程量"列，其余列留空
                vidx = ncol - 1
                for _i, _c in enumerate(cols):
                    if _c['key'] == 'value':
                        vidx = _i
                        break
                tail = ''.join('<td></td>' for _ in range(ncol - vidx - 1))
                parts.append(
                    '<tr style="background:#e2efda">'
                    f'<td colspan="{max(1, vidx)}" align="right"><b>小计</b></td>'
                    f'<td width="{pct[vidx]:.1f}%"><b>{fmt_num(t, section_precision(sec, self.precision))}</b></td>'
                    + tail + '</tr>')
            for c in sec.children:
                walk(c, level + 1)
        for book in (self.project.books or [self.project.root]):
            walk(book, 0)
        parts.append('</table>')
        # 页脚统一由 ReportRenderer 在每页底部绘制（左/中/右 + 自动页码），这里不再插入
        return ''.join(parts)

    def report_html(self, with_margins=True):
        """生成报表 HTML（整份 = 头部 + 正文）。with_margins=False 用于真实打印
        （页边距交给打印机）；打印时按页拆分，见 report_head_html / report_body_html。"""
        rep = self.project.report or {}
        margins = rep.get('margins') or {}
        m0 = float(rep.get('margin', 1.9))
        mt = float(margins.get('top', m0))
        mb = float(margins.get('bottom', m0))
        ml = float(margins.get('left', m0))
        mr = float(margins.get('right', m0))
        inner = self.report_head_html() + self.report_body_html()
        if with_margins:
            return f'<div style="padding:{mt}cm {mr}cm {mb}cm {ml}cm;">{inner}</div>'
        return inner

    def selected_print_area(self):
        """把表格当前选中区域转成打印区域（行索引 + 列 key + 所属分部）。"""
        sel = self.table.selectionModel()
        indexes = sel.selectedIndexes() if sel is not None else []
        if not indexes:
            r, c = self.table.currentRow(), self.table.currentColumn()
            if r < 0 or c < 0:
                return None
            indexes = [self.table.model().index(r, c)]
        rows = sorted({i.row() for i in indexes if i.row() >= 0})
        keys = []
        for c in sorted({i.column() for i in indexes if i.column() >= 0}):
            key = self.col_key(c)
            if key and key not in keys:
                keys.append(key)
        if not rows or not keys:
            return None
        return {'section': self.current_section, 'rows': set(rows), 'keys': keys}

    def set_print_area(self):
        """把当前选中单元格设为打印区域，并打开预览确认。"""
        area = self.selected_print_area()
        if area is None:
            QMessageBox.information(self, APP_NAME, '请先在表格里选中要打印的行 / 列区域。')
            return
        self._print_area = area
        self.print_preview()
        self.status.showMessage('已设置打印区域：%d 行 × %d 列（点“清除打印区域”可恢复全部）'
                                % (len(area['rows']), len(area['keys'])), 5000)

    def clear_print_area(self):
        """清除打印区域，恢复打印整张计算书。"""
        self._print_area = None
        if hasattr(self, 'preview_pane'):
            self.preview_pane.refresh()
        self.status.showMessage('已清除打印区域，将打印整张计算书', 3000)

    def print_preview(self):
        """切到报表页并刷新纸张预览（页边距可直接拖动调整）。"""
        self.preview_pane.refresh()
        self.tabs.setCurrentIndex(1)

    def make_printer(self):
        """按当前报表设置创建 QPrinter（纸张大小/方向/页边距已应用）。"""
        from PyQt5.QtPrintSupport import QPrinter
        from PyQt5.QtGui import QPageLayout, QPageSize
        rep = self.project.report or {}
        printer = QPrinter(QPrinter.HighResolution)
        printer.setDocName('%s 工程量计算书' % self.project.name)
        try:
            # 纸张：取页面设置里选定的尺寸（mm），交给打印机驱动匹配最接近的纸型
            paper = paper_of(rep)
            printer.setPageSize(QPageSize(QSizeF(paper['w'], paper['h']),
                                          QPageSize.Millimeter, paper['name']))
            landscape = rep.get('orientation') == 'landscape'
            printer.setPageOrientation(QPageLayout.Landscape if landscape else QPageLayout.Portrait)
            margins = rep.get('margins') or {}
            m0 = float(rep.get('margin', 1.9))
            # PyQt5 的 setPageMargins 是 (左, 上, 右, 下, 单位) 五参数版
            printer.setPageMargins(float(margins.get('left', m0)) * 10.0,
                                   float(margins.get('top', m0)) * 10.0,
                                   float(margins.get('right', m0)) * 10.0,
                                   float(margins.get('bottom', m0)) * 10.0,
                                   QPrinter.Millimeter)
        except Exception:
            pass    # 个别驱动不支持以上设置时用驱动默认值
        return printer

    def print_report(self):
        """连系统打印机真实打印（Ctrl+P）：每一页都带表头。"""
        try:
            from PyQt5.QtPrintSupport import QPrintDialog
            from PyQt5.QtGui import QPainter
        except Exception as e:      # 缺少打印模块时给出提示
            QMessageBox.warning(self, APP_NAME, '当前环境缺少打印组件：%s' % e)
            return
        printer = self.make_printer()
        dlg = QPrintDialog(printer, self)
        dlg.setWindowTitle('打印工程量计算书')
        if dlg.exec_() != QDialog.Accepted:
            return
        try:
            pages = self.paint_report_to_printer(printer)
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, '打印失败：%s' % e)
            return
        self.status.showMessage('已发送到打印机：%s（共 %d 页，每页都带表头）'
                                % (printer.printerName(), pages), 5000)

    def paint_report_to_printer(self, printer):
        """逐页绘制到打印机：每页先画表头，再画该页正文。

        Qt 富文本不支持跨页重复表头（<thead> 只在表格顶部出现），
        所以这里把表头单独做成一个文档，按页切片自己画 —— 打印与预览完全一致。

        Qt 富文本按“屏幕逻辑 DPI”排 px（本机 150% 缩放时为 144dpi），
        所以纸张可打印区要按 打印机DPI/屏幕DPI 折算成文档 px 宽度，
        否则字号会被放大 1.5 倍、表格列宽不够导致编号折行。
        """
        from PyQt5.QtGui import QPainter
        from ui_printpreview import ReportRenderer
        scale = max(1.0, printer.resolution() / max(72.0, float(screen_dpi())))
        rect = printer.pageRect()
        cw = rect.width() / scale
        ch = rect.height() / scale
        fl, fc, fr = self.footer_fields()
        renderer = ReportRenderer(self.report_head_html(), self.report_body_html(), cw, ch,
                                  fl, fc, fr, show_page_no=True)
        painter = QPainter(printer)
        try:
            painter.scale(scale, scale)
            for i in range(renderer.page_count):
                if i:
                    printer.newPage()
                renderer.draw_page(painter, i, 0.0, 0.0)
        finally:
            painter.end()
        return renderer.page_count
        self.status.showMessage('已发送到打印机：%s' % printer.printerName(), 5000)

    def insert_symbol(self, ch):
        """把特殊字符插入到当前输入位置（内联编辑器 / 公式辅助框 / 单元格）。"""
        self._insert_formula_text(ch, 0)

    def insert_function(self, text, back=0):
        """插入函数模板（如 sqrt()），back=插入后光标回退字符数。"""
        self._insert_formula_text(text, back)

    def _insert_formula_text(self, text, back=0):
        """按优先级插入文本：正在编辑的表格单元格 → 公式辅助框 → 当前单元格末尾。"""
        # 1) 表格里正在编辑的公式单元格
        ed = self._cell_editor
        if ed is not None and ed.isVisible():
            ed.insert(text)
            if back:
                ed.setCursorPosition(max(0, ed.cursorPosition() - back))
            ed.setFocus()
            return
        # 2) 公式辅助框（只读时尝试激活当前公式单元格）
        editor = self._ensure_formula_edit()
        if editor is not None:
            cur = editor.textCursor()
            cur.insertText(text)
            if back:
                cur.movePosition(QTextCursor.Left, QTextCursor.MoveAnchor, back)
            editor.setTextCursor(cur)
            editor.setFocus()
            return
        # 3) 当前单元格文本末尾
        row = self.table.currentRow()
        col = self.table.currentColumn()
        if row < 0 or col < 0:
            return
        item = self.table.item(row, col)
        if item is None or self.col_key(col) in ('value', 'no_count'):
            return
        item.setText((item.text() or '') + text)
        self.table.setCurrentCell(row, col)

    def _ensure_formula_edit(self):
        """返回可写的公式辅助框；只读时若当前是公式单元格则激活它。"""
        if not self.formula_help.isReadOnly():
            return self.formula_help
        row = self.table.currentRow()
        col = self.table.currentColumn()
        if col < 0 or self.col_key(col) != 'formula':
            return None
        if row < 0 or row >= len(self.current_section.rows):
            return None
        self._formula_row = row
        self.formula_help.setReadOnly(False)
        self.formula_help.setStyleSheet('background:#ffffff;')
        self.formula_help.setTextSilent(self.current_section.rows[row].formula)
        return self.formula_help

    def insert_bracket(self, mode):
        """F4-F7 括号插入：pre=公式前插(), cur=插入点插(), cur_sq=插入点插[], end_sq=末尾插[]

        在表格单元格里直接打字时（内联编辑器激活）也生效，插入到光标处；
        [] 内的内容按备注处理，不参与计算。
        """
        # 1) 正在编辑的表格单元格 → 直接改编辑器（textChanged 会同步辅助框与工程量）
        ed = self._cell_editor
        if ed is not None and ed.isVisible() and ed.hasFocus():
            text = ed.text()
            pos = ed.cursorPosition()
            if mode == 'pre':
                ed.setText('()' + text)
                ed.setCursorPosition(1)
            elif mode == 'cur':
                ed.insert('()')
                ed.setCursorPosition(pos + 1)
            elif mode == 'cur_sq':
                ed.insert('[]')
                ed.setCursorPosition(pos + 1)
            elif mode == 'end_sq':
                ed.setText(text + '[]')
                ed.setCursorPosition(len(text) + 1)
            ed.setFocus()
            return
        # 2) 公式辅助框
        editor = self._ensure_formula_edit()
        if editor is None:
            QMessageBox.information(self, APP_NAME, '请先选中"计算公式"列的单元格。')
            return
        cur = editor.textCursor()
        if mode == 'pre':
            editor.setTextCursor(QTextCursor(editor.document()))
            cur = editor.textCursor()
            cur.insertText('()')
            cur.movePosition(QTextCursor.Left, QTextCursor.MoveAnchor, 1)
        elif mode == 'cur':
            pos = cur.position()
            cur.insertText('()')
            cur.setPosition(pos + 1)
        elif mode == 'cur_sq':
            pos = cur.position()
            cur.insertText('[]')
            cur.setPosition(pos + 1)
        elif mode == 'end_sq':
            cur.movePosition(QTextCursor.End)
            cur.insertText('[]')
            cur.movePosition(QTextCursor.Left, QTextCursor.MoveAnchor, 1)
        editor.setTextCursor(cur)
        editor.setFocus()

    # ---------------- 撤销 / 恢复 ----------------
    # 约定：栈里存“改动之前”的状态，撤销时恢复它、并把当前状态压入恢复栈。
    # 公式是边打字边写进模型的，所以连续编辑用 begin_edit()/commit_edit()
    # 只记录一次改动前快照，否则每次按键都算一步、撤销等于没反应。
    def snapshot_state(self):
        import copy
        return copy.deepcopy(self.project.to_dict())

    @staticmethod
    def _state_sig(data):
        try:
            return json.dumps(data, sort_keys=True, ensure_ascii=False)
        except Exception:
            return repr(id(data))

    def push_undo(self, label='编辑', data=None):
        """把改动前的快照压入撤销栈（必须在真正修改模型之前调用）。"""
        if data is None:
            data = self.snapshot_state()
        sig = self._state_sig(data)
        if self._undo_stack and self._undo_stack[-1][2] == sig:
            return False            # 状态没变化，不产生空的撤销步骤
        self._undo_stack.append((label or '编辑', data, sig))
        if len(self._undo_stack) > 60:
            self._undo_stack.pop(0)
        self._redo_stack.clear()
        return True

    def begin_edit(self, label='编辑'):
        """开始一段连续编辑（如逐字输入公式）：只在第一次记快照。"""
        if self._pending_undo is None:
            self._pending_undo = self.snapshot_state()
            self._pending_label = label

    def commit_edit(self):
        """连续编辑结束：把改动前快照正式压入撤销栈（无实际改动则跳过）。"""
        if self._pending_undo is None:
            return
        data, label = self._pending_undo, self._pending_label
        self._pending_undo = None
        self._pending_label = ''
        if self._state_sig(data) == self._state_sig(self.snapshot_state()):
            return
        self.push_undo(label or '编辑', data)

    def push_undo_coalesce(self, label, gap=1.2):
        """把短时间内的连续同类改动合并成一步（常量框逐字输入用）。"""
        now = time.time()
        if (self._undo_stack and self._undo_stack[-1][0] == label
                and now - getattr(self, '_last_undo_t', 0.0) < gap):
            self._last_undo_t = now
            return False
        ok = self.push_undo(label)
        self._last_undo_t = now
        return ok

    def undo(self):
        self.commit_edit()
        if not self._undo_stack:
            self.status.showMessage('没有可撤销的操作', 2000)
            return
        cur = self.snapshot_state()
        label, data, _sig = self._undo_stack.pop()
        self._redo_stack.append((label, cur, self._state_sig(cur)))
        self._load_project_data(data)
        self._dirty = True
        self.status.showMessage('已撤销：' + label, 2500)

    def redo(self):
        self.commit_edit()
        if not self._redo_stack:
            self.status.showMessage('没有可恢复的操作', 2000)
            return
        cur = self.snapshot_state()
        label, data, _sig = self._redo_stack.pop()
        self._undo_stack.append((label, cur, self._state_sig(cur)))
        self._load_project_data(data)
        self._dirty = True
        self.status.showMessage('已恢复：' + label, 2500)

    def _load_project_data(self, data):
        from core.model import Project as PJ
        self.project = PJ.from_dict(data)
        # 保留当前选中分部（按 id 找回）
        sid = getattr(self.current_section, 'id', None)
        found = self.find_section(sid) if sid else None
        self.current_section = found or self.project.root
        self.refresh_tree()
        if found:
            self.select_section(found)
        else:
            self.refresh_table()
        self.update_constants_display()

    # ---------------- 自动保存 ----------------
    def reset_autosave_timer(self):
        # 主文件已保存 或 设置了备份目录，均可启用自动保存
        if self.autosave_enabled and (self.current_file or self.autosave_dir):
            self.autosave_timer.start(max(1, self.autosave_min) * 60 * 1000)
        else:
            self.autosave_timer.stop()

    def autosave_tick(self):
        if not self._dirty:
            return
        now = datetime.datetime.now().strftime('%Y%m%d-%H%M%S')
        saved = []
        try:
            if self.current_file:
                self.project.save(self.current_file)
                saved.append('主文件')
        except Exception as e:
            self.status.showMessage('自动保存主文件失败：%s' % e, 4000)
        # 备份目录：另存带时间戳副本，防止主文件被覆盖后无法回溯
        if self.autosave_dir and os.path.isdir(self.autosave_dir):
            try:
                base = self.project.name or '工程'
                bpath = os.path.join(self.autosave_dir, f'{base}-自动备份-{now}{FILE_EXT}')
                self.project.save(bpath)
                saved.append('备份')
            except Exception as e:
                self.status.showMessage('自动备份失败：%s' % e, 4000)
        if saved:
            self._dirty = False
            self.status.showMessage('自动保存完成[%s] %s' %
                                    ('+'.join(saved), datetime.datetime.now().strftime('%H:%M:%S')), 3000)

    def autosave_settings(self):
        dlg = QDialog(self)
        dlg.setWindowTitle('自动保存设置')
        dlg.setMinimumWidth(480)
        form = QFormLayout(dlg)
        cb = QComboBox()
        cb.addItems(['关闭', '开启'])
        cb.setCurrentIndex(1 if self.autosave_enabled else 0)
        form.addRow('自动保存：', cb)
        sp = QDoubleSpinBox()
        sp.setRange(1, 120)
        sp.setDecimals(0)
        sp.setSingleStep(1)
        sp.setValue(self.autosave_min)
        sp.setSuffix(' 分钟')
        form.addRow('保存间隔：', sp)
        # 备份目录
        dir_row = QHBoxLayout()
        self.ed_autosave_dir = QLineEdit(self.autosave_dir)
        self.ed_autosave_dir.setPlaceholderText('留空则只覆盖主文件')
        btn_browse = QPushButton('浏览…')
        def browse():
            d = QFileDialog.getExistingDirectory(dlg, '选择自动保存目录',
                                                  self.ed_autosave_dir.text() or '')
            if d:
                self.ed_autosave_dir.setText(d)
        btn_browse.clicked.connect(browse)
        dir_row.addWidget(self.ed_autosave_dir, 1)
        dir_row.addWidget(btn_browse)
        form.addRow('备份目录：', dir_row)
        lbl = QLabel('说明：到点时若工程有改动，自动写回主文件；若指定了备份目录，'
                     '会同时在该目录生成带时间戳的备份副本，便于回溯。')
        lbl.setWordWrap(True)
        lbl.setStyleSheet('color:#666;')
        form.addRow(lbl)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        form.addRow(bb)
        if dlg.exec_() == QDialog.Accepted:
            self.autosave_enabled = (cb.currentIndex() == 1)
            self.autosave_min = int(sp.value())
            self.autosave_dir = self.ed_autosave_dir.text().strip()
            self.settings.setValue('autosave_enabled', self.autosave_enabled)
            self.settings.setValue('autosave_min', self.autosave_min)
            self.settings.setValue('autosave_dir', self.autosave_dir)
            self.reset_autosave_timer()
            state = '已开启，每 %d 分钟保存一次' % self.autosave_min if self.autosave_enabled else '已关闭'
            if self.autosave_dir:
                state += '，备份目录：' + self.autosave_dir
            self.status.showMessage('自动保存' + state, 4000)

    # ---------------- 其他功能 ----------------
    def view_formula(self):
        """查看当前分部所有公式。"""
        rows = self.current_section.rows
        if not rows:
            QMessageBox.information(self, APP_NAME, '当前分部没有公式。')
            return
        text = '\n'.join('%s、%s：%s' % (r.code, r.name, r.formula) for r in rows)
        dlg = QDialog(self)
        dlg.setWindowTitle('查看公式 - ' + self.current_section.name)
        dlg.resize(760, 520)
        lay = QVBoxLayout(dlg)
        ed = QTextEdit()
        ed.setReadOnly(True)
        ed.setPlainText(text)
        lay.addWidget(ed)
        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(dlg.reject)
        bb.accepted.connect(dlg.accept)
        lay.addWidget(bb)
        dlg.exec_()

    def data_summary(self):
        """分类汇总：打开右侧汇总面板（按编号 / 项目名称 / 部位 / 分部 / 单位）。"""
        if self._sum_folded:
            self.fold_summary(False)
        self.tabs.setCurrentIndex(0)
        self.summary_panel.refresh()
        n = len(self.summary_panel._groups)
        if n == 0:
            self.status.showMessage('当前范围没有可汇总的数据（检查是否都勾了“不计”）', 5000)
        else:
            self.status.showMessage('分类汇总已更新：%d 项（可切换汇总方式，或点“导出”存 Excel）'
                                    % n, 6000)

    def schedule_summary_refresh(self):
        """数据变化后延时刷新右侧汇总（面板收起时不刷）。"""
        if not hasattr(self, 'summary_panel') or getattr(self, '_sum_folded', False):
            return
        if getattr(self, '_sum_timer', None) is None:
            self._sum_timer = QTimer(self)
            self._sum_timer.setSingleShot(True)
            self._sum_timer.timeout.connect(self._refresh_summary_now)
        self._sum_timer.start(280)

    def _refresh_summary_now(self):
        try:
            self.summary_panel.refresh()
        except Exception:
            pass

    def export_number_format(self):
        """导出纯数字格式（每行工程量，便于复制到其他软件）。"""
        self.project.recalc()
        lines = []
        for sec in self.project.all_sections():
            for r in sec.rows:
                if not r.no_count and r.value is not None:
                    lines.append(str(round(r.value, precision_for_unit(r.unit, self.precision))))
        if not lines:
            QMessageBox.information(self, APP_NAME, '暂无可导出的数值。')
            return
        from PyQt5.QtWidgets import QApplication as QA
        QA.clipboard().setText('\n'.join(lines))
        self.status.showMessage('已复制 %d 个数字到剪贴板' % len(lines), 4000)

    def show_number_format(self):
        QMessageBox.information(self, APP_NAME,
                                '数字格式规则（V2.2 起按单位自动决定）：\n\n'
                                '  m、m2、m3、kg     → 保留 2 位小数\n'
                                '  t（吨）           → 保留 3 位小数\n'
                                '  个、块、套、处、台、组 → 取整\n'
                                '  其他单位           → 默认 3 位小数\n\n'
                                '表格、打印预览与导出的 Excel 均按此规则显示。')

    def rmb_convert(self):
        """把选中行的工程量转换为人民币大写，写入备注。"""
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()})
        if not rows:
            QMessageBox.information(self, APP_NAME, '请先选中要转换的行。')
            return
        def to_rmb(n):
            units = '零壹贰叁肆伍陆柒捌玖'
            parts = '元拾佰仟万拾佰仟亿'
            n = round(n, 2)
            integer = int(n)
            decimal = int(round((n - integer) * 100))
            if integer == 0:
                s = '零'
            else:
                digits = str(integer)
                s = ''
                for i, d in enumerate(digits):
                    s += units[int(d)] + parts[len(digits) - 1 - i]
            if decimal:
                s += '点' + units[decimal // 10] + units[decimal % 10]
            else:
                s += '整'
            return s
        cnt = 0
        self.push_undo('人民币大写')
        for r in rows:
            if r < len(self.current_section.rows):
                row = self.current_section.rows[r]
                if row.value is not None:
                    row.remark = (row.remark + ' ' if row.remark else '') + to_rmb(row.value)
                    cnt += 1
        self.refresh_table()
        self._dirty = True
        self.status.showMessage('已为 %d 行追加人民币大写' % cnt, 4000)

    def setup_file_assoc(self):
        """把 .gcls 工程文件关联到本程序（资源管理器显示本软件图标、双击即打开）。"""
        if not current_exe():
            QMessageBox.information(self, '工程文件图标',
                                    '当前是以源码方式运行，文件关联只在打包后的 exe 中生效。\n\n'
                                    '打包成 exe 后，首次运行会自动关联；换过存放目录后'
                                    '再点一次此按钮即可重新指向。')
            return
        ok, info = register_file_assoc()
        if ok:
            QMessageBox.information(
                self, '工程文件图标',
                '已设置完成。\n\n工程文件（*%s）将显示为本软件图标，双击即用本程序打开。\n\n'
                '当前程序：%s\n\n'
                '（若资源管理器里图标还是旧的，按 F5 刷新一下，或注销后重新登录。）'
                % (FILE_EXT, info))
        else:
            QMessageBox.warning(self, '设置失败', '未能写入文件关联：\n%s' % info)

    def show_about(self):
        QMessageBox.about(self, APP_NAME,
                          '<h2>工程量计算书 V%s</h2>' % APP_VER +
                          '<p>一款手算工程量计算书工具，支持手写公式、备注、报表与导出 WPS 表格。</p>'
                          '<p>功能：多个单项工程、可自定义层级（单位/分部/分项工程）、'
                          '公式自动计算与校核、公式与辅助框同色对照、钢筋与工程符号、'
                          '单位下拉与按单位自动定小数位、自定义常量（三列，含说明）、'
                          '按编号/项目名称/部位/分部/单位分类汇总并可导出、'
                          '行/列/单元格填充色、默认自动行高列宽、撤销与恢复、查找定位、'
                          '面板可收起、删除全部数据、'
                          '纸张大小可选（跟随打印机支持的纸型）、报表打印预览与直接打印、'
                          'Excel/WPS 导出、自动保存与备份。</p>')

    def show_formula_help(self):
        self.tabs.setCurrentIndex(2)

    def show_functions(self):
        """“插入函数”旧入口，保留为打开函数面板。"""
        self.show_function_picker()

    def _place_popup(self, dlg, btn=None):
        """把弹出面板放到按钮下方；下方空间不够则放到上方，并保证不出屏幕。"""
        if btn is not None:
            pos = btn.mapToGlobal(QPoint(0, btn.height()))
            scr = QApplication.screenAt(btn.mapToGlobal(QPoint(0, 0))) or QApplication.primaryScreen()
            screen = scr.availableGeometry()
            x = min(pos.x(), screen.right() - dlg.width() - 8)
            y = pos.y() if pos.y() + dlg.height() < screen.bottom() else pos.y() - btn.height() - dlg.height()
            dlg.move(max(screen.left(), x), max(screen.top(), y))
        else:
            c = self.mapToGlobal(self.rect().center())
            dlg.move(max(0, c.x() - dlg.width() // 2), max(0, c.y() - dlg.height() // 2))

    def show_function_picker(self):
        """弹出函数面板：点选 abs/sqrt/round/power/sin/max… 直接插入公式。"""
        s = self.sender()
        btn = s if isinstance(s, QWidget) else None
        dlg = FunctionPicker(self.insert_function, self)
        dlg.adjustSize()
        self._place_popup(dlg, btn)
        dlg.show()
        self._func_picker = dlg

    def closeEvent(self, event):
        # 测试模式（CALCBOOK_NO_RESTORE=1）：不写回界面设置、不弹保存提示，
        # 免得自动化测试把用户的窗口尺寸 / 面板宽度 / 自适应开关改掉
        test_mode = bool(os.environ.get('CALCBOOK_NO_RESTORE'))
        if not test_mode:
            self.settings.setValue('precision', self.precision)
            self.settings.setValue('row_autofit', bool(self._row_autofit))
            self.settings.setValue('col_autofit', bool(self._col_autofit))
            if hasattr(self, 'left_split'):
                self.settings.setValue('left_split_sizes', self.left_split.sizes())
                if not getattr(self, '_const_folded', False):
                    self.settings.setValue('const_h', self.const_box.height())
            if hasattr(self, 'main_split'):
                self.settings.setValue('main_left_w', self._left_width)
                self.settings.setValue('main_sum_w', self._sum_width)
            if getattr(self, 'current_section', None) is not None:
                self.settings.setValue('last_section', self.current_section.id)
        if self._dirty and self.current_file and not test_mode:
            ret = QMessageBox.question(self, APP_NAME, '当前工程有未保存的修改，是否保存？',
                                       QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if ret == QMessageBox.Cancel:
                event.ignore()
                return
            if ret == QMessageBox.Yes:
                self.save_project()
        event.accept()

    def open_project(self, path=None):
        """打开工程（覆盖原定义以补充自动保存重置）。"""
        if not path:
            path, _ = QFileDialog.getOpenFileName(
                self, '打开工程', '',
                '工程量计算书工程 (*%s *.json);;工程文件 (*%s);;所有文件 (*)' % (FILE_EXT, FILE_EXT))
        if not path:
            return
        try:
            self.project = Project.load(path)
            if (self.project.name or '').strip() in ('', '未命名工程'):
                self.project.name = self.name_from_path(path)
            self.current_file = path
            self.current_section = self.project.root
            # 根节点常常是空壳（数据在下级分部/分项），自动跳到内容最多的节点，
            # 否则一打开工作区一片空白，容易被当成“数据丢了”
            self.current_section = (self._pick_section_with_rows(None)
                                   or self.project.root)
            self._dirty = False
            self._undo_stack.clear()
            self._redo_stack.clear()
            self._pending_undo = None
            self.update_title()
            self.refresh_tree()
            self.refresh_table()
            # 工程文件里存的列宽常常偏大（上次是别的窗口宽度下排的），
            # 自动列宽开启时重新排一次，直接铺满当前工作区，不留横向滚动条
            if getattr(self, '_col_autofit', False):
                QTimer.singleShot(0, self._safe_auto_size)
            self.reset_autosave_timer()
            self.remember_file(path)
            self.add_recent(path)
            self.status.showMessage(f'已打开 {path}', 4000)
        except Exception as e:
            QMessageBox.critical(self, '打开失败', str(e))


def main():
    # 与 Delphi/VCL 程序行为一致：不做 Qt 高 DPI 自动缩放，
    # 由 Windows 统一做位图虚拟化，保证在不同缩放比例的屏幕上控件尺寸一致，
    # 不会出现"图标文字随屏幕变化、忽大忽小"的问题。

    # 全局异常钩子：打包成 exe 后没有控制台，未处理异常写日志文件方便排查
    def _excepthook(tp, val, tb):
        try:
            with open(LOG_FILE, 'a', encoding='utf-8') as f:
                f.write('\n[%s] 未处理异常\n' % time.strftime('%Y-%m-%d %H:%M:%S'))
                traceback.print_exception(tp, val, tb, file=f)
        except Exception:
            pass
        try:
            sys.__excepthook__(tp, val, tb)
        except Exception:
            pass
    sys.excepthook = _excepthook

    app = QApplication(sys.argv)
    app.setFont(QFont('微软雅黑', 10))
    # 所有弹窗的 OK / Cancel 等按钮统一显示中文
    app._calcbook_i18n_filter = DialogButtonI18nFilter(app)
    app.installEventFilter(app._calcbook_i18n_filter)
    # 任务栏使用本程序自己的图标与分组（否则会跟 python.exe 混在一起）
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            'CalcBook.Software.%s' % APP_VER)
    except Exception:
        pass
    # 双击工程文件启动时，Windows 会把文件路径作为第一个参数传进来
    arg_file = ''
    for a in sys.argv[1:]:
        if a and not a.startswith('-') and a.lower().endswith(FILE_EXT):
            arg_file = a
            break
    w = MainWindow(app)
    if arg_file and os.path.isfile(arg_file):
        w.open_project(arg_file)
    w.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
