# -*- coding: utf-8 -*-
"""查看报表窗口。

与“打印预览”不同，这里先给用户一份**可读、可核对**的表格化报表：
左边是分部/分项的层级，右边是各列数据，每个节点后面带合计。
确认没问题再点“打印”或“导出表格”。

V3.7.1 起：
- 报表范围可选「哪一本计算书」「哪一层级」；
- 打开时窗口自动放大到主界面编辑区大小，不用缩在角落里眯眼看；
- 每个单元格画浅灰虚线框，格子边界清楚；
- 所有列（含项目名称列）都能拖动改宽度。
"""

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor, QFont, QPen
from PyQt5.QtWidgets import (
    QAbstractItemView, QComboBox, QDialog, QDialogButtonBox, QHBoxLayout,
    QHeaderView, QLabel, QMenu, QPushButton, QStyledItemDelegate,
    QTreeWidget, QTreeWidgetItem, QVBoxLayout
)

from core.units import precision_for_unit, section_precision, fmt_num


LEVEL_COLORS = ['#1f3864', '#1f6f4a', '#9c4a1a', '#5b2a86', '#7a5c00', '#215b6b']


def _row_field(row, key):
    """取行上某个字段的文本（内置字段在属性上，自定义列在 custom 里）。"""
    if hasattr(row, key):
        return getattr(row, key) or ''
    return (getattr(row, 'custom', {}) or {}).get(key, '') or ''


class DashedCellDelegate(QStyledItemDelegate):
    """给每个单元格画一圈浅灰虚线框：一眼看清行列边界，不刺眼。"""

    def paint(self, painter, option, index):
        super().paint(painter, option, index)
        painter.save()
        pen = QPen(QColor('#c9ced6'))
        pen.setStyle(Qt.DashLine)
        pen.setWidth(1)
        painter.setPen(pen)
        r = option.rect
        painter.drawRect(r.left(), r.top(), max(1, r.width() - 1), max(1, r.height() - 1))
        painter.restore()


class ReportViewerDialog(QDialog):
    """报表查看：树形列出整个工程的工程量数据，并给出各级合计。"""

    COLS = ['编号', '项目名称', '单位', '工程量', '计算公式', '图纸编号', '备注']

    def __init__(self, win, parent=None):
        super().__init__(parent)
        self.win = win
        self.setWindowTitle('查看报表')
        self.setMinimumSize(960, 600)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(6)

        # —— 顶部：范围选择 + 提示 ——
        top = QHBoxLayout()
        top.addWidget(QLabel('计算书：'))
        self.cb_book = QComboBox()
        self.cb_book.setMinimumWidth(160)
        top.addWidget(self.cb_book)
        top.addWidget(QLabel('层级：'))
        self.cb_level = QComboBox()
        self.cb_level.setMinimumWidth(150)
        top.addWidget(self.cb_level)
        top.addSpacing(10)
        self.lb_range = QLabel('')
        self.lb_range.setStyleSheet('color:#444;')
        top.addWidget(self.lb_range)
        top.addStretch(1)
        self.lb_total = QLabel('')
        self.lb_total.setStyleSheet('color:#1f5f99; font-weight:600;')
        top.addWidget(self.lb_total)
        lay.addLayout(top)

        # —— 表格 ——
        self.tree = QTreeWidget()
        self.tree.setColumnCount(len(self.COLS))
        self.tree.setHeaderLabels(self.COLS)
        self.tree.setAlternatingRowColors(True)
        self.tree.setRootIsDecorated(True)
        self.tree.setUniformRowHeights(False)
        self.tree.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tree.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.setItemDelegate(DashedCellDelegate(self.tree))
        hh = self.tree.header()
        # 所有列都能拖动宽度（原来“项目名称”列是 Stretch，拖不动）
        for i in range(len(self.COLS)):
            hh.setSectionResizeMode(i, QHeaderView.Interactive)
        hh.setStretchLastSection(True)
        self.tree.setColumnWidth(0, 110)
        self.tree.setColumnWidth(1, 260)
        self.tree.setColumnWidth(2, 70)
        self.tree.setColumnWidth(3, 110)
        self.tree.setColumnWidth(4, 320)
        self.tree.setColumnWidth(5, 110)
        self.tree.setColumnWidth(6, 160)
        lay.addWidget(self.tree, 1)

        # —— 底部按钮 ——
        bar = QHBoxLayout()
        self.btn_refresh = QPushButton('刷新')
        self.btn_refresh.clicked.connect(self.refresh)
        self.btn_print = QPushButton('打印…')
        self.btn_print.clicked.connect(self._on_print)
        self.btn_export = QPushButton('导出表格…')
        self.btn_export.clicked.connect(self._on_export)
        self.btn_page = QPushButton('页面设置…')
        self.btn_page.clicked.connect(self._on_page_setup)
        bar.addWidget(self.btn_refresh)
        bar.addWidget(self.btn_print)
        bar.addWidget(self.btn_export)
        bar.addWidget(self.btn_page)
        bar.addStretch(1)
        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.button(QDialogButtonBox.Close).setText('关闭')
        bb.rejected.connect(self.reject)
        bar.addWidget(bb)
        lay.addLayout(bar)

        self._loading = False
        self.cb_book.currentIndexChanged.connect(self._on_scope_changed)
        self.cb_level.currentIndexChanged.connect(self._on_scope_changed)
        self._fill_scope_boxes()
        self.refresh()
        self._fit_to_editor()

    # ---------------- 范围选择 ----------------
    def _fill_scope_boxes(self):
        """把「有哪些计算书」「有哪些层级」填进下拉框。

        刷新时保留用户已经选好的项（按序号还原），免得每次刷新都跳回“全部”。
        """
        keep_book = self.cb_book.currentIndex()
        keep_level = self.cb_level.currentIndex()
        self._loading = True
        proj = self.win.project
        self.cb_book.clear()
        self.cb_book.addItem('全部计算书', None)
        for b in (proj.books or []):
            self.cb_book.addItem(b.name, b.id)
        self.cb_level.clear()
        self.cb_level.addItem('全部层级', -1)
        levels = list(proj.levels or [])
        maxd = self._max_depth(proj.books or [])
        for d in range(0, maxd + 1):
            name = levels[min(d, len(levels) - 1)] if levels else ('第%d层' % (d + 1))
            self.cb_level.addItem('只看「%s」层' % name, d)
        if keep_book >= 0 and keep_book < self.cb_book.count():
            self.cb_book.setCurrentIndex(keep_book)
        if keep_level >= 0 and keep_level < self.cb_level.count():
            self.cb_level.setCurrentIndex(keep_level)
        self._loading = False

    @staticmethod
    def _max_depth(books):
        m = 0

        def walk(sec, d):
            nonlocal m
            m = max(m, d)
            for c in getattr(sec, 'children', []) or []:
                walk(c, d + 1)
        for b in books:
            walk(b, 0)
        return m

    def _on_scope_changed(self):
        if not self._loading:
            self.refresh()

    def _selected_books(self):
        """按下拉框选择取出要统计的计算书列表。"""
        proj = self.win.project
        books = proj.books or []
        bid = self.cb_book.currentData()
        if bid is None:
            return books
        for b in books:
            if b.id == bid:
                return [b]
        return books

    # ---------------- 窗口大小 ----------------
    def _fit_to_editor(self):
        """打开时自动放大到主界面编辑区（表格）的大小，看着不费劲。"""
        try:
            tw = getattr(self.win, 'table', None)
            if tw is None:
                return
            w = int(tw.width())
            h = int(tw.height())
        except Exception:
            return
        self.resize(max(960, w + 60), max(600, h + 80))

    # ---------------- 数据 ----------------
    def refresh(self):
        win = self.win
        try:
            win.project.recalc()
        except Exception:
            pass
        self._fill_scope_boxes()
        self.tree.clear()
        proj = win.project
        prec = getattr(win, 'precision', 3)
        books = self._selected_books()
        level = self.cb_level.currentData()
        if level is None:
            level = -1

        grand = 0.0
        grand_rows = 0
        for book in books:
            if level < 0:
                item = self._add_node(book, 0, proj, prec)
                self.tree.addTopLevelItem(item)
                item.setExpanded(True)
            else:
                for node in self._nodes_at_depth(book, level):
                    item = self._add_level_node(node, level, proj, prec)
                    self.tree.addTopLevelItem(item)
            tot, cnt = self._node_total(book)
            grand += tot
            grand_rows += cnt

        scope_txt = '报表范围：%s' % (self.cb_book.currentText() or '全部计算书')
        if level >= 0:
            scope_txt += ' · %s' % (self.cb_level.currentText() or '')
        self.lb_range.setText(scope_txt)
        self.lb_total.setText('合计：%.3f    共 %d 条计算式'
                              % (grand, grand_rows))
        for i in range(self.tree.columnCount()):
            self.tree.resizeColumnToContents(i) if i == 0 else None

    def _nodes_at_depth(self, sec, depth, cur=0):
        """取某一本计算书里第 depth 层的全部节点。"""
        out = []
        if cur == depth:
            return [sec]
        for c in getattr(sec, 'children', []) or []:
            out.extend(self._nodes_at_depth(c, depth, cur + 1))
        return out

    def _node_total(self, sec):
        """节点及其所有后代的工程量合计与条数。"""
        total, cnt = 0.0, 0
        for r in getattr(sec, 'rows', []) or []:
            if not r.no_count and r.value is not None:
                total += r.value
                cnt += 1
        for c in getattr(sec, 'children', []) or []:
            t, n = self._node_total(c)
            total += t
            cnt += n
        return total, cnt

    def _add_level_node(self, sec, depth, proj, prec):
        """按层级查看时的一个节点：本级条数 + 含下级的合计。"""
        own = len(getattr(sec, 'rows', []) or [])
        tot, cnt = self._node_total(sec)
        kids = len(getattr(sec, 'children', []) or [])
        item = QTreeWidgetItem([sec.name, '', '', '', '', '', ''])
        color = QColor(LEVEL_COLORS[depth % len(LEVEL_COLORS)])
        f = item.font(0)
        f.setBold(True)
        item.setFont(0, f)
        item.setForeground(0, QBrush(color))
        item.setText(1, '本级 %d 条 / 含下级 %d 条' % (own, cnt))
        item.setText(3, fmt_num(tot, section_precision(sec, prec)) if cnt else '')
        item.setText(4, '含 %d 个下级节点' % kids if kids else '')
        item.setData(0, Qt.UserRole, sec.id)
        item.setData(1, Qt.UserRole, 'total')
        return item

    def _add_node(self, sec, depth, proj, prec):
        """把一个节点（及其子节点、计算式行）建成报表条目。"""
        label = sec.name
        tot, cnt = self._node_total(sec)
        item = QTreeWidgetItem([label, '', '', '', '', '', ''])
        item.setFirstColumnSpanned(False)
        color = QColor(LEVEL_COLORS[depth % len(LEVEL_COLORS)])
        f = item.font(0)
        f.setBold(True)
        item.setFont(0, f)
        item.setForeground(0, QBrush(color))
        item.setText(1, '（%d 条）' % cnt)
        item.setText(3, fmt_num(tot, section_precision(sec, prec)) if cnt else '')
        item.setData(0, Qt.UserRole, sec.id)
        item.setData(1, Qt.UserRole, 'total')

        for r in getattr(sec, 'rows', []) or []:
            val = '' if r.no_count else (
                fmt_num(r.value, precision_for_unit(r.unit, prec))
                if r.value is not None else '')
            sub = QTreeWidgetItem([r.code or '', r.name or '', r.unit or '', val,
                                   r.formula or '', _row_field(r, 'drawing_no'),
                                   r.remark or ''])
            for c in range(self.tree.columnCount()):
                sub.setForeground(c, QBrush(QColor('#333333')))
            item.addChild(sub)

        for c in getattr(sec, 'children', []) or []:
            item.addChild(self._add_node(c, depth + 1, proj, prec))
        return item

    # ---------------- 按钮回调 ----------------
    def _on_print(self):
        try:
            self.win.print_report()
        except Exception as e:
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.warning(self, '打印失败', str(e))

    def _on_export(self):
        try:
            self.win.export_excel()
        except Exception as e:
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.warning(self, '导出失败', str(e))

    def _on_page_setup(self):
        try:
            self.win.page_setup()
        except Exception as e:
            from PyQt5.QtWidgets import QMessageBox
            QMessageBox.warning(self, '页面设置失败', str(e))
