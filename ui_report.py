# -*- coding: utf-8 -*-
"""查看报表窗口。

与“打印预览”不同，这里先给用户一份**可读、可核对**的表格化报表：
左边是分部/分项的层级，右边是各列数据，每个节点后面带合计。
确认没问题再点“打印”或“导出表格”。
"""

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import (
    QAbstractItemView, QDialog, QDialogButtonBox, QHBoxLayout, QHeaderView,
    QLabel, QMenu, QPushButton, QTreeWidget, QTreeWidgetItem, QVBoxLayout
)

from core.units import precision_for_unit, section_precision, fmt_num


LEVEL_COLORS = ['#1f3864', '#1f6f4a', '#9c4a1a', '#5b2a86', '#7a5c00', '#215b6b']


def _row_field(row, key):
    """取行上某个字段的文本（内置字段在属性上，自定义列在 custom 里）。"""
    if hasattr(row, key):
        return getattr(row, key) or ''
    return (getattr(row, 'custom', {}) or {}).get(key, '') or ''


class ReportViewerDialog(QDialog):
    """报表查看：树形列出整个工程的工程量数据，并给出各级合计。"""

    COLS = ['编号', '项目名称', '单位', '工程量', '计算公式', '图纸编号', '备注']

    def __init__(self, win, parent=None):
        super().__init__(parent)
        self.win = win
        self.setWindowTitle('查看报表')
        self.setMinimumSize(960, 600)
        self.resize(1080, 680)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(6)

        # —— 顶部：范围 + 提示 ——
        top = QHBoxLayout()
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
        hh = self.tree.header()
        hh.setSectionResizeMode(0, QHeaderView.Interactive)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        for i in range(2, len(self.COLS)):
            hh.setSectionResizeMode(i, QHeaderView.Interactive)
        self.tree.setColumnWidth(0, 110)
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

        self.refresh()

    # ---------------- 数据 ----------------
    def refresh(self):
        win = self.win
        try:
            win.project.recalc()
        except Exception:
            pass
        self.tree.clear()
        proj = win.project
        prec = getattr(win, 'precision', 3)
        only_cur = False
        try:
            only_cur = bool(self.win.summary_scope_only_cur())
        except Exception:
            only_cur = False

        scope = None
        if only_cur:
            scope = getattr(win, 'current_section', None)

        books = proj.books if scope is None else [scope]
        grand = 0.0
        grand_rows = 0
        for book in books:
            item = self._add_node(book, 0, proj, prec)
            self.tree.addTopLevelItem(item)
            item.setExpanded(True)
            tot, cnt = self._node_total(book)
            grand += tot
            grand_rows += cnt

        self.lb_range.setText('报表范围：%s' % ('全部计算书' if scope is None else scope.name))
        self.lb_total.setText('合计：%.3f    共 %d 条计算式'
                              % (grand, grand_rows))
        for i in range(self.tree.columnCount()):
            self.tree.resizeColumnToContents(i) if i == 0 else None

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
