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

V3.7.9 起：
- 增加**分类汇总**：按清单编号 / 项目名称 / 计量单位 / 备注 / 图纸编号 /
  自定义列分组，每组小计工程量与条数；
- 汇总结果可「打印汇总预览…」——按页面设置的纸张分页打印。

V3.7.10 起（修 bug）：
- **正则改成和「筛选」完全一样的口径**：正则命中才参与汇总，不命中的行
  直接排除（不再塞进「(未匹配)」组里造成显示混乱）；正则可指定作用列
  （默认跟随分类列），筛选列与分组列可以不同；
- 保留「按正则提取分组名」高级模式（原来的按匹配内容归组）；
- **全空行不再参与汇总**（默认 15 行空行以前会自成一组「(空) 共 15 条」，
  这是"显示混乱"的最大来源）；
- 分类汇总单独成表（分组 / 单位 / 条数 / 工程量合计），双击分组行看该组明细；
- 汇总打印预览窗口与报表窗口同尺寸，工具条按钮全部中文。

V3.7.11 起：
- **打印预览两种模式都能用**：选「不汇总」时预览的是报表明细（按当前计算书 /
  层级范围，带分部小计与总计），不再只能汇总才能预览。

V3.8 起：
- **预览窗的打印换成我们自己的打印窗口**：Qt 自带的打印按钮走它内部的打印框，
  页码范围既拿不到也没法预填（已隐藏），改由预览窗底部的「打印…」按钮调起
  我们的打印窗口 —— 同样支持页码范围，打印后状态栏说明本次打了第几页。
"""

import re as _re

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor, QFont, QPen
from PyQt5.QtWidgets import (
    QAbstractItemView, QAction, QComboBox, QDialog, QDialogButtonBox,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMenu, QMessageBox,
    QPushButton, QStyledItemDelegate, QTableWidget, QTableWidgetItem,
    QToolButton, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget
)

from core.units import precision_for_unit, section_precision, fmt_num
# V3.7.10：正则口径与「筛选」完全一致（同一套 norm_text + filter_match）
from ui_sortfilter import filter_match, norm_text


LEVEL_COLORS = ['#1f3864', '#1f6f4a', '#9c4a1a', '#5b2a86', '#7a5c00', '#215b6b']

# 打印预览工具条按钮中文化（Qt 未加载中文翻译时默认是英文）
_PREVIEW_ZH = {
    'Next page': '下一页', 'Previous page': '上一页',
    'First page': '首页', 'Last page': '末页',
    'Fit width': '按宽度显示', 'Fit page': '整页显示',
    'Zoom in': '放大', 'Zoom out': '缩小',
    'Portrait': '纵向', 'Landscape': '横向',
    'Show single page': '单页显示', 'Show facing pages': '双页显示',
    'Show overview of all pages': '总览所有页',
    'Print': '打印', 'Page setup': '页面设置',
}


def localize_print_preview(dlg):
    """把 QPrintPreviewDialog 的工具条按钮（Zoom in / Print…）改成中文。

    V3.8：Qt 自带的「打印」按钮弹的是它**内部**的打印框，页码范围既拿不到，
    也无法预填页面设置里的范围 —— 隐藏它，改由我们在预览窗底部提供
    自己的「打印…」（走统一的页码范围解析）。
    """
    try:
        # 先按原始英文找到「打印」按钮并隐藏（要在改名之前找）。
        # 注意：光调 setVisible(False) 不够 —— 窗口还没显示时隐藏的子控件，
        # 窗口显示后又会冒出来；工具条上的按钮跟随 QAction 的可见性，
        # 所以必须把 QAction 也隐藏掉。
        for b in dlg.findChildren(QToolButton):
            act = b.defaultAction()
            if act is not None and act.text() == 'Print':
                act.setVisible(False)
                b.setVisible(False)
        for act in dlg.findChildren(QAction):
            zh = _PREVIEW_ZH.get(act.text())
            if zh:
                act.setText(zh)
                act.setToolTip(zh)
        for b in dlg.findChildren(QToolButton):
            if b.isHidden():        # 自己被隐藏的按钮（Qt 的「打印」）不再改名
                continue
            zh = _PREVIEW_ZH.get(b.text())
            if zh:
                b.setText(zh)
                b.setToolTip(zh)
    except Exception:
        pass



def _row_field(row, key):
    """取行上某个字段的文本（内置字段在属性上，自定义列在 custom 里）。

    V3.9：备注列 / 自定义列里写了公式时，返回算出来的数值。
    """
    if key == 'img':
        return ''
    if key == 'value':
        return fmt_num(getattr(row, 'value', None),
                       precision_for_unit(getattr(row, 'unit', ''), 3))
    calc = (getattr(row, 'cell_calc', None) or {}).get(key)
    if calc:
        v = calc.get('value')
        if v is not None:
            return fmt_num(v, precision_for_unit(getattr(row, 'unit', ''), 3))
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
    # V3.7.10：分类汇总模式单独一套表头（干净的一张小计表）
    GROUP_COLS = ['分组', '单位', '条数', '工程量合计']

    def __init__(self, win, parent=None):
        super().__init__(parent)
        self.win = win
        self.setWindowTitle('查看报表')
        self.setMinimumSize(960, 600)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(6)

        # —— 顶部第一行：范围选择 ——
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

        # —— 顶部第二行：分类汇总（V3.7.9 新增，V3.7.10 修正则口径）——
        grp = QHBoxLayout()
        grp.addWidget(QLabel('分类汇总：'))
        self.cb_group = QComboBox()
        self.cb_group.addItem('不汇总', None)
        for _k, _t in (('code', '清单编号'), ('name', '项目名称'),
                       ('unit', '计量单位'), ('remark', '备注'),
                       ('drawing_no', '图纸编号')):
            self.cb_group.addItem(_t, _k)
        self.cb_group.setMinimumWidth(100)
        self.cb_group.currentIndexChanged.connect(self._on_scope_changed)
        grp.addWidget(self.cb_group)
        grp.addWidget(QLabel('正则列：'))
        self.cb_regex_col = QComboBox()
        self.cb_regex_col.addItem('自动（跟随分类列）', '')
        for _k, _t in (('code', '清单编号'), ('name', '项目名称'),
                       ('unit', '计量单位'), ('formula', '计算公式'),
                       ('remark', '备注')):
            self.cb_regex_col.addItem(_t, _k)
        self.cb_regex_col.setMinimumWidth(120)
        self.cb_regex_col.currentIndexChanged.connect(self._on_scope_changed)
        grp.addWidget(self.cb_regex_col)
        grp.addWidget(QLabel('正则：'))
        self.ed_regex = QLineEdit()
        self.ed_regex.setPlaceholderText('与「筛选」一致：命中的行才参与汇总')
        self.ed_regex.setMaximumWidth(240)
        self.ed_regex.setToolTip(
            '正则口径与表格右键「筛选」里的正则匹配完全一样：\n'
            '· 在「正则列」上匹配，命中的行才参与汇总，不命中的直接排除；\n'
            '· 留空则按分类列的值相同归组；\n'
            '· 例：^0304 → 清单编号 0304 开头的；门窗 → 名称含“门窗”的。')
        self.ed_regex.textChanged.connect(self._on_scope_changed)
        grp.addWidget(self.ed_regex)
        grp.addWidget(QLabel('方式：'))
        self.cb_regex_mode = QComboBox()
        self.cb_regex_mode.addItem('筛选匹配行（同筛选）', 'filter')
        self.cb_regex_mode.addItem('按正则提取分组名', 'extract')
        self.cb_regex_mode.setMinimumWidth(150)
        self.cb_regex_mode.currentIndexChanged.connect(self._on_scope_changed)
        grp.addWidget(self.cb_regex_mode)
        self.btn_print_sum = QPushButton('打印预览…')
        self.btn_print_sum.clicked.connect(self._on_print_preview)
        self.btn_print_sum.setToolTip('预览当前报表：\n'
                                      '· 选了「分类汇总」→ 预览汇总表；\n'
                                      '· 选「不汇总」→ 预览报表明细（按当前计算书 / 层级范围）。')
        grp.addWidget(self.btn_print_sum)
        grp.addStretch(1)
        lay.addLayout(grp)

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
        # V3.7.10：双击分组行看该组明细
        self.tree.itemDoubleClicked.connect(self._on_group_double_clicked)
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

    # ---------------- V3.7.9 分类汇总 ----------------
    def _refresh_group_box(self):
        """把自定义列也加进「分类汇总」下拉框（保留当前选择）。"""
        keep = self.cb_group.currentData()
        self._loading = True
        while self.cb_group.count() > 6:      # 前 6 项固定：不汇总 + 5 个内置
            self.cb_group.removeItem(self.cb_group.count() - 1)
        for cd in (self.win.project.columns or []):
            if cd.get('builtin'):
                continue
            self.cb_group.addItem(str(cd.get('title', '列')), str(cd.get('key')))
        idx = self.cb_group.findData(keep)
        if idx >= 0:
            self.cb_group.setCurrentIndex(idx)
        self._loading = False

    def _refresh_regex_col_box(self):
        """正则列下拉框：内置几列 + 自定义列（保留当前选择）。"""
        keep = self.cb_regex_col.currentData()
        self._loading = True
        while self.cb_regex_col.count() > 6:   # 自动 + 5 个内置
            self.cb_regex_col.removeItem(self.cb_regex_col.count() - 1)
        for cd in (self.win.project.columns or []):
            if cd.get('builtin'):
                continue
            self.cb_regex_col.addItem(str(cd.get('title', '列')),
                                      str(cd.get('key')))
        idx = self.cb_regex_col.findData(keep)
        if idx >= 0:
            self.cb_regex_col.setCurrentIndex(idx)
        self._loading = False

    def _scope_rows(self):
        """当前范围（计算书 × 层级）内的全部数据行（跳过合计行）。"""
        books = self._selected_books()
        level = self.cb_level.currentData()
        rows = []
        if level is None or level < 0:
            for b in books:
                for sec in b.iter_self_and_desc():
                    rows.extend(sec.rows)
        else:
            for b in books:
                for node in self._nodes_at_depth(b, level):
                    rows.extend(node.rows)
        return [r for r in rows if not getattr(r, 'total_row', False)]

    @staticmethod
    def _is_blank_row(r):
        """行是否「什么都没有」——默认那 15 行空行以前会自成一组「(空)」，
        是分类汇总显示混乱的最大来源。V3.7.10 起空行不参与汇总。"""
        if r.value is not None:
            return False
        for k in ('code', 'name', 'unit', 'formula', 'remark'):
            if str(getattr(r, k, '') or '').strip():
                return False
        for v in (getattr(r, 'custom', {}) or {}).values():
            if str(v or '').strip():
                return False
        return True

    def _stat_rows(self):
        """参与统计的行：范围内、跳过合计行与全空行。"""
        return [r for r in self._scope_rows() if not self._is_blank_row(r)]

    def _regex_col_key(self):
        """正则作用列；为空表示跟随分类列。"""
        k = self.cb_regex_col.currentData()
        return str(k) if k else None

    def _build_groups(self):
        """按当前分类字段（+可选正则）分组，返回 ([(组名, 行列表)...], 错误)。

        V3.7.10：正则口径与「筛选」一致——
        · 方式=筛选匹配行（默认）：在「正则列」上匹配，**命中的行才参与汇总**，
          不命中的直接排除（不再归进「(未匹配)」组）；
        · 方式=按正则提取分组名：保持旧行为，按匹配内容归组，
          没匹配上的归「(未匹配)」。
        """
        key = self.cb_group.currentData()
        if not key:
            return None, '没有选择分类汇总列'
        pat = self.ed_regex.text().strip()
        mode = self.cb_regex_mode.currentData() or 'filter'
        rx = None
        self._regex_error = ''
        if pat:
            try:
                rx = _re.compile(norm_text(pat))
            except _re.error as e:
                # 正则写错：跟「筛选」一样不排除任何行，只提示一下
                self._regex_error = ('正则表达式无效：%s（本次暂按全部统计）'
                                     % e)
                rx = None
        rkey = self._regex_col_key() or key
        groups = {}
        order = []
        matched = 0
        for r in self._stat_rows():
            gval = str(_row_field(r, key) or '').strip()
            if rx is not None and mode == 'filter':
                # 与筛选同一套口径：整行归一化后做 re.search
                txt = str(_row_field(r, rkey) or '')
                if not filter_match(txt, '', '正则匹配', pat):
                    continue
                matched += 1
                g = gval
            elif rx is not None and mode == 'extract':
                m = rx.search(norm_text(gval))
                if m:
                    g = (m.group(1) if m.groups() else m.group(0)) or '(未匹配)'
                else:
                    g = '(未匹配)'
            else:
                g = gval
            g = g or '(空)'
            if g not in groups:
                groups[g] = []
                order.append(g)
            groups[g].append(r)
        self._last_matched = matched
        return [(g, groups[g]) for g in order], None

    @staticmethod
    def _group_unit(rows):
        """组内单位：只有一种就显示它；多种就都列出来（混单位的小计要当心）。"""
        units = []
        for r in rows:
            u = (r.unit or '').strip()
            if u and u not in units:
                units.append(u)
        return '/'.join(units), len(units) > 1

    def _apply_tree_mode(self, group_mode):
        """分组模式与普通报表模式的表头/列显隐切换。"""
        if group_mode:
            self.tree.setHeaderLabels(self.GROUP_COLS)
            for i in range(len(self.GROUP_COLS), self.tree.columnCount()):
                self.tree.setColumnHidden(i, True)
            self.tree.setColumnWidth(0, 280)
            self.tree.setColumnWidth(1, 110)
            self.tree.setColumnWidth(2, 100)
            self.tree.setColumnWidth(3, 150)
        else:
            self.tree.setHeaderLabels(self.COLS)
            for i in range(len(self.GROUP_COLS), self.tree.columnCount()):
                self.tree.setColumnHidden(i, False)
            self.tree.setColumnWidth(0, 110)
            self.tree.setColumnWidth(1, 260)
            self.tree.setColumnWidth(2, 70)
            self.tree.setColumnWidth(3, 110)
            self.tree.setColumnWidth(4, 320)
            self.tree.setColumnWidth(5, 110)
            self.tree.setColumnWidth(6, 160)

    def _refresh_group_tree(self, prec):
        """分类汇总模式：一张干净的汇总表（分组 / 单位 / 条数 / 工程量合计），
        双击某一行看该组明细。"""
        self.tree.clear()
        self._apply_tree_mode(True)
        groups, err = self._build_groups()
        if err:
            self.lb_range.setText(err)
            self.lb_range.setStyleSheet('color:#b03030;')
            self.lb_total.setText('')
            return
        label = self.cb_group.currentText()
        scope = '全部计算书' if self.cb_book.currentData() is None \
            else (self.cb_book.currentText() or '全部计算书')
        txt = '分类汇总：按「%s」· %s' % (label, scope)
        if self.cb_level.currentData() not in (None, -1):
            txt += ' · %s' % (self.cb_level.currentText() or '')
        pat = self.ed_regex.text().strip()
        total_rows = len(self._stat_rows())
        if pat and (self.cb_regex_mode.currentData() or 'filter') == 'filter' \
                and not self._regex_error:
            rk = self._regex_col_key() or label
            txt += ' · 正则「%s」筛出 %d / %d 条' % (
                pat, getattr(self, '_last_matched', 0), total_rows)
        elif pat and (self.cb_regex_mode.currentData() or 'filter') == 'extract':
            txt += ' · 正则「%s」（提取分组名）' % pat
        self.lb_range.setText(txt)
        self.lb_range.setStyleSheet('color:#b03030;' if self._regex_error
                                    else 'color:#444;')
        if self._regex_error:
            self.lb_range.setText(self._regex_error + '　' + txt)
        grand = 0.0
        grand_rows = 0
        mixed = False
        for g, rows in groups:
            vals = [r.value for r in rows
                    if not r.no_count and r.value is not None]
            tot = sum(vals)
            unit, is_mixed = self._group_unit(rows)
            mixed = mixed or is_mixed
            item = QTreeWidgetItem([g, unit, '共 %d 条' % len(rows),
                                    fmt_num(tot, prec) if vals else ''])
            f = item.font(0)
            f.setBold(True)
            item.setFont(0, f)
            item.setForeground(0, QBrush(QColor('#1f5f99')))
            item.setForeground(3, QBrush(QColor('#1f5f99')))
            if is_mixed:
                item.setToolTip(0, '本组含多种计量单位（%s），小计仅供参考' % unit)
                item.setForeground(1, QBrush(QColor('#b03030')))
            item.setData(0, Qt.UserRole, rows)
            self.tree.addTopLevelItem(item)
            grand += tot
            grand_rows += len(rows)
        tip = '合计：%.3f    共 %d 条' % (grand, grand_rows)
        if mixed:
            tip += '（有组含多种计量单位）'
        self.lb_total.setText(tip)

    def _on_group_double_clicked(self, item, col):
        """双击分组行 → 看该组明细（普通报表模式下不做任何事）。"""
        rows = item.data(0, Qt.UserRole)
        if not isinstance(rows, list) or not rows:
            return
        dlg = QDialog(self)
        dlg.setWindowTitle('分组明细 —— %s' % item.text(0))
        dlg.resize(900, 460)
        v = QVBoxLayout(dlg)
        cols = ['编号', '项目名称', '单位', '工程量', '计算公式', '备注']
        tb = QTableWidget(len(rows), len(cols))
        tb.setHorizontalHeaderLabels(cols)
        tb.setEditTriggers(QAbstractItemView.NoEditTriggers)
        tb.setSelectionBehavior(QAbstractItemView.SelectRows)
        prec = getattr(self.win, 'precision', 3)
        for i, r in enumerate(rows):
            val = '' if r.no_count else (
                fmt_num(r.value, precision_for_unit(r.unit, prec))
                if r.value is not None else '')
            for j, txt in enumerate([r.code or '', r.name or '', r.unit or '',
                                     val, r.formula or '', r.remark or '']):
                it = QTableWidgetItem(str(txt))
                tb.setItem(i, j, it)
        tb.resizeColumnsToContents()
        v.addWidget(tb)
        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.button(QDialogButtonBox.Close).setText('关闭')
        bb.rejected.connect(dlg.reject)
        v.addWidget(bb)
        dlg.exec_()

    def _summary_html(self):
        """把分类汇总结果转成 HTML（供打印预览）。返回 (html, 错误)。"""
        import html as _html
        groups, err = self._build_groups()
        if err:
            return None, err
        prec = getattr(self.win, 'precision', 3)
        rows_html = []
        grand = 0.0
        for g, rows in groups:
            vals = [r.value for r in rows
                    if not r.no_count and r.value is not None]
            tot = sum(vals)
            grand += tot
            unit, is_mixed = self._group_unit(rows)
            rows_html.append(
                '<tr><td>%s</td><td align="center">%s</td>'
                '<td align="center">%d</td><td align="right">%s</td></tr>'
                % (_html.escape(g), _html.escape(unit), len(rows),
                   fmt_num(tot, prec) if vals else ''))
        label = self.cb_group.currentText()
        pat = self.ed_regex.text().strip()
        meta = '按「%s」分类%s' % (label,
                                '（正则：%s）' % _html.escape(pat) if pat else '')
        return (
            '<h2 style="font-family:sans-serif">工程量分类汇总表</h2>'
            '<p style="color:#444">%s ｜ 汇总范围：%s</p>'
            '<table width="100%%" border="1" cellspacing="0" cellpadding="4" '
            'style="border-collapse:collapse">'
            '<tr style="background:#eef2f7"><th align="left">分组</th>'
            '<th>单位</th><th>条数</th><th align="right">工程量合计</th></tr>'
            '%s'
            '<tr><th align="left">总计</th><th></th><th></th>'
            '<th align="right">%.3f</th></tr>'
            '</table>' % (meta, _html.escape(self.cb_book.currentText()
                                             or '全部计算书'),
                          ''.join(rows_html), grand)), None

    def _build_print_preview(self, printer, doc, title='打印预览'):
        """建打印预览窗：窗口与报表同尺寸、工具条按钮中文、打印走我们自己的对话框。"""
        from PyQt5.QtPrintSupport import QPrintPreviewDialog
        dlg = QPrintPreviewDialog(printer, self)
        dlg.setWindowTitle(title)
        localize_print_preview(dlg)              # 工具条按钮中文化（并隐藏 Qt 的打印）
        dlg.paintRequested.connect(lambda p: doc.print_(p))
        # 底部加我们自己的「打印…」：Qt 自带的打印按钮用的是它内部的打印框，
        # 页码范围拿不到、也没法预填页面设置里的范围 → 统一走我们的打印流程
        bar = QWidget()
        bl = QHBoxLayout(bar)
        bl.setContentsMargins(8, 4, 8, 6)
        bl.addStretch(1)
        btn_print = QPushButton('打印…')
        btn_print.setToolTip('打开系统打印窗口，可在其中选择页码范围')
        btn_print.clicked.connect(lambda: self._on_print_doc(dlg, doc, title))
        bl.addWidget(btn_print)
        btn_close = QPushButton('关闭')
        btn_close.clicked.connect(dlg.close)
        bl.addWidget(btn_close)
        lay = dlg.layout()
        if lay is not None:
            lay.addWidget(bar)
        # 窗口大小与「查看报表」协调（Qt 默认只有 640×480，太小）
        w = max(self.width(), 1080)
        h = max(self.height(), 680)
        dlg.resize(w, h)
        try:
            g = self.geometry()
            dlg.move(g.center().x() - w // 2, g.center().y() - h // 2)
        except Exception:
            pass
        return dlg

    def _on_print_doc(self, dlg, doc, title='打印'):
        """预览窗里的「打印…」：用我们自己的打印窗口（页码范围可控）。"""
        try:
            printer = dlg.printer()          # 复用预览窗的打印机，分页与预览一致
        except Exception:
            printer = None
        if printer is None:
            printer = self.win.make_printer()
        # 预览控件知道真实总页数（比估算准），用它来限制页码范围
        total = 0
        try:
            from PyQt5.QtPrintSupport import QPrintPreviewWidget
            wid = dlg.findChild(QPrintPreviewWidget)
            if wid is not None:
                total = int(wid.pageCount())
        except Exception:
            total = 0
        try:
            self.win.print_document(printer, doc, title, total=total)
        except Exception as e:
            QMessageBox.warning(self, '打印失败', str(e))

    def _detail_html(self):
        """不汇总时：把当前报表明细（按范围）转成 HTML，供打印预览。

        V3.7.11：以前「不汇总」点打印预览只会弹「没有选择分类汇总列」，
        现在两种模式都能预览 —— 汇总看汇总表，不汇总看明细表。
        """
        import html as _html
        prec = getattr(self.win, 'precision', 3)
        books = self._selected_books()
        level = self.cb_level.currentData()
        rows_html = []
        grand = 0.0
        grand_rows = 0

        def walk(sec, depth=0):
            rows_html.append(
                '<tr><td colspan="%d" style="background:#fff2cc">'
                '<b>%s%s</b></td></tr>'
                % (len(self.COLS), '　' * depth, _html.escape(sec.name)))
            for r in getattr(sec, 'rows', []) or []:
                if getattr(r, 'total_row', False):
                    continue          # 合计行不重复列（下面有小计）
                val = '' if r.no_count else (
                    fmt_num(r.value, precision_for_unit(r.unit, prec))
                    if r.value is not None else '')
                cells = [r.code or '', r.name or '', r.unit or '', val,
                         r.formula or '', _row_field(r, 'drawing_no'),
                         _row_field(r, 'remark')]
                rows_html.append('<tr>' + ''.join(
                    '<td>%s</td>' % _html.escape(str(c)) for c in cells) + '</tr>')
            # 小计只算本分部自己的行（与打印出来的计算书一致，不含下级，
            # 否则上级小计与下级小计会重复显示同一个数）
            t = sec.total()
            if t is not None:
                rows_html.append(
                    '<tr style="background:#e2efda">'
                    '<td colspan="3" align="right"><b>小计</b></td>'
                    '<td align="right"><b>%s</b></td>'
                    '<td colspan="3"></td></tr>'
                    % fmt_num(t, section_precision(sec, prec)))
            for c in getattr(sec, 'children', []) or []:
                walk(c, depth + 1)

        if level is None or level < 0:
            for book in books:
                walk(book, 0)
        else:
            for book in books:
                for node in self._nodes_at_depth(book, level):
                    walk(node, 0)
        grand, grand_rows = self._node_total_of(books, level)
        scope = self.cb_book.currentText() or '全部计算书'
        if level is not None and level >= 0:
            scope += ' · %s' % (self.cb_level.currentText() or '')
        headers = ''.join('<th align="left">%s</th>' % h for h in self.COLS)
        return (
            '<h2 style="font-family:sans-serif">工程量明细报表</h2>'
            '<p style="color:#444">汇总范围：%s ｜ 共 %d 条计算式</p>'
            '<table width="100%%" border="1" cellspacing="0" cellpadding="4" '
            'style="border-collapse:collapse;font-size:9pt">'
            '<tr style="background:#eef2f7">%s</tr>'
            '%s'
            '<tr style="background:#e2efda"><th align="left">总计</th>'
            '<th></th><th></th><th align="right">%.3f</th>'
            '<th></th><th></th><th></th></tr>'
            '</table>'
            % (_html.escape(scope), grand_rows, headers,
               ''.join(rows_html), grand)), None

    def _node_total_of(self, books, level):
        """范围内（计算书 × 层级）的合计与条数。"""
        total, cnt = 0.0, 0
        if level is None or level < 0:
            for b in books:
                t, c = self._node_total(b)
                total += t
                cnt += c
        else:
            for b in books:
                for node in self._nodes_at_depth(b, level):
                    t, c = self._node_total(node)
                    total += t
                    cnt += c
        return total, cnt

    def _on_print_preview(self):
        """打印预览：汇总模式预览汇总表；不汇总模式预览报表明细。"""
        if self.cb_group.currentData():
            html, err = self._summary_html()
            title = '打印预览——工程量分类汇总表'
        else:
            html, err = self._detail_html()
            title = '打印预览——工程量明细报表'
        if err:
            QMessageBox.warning(self, '打印预览', err)
            return
        from PyQt5.QtGui import QTextDocument
        printer = self.win.make_printer()
        doc = QTextDocument()
        doc.setHtml(html)
        dlg = self._build_print_preview(printer, doc, title)
        dlg.exec_()

    # 兼容旧调用名
    _on_print_summary = _on_print_preview

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
        self._refresh_group_box()
        self._refresh_regex_col_box()
        self.tree.clear()
        proj = win.project
        prec = getattr(win, 'precision', 3)
        # V3.7.9：选了分类字段 → 分组汇总视图（每组小计）
        if self.cb_group.currentData():
            self._refresh_group_tree(prec)
            return
        self._apply_tree_mode(False)      # V3.7.10：普通报表模式恢复原表头
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
                                   _row_field(r, 'remark')])
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
