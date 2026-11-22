# -*- coding: utf-8 -*-
"""导入表格文件对话框。

从 .xlsx / .csv / .txt 表格文件读取数据，使用者可以：
1. 选择文件与工作表（xlsx 多工作表时）；
2. 指定表头行 / 数据起始行；
3. 在预览中查看源数据；
4. 逐列勾选要导入的字段，并指定该列写入计算书的哪个字段（列）；
5. 指定写入位置：追加到当前分部末尾，或从指定行开始覆盖写入。
"""

import os
import csv

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QFileDialog, QComboBox, QCheckBox, QSpinBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QDialogButtonBox,
    QMessageBox, QRadioButton, QButtonGroup, QGroupBox,
)
from PyQt5.QtCore import Qt

MAX_PREVIEW_ROWS = 100     # 预览最多显示的行数
MAX_COLS = 60              # 超过这么多列的文件不再处理（异常保护）


def read_table_file(path, sheet_name=None):
    """读取表格文件，返回 (工作表名列表, 二维字符串列表)。

    支持 .xlsx（openpyxl，读单元格显示值）、.csv（逗号分隔，自动识别
    utf-8 / gbk / utf-16 编码）、.txt（制表符分隔）。
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == '.xlsx':
        try:
            from openpyxl import load_workbook
        except ImportError:
            raise RuntimeError('读取 .xlsx 需要 openpyxl 库（安装版请先 pip install openpyxl）')
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            names = list(wb.sheetnames)
            ws = wb[sheet_name] if sheet_name in names else wb[names[0]]
            data = [['' if v is None else str(v) for v in row]
                    for row in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
        return names, data
    # ---- 纯文本：.csv 逗号分隔 / .txt 制表符分隔 ----
    text = None
    last_err = None
    for enc in ('utf-8-sig', 'gbk', 'utf-16'):
        try:
            with open(path, 'r', encoding=enc) as f:
                text = f.read()
            break
        except (UnicodeDecodeError, UnicodeError) as e:
            last_err = e
    if text is None:
        raise RuntimeError('无法识别文件编码（%s）' % last_err)
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    if lines and lines[-1] == '':
        lines = lines[:-1]
    if ext == '.csv':
        rows = list(csv.reader(lines))
    else:
        rows = [ln.split('\t') for ln in lines]
    return ['文本'], rows


def trim_data(rows):
    """去掉末尾整行 / 整列为空的部分，并把每行截到实际有内容的最大列。"""
    if not rows:
        return []
    ncol = 0
    for r in rows:
        for j in range(len(r) - 1, -1, -1):
            if str(r[j]).strip():
                ncol = max(ncol, j + 1)
                break
    ncol = min(ncol, MAX_COLS)
    data = []
    for r in rows:
        if any(str(x).strip() for x in r):
            data.append([str(x) for x in (r[:ncol] if ncol else r)])
    return data


class ImportTableDialog(QDialog):
    """导入表格对话框：选文件 → 勾选字段并映射目标列 → 指定写入起始行。"""

    def __init__(self, parent, columns, default_row=1, path=''):
        super().__init__(parent)
        self.setWindowTitle('导入表格文件')
        self.resize(780, 660)
        # 可映射的目标字段（工程量列由公式计算得出，不作为导入目标）
        self.targets = [(c['key'], c['title']) for c in columns if c['key'] != 'value']
        # 计算书全部列名（含工程量），用于识别“工程量”列自动跳过
        all_titles = {c['title'].strip(): c['key'] for c in columns}
        self._value_titles = {t for t, k in all_titles.items() if k == 'value'}
        self._saved_map = {}      # 换文件/换表时记住已选映射 {列号: 目标key}
        self._data = []           # 当前文件的整理后数据
        self._headers = None      # 表头（首行是表头时）

        v = QVBoxLayout(self)

        # ===== 第 1 步：选择文件 =====
        g1 = QGroupBox('1. 选择表格文件')
        r1 = QHBoxLayout(g1)
        r1.addWidget(QLabel('文件：'))
        self.ed_file = QLineEdit(path)
        self.ed_file.textChanged.connect(self._on_file_changed)
        r1.addWidget(self.ed_file, 1)
        btn = QPushButton('浏览...')
        btn.clicked.connect(self._browse)
        r1.addWidget(btn)
        v.addWidget(g1)

        # ===== 第 2 步：工作表 / 表头 / 起始行 =====
        g2 = QGroupBox('2. 工作表与读取范围')
        r2 = QHBoxLayout(g2)
        r2.addWidget(QLabel('工作表：'))
        self.cb_sheet = QComboBox()
        self.cb_sheet.currentTextChanged.connect(self._reload)
        r2.addWidget(self.cb_sheet)
        self.ck_header = QCheckBox('首行是表头（用于自动对应字段）')
        self.ck_header.setChecked(True)
        self.ck_header.stateChanged.connect(self._reload)
        r2.addWidget(self.ck_header)
        r2.addWidget(QLabel('数据起始行：'))
        self.sp_from = QSpinBox()
        self.sp_from.setRange(1, 999999)
        self.sp_from.setValue(2)
        self.sp_from.setToolTip('源文件从第几行开始是数据。勾选表头时，\n'
                                '表头取“数据起始行的上一行”。')
        self.sp_from.valueChanged.connect(self._reload)
        r2.addWidget(self.sp_from)
        r2.addStretch(1)
        v.addWidget(g2)

        # ===== 第 3 步：预览 =====
        g3 = QGroupBox('3. 预览（最多显示前 %d 行）' % MAX_PREVIEW_ROWS)
        pv = QVBoxLayout(g3)
        self.preview = QTableWidget()
        self.preview.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.preview.setSelectionMode(QAbstractItemView.NoSelection)
        self.preview.verticalHeader().setDefaultSectionSize(22)
        self.preview.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        pv.addWidget(self.preview)
        v.addWidget(g3, 3)

        # ===== 第 4 步：字段映射 =====
        g4 = QGroupBox('4. 选择要导入的字段（列）—— 每列勾选后指定写入计算书的哪个字段')
        mv = QVBoxLayout(g4)
        self.map_table = QTableWidget(0, 3)
        self.map_table.setHorizontalHeaderLabels(['导入', '源文件列', '写入到计算书的字段'])
        self.map_table.verticalHeader().setVisible(False)
        self.map_table.verticalHeader().setDefaultSectionSize(24)
        th = self.map_table.horizontalHeader()
        th.setSectionResizeMode(0, QHeaderView.Fixed)
        self.map_table.setColumnWidth(0, 46)
        th.setSectionResizeMode(1, QHeaderView.Stretch)
        th.setSectionResizeMode(2, QHeaderView.Stretch)
        mv.addWidget(self.map_table)
        v.addWidget(g4, 3)

        # ===== 第 5 步：写入位置 =====
        g5 = QGroupBox('5. 写入位置（当前分部）')
        r5 = QHBoxLayout(g5)
        self.rb_append = QRadioButton('追加到末尾')
        self.rb_row = QRadioButton('从指定行开始写入（覆盖已有数据）')
        self._tgt_group = QButtonGroup(self)
        self._tgt_group.addButton(self.rb_append)
        self._tgt_group.addButton(self.rb_row)
        self.rb_append.setChecked(True)
        r5.addWidget(self.rb_append)
        r5.addWidget(self.rb_row)
        r5.addWidget(QLabel('目标起始行：'))
        self.sp_target = QSpinBox()
        self.sp_target.setRange(1, 999999)
        self.sp_target.setValue(max(default_row, 1))
        r5.addWidget(self.sp_target)
        r5.addStretch(1)
        v.addWidget(g5)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(self._on_accept)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)

        if path and os.path.exists(path):
            self._on_file_changed()

    # ---------------- 文件读取与界面刷新 ----------------
    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(
            self, '选择表格文件', '',
            '表格文件 (*.xlsx *.csv *.txt);;Excel 工作簿 (*.xlsx);;'
            'CSV 文件 (*.csv);;文本文件 (*.txt);;所有文件 (*)')
        if path:
            self.ed_file.setText(path)

    def _on_file_changed(self):
        path = self.file_path()
        if not path or not os.path.exists(path):
            return
        try:
            names, _ = read_table_file(path)
        except Exception as e:
            self._data = []
            self._fill_preview()
            self._fill_mapping()
            QMessageBox.warning(self, '导入表格文件',
                                '无法读取该文件：\n%s' % e)
            return
        self.cb_sheet.blockSignals(True)
        self.cb_sheet.clear()
        self.cb_sheet.addItems(names)
        self.cb_sheet.setEnabled(len(names) > 1)
        self.cb_sheet.blockSignals(False)
        self._reload()

    def _reload(self, *args):
        path = self.file_path()
        if not path or not os.path.exists(path):
            self._data = []
        else:
            try:
                _, raw = read_table_file(path, self.cb_sheet.currentText())
                self._data = trim_data(raw)
            except Exception as e:
                self._data = []
                QMessageBox.warning(self, '导入表格文件', '无法读取该文件：\n%s' % e)
        self._headers = None
        self._fill_preview()
        self._fill_mapping()

    def _fill_preview(self):
        data = self._data[:MAX_PREVIEW_ROWS]
        ncol = max((len(r) for r in data), default=0)
        self.preview.setRowCount(len(data))
        self.preview.setColumnCount(ncol)
        for i, row in enumerate(data):
            for j in range(ncol):
                it = QTableWidgetItem(row[j] if j < len(row) else '')
                it.setFlags(Qt.ItemIsEnabled)
                self.preview.setItem(i, j, it)
        self.preview.resizeColumnsToContents()
        # 预览列宽限制，避免一个超宽单元格撑爆对话框
        for j in range(ncol):
            w = self.preview.columnWidth(j)
            if w > 220:
                self.preview.setColumnWidth(j, 220)

    def _source_headers(self):
        """表头行文本：勾选“首行是表头”时取数据起始行的上一行。"""
        if not self.ck_header.isChecked():
            return None
        hrow = self.sp_from.value() - 2          # 数据起始行的上一行（0 基）
        if 0 <= hrow < len(self._data):
            return self._data[hrow]
        return None

    def _col_label(self, j):
        names = self._source_headers()
        h = ''
        if names and j < len(names):
            h = str(names[j]).strip()
        return ('第%d列' % (j + 1)) + (('（%s）' % h) if h else '')

    def _fill_mapping(self):
        """按当前数据列数重建映射表，尽量保留 / 猜测已有映射。"""
        ncol = max((len(r) for r in self._data), default=0)
        headers = self._source_headers()
        self.map_table.setRowCount(ncol)
        guess = self._guess_map(headers, ncol)
        tkeys = [k for k, _ in self.targets]
        for j in range(ncol):
            # ---- 勾选框 ----
            ck = QCheckBox()
            ck.setChecked(guess.get(j, '') != '')
            ck.setStyleSheet('QCheckBox::indicator{width:16px;height:16px;}')
            ck.stateChanged.connect(
                lambda st, row=j: self._on_check(row, st))
            self.map_table.setCellWidget(j, 0, ck)
            # ---- 源列名 ----
            it = QTableWidgetItem(self._col_label(j))
            it.setFlags(Qt.ItemIsEnabled)
            self.map_table.setItem(j, 1, it)
            # ---- 目标字段下拉 ----
            cb = QComboBox()
            cb.addItem('不导入', '')
            for key, title in self.targets:
                cb.addItem(title, key)
            cur = self._saved_map.get(j, guess.get(j, ''))
            idx = cb.findData(cur)
            cb.setCurrentIndex(max(idx, 0))
            cb.currentIndexChanged.connect(
                lambda i, row=j, box=cb: self._saved_map.__setitem__(
                    row, box.itemData(i)))
            self.map_table.setCellWidget(j, 2, cb)
            self._sync_row_color(j, ck.isChecked())

    def _on_check(self, row, state):
        self._sync_row_color(row, bool(state))
        if not state:
            cb = self.map_table.cellWidget(row, 2)
            if cb is not None:
                self._saved_map[row] = ''

    def _sync_row_color(self, row, checked):
        it = self.map_table.item(row, 1)
        if it is not None:
            it.setForeground(Qt.darkGray if checked else Qt.gray)

    def _guess_map(self, headers, ncols):
        """默认映射：先按表头名对应，识别不出的按位置顺延到相同序号的列。"""
        titles = {}
        for key, title in self.targets:
            titles[title.strip()] = key
        guess = {}
        # 1) 表头名完全相同的直接对应
        if headers:
            for j in range(min(len(headers), ncols)):
                h = str(headers[j]).strip()
                if not h:
                    continue
                if h in self._value_titles:
                    guess[j] = ''          # 工程量列：公式算出来的，默认不导入
                elif h in titles:
                    guess[j] = titles[h]
        # 2) 其余按位置顺延（源第 j 列 → 计算书第 j 个字段）
        used = {v for v in guess.values() if v}
        for j in range(ncols):
            if j in guess:
                continue
            if j < len(self.targets) and self.targets[j][0] not in used:
                guess[j] = self.targets[j][0]
                used.add(self.targets[j][0])
            else:
                guess[j] = ''
        return guess

    # ---------------- 对外取值 ----------------
    def file_path(self):
        return self.ed_file.text().strip()

    def map_pairs(self):
        """返回 [(源列号, 目标字段key), ...]，只含勾选且选了目标字段的列。"""
        pairs = []
        for j in range(self.map_table.rowCount()):
            ck = self.map_table.cellWidget(j, 0)
            cb = self.map_table.cellWidget(j, 2)
            if ck is None or cb is None:
                continue
            if ck.isChecked() and cb.currentData():
                pairs.append((j, cb.currentData()))
        return pairs

    def data_from(self):
        """返回选中的数据行（从数据起始行到末尾）。"""
        start = self.sp_from.value() - 1
        return self._data[start:]

    def target_row0(self):
        """写入位置的 0 基行号；追加模式返回 None。"""
        if self.rb_append.isChecked():
            return None
        return self.sp_target.value() - 1

    # ---------------- 校验 ----------------
    def _on_accept(self):
        if not self.file_path():
            QMessageBox.warning(self, '导入表格文件', '请先选择表格文件。')
            return
        if not self._data:
            QMessageBox.warning(self, '导入表格文件', '文件里没有读到数据。')
            return
        if not self.data_from():
            QMessageBox.warning(self, '导入表格文件',
                                '数据起始行之后没有数据，请调整“数据起始行”。')
            return
        if not self.map_pairs():
            QMessageBox.warning(self, '导入表格文件',
                                '请至少勾选一列并指定要写入的字段。')
            return
        keys = [k for _, k in self.map_pairs()]
        if len(keys) != len(set(keys)):
            QMessageBox.warning(self, '导入表格文件',
                                '有两列写到了同一个字段，请调整字段映射。')
            return
        self.accept()
