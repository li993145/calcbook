# -*- coding: utf-8 -*-
"""WPS 式列字母标尺：在表格水平表头上方显示 A/B/C… 列字母。

字母与公式引用一致（见 core.model.column_letters）：
工程量=A、构件数量=B、计算公式=C、备注=D、自定义列依次 E、F…
序号/编号/项目名称/计量单位/图纸/不计 是标识列，不参与引用，留空。

标尺与主表列宽、列拖动、水平滚动完全同步；点击字母可选中整列。
"""

from PyQt5.QtCore import Qt, QRect
from PyQt5.QtGui import QPainter, QColor, QFont, QPen
from PyQt5.QtWidgets import QWidget


class ColumnLetterBar(QWidget):
    def __init__(self, win, parent=None):
        super().__init__(parent)
        self.win = win
        self.setFixedHeight(24)
        self.setMouseTracking(True)
        self.setToolTip('列字母标尺：与公式引用一致（工程量=A、构件数量=B、计算公式=C…）；'
                        '点击字母选中整列')

        table = self.win.table
        hh = table.horizontalHeader()
        hh.sectionResized.connect(self.update)
        hh.sectionMoved.connect(self.update)
        hh.geometriesChanged.connect(self.update)
        table.verticalHeader().geometriesChanged.connect(self.update)
        table.horizontalScrollBar().valueChanged.connect(self.update)

    # ---------- 几何 ----------
    def _columns(self):
        try:
            return list(self.win._cols())
        except Exception:
            return []

    def _letters(self):
        try:
            return dict(self.win.calc_letters())
        except Exception:
            return {}

    def _column_at_x(self, x):
        """返回鼠标 x 处的逻辑列号（考虑水平滚动与列拖动），不在任何列上返回 -1。"""
        table = self.win.table
        hh = table.horizontalHeader()
        left = -table.horizontalScrollBar().value()
        cx = x
        for vis in range(hh.count()):
            logical = hh.logicalIndex(vis)
            w = table.columnWidth(logical)
            if left <= cx < left + w:
                return logical
            left += w
        return -1

    # ---------- 交互 ----------
    def mousePressEvent(self, ev):
        if ev.button() != Qt.LeftButton:
            return
        col = self._column_at_x(ev.pos().x())
        if col < 0:
            return
        table = self.win.table
        table.setCurrentCell(table.currentRow() if table.currentRow() >= 0 else 0, col)
        table.selectColumn(col)
        table.setFocus()

    # ---------- 绘制 ----------
    def paintEvent(self, ev):
        super().paintEvent(ev)
        p = QPainter(self)
        p.fillRect(self.rect(), QColor('#eef2f7'))

        table = self.win.table
        hh = table.horizontalHeader()
        cols = self._columns()
        letters = self._letters()
        if not cols:
            return

        # 关键：sectionViewportPosition() 已含水平滚动，直接给出每段在表头
        # viewport 里的 x；再累加表头在表格中的位置（行号列宽 31px）即是
        # 标尺坐标。切勿用 horizontalScrollBar().value() 当像素偏移——
        # 它的单位是「列」，与像素位移完全不成比例，会导致字母整体错位。
        hx = hh.x()
        cur_visual = hh.visualIndex(table.currentColumn()) if table.currentColumn() >= 0 else -1
        font = QFont(self.font())
        font.setPointSize(max(8, font.pointSize() - 1))
        p.setFont(font)
        border = QColor('#d5dce5')

        for vis in range(hh.count()):
            logical = hh.logicalIndex(vis)
            if logical >= len(cols):
                break
            w = table.columnWidth(logical)
            if w <= 0:
                continue
            x = hx + hh.sectionViewportPosition(vis)
            rect = QRect(int(x), 0, int(w), self.height())
            # 当前列底色（与主表选中列呼应）
            if vis == cur_visual:
                p.fillRect(rect, QColor('#dce9f7'))
            # 可计算列显示引用字母；标识列留空
            letter = letters.get(cols[logical]['key'], '')
            if letter:
                p.setPen(QColor('#305080'))
                p.drawText(rect, Qt.AlignCenter, letter)
            # 列分隔线
            p.setPen(border)
            p.drawLine(rect.right(), 0, rect.right(), self.height() - 1)

        # 底边线
        p.setPen(QColor('#c3ccd8'))
        p.drawLine(0, self.height() - 1, self.width(), self.height() - 1)
