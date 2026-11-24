# -*- coding: utf-8 -*-
"""冻结窗格（V3.2）：仿 WPS 锁定行 / 列 / 行列同时锁定。

在主表视口上叠加最多三块只读镜像表：
  fc —— 左侧冻结列（整高，垂直滚动随主表同步）
  fr —— 顶部冻结行（整宽，水平滚动随主表同步）
  fx —— 左上角行列交叉区
镜像内容从主表 QTableWidget 克隆（文字/字体/颜色/底色/对齐/勾选）。
单击镜像会把主表光标移到对应单元格；双击/主表进入编辑时暂时隐藏镜像，
编辑结束自动恢复。
"""

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (QTableWidget, QTableWidgetItem, QAbstractItemView,
                             QAbstractItemDelegate)


def _clone(item):
    if item is None:
        return None
    return QTableWidgetItem(item)


class _MirrorTable(QTableWidget):
    """冻结区镜像：自身不滚动，鼠标滚轮转给主表（WPS 同款体验）。"""

    def __init__(self, panes):
        super().__init__(0, 0, panes.t)
        self._panes = panes

    def contextMenuEvent(self, ev):
        """V3.7.4：冻结区右键 → 转发给主表的单元格右键菜单（以前无任何菜单）。"""
        try:
            t = self._panes.t
            pos = t.viewport().mapFromGlobal(ev.globalPos())
            t.customContextMenuRequested.emit(pos)
            ev.accept()
        except Exception:
            pass

    def wheelEvent(self, ev):
        try:
            dx = ev.angleDelta().x()
            if abs(dx) > abs(ev.angleDelta().y()) and dx:
                # 横向滚轮：Qt 默认只处理纵向，这里直接滚主表横向滚动条
                bar = self._panes.t.horizontalScrollBar()
                bar.setValue(bar.value() - int(dx / 8))
            else:
                # 滚轮落在冻结行/列上时，应该滚动的是整张表，而不是镜像本身
                self._panes.t.wheelEvent(ev)
        except Exception:
            pass


class FrozenPanes(object):

    def __init__(self, table):
        self.t = table
        self._rows = 0
        self._cols = 0
        self._editing = False          # 主表正在编辑（镜像临时让位）
        self.fc = None   # 左侧冻结列镜像
        self.fr = None   # 顶部冻结行镜像
        self.fx = None   # 左上角交叉镜像

    # ---- 状态 ----

    @property
    def rows(self):
        return self._rows

    @property
    def cols(self):
        return self._cols

    def set_freeze(self, rows, cols):
        """设置冻结行数/列数（自动收敛到合法范围）。"""
        rc, cc = self.t.rowCount(), self.t.columnCount()
        rows = max(0, min(int(rows), rc - 1)) if rc > 1 else 0
        cols = max(0, min(int(cols), cc - 1)) if cc > 1 else 0
        self._rows, self._cols = rows, cols
        if rows == 0 and cols == 0:
            self.teardown()
        else:
            self._ensure()
            self.rebuild()

    def teardown(self):
        for w in (self.fc, self.fr, self.fx):
            if w is not None:
                w.hide()
                w.deleteLater()
        self.fc = self.fr = self.fx = None

    def _mk(self):
        t = self.t
        w = _MirrorTable(self)
        w.setFocusPolicy(Qt.NoFocus)
        w.setSelectionMode(QAbstractItemView.NoSelection)
        w.setEditTriggers(QAbstractItemView.NoEditTriggers)
        w.setVerticalScrollMode(t.verticalScrollMode())
        w.setHorizontalScrollMode(t.horizontalScrollMode())
        w.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        w.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        w.verticalHeader().hide()
        w.horizontalHeader().hide()
        w.setAlternatingRowColors(t.alternatingRowColors())
        w.setWordWrap(t.wordWrap())
        w.setFrameShape(t.frameShape())
        w.setStyleSheet('QTableWidget { border: none; }')
        w.cellClicked.connect(self._clicked)
        w.cellDoubleClicked.connect(self._dbl_clicked)
        return w

    def _ensure(self):
        if self.fc is None:
            self.fc = self._mk()
        if self.fr is None:
            self.fr = self._mk()
        if self.fx is None:
            self.fx = self._mk()

    # ---- 交互 ----

    def _clicked(self, r, c):
        """点镜像 → 主表光标/选中跟随。"""
        self.t.setCurrentCell(r, c)
        self.t.setFocus()

    def _dbl_clicked(self, r, c):
        """双击镜像 → 主表在原位编辑该格（镜像暂时让位）。"""
        if self._rows <= 0 and self._cols <= 0:
            return
        self.t.setCurrentCell(r, c)
        self.hide_for_edit()
        QTimer.singleShot(0, lambda: self._start_edit(r, c))

    def _start_edit(self, r, c):
        try:
            self.t.edit(self.t.model().index(r, c),
                        QAbstractItemView.AllEditTriggers, None)
        except Exception:
            self.update_geometry()

    def hide_for_edit(self):
        """主表进入冻结区编辑时隐藏镜像，编辑器才看得见。"""
        self._editing = True
        for w in (self.fc, self.fr, self.fx):
            if w is not None:
                w.hide()

    def show_after_edit(self):
        """编辑结束恢复镜像。"""
        self._editing = False
        self.update_geometry()

    # ---- 内容 ----

    def rebuild(self):
        """refresh_table / 行列数变化后调用：重建镜像内容。"""
        if self._rows <= 0 and self._cols <= 0:
            return
        self._ensure()
        t = self.t
        rc, cc = t.rowCount(), t.columnCount()
        if self._cols > 0 and self.fc is not None:
            fc = self.fc
            fc.blockSignals(True)
            fc.setRowCount(rc)
            fc.setColumnCount(self._cols)
            for r in range(rc):
                for c in range(self._cols):
                    fc.setItem(r, c, _clone(t.item(r, c)))
                fc.setRowHeight(r, t.rowHeight(r))
                fc.setRowHidden(r, t.isRowHidden(r))
            for c in range(self._cols):
                fc.setColumnWidth(c, t.columnWidth(c))
            fc.blockSignals(False)
        if self._rows > 0 and self.fr is not None:
            fr = self.fr
            fr.blockSignals(True)
            fr.setRowCount(self._rows)
            fr.setColumnCount(cc)
            for r in range(self._rows):
                for c in range(cc):
                    fr.setItem(r, c, _clone(t.item(r, c)))
                fr.setRowHeight(r, t.rowHeight(r))
                fr.setRowHidden(r, t.isRowHidden(r))
            for c in range(cc):
                fr.setColumnWidth(c, t.columnWidth(c))
            fr.blockSignals(False)
        if self._rows > 0 and self._cols > 0 and self.fx is not None:
            fx = self.fx
            fx.blockSignals(True)
            fx.setRowCount(self._rows)
            fx.setColumnCount(self._cols)
            for r in range(self._rows):
                for c in range(self._cols):
                    fx.setItem(r, c, _clone(t.item(r, c)))
                fx.setRowHeight(r, t.rowHeight(r))
            for c in range(self._cols):
                fx.setColumnWidth(c, t.columnWidth(c))
            fx.blockSignals(False)
        self.sync_v(self.t.verticalScrollBar().value())
        self.sync_h(self.t.horizontalScrollBar().value())
        self.update_geometry()

    def patch_cell(self, r, c):
        """主表单个单元格内容/格式变化后，只更新镜像里对应格（省性能）。"""
        if self._rows <= 0 and self._cols <= 0:
            return
        it = self.t.item(r, c)
        if self.fc is not None and self._cols > 0 and c < self._cols \
                and 0 <= r < self.fc.rowCount():
            self.fc.setItem(r, c, _clone(it))
        if self.fr is not None and self._rows > 0 and r < self._rows \
                and 0 <= c < self.fr.columnCount():
            self.fr.setItem(r, c, _clone(it))
        if self.fx is not None and self._rows > 0 and self._cols > 0 \
                and r < self._rows and c < self._cols:
            self.fx.setItem(r, c, _clone(it))

    def sync_hidden(self):
        """筛选/取消筛选后镜像隐藏行跟随主表。"""
        t = self.t
        if self.fc is not None and self._cols > 0:
            for r in range(min(t.rowCount(), self.fc.rowCount())):
                self.fc.setRowHidden(r, t.isRowHidden(r))
        if self.fr is not None and self._rows > 0:
            for r in range(self._rows):
                self.fr.setRowHidden(r, t.isRowHidden(r))
        self.update_geometry()

    # ---- 同步 ----

    def _keep_visible(self):
        """V3.7.4：滚动时兜底恢复被意外隐藏的镜像（编辑让位期间除外）。"""
        if self._editing:
            return
        for w in (self.fc, self.fr, self.fx):
            if w is None:
                continue
            active = ((w is self.fc and self._cols > 0) or
                      (w is self.fr and self._rows > 0) or
                      (w is self.fx and self._rows > 0 and self._cols > 0))
            if active and not w.isVisible():
                w.show()
                w.raise_()

    def sync_one_height(self, r, h):
        """V3.7.4：主表某一行行高变化 → 只同步镜像里该行的行高（省性能）。"""
        if self._rows <= 0 and self._cols <= 0:
            return
        if self.fc is not None and self._cols > 0 and 0 <= r < self.fc.rowCount():
            self.fc.setRowHeight(r, h)
        if self.fr is not None and self._rows > 0 and r < self._rows:
            self.fr.setRowHeight(r, h)
        if self.fx is not None and self._rows > 0 and self._cols > 0 and r < self._rows:
            self.fx.setRowHeight(r, h)
        if r < self._rows:
            self.update_geometry()

    def sync_v(self, value):
        """主表垂直滚动 → 冻结列镜像跟随。"""
        if self.fc is not None and self._cols > 0:
            self.fc.verticalScrollBar().setValue(int(value))
            self._keep_visible()

    def sync_h(self, value):
        """主表水平滚动 → 冻结行镜像跟随。"""
        if self.fr is not None and self._rows > 0:
            self.fr.horizontalScrollBar().setValue(int(value))
            self._keep_visible()

    def sync_widths(self):
        """主表列宽变化（拖动/自适应）后同步镜像列宽与几何。"""
        if self._rows <= 0 and self._cols <= 0:
            return
        t = self.t
        if self.fc is not None and self._cols > 0:
            for c in range(self._cols):
                self.fc.setColumnWidth(c, t.columnWidth(c))
        if self.fr is not None and self._rows > 0:
            for c in range(t.columnCount()):
                self.fr.setColumnWidth(c, t.columnWidth(c))
        if self.fx is not None and self._rows > 0 and self._cols > 0:
            for c in range(self._cols):
                self.fx.setColumnWidth(c, t.columnWidth(c))
        self.update_geometry()

    def sync_heights(self):
        """主表行高变化（自适应/手动设定）后同步镜像行高与几何。"""
        if self._rows <= 0 and self._cols <= 0:
            return
        t = self.t
        if self.fc is not None and self._cols > 0:
            for r in range(min(t.rowCount(), self.fc.rowCount())):
                self.fc.setRowHeight(r, t.rowHeight(r))
        if self.fr is not None and self._rows > 0:
            for r in range(self._rows):
                self.fr.setRowHeight(r, t.rowHeight(r))
        if self.fx is not None and self._rows > 0 and self._cols > 0:
            for r in range(self._rows):
                self.fx.setRowHeight(r, t.rowHeight(r))
        self.update_geometry()

    def update_geometry(self):
        """按当前冻结行列数与主表尺寸摆放镜像。"""
        if self._rows <= 0 and self._cols <= 0:
            return
        t = self.t
        vp = t.viewport().geometry()
        fcw = sum(t.columnWidth(c) for c in range(self._cols)) if self._cols else 0
        frh = sum(t.rowHeight(r) for r in range(self._rows)) if self._rows else 0
        if self._cols > 0 and self.fc is not None:
            self.fc.setGeometry(vp.x(), vp.y(), fcw, max(1, vp.height()))
            self.fc.show()
            self.fc.raise_()
        elif self.fc is not None:
            self.fc.hide()
        if self._rows > 0 and self.fr is not None:
            self.fr.setGeometry(vp.x() + fcw, vp.y(),
                                max(1, vp.width() - fcw), frh)
            self.fr.show()
            self.fr.raise_()
        elif self.fr is not None:
            self.fr.hide()
        if self._rows > 0 and self._cols > 0 and self.fx is not None:
            self.fx.setGeometry(vp.x(), vp.y(), fcw, frh)
            self.fx.show()
            self.fx.raise_()
        elif self.fx is not None:
            self.fx.hide()
