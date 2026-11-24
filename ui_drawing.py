# -*- coding: utf-8 -*-
"""V3.7.5 图纸功能：

* 区域截图覆盖层（全屏置顶，拖拽选区，Esc 取消，类似 QQ 截图）；
* 图片放大查看器（滚轮 / 按钮缩放，拖动平移，可打开原图路径）；
* 工作区右侧「图纸预览」面板（跟随选中行联动，点击放大）。

图纸数据统一存成 Row.img 字典：
    {'data': b64(全图), 'fmt': 'png'/'jpg', 'thumb': b64(缩略图),
     'path': 原始文件路径（可空）, 't': 添加时间}
"""

import os
import base64
import hashlib
from datetime import datetime

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QDialog,
    QGraphicsView, QGraphicsScene, QGraphicsPixmapItem, QFileDialog,
    QApplication, QFrame
)
from PyQt5.QtCore import Qt, QRect, QRectF, QSize, QEventLoop, QPoint, pyqtSignal, QTimer
from PyQt5.QtGui import (QPixmap, QPainter, QColor, QPen, QBrush, QGuiApplication,
                         QPixmapCache, QImage, QKeySequence)

MAX_IMG_W = 1600
MAX_IMG_H = 1200


def pixmap_bytes(pix, fmt='PNG', quality=-1):
    from PyQt5.QtCore import QBuffer, QIODevice
    b = QBuffer()
    b.open(QIODevice.WriteOnly)
    pix.save(b, fmt, quality)
    return bytes(b.data())


def encode_img(pix, path=''):
    """QPixmap → Row.img 字典（自动限尺寸；PNG 优先，过大转 JPEG）。"""
    if pix is None or pix.isNull():
        return None
    if pix.width() > MAX_IMG_W or pix.height() > MAX_IMG_H:
        pix = pix.scaled(MAX_IMG_W, MAX_IMG_H,
                         Qt.KeepAspectRatio, Qt.SmoothTransformation)
    raw = pixmap_bytes(pix, 'PNG')
    fmt = 'png'
    if len(raw) > 600 * 1024:
        jpg = pixmap_bytes(pix, 'JPG', 88)
        if 0 < len(jpg) < len(raw):
            raw, fmt = jpg, 'jpg'
    thumb = pix.scaled(200, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    tb = pixmap_bytes(thumb, 'PNG')
    return {'data': base64.b64encode(raw).decode('ascii'),
            'fmt': fmt,
            'thumb': base64.b64encode(tb).decode('ascii'),
            'path': path or '',
            't': datetime.now().strftime('%Y-%m-%d %H:%M')}


_PM_CACHE = {}          # key -> QPixmap（自管缓存，绘制不反复解码）
_PM_CACHE_MAX = 64


def decode_img(d, thumb=False):
    """Row.img 字典 → QPixmap（带缓存）。"""
    if not d:
        return None
    b64 = d.get('thumb') if (thumb and d.get('thumb')) else d.get('data')
    if not b64:
        return None
    key = ('t' if thumb else 'f') + hashlib.md5(b64.encode('ascii')).hexdigest()
    pm = _PM_CACHE.get(key)
    if pm is not None:
        return pm
    try:
        raw = base64.b64decode(b64)
    except Exception:
        return None
    img = QImage.fromData(raw)
    if img.isNull():
        return None
    pm = QPixmap.fromImage(img)
    if len(_PM_CACHE) >= _PM_CACHE_MAX:
        _PM_CACHE.clear()
    _PM_CACHE[key] = pm
    return pm


def img_size(d):
    """取图纸完整尺寸（解码全图，结果有缓存）。"""
    pm = decode_img(d)
    return (pm.width(), pm.height()) if pm is not None else (0, 0)


# ---------------- 区域截图覆盖层 ----------------

class _CaptureOverlay(QWidget):
    """单屏覆盖层：显示冻结的全屏画面 + 半透明暗罩，拖拽选出亮区。"""

    def __init__(self, screen_pix, result, loop):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint |
                         Qt.Tool)
        self.setWindowModality(Qt.ApplicationModal)
        self._pix = screen_pix
        self._result = result
        self._loop = loop
        self._start = None
        self._end = None
        self.setCursor(Qt.CrossCursor)
        self.setWindowTitle('CalcBook 截图')

    def keyPressEvent(self, ev):
        if ev.key() == Qt.Key_Escape:
            self._loop.quit()
            return
        super().keyPressEvent(ev)

    def mousePressEvent(self, ev):
        if ev.button() == Qt.LeftButton:
            self._start = ev.pos()
            self._end = ev.pos()
            self.update()
            return
        self._loop.quit()            # 右键 / 中键 = 取消

    def mouseMoveEvent(self, ev):
        if self._start is not None:
            self._end = ev.pos()
            self.update()

    def mouseReleaseEvent(self, ev):
        if self._start is None:
            return
        self._end = ev.pos()
        r = self._norm()
        if r.width() < 4 or r.height() < 4:
            self._start = None       # 太小当作误点，继续选
            self.update()
            return
        dpr = self._pix.devicePixelRatio() or 1.0
        crop = QRect(int(r.x() * dpr), int(r.y() * dpr),
                     int(r.width() * dpr), int(r.height() * dpr))
        self._result.append(self._pix.copy(crop))
        self._loop.quit()

    def _norm(self):
        x1, y1 = self._start.x(), self._start.y()
        x2, y2 = self._end.x(), self._end.y()
        return QRect(min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1))

    def paintEvent(self, ev):
        p = QPainter(self)
        p.drawPixmap(0, 0, self._pix)
        p.fillRect(self.rect(), QColor(0, 0, 0, 70))     # 暗罩
        if self._start is not None and self._end is not None:
            r = self._norm()
            # 亮出选区：把原画面再画一遍（不带罩），加边框
            p.drawPixmap(0, 0, self._pix, r.x(), r.y(), r.width(), r.height())
            p.setPen(QPen(QColor('#2f7de0'), 2))
            p.setBrush(Qt.NoBrush)
            p.drawRect(r)
            tip = '%d × %d  松开完成，Esc 取消' % (r.width(), r.height())
            p.fillRect(QRect(r.x(), max(0, r.y() - 22), r.width(), 20),
                       QColor('#2f7de0'))
            p.setPen(QPen(QColor('#ffffff')))
            p.drawText(QRect(r.x(), max(0, r.y() - 22), r.width(), 20),
                       Qt.AlignCenter, tip)


def capture_region():
    """全屏区域截图：每个屏幕一个覆盖层，用户拖拽选区后返回 QPixmap；取消返回 None。"""
    try:
        screens = QGuiApplication.screens()
        if not screens:
            return None
        result = []
        loop = QEventLoop()
        overlays = []
        for scr in screens:
            pix = scr.grabWindow(0)
            if pix.isNull():
                continue
            ov = _CaptureOverlay(pix, result, loop)
            ov.setGeometry(scr.geometry())
            overlays.append(ov)
        if not overlays:
            return None
        for ov in overlays:
            ov.show()
        loop.exec_()
        for ov in overlays:
            ov.close()
            ov.deleteLater()
        return result[0] if result else None
    except Exception:
        import traceback
        traceback.print_exc()
        return None


# ---------------- 放大查看器 ----------------

class _ZoomView(QGraphicsView):
    """支持滚轮缩放的画布（不用 eventFilter：样式表的 DynamicPropertyChange
    事件打进自定义 eventFilter 会触发 PyQt5 崩溃，改用子类事件最稳）。"""

    zoom_requested = pyqtSignal(int)          # +1 放大 / -1 缩小

    def wheelEvent(self, ev):
        d = ev.angleDelta().y()
        if d:
            self.zoom_requested.emit(1 if d > 0 else -1)
            ev.accept()
            return
        super().wheelEvent(ev)


class ImageViewerDialog(QDialog):
    """图纸放大查看：滚轮 / 按钮缩放，拖动平移，Esc 关闭。"""

    def __init__(self, parent, pix, title='图纸', path=''):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(860, 640)
        self._path = path or ''
        self._pix = pix
        self._zoom = 1.0

        self.view = _ZoomView(self)
        self.view.zoom_requested.connect(self._on_zoom)
        self.view.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform)
        self.view.setDragMode(QGraphicsView.ScrollHandDrag)
        self.view.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.scene = QGraphicsScene(self)
        self._item = None
        if pix is not None and not pix.isNull():
            self._item = QGraphicsPixmapItem(pix)
            self.scene.setSceneRect(QRectF(pix.rect()))
            self.scene.addItem(self._item)
        else:
            self.scene.addText('（无图纸）')
        self.view.setScene(self.scene)

        v = QVBoxLayout(self)
        v.setContentsMargins(6, 6, 6, 6)
        v.addWidget(self.view, 1)
        hb = QHBoxLayout()
        tip = QLabel('滚轮缩放 · 拖动平移 · Esc 关闭' +
                     ('' if not self._path else '\n原图：%s' % self._path))
        hb.addWidget(tip, 1)
        for text, slot in (('适应窗口', self.fit_window), ('1:1', self.zoom_one),
                           ('＋', self.zoom_in), ('－', self.zoom_out)):
            b = QPushButton(text)
            b.clicked.connect(slot)
            hb.addWidget(b)
        if self._path and os.path.exists(self._path):
            ob = QPushButton('打开原图')
            ob.setToolTip('用系统看图程序打开原文件')
            ob.clicked.connect(self.open_external)
            hb.addWidget(ob)
        v.addLayout(hb)
        self.fit_window()

    def _on_zoom(self, direction):
        self._zoom_by(1.25 if direction > 0 else 0.8)

    def _zoom_by(self, f):
        self._zoom = max(0.05, min(20, self._zoom * f))
        self.view.scale(f, f)

    def zoom_in(self):
        self._zoom_by(1.25)

    def zoom_out(self):
        self._zoom_by(0.8)

    def zoom_one(self):
        self.view.resetTransform()
        self._zoom = 1.0

    def fit_window(self):
        if self._item is None:
            return
        self.view.resetTransform()
        self._zoom = 1.0
        sr = self.scene.sceneRect()
        vr = self.view.viewport().rect()
        if sr.width() <= 0 or sr.height() <= 0 or vr.width() <= 0:
            return
        f = min(vr.width() / sr.width(), vr.height() / sr.height())
        if 0 < f < 1 or f > 1:
            self.view.scale(f, f)
            self._zoom = f

    def keyPressEvent(self, ev):
        if ev.key() == Qt.Key_Escape:
            self.close()
            return
        super().keyPressEvent(ev)

    def open_external(self):
        try:
            if self._path and os.path.exists(self._path):
                os.startfile(self._path)
        except Exception:
            pass


# ---------------- 右侧图纸预览面板 ----------------

class _ClickLabel(QLabel):
    """可点击的图片标签（mousePressEvent 子类化，不用 eventFilter）。"""

    clicked = pyqtSignal()

    def mousePressEvent(self, ev):
        if ev.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(ev)


class PreviewPanel(QFrame):
    """工作区右侧：显示当前行图纸缩略图，点击放大。"""

    open_requested = pyqtSignal()
    hide_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet(
            'PreviewPanel { background:#f6f8fb; border:1px solid #d8e0ea; }'
            'QLabel#ppTitle { color:#24507a; font-weight:600; }'
            'QLabel#ppCap { color:#5c738c; font-size:8.5pt; }'
            'QPushButton { border:1px solid #c9d6e3; background:#fff; color:#24507a;'
            ' font-size:8.5pt; padding:1px 8px; border-radius:2px; }'
            'QPushButton:hover { background:#e8f0fa; }')
        v = QVBoxLayout(self)
        v.setContentsMargins(6, 5, 6, 6)
        v.setSpacing(4)
        head = QHBoxLayout()
        t = QLabel('图纸预览')
        t.setObjectName('ppTitle')
        head.addWidget(t)
        head.addStretch(1)
        big = QPushButton('放大')
        big.setToolTip('放大查看图纸（滚轮缩放 / 拖动平移）')
        big.clicked.connect(self.open_requested.emit)
        head.addWidget(big)
        x = QPushButton('×')
        x.setFixedWidth(24)
        x.setToolTip('收起图纸预览（功能区「编辑 → 图纸预览」可再打开）')
        x.clicked.connect(self.hide_requested.emit)
        head.addWidget(x)
        v.addLayout(head)
        self.lbl_img = _ClickLabel('点击带图纸的行，这里显示该行图纸')
        self.lbl_img.setAlignment(Qt.AlignCenter)
        self.lbl_img.setWordWrap(True)
        self.lbl_img.setMinimumSize(150, 180)
        self.lbl_img.setStyleSheet('QLabel { color:#9aa8b8; background:#ffffff;'
                                   ' border:1px dashed #c9d6e3; }')
        self.lbl_img.setCursor(Qt.PointingHandCursor)
        self.lbl_img.clicked.connect(self.open_requested.emit)
        v.addWidget(self.lbl_img, 1)
        self.lbl_cap = QLabel('')
        self.lbl_cap.setObjectName('ppCap')
        self.lbl_cap.setWordWrap(True)
        v.addWidget(self.lbl_cap)
        self._pix = None
        self._path = ''

    def set_image(self, pix, caption='', path=''):
        self._pix = pix
        self._path = path or ''
        if pix is None or pix.isNull():
            self.lbl_img.setPixmap(QPixmap())
            self.lbl_img.setText('该行还没有图纸\n\n右键「图纸」列：\n截图 / 选图 / 粘贴')
            self.lbl_cap.setText('')
            return
        self.lbl_img.setText('')
        avail = QSize(max(80, self.lbl_img.width() - 12),
                      max(60, self.lbl_img.height() - 12))
        self.lbl_img.setPixmap(pix.scaled(avail, Qt.KeepAspectRatio,
                                          Qt.SmoothTransformation))
        self.lbl_cap.setText(caption)

    def resizeEvent(self, ev):
        super().resizeEvent(ev)
        if self._pix is not None and not self._pix.isNull():
            self.set_image(self._pix,
                           self.lbl_cap.text(), self._path)
