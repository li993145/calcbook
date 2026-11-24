# -*- coding: utf-8 -*-
"""项目选项 / 选项配置 对话框。

按成熟工程量软件的界面习惯提供两个对话框：

* ``ProjectOptionsDialog``  —— 项目选项：工程名称、编制单位、编制人、创建时间、
  工程量精度，以及工程密码的设置 / 修改。
* ``OptionConfigDialog``   —— 选项配置：回车键效果、显示配置、注释设置、定时、
  分部默认计算式个数、快速启动、特殊符号，以及“编辑辅助功能”里的十余个开关。

两个对话框都只负责“读出/写回”一个普通 dict，真正的生效逻辑在主窗口里，
方便以后再加新的选项而不必改动对话框。
"""

import json

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QPalette
from PyQt5.QtWidgets import (
    QCheckBox, QColorDialog, QComboBox, QDialog, QDialogButtonBox,
    QFormLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QRadioButton, QSpinBox, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget, QHeaderView, QAbstractItemView
)


# ---------------------------------------------------------------- 选项默认值
DEFAULT_OPTIONS = {
    # 回车键效果设置
    'enter_dir': 'down',            # down=向下换行/跳下一行；right=向右移动一格
    # 显示配置
    'font_size': 9,                 # 表格字体大小
    'wide_bracket': False,          # 加宽括号
    # 注释设置
    'note_color': '#008000',        # 注释颜色
    'var_autovalue': False,         # 变量自动值（总开关）
    'var_values': '{}',             # 变量自动值表：{"@层高": "3.6"}
    # 定时
    'autosave_enabled': True,
    'autosave_min': 3,
    'autosave_dir': '',
    # 分部
    'default_rows': 15,             # 新建分部时预置的计算式行数（V3.7.4 从 1 提到 15）
    # 快速启动
    'autoload_last': False,         # 自动打开上次编辑工程
    'save_geometry': True,          # 记住上次关闭窗口大小及位置
    # 编辑辅助功能
    'auto_unit': True,              # 分部单位自动应用到计算式
    'mul_star': True,               # * 号自动翻译为 ×
    'mul_div': True,                # / 号自动翻译为 ÷
    'mul_plus': False,              # + 号自动翻译为 ＋
    'mul_minus': False,             # - 号自动翻译为 －
    'bracket_gbq': False,           # [] 自动翻译为 {}（兼容 GBQ）
    'show_helper': True,            # 显示辅助计算书公式窗口
    'ribbon_expanded': True,        # 默认展开 Ribbon 菜单
    'paste_gcl': True,              # 粘贴时自动识别 ◇注释(GCL)
    'show_ref_code': True,          # 自动显示引用代码工程量注释
    'level_colors': True,           # 分部按层级显示不同颜色
}


def merged_options(raw):
    """把读出来的选项和默认值合并（缺项补默认，多余项丢弃）。"""
    out = dict(DEFAULT_OPTIONS)
    if isinstance(raw, dict):
        for k, v in raw.items():
            if k in out:
                out[k] = v
    # 类型兜底（设置文件被手工改坏时也不崩）
    out['font_size'] = _as_int(out['font_size'], 9, 7, 20)
    out['autosave_min'] = _as_int(out['autosave_min'], 3, 1, 240)
    out['default_rows'] = _as_int(out['default_rows'], 15, 0, 200)
    out['enter_dir'] = 'right' if str(out['enter_dir']) == 'right' else 'down'
    if not str(out['note_color']).startswith('#') or len(str(out['note_color'])) != 7:
        out['note_color'] = '#008000'
    if not isinstance(out['var_values'], str):
        out['var_values'] = '{}'
    for key in ('wide_bracket', 'var_autovalue', 'autosave_enabled', 'autoload_last',
                'save_geometry', 'auto_unit', 'mul_star', 'mul_div', 'mul_plus',
                'mul_minus', 'bracket_gbq', 'show_helper', 'ribbon_expanded',
                'paste_gcl', 'show_ref_code', 'level_colors'):
        out[key] = bool(out[key])
    return out


def _as_int(v, default, lo, hi):
    try:
        v = int(v)
    except (TypeError, ValueError):
        return default
    return max(lo, min(hi, v))


# ---------------------------------------------------------------- 项目选项
class ProjectOptionsDialog(QDialog):
    """项目选项：工程信息 + 数据配置（工程量精度、密码）。"""

    def __init__(self, project, parent=None):
        super().__init__(parent)
        self.setWindowTitle('项目选项')
        self.setMinimumWidth(520)
        self.project = project
        info = project.info or {}

        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(8)

        # —— 工程信息 ——
        g1 = QGroupBox('工程信息')
        f1 = QFormLayout(g1)
        f1.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.ed_name = QLineEdit(project.name or '')
        self.ed_name.setPlaceholderText('工程名称')
        f1.addRow('工程名称：', self.ed_name)

        self.ed_org = QLineEdit(str(info.get('org', '') or ''))
        self.ed_org.setPlaceholderText('编制单位')
        f1.addRow('编制单位：', self.ed_org)

        row = QHBoxLayout()
        self.ed_author = QLineEdit(str(info.get('author', '') or ''))
        self.ed_author.setPlaceholderText('编制人')
        row.addWidget(self.ed_author, 1)
        self.lb_created = QLabel(str(info.get('created', '') or project.created or ''))
        self.lb_created.setStyleSheet('color:#444;')
        row.addWidget(QLabel('创建时间：'))
        row.addWidget(self.lb_created)
        f1.addRow('编制人：', row)
        lay.addWidget(g1)

        # —— 数据配置 ——
        g2 = QGroupBox('数据配置')
        f2 = QFormLayout(g2)
        f2.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.sp_precision = QSpinBox()
        self.sp_precision.setRange(0, 6)
        self.sp_precision.setValue(int(getattr(parent, 'precision', 3) or 3))
        self.sp_precision.setSuffix(' 位')
        f2.addRow('工程量精度：', self.sp_precision)

        row2 = QHBoxLayout()
        row2.addStretch(1)
        self.btn_pwd = QPushButton('设置 / 更改密码…')
        self.btn_pwd.clicked.connect(self.edit_password)
        row2.addWidget(self.btn_pwd)
        f2.addRow('', row2)

        self.lb_pwd_state = QLabel('')
        self.lb_pwd_state.setStyleSheet('color:#666;')
        f2.addRow('', self.lb_pwd_state)
        lay.addWidget(g2)
        self._refresh_pwd_state()

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    # ---------------- 密码 ----------------
    def _refresh_pwd_state(self):
        has = bool((self.project.info or {}).get('password'))
        self.lb_pwd_state.setText('当前工程已设置密码。' if has else '当前工程未设置密码。')

    def edit_password(self):
        dlg = QDialog(self)
        dlg.setWindowTitle('设置 / 更改密码')
        dlg.setMinimumWidth(360)
        f = QFormLayout(dlg)
        ed1 = QLineEdit()
        ed1.setEchoMode(QLineEdit.Password)
        ed1.setPlaceholderText('留空表示清除密码')
        f.addRow('新密码：', ed1)
        ed2 = QLineEdit()
        ed2.setEchoMode(QLineEdit.Password)
        ed2.setPlaceholderText('再次输入新密码')
        f.addRow('确认密码：', ed2)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        f.addRow(bb)
        if dlg.exec_() != QDialog.Accepted:
            return
        p1 = ed1.text()
        p2 = ed2.text()
        if p1 != p2:
            QMessageBoxWarning(self, '两次输入的密码不一致。')
            return
        if not p1:
            (self.project.info or {}).pop('password', None)
        else:
            self.project.info['password'] = hash_password(p1)
        self._refresh_pwd_state()

    def apply(self):
        """把界面上的值写回工程对象，返回 True 表示有改动。"""
        changed = False
        name = self.ed_name.text().strip()
        if name and name != self.project.name:
            self.project.name = name
            changed = True
        info = self.project.info if self.project.info is not None else {}
        self.project.info = info
        if info.get('org', '') != self.ed_org.text().strip():
            info['org'] = self.ed_org.text().strip()
            changed = True
        if info.get('author', '') != self.ed_author.text().strip():
            info['author'] = self.ed_author.text().strip()
            changed = True
        if not info.get('created'):
            info['created'] = self.lb_created.text()
        return changed


def QMessageBoxWarning(parent, text):
    """小工具：避免顶部再引一次 QMessageBox 的名字冲突。"""
    from PyQt5.QtWidgets import QMessageBox
    QMessageBox.warning(parent, '提示', text)


def hash_password(text):
    """工程密码只存哈希，不存明文。"""
    import hashlib
    return hashlib.sha256(('calcbook:' + text).encode('utf-8')).hexdigest()


def check_password(text, digest):
    return bool(digest) and hash_password(text) == digest


# ---------------------------------------------------------------- 选项配置
class OptionConfigDialog(QDialog):
    """选项配置（按截图布局：左侧分组设置，右侧“编辑辅助功能”）。"""

    def __init__(self, options, parent=None):
        super().__init__(parent)
        self.setWindowTitle('选项配置')
        self.setMinimumSize(760, 520)
        self.opt = dict(options or {})
        self._custom_labels = list(getattr(parent, 'annotation_labels', []) or [])

        root = QHBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(10)
        left = QVBoxLayout()
        left.setSpacing(6)
        root.addLayout(left, 3)
        right = QVBoxLayout()
        right.setSpacing(6)
        root.addLayout(right, 2)

        # ——— 回车键效果设置 ———
        gb = QGroupBox('回车键效果设置')
        hb = QHBoxLayout(gb)
        self.rb_down = QRadioButton('向下')
        self.rb_right = QRadioButton('向右')
        if self.opt.get('enter_dir') == 'right':
            self.rb_right.setChecked(True)
        else:
            self.rb_down.setChecked(True)
        hb.addWidget(self.rb_down)
        hb.addWidget(self.rb_right)
        hb.addStretch(1)
        left.addWidget(gb)

        # ——— 显示配置 ———
        gb = QGroupBox('显示配置')
        fh = QHBoxLayout(gb)
        fh.addWidget(QLabel('字体大小'))
        self.cb_font = QComboBox()
        for i in range(7, 21):
            self.cb_font.addItem(str(i), i)
        try:
            self.cb_font.setCurrentIndex(self.cb_font.findData(int(self.opt.get('font_size', 9))))
        except Exception:
            self.cb_font.setCurrentIndex(2)
        fh.addWidget(self.cb_font)
        self.chk_wide = QCheckBox('加宽括号')
        self.chk_wide.setChecked(bool(self.opt.get('wide_bracket')))
        fh.addWidget(self.chk_wide)
        fh.addStretch(1)
        left.addWidget(gb)

        # ——— 注释设置 ———
        gb = QGroupBox('注释设置')
        fh = QHBoxLayout(gb)
        fh.addWidget(QLabel('注释颜色'))
        self.btn_note_color = QPushButton('点击设置')
        self.btn_note_color.setFixedWidth(96)
        self.btn_note_color.clicked.connect(self.pick_note_color)
        fh.addWidget(self.btn_note_color)
        fh.addSpacing(16)
        fh.addWidget(QLabel('变量自动值'))
        self.btn_var = QPushButton('点击设置')
        self.btn_var.setFixedWidth(96)
        self.btn_var.clicked.connect(self.edit_var_values)
        fh.addWidget(self.btn_var)
        fh.addStretch(1)
        left.addWidget(gb)

        # ——— 定时 ———
        gb = QGroupBox('定时')
        fh = QHBoxLayout(gb)
        fh.addWidget(QLabel('保存周期'))
        self.sp_save = QSpinBox()
        self.sp_save.setRange(1, 240)
        self.sp_save.setValue(int(self.opt.get('autosave_min', 3) or 3))
        self.sp_save.setSuffix(' (分钟)')
        fh.addWidget(self.sp_save)
        fh.addStretch(1)
        left.addWidget(gb)

        # ——— 分部 ———
        gb = QGroupBox('分部')
        fh = QHBoxLayout(gb)
        fh.addWidget(QLabel('默认计算式个数'))
        self.sp_rows = QSpinBox()
        self.sp_rows.setRange(0, 50)
        self.sp_rows.setValue(int(self.opt.get('default_rows', 1) or 0))
        fh.addWidget(self.sp_rows)
        fh.addStretch(1)
        left.addWidget(gb)

        # ——— 快速启动 ———
        gb = QGroupBox('快速启动')
        fh = QHBoxLayout(gb)
        self.chk_autoload = QCheckBox('自动打开上次编辑工程')
        self.chk_autoload.setChecked(bool(self.opt.get('autoload_last')))
        fh.addWidget(self.chk_autoload)
        self.chk_geom = QCheckBox('记住上次关闭窗口大小及位置')
        self.chk_geom.setChecked(bool(self.opt.get('save_geometry')))
        fh.addWidget(self.chk_geom)
        fh.addStretch(1)
        left.addWidget(gb)

        # ——— 特殊符号 ———
        gb = QGroupBox('特殊符号（变化在软件重启后生效）')
        fh = QHBoxLayout(gb)
        self.btn_sym = QPushButton('编辑特殊字符')
        self.btn_sym.clicked.connect(self.edit_symbols)
        fh.addWidget(self.btn_sym)
        self.btn_label = QPushButton('编辑注释标签')
        self.btn_label.clicked.connect(self.edit_labels)
        fh.addWidget(self.btn_label)
        fh.addStretch(1)
        left.addWidget(gb)
        left.addStretch(1)

        # ——— 编辑辅助功能 ———
        gb = QGroupBox('编辑辅助功能')
        gv = QVBoxLayout(gb)
        gv.setSpacing(2)
        self.helpers = {}
        items = [
            ('auto_unit', '分部单位自动应用到计算式'),
            ('mul_star', '* 号自动翻译为 ×'),
            ('mul_div', '/ 号自动翻译为 ÷'),
            ('mul_plus', '+ 号自动翻译为 ＋'),
            ('mul_minus', '- 号自动翻译为 －'),
            ('bracket_gbq', '[] 自动翻译为 {}（兼容GBQ）'),
            ('show_helper', '显示辅助计算书公式窗口'),
            ('ribbon_expanded', '默认展开Ribbon菜单'),
            ('paste_gcl', '粘贴时自动识别◇注释(GCL)'),
            ('show_ref_code', '自动显示引用代码工程量注释'),
            ('level_colors', '分部按层级显示不同颜色'),
        ]
        for key, label in items:
            cb = QCheckBox(label)
            cb.setChecked(bool(self.opt.get(key)))
            gv.addWidget(cb)
            self.helpers[key] = cb
        right.addWidget(gb)
        right.addStretch(1)

        self._note_color = str(self.opt.get('note_color', '#008000'))
        self._apply_color_btn()

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        root.addWidget(bb)

    # ---------------- 各设置项 ----------------
    def _apply_color_btn(self):
        c = QColor(self._note_color)
        if not c.isValid():
            c = QColor('#008000')
        self.btn_note_color.setStyleSheet(
            'QPushButton { background:%s; color:#ffffff; border:1px solid #999; '
            'padding:3px 8px; }' % c.name())

    def pick_note_color(self):
        c = QColorDialog.getColor(QColor(self._note_color), self, '选择注释颜色')
        if c.isValid():
            self._note_color = c.name()
            self._apply_color_btn()

    def edit_var_values(self):
        dlg = VarValueDialog(self.opt.get('var_values'), self)
        if dlg.exec_() == QDialog.Accepted:
            self.opt['var_values'] = dlg.json_text()

    def edit_symbols(self):
        dlg = CustomSymbolDialog(getattr(self.parent(), 'custom_symbols', []) or [], self)
        if dlg.exec_() == QDialog.Accepted:
            if self.parent() is not None:
                try:
                    self.parent().custom_symbols = dlg.symbols()
                except Exception:
                    pass

    def edit_labels(self):
        dlg = LabelDialog(self._custom_labels, self)
        if dlg.exec_() == QDialog.Accepted:
            self._custom_labels = dlg.labels()

    # ---------------- 结果 ----------------
    def result_options(self):
        out = dict(self.opt)
        out['enter_dir'] = 'right' if self.rb_right.isChecked() else 'down'
        out['font_size'] = int(self.cb_font.currentData() or 9)
        out['wide_bracket'] = self.chk_wide.isChecked()
        out['note_color'] = self._note_color
        out['autosave_min'] = int(self.sp_save.value())
        out['default_rows'] = int(self.sp_rows.value())
        out['autoload_last'] = self.chk_autoload.isChecked()
        out['save_geometry'] = self.chk_geom.isChecked()
        for key, cb in self.helpers.items():
            out[key] = cb.isChecked()
        return out

    def labels(self):
        return list(self._custom_labels)


# ---------------------------------------------------------------- 变量自动值
class VarValueDialog(QDialog):
    """变量自动值：给变量（常量）预设一个默认取值。"""

    def __init__(self, json_text, parent=None):
        super().__init__(parent)
        self.setWindowTitle('变量自动值')
        self.setMinimumSize(460, 320)
        lay = QVBoxLayout(self)
        tip = QLabel('公式里引用到未定义的变量时，自动套用这里设置的取值。\n'
                     '格式：变量名 = 取值，例如  @层高 = 3.6')
        tip.setWordWrap(True)
        tip.setStyleSheet('color:#666;')
        lay.addWidget(tip)

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(['变量名', '取值'])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        lay.addWidget(self.table, 1)

        data = {}
        try:
            data = json.loads(json_text or '{}')
            if not isinstance(data, dict):
                data = {}
        except Exception:
            data = {}
        for k in sorted(data):
            self._add_row(str(k), str(data[k]))

        row = QHBoxLayout()
        b_add = QPushButton('添加')
        b_add.clicked.connect(lambda: self._add_row('', ''))
        b_del = QPushButton('删除')
        b_del.clicked.connect(self._del_row)
        row.addWidget(b_add)
        row.addWidget(b_del)
        row.addStretch(1)
        lay.addLayout(row)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _add_row(self, name, value):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(name))
        self.table.setItem(r, 1, QTableWidgetItem(value))

    def _del_row(self):
        r = self.table.currentRow()
        if r >= 0:
            self.table.removeRow(r)

    def json_text(self):
        data = {}
        for r in range(self.table.rowCount()):
            it0 = self.table.item(r, 0)
            it1 = self.table.item(r, 1)
            if it0 is None:
                continue
            k = it0.text().strip()
            if not k:
                continue
            v = it1.text().strip() if it1 is not None else ''
            data[k] = v
        return json.dumps(data, ensure_ascii=False)


# ---------------------------------------------------------------- 自定义特殊字符
class CustomSymbolDialog(QDialog):
    """编辑特殊字符：自定义一组符号，显示在“特殊符号”面板最前面。"""

    def __init__(self, symbols, parent=None):
        super().__init__(parent)
        self.setWindowTitle('编辑特殊字符')
        self.setMinimumSize(460, 320)
        lay = QVBoxLayout(self)
        tip = QLabel('自定义的特殊字符会显示在“特殊符号”面板的第一组，点一下即可插入公式。')
        tip.setWordWrap(True)
        tip.setStyleSheet('color:#666;')
        lay.addWidget(tip)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(['符号', '名称/说明', '快捷码'])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        lay.addWidget(self.table, 1)
        for s in symbols or []:
            self._add_row(s.get('ch', ''), s.get('name', ''), s.get('code', ''))

        row = QHBoxLayout()
        b_add = QPushButton('添加')
        b_add.clicked.connect(lambda: self._add_row('', '', ''))
        b_del = QPushButton('删除')
        b_del.clicked.connect(self._del_row)
        row.addWidget(b_add)
        row.addWidget(b_del)
        row.addStretch(1)
        lay.addLayout(row)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _add_row(self, ch, name, code):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(ch))
        self.table.setItem(r, 1, QTableWidgetItem(name))
        self.table.setItem(r, 2, QTableWidgetItem(code))

    def _del_row(self):
        r = self.table.currentRow()
        if r >= 0:
            self.table.removeRow(r)

    def symbols(self):
        out = []
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 0)
            if it is None:
                continue
            ch = it.text()
            if not ch:
                continue
            n = self.table.item(r, 1)
            c = self.table.item(r, 2)
            out.append({'ch': ch,
                        'name': n.text() if n else ch,
                        'code': c.text() if c else ''})
        return out


# ---------------------------------------------------------------- 注释标签
class LabelDialog(QDialog):
    """编辑注释标签：注释的常用标记，点一下插入到备注/公式。"""

    def __init__(self, labels, parent=None):
        super().__init__(parent)
        self.setWindowTitle('编辑注释标签')
        self.setMinimumWidth(400)
        lay = QVBoxLayout(self)
        tip = QLabel('注释标签用于在“备注”里快速打标记，格式形如  [标签] 内容。')
        tip.setWordWrap(True)
        tip.setStyleSheet('color:#666;')
        lay.addWidget(tip)

        self.ed = QLineEdit()
        self.ed.setPlaceholderText('标签名称（多个用顿号或逗号分隔）')
        self.ed.setText('、'.join(labels or []))
        lay.addWidget(self.ed)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText('确定')
        bb.button(QDialogButtonBox.Cancel).setText('取消')
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def labels(self):
        raw = self.ed.text().replace(',', '、').replace('，', '、')
        out = []
        for x in raw.split('、'):
            x = x.strip()
            if x and x not in out:
                out.append(x)
        return out
