# -*- coding: utf-8 -*-
"""排序与筛选对话框（V2.10.0）。

纯逻辑函数（sort_key / filter_match）不依赖 Qt，便于自动化测试；
对话框只负责收集参数。
"""

import re

from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                             QComboBox, QLineEdit, QSpinBox, QCheckBox,
                             QRadioButton, QButtonGroup, QDialogButtonBox,
                             QPushButton)

# ---------------- 拖动填充（自动序号） ----------------

_TRAIL = re.compile(r'^(.*?)(\d+)$')


def _fmt_num_val(v):
    """数值 → 显示文本（整数不带小数点，小数最多 6 位）。"""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, int):
        return str(v)
    return '%g' % round(v, 6)


def fill_series(svals, offsets):
    """按 WPS 规则由源值列表生成偏移位置的填充值。

    offsets: 相对源块起点的整数序位（向下/右为正，向上/左为负）。
    规则（整列判断）：
      * 全部是数字 → 等差序列：步长=(末-首)/(n-1)；单格步长 1；多个相同数字步长 0（复制）
      * 全部是"前缀+尾号数字"（A-001 / TG102.103 / 深度3）→ 尾号等差，保留前导零与前缀
      * 其它 → 循环复制源值
    """
    svals = [str(s) for s in (svals or [''])]
    n = len(svals)
    if n == 0:
        return ['' for _ in offsets]

    def _step(nums):
        if n >= 2 and nums[-1] != nums[0]:
            return (nums[-1] - nums[0]) / (n - 1)
        if n == 1:
            return 1
        return 0

    nums = [parse_num(s) for s in svals]
    if all(x is not None for x in nums):
        step = _step(nums)
        start = nums[0]
        return [_fmt_num_val(start + i * step) for i in offsets]

    tails = [_TRAIL.match(s) for s in svals]
    if all(t is not None for t in tails):
        prefixes = [t.group(1) for t in tails]
        if len(set(prefixes)) == 1:
            bs = [int(t.group(2)) for t in tails]
            step = _step(bs)
            width = max(len(t.group(2)) for t in tails)
            out = []
            for i in offsets:
                v = int(round(bs[0] + i * step))
                if v < 0:
                    v = 0
                out.append(prefixes[0] + ('%0*d' % (width, v) if width else str(v)))
            return out

    return [svals[i % n] for i in offsets]


# ---------------- 排序 ----------------

SORT_MODES = ['按数字', '按文本', '按填充颜色', '按正则提取']

# 数字容差：浮点比较用
_EPS = 1e-12


def parse_num(s):
    """把文本解析为数字（支持千分位、百分号、全角数字、单位前缀如 ≈、约）。失败返回 None。"""
    if s is None:
        return None
    t = str(s).strip().replace(',', '').replace('，', '')
    # 全角数字/字母/标点先转半角，避免 １２３ 解析失败
    t = t.translate(str.maketrans(
        '０１２３４５６７８９．－＋％',
        '0123456789.-+%'))
    t = re.sub(r'^[≈约±+]', '', t)
    neg = t.startswith('-')
    t = t.lstrip('+-')
    m = re.match(r'^(\d+(?:\.\d+)?)\s*%$', t)
    if m:
        try:
            return (-1 if neg else 1) * float(m.group(1)) / 100.0
        except ValueError:
            return None
    m = re.match(r'^(\d+(?:\.\d+)?)', t)
    if not m:
        return None
    try:
        return (-1 if neg else 1) * float(m.group(1))
    except ValueError:
        return None


def regex_sort_key(text, pattern, prefer_num=True):
    """正则提取排序键：优先第 1 个捕获组，否则整个匹配；能转数字按数字。"""
    try:
        m = re.search(pattern, str(text))
    except re.error:
        return (2, 0.0, '')
    if not m:
        return (2, 0.0, '')          # 不匹配的排最后
    s = m.group(1) if (m.lastindex or 0) >= 1 else m.group(0)
    if prefer_num:
        n = parse_num(s)
        if n is not None:
            return (0, n, '')
    return (1, 0.0, str(s))


def make_sort_key(mode, pattern=None, prefer_num=True):
    """生成排序 key 函数。返回值形如 (分组, 数字, 文本) 元组，可直接 sorted。

    分组：0=有效数字 1=纯文本 2=空/不匹配（恒排最后，与方向无关地聚在一起）
    """
    if mode == '按数字':
        def key(text):
            n = parse_num(text)
            if n is None:
                t = str(text).strip()
                return (2 if not t else 1, 0.0, t)
            return (0, n, '')
        return key
    if mode == '按文本':
        def key(text):
            t = str(text).strip()
            return (2 if not t else 1, 0.0, t)
        return key
    if mode == '按填充颜色':
        # color 值由调用方传入（返回 (key, color) 的包装见 sort_rows）
        return None
    if mode == '按正则提取':
        def key(text):
            return regex_sort_key(text, pattern or '', prefer_num)
        return key
    return key


def sort_rows_with_colors(items, mode, pattern=None, reverse=False, prefer_num=True):
    """items: [(text, color), ...] → 返回排序后的下标列表（对原列表的重排顺序）。

    颜色模式：无填充排最前，其余按色号排序；升/降序只反转有填充的部分内
    部顺序与其它模式一致。
    """
    if mode == '按填充颜色':
        # 稳定分组：无填充(0) → 有填充按颜色值(1)
        keyed = []
        for i, (text, color) in enumerate(items):
            c = str(color or '')
            keyed.append(((0, '', '') if not c else (1, c, ''),
                          (0, 0.0, ''), i))
        keyed.sort(key=lambda x: (x[0], x[1], x[2]))
        order = [k[2] for k in keyed]
        if reverse:
            # 反转只作用于"有填充"块，无填充仍排最前
            head = [i for i in order if not str(items[i][1] or '')]
            tail = [i for i in order if str(items[i][1] or '')]
            return head + tail[::-1]
        return order
    keyfn = make_sort_key(mode, pattern, prefer_num)
    keyed = []
    for i, (text, _c) in enumerate(items):
        keyed.append((keyfn(text), i))
    keyed.sort(key=lambda x: (x[0], x[1]))
    order = [k[1] for k in keyed]
    if reverse:
        # 空值组（分组2）保持在最后：反转非空部分
        nonempty = [i for i in order if keyfn(items[i][0])[0] != 2]
        empty = [i for i in order if keyfn(items[i][0])[0] == 2]
        return nonempty[::-1] + empty
    return order


# ---------------- 筛选 ----------------

FILTER_MODES = ['等于', '不等于', '包含', '不包含', '开头是',
                '大于', '大于等于', '小于', '小于等于',
                '正则匹配', '按填充颜色是']


def filter_match(text, color, mode, value, color_value=None):
    """单元格是否通过筛选（True=显示）。"""
    t = str(text if text is not None else '')
    v = str(value if value is not None else '')
    if mode == '等于':
        return t.strip() == v.strip()
    if mode == '不等于':
        return t.strip() != v.strip()
    if mode == '包含':
        return v in t
    if mode == '不包含':
        return v not in t
    if mode == '开头是':
        return t.strip().startswith(v.strip())
    if mode in ('大于', '大于等于', '小于', '小于等于'):
        n1, n2 = parse_num(t), parse_num(v)
        if n1 is None:
            return False
        if n2 is None:
            return True
        if mode == '大于':
            return n1 > n2 + _EPS
        if mode == '大于等于':
            return n1 >= n2 - _EPS
        if mode == '小于':
            return n1 < n2 - _EPS
        return n1 <= n2 + _EPS
    if mode == '正则匹配':
        try:
            return re.search(v, t) is not None
        except re.error:
            return True            # 正则写错时不隐藏任何行
    if mode == '按填充颜色是':
        return str(color or '') == (color_value or '')
    return True


# ---------------- 对话框 ----------------

def _build_dlg(title, parent):
    from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel
    dlg = QDialog(parent)
    dlg.setWindowTitle(title)
    v = QVBoxLayout(dlg)
    return dlg, v


def _add_combo(v, label, items):
    from PyQt5.QtWidgets import QHBoxLayout, QLabel, QComboBox
    h = QHBoxLayout()
    lab = QLabel(label)
    h.addWidget(lab)
    cb = QComboBox()
    cb.addItems(items)
    h.addWidget(cb, 1)
    v.addLayout(h)
    return cb, lab


def _add_edit(v, label, default=''):
    from PyQt5.QtWidgets import QHBoxLayout, QLabel, QLineEdit
    h = QHBoxLayout()
    lab = QLabel(label)
    h.addWidget(lab)
    ed = QLineEdit(default)
    h.addWidget(ed, 1)
    v.addLayout(h)
    return ed, lab


class SortDialog:
    """排序参数收集器（配合 _build_dlg 使用，避免多重继承 QWidget 的坑）。"""

    @staticmethod
    def exec(parent, col_titles, cur_col=0):
        from PyQt5.QtWidgets import (QDialog, QRadioButton, QButtonGroup,
                                     QCheckBox, QDialogButtonBox, QHBoxLayout)
        dlg, v = _build_dlg('排序（当前分部）', parent)
        cb_col, _ = _add_combo(v, '排序列：', col_titles)
        cb_col.setCurrentIndex(max(0, min(cur_col, len(col_titles) - 1)))
        cb_mode, _ = _add_combo(v, '依据：', SORT_MODES)
        ed_re, _ = _add_edit(v, '正则（按正则提取时使用）：')
        h = QHBoxLayout()
        h.addWidget(QLabel('方向：'))
        rb_asc = QRadioButton('升序')
        rb_desc = QRadioButton('降序')
        rb_asc.setChecked(True)
        grp = QButtonGroup(dlg)
        grp.addButton(rb_asc)
        grp.addButton(rb_desc)
        h.addWidget(rb_asc)
        h.addWidget(rb_desc)
        h.addStretch(1)
        v.addLayout(h)
        ck_regex_num = QCheckBox('正则结果转数字比较（勾选后 10 排在 9 后面）')
        ck_regex_num.setChecked(True)
        v.addWidget(ck_regex_num)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        v.addWidget(bb)

        def sync_mode(i):
            ed_re.setEnabled(SORT_MODES[i] == '按正则提取')
            ck_regex_num.setVisible(SORT_MODES[i] == '按正则提取')
        cb_mode.currentIndexChanged.connect(sync_mode)
        sync_mode(cb_mode.currentIndex())

        if dlg.exec_() != QDialog.Accepted:
            return None
        return {
            'col': cb_col.currentIndex(),
            'mode': cb_mode.currentText(),
            'pattern': ed_re.text().strip(),
            'reverse': rb_desc.isChecked(),
            'regex_num': ck_regex_num.isChecked(),
        }


class FilterDialog:
    """筛选参数收集器。colors: [(色号或'', 显示名), ...] 供颜色模式选择。"""

    @staticmethod
    def exec(parent, col_titles, cur_col=0, colors=None):
        from PyQt5.QtWidgets import (QDialog, QComboBox, QDialogButtonBox,
                                     QHBoxLayout, QLabel)
        dlg, v = _build_dlg('筛选（当前分部，只隐藏不删数据）', parent)
        cb_col, _ = _add_combo(v, '筛选列：', col_titles)
        cb_col.setCurrentIndex(max(0, min(cur_col, len(col_titles) - 1)))
        cb_mode, lab_mode = _add_combo(v, '条件：', FILTER_MODES)
        ed_val, lab_val = _add_edit(v, '内容：')
        h = QHBoxLayout()
        h.addWidget(QLabel('颜色：'))
        cb_color = QComboBox()
        for cv, name in (colors or [('', '无填充')]):
            cb_color.addItem(name, cv)
        h.addWidget(cb_color, 1)
        v.addLayout(h)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        # V3.4.1：筛选窗口内直接“清除筛选”，不用退回右键菜单找入口。
        # 从父窗口 _filter 判断当前是否正在筛选，不新增函数签名（兼容旧调用）。
        btn_clear = QPushButton('清除筛选')
        btn_clear.setEnabled(bool(getattr(parent, '_filter', None)))
        btn_clear.setToolTip('取消当前筛选条件，恢复显示全部行')

        def _do_clear():
            dlg.setProperty('clearFilter', True)
            dlg.accept()

        btn_clear.clicked.connect(_do_clear)
        bb.addButton(btn_clear, QDialogButtonBox.ActionRole)
        v.addWidget(bb)

        def sync(i):
            m = FILTER_MODES[i]
            is_color = m == '按填充颜色是'
            ed_val.setVisible(not is_color)
            lab_val.setVisible(not is_color)
            cb_color.setVisible(is_color)
            lab_mode.setText('颜色：' if is_color else '条件：')
        cb_mode.currentIndexChanged.connect(sync)
        sync(0)

        if dlg.exec_() != QDialog.Accepted:
            return None
        if dlg.property('clearFilter'):
            return {'clear': True}
        return {
            'col': cb_col.currentIndex(),
            'mode': cb_mode.currentText(),
            'value': ed_val.text(),
            'color_value': cb_color.currentData(),
        }
