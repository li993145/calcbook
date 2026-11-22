# -*- coding: utf-8 -*-
"""公式语法配色规则（公式辅助框与表格“计算公式”列共用同一套规则）。

规则（仿兔兔计算书配色习惯）：
  数字      -> 红色
  运算符    -> 黑色（加粗）
  备注/汉字 -> 淡绿色
  方括号    -> 深绿色
  上标 ² ³  -> 蓝色加粗

表格里没有富文本，所以额外提供 compute_runs()：
把一行公式拆成 [(起始位置, 长度, 类型)]，由单元格绘制委托按同一配色着色，
保证“表格里看到的颜色”和“辅助框里看到的颜色”完全一致。
"""

import re

from PyQt5.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor, QFont

# 各类记号的颜色（表格单元格绘制也取这里，两边永远一致）
FCOLOR = {
    'number': '#C00000',
    'op': '#000000',
    'comment': '#3a9d3a',
    'bracket': '#227c22',
    'super': '#1f5fbf',
    'const': '#6a3d9a',      # @常量：紫色，一眼能看出引用了哪个常量
}

_NUM_RE = re.compile(r'\d+(?:\.\d+)?')
_OP_RE = re.compile(r'[+\-*/^()=×÷,，]+')
_BRACKET_RE = re.compile(r'[\[{【][^\]}]*[\]}】]')
_CJK_RE = re.compile(r'[\u4e00-\u9fff]+')
_SUPER_RE = re.compile(r'[²³¹⁰]')
# @常量（含 @Φ8 这种带希腊字母的），遇到运算符/括号即结束
_CONST_RE = re.compile(r'@[^\s+\-*/^()=×÷,，\[\]{}【】]+')


def compute_runs(text):
    """把一行公式拆成 [(start, length, kind), ...]，kind 见 FCOLOR。

    着色优先级与 highlightBlock 完全一致（后应用的类型覆盖前面的）：
    @常量 → 数字 → 运算符 → 方括号 → 汉字 → 上标。
    """
    text = text or ''
    n = len(text)
    if not n:
        return []
    kinds = [None] * n
    for m in _CONST_RE.finditer(text):
        for i in range(m.start(), m.end()):
            kinds[i] = 'const'
    for m in _NUM_RE.finditer(text):
        for i in range(m.start(), m.end()):
            if kinds[i] == 'const':
                continue
            kinds[i] = 'number'
    for m in _OP_RE.finditer(text):
        for i in range(m.start(), m.end()):
            if kinds[i] == 'const':
                continue
            kinds[i] = 'op'
    for m in _BRACKET_RE.finditer(text):
        for i in range(m.start(), m.end()):
            kinds[i] = 'bracket'
    for m in _CJK_RE.finditer(text):
        for i in range(m.start(), m.end()):
            if kinds[i] in ('bracket', 'comment', 'const'):
                continue
            kinds[i] = 'comment'
    for m in _SUPER_RE.finditer(text):
        for i in range(m.start(), m.end()):
            kinds[i] = 'super'
    runs = []
    start = 0
    for i in range(1, n + 1):
        if i == n or kinds[i] != kinds[start]:
            if kinds[start]:
                runs.append((start, i - start, kinds[start]))
            start = i
    return runs


class FormulaHighlighter(QSyntaxHighlighter):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.fmt_number = self._fmt(FCOLOR['number'])
        self.fmt_op = self._fmt(FCOLOR['op'], bold=True)
        self.fmt_comment = self._fmt(FCOLOR['comment'])
        self.fmt_bracket = self._fmt(FCOLOR['bracket'], bold=True)
        self.fmt_super = self._fmt(FCOLOR['super'], bold=True)
        self.fmt_const = self._fmt(FCOLOR['const'])
        self._fmts = {
            'number': self.fmt_number, 'op': self.fmt_op,
            'bracket': self.fmt_bracket, 'comment': self.fmt_comment,
            'super': self.fmt_super, 'const': self.fmt_const,
        }

    @staticmethod
    def _fmt(color, bold=False):
        f = QTextCharFormat()
        f.setForeground(QColor(color))
        if bold:
            f.setFontWeight(QFont.Bold)
        return f

    def highlightBlock(self, text):
        for start, length, kind in compute_runs(text):
            fmt = self._fmts.get(kind)
            if fmt is not None:
                self.setFormat(start, length, fmt)
