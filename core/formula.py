# -*- coding: utf-8 -*-
"""
公式引擎 —— 工程量手算计算式解析与求值

设计目标：
  1. 全角数学符号自动转半角：＋ － × ÷ （ ） ＾ ． 等
  2. 注释剥离：[...] {...} 【...】 ｛...｝ 以及直接输入的汉字注释
  3. 常量：@汉字名，如 @层高
  4. 函数：ceil floor sqr sqrt round power pi radians log XJ Du max min abs sin cos tan
  5. 度分秒：34°56′78″
  6. 安全求值：使用 ast 白名单，绝不调用 eval

用法：
    eng = FormulaEngine(constants={'@层高': 3.6})
    value, cleaned, err = eng.evaluate('4×9.6×2.4×@层高[柱基预埋]', context={})
"""

import ast
import math
import operator
import re

# ------------------------- 全角 → 半角 -------------------------
_FULL2HALF = {
    '＋': '+', '－': '-', '−': '-', '—': '-', '–': '-',
    '×': '*', '✕': '*', '＊': '*', '·': '*', '': '*',
    '÷': '/', '／': '/', '∕': '/',
    '（': '(', '）': ')', '［': '[', '］': ']',
    '｛': '{', '｝': '}', '【': '[', '】': ']',
    '＾': '^', '＀': '^',
    '．': '.', '。': '.',
    '＝': '=', '，': ',', '、': ',',
    '％': '%', '：': ':', '；': ';',
    '＂': '"', '＇': "'",
}


def to_halfwidth(s: str) -> str:
    out = []
    for ch in s:
        if ch in _FULL2HALF:
            out.append(_FULL2HALF[ch])
        elif '\uff01' <= ch <= '\uff5e':  # 全角 ASCII 区
            out.append(chr(ord(ch) - 0xfee0))
        else:
            out.append(ch)
    return ''.join(out)


# ------------------------- 注释剥离 -------------------------
# 四套注释括号（含全角，to_halfwidth 后【】已变 []）
_COMMENT_BRACKETS = re.compile(r'\[[^\]]*\]|\{[^\}]*\}')
# 汉字/中文标点（作为行内注释）
_CJK = re.compile(r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]+')
# 常量占位符
_CONST_TOKEN = re.compile(r'__CONST_(\d+)__')
# 度分秒
_DMS = re.compile(r'(\d+(?:\.\d+)?)\s*°\s*(\d+(?:\.\d+)?)\s*[′\']\s*(\d+(?:\.\d+)?)\s*[″"]?')
# 钢筋级别罗马数字（ⅠⅡⅢⅣ）与直径/半径标记，公式中一律当注释忽略
_GRADE_CHARS = 'ⅠⅡⅢⅣⅤ'

# ------------------------- 单元格引用（V3.9） -------------------------
# 单元格地址 = 列字母 + 行号，如 A2、B12、AA3。
# 边界取得很严：前后不能紧贴字母 / 数字 / 下划线 / 汉字，
# 否则 C30混凝土、M10砂浆、HRB400 这类规格串会被误当成单元格引用。
CELL_REF_RE = re.compile(
    r'(?<![0-9A-Za-z_\u4e00-\u9fff])'
    r'([A-Za-z]{1,2})(\d{1,7})'
    r'(?![0-9A-Za-z_\u4e00-\u9fff])')


def find_cell_refs(text, valid=None):
    """找出文本里的单元格引用。

    valid 为合法地址集合（如 {'A1','A2','B1'}）；None 表示不校验。
    返回 [(起点, 终点, 地址), …]，地址统一大写。
    """
    out = []
    for m in CELL_REF_RE.finditer(str(text or '')):
        addr = m.group(1).upper() + m.group(2)
        if valid is not None and addr not in valid:
            continue
        out.append((m.start(), m.end(), addr))
    return out


# ------------------------- 开根号 √ → sqrt() -------------------------
def _sqrt_operand_end(s, i):
    """返回 √ 作用对象的结束下标（不含）。i 指向 √ 之后第一个位置。"""
    n = len(s)
    j = i
    while j < n and s[j] in ' \t':
        j += 1
    if j >= n:
        return j
    ch = s[j]
    if ch == '(':                      # √(3+4)
        depth = 0
        k = j
        while k < n:
            if s[k] == '(':
                depth += 1
            elif s[k] == ')':
                depth -= 1
                if depth == 0:
                    return k + 1
            k += 1
        return n
    if ch in '+-':                     # √-3
        return _sqrt_operand_end(s, j + 1)
    if ch.isdigit() or ch == '.':      # √9、√9.5
        k = j
        while k < n and (s[k].isdigit() or s[k] == '.'):
            k += 1
        return k
    if ch.isalpha() or ch == '_':      # √pi、√sqrt(9)、√__CONST_0__
        k = j
        while k < n and (s[k].isalnum() or s[k] == '_'):
            k += 1
        # 若紧跟括号，连括号一起作为对象：√ceil(3.2)
        m = k
        while m < n and s[m] in ' \t':
            m += 1
        if m < n and s[m] == '(':
            depth = 0
            while m < n:
                if s[m] == '(':
                    depth += 1
                elif s[m] == ')':
                    depth -= 1
                    if depth == 0:
                        return m + 1
                m += 1
            return n
        return k
    return j + 1


def convert_sqrt(s):
    """把 √ 转换成 sqrt(...)：√9 → sqrt(9)，2√9 → 2sqrt(9)，可嵌套。"""
    while '√' in s:
        i = s.index('√')
        end = _sqrt_operand_end(s, i + 1)
        operand = s[i + 1:end].strip()
        if not operand:                         # 孤立的 √，直接丢弃
            s = s[:i] + s[i + 1:]
            continue
        s = s[:i] + 'sqrt(' + operand + ')' + s[end:]
    return s


def strip_comments(expr: str) -> str:
    """剥离 [...]/{...} 注释与汉字注释，保留数学结构。"""
    s = expr
    # 多轮剥离，处理嵌套注释
    for _ in range(5):
        new = _COMMENT_BRACKETS.sub(' ', s)
        if new == s:
            break
        s = new
    return s


# ------------------------- 函数表 -------------------------
def _log(x, base=None):
    if base is None:
        return math.log(x)
    return math.log(x, base)


def _avg(*args):
    """求平均值：avg(1,2,3) = 2；无参数时返回 0。"""
    if not args:
        return 0
    return sum(args) / len(args)


# ---------------- 台体体积（V3.7.3） ----------------
def _tri_area(a, b, c):
    """海伦公式：三边求三角形面积；三边不构成三角形时报错。"""
    p = (a + b + c) / 2.0
    v = p * (p - a) * (p - b) * (p - c)
    if v <= 0:
        raise ValueError('三边无法构成三角形：%g, %g, %g' % (a, b, c))
    return math.sqrt(v)


def _frustum4(a, b, c, d, h):
    """四棱台体积（拟柱体公式）：下底长a×宽b、上底长c×宽d、高h。

    V = h/6 × [A·B + (A+a)·(B+b) + a·b]
    """
    return h / 6.0 * (a * b + (a + c) * (b + d) + c * d)


def _frustum3(a1, b1, c1, a2, b2, c2, h):
    """三棱台体积：下底三边 a1,b1,c1、上底三边 a2,b2,c2、高h。

    V = h/6 × (S下 + S上 + 4×S中)，三个三角形面积按海伦公式由边长算出。
    """
    s1 = _tri_area(a1, b1, c1)
    s2 = _tri_area(a2, b2, c2)
    s0 = _tri_area((a1 + a2) / 2.0, (b1 + b2) / 2.0, (c1 + c2) / 2.0)
    return h / 6.0 * (s1 + s2 + 4.0 * s0)


def _frustum_c(r_big, r_small, h):
    """圆台体积：下口半径R、上口半径r、高h。V = πh/3 × (R²+r²+R·r)"""
    return math.pi * h / 3.0 * (r_big * r_big + r_small * r_small
                                + r_big * r_small)


def _frustum_ring(r_out_big, r_out_small, r_in_big, r_in_small, h):
    """空心圆台（管状）体积：外圆台下口半径R、上口半径r，内圆台下口半径R1、上口半径r1，高h。

    V = 外圆台 − 内圆台（内外高度相同）。
    """
    outer = _frustum_c(r_out_big, r_out_small, h)
    inner = _frustum_c(r_in_big, r_in_small, h)
    return outer - inner


FUNCTIONS = {
    'ceil': math.ceil,
    'floor': math.floor,
    'sqr': lambda x: x * x,
    'cube': lambda x: x * x * x,   # V3.7.4：立方（面板上取代幂运算 power）
    'sqrt': math.sqrt,
    'round': round,
    'power': math.pow,
    'radians': math.radians,
    'log': _log,
    'du': math.degrees,
    'abs': abs,
    'max': max,
    'min': min,
    'sin': math.sin,
    'cos': math.cos,
    'tan': math.tan,
    'sum': sum,
    'avg': _avg,
    # V3.7.3 台体公式
    'frustum4': _frustum4,
    'frustum3': _frustum3,
    'frustum_c': _frustum_c,
    'frustum_ring': _frustum_ring,
}

CONSTANTS = {
    'pi': math.pi, 'π': math.pi,
    'e': math.e,
}

# ast 允许的运算
_BINOPS = {
    ast.Add: operator.add, ast.Sub: operator.sub,
    ast.Mult: operator.mul, ast.Div: operator.truediv,
    ast.Pow: operator.pow, ast.Mod: operator.mod,
    ast.FloorDiv: operator.floordiv,
}
_UNARYOPS = {ast.UAdd: operator.pos, ast.USub: operator.neg}


class FormulaError(Exception):
    pass


class Evaluator(ast.NodeVisitor):
    """ast 白名单求值器，杜绝任意代码执行。"""

    def __init__(self, variables, xj_callback, functions):
        self.vars = variables
        self.xj = xj_callback
        self.funcs = functions

    def visit_Expression(self, node):
        return self.visit(node.body)

    def visit_Constant(self, node):
        if isinstance(node.value, (int, float)):
            return node.value
        raise FormulaError(f"不支持的常量类型: {type(node.value).__name__}")

    def visit_Num(self, node):  # py<3.8 兼容
        return node.n

    def visit_Name(self, node):
        name = node.id.lower()
        if name in CONSTANTS:
            return CONSTANTS[name]
        if node.id in self.vars:
            return self.vars[node.id]
        if node.id in self.funcs:
            return self.funcs[node.id]
        raise FormulaError(f"未定义的名称: {node.id}")

    def visit_BinOp(self, node):
        op = _BINOPS.get(type(node.op))
        if op is None:
            raise FormulaError(f"不支持的运算符: {type(node.op).__name__}")
        left, right = self.visit(node.left), self.visit(node.right)
        try:
            return op(left, right)
        except ZeroDivisionError:
            raise FormulaError("除数为 0")

    def visit_UnaryOp(self, node):
        op = _UNARYOPS.get(type(node.op))
        if op is None:
            raise FormulaError(f"不支持的一元运算符: {type(node.op).__name__}")
        return op(self.visit(node.operand))

    def visit_Call(self, node):
        if not isinstance(node.func, ast.Name):
            raise FormulaError("只支持简单函数名调用")
        name = node.func.id.lower()
        args = [self.visit(a) for a in node.args]
        if name == 'xj':
            if self.xj is None:
                raise FormulaError("XJ 汇总需要分部上下文")
            return self.xj(args)
        fn = self.funcs.get(name)
        if fn is None:
            raise FormulaError(f"未知函数: {node.func.id}")
        try:
            return fn(*args)
        except TypeError as e:
            raise FormulaError(f"{node.func.id} 参数错误: {e}")
        except ValueError as e:
            raise FormulaError(f"{node.func.id} 取值错误: {e}")

    def generic_visit(self, node):
        raise FormulaError(f"不支持的语法: {type(node).__name__}")


class FormulaEngine:
    def __init__(self, constants=None):
        self.constants = dict(constants or {})  # {'@层高': 3.6}
        self.functions = dict(FUNCTIONS)
        # V3.9：合法的单元格地址集合（列字母 + 行号）。
        # None = 旧版模式：只把 A<行号> 当作引用（V3.7.9 行为）。
        self.ref_addrs = None
        self.ref_cols = set()

    def set_constants(self, constants):
        self.constants = dict(constants or {})

    def set_refs(self, addrs=None, cols=None):
        """设置本轮可用的单元格地址（Project.recalc 每个分部调一次）。

        addrs 为 None 时退回旧版模式（只认 A<行号>）。
        """
        self.ref_addrs = set(addrs) if addrs is not None else None
        self.ref_cols = set(cols or ())

    def clean(self, expr: str) -> str:
        """返回清理后可直接求值的表达式（用于展示与调试）。"""
        self._missing = []      # 公式里引用、但常量表里没有的名字
        s = to_halfwidth(str(expr))
        # V3.9：允许像 Excel 那样以 = 开头（=4*5、＝4×5 都行）
        if s.startswith('='):
            s = s[1:].lstrip()
        # V3.7.9 / V3.9：单元格引用。
        #   SUM(A1:A10) 先展开成 A1+A2+…+A10（同一列区间）；
        #   合法引用（列字母 + 行号）先换成 \x01序号\x01 占位符，
        #   避免后面的钢筋规格剥离、注释剥离把它破坏；
        #   隐式乘号规则跑之前再还原成地址文本，数值由 context 提供。
        #   ref_addrs 为 None 时是旧版模式：只把 A<行号> 当作引用（V3.7.9 行为）。
        refs = []          # 占位符序号 -> 地址文本

        def _col_no(letters):
            n = 0
            for ch in letters:
                n = n * 26 + (ord(ch) - 64)
            return n

        def _expand_sum(m):
            c1, r1 = m.group(1).upper(), int(m.group(2))
            c2, r2 = m.group(3).upper(), int(m.group(4))
            a, b = (r1, r2) if r1 <= r2 else (r2, r1)
            ci1, ci2 = _col_no(c1), _col_no(c2)
            if ci1 > ci2:
                ci1, ci2 = ci2, ci1
            # V3.9.1：支持跨列矩形区间（点选引用时 Shift+点击会生成 SUM(B2:D5)）
            cells = []
            for ci in range(ci1, ci2 + 1):
                col = ''
                k = ci
                while k:
                    k, rem = divmod(k - 1, 26)
                    col = chr(65 + rem) + col
                for ri in range(a, b + 1):
                    cells.append('%s%d' % (col, ri))
                    if len(cells) >= 2000:
                        break
                if len(cells) >= 2000:
                    break
            return '(' + '+'.join(cells) + ')' if cells else m.group(0)

        s = re.sub(r'(?i)\bSUM\s*\(\s*([A-Za-z]{1,2})(\d{1,7})\s*:\s*'
                   r'([A-Za-z]{1,2})(\d{1,7})\s*\)', _expand_sum, s)

        def _hold_ref(m):
            letter, num = m.group(1).upper(), m.group(2)
            addr = letter + num
            if self.ref_addrs is None:
                if letter != 'A':
                    return m.group(0)      # 旧版模式只认 A<行号>
            elif addr not in self.ref_addrs:
                return m.group(0)          # 不存在的地址原样保留 → 报「未定义的名称」
            refs.append(addr)
            return '\x01%d\x01' % (len(refs) - 1)

        s = CELL_REF_RE.sub(_hold_ref, s)
        # V3.6.1：TssdEng 字体的标准钢筋等级符号（私用区 U+E000 等）统一按 Φ 处理，
        # 这样「<三级钢符号>8@200 4×5」这类规格说明照常被剥离，不参与计算。
        s = re.sub('[\ue000-\ue00e\ue530-\ue53d]', 'Φ', s)
        # 钢筋级别标记 Ⅰ Ⅱ Ⅲ Ⅳ 视为注释（ΦⅢ20 这类写法不影响计算）
        s = s.translate({ord(c): ' ' for c in _GRADE_CHARS})
        # 行业习惯写法 Φ²、Φ³、Φ⁴ 里的上标是“钢筋级别”，不是乘方，先统一成 Φ
        s = re.sub(r'[Φφ]\s*[¹²³⁴]', 'Φ', s)
        # 上标数字 → 幂：²=^2、³=^3、¹=^1、⁰=^0（跟在数字/右括号后表示乘方）
        _sup = '⁰¹²³⁴⁵⁶⁷⁸⁹'
        s = re.sub(r'([0-9)\]])\s*([⁰¹²³⁴⁵⁶⁷⁸⁹]+)',
                   lambda m: m.group(1) + ''.join('^' + str(_sup.index(ch)) for ch in m.group(2)),
                   s)
        s = re.sub(r'([⁰¹²³⁴⁵⁶⁷⁸⁹])', lambda m: '^' + str(_sup.index(m.group(1))), s)
        # 剥离行首规格描述：Φ16@100、C30、HRB400 等（钢筋/材料规格）直到出现明确计算式
        # 规则：若开头是 Φ/φ/钢筋级别符号 或 字母+数字规格，且后面还跟汉字(共/计/等)，整段作注释
        # 钢筋规格：3Φ8@200、Φ20、Φ8@150（可带根数与间距）
        s = re.sub(r'^\s*(?:\d+\s*)?[Φφ]\s*\d+(?:\.\d+)?(?:\s*@\s*\d+(?:\.\d+)?)?\s*[\u4e00-\u9fff]*\s*', '', s)
        # 仅剩间距写法：8@200
        s = re.sub(r'^\s*\d+(?:\.\d+)?\s*@\s*\d+(?:\.\d+)?\s*[\u4e00-\u9fff]*\s*', '', s)
        s = re.sub(r'^[Φφ＠@]?\s*[A-Za-z]?\s*\d+[A-Za-z]?\s*@\s*\d+\s*[\u4e00-\u9fff]*\s*', '', s)
        s = re.sub(r'^[A-Za-z]{1,4}\d+\s+[\u4e00-\u9fff]*\s*', '', s)
        # C30混凝土 5×4 / M10砂浆 3×2：规格字母数字 + 可选汉字 + 空格 + 计算式
        s = re.sub(r'^[A-Za-z]{1,4}\d+\s*[\u4e00-\u9fff]{1,8}\s+(?=[\d(（])', '', s)
        # 度分秒 → 十进制度
        def dms(m):
            d, mi, sec = float(m.group(1)), float(m.group(2)), float(m.group(3))
            return repr(d + mi / 60.0 + sec / 3600.0)
        s = _DMS.sub(dms, s)
        s = strip_comments(s)
        # ^ 在 Python 中是异或，统一改成幂运算 **
        # 先保护连续的 **（用户若写了），再把单个 ^ 换 **
        s = re.sub(r'\*\s*\*', '^', s)
        s = s.replace('^', '**')
        # 常量先占位，避免汉字被剥离
        consts = []
        def hold_const(m):
            key = m.group(0)
            if key in self.constants:
                consts.append(self.constants[key])
            else:
                # 未定义的常量：记下来，报错时给人看得懂的提示
                self._missing.append(key)
                consts.append(key)
            return f"__CONST_{len(consts)-1}__"
        s = re.sub(r'@[\u4e00-\u9fff\w]+', hold_const, s)
        # 去掉汉字注释（函数名、常量占位符为 ASCII，不受影响）
        s = _CJK.sub(' ', s)
        # 保护 XJ() —— 汇总函数的空括号不能删
        s = re.sub(r'(?i)\bXJ\s*\(\s*\)', '__XJ_CALL__', s)
        # 去掉空的括号对
        for _ in range(3):
            new = re.sub(r'\(\s*\)', ' ', s)
            if new == s:
                break
            s = new
        s = s.replace('__XJ_CALL__', 'xj()')
        # 还原常量
        def restore(m):
            v = consts[int(m.group(1))]
            return f'({v})' if isinstance(v, str) else f'({v!r})'
        s = _CONST_TOKEN.sub(restore, s)
        # 开根号 √ → sqrt()：√9、√(3+4)、2√9、√9×4 均可直接算
        s = convert_sqrt(s)
        # 还原单元格引用占位符 → 地址文本（数值由 context 提供）
        # 必须放在隐式乘号规则之前：2A12 这类写法要靠隐式乘号补上 "*"
        if refs:
            s = re.sub(r'\x01(\d+)\x01', lambda m: refs[int(m.group(1))], s)
        # 隐式乘法：)( → )*(  、数字( → 数字*(   、数字sqrt → 数字*sqrt
        # V3.7.3：数字前若是字母（函数名以数字结尾，如 frustum4(...)）不插乘号
        s = re.sub(r'\)\s*\(', ')*(', s)
        s = re.sub(r'(?<![A-Za-z_])(\d|\))\s*\(', r'\1*(', s)
        s = re.sub(r'(\d)(?=[A-Za-z_])', r'\1*', s)
        s = re.sub(r'\)(?=[A-Za-z_])', r')*', s)
        s = re.sub(r'\s+', '', s)
        return s

    def evaluate(self, expr, context=None, xj_callback=None):
        """
        求值。返回 (值, 清理后的表达式, 错误信息)。
        context: {'@常量名': 值, '变量名': 值}
        """
        context = context or {}
        try:
            cleaned = self.clean(expr)
            if not cleaned:
                return None, cleaned, "空表达式"
            if self._missing and not any(k in context for k in self._missing):
                names = '、'.join(dict.fromkeys(self._missing))
                return None, cleaned, f"未定义的常量：{names}（在左下常量框里新增即可）"
            tree = ast.parse(cleaned, mode='eval')
            # 上下文直接交给求值器（大表反复求值时省掉一次整体复制；
            # 求值是同步的，期间调用方不会改动这个 dict）
            variables = context
            ev = Evaluator(variables, xj_callback, self.functions)
            val = ev.visit(tree)
            if isinstance(val, bool):
                val = int(val)
            return float(val), cleaned, None
        except FormulaError as e:
            return None, locals().get('cleaned', ''), str(e)
        except SyntaxError as e:
            return None, locals().get('cleaned', ''), f"公式语法错误: {e.msg}"
        except Exception as e:
            return None, locals().get('cleaned', ''), f"计算错误: {e}"


# ------------------------- 自测 -------------------------
if __name__ == '__main__':
    eng = FormulaEngine(constants={'@层高': 3.6, '@墙厚': 0.24})
    cases = [
        ("4×9.6×2.4×0.4[柱基预埋]", 4 * 9.6 * 2.4 * 0.4),
        ("（3＋5）×2", 16),
        ("2^3+sqr(2)", 12),
        ("ceil(3.2)+floor(3.8)", 7),
        ("round(3.14159,2)", 3.14),
        ("4×9.6×@层高[xxx]", 4 * 9.6 * 3.6),
        ("34°56′30″", 34 + 56/60 + 30/3600),
        ("sqrt(16)+abs(-3)", 7),
        ("power(2,10)", 1024),
        ("max(3,9,5)", 9),
        ("5[纯注释]", 5),
        ("2×（3+4）圆钢框架", 14),
    ]
    ok = 0
    for expr, expect in cases:
        val, cleaned, err = eng.evaluate(expr)
        good = err is None and abs(val - expect) < 1e-9
        ok += good
        print(f"{'OK ' if good else 'FAIL'} | {expr:28s} -> {val} | {cleaned} | {err or ''}")
    print(f"\n{ok}/{len(cases)} 通过")
