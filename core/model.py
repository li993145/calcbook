# -*- coding: utf-8 -*-
"""数据模型：工程 -> 分部树 -> 数据行。JSON 可序列化。"""

import json, os, re, uuid, datetime
from .formula import FormulaEngine, find_cell_refs, to_halfwidth
from .units import default_constants, fmt_num


def new_id():
    return uuid.uuid4().hex[:12]


# 默认层级方案（V2.5 起）：单项工程 > 单位工程 > 分部工程 > 分项工程
# 使用者可在“层级设置”里改名、增删，软件按这里的名字给新建节点起名/标注层级
DEFAULT_LEVELS = ['单项工程', '单位工程', '分部工程', '分项工程']


# 默认列定义：每列 {key, title, width, builtin}
# V3.6.0：按用户要求重排默认列并固定默认宽度。新增列时这些宽度不得被重置。
# V3.7.9：默认列 9 列——序号/清单编号/项目名称/计量单位/工程量/计算公式/备注/图纸/不计
# （图纸编号列从默认列取消；老工程打开时一次性迁移，数据不丢，需要时可自己加回列）
DEFAULT_COLUMNS = [
    {'key': 'seq', 'title': '序号', 'width': 80, 'builtin': True},
    {'key': 'code', 'title': '清单编号', 'width': 135, 'builtin': True},
    {'key': 'name', 'title': '项目名称', 'width': 340, 'builtin': True},
    {'key': 'unit', 'title': '计量单位', 'width': 80, 'builtin': True},
    {'key': 'value', 'title': '工程量', 'width': 100, 'builtin': True},
    {'key': 'formula', 'title': '计算公式', 'width': 600, 'builtin': True},
    {'key': 'remark', 'title': '备注', 'width': 200, 'builtin': True},
    {'key': 'img', 'title': '图纸', 'width': 110, 'builtin': True},
    {'key': 'no_count', 'title': '不计', 'width': 80, 'builtin': True},
]

# V3.9.1：明细子目自动附加列（不放进 DEFAULT_COLUMNS，默认 9 列保持不变）
DETAIL_QTY_COLUMN = {'key': 'qty', 'title': '构件数量', 'width': 110,
                     'builtin': False, 'align': 'center'}

# V3.5 及以前的默认宽度：用于判断老工程里某列宽度是否被用户手动改过
LEGACY_DEFAULT_WIDTHS = {'code': 120, 'name': 200, 'part': 130,
                         'formula': 480, 'unit': 60, 'value': 100,
                         'remark': 160, 'no_count': 50}


# ------------------------------------------------------------------
# 单元格地址（列字母 + 行号）
# ------------------------------------------------------------------
# V3.10：完全 WPS 式——所有列按表格里的物理顺序依次分配 A/B/C…，
# 序号=A、清单编号=B、项目名称=C、计量单位=D、工程量=E、计算公式=F…
# 标识列（序号/清单编号/项目名称/计量单位/图纸/不计）虽然也有字母、
# 也能被公式引用（普通数字格按数字取），但不允许在这些列里写公式，
# 避免「清单编号 A2-001」这类文本被当成公式。
NON_CALC_COLUMNS = {'seq', 'code', 'name', 'unit', 'img', 'no_count'}

# 工程文件里的列字母方案版本号。1=V3.9 旧方案（工程量固定 A），
# 2=V3.10 物理列方案；老工程打开时按旧映射迁移公式里的地址。
COL_LETTER_SCHEME = 2

# 中文（含全角标点）——用于判断一个格子的文本是不是「普通文字」
_CJK_RE = re.compile(r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]')
# V3.10：地址后紧跟「-数字」是型号 / 清单编号（A2-001、C30-1），不是公式。
# 标识列有了字母后，这类文本原来会被 find_cell_refs 命中而误算。
_MODEL_NO_RE = re.compile(r'\b[A-Za-z]{1,2}\d+\s*[-－]\s*\d')


def _col_letters(n):
    """0 起序号 → Excel 式列字母（A, B, … Z, AA, AB …）。"""
    out = []
    for i in range(max(0, int(n))):
        s = ''
        k = i
        while True:
            s = chr(ord('A') + k % 26) + s
            k = k // 26 - 1
            if k < 0:
                break
        out.append(s)
    return out


def column_letters(columns):
    """列 key → 单元格地址字母（V3.10：完全 WPS 式物理列）。

    表格里的每一列都有字母：第 1 列 A、第 2 列 B…第 27 列 AA。
    默认 9 列 → 序号 A、清单编号 B、项目名称 C、计量单位 D、
    工程量 E、计算公式 F、备注 G、图纸 H、不计 I；
    明细表再加「构件数量」插在工程量后 → 构件数量 F、计算公式 G…

    公式引用与标尺完全一致：点击哪一列，标尺上显示什么字母，
    公式里就写什么地址。
    """
    keys = [str(c.get('key', '')) for c in (columns or [])]
    letters = _col_letters(len(keys))
    out = {}
    for idx, k in enumerate(keys):
        if k:
            out[k] = letters[idx] if idx < len(letters) else ''
    return out


def legacy_column_letters(columns):
    """V3.9 旧方案的列字母映射——仅用于老工程公式迁移。

    旧规则：工程量固定 A，其余可计算列（计算公式/备注/自定义列）
    按表格顺序依次 B、C…；标识列无字母。
    """
    keys = [str(c.get('key', '')) for c in (columns or [])]
    calc = [k for k in keys if k and k not in NON_CALC_COLUMNS]
    letters = _col_letters(len(calc))
    out = {}
    for idx, k in enumerate(calc):
        out[k] = letters[idx] if idx < len(letters) else ''
    if 'value' in out and calc and calc[0] != 'value':
        v_idx = calc.index('value')
        out[calc[0]] = letters[v_idx] if v_idx < len(letters) else ''
        out['value'] = 'A'
    return out


def migrate_letter_scheme(proj):
    """老工程（方案 1）公式里的地址字母 → 新物理列字母。

    逐分部按各自的实际列建立「旧字母 → 新字母」映射，只改公式文本
    （计算公式列 / 备注 / 自定义列），普通数据和清单编号一律不动。
    用单元格地址正则整体替换，保证 SUM(A1:A10) 里的两个端点都被迁移。
    """
    from .formula import CELL_REF_RE
    for sec in proj.all_sections():
        cols = proj.effective_columns(sec)
        old, new = legacy_column_letters(cols), column_letters(cols)
        mapping = {}
        for key, letter in old.items():
            if letter and new.get(key):
                mapping[letter] = new[key]
        if not mapping:
            continue

        def repl(m):
            letter = m.group(1).upper()
            return '%s%s' % (mapping.get(letter, letter), m.group(2))

        def migrate_text(s):
            return CELL_REF_RE.sub(repl, str(s or ''))

        for row in sec.rows:
            row.formula = migrate_text(row.formula)
            row.remark = migrate_text(row.remark)
            if isinstance(getattr(row, 'custom', None), dict):
                for k, v in row.custom.items():
                    row.custom[k] = migrate_text(v)


def cell_formula_text(text, addrs=None):
    """判断某个「非公式列」的单元格文本是否按公式计算。

    返回 (是否公式, 公式文本)：
    * 以 '=' 开头 → 明确是公式；
    * 否则必须含「本表真实存在的单元格引用」且不含中文，才当公式
      —— 这样清单编号、备注里的普通文字不会被误算。

    addrs 为本分部合法的单元格地址集合。
    """
    s = str(text or '').strip()
    if not s:
        return False, ''
    if s.startswith('='):
        body = s[1:].strip()
        # 只写了一个 = 是「还没开始输入」，不当公式，免得单元格误标红
        return (True, body) if body else (False, '')
    if _CJK_RE.search(s):
        return False, ''
    if _MODEL_NO_RE.search(s):
        return False, ''
    if not find_cell_refs(s, addrs):
        return False, ''
    return True, s


def plain_cell_number(text):
    """普通格文本若是数字则返回 float，否则 None。

    V3.9.1：像「构件数量」列里直接填 6 这种普通数字格，也必须能被
    =C3*B3 引用。支持全角数字、千分位逗号与前后空格；单位文字（如 6个）
    不当数字，避免把说明文字算进去。
    """
    s = to_halfwidth(str(text or '')).strip().replace(',', '').replace(' ', '')
    if not re.match(r'^[-+]?\d+(?:\.\d+)?$', s):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def cell_source_text(row, key):
    """单元格里的原始文本（公式格是公式本身，普通格是文字）。

    内置列在 Row 属性上，自定义列在 Row.custom 里。
    """
    if hasattr(row, key):
        return str(getattr(row, key) or '')
    return str((getattr(row, 'custom', None) or {}).get(key, '') or '')


def cell_display_text(row, key, precision=3, no_count_text=''):
    """单元格的显示文本（统一口径：公式格显示计算结果）。

    main.py 的表格、报表预览与 Excel 导出都用它，保证三处显示一致。
    """
    if key == 'no_count':
        return no_count_text
    if key == 'img':
        return ''
    if key == 'value':
        return fmt_num(getattr(row, 'value', None), precision)
    if key == 'formula':
        return str(getattr(row, 'formula', '') or '')
    # V3.9：备注列 / 自定义列里写了公式 → 显示算出来的结果
    calc = (getattr(row, 'cell_calc', None) or {}).get(key)
    if calc:
        v = calc.get('value')
        if v is not None:
            return fmt_num(v, precision)
    if hasattr(row, key):
        return str(getattr(row, key) or '')
    return str((getattr(row, 'custom', None) or {}).get(key, '') or '')


def migrate_columns(cols):
    """老工程列定义升级：补齐 V3.6 新增内置列、按新顺序排列，保留用户改动。

    * 老工程已有列的名称、颜色等属性原样保留；
    * 宽度只有在“等于老默认值”时才换成新默认值，用户手动调过的宽度保留；
    * 老工程里的「部位」等不在新默认列里的列放到末尾，数据不会丢失。
    """
    if not cols:
        return [dict(c) for c in DEFAULT_COLUMNS]
    old = {str(c.get('key', '')): c for c in cols}
    if 'seq' in old or 'part' not in old:
        return ensure_img_column([dict(c) for c in cols])
    result = []
    used = set()
    for dcol in DEFAULT_COLUMNS:
        c = dict(dcol)
        oc = old.get(dcol['key'])
        if oc is not None:
            used.add(dcol['key'])
            for k, v in oc.items():
                if k != 'width':
                    c[k] = v
            ow = oc.get('width')
            base = LEGACY_DEFAULT_WIDTHS.get(dcol['key'])
            try:
                if ow is not None and base is not None and int(ow) != int(base):
                    c['width'] = int(ow)    # 用户手动调过的宽度保留
            except (TypeError, ValueError):
                pass
        result.append(c)
    for c in cols:
        if str(c.get('key', '')) not in used:
            result.append(dict(c))
    return ensure_img_column(result)


def ensure_img_column(cols):
    """V3.7.5：老工程没有「图纸」列时，自动插到「计算公式」之后（数据不丢）。"""
    if any(str(c.get('key', '')) == 'img' for c in cols):
        return cols
    result = []
    pos = -1
    for i, c in enumerate(cols):
        result.append(dict(c))
        if str(c.get('key', '')) == 'formula':
            pos = len(result)
    if pos < 0:
        # 没有 formula 列就插在「图纸编号」前，再不行放最前
        pos = next((i for i, c in enumerate(result)
                    if str(c.get('key', '')) == 'drawing_no'), 0)
    result.insert(pos, {'key': 'img', 'title': '图纸', 'width': 110,
                        'builtin': True})
    return result


# V3.7.7 新默认列顺序（apply_column_order_v377 用）
_V377_ORDER = [c['key'] for c in DEFAULT_COLUMNS]
# 内置列改名的映射：只有当老工程里还是旧默认名时才改，自定义名称不动
_V377_RENAME = {'code': {'编号': '清单编号'}, 'unit': {'单位': '计量单位'}}


def apply_column_order_v377(cols):
    """V3.7.7：把工程列按新默认顺序重排 + 内置列改名（清单编号/计量单位），
    并去掉「大样图来源」列。

    * 只动顺序和内置列标题，**每列自己的宽度等设置原样保留**；
    * 大样图来源列从表里取消（单元格数据存在 Row.custom 里，不丢）；
    * 用户自建的自定义列排在内置列后面，相对顺序不变。
    """
    if not cols:
        return [dict(c) for c in DEFAULT_COLUMNS]
    by_key = {}
    rest = []
    for c in cols:
        k = str(c.get('key', ''))
        if k in _V377_ORDER and k not in by_key:
            by_key[k] = dict(c)
        else:
            rest.append(dict(c))         # 自定义列 / 未识别列（含大样图来源外溢）
    result = []
    for k in _V377_ORDER:
        c = by_key.pop(k, None)
        if c is None:
            c = next((dict(d) for d in DEFAULT_COLUMNS if d['key'] == k), None)
        if c is None:
            continue
        ren = _V377_RENAME.get(k)
        if ren and c.get('title') in ren:
            c['title'] = ren[c.get('title')]
        result.append(c)
    # 大样图来源列不再显示；其余自定义列原样接在后面
    result.extend(c for c in rest if str(c.get('key', '')) != 'detail_source')
    return result


# V3.7.9 新默认列顺序（apply_column_order_v379 用）；图纸编号列取消
_V379_ORDER = [c['key'] for c in DEFAULT_COLUMNS]


def apply_column_order_v379(cols):
    """V3.7.9：老工程一次性按 9 列新顺序重排，并去掉「图纸编号」列。

    * 只动顺序，**每列自己的宽度/名称等设置原样保留**；
    * 图纸编号列从表中取消（Row.drawing_no 数据仍存在，需要时可在
      表头右键「添加列」后重新显示——自定义列数据不丢）；
    * 用户自建的自定义列排在内置列后面，相对顺序不变。
    """
    if not cols:
        return [dict(c) for c in DEFAULT_COLUMNS]
    by_key = {}
    rest = []
    for c in cols:
        k = str(c.get('key', ''))
        if k in _V379_ORDER and k not in by_key:
            by_key[k] = dict(c)
        else:
            rest.append(dict(c))          # 自定义列（drawing_no 在 rest 里剔除）
    result = []
    for k in _V379_ORDER:
        c = by_key.pop(k, None)
        if c is None:
            c = next((dict(d) for d in DEFAULT_COLUMNS if d['key'] == k), None)
        if c is not None:
            result.append(c)
    result.extend(c for c in rest if str(c.get('key', '')) != 'drawing_no')
    return result


class Row:
    """一行计算式。"""
    def __init__(self, code="", name="", part="", formula="", unit="",
                 remark="", no_count=False, rate=0.0):
        self.id = new_id()
        self.code = code
        self.name = name
        self.part = part
        self.formula = formula
        self.unit = unit
        self.remark = remark
        self.no_count = no_count
        self.rate = rate
        self.custom = {}          # 自定义列值：{列key: 文本}
        # V3.9：非公式列（备注 / 自定义列）里公式的计算结果
        # {列key: {'value': 数值或 None, 'error': 错误信息或 None}}
        self.cell_calc = {}
        self.fill = ''            # 整行填充色（#RRGGBB，空=无）
        self.cell_fill = {}       # 单元格填充色：{列key: #RRGGBB}
        self.cell_font = {}       # 单元格字体：{列key: {family,size,bold,italic,underline,strikeout,color}}
        self.cell_align = {}      # 单元格对齐：{列key: 'left'/'center'/'right'}（空=默认对齐）
        self.value = None
        self.error = None
        self.cleaned = ""
        self.img = None            # V3.7.5 图纸：{'data':b64,'fmt','thumb':b64,'path','t'} 或 None
        # V3.7.9 合计行：工程量=自动汇总自上一个合计行以来的本分部普通行，
        # 不参与上级 XJ() 汇总与分类汇总（避免重复计数）
        self.total_row = False

    def to_dict(self):
        d = dict(id=self.id, code=self.code, name=self.name, part=self.part,
                 formula=self.formula, unit=self.unit, remark=self.remark,
                 no_count=self.no_count, rate=self.rate, value=self.value,
                 error=self.error, cleaned=self.cleaned, fill=self.fill)
        d['custom'] = self.custom
        d['cell_calc'] = self.cell_calc
        d['cell_fill'] = self.cell_fill
        d['cell_font'] = self.cell_font
        d['cell_align'] = self.cell_align
        if self.img:
            d['img'] = self.img
        if getattr(self, 'total_row', False):
            d['total_row'] = True
        return d

    @classmethod
    def from_dict(cls, d):
        r = cls()
        for k, v in d.items():
            if hasattr(r, k):
                setattr(r, k, v)
        if not isinstance(r.cell_calc, dict):
            r.cell_calc = {}
        if not isinstance(r.custom, dict):
            r.custom = {}
        return r


class Section:
    """分部 / 计算书节点（树形）。"""
    def __init__(self, name="新建分部", parent=None):
        self.id = new_id()
        self.name = name
        self.parent = parent
        self.children = []
        self.rows = []
        self.expanded = True
        # V3.7.6：节点外链（网址或文件路径），明细子目等节点可挂链接
        self.link = ''
        # V3.7.9：分部级列名覆盖 {列key: 标题}——同一个项目里，
        # 明细表可以单独把某列叫别的名字（空 dict = 跟随项目级列名）
        self.col_titles = {}
        # V3.9.1：分部自己的附加列（如新建明细子目时自动加的「构件数量」）。
        # 项目默认列保持 9 列不变，附加列只在本分部显示、参与本分部运算。
        self.extra_columns = []

    def add_child(self, name="新建分部"):
        c = Section(name, self)
        self.children.append(c)
        return c

    def depth(self):
        d, p = 0, self.parent
        while p:
            d += 1
            p = p.parent
        return d

    def to_dict(self):
        d = dict(id=self.id, name=self.name, expanded=self.expanded,
                 children=[c.to_dict() for c in self.children],
                 rows=[r.to_dict() for r in self.rows])
        if getattr(self, 'link', ''):        # 有链接才写盘，老文件零改动
            d['link'] = self.link
        ct = getattr(self, 'col_titles', None)
        if ct:                                # 有覆盖才写盘，老文件零改动
            d['col_titles'] = dict(ct)
        ec = getattr(self, 'extra_columns', None)
        if ec:                                # 有分部自己的列才写盘
            d['extra_columns'] = [dict(c) for c in ec]
        return d

    @classmethod
    def from_dict(cls, d, parent=None):
        s = cls(d.get('name', '分部'), parent)
        s.id = d.get('id', new_id())
        s.expanded = d.get('expanded', True)
        s.link = d.get('link', '')
        ct = d.get('col_titles')
        s.col_titles = dict(ct) if isinstance(ct, dict) else {}
        ec = d.get('extra_columns')
        s.extra_columns = [dict(c) for c in ec] if isinstance(ec, list) else []
        s.rows = [Row.from_dict(x) for x in d.get('rows', [])]
        s.children = [cls.from_dict(x, s) for x in d.get('children', [])]
        return s

    def iter_self_and_desc(self):
        yield self
        for c in self.children:
            yield from c.iter_self_and_desc()


def unit_level_section(sec):
    """sec 所属的单位工程节点（树深度=1）；不在单位工程树下返回 None。

    V4.1 分类汇总「仅当前单位工程的下一级」要用：从选中的节点往上找到
    单位工程，再只取它的直接下级表，更深层的明细表自然被排除。
    """
    s = sec
    while s is not None and s.depth() > 1:
        s = s.parent
    if s is not None and s.parent is not None and s.depth() == 1:
        return s
    return None


class Project:
    def __init__(self, name="未命名工程"):
        self.name = name
        self.id = new_id()
        self.created = datetime.datetime.now().isoformat(timespec='seconds')
        self.modified = self.created
        self.info = {}
        # 常量：V2.4 起不再内置任何常量，全部由使用者自定义（如 @层高=3.6）
        self.constants = default_constants()
        self.constant_notes = {}     # 常量说明：{代码: 说明}，只作备注
        # 层级方案：树的第一层是单项工程，往下依次是单位/分部/分项工程
        self.levels = list(DEFAULT_LEVELS)
        self.columns = [dict(c) for c in DEFAULT_COLUMNS]   # 列定义（可增删）
        self.report = {       # 报表设置
            'title': '', 'header': '', 'footer': '',
            'footer_left': '', 'footer_center': '', 'footer_right': '',
            'margin': 1.9, 'halign': 'left', 'valign': 'center',
            # 纸张（mm）：默认 A4，可在页面设置里按打印机支持的纸张选择
            'paper': {'name': 'A4', 'w': 210.0, 'h': 297.0},
        }
        # 一个工程可包含多个单项工程，每个单项工程是一棵独立的树
        self.books = [Section("单项工程1")]
        self.engine = FormulaEngine(self.constants)
        # 冻结窗格（V3.2）：视图状态，随工程保存
        self.freeze_rows = 0
        self.freeze_cols = 0
        # V3.7.7：新工程直接就是新列序，不需要再迁移
        self.col_order_v377 = True

    @property
    def root(self):
        """兼容别名：指向第一个单项工程。新代码请优先用 self.books。"""
        return self.books[0] if self.books else None

    # ---------------- 层级 ----------------
    def level_name(self, depth):
        """第 depth 层的名称（0=单项工程）；超出方案层数时沿用最后一级。"""
        levels = self.levels or DEFAULT_LEVELS
        depth = max(0, int(depth))
        return levels[min(depth, len(levels) - 1)]

    def level_of(self, sec):
        """某节点在层级方案中的名称（按树深度取）。"""
        if sec is None:
            return self.level_name(0)
        return self.level_name(0 if sec.parent is None else sec.depth())

    def next_level_of(self, sec):
        """sec 的下一级叫什么。"""
        if sec is None:
            return self.level_name(0)
        return self.level_name((0 if sec.parent is None else sec.depth()) + 1)

    # ---------------- 列（项目列 + 分部附加列） ----------------
    def effective_columns(self, sec):
        """某分部实际显示的列：项目列 + 该分部自己的附加列。

        V3.9.1：默认 9 列不含「构件数量」；新建明细子目时只给那个分部
        加一列，其它分部看到的仍是项目默认列。附加列插在「工程量」后面。
        """
        out = list(self.columns or [])
        extra = list(getattr(sec, 'extra_columns', None) or []) if sec else []
        if not extra:
            return out
        keys = set(str(c.get('key', '')) for c in out)
        titles = set(str(c.get('title', '')).strip() for c in out)
        pos = next((i + 1 for i, c in enumerate(out)
                    if c.get('key') == 'value'), len(out))
        for c in extra:
            k, t = str(c.get('key', '')), str(c.get('title', '')).strip()
            if k in keys or t in titles:      # 项目里已有同名列就不再重复加
                continue
            out.insert(pos, c)      # 共享同一个 dict，改列宽/列名能写回分部定义
            pos += 1
            keys.add(k)
            titles.add(t)
        return out

    def all_columns(self):
        """导出 / 报表用：项目列 + 所有分部附加列的并集。"""
        out = list(self.columns or [])
        keys = set(str(c.get('key', '')) for c in out)
        for sec in self.all_sections():
            for c in (getattr(sec, 'extra_columns', None) or []):
                if str(c.get('key', '')) not in keys:
                    out.append(dict(c))
                    keys.add(str(c.get('key', '')))
        return out

    def ensure_detail_qty(self, sec):
        """新建明细子目：给该分部加一列「构件数量」（项目默认列不变）。"""
        if sec is None:
            return False
        existing = list(self.columns or []) + list(getattr(sec, 'extra_columns', None) or [])
        if any(str(c.get('title', '')).strip() == '构件数量'
               or c.get('key') == 'qty' for c in existing):
            return False                       # 项目里已有同名列，直接用它
        sec.extra_columns = list(getattr(sec, 'extra_columns', None) or [])
        sec.extra_columns.append(dict(DETAIL_QTY_COLUMN))
        return True

    @staticmethod
    def _pick_name(label, existing):
        """在 existing 里找一个没用过的“层级名+序号”。"""
        used = set(existing or [])
        n = 1
        while ('%s%d' % (label, n)) in used:
            n += 1
        return '%s%d' % (label, n)

    def new_book_name(self):
        return self._pick_name(self.level_name(0), [b.name for b in self.books])

    def new_child_name(self, parent):
        """parent 下新建节点时的默认名（用下一级的层级名 + 序号）。"""
        label = self.next_level_of(parent)
        existing = [c.name for c in parent.children] if parent is not None else []
        return self._pick_name(label, existing)

    def add_book(self, name=None):
        b = Section(name or self.new_book_name())
        self.books.append(b)
        return b

    def remove_book(self, book):
        if book in self.books:
            self.books.remove(book)

    def is_book(self, sec):
        """判断某节点是否为顶级（单项工程）。"""
        return sec is not None and sec.parent is None and sec in self.books

    def is_first_book(self, sec):
        """第一个单项工程：不可删除（其余可自行增删）。"""
        return bool(self.books) and sec is self.books[0]

    def all_sections(self):
        result = []
        for b in self.books:
            result.extend(b.iter_self_and_desc())
        return result

    def recalc(self):
        """重算所有行；XJ 汇总子节点有效工程量。

        V3.7.9：分部内支持单元格引用——公式里 A<行号> 表示本分部第 <行号> 行
        的工程量（如 A12*0.35、SUM(A1:A10)）。
        V3.9：扩展为二维引用（列字母 + 行号，如 A2*B3）：
            * 列字母按表格列顺序分配（工程量列固定 A，其余可计算列依次 B、C…）；
            * 公式不再只写在「计算公式」列——备注列与自定义列都能写公式
              （以 = 开头，或含本表真实存在的单元格引用），算出的结果可以
              被别的格引用，表格里显示的是算出来的数值；
            * 求值改为迭代（多遍直到收敛），上下行、跨列互相引用都能取到值。
        合计行（Row.total_row）不进公式引擎，直接汇总本分部内
        自上一个合计行以来的普通行。
        """
        self.engine.set_constants(self.constants)

        # 后序遍历：先子节点后父节点，保证 XJ 能取到子分部结果
        def walk(sec):
            for c in sec.children:
                walk(c)
            # V3.9.1：列字母按「本分部实际列」分配——明细子目自己加的
            # 构件数量列只影响本表的 A/B/C 地址，其它分部不受影响。
            letters = column_letters(self.effective_columns(sec))
            # V3.10：每一列都有字母（物理列顺序），全部进入引用上下文
            all_keys = [k for k in letters if letters[k]]
            # 标识列 + 工程量结果列不能写公式；其它列都能写
            other_keys = [k for k in all_keys
                          if k not in NON_CALC_COLUMNS
                          and k not in ('value', 'formula')]
            # 标识列虽不能写公式，但格里的普通数字（如序号 1/2/3）
            # 要能被公式引用，所以照样灌进上下文
            plain_keys = [k for k in all_keys if k in NON_CALC_COLUMNS]
            f_letter = letters.get('formula', '')
            v_letter = letters.get('value', '')
            rows = sec.rows
            n = len(rows)
            if not n:
                return
            # 本分部合法的单元格地址（列字母 × 行号）
            addrs = set()
            for k in all_keys:
                L = letters[k]
                for i in range(1, n + 1):
                    addrs.add('%s%d' % (L, i))
            self.engine.set_refs(addrs, set(letters[k] for k in all_keys))
            # 地址 → 数值（初值 0；求值时按行推进）
            values = dict((a, 0.0) for a in addrs)

            def xj(args):
                # XJ() 汇总当前分部下所有后代分部的工程量
                total = 0.0
                for c in sec.children:
                    for descendant in c.iter_self_and_desc():
                        if descendant.total() is not None:
                            total += descendant.total()
                return total

            # 迭代求解：每遍按行顺序重算，值不再变化即收敛。
            # 上限必须随行数增长——反向引用链（本行=下一行+1）每遍只往前
            # 传播一行，n 行就要 n 遍；写死 40 会让 45 行以上的链条算不完。
            # 循环引用不会收敛，跑满上限即停（不崩、不无限循环）。
            max_pass = min(max(4, n + 3), 500)
            for _pass in range(max_pass):
                before = dict(values)
                for idx, row in enumerate(rows):
                    rno = idx + 1
                    addr_v = ('%s%d' % (v_letter, rno)) if v_letter else None
                    addr_f = ('%s%d' % (f_letter, rno)) if f_letter else None
                    if getattr(row, 'total_row', False):
                        # 合计行：汇总自上一个合计行以来的普通行（跳过不计行）
                        start = 0
                        for j in range(idx - 1, -1, -1):
                            if getattr(rows[j], 'total_row', False):
                                start = j + 1
                                break
                        tot = sum(rows[j].value for j in range(start, idx)
                                  if not rows[j].no_count
                                  and rows[j].value is not None
                                  and not getattr(rows[j], 'total_row', False))
                        row.value, row.cleaned, row.error = tot, '合计', None
                        if addr_v:
                            values[addr_v] = tot
                        if addr_f:
                            values[addr_f] = tot
                        continue
                    if row.no_count:
                        row.value, row.cleaned, row.error = None, row.formula, None
                        if addr_v:
                            values[addr_v] = 0.0
                        if addr_f:
                            values[addr_f] = 0.0
                        for k in other_keys:
                            row.cell_calc.pop(k, None)
                        continue
                    # 本行先归零：自己引用自己按 0 处理
                    for k in all_keys:
                        values['%s%d' % (letters[k], rno)] = 0.0
                    # 标识列里的普通数字也进引用上下文；序号列没有存文本，
                    # 按它显示的行号取值（标尺 A 列 → A1=1、A2=2…）
                    for k in plain_keys:
                        num = float(rno) if k == 'seq' else \
                            plain_cell_number(cell_source_text(row, k))
                        values['%s%d' % (letters[k], rno)] = \
                            num if num is not None else 0.0
                    # 1) 计算公式列（结果写进行工程量）
                    row.value, row.cleaned, row.error = self.engine.evaluate(
                        row.formula, context=values, xj_callback=xj)
                    v = row.value if row.value is not None else 0.0
                    if addr_v:
                        values[addr_v] = v
                    if addr_f:
                        values[addr_f] = v
                    # 2) 备注列 / 自定义列里自己写的公式
                    for k in other_keys:
                        text = cell_source_text(row, k)
                        is_formula, ftext = cell_formula_text(text, addrs)
                        if not is_formula:
                            row.cell_calc.pop(k, None)
                            # 普通数字格（如构件数量填 6）同样进入引用上下文，
                            # 否则 =C3*B3 会把 6 当 0 算。
                            num = plain_cell_number(text)
                            values['%s%d' % (letters[k], rno)] = \
                                num if num is not None else 0.0
                            continue
                        val, _cleaned, err = self.engine.evaluate(
                            ftext, context=values, xj_callback=xj)
                        values['%s%d' % (letters[k], rno)] = \
                            val if val is not None else 0.0
                        row.cell_calc[k] = {'value': val, 'error': err}
                if values == before:
                    break
        for book in self.books:
            walk(book)

    def to_dict(self):
        return dict(id=self.id, name=self.name, created=self.created,
                    modified=self.modified, info=self.info,
                    constants=self.constants, constant_notes=self.constant_notes,
                    levels=self.levels,
                    columns=self.columns, report=self.report,
                    freeze_rows=getattr(self, 'freeze_rows', 0),
                    freeze_cols=getattr(self, 'freeze_cols', 0),
                    col_order_v377=getattr(self, 'col_order_v377', False),
                    col_order_v379=getattr(self, 'col_order_v379', False),
                    col_letter_scheme=COL_LETTER_SCHEME,
                    books=[b.to_dict() for b in self.books],
                    root=self.books[0].to_dict() if self.books else None, version=4)

    @classmethod
    def from_dict(cls, d):
        p = cls(d.get('name', '工程'))
        for k in ('id', 'created', 'modified', 'info', 'constants',
                  'constant_notes', 'report'):
            if k in d:
                setattr(p, k, d[k])
        # 常量全部由用户自定义：旧工程里已有的常量原样保留，不再补内置钢筋单重
        if not isinstance(p.constants, dict):
            p.constants = {}
        if not isinstance(getattr(p, 'constant_notes', None), dict):
            p.constant_notes = {}
        # 层级方案（V2.5）：没有或为空则用默认四级
        lv = d.get('levels')
        if isinstance(lv, list):
            lv = [str(x).strip() for x in lv if str(x).strip()]
        p.levels = lv or list(DEFAULT_LEVELS)
        # 列定义：兼容旧工程（无 columns 或空列表则用默认）
        p.columns = migrate_columns(d.get('columns') or [dict(c) for c in DEFAULT_COLUMNS])
        # V3.7.7：老工程一次性按新默认顺序重排 + 内置列改名（清单编号/计量单位），
        # 之后再打开不再动（尊重用户自己拖动过的列序）；每列宽度原样保留
        p.col_order_v377 = bool(d.get('col_order_v377', False))
        if not p.col_order_v377:
            p.columns = apply_column_order_v377(p.columns)
            p.col_order_v377 = True
        # V3.7.9：一次性按 9 列新顺序重排并去掉「图纸编号」列，之后不再动
        p.col_order_v379 = bool(d.get('col_order_v379', False))
        if not p.col_order_v379:
            p.columns = apply_column_order_v379(p.columns)
            p.col_order_v379 = True
        # 冻结窗格（V3.2）：旧工程无此字段默认不冻结
        try:
            p.freeze_rows = max(0, int(d.get('freeze_rows') or 0))
        except (TypeError, ValueError):
            p.freeze_rows = 0
        try:
            p.freeze_cols = max(0, int(d.get('freeze_cols') or 0))
        except (TypeError, ValueError):
            p.freeze_cols = 0
        # 单项工程列表：新格式用 books；兼容旧格式只有 root
        books_data = d.get('books')
        if books_data:
            p.books = [Section.from_dict(x, None) for x in books_data]
        else:
            p.books = [Section.from_dict(d.get('root', {'name': '单项工程1'}), None)]
        # 兼容 V2.4 及更早：顶级节点叫“计算书N”的改名为“单项工程N”
        for b in p.books:
            m = re.match(r'^计算书\s*(\d*)$', (b.name or '').strip())
            if m:
                b.name = p.level_name(0) + (m.group(1) or '')
        # V3.10：老工程的公式地址按旧映射（工程量=A）迁移到物理列字母
        try:
            scheme = int(d.get('col_letter_scheme', 1) or 1)
        except (TypeError, ValueError):
            scheme = 1
        if scheme < COL_LETTER_SCHEME:
            migrate_letter_scheme(p)
        p.recalc()
        return p

    def save(self, path):
        d = self.to_dict()
        d['modified'] = datetime.datetime.now().isoformat(timespec='seconds')
        self.modified = d['modified']
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
        return path

    @classmethod
    def load(cls, path):
        with open(path, encoding='utf-8') as f:
            return cls.from_dict(json.load(f))


# Section.total 在下面定义（避免类体内引用未定义）
def _section_total(self):
    """本部所有有效行工程量之和（合计行是汇总展示，不重复计入）。"""
    total = 0.0
    any_valid = False
    for r in self.rows:
        if getattr(r, 'total_row', False):
            continue
        if not r.no_count and r.value is not None:
            total += r.value
            any_valid = True
    return total if any_valid else None

Section.total = _section_total
