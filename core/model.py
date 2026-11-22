# -*- coding: utf-8 -*-
"""数据模型：工程 -> 分部树 -> 数据行。JSON 可序列化。"""

import json, os, re, uuid, datetime
from .formula import FormulaEngine
from .units import default_constants


def new_id():
    return uuid.uuid4().hex[:12]


# 默认层级方案（V2.5 起）：单项工程 > 单位工程 > 分部工程 > 分项工程
# 使用者可在“层级设置”里改名、增删，软件按这里的名字给新建节点起名/标注层级
DEFAULT_LEVELS = ['单项工程', '单位工程', '分部工程', '分项工程']


# 默认列定义：每列 {key, title, width, builtin}
# 列宽：计算公式最宽，编号/单位/不计较窄，备注放右侧面板
DEFAULT_COLUMNS = [
    {'key': 'code', 'title': '编号', 'width': 120, 'builtin': True},
    {'key': 'name', 'title': '项目名称', 'width': 200, 'builtin': True},
    {'key': 'part', 'title': '部位', 'width': 130, 'builtin': True},
    {'key': 'formula', 'title': '计算公式', 'width': 480, 'builtin': True},
    {'key': 'unit', 'title': '单位', 'width': 60, 'builtin': True},
    {'key': 'value', 'title': '工程量', 'width': 100, 'builtin': True},
    {'key': 'remark', 'title': '备注', 'width': 160, 'builtin': True},
    {'key': 'no_count', 'title': '不计', 'width': 50, 'builtin': True},
]


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
        self.fill = ''            # 整行填充色（#RRGGBB，空=无）
        self.cell_fill = {}       # 单元格填充色：{列key: #RRGGBB}
        self.value = None
        self.error = None
        self.cleaned = ""

    def to_dict(self):
        d = dict(id=self.id, code=self.code, name=self.name, part=self.part,
                 formula=self.formula, unit=self.unit, remark=self.remark,
                 no_count=self.no_count, rate=self.rate, value=self.value,
                 error=self.error, cleaned=self.cleaned, fill=self.fill)
        d['custom'] = self.custom
        d['cell_fill'] = self.cell_fill
        return d

    @classmethod
    def from_dict(cls, d):
        r = cls()
        for k, v in d.items():
            if hasattr(r, k):
                setattr(r, k, v)
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
        return dict(id=self.id, name=self.name, expanded=self.expanded,
                    children=[c.to_dict() for c in self.children],
                    rows=[r.to_dict() for r in self.rows])

    @classmethod
    def from_dict(cls, d, parent=None):
        s = cls(d.get('name', '分部'), parent)
        s.id = d.get('id', new_id())
        s.expanded = d.get('expanded', True)
        s.rows = [Row.from_dict(x) for x in d.get('rows', [])]
        s.children = [cls.from_dict(x, s) for x in d.get('children', [])]
        return s

    def iter_self_and_desc(self):
        yield self
        for c in self.children:
            yield from c.iter_self_and_desc()


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
        """重算所有行；XJ 汇总子节点有效工程量。"""
        self.engine.set_constants(self.constants)
        # 后序遍历：先子节点后父节点，保证 XJ 能取到子分部结果
        def walk(sec):
            for c in sec.children:
                walk(c)
            for row in sec.rows:
                if row.no_count:
                    row.value, row.cleaned, row.error = None, row.formula, None
                    continue
                def xj(args):
                    # XJ() 汇总当前分部下所有后代分部的工程量
                    total = 0.0
                    for c in sec.children:
                        for descendant in c.iter_self_and_desc():
                            if descendant.total() is not None:
                                total += descendant.total()
                    return total
                row.value, row.cleaned, row.error = self.engine.evaluate(
                    row.formula, xj_callback=xj)
        for book in self.books:
            walk(book)

    def to_dict(self):
        return dict(id=self.id, name=self.name, created=self.created,
                    modified=self.modified, info=self.info,
                    constants=self.constants, constant_notes=self.constant_notes,
                    levels=self.levels,
                    columns=self.columns, report=self.report,
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
        cols = d.get('columns') or [dict(c) for c in DEFAULT_COLUMNS]
        p.columns = [dict(c) for c in cols]
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
    """本部所有有效行工程量之和。"""
    total = 0.0
    any_valid = False
    for r in self.rows:
        if not r.no_count and r.value is not None:
            total += r.value
            any_valid = True
    return total if any_valid else None

Section.total = _section_total
