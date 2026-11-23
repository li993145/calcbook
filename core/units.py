# -*- coding: utf-8 -*-
"""计量单位与工程量小数位规则。

规则（V2.2 起）：
  * m、m2、kg 一类（长度/面积/体积/重量公斤） → 保留 2 位小数
  * t（吨）                                     → 保留 3 位小数
  * 个、块、套、处、台、组 等计件类             → 取整（0 位小数）
未列出的单位仍按程序的默认精度（默认 3 位）显示。
"""

# 单位列下拉框：常用工程计量单位（可在单元格里直接输入自定义单位）
UNIT_GROUPS = [
    ('最常用', ['m', 'm2', 'm3', 'kg', 't', '个', '块', '套', '处', '台', '组']),
    ('计件/构件', ['根', '樘', '座', '件', '项', '只', '系统', '台班', '组日', '千块', '百块']),
    ('长度/面积扩展', ['mm', 'cm', 'km', '延长米', '10m', '100m', 'm·d', 'm2·d', 'm3·d']),
]

UNIT_LIST = [u for _g, items in UNIT_GROUPS for u in items]

# 单位 → 小数位数（键已归一：小写、² → 2、³ → 3、去掉空格）
UNIT_DECIMALS = {
    # 长度：2 位
    'm': 2, '米': 2, 'mm': 2, 'cm': 2, 'km': 2, '10m': 2, '100m': 2,
    '延长米': 2, '延米': 2,
    # 面积 / 体积：2 位
    'm2': 2, '平方米': 2, 'm3': 2, '立方米': 2,
    'm·d': 2, 'm2·d': 2, 'm3·d': 2,
    # 重量：kg 2 位，t 3 位
    'kg': 2, '千克': 2, '公斤': 2,
    't': 3, '吨': 3, 't/m3': 3,
    # 计件：取整
    '个': 0, '块': 0, '套': 0, '处': 0, '台': 0, '组': 0,
    '根': 0, '樘': 0, '座': 0, '件': 0, '项': 0, '只': 0,
    '系统': 0, '台班': 0, '组日': 0, '千块': 0, '百块': 0,
}

# 取整的单位（用于界面提示）
INTEGER_UNITS = ['个', '块', '套', '处', '台', '组']


def norm_unit(unit):
    """单位归一化：小写、全角/上标换算、去空格。"""
    u = str(unit or '').strip().lower().replace(' ', '')
    return (u.replace('㎡', 'm2').replace('m²', 'm2').replace('m³', 'm3')
             .replace('²', '2').replace('³', '3'))


def precision_for_unit(unit, default=3):
    """按单位返回工程量应保留的小数位数；未知单位用 default。"""
    u = norm_unit(unit)
    if not u:
        return default
    if u in UNIT_DECIMALS:
        return UNIT_DECIMALS[u]
    # 复合单位（如 m2/人、m3·台班）按前缀判断
    for key in ('m2', 'm3', 'm', 'kg', 't'):
        if u.startswith(key):
            return UNIT_DECIMALS[key]
    return default


def section_precision(section, default=3):
    """一个分部小计的小数位：取该分部各行单位所需位数的最大值。"""
    ps = []
    for r in getattr(section, 'rows', []) or []:
        if getattr(r, 'no_count', False) or getattr(r, 'value', None) is None:
            continue
        ps.append(precision_for_unit(getattr(r, 'unit', ''), default))
    return max(ps) if ps else default


# 常用规格钢筋单位重量（kg/m）——V2.4 起**不再自动写进工程**，
# 仅作为“常量模板”供用户一键插入（常量全部由使用者自己定义）
REBAR_CONSTANTS = [
    ('@Φ6', 0.222), ('@Φ8', 0.395), ('@Φ10', 0.617), ('@Φ12', 0.888),
    ('@Φ14', 1.21), ('@Φ16', 1.58), ('@Φ18', 2.00), ('@Φ20', 2.47),
    ('@Φ22', 2.98), ('@Φ25', 3.85), ('@Φ28', 4.83), ('@Φ32', 6.31),
    ('@Φ36', 7.99), ('@Φ40', 9.87),
]


def default_constants():
    """新建工程时的常量：V2.4 起为空，全部交给用户自定义。

    （钢筋单重等只是“模板”，见 REBAR_CONSTANTS，可在常量框里一键插入）
    """
    return {}


def fmt_const(v):
    """常量值显示：整数不带小数点，小数最多 6 位并去尾零。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    if abs(f - round(f)) < 1e-9:
        return str(int(round(f)))
    return ('%.6f' % f).rstrip('0').rstrip('.')


def fmt_num(v, prec=3):
    """按位数格式化工程量：四舍五入（half-up）、带千分位；None 显示空。"""
    if v is None:
        return ''
    try:
        from decimal import Decimal, ROUND_HALF_UP
        q = Decimal(1).scaleb(-int(prec))
        d = Decimal(str(float(v))).quantize(q, rounding=ROUND_HALF_UP)
        if d == 0:
            d = abs(d)          # 避免 -0.00
    except Exception:
        return str(v)
    return format(d, ',.%df' % int(prec))
