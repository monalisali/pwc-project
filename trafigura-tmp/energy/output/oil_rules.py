# -*- coding: utf-8 -*-
"""
成品油判定共用规则与工具: 被 oil_classify.py(进项) / oil_classify_sales.py(销项) 引用
规则: 用【规格型号+商品名称】与 Mapping 表条目做"包含"比较
"""
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# 细类 -> (大类, 关键字条件列表)  条件全部满足(包含)才算命中
# 顺序即判定优先级: 京VIB 在 VIB 之前(京VIB包含VIB), -35/-10 在 0 之前(-10/-35包含0), 京VI 在 VI 之前
RULES = [
    # ---- 汽油 ----
    ('92号车用汽油（京VIB）', '汽油', ['92', '京VIB', '车用汽油']),
    ('95号车用汽油（京VIB）', '汽油', ['95', '京VIB', '车用汽油']),
    ('98号车用汽油（京VIB）', '汽油', ['98', '京VIB', '车用汽油']),
    ('92号车用汽油（VIB）',   '汽油', ['92', 'VIB', '车用汽油']),
    ('95号车用汽油（VIB）',   '汽油', ['95', 'VIB', '车用汽油']),
    ('98号车用汽油（VIB）',   '汽油', ['98', 'VIB', '车用汽油']),
    # 无标号车用汽油(如"车用汽油（VIB)", 无92/95/98前缀) -> 细类 NA; 有标号的会先命中上方细类规则
    ('NA',                    '汽油', ['VIB', '车用汽油']),
    # ---- 车用乙醇汽油调合组分油 ----
    ('95号车用乙醇汽油调合组分', '车用乙醇汽油调合组分油', ['95', '车用乙醇汽油']),
    ('92号车用乙醇汽油调合组分', '车用乙醇汽油调合组分油', ['92', '车用乙醇汽油']),
    # ---- 柴油 ----
    ('0号车用柴油（京VI）',  '柴油', ['0', '京VI', '车用柴油']),
    ('-35号车用柴油（VI）', '柴油', ['-35', 'VI', '车用柴油']),
    ('-10号车用柴油（VI）', '柴油', ['-10', 'VI', '车用柴油']),
    ('0号车用柴油（VI）',   '柴油', ['0', 'VI', '车用柴油']),
    # ---- 船用燃料油(柴油大类, 见油品分类对应关系表) ----
    ('DMX船用燃料油', '柴油', ['DMX', '船用燃料油']),
    ('DMA船用燃料油', '柴油', ['DMA', '船用燃料油']),
    ('船用燃料油',     '柴油', ['船用燃料油']),
    # ---- 石脑油 ----
    ('石脑油', '石脑油', ['石脑油']),
]

# 罗马数字等常见变体归一化(仅用于第二遍兜底匹配, 并在结果中标注)
NORMALIZE = {'Ⅵ': 'VI', 'Ⅴ': 'V', '＃': '#', '　': ' '}


def norm(text):
    for k, v in NORMALIZE.items():
        text = text.replace(k, v)
    return text


def match(text, rules=RULES):
    """返回 (细类, 大类, 命中关键字) 或 None"""
    for sub, cat, kws in rules:
        if all(k in text for k in kws):
            return sub, cat, '+'.join(kws)
    return None


def s(v):
    """None -> '', 其余转字符串"""
    return '' if v is None else str(v)


# ---------- Excel 输出样式 ----------
HDR_FONT = Font(bold=True, color='FFFFFF')
HDR_FILL = PatternFill('solid', fgColor='4472C4')
OIL_FILL = PatternFill('solid', fgColor='FFF2CC')   # 成品油行淡黄


def write_sheet(ws, header, rows, widths, flag_col):
    ws.append(header)
    for c in ws[1]:
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(horizontal='center')
    for row in rows:
        ws.append(row)
        if row[flag_col] == '是':
            for c in ws[ws.max_row]:
                c.fill = OIL_FILL
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions


def classify_detail(no, date, inv_type, seller, spec, name, unit, qty, text):
    """对一行明细做判定, 返回该行的输出列表(与明细判定 sheet 列序一致)"""
    m = match(text)
    note = ''
    if not m:
        m2 = match(norm(text))
        if m2:
            m = m2
            note = '归一化匹配(罗马数字Ⅵ→VI等)'
    if m:
        sub, cat, kws = m
        flag, big, small = '是', cat, sub
    else:
        flag, big, small = '否', '', ''
        kws = ''
    row = [
        no, date, inv_type, seller,
        name, spec, unit, qty, text, flag, big, small,
        (kws + ('；' + note if note else '')) if (kws or note) else '',
        note,
    ]
    return row, (big, small, note) if flag == '是' else None
