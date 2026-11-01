# -*- coding: utf-8 -*-
"""
油品大类台账 (流程步骤3), 按大类参数化: 初始 + 进项比对 - 销项比对(税局口径)
- 初始数据: Energy初始数据.xlsx 中 f'{category}合计' 行, 可配置多个初始列(每个生成一个场景)
- 进项: {category}数量比对_进项_三方.xlsx 逐票, 数量取 结果文件数量 -> 税局数量 -> TMP购进数量
- 销项: {category}数量比对_销项vs税局.xlsx 逐票, 数量取 税局数量(无税局数量的已预制票剔除并记录)
- 按开票日期滚动计算累计结余, 标记首次转为负数的发票
- 输出: energy/output/{category}数量结果_初始加进项减销项.xlsx
可被流程调用: from ledger_category import run; run(category, ...)
"""
import io
import os
import sys
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # .../energy
OUT = os.path.join(BASE, 'output')

TOL = 1e-6


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return date.fromisoformat(str(v)[:10])


def f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def coalesce(*vals):
    for v in vals:
        if v is not None:
            return v
    return None


def _window_of(cmp_path):
    """从比对文件汇总sheet读 比对窗口 行, 形如 '2026-07-16 ~ 2026-10-13'"""
    wb = openpyxl.load_workbook(cmp_path, data_only=True)
    for r in wb['汇总'].iter_rows(values_only=True):
        if r and r[0] == '比对窗口':
            return str(r[1]).replace(' ~ ', '~')
    return ''


def run(category, init_path, cmp_in_path, cmp_out_path, dst, start_date, init_cols):
    """init_cols: [(列字母, 说明), ...], 每个"初始列"生成一个场景; 返回统计 dict"""
    # ---------- 初始数据: {category}合计行的各初始列 ----------
    wb0 = openpyxl.load_workbook(init_path, data_only=True)
    rows0 = list(wb0.active.iter_rows(values_only=True))
    init_vals = {}
    for r in rows0[1:]:
        if str(r[0] or '').strip() == f'{category}合计':
            for col, _desc in init_cols:
                init_vals[col] = f(r[column_index_from_string(col) - 1])
            break
    missing = [c for c, _ in init_cols if init_vals.get(c) is None]
    assert not missing, f'未找到{category}合计行的初始列: {missing}'

    # ---------- 进项事件: 结果文件数量 -> 税局数量 -> TMP购进数量 ----------
    events = []   # (date, direction, no, sub, qty, remark)
    wb_i = openpyxl.load_workbook(cmp_in_path, data_only=True)
    rows_i = list(wb_i['逐票比对'].iter_rows(values_only=True))
    h = {c: i for i, c in enumerate(rows_i[0])}
    for r in rows_i[1:]:
        qty = coalesce(f(r[h['结果文件数量']]), f(r[h['税局数量']]),
                       f(r[h['TMP购进数量']]) if 'TMP购进数量' in h else None)
        if qty is None:
            continue
        events.append((to_date(r[h['开票日期']]), '进项', str(r[h['发票号码']]),
                       str(r[h['细类（结果文件）']] or ''), qty,
                       str(r[h['比对结果']] or '') + ('；' + str(r[h['说明']]) if r[h['说明']] else '')))

    # ---------- 销项事件: 税局数量(无税局数量的已预制票剔除) ----------
    excluded_sales = []
    wb_s = openpyxl.load_workbook(cmp_out_path, data_only=True)
    rows_s = list(wb_s['逐票比对'].iter_rows(values_only=True))
    h = {c: i for i, c in enumerate(rows_s[0])}
    for r in rows_s[1:]:
        qty = f(r[h['税局数量']])
        if qty is None:
            excluded_sales.append((str(r[h['发票号码']]), str(r[h['开票日期']]),
                                   f(r[h['结果文件数量']]), str(r[h['说明']] or '')))
            continue
        events.append((to_date(r[h['开票日期']]), '销项', str(r[h['发票号码']]),
                       str(r[h['细类（结果文件）']] or ''), qty,
                       str(r[h['比对结果']] or '') + ('；' + str(r[h['说明']]) if r[h['说明']] else '')))

    events.sort(key=lambda e: (e[0], e[1], e[2]))
    total_in = round(sum(e[4] for e in events if e[1] == '进项'), 4)
    total_out = round(sum(e[4] for e in events if e[1] == '销项'), 4)

    def build_ledger(init_qty):
        """返回 (ledger行, 首次转负tuple或None): 行=(日期,方向,发票,细类,变动,结余,备注)"""
        ledger = [(start_date, '初始', f'(初始数据 {category}合计)', '', round(init_qty, 6),
                   round(init_qty, 6), '')]
        balance = init_qty
        first_neg = None
        for d, direction, no, sub, qty, remark in events:
            delta = qty if direction == '进项' else -qty   # 进项为加(红冲为负即减), 销项为减(红冲为负即加回)
            balance += delta
            ledger.append((d, direction, no, sub, round(delta, 6), round(balance, 6), remark))
            if first_neg is None and balance < -TOL:
                first_neg = (d, direction, no, sub, round(delta, 4), round(balance, 6))
        return ledger, first_neg

    win_in = _window_of(cmp_in_path)
    win_out = _window_of(cmp_out_path)
    cmp_in_name = os.path.basename(cmp_in_path)[:-5]   # 去掉 .xlsx
    cmp_out_name = os.path.basename(cmp_out_path)[:-5]

    scens = []
    for i, (col, desc) in enumerate(init_cols, 1):
        qty = init_vals[col]
        scens.append((f'场景{i}: 初始取{col}列({desc})', qty, *build_ledger(qty)))

    # ---------- 写 Excel ----------
    out = openpyxl.Workbook()
    hdr_font = Font(bold=True, color='FFFFFF')
    hdr_fill = PatternFill('solid', fgColor='4472C4')
    neg_fill = PatternFill('solid', fgColor='FCE4EC')
    init_fill = PatternFill('solid', fgColor='E2EFDA')

    ws1 = out.active
    ws1.title = '汇总'
    for name, qty, ledger, first_neg in scens:
        final = round(ledger[-1][5], 4)
        ws1.append([name])
        ws1[ws1.max_row][0].font = Font(bold=True, size=12)
        for row in [
            [f'初始数量({start_date})', round(qty, 4), '吨'],
            [f'进项合计({cmp_in_name} 逐票, 窗口{win_in})', total_in, '吨'],
            [f'销项合计(税局口径, {cmp_out_name} 逐票, 窗口{win_out})', total_out, '吨'],
            ['最终数量 = 初始 + 进项 - 销项', final, '吨'],
            ['是否出现负数', '是' if first_neg else '否'],
        ]:
            ws1.append(list(row))
        if first_neg:
            for row in [
                ['首次转负日期', str(first_neg[0])],
                ['首次转负发票', first_neg[2]],
                ['该票方向/细类', f'{first_neg[1]} / {first_neg[3]}'],
                ['该票数量变动', first_neg[4]],
                ['该票结余(转负时)', first_neg[5]],
            ]:
                ws1.append(list(row))
        ws1.append([])
    if excluded_sales:
        ws1.append(['销项中无税局数量、已剔除的发票(均为已预制):'])
        for no, d, q, note in excluded_sales:
            ws1.append([f'  {no} {d} 结果文件数量={q} {note}'])
    for c in ws1['A']:
        if c.value:
            c.font = Font(bold=True)
    ws1.column_dimensions['A'].width = 66
    ws1.column_dimensions['B'].width = 22
    ws1.column_dimensions['C'].width = 8

    for (name, qty, ledger, first_neg), (col, _desc) in zip(scens, init_cols):
        ws = out.create_sheet(f'流水_初始{col}列')
        ws.append(['日期', '方向', '发票号码', '细类', '数量变动(+进/-销)', '累计结余', '备注(比对状态)'])
        for c in ws[1]:
            c.font, c.fill = hdr_font, hdr_fill
            c.alignment = Alignment(horizontal='center')
        for d, direction, no, sub, change, bal, remark in ledger:
            ws.append([str(d), direction, no, sub, change, bal, remark])
            row = ws[ws.max_row]
            if direction == '初始':
                for c in row:
                    c.fill = init_fill
            elif bal < -TOL:
                for c in row:
                    c.fill = neg_fill
        for i, w in enumerate([12, 8, 24, 24, 17, 14, 44], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions

    out.save(dst)

    # ---------- 统计 ----------
    stats = {'输出': dst, '进项合计': total_in, '销项合计(税局)': total_out, '销项剔除': excluded_sales,
             '场景': []}
    for name, qty, ledger, first_neg in scens:
        neg = [l for l in ledger if l[5] < -TOL]
        stats['场景'].append({'名称': name, '初始': round(qty, 4), '最终': round(ledger[-1][5], 4),
                              '首次转负': first_neg, '负数流水行数': len(neg),
                              '负数区间末行': neg[-1] if neg else None})
    return stats


def format_stats(st):
    buf = io.StringIO()
    p = lambda *a: print(*a, file=buf)
    p('输出:', st['输出'])
    for s_ in st['场景']:
        p()
        p('==', s_['名称'])
        p('   初始:', s_['初始'], '| 进项合计:', st['进项合计'], '| 销项合计(税局):', st['销项合计(税局)'],
          '| 最终:', s_['最终'])
        p('   首次转负:', s_['首次转负'])
        p('   负数流水行数:', s_['负数流水行数'], '| 负数区间末行:', s_['负数区间末行'])
    p()
    p('销项剔除(无税局数量):', st['销项剔除'])
    return buf.getvalue()
