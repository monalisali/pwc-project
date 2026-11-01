# -*- coding: utf-8 -*-
"""
销项比对: 结果文件(成品油发票判定_Energy销项) vs 税局Energy增值税发票, 大类=汽油, 数量差异
1. 日期对齐: 取双方汽油日期范围交集为比对窗口(结果文件范围 >= 税局)
2. 税局剔除变动方式 已勾选红冲/增值税发票勾选(进项变动); 保留 增值税发票开具/销售红冲(销项)
3. 销售红冲数量记负数
输出: energy/output/汽油数量比对_销项vs税局.xlsx
"""
import os
import io
import sys
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from oil_rules import s

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # .../energy
RESULT = os.path.join(BASE, 'output', '成品油发票判定_Energy销项.xlsx')
TAX = os.path.join(BASE, '税局Energy增值税发票7.1-10.23.xlsx')
DST = os.path.join(BASE, 'output', '汽油数量比对_销项vs税局.xlsx')

KEEP_WAY = ('增值税发票开具', '销售红冲')   # 销项变动方式, 保留; 其余(勾选类, 属进项)剔除
NEG_WAY = ('销售红冲',)                     # 数量记负数
TAX_FROM = date(2026, 7, 16)                # 税局取数起始日(含当天)
TAX_WINDOW = (TAX_FROM, date(2026, 10, 23)) # 税局名义取数窗口(源文件名 7.1-10.23, 起点改 07-16;
                                            # 即使07-16当天无数据, 窗口也记为07-16起, 避免误读)
TOL = 1e-4


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


# ---------- 结果文件(销项): 大类=汽油 ----------
wb_r = openpyxl.load_workbook(RESULT, data_only=True)
ws_r = wb_r['发票汇总判定']
rows_r = list(ws_r.iter_rows(values_only=True))
h_r = {c: i for i, c in enumerate(rows_r[0])}
res, res_sub, res_pre = {}, {}, {}
for r in rows_r[1:]:
    if r[h_r['是否成品油发票']] == '是' and r[h_r['油品分类（大类）']] == '汽油':
        no = str(r[h_r['发票号码']])
        res[no] = to_date(r[h_r['开票日期']])
        res[no + '_qty'] = f(r[h_r['数量']]) or 0.0
        res_sub[no] = str(r[h_r['油品细分（细类）']])
        res_pre[no] = '已预制' in str(r[h_r['备注']] or '')

# ---------- 税局文件: 汽油, 仅保留销项变动方式, 销售红冲记负 ----------
wb_t = openpyxl.load_workbook(TAX, data_only=True)
ws_t = wb_t['增值税发票']
rows_t = list(ws_t.iter_rows(values_only=True))
h_t = {c: i for i, c in enumerate(rows_t[0])}
tax, tax_ways = {}, {}
n_excluded = 0
for r in rows_t[1:]:
    if s(r[h_t['商品和服务名称']]).strip() != '汽油':
        continue
    if to_date(r[h_t['开具日期']]) < TAX_FROM:   # 税局取数从 2026-07-16 起(含)
        continue
    way = s(r[h_t['变动方式']]).strip()
    if way not in KEEP_WAY:
        n_excluded += 1
        continue
    no = s(r[h_t['数电发票号码']] or r[h_t['发票号码']]).strip()
    if not no:
        continue
    qty = f(r[h_t['数量（吨）']]) or 0.0
    if way in NEG_WAY:
        qty = -qty
    rec = tax.setdefault(no, {'date': to_date(r[h_t['开具日期']]), 'qty': 0.0, 'ways': []})
    rec['qty'] += qty
    rec['ways'].append(way)

# ---------- 日期窗口 ----------
# 起点: 名义取数窗口起点的最大者(税局自07-16起, 即使07-16当天无数据窗口也记为07-16)
# 终点: 双方实际数据最大日期的最小者
r_dates = [v for k, v in res.items() if not k.endswith('_qty')]
t_dates = [v['date'] for v in tax.values()]
w_min = max(TAX_WINDOW[0], min(r_dates))
w_max = min(max(r_dates), max(t_dates))
in_win = lambda d: w_min <= d <= w_max

res_w = {k: res[k + '_qty'] for k in res if not k.endswith('_qty') and in_win(res[k])}
tax_w = {k: v['qty'] for k, v in tax.items() if in_win(v['date'])}

# ---------- 逐票比对 ----------
detail = []
for no in sorted(set(res_w) | set(tax_w),
                 key=lambda x: (res.get(x) or tax.get(x, {}).get('date')).isoformat()):
    a, b = res_w.get(no), tax_w.get(no)
    if a is not None and b is not None:
        diff = round(a - b, 4)
        status = '一致' if abs(diff) < TOL else '数量差异'
        note = ''
    elif a is not None:
        status = '仅结果文件'
        note = '已预制发票(未开票, 税局无号可对)' if res_pre.get(no) else '税局文件中无此发票(或变动方式已剔除)'
    else:
        status = '仅税局文件'
        note = '结果文件中无此汽油发票'
    date_s = (res.get(no) or tax.get(no, {}).get('date')).isoformat()
    detail.append([no, date_s, res_sub.get(no, ''), a, b,
                   round(a - b, 4) if (a is not None and b is not None) else None,
                   status, '&'.join(tax_ways.get(no, tax.get(no, {}).get('ways', []))) or '', note])

detail.sort(key=lambda d: (d[6] == '一致', d[1], d[0]))

# ---------- 汇总 ----------
def T(d):
    return round(sum(d.values()), 4)

n_ok = sum(1 for d in detail if d[6] == '一致')
n_diff = sum(1 for d in detail if d[6] == '数量差异')
n_res = sum(1 for d in detail if d[6] == '仅结果文件')
n_tax = sum(1 for d in detail if d[6] == '仅税局文件')
n_res_pre = sum(1 for d in detail if d[6] == '仅结果文件' and res_pre.get(d[0]))
summary = [
    ['结果文件(销项)汽油实际数据范围', f'{min(r_dates)} ~ {max(r_dates)}'],
    ['税局取数窗口', f'{TAX_WINDOW[0]} ~ {TAX_WINDOW[1]}'],
    ['税局汽油实际数据范围(仅销项变动后)', f'{min(t_dates)} ~ {max(t_dates)}'],
    ['比对窗口', f'{w_min} ~ {w_max}'],
    ['窗口内发票数: 结果文件', len(res_w)],
    ['窗口内发票数: 税局', len(tax_w)],
    ['双方都有', n_ok + n_diff],
    ['  其中数量一致', n_ok],
    ['  其中数量差异', n_diff],
    ['仅结果文件有', n_res, f'（其中已预制 {n_res_pre} 张）'],
    ['仅税局文件有', n_tax],
    ['税局被剔除的汽油行数(勾选类, 属进项)', n_excluded],
    ['汽油数量合计-结果文件', T(res_w), '吨'],
    ['汽油数量合计-税局', T(tax_w), '吨'],
    ['数量差异合计(结果-税局)', round(T(res_w) - T(tax_w), 4), '吨'],
]

# ---------- 写 Excel ----------
out = openpyxl.Workbook()
hdr_font = Font(bold=True, color='FFFFFF')
hdr_fill = PatternFill('solid', fgColor='4472C4')
diff_fill = PatternFill('solid', fgColor='FCE4EC')   # 差异淡红
only_fill = PatternFill('solid', fgColor='FFF3E0')   # 单边淡橙

ws1 = out.active
ws1.title = '汇总'
for row in summary:
    ws1.append(list(row))
for c in ws1['A']:
    c.font = Font(bold=True)
ws1.column_dimensions['A'].width = 40
ws1.column_dimensions['B'].width = 32
ws1.column_dimensions['C'].width = 26

ws2 = out.create_sheet('逐票比对')
ws2.append(['发票号码', '开票日期', '细类（结果文件）', '结果文件数量', '税局数量',
            '差异(结果-税局)', '比对结果', '税局变动方式', '说明'])
for c in ws2[1]:
    c.font, c.fill = hdr_font, hdr_fill
    c.alignment = Alignment(horizontal='center')
for row in detail:
    ws2.append(row)
    fill = diff_fill if row[6] == '数量差异' else (only_fill if row[6].startswith('仅') else None)
    if fill:
        for c in ws2[ws2.max_row]:
            c.fill = fill
for i, w in enumerate([24, 12, 26, 14, 14, 15, 12, 16, 36], 1):
    ws2.column_dimensions[get_column_letter(i)].width = w
ws2.freeze_panes = 'A2'
ws2.auto_filter.ref = ws2.dimensions

out.save(DST)

# ---------- 统计输出 ----------
buf = io.StringIO()
p = lambda *a: print(*a, file=buf)
p('输出:', DST)
for row in summary:
    p('  ', row[0], ':', row[1] if len(row) < 3 else f'{row[1]} {row[2]}')
p()
p('非一致明细(差异/仅单边):')
for d in detail:
    if d[6] != '一致':
        p('  ', d)
open(os.path.join(BASE, 'output', '_cmp_sales_stats.txt'), 'w', encoding='utf-8').write(buf.getvalue())
print('saved')
