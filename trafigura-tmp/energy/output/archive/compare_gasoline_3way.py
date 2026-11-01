# -*- coding: utf-8 -*-
"""
三方比对汽油发票数量: 结果文件(成品油发票判定_进项) vs 税局成品油发票 vs TMP购进数据
1. 日期对齐: 取三方汽油日期范围的交集为比对窗口(结果文件范围 >= 税局)
2. 税局剔除变动方式 销售红冲/增值税发票开具(销项); 已勾选红冲数量记负数
3. TMP: 油品类型=1(汽油), 剔除油品细类=10/11, 按发票号码聚合油品数量
输出: energy/output/汽油数量比对_三方.xlsx
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
RESULT = os.path.join(BASE, 'output', '成品油发票判定_Energy进项20260501-20261026.xlsx')
TAX = os.path.join(BASE, '税局Energy增值税发票7.1-10.23.xlsx')
TMP = os.path.join(BASE, 'output', '购进数据选择.xlsx')
DST = os.path.join(BASE, 'output', '汽油数量比对_进项_三方.xlsx')

EXCLUDE_WAY = ('销售红冲', '增值税发票开具')   # 销项变动方式, 剔除
NEG_WAY = ('已勾选红冲',)                      # 数量记负数
TMP_EXCLUDE_SUB = (10, 11)                     # TMP 剔除的油品细类
TAX_FROM = date(2026, 7, 16)                   # 税局取数起始日(含当天)
# 名义取数窗口(即使窗口起点无数据也按名义窗口记录, 避免被误读为实际首张日期)
RES_WINDOW = (date(2026, 5, 1), date(2026, 10, 26))   # 进项结果文件名义窗口(见源文件名 20260501-20261026)
TAX_WINDOW = (TAX_FROM, date(2026, 10, 23))           # 税局文件名义窗口(见源文件名 7.1-10.23, 起点改 07-16)
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


# ---------- 结果文件: 大类=汽油 ----------
wb_r = openpyxl.load_workbook(RESULT, data_only=True)
ws_r = wb_r['发票汇总判定']
rows_r = list(ws_r.iter_rows(values_only=True))
h_r = {c: i for i, c in enumerate(rows_r[0])}
res, res_sub = {}, {}
for r in rows_r[1:]:
    if r[h_r['是否成品油发票']] == '是' and r[h_r['油品分类（大类）']] == '汽油':
        no = str(r[h_r['发票号码']])
        res[no] = to_date(r[h_r['开票日期']])
        res_sub[no] = str(r[h_r['油品细分（细类）']])
        res[no + '_qty'] = f(r[h_r['数量']]) or 0.0

# ---------- 税局文件: 汽油, 剔除销项, 已勾选红冲记负 ----------
wb_t = openpyxl.load_workbook(TAX, data_only=True)
ws_t = wb_t['增值税发票']
rows_t = list(ws_t.iter_rows(values_only=True))
h_t = {c: i for i, c in enumerate(rows_t[0])}
tax = {}
for r in rows_t[1:]:
    if s(r[h_t['商品和服务名称']]).strip() != '汽油':
        continue
    if to_date(r[h_t['开具日期']]) < TAX_FROM:   # 税局取数从 2026-07-16 起(含)
        continue
    way = s(r[h_t['变动方式']]).strip()
    if way in EXCLUDE_WAY:
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

# ---------- TMP购进数据: 油品类型=1, 剔除细类10/11 ----------
wb_p = openpyxl.load_workbook(TMP, data_only=True)
ws_p = wb_p['1']
rows_p = list(ws_p.iter_rows(values_only=True))
h_p = {c: i for i, c in enumerate(rows_p[0])}
tmp, tmp_sub = {}, {}
tmp_dropped = 0
for r in rows_p[1:]:
    if r[h_p['油品类型']] != 1:
        continue
    if r[h_p['油品细类']] in TMP_EXCLUDE_SUB:
        tmp_dropped += 1
        continue
    no = s(r[h_p['发票号码']]).strip()
    if not no:
        continue
    tmp[no] = to_date(r[h_p['开票日期']])
    tmp[no + '_qty'] = (tmp.get(no + '_qty') or 0.0) + (f(r[h_p['油品数量（吨）']]) or 0.0)
    tmp_sub[no] = r[h_p['油品细类']]

# ---------- 日期窗口 ----------
# 起点: 各名义取数窗口起点的最大者(税局自07-16起, 即使07-16当天无数据窗口也记为07-16)
# 终点: 各方实际数据最大日期的最小者
r_dates = [v for k, v in res.items() if not k.endswith('_qty')]
t_dates = [v['date'] for v in tax.values()]
p_dates = [v for k, v in tmp.items() if not k.endswith('_qty')]
w_min = max(RES_WINDOW[0], TAX_WINDOW[0], min(p_dates))
w_max = min(max(r_dates), max(t_dates), max(p_dates))
in_win = lambda d: w_min <= d <= w_max

res_w = {k: res[k + '_qty'] for k in res if not k.endswith('_qty') and in_win(res[k])}
tax_w = {k: v['qty'] for k, v in tax.items() if in_win(v['date'])}
tmp_w = {k: tmp[k + '_qty'] for k in tmp if not k.endswith('_qty') and in_win(tmp[k])}

# ---------- 逐票三方比对 ----------
def qty_of(d, no):
    return d.get(no)


detail = []
nos = sorted(set(res_w) | set(tax_w) | set(tmp_w),
             key=lambda x: (res.get(x) or tax.get(x, {}).get('date') or tmp.get(x)).isoformat())
for no in nos:
    a, b, c = res_w.get(no), tax_w.get(no), tmp_w.get(no)
    present = [v for v in (a, b, c) if v is not None]
    diffs = []
    if len(present) == 3:
        if abs(a - b) > TOL:
            diffs.append('结果≠税局')
        if abs(a - c) > TOL:
            diffs.append('结果≠TMP')
        if abs(b - c) > TOL:
            diffs.append('税局≠TMP')
        status = '一致' if not diffs else '差异:' + '&'.join(diffs)
    else:
        missing = [n for v, n in ((a, '结果'), (b, '税局'), (c, 'TMP')) if v is None]
        status = '缺失:' + '&'.join(missing)
    date_s = (res.get(no) or tax.get(no, {}).get('date') or tmp.get(no)).isoformat()
    note = ''
    if a is None:
        note = '结果文件无此汽油发票'
    if b is None:
        note = (note + '；' if note else '') + '税局无此发票(或销项已剔除)'
    if c is None:
        note = (note + '；' if note else '') + 'TMP购进无此发票'
    detail.append([
        no, date_s, res_sub.get(no, ''),
        a, b, c,
        round(a - b, 4) if (a is not None and b is not None) else None,
        round(a - c, 4) if (a is not None and c is not None) else None,
        status,
        '&'.join(tax.get(no, {}).get('ways', [])) or 'TMP细类:' + str(tmp_sub.get(no, '')) if no in tmp_sub or no in tax else '',
        note,
    ])

# 排序: 差异/缺失在前, 再按日期
detail.sort(key=lambda d: (d[8] == '一致', d[1], d[0]))

# ---------- 汇总 ----------
def T(d):
    return round(sum(v for v in d.values()), 4)

n3 = sum(1 for d in detail if d[8] == '一致')
n_diff = sum(1 for d in detail if d[8].startswith('差异'))
n_miss = sum(1 for d in detail if d[8].startswith('缺失'))
summary = [
    ['结果文件(进项)取数窗口', f'{RES_WINDOW[0]} ~ {RES_WINDOW[1]}'],
    ['结果文件汽油实际数据范围', f'{min(r_dates)} ~ {max(r_dates)}'],
    ['税局取数窗口', f'{TAX_WINDOW[0]} ~ {TAX_WINDOW[1]}'],
    ['税局汽油实际数据范围(剔除销项后)', f'{min(t_dates)} ~ {max(t_dates)}'],
    ['TMP购进汽油实际数据范围(剔除细类10/11前)', f'{min(p_dates)} ~ {max(p_dates)}'],
    ['比对窗口', f'{w_min} ~ {w_max}'],
    ['窗口内发票数: 结果文件', len(res_w)],
    ['窗口内发票数: 税局', len(tax_w)],
    ['窗口内发票数: TMP购进', len(tmp_w)],
    ['三方一致', n3],
    ['有差异', n_diff],
    ['存在缺失(仅一方或两方有)', n_miss],
    ['TMP因细类10/11剔除的汽油行数', tmp_dropped],
    ['汽油数量合计-结果文件', T(res_w), '吨'],
    ['汽油数量合计-税局', T(tax_w), '吨'],
    ['汽油数量合计-TMP购进', T(tmp_w), '吨'],
    ['差异合计(结果-税局)', round(T(res_w) - T(tax_w), 4), '吨'],
    ['差异合计(结果-TMP)', round(T(res_w) - T(tmp_w), 4), '吨'],
    ['差异合计(税局-TMP)', round(T(tax_w) - T(tmp_w), 4), '吨'],
]

# ---------- 写 Excel ----------
out = openpyxl.Workbook()
hdr_font = Font(bold=True, color='FFFFFF')
hdr_fill = PatternFill('solid', fgColor='4472C4')
diff_fill = PatternFill('solid', fgColor='FCE4EC')   # 差异淡红
miss_fill = PatternFill('solid', fgColor='FFF3E0')   # 缺失淡橙

ws1 = out.active
ws1.title = '汇总'
for row in summary:
    ws1.append(list(row))
for c in ws1['A']:
    c.font = Font(bold=True)
ws1.column_dimensions['A'].width = 40
ws1.column_dimensions['B'].width = 34
ws1.column_dimensions['C'].width = 8

ws2 = out.create_sheet('逐票比对')
ws2.append(['发票号码', '开票日期', '细类（结果文件）', '结果文件数量', '税局数量',
            'TMP购进数量', '差异(结果-税局)', '差异(结果-TMP)', '比对结果',
            '税局变动方式/TMP细类', '说明'])
for c in ws2[1]:
    c.font, c.fill = hdr_font, hdr_fill
    c.alignment = Alignment(horizontal='center')
for row in detail:
    ws2.append(row)
    fill = miss_fill if row[8].startswith('缺失') else (diff_fill if row[8].startswith('差异') else None)
    if fill:
        for c in ws2[ws2.max_row]:
            c.fill = fill
for i, w in enumerate([24, 12, 26, 14, 14, 14, 15, 15, 16, 18, 36], 1):
    ws2.column_dimensions[get_column_letter(i)].width = w
ws2.freeze_panes = 'A2'
ws2.auto_filter.ref = ws2.dimensions

out.save(DST)

# ---------- 统计输出 ----------
buf = io.StringIO()
p = lambda *a: print(*a, file=buf)
p('输出:', DST)
for row in summary:
    p('  ', row[0], ':', row[1])
p()
p('非一致明细(差异/缺失):')
for d in detail:
    if d[8] != '一致':
        p('  ', d)
open(os.path.join(BASE, 'output', '_cmp3_stats.txt'), 'w', encoding='utf-8').write(buf.getvalue())
print('saved')
