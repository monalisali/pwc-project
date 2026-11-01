# -*- coding: utf-8 -*-
"""
比对 结果文件(成品油发票判定) 与 税局Energy增值税发票 的汽油发票数量
- 税局文件剔除变动方式为 销售红冲 / 增值税发票开具 的行(销项, 非进项)
- 开票日期对齐: 结果文件发票限定在税局文件汽油行的日期范围内(结果范围 >= 税局范围)
- 按发票号码 join, 逐张比对数量
输出: energy/output/汽油数量比对_结果vs税局.xlsx
"""
import os
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # .../energy
RESULT = os.path.join(BASE, 'output', '成品油发票判定_Energy进项20260501-20261026.xlsx')
TAX = os.path.join(BASE, '税局Energy增值税发票7.1-10.23.xlsx')
DST = os.path.join(BASE, 'output', '汽油数量比对_结果vs税局.xlsx')

EXCLUDE = ('销售红冲', '增值税发票开具')   # 销项变动方式, 剔除
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


# ---------- 结果文件: 大类=汽油 ----------
wb_r = openpyxl.load_workbook(RESULT, data_only=True)
ws_r = wb_r['发票汇总判定']
rows_r = list(ws_r.iter_rows(values_only=True))
h_r = {c: i for i, c in enumerate(rows_r[0])}
res = {}
for r in rows_r[1:]:
    if r[h_r['是否成品油发票']] == '是' and r[h_r['油品分类（大类）']] == '汽油':
        d = to_date(r[h_r['开票日期']])
        res[str(r[h_r['发票号码']])] = {
            'date': d, 'sub': str(r[h_r['油品细分（细类）']]),
            'qty': f(r[h_r['数量']]) or 0.0, 'unit': str(r[h_r['单位']] or '')}

# ---------- 税局文件: 汽油, 剔除销项 ----------
wb_t = openpyxl.load_workbook(TAX, data_only=True)
ws_t = wb_t['增值税发票']
rows_t = list(ws_t.iter_rows(values_only=True))
h_t = {c: i for i, c in enumerate(rows_t[0])}
tax = {}
for r in rows_t[1:]:
    if str(r[h_t['商品和服务名称']]).strip() != '汽油':
        continue
    if str(r[h_t['变动方式']]).strip() in EXCLUDE:
        continue
    no = str(r[h_t['数电发票号码']] or r[h_t['发票号码']] or '').strip()
    if not no:
        continue
    tax.setdefault(no, {'date': to_date(r[h_t['开具日期']]),
                        'qty': (f(r[h_t['数量（吨）']]) or 0.0),
                        'way': str(r[h_t['变动方式']])})

# ---------- 日期对齐: 取税局汽油行的日期范围, 限定结果文件 ----------
t_dates = [v['date'] for v in tax.values()]
d_min, d_max = min(t_dates), max(t_dates)
res_in_range = {k: v for k, v in res.items() if d_min <= v['date'] <= d_max}
res_out_range = {k: v for k, v in res.items() if not (d_min <= v['date'] <= d_max)}

# ---------- 逐发票比对 ----------
detail = []
for no in sorted(set(res_in_range) | set(tax), key=lambda x: res_in_range.get(x, tax.get(x))['date']):
    a, b = res_in_range.get(no), tax.get(no)
    if a and b:
        diff = round(a['qty'] - b['qty'], 4)
        status = '一致' if abs(diff) < TOL / 2 else '数量差异'
        detail.append([no, a['date'].isoformat(), a['sub'], a['qty'], b['qty'], diff,
                       status, b['way'], ''])
    elif a:
        detail.append([no, a['date'].isoformat(), a['sub'], a['qty'], None, None,
                       '仅结果文件', '', '税局文件中无此发票(或为销项已剔除)'])
    else:
        detail.append([no, b['date'].isoformat(), '', None, b['qty'], None,
                       '仅税局文件', b['way'], '结果文件中无此汽油发票'])

# ---------- 汇总 ----------
def total(rows, col):
    return round(sum(r[col] or 0 for r in rows), 4)

n_ok = sum(1 for d in detail if d[6] == '一致')
n_diff = sum(1 for d in detail if d[6] == '数量差异')
n_res_only = sum(1 for d in detail if d[6] == '仅结果文件')
n_tax_only = sum(1 for d in detail if d[6] == '仅税局文件')
both = [d for d in detail if d[6] in ('一致', '数量差异')]
summary = [
    ['结果文件范围', '税局文件汽油日期范围(比对窗口)', f'{d_min.isoformat()} ~ {d_max.isoformat()}'],
    ['比对窗口内结果文件汽油发票数', len(res_in_range), '张'],
    ['比对窗口外结果文件汽油发票数(未参与比对)', len(res_out_range), '张'],
    ['税局汽油发票数(剔除销项后)', len(tax), '张'],
    ['双方都有', len(both), '张'],
    ['  其中数量一致', n_ok, '张'],
    ['  其中数量差异', n_diff, '张'],
    ['仅结果文件有', n_res_only, '张'],
    ['仅税局文件有', n_tax_only, '张'],
    ['结果文件汽油数量合计(窗口内)', total(both, 3) + total([d for d in detail if d[6] == '仅结果文件'], 3), '吨'],
    ['税局汽油数量合计', total(both, 4) + total([d for d in detail if d[6] == '仅税局文件'], 4), '吨'],
    ['数量差异合计(结果-税局)', round(total(both, 5) or 0, 4), '吨'],
]
# 按差异绝对值排序, 差异在前
detail.sort(key=lambda d: (d[6] not in ('数量差异', '仅结果文件', '仅税局文件'),
                           -(abs(d[5]) if d[5] is not None else 0), d[1]))

# ---------- 写 Excel ----------
out = openpyxl.Workbook()
hdr_font = Font(bold=True, color='FFFFFF')
hdr_fill = PatternFill('solid', fgColor='4472C4')
diff_fill = PatternFill('solid', fgColor='FCE4EC')   # 差异行淡红
only_fill = PatternFill('solid', fgColor='FFF3E0')   # 单边行淡橙

ws1 = out.active
ws1.title = '汇总'
for row in summary:
    ws1.append(row)
for c in ws1['A']:
    c.font = Font(bold=True)
ws1.column_dimensions['A'].width = 40
ws1.column_dimensions['B'].width = 32
ws1.column_dimensions['C'].width = 8

ws2 = out.create_sheet('逐票比对')
ws2.append(['发票号码', '开票日期', '油品细分（细类）', '结果文件数量', '税局数量',
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
for i, w in enumerate([24, 12, 26, 14, 14, 16, 12, 14, 34], 1):
    ws2.column_dimensions[get_column_letter(i)].width = w
ws2.freeze_panes = 'A2'
ws2.auto_filter.ref = ws2.dimensions

out.save(DST)

# ---------- 统计输出 ----------
import io
buf = io.StringIO()
p = lambda *a: print(*a, file=buf)
p('输出:', DST)
for row in summary:
    p('  ', row[0], ':', row[1], row[2] if len(row) > 2 else '')
p()
p('差异明细(数量差异/仅单边):')
for d in detail:
    if d[6] != '一致':
        p('  ', d)
open(os.path.join(BASE, 'output', '_cmp_stats.txt'), 'w', encoding='utf-8').write(buf.getvalue())
print('saved')
