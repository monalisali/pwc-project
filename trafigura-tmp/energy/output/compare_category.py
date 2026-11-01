# -*- coding: utf-8 -*-
"""
油品大类比对 (流程步骤2), 按大类参数化, 支持 汽油/柴油/... 任意大类
- run_input:  进项三方比对  结果文件(进项判定) vs 税局 vs TMP购进 -> {大类}数量比对_进项_三方.xlsx
- run_output: 销项比对      结果文件(销项判定) vs 税局           -> {大类}数量比对_销项vs税局.xlsx
税局侧通用规则:
  进项比对: 剔除 销售红冲/增值税发票开具(销项); 已勾选红冲数量记负; 自 TAX_FROM 起取数
  销项比对: 剔除 已勾选红冲/增值税发票勾选(进项); 销售红冲数量记负; 自 TAX_FROM 起取数
"""
import io
import os
import sys
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from oil_rules import s

TOL = 1e-4

HDR_FONT = Font(bold=True, color='FFFFFF')
HDR_FILL = PatternFill('solid', fgColor='4472C4')
DIFF_FILL = PatternFill('solid', fgColor='FCE4EC')   # 差异淡红
MISS_FILL = PatternFill('solid', fgColor='FFF3E0')   # 缺失淡橙


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


# ============ 进项三方比对 ============
def run_input(category, result_path, tax_path, tmp_path, dst,
              tax_from, res_window, tax_window, tmp_type, tmp_exclude_sub):
    """进项三方比对, 输出 {category}数量比对_进项_三方.xlsx, 返回(汇总行, 逐票行)"""
    excl_way = ('销售红冲', '增值税发票开具')
    neg_way = ('已勾选红冲',)

    # 结果文件(进项判定): 大类=category
    wb_r = openpyxl.load_workbook(result_path, data_only=True)
    rows_r = list(wb_r['发票汇总判定'].iter_rows(values_only=True))
    h_r = {c: i for i, c in enumerate(rows_r[0])}
    res, res_sub = {}, {}
    for r in rows_r[1:]:
        if r[h_r['是否成品油发票']] == '是' and r[h_r['油品分类（大类）']] == category:
            no = str(r[h_r['发票号码']])
            res[no] = to_date(r[h_r['开票日期']])
            res_sub[no] = str(r[h_r['油品细分（细类）']])
            res[no + '_qty'] = f(r[h_r['数量']]) or 0.0

    # 税局: 剔除销项变动, 红冲记负, 自 tax_from 起取数
    wb_t = openpyxl.load_workbook(tax_path, data_only=True)
    rows_t = list(wb_t['增值税发票'].iter_rows(values_only=True))
    h_t = {c: i for i, c in enumerate(rows_t[0])}
    tax = {}
    for r in rows_t[1:]:
        if s(r[h_t['商品和服务名称']]).strip() != category:
            continue
        if to_date(r[h_t['开具日期']]) < tax_from:
            continue
        way = s(r[h_t['变动方式']]).strip()
        if way in excl_way:
            continue
        no = s(r[h_t['数电发票号码']] or r[h_t['发票号码']]).strip()
        if not no:
            continue
        qty = f(r[h_t['数量（吨）']]) or 0.0
        if way in neg_way:
            qty = -qty
        rec = tax.setdefault(no, {'date': to_date(r[h_t['开具日期']]), 'qty': 0.0, 'ways': []})
        rec['qty'] += qty
        rec['ways'].append(way)

    # TMP购进: 油品类型=指定编码, 剔除指定细类, 按发票号聚合
    wb_p = openpyxl.load_workbook(tmp_path, data_only=True)
    rows_p = list(wb_p.worksheets[0].iter_rows(values_only=True))
    h_p = {c: i for i, c in enumerate(rows_p[0])}
    tmp, tmp_sub = {}, {}
    tmp_dropped = 0
    for r in rows_p[1:]:
        if r[h_p['油品类型']] != tmp_type:
            continue
        if r[h_p['油品细类']] in tmp_exclude_sub:
            tmp_dropped += 1
            continue
        no = s(r[h_p['发票号码']]).strip()
        if not no:
            continue
        tmp[no] = to_date(r[h_p['开票日期']])
        tmp[no + '_qty'] = (tmp.get(no + '_qty') or 0.0) + (f(r[h_p['油品数量（吨）']]) or 0.0)
        tmp_sub[no] = r[h_p['油品细类']]

    # 日期窗口: 起点=名义取数窗口起点的最大者; 终点=各方实际数据最大日期的最小者
    r_dates = [v for k, v in res.items() if not k.endswith('_qty')]
    t_dates = [v['date'] for v in tax.values()]
    p_dates = [v for k, v in tmp.items() if not k.endswith('_qty')]
    w_min = max(res_window[0], tax_window[0], min(p_dates))
    w_max = min(max(r_dates), max(t_dates), max(p_dates))
    in_win = lambda d: w_min <= d <= w_max

    res_w = {k: res[k + '_qty'] for k in res if not k.endswith('_qty') and in_win(res[k])}
    tax_w = {k: v['qty'] for k, v in tax.items() if in_win(v['date'])}
    tmp_w = {k: tmp[k + '_qty'] for k in tmp if not k.endswith('_qty') and in_win(tmp[k])}

    # 逐票三方比对
    detail = []
    for no in sorted(set(res_w) | set(tax_w) | set(tmp_w),
                     key=lambda x: (res.get(x) or tax.get(x, {}).get('date') or tmp.get(x)).isoformat()):
        a, b, c = res_w.get(no), tax_w.get(no), tmp_w.get(no)
        present = [v for v in (a, b, c) if v is not None]
        if len(present) == 3:
            diffs = []
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
            note = f'结果文件无此{category}发票'
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
    detail.sort(key=lambda d: (d[8] == '一致', d[1], d[0]))

    def T(d):
        return round(sum(v for v in d.values()), 4)

    summary = [
        [f'结果文件(进项)取数窗口', f'{res_window[0]} ~ {res_window[1]}'],
        [f'结果文件{category}实际数据范围', f'{min(r_dates)} ~ {max(r_dates)}'],
        ['税局取数窗口', f'{tax_window[0]} ~ {tax_window[1]}'],
        [f'税局{category}实际数据范围(剔除销项后)', f'{min(t_dates)} ~ {max(t_dates)}'],
        [f'TMP购进{category}实际数据范围(剔除细类{"/".join(str(x) for x in tmp_exclude_sub)}前)',
         f'{min(p_dates)} ~ {max(p_dates)}'],
        ['比对窗口', f'{w_min} ~ {w_max}'],
        ['窗口内发票数: 结果文件', len(res_w)],
        ['窗口内发票数: 税局', len(tax_w)],
        ['窗口内发票数: TMP购进', len(tmp_w)],
        ['三方一致', sum(1 for d in detail if d[8] == '一致')],
        ['有差异', sum(1 for d in detail if d[8].startswith('差异'))],
        ['存在缺失(仅一方或两方有)', sum(1 for d in detail if d[8].startswith('缺失'))],
        [f'TMP因细类{"/".join(str(x) for x in tmp_exclude_sub)}剔除的{category}行数', tmp_dropped],
        [f'{category}数量合计-结果文件', T(res_w), '吨'],
        [f'{category}数量合计-税局', T(tax_w), '吨'],
        [f'{category}数量合计-TMP购进', T(tmp_w), '吨'],
        ['差异合计(结果-税局)', round(T(res_w) - T(tax_w), 4), '吨'],
        ['差异合计(结果-TMP)', round(T(res_w) - T(tmp_w), 4), '吨'],
        ['差异合计(税局-TMP)', round(T(tax_w) - T(tmp_w), 4), '吨'],
    ]

    _write_compare_workbook(dst, summary, detail,
                            ['发票号码', '开票日期', f'细类（结果文件）', '结果文件数量', '税局数量',
                             'TMP购进数量', '差异(结果-税局)', '差异(结果-TMP)', '比对结果',
                             '税局变动方式/TMP细类', '说明'],
                            [24, 12, 26, 14, 14, 14, 15, 15, 16, 18, 36])
    return summary, detail


# ============ 销项 vs 税局 ============
def run_output(category, result_path, tax_path, dst, tax_from, tax_window):
    """销项比对, 输出 {category}数量比对_销项vs税局.xlsx, 返回(汇总行, 逐票行)"""
    keep_way = ('增值税发票开具', '销售红冲')
    neg_way = ('销售红冲',)

    # 结果文件(销项判定): 大类=category
    wb_r = openpyxl.load_workbook(result_path, data_only=True)
    rows_r = list(wb_r['发票汇总判定'].iter_rows(values_only=True))
    h_r = {c: i for i, c in enumerate(rows_r[0])}
    res, res_sub, res_pre = {}, {}, {}
    for r in rows_r[1:]:
        if r[h_r['是否成品油发票']] == '是' and r[h_r['油品分类（大类）']] == category:
            no = str(r[h_r['发票号码']])
            res[no] = to_date(r[h_r['开票日期']])
            res[no + '_qty'] = f(r[h_r['数量']]) or 0.0
            res_sub[no] = str(r[h_r['油品细分（细类）']])
            res_pre[no] = '已预制' in str(r[h_r['备注']] or '')

    # 税局: 仅保留销项变动, 销售红冲记负, 自 tax_from 起取数
    wb_t = openpyxl.load_workbook(tax_path, data_only=True)
    rows_t = list(wb_t['增值税发票'].iter_rows(values_only=True))
    h_t = {c: i for i, c in enumerate(rows_t[0])}
    tax = {}
    n_excluded = 0
    for r in rows_t[1:]:
        if s(r[h_t['商品和服务名称']]).strip() != category:
            continue
        if to_date(r[h_t['开具日期']]) < tax_from:
            continue
        way = s(r[h_t['变动方式']]).strip()
        if way not in keep_way:
            n_excluded += 1
            continue
        no = s(r[h_t['数电发票号码']] or r[h_t['发票号码']]).strip()
        if not no:
            continue
        qty = f(r[h_t['数量（吨）']]) or 0.0
        if way in neg_way:
            qty = -qty
        rec = tax.setdefault(no, {'date': to_date(r[h_t['开具日期']]), 'qty': 0.0, 'ways': []})
        rec['qty'] += qty
        rec['ways'].append(way)

    # 日期窗口: 起点=名义取数窗口起点的最大者; 终点=实际数据最大日期的最小者
    r_dates = [v for k, v in res.items() if not k.endswith('_qty')]
    t_dates = [v['date'] for v in tax.values()]
    w_min = max(tax_window[0], min(r_dates))
    w_max = min(max(r_dates), max(t_dates))
    in_win = lambda d: w_min <= d <= w_max

    res_w = {k: res[k + '_qty'] for k in res if not k.endswith('_qty') and in_win(res[k])}
    tax_w = {k: v['qty'] for k, v in tax.items() if in_win(v['date'])}

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
            note = f'结果文件中无此{category}发票'
        date_s = (res.get(no) or tax.get(no, {}).get('date')).isoformat()
        detail.append([no, date_s, res_sub.get(no, ''), a, b,
                       round(a - b, 4) if (a is not None and b is not None) else None,
                       status, '&'.join(tax.get(no, {}).get('ways', [])), note])
    detail.sort(key=lambda d: (d[6] == '一致', d[1], d[0]))

    def T(d):
        return round(sum(d.values()), 4)

    summary = [
        [f'结果文件(销项){category}实际数据范围', f'{min(r_dates)} ~ {max(r_dates)}'],
        ['税局取数窗口', f'{tax_window[0]} ~ {tax_window[1]}'],
        [f'税局{category}实际数据范围(仅销项变动后)', f'{min(t_dates)} ~ {max(t_dates)}'],
        ['比对窗口', f'{w_min} ~ {w_max}'],
        ['窗口内发票数: 结果文件', len(res_w)],
        ['窗口内发票数: 税局', len(tax_w)],
        ['双方都有', sum(1 for d in detail if d[6] != '仅结果文件' and d[6] != '仅税局文件')],
        ['  其中数量一致', sum(1 for d in detail if d[6] == '一致')],
        ['  其中数量差异', sum(1 for d in detail if d[6] == '数量差异')],
        ['仅结果文件有', sum(1 for d in detail if d[6] == '仅结果文件'),
         f"（其中已预制 {sum(1 for d in detail if d[6] == '仅结果文件' and res_pre.get(d[0]))} 张）"],
        ['仅税局文件有', sum(1 for d in detail if d[6] == '仅税局文件')],
        [f'税局被剔除的{category}行数(勾选类, 属进项)', n_excluded],
        [f'{category}数量合计-结果文件', T(res_w), '吨'],
        [f'{category}数量合计-税局', T(tax_w), '吨'],
        ['数量差异合计(结果-税局)', round(T(res_w) - T(tax_w), 4), '吨'],
    ]

    _write_compare_workbook(dst, summary, detail,
                            ['发票号码', '开票日期', f'细类（结果文件）', '结果文件数量', '税局数量',
                             '差异(结果-税局)', '比对结果', '税局变动方式', '说明'],
                            [24, 12, 26, 14, 14, 15, 12, 16, 36])
    return summary, detail


# ============ 输出 Excel ============
def _write_compare_workbook(dst, summary, detail, detail_header, widths):
    out = openpyxl.Workbook()
    ws1 = out.active
    ws1.title = '汇总'
    for row in summary:
        ws1.append(list(row))
    for c in ws1['A']:
        if c.value:
            c.font = Font(bold=True)
    ws1.column_dimensions['A'].width = 52
    ws1.column_dimensions['B'].width = 32
    ws1.column_dimensions['C'].width = 26

    ws2 = out.create_sheet('逐票比对')
    ws2.append(detail_header)
    for c in ws2[1]:
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(horizontal='center')
    status_col = detail_header.index('比对结果')
    for row in detail:
        ws2.append(row)
        st = row[status_col]
        fill = MISS_FILL if st.startswith('缺失') or st.startswith('仅') else (DIFF_FILL if st.startswith('差异') else None)
        if fill:
            for c in ws2[ws2.max_row]:
                c.fill = fill
    for i, w in enumerate(widths, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = 'A2'
    ws2.auto_filter.ref = ws2.dimensions
    out.save(dst)


def format_stats(title, summary):
    buf = io.StringIO()
    print('==', title, file=buf)
    for row in summary:
        print('  ', row[0], ':', row[1] if len(row) < 3 else f'{row[1]} {row[2]}', file=buf)
    return buf.getvalue()
