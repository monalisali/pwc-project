# -*- coding: utf-8 -*-
"""
TMP日志 vs 税局 变动差异比对 (独立工具, 不并入 run_pipeline)
=============================================================
背景: TMP 成品油初始值(G列) 与 税局初始值(H列) 天然有差异。若两者期间变动一致,
则 (TMP期末 - 税局期末) 应始终等于初始差异。实际 TMP 日志变动与税局变动不一致,
本工具定位差异发生在哪里。

口径:
- TMP侧: TMP日志的"变动值"列不完整(部分为空/重复回放), 故用【相邻更新剩余库存之差】
  还原变动 (余额上升=入库+, 下降=出库-)。期初锚点 = 期初日之后余额==G列初始值的
  第一个事件 (TMP 在该时点按 G列 校准; 锚点前的波动视为期初整理, 不计变动)
- 税局侧: 税局文件按"商品和服务名称"=大类取行, 发票号码经成品油判定文件映射细类
  (进项判定+销项判定), 只取指定细类; 方向: 勾选/销售红冲=入库+, 开具/已勾选红冲=出库-
- 配对: TMP事件与税局事件按【数量相同(容差1e-3)且日期相差<=5天】贪婪配对,
  配上=日期错位/当日一致, 配不上=单方变动 (TMP 23:50 的小额变动为盘点/损耗调整)

用法 (在 energy/output 下):
  python -I tmp_tax_diff.py                 # 默认: 92号车用汽油（VIB）
  可选参数:
    --tmp-log 文件名   TMP日志 (相对 energy/, 默认 TMP日志_92号车用汽油（VIB）.xlsx)
    --cat 大类         (默认 汽油)
    --sub 细类         (默认 92号车用汽油（VIB）)
    --g-key 关键词     初始数据G列行匹配 (默认 92#小计)
    --h-key 关键词     初始数据H列行匹配 (默认同 g-key; H列取不到则留空)
    --start 日期       税局取数起始日 (默认 2026-07-16)
    --g-init/--h-init  手工指定初始值, 覆盖文件读取
输出: output/tmp税局变动差异_{细类}_{yyyymmddhhmmss}.xlsx
  汇总 / 逐日比对 / 差异配对 / TMP流水 / 税局流水
"""
import argparse
import glob
import os
import sys
from datetime import datetime, date, timedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # .../energy
OUT = os.path.join(BASE, 'output')
TOL = 1e-3
PAIR_DAYS = 5          # 配对允许的最大日期差
COUNT_TOL = 25.0       # 视为"小额盘点调整"的数量上限 (TMP 23:50 日结调整)

HDR_FONT = Font(bold=True, color='FFFFFF')
HDR_FILL = PatternFill('solid', fgColor='4472C4')
DIFF_FILL = PatternFill('solid', fgColor='FCE4EC')   # 有差异 淡红
TMP_ONLY_FILL = PatternFill('solid', fgColor='FFF3E0')   # 单方变动 淡橙
OK_FILL = PatternFill('solid', fgColor='E2EFDA')     # 一致 淡绿

# 税局变动方式 -> 方向 (+1 入库 / -1 出库)
TAX_DIR = {'增值税发票勾选': 1, '销售红冲': 1, '增值税发票开具': -1, '已勾选红冲': -1}


def p(*a):
    print(*a, file=sys.stderr)


def to_dt(v):
    if isinstance(v, datetime):
        return v
    return datetime.strptime(str(v)[:10], '%Y-%m-%d')


def f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# ---------------- TMP 日志 ----------------
def load_tmp_log(path, g_init, start):
    """余额序列 -> (锚点索引, 事件列表[{time,delta,bal}])"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rws = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0] is not None]
    rws.sort(key=lambda r: r[0])
    # 锚点: 期初日之后余额==G列的第一个事件; 取不到则退回期初日前最后一个事件
    anchor = next((i for i, r in enumerate(rws)
                   if r[0] >= datetime.combine(start, datetime.min.time())
                   and g_init is not None and abs((r[1] or 0) - g_init) < TOL), None)
    anchor_note = ''
    if anchor is None:
        before = [i for i, r in enumerate(rws) if r[0] < datetime.combine(start, datetime.min.time())]
        anchor = before[-1] if before else 0
        anchor_note = (f'未找到余额==G列({g_init})的事件, 退用期初日前最后事件; '
                       f'锚点余额={rws[anchor][1]}, 与G列差={round((rws[anchor][1] or 0) - (g_init or 0), 3)}')
    else:
        anchor_note = f'锚点余额==G列({g_init}), 该时点前波动视为期初整理'
    evs = []
    for a, b in zip(rws[anchor:], rws[anchor + 1:]):
        d = round((b[1] or 0) - (a[1] or 0), 6)
        if abs(d) > 0.0005:
            evs.append({'time': b[0], 'delta': d, 'bal': b[1], 'raw': b[2]})
    return rws[anchor], evs, anchor_note


# ---------------- 税局 + 判定文件 ----------------
def load_det_map(paths):
    """判定文件(进项+销项) 发票号码 -> 细类"""
    m = {}
    for path in paths:
        if not path or not os.path.exists(path):
            continue
        wb = openpyxl.load_workbook(path, data_only=True)
        rows = list(wb['发票汇总判定'].iter_rows(values_only=True))
        h = {c: i for i, c in enumerate(rows[0])}
        for r in rows[1:]:
            no = str(r[h['发票号码']]).strip()
            if no:
                m[no] = str(r[h['油品细分（细类）']])
    return m


def load_tax(path, cat, sub, det, start):
    wb = openpyxl.load_workbook(path, data_only=True)
    rows = list(wb['增值税发票'].iter_rows(values_only=True))
    h = {c: i for i, c in enumerate(rows[0])}
    evs, n_unmapped, n_other_sub = [], 0, 0
    for r in rows[1:]:
        if str(r[h['商品和服务名称']]).strip() != cat:
            continue
        d = to_dt(r[h['开具日期']])
        if d < datetime.combine(start, datetime.min.time()):
            continue
        no = str(r[h['数电发票号码']] or r[h['发票号码']]).strip()
        sub_hit = det.get(no)
        if sub_hit is None:
            n_unmapped += 1
        if sub_hit != sub:
            n_other_sub += 1
            continue
        way = str(r[h['变动方式']]).strip()
        qty = f(r[h['数量（吨）']]) or 0.0
        evs.append({'date': d.date(), 'no': no, 'way': way, 'qty': round(qty * TAX_DIR[way], 6),
                    'time': str(r[h['变动时间']] or '')})
    return evs, n_unmapped, n_other_sub


# ---------------- 配对 ----------------
def pair_events(tmp_evs, tax_evs):
    """按 数量相同(容差TOL)且日期差<=PAIR_DAYS 贪婪配对; 返回 (pairs, tmp_only, tax_only)"""
    tax_left = list(range(len(tax_evs)))
    pairs, tmp_only = [], []
    for te in sorted(tmp_evs, key=lambda e: e['time']):
        cands = []
        for i in tax_left:
            xe = tax_evs[i]
            if (te['delta'] > 0) != (xe['qty'] > 0):
                continue
            if abs(abs(te['delta']) - abs(xe['qty'])) > TOL:
                continue
            gap = abs((te['time'].date() - xe['date']).days)
            if gap <= PAIR_DAYS:
                cands.append((gap, i))
        if cands:
            _, i = min(cands)
            tax_left.remove(i)
            xe = tax_evs[i]
            gap = (te['time'].date() - xe['date']).days
            kind = '当日一致' if gap == 0 else f'日期错位({"税局晚%d天" % gap if gap > 0 else "税局早%d天" % -gap})'
            pairs.append((te, xe, kind))
        else:
            kind = '盘点/损耗调整' if (te['time'].hour == 23 and te['time'].minute == 50
                                       and abs(te['delta']) <= COUNT_TOL) else '税局无对应变动'
            tmp_only.append((te, kind))
    tax_only = [tax_evs[i] for i in tax_left]
    return pairs, tmp_only, tax_only


# ---------------- Excel 输出 ----------------
def _hdr(ws, header, widths):
    ws.append(header)
    for c in ws[1]:
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(horizontal='center')
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'


def _summary_block(ws, title, rows):
    ws.append([title])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    for label, *rest in rows:
        ws.append([label] + list(rest))
    return ws


def write_report(dst, ctx):
    (cat, sub, g_init, h_init, anchor_row, anchor_note, tmp_evs, tax_evs,
     pairs, tmp_only, tax_only, n_unmapped, n_other_sub, start) = ctx
    wb = openpyxl.Workbook()

    # --- 汇总 ---
    ws = wb.active
    ws.title = '汇总'
    t_in = round(sum(e['delta'] for e in tmp_evs if e['delta'] > 0), 3)
    t_out = round(sum(e['delta'] for e in tmp_evs if e['delta'] < 0), 3)
    t_net = round(t_in + t_out, 3)
    x_in = round(sum(e['qty'] for e in tax_evs if e['qty'] > 0), 3)
    x_out = round(sum(e['qty'] for e in tax_evs if e['qty'] < 0), 3)
    x_net = round(x_in + x_out, 3)
    move_diff = round(t_net - x_net, 3)
    t_end = round((g_init or 0) + t_net, 3)
    x_end = round((g_init or 0) + x_net, 3)
    tmp_span = f"{tmp_evs[0]['time']:%Y-%m-%d} ~ {tmp_evs[-1]['time']:%Y-%m-%d}" if tmp_evs else '-'
    tax_span = f"{min(e['date'] for e in tax_evs)} ~ {max(e['date'] for e in tax_evs)}" if tax_evs else '-'
    rows = [
        ('比对对象', f'{cat} / {sub}'),
        ('期初日(税局取数起始日)', start),
        ('TMP期初(G列)', g_init or '空', '来源: Energy初始数据.xlsx G列'),
        ('税局期初(H列)', h_init if h_init is not None else '空(初始数据H列该细类无值, H列仅大类合计有值)'),
        ('初始差异(TMP-税局)', round((g_init or 0) - (h_init or 0), 3) if h_init is not None else 'H列缺失, 无法计算'),
        ('', ''),
        ('TMP锚点事件', f"{anchor_row[0]:%Y-%m-%d %H:%M:%S} 余额={anchor_row[1]}", anchor_note),
        ('TMP变动窗口', tmp_span),
        ('税局变动窗口', tax_span),
        ('', ''),
        ('TMP 入库合计', t_in, '吨 (由余额增量还原)'),
        ('TMP 出库合计', t_out, '吨'),
        ('TMP 净变动', t_net, '吨'),
        ('税局 入库合计', x_in, '吨 (勾选+销售红冲, 按判定文件映射细类)'),
        ('税局 出库合计', x_out, '吨 (开具+已勾选红冲)'),
        ('税局 净变动', x_net, '吨'),
        ('变动差异 (TMP-税局)', move_diff, '吨  <<< 核心结果'),
        ('', ''),
        ('期末推算-TMP口径 (G列+TMP净变动)', t_end, '应等于TMP日志最末余额'),
        ('期末推算-税局口径 (G列+税局净变动)', x_end, '若H列税局期初可得, 期末差异=初始差异+变动差异'),
        ('期末差异 (TMP口径-税局口径)', round(t_end - x_end, 3), '== 变动差异 (期初同取G列)'),
        ('', ''),
        ('配对成功事件', len(pairs), '其中 当日一致 %d, 日期错位 %d' % (
            sum(1 for _, _, k in pairs if k == '当日一致'),
            sum(1 for _, _, k in pairs if k.startswith('日期错位')))),
        ('TMP单方变动 (税局无对应)', len(tmp_only)),
        ('税局单方变动 (TMP无对应)', len(tax_only)),
        ('税局行未映射到细类(未计入)', n_unmapped),
        ('税局行属其他细类(未计入)', n_other_sub),
    ]
    for row in rows:
        ws.append(list(row))
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value:
            ws.cell(r, 1).font = Font(bold=True)
        if ws.cell(r, 1).value == '变动差异 (TMP-税局)':
            for c in ws[r]:
                c.fill = DIFF_FILL
    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 52

    # --- 逐日比对 ---
    ws2 = wb.create_sheet('逐日比对')
    _hdr(ws2, ['日期', 'TMP净变动', '税局净变动', '差异(TMP-税局)', '累计差异', '当日配对情况'],
         [12, 13, 13, 14, 12, 60])
    tmp_day, tax_day = {}, {}
    for e in tmp_evs:
        tmp_day[e['time'].date()] = round(tmp_day.get(e['time'].date(), 0) + e['delta'], 3)
    for e in tax_evs:
        tax_day[e['date']] = round(tax_day.get(e['date'], 0) + e['qty'], 3)
    pair_by_day = {}
    for te, xe, kind in pairs:
        pair_by_day.setdefault(te['time'].date(), []).append(
            f"{te['delta']:+.3f}↔{xe['date']}({xe['way']},{kind})")
    cum = 0.0
    for d in sorted(set(tmp_day) | set(tax_day)):
        t, x = tmp_day.get(d, 0), tax_day.get(d, 0)
        diff = round(t - x, 3)
        cum = round(cum + diff, 3)
        note = '；'.join(pair_by_day.get(d, []))
        if diff and not note:
            note = '当日变动未配对, 详见 差异配对'
        ws2.append([d, t, x, diff, cum, note])
        st = ws2[ws2.max_row]
        fill = OK_FILL if abs(diff) < TOL else (DIFF_FILL if note else TMP_ONLY_FILL)
        for c in st:
            c.fill = fill

    # --- 差异配对 ---
    ws3 = wb.create_sheet('差异配对')
    _hdr(ws3, ['类型', 'TMP时间', 'TMP变动', 'TMP余额', '税局日期', '税局发票号码',
               '税局变动方式', '税局数量', '天数差', '说明'], [16, 20, 12, 12, 12, 24, 14, 11, 8, 30])
    for te, xe, kind in sorted(pairs, key=lambda x: x[0]['time']):
        ws3.append([kind, te['time'], te['delta'], te['bal'], xe['date'], xe['no'],
                    xe['way'], xe['qty'], (te['time'].date() - xe['date']).days, ''])
    for te, kind in sorted(tmp_only, key=lambda x: x[0]['time']):
        ws3.append([f'TMP单方({kind})', te['time'], te['delta'], te['bal'],
                    '', '', '', '', '', '税局无对应记录'])
    for xe in sorted(tax_only, key=lambda x: x['date']):
        ws3.append(['税局单方', '', '', '', xe['date'], xe['no'], xe['way'], xe['qty'],
                    '', 'TMP日志无对应变动'])
    for row in ws3.iter_rows(min_row=2):
        if str(row[0].value).startswith('TMP单方'):
            for c in row:
                c.fill = TMP_ONLY_FILL
        elif str(row[0].value).startswith('税局单方'):
            for c in row:
                c.fill = DIFF_FILL

    # --- TMP流水 ---
    ws4 = wb.create_sheet('TMP流水')
    _hdr(ws4, ['时间', '变动(余额增量)', '更新后余额', '日志变动值列', '方向'], [22, 14, 13, 13, 8])
    for e in tmp_evs:
        ws4.append([e['time'], e['delta'], e['bal'], e['raw'], '入库' if e['delta'] > 0 else '出库'])

    # --- 税局流水 ---
    ws5 = wb.create_sheet('税局流水')
    _hdr(ws5, ['开具日期', '发票号码', '变动方式', '数量(方向后)', '变动时间', '方向'], [12, 24, 14, 13, 20, 8])
    for e in sorted(tax_evs, key=lambda x: (x['date'], x['no'])):
        ws5.append([e['date'], e['no'], e['way'], e['qty'], e['time'], '入库' if e['qty'] > 0 else '出库'])

    wb.save(dst)


# ---------------- 主流程 ----------------
def find_det_file(prefix):
    """优先无时间戳判定文件, 否则取最新时间戳版本"""
    fp = os.path.join(OUT, prefix)
    if os.path.exists(fp):
        return fp
    hits = sorted(glob.glob(os.path.join(OUT, prefix.replace('.xlsx', '') + '_*.xlsx')))
    return hits[-1] if hits else None


def run(args):
    # 初始数据: G/H 列按关键词匹配行 (品名/规格包含全部关键词)
    wb = openpyxl.load_workbook(os.path.join(BASE, args.init), data_only=True)
    rows = list(wb[wb.sheetnames[0]].iter_rows(values_only=True))

    def find_col(keys, col):
        if not keys:
            return None
        for r in rows[1:]:
            text = ' '.join(str(v) for v in r[:2] if v)
            if all(k.lstrip('!') in text for k in keys):
                return f(r[col])
        return None

    g_keys = [k for k in args.g_key.split(',') if k.strip()]
    h_keys = [k for k in args.h_key.split(',') if k.strip()] if args.h_key else g_keys
    g_init = args.g_init if args.g_init is not None else find_col(g_keys, 6)
    h_init = args.h_init if args.h_init is not None else find_col(h_keys, 7)
    if g_init is None:
        p('!! G列初始值未取到 (关键词: %s), TMP锚点将退用期初日前最后事件' % args.g_key)

    start = date.fromisoformat(args.start)
    anchor_row, tmp_evs, anchor_note = load_tmp_log(
        os.path.join(BASE, args.tmp_log) if not os.path.isabs(args.tmp_log) else args.tmp_log,
        g_init, start)
    det = load_det_map([find_det_file('成品油发票判定_Energy进项20260501-20261026.xlsx'),
                        find_det_file('成品油发票判定_Energy销项.xlsx')])
    tax_path = os.path.join(BASE, args.tax) if not os.path.isabs(args.tax) else args.tax
    tax_evs, n_unmapped, n_other_sub = load_tax(tax_path, args.cat, args.sub, det, start)
    pairs, tmp_only, tax_only = pair_events(tmp_evs, tax_evs)

    ctx = (args.cat, args.sub, g_init, h_init, anchor_row, anchor_note, tmp_evs, tax_evs,
           pairs, tmp_only, tax_only, n_unmapped, n_other_sub, start)
    dst = os.path.join(OUT, 'tmp税局变动差异_%s_%s.xlsx' % (
        args.sub, datetime.now().strftime('%Y%m%d%H%M%S')))
    write_report(dst, ctx)

    t_net = round(sum(e['delta'] for e in tmp_evs), 3)
    x_net = round(sum(e['qty'] for e in tax_evs), 3)
    p('比对对象: %s / %s' % (args.cat, args.sub))
    p('TMP期初(G列)=%s, 税局期初(H列)=%s' % (g_init, h_init if h_init is not None else '空'))
    p('锚点: %s %s (%s)' % (anchor_row[0], anchor_row[1], anchor_note))
    p('TMP 净变动=%s | 税局 净变动=%s | 变动差异(TMP-税局)=%s' % (t_net, x_net, round(t_net - x_net, 3)))
    p('配对 %d (当日一致 %d / 日期错位 %d), TMP单方 %d, 税局单方 %d' % (
        len(pairs), sum(1 for _, _, k in pairs if k == '当日一致'),
        sum(1 for _, _, k in pairs if k.startswith('日期错位')), len(tmp_only), len(tax_only)))
    p('报告:', dst)


def main():
    ap = argparse.ArgumentParser(description='TMP日志 vs 税局 变动差异比对 (独立工具)')
    ap.add_argument('--tmp-log', default='TMP日志_92号车用汽油（VIB）.xlsx')
    ap.add_argument('--tax', default='税局Energy增值税发票7.1-10.23.xlsx')
    ap.add_argument('--init', default='Energy初始数据.xlsx')
    ap.add_argument('--cat', default='汽油')
    ap.add_argument('--sub', default='92号车用汽油（VIB）')
    ap.add_argument('--g-key', default='92#小计', help='G列(TMP初始)行匹配关键词, 逗号分隔')
    ap.add_argument('--h-key', default='', help='H列(税局初始)行匹配关键词, 缺省同 g-key')
    ap.add_argument('--start', default='2026-07-16')
    ap.add_argument('--g-init', type=float, default=None, help='手工指定TMP初始值')
    ap.add_argument('--h-init', type=float, default=None, help='手工指定税局初始值')
    run(ap.parse_args())


if __name__ == '__main__':
    main()
