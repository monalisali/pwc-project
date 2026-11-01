# -*- coding: utf-8 -*-
"""
库存计算台账 vs TMP系统日志 逐笔比对 (独立工具, 不并入 run_pipeline)
=====================================================================
背景: run_pipeline 生成的 库存计算_细类_{大类}.xlsx 中某细类的剩余库存,
与 TMP 系统日志的余额差距较大。本工具把两边逐笔还原成变动事件并配对,
定位差异出在哪些笔上。

口径:
- 台账侧: 流水_{细类} sheet, 数量变动(+进/-销) 即事件; 初始行给出期初;
  尾部"以下为已剔除数据"之后的行(已预制剔除票等)不参与配对, 单独列出
- TMP侧: 日志"变动值"列不完整, 用【相邻更新剩余库存之差】还原变动
  (余额上升=入库+, 下降=出库-); 锚点 = 起始日后余额==台账初始值的第一个事件
  (TMP 在该时点按初始值校准, 锚点前波动视为期初整理)
- 配对: 数量相同(容差1e-3)且日期差<=5天, 贪婪配对; 配不上即为单方变动
  - TMP单方分类: 23:50 且小额<=25吨 = 盘点/损耗调整;
    与已配对TMP事件同数量 = 疑似系统重复回放; 其余 = 无对应记录
- 恒等式: 变动差异(台账净-TMP净) = 台账单方合计 - TMP单方合计 (汇总中互相验证)

用法 (在 energy/output 下):
  python -I ledger_vs_tmplog.py           # 启动后逐项询问, 回车=默认值
  python -I ledger_vs_tmplog.py --batch   # 不询问, 直接用默认值/命令行参数
  可选参数 (作为询问时的默认值):
    --ledger  流水计算文件完整路径 (默认 output/库存计算/库存计算_细类_汽油.xlsx, 相对 energy/)
    --sheet   流水sheet名          (默认 流水_92号车用汽油（VIB）)
    --tmp-log TMP日志文件完整路径  (默认 TMP日志_92号车用汽油（VIB）.xlsx, 相对 energy/)
    --start   取数起始日           (默认 2026-07-16, 含当天)
    --end     取数截止日           (默认 2026-10-23, 含当天)
输出: output/库存比对/日志比对/台账vsTMP日志_{细类}_{yyyymmddhhmmss}.xlsx, 结束时打印完整路径
  汇总 / 逐笔配对 / TMP单方变动 / 台账单方变动 / 台账已剔除票 / TMP流水 / 台账流水
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
OUT_DIR = os.path.join(OUT, '库存比对', '日志比对')   # 报告输出目录

TOL = 1e-3
PAIR_DAYS = 5
COUNT_TOL = 25.0

HDR_FONT = Font(bold=True, color='FFFFFF')
HDR_FILL = PatternFill('solid', fgColor='4472C4')
DIFF_FILL = PatternFill('solid', fgColor='FCE4EC')   # 台账单方 淡红
TMP_ONLY_FILL = PatternFill('solid', fgColor='FFF3E0')   # TMP单方 淡橙
OK_FILL = PatternFill('solid', fgColor='E2EFDA')     # 一致 淡绿
RED_FONT = Font(color='FF0000')                     # 重点数字标红


def p(*a):
    print(*a, file=sys.stderr)


def resolve(path):
    return path if os.path.isabs(path) else os.path.join(BASE, path)


def f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return datetime.strptime(str(v)[:10], '%Y-%m-%d').date()


# ---------------- 台账侧 ----------------
def load_ledger(path, sheet):
    """读流水 sheet -> (初始值, 事件列表, 已剔除行列表, 期末结余)"""
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet not in wb.sheetnames:
        raise SystemExit('!! 文件中找不到 sheet: %s (现有: %s)' % (sheet, wb.sheetnames))
    rows = list(wb[sheet].iter_rows(values_only=True))
    evs, excluded, initial, last_bal = [], [], None, None
    in_excluded = False
    for r in rows[1:]:
        c0 = str(r[0]) if r[0] is not None else ''
        c1 = str(r[1]).strip() if r[1] is not None else ''
        if c0.startswith('以下为已剔除'):
            in_excluded = True
            continue
        qty = f(r[8])
        if in_excluded:
            if qty is not None:
                excluded.append({'date': to_date(r[0]), 'dir': c1, 'no': str(r[2] or ''),
                                 'qty': qty, 'note': str(r[10] or '')})
            continue
        if not c1 or qty is None:
            continue
        if c1 == '初始':
            initial = qty
            continue
        evs.append({'date': to_date(r[0]), 'dir': c1, 'no': str(r[2] or ''),
                    'seller': str(r[3] or ''), 'buyer': str(r[4] or ''),
                    'qty': qty, 'bal': f(r[9])})
        last_bal = f(r[9]) or last_bal
    return initial, evs, excluded, last_bal


# ---------------- TMP侧 ----------------
def load_tmp_log(path, init_val, start):
    """余额序列 -> (锚点行, 事件列表, 锚点说明)"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rws = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0] is not None]
    rws.sort(key=lambda r: r[0])
    anchor = next((i for i, r in enumerate(rws)
                   if r[0] >= datetime.combine(start, datetime.min.time())
                   and init_val is not None and abs((r[1] or 0) - init_val) < TOL), None)
    if anchor is None:
        before = [i for i, r in enumerate(rws) if r[0] < datetime.combine(start, datetime.min.time())]
        anchor = before[-1] if before else 0
        note = ('未找到余额==初始值(%s)的事件, 退用起始日前最后事件; 锚点余额=%s, 与初始值差=%s'
                % (init_val, rws[anchor][1], round((rws[anchor][1] or 0) - (init_val or 0), 3)))
    else:
        note = '锚点余额==台账初始值(%s), 该时点前波动视为期初整理' % init_val
    evs = []
    for a, b in zip(rws[anchor:], rws[anchor + 1:]):
        d = round((b[1] or 0) - (a[1] or 0), 6)
        if abs(d) > 0.0005:
            evs.append({'time': b[0], 'delta': d, 'bal': b[1], 'raw': b[2]})
    return rws[anchor], evs, note


# ---------------- 配对 ----------------
def pair_events(led_evs, tmp_evs):
    """按 数量相同(容差TOL)且日期差<=PAIR_DAYS 贪婪配对 -> (pairs, led_only, tmp_only)"""
    tmp_left = list(range(len(tmp_evs)))
    pairs, led_only = [], []
    for le in sorted(led_evs, key=lambda e: (e['date'], e['no'])):
        cands = []
        for i in tmp_left:
            te = tmp_evs[i]
            if (le['qty'] > 0) != (te['delta'] > 0):
                continue
            if abs(abs(le['qty']) - abs(te['delta'])) > TOL:
                continue
            gap = abs((le['date'] - te['time'].date()).days)
            if gap <= PAIR_DAYS:
                cands.append((gap, i))
        if cands:
            _, i = min(cands)
            tmp_left.remove(i)
            te = tmp_evs[i]
            gap = (te['time'].date() - le['date']).days
            kind = '当日一致' if gap == 0 else ('TMP晚%d天' % gap if gap > 0 else 'TMP早%d天' % -gap)
            pairs.append((le, te, kind))
        else:
            led_only.append(le)
    tmp_only = [tmp_evs[i] for i in tmp_left]
    # TMP单方分类: 与已配对TMP事件同数量 -> 疑似系统重复回放; 23:50小额 -> 盘点/损耗调整
    paired_qtys = [round(abs(te['delta']), 3) for _, te, _ in pairs]
    out = []
    for te in tmp_only:
        if te['time'].hour == 23 and te['time'].minute == 50 and abs(te['delta']) <= COUNT_TOL:
            kind = 'gain/loss'
        elif any(abs(round(abs(te['delta']), 3) - q) <= TOL for q in paired_qtys):
            kind = '疑似系统重复回放'
        else:
            kind = '无对应记录'
        out.append((te, kind))
    return pairs, led_only, out


# ---------------- Excel ----------------
def _hdr(ws, header, widths):
    ws.append(header)
    for c in ws[1]:
        c.font, c.fill = HDR_FILL and HDR_FONT, HDR_FILL
        c.alignment = Alignment(horizontal='center')
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'


def write_report(dst, sub, initial, led_evs, led_excl, led_last, anchor_row, anchor_note,
                 tmp_evs, pairs, led_only, tmp_only, start, end):
    wb = openpyxl.Workbook()

    # --- 汇总 ---
    ws = wb.active
    ws.title = '汇总'
    ws['A1'] = '库存计算台账 vs TMP日志 逐笔比对 (细类: %s; 窗口 %s ~ %s 含当天)' % (sub, start, end)
    ws['A1'].font = Font(bold=True)
    l_in = round(sum(e['qty'] for e in led_evs if e['qty'] > 0), 3)
    l_out = round(sum(e['qty'] for e in led_evs if e['qty'] < 0), 3)
    l_net = round(l_in + l_out, 3)
    t_in = round(sum(e['delta'] for e in tmp_evs if e['delta'] > 0), 3)
    t_out = round(sum(e['delta'] for e in tmp_evs if e['delta'] < 0), 3)
    t_net = round(t_in + t_out, 3)
    tmp_only_sum = round(sum(te['delta'] for te, _ in tmp_only), 3)
    led_only_sum = round(sum(e['qty'] for e in led_only), 3)
    move_diff = round(l_net - t_net, 3)
    check = round(led_only_sum - tmp_only_sum, 3)
    rows = [
        ('台账初始值 (流水初始行)', initial),
        ('TMP锚点事件', '%s 余额=%s' % (anchor_row[0], anchor_row[1]), anchor_note),
        ('', ''),
        ('台账 进项合计(+)', l_in, '吨'),
        ('台账 销项合计(-)', l_out, '吨'),
        ('台账 净变动', l_net, '吨'),
        ('台账 期末结余', led_last, '吨'),
        ('', ''),
        ('TMP 入库合计(+)', t_in, '吨 (余额增量还原)'),
        ('TMP 出库合计(-)', t_out, '吨'),
        ('TMP 净变动', t_net, '吨'),
        ('TMP 最末余额', tmp_evs[-1]['bal'] if tmp_evs else '-', '吨'),
        ('', ''),
        ('变动差异 (台账净-TMP净)', move_diff, '吨  <<< 核心结果'),
        ('期末差异 (台账期末-TMP最末余额)', round((led_last or 0) - (tmp_evs[-1]['bal'] if tmp_evs else 0), 3), '吨'),
        ('', ''),
        ('配对成功笔数', len(pairs), '其中 当日一致 %d, 日期错位 %d' % (
            sum(1 for _, _, k in pairs if k == '当日一致'),
            sum(1 for _, _, k in pairs if k != '当日一致'))),
        ('TMP单方笔数 / 合计', '%d / %s' % (len(tmp_only), tmp_only_sum), '吨'),
        ('台账单方笔数 / 合计', '%d / %s' % (len(led_only), led_only_sum), '吨'),
        ('验证: 台账单方合计-TMP单方合计', check, '应等于 变动差异 %s' % move_diff),
        ('', ''),
        ('TMP单方-疑似系统重复回放', '%d 笔 / %s 吨' % (
            sum(1 for _, k in tmp_only if k == '疑似系统重复回放'),
            round(sum(te['delta'] for te, k in tmp_only if k == '疑似系统重复回放'), 3))),
        ('TMP单方-gain/loss', '%d 笔 / %s 吨' % (
            sum(1 for _, k in tmp_only if k == 'gain/loss'),
            round(sum(te['delta'] for te, k in tmp_only if k == 'gain/loss'), 3))),
        ('台账已剔除票(未计入台账, TMP若有扣减即为差异)', '%d 笔 / %s 吨' % (
            len(led_excl), round(sum(e['qty'] for e in led_excl), 3))),
    ]
    for row in rows:
        ws.append(list(row))
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value:
            ws.cell(r, 1).font = Font(bold=True)
        label = str(ws.cell(r, 1).value or '')
        if label.startswith(('变动差异', '期末差异', '验证')):
            for c in ws[r]:
                c.fill = DIFF_FILL
        if label.startswith(('台账 期末结余', 'TMP 最末余额', 'TMP单方-gain/loss')):
            ws.cell(r, 2).font = RED_FONT   # 重点数字标红
    ws.column_dimensions['A'].width = 42
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 48

    # --- 逐笔配对 ---
    ws2 = wb.create_sheet('逐笔配对')
    _hdr(ws2, ['配对结果', '台账日期', '台账方向', '台账发票号码', '台账数量', 'TMP时间',
               'TMP变动', 'TMP余额', '天数差', '说明'], [10, 12, 8, 24, 12, 20, 11, 12, 8, 14])
    for le, te, kind in sorted(pairs, key=lambda x: x[1]['time']):
        ws2.append([kind, le['date'], le['dir'], le['no'], le['qty'], te['time'],
                    te['delta'], te['bal'], (te['time'].date() - le['date']).days, ''])

    # --- TMP单方变动 ---
    ws3 = wb.create_sheet('TMP单方变动')
    _hdr(ws3, ['分类', 'TMP时间', '变动', '更新后余额', '日志变动值列', '说明'],
         [18, 22, 12, 13, 13, 24])
    for te, kind in sorted(tmp_only, key=lambda x: x[0]['time']):
        ws3.append([kind, te['time'], te['delta'], te['bal'], te['raw'], '台账无对应记录'])
        for c in ws3[ws3.max_row]:
            c.fill = TMP_ONLY_FILL

    # --- 台账单方变动 ---
    ws4 = wb.create_sheet('台账单方变动')
    _hdr(ws4, ['台账日期', '方向', '发票号码', '销方名称', '购方名称', '数量变动', '说明'],
         [12, 8, 24, 30, 30, 12, 26])
    for le in sorted(led_only, key=lambda e: e['date']):
        ws4.append([le['date'], le['dir'], le['no'], le['seller'], le['buyer'], le['qty'],
                    'TMP日志无对应变动'])
        for c in ws4[ws4.max_row]:
            c.fill = DIFF_FILL

    # --- 台账已剔除票 ---
    ws5 = wb.create_sheet('台账已剔除票')
    _hdr(ws5, ['日期', '方向', '发票号码', '数量', '备注'], [12, 16, 24, 12, 24])
    for e in led_excl:
        ws5.append([e['date'], e['dir'], e['no'], e['qty'], e['note']])

    # --- 流水 ---
    ws6 = wb.create_sheet('TMP流水')
    _hdr(ws6, ['时间', '变动(余额增量)', '更新后余额', '日志变动值列', '方向'], [22, 14, 13, 13, 8])
    for e in tmp_evs:
        ws6.append([e['time'], e['delta'], e['bal'], e['raw'], '入库' if e['delta'] > 0 else '出库'])
    ws7 = wb.create_sheet('台账流水')
    _hdr(ws7, ['日期', '方向', '发票号码', '销方名称', '购方名称', '数量变动', '累计结余'],
         [12, 8, 24, 30, 30, 12, 13])
    for e in led_evs:
        ws7.append([e['date'], e['dir'], e['no'], e['seller'], e['buyer'], e['qty'], e['bal']])

    wb.save(dst)


# ---------------- 启动交互 ----------------
def ask(prompt, default):
    try:
        v = input('%s (回车=%s): ' % (prompt, default)).strip()
    except EOFError:
        v = ''
    return v or default


def ask_file(prompt, default):
    while True:
        v = ask(prompt, default)
        fp = resolve(v)
        if os.path.exists(fp):
            return v
        p('!! 文件不存在: %s , 请重新输入' % fp)


def ask_date(prompt, default):
    while True:
        v = ask(prompt + ' (yyyy-mm-dd)', default)
        try:
            date.fromisoformat(v)
            return v
        except ValueError:
            p('!! 日期格式不正确: %s , 请输入 yyyy-mm-dd' % v)


def prompt_inputs(args):
    p('=' * 64)
    p('库存计算台账 vs TMP日志 逐笔比对 - 请输入参数 (直接回车使用括号内默认值)')
    p('=' * 64)
    args.ledger = ask_file('流水计算文件完整路径', args.ledger)
    args.sheet = ask('流水sheet名', args.sheet)
    args.tmp_log = ask_file('TMP日志文件完整路径', args.tmp_log)
    args.start = ask_date('取数起始日', args.start)
    args.end = ask_date('取数截止日', args.end)


# ---------------- 主流程 ----------------
def run(args):
    start = date.fromisoformat(args.start)
    end = date.fromisoformat(args.end)
    sub = args.sheet[len('流水_'):] if args.sheet.startswith('流水_') else args.sheet
    initial, led_evs, led_excl, led_last = load_ledger(resolve(args.ledger), args.sheet)
    led_evs = [e for e in led_evs if start <= e['date'] <= end]
    anchor_row, tmp_evs, anchor_note = load_tmp_log(resolve(args.tmp_log), initial, start)

    pairs, led_only, tmp_only = pair_events(led_evs, tmp_evs)

    os.makedirs(OUT_DIR, exist_ok=True)
    dst = os.path.join(OUT_DIR, '台账vsTMP日志_%s_%s.xlsx' % (
        sub, datetime.now().strftime('%Y%m%d%H%M%S')))
    write_report(dst, sub, initial, led_evs, led_excl, led_last, anchor_row, anchor_note,
                 tmp_evs, pairs, led_only, tmp_only, start, end)

    l_net = round(sum(e['qty'] for e in led_evs), 3)
    t_net = round(sum(e['delta'] for e in tmp_evs), 3)
    p('sheet: %s | 窗口: %s ~ %s (含当天)' % (args.sheet, start, end))
    p('台账初始=%s | 锚点: %s %s' % (initial, anchor_row[0], anchor_row[1]))
    p('台账净变动=%s (期末 %s) | TMP净变动=%s (最末余额 %s)' % (
        l_net, led_last, t_net, tmp_evs[-1]['bal'] if tmp_evs else '-'))
    p('变动差异(台账-TMP)=%s | 配对 %d, TMP单方 %d (%s), 台账单方 %d, 台账已剔除票 %d' % (
        round(l_net - t_net, 3), len(pairs), len(tmp_only),
        round(sum(te['delta'] for te, _ in tmp_only), 3), len(led_only), len(led_excl)))
    p('输出文件完整路径: %s' % dst)


def find_default_ledger():
    fp = os.path.join(BASE, 'output', '库存计算', '库存计算_细类_汽油.xlsx')
    return fp if os.path.exists(fp) else 'output/库存计算/库存计算_细类_汽油.xlsx'


def main():
    ap = argparse.ArgumentParser(description='库存计算台账 vs TMP日志 逐笔比对 (独立工具)')
    ap.add_argument('--ledger', default=find_default_ledger(), help='流水计算文件完整路径')
    ap.add_argument('--sheet', default='流水_92号车用汽油（VIB）', help='流水sheet名')
    ap.add_argument('--tmp-log', default='TMP日志_92号车用汽油（VIB）.xlsx', help='TMP日志文件完整路径')
    ap.add_argument('--start', default='2026-07-16', help='取数起始日 (含当天)')
    ap.add_argument('--end', default='2026-10-23', help='取数截止日 (含当天)')
    ap.add_argument('--batch', action='store_true', help='不询问, 直接用默认值/命令行参数')
    args = ap.parse_args()
    if not args.batch:
        prompt_inputs(args)
    run(args)


if __name__ == '__main__':
    main()
