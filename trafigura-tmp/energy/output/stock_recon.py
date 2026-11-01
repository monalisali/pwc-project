# -*- coding: utf-8 -*-
"""
税局 vs TMP 剩余库存核对 (独立工具, 不并入 run_pipeline)
=========================================================
逻辑 (按需求):
1. 税局剩余库存 = 初始数据H列初始值(取不到按0) + 税局文件变动合计(开具日期在窗口内)
   - 税局文件【商品和服务名称】= 大类
   - 方向: 增值税发票勾选/销售红冲 = 入库(+), 增值税发票开具/已勾选红冲 = 出库(-)
2. TMP剩余库存 = TMP成品油库存.xlsx 按【税收分类名称】(大类)合计剩余库存(吨)
   - 剔除 gain/loss 影响: TMP库存(剔GL) = 快照剩余库存 - Σ gain/loss(创建时间在窗口内)
   - gain/loss 来源 TMPGainLoss明细.xlsx【数量】, 货物名称映射大类
     (Gasoline=汽油, Gas oil=柴油, Naphtha=石脑油)
3. 校验: 库存差异(TMP剔GL - 税局) 是否等于 初始值差异(初始数据I列 = G列TMP初始 - H列税局初始)
   - 理论上 TMP库存(剔GL) = G列 + TMP非GL变动, 税局库存 = H列 + 税局变动,
     故 库存差异 - 初始差异 = 变动差异(TMP非GL变动 - 税局变动), 两边变动一致时为 0

用法 (在 energy/output 下):
  python -I stock_recon.py                # 启动后逐项询问输入文件与取数窗口, 回车=默认值
  python -I stock_recon.py --batch        # 不询问, 直接用默认值/命令行参数
  可在询问时输入, 也可用命令行参数预置默认值:
    --tax       税局发票文件   (默认 税局Energy增值税发票7.1-10.23.xlsx, 相对 energy/)
    --tmp-stock TMP库存快照    (默认 TMP成品油库存.xlsx,       相对 energy/)
    --gl        GainLoss明细   (默认 TMPGainLoss明细.xlsx,     相对 energy/)
    --init      初始数据       (默认 Energy初始数据.xlsx,      相对 energy/)
    --start     取数起始日     (默认 2026-07-16, 含当天)
    --end       取数截止日     (默认 2026-10-23, 含当天)
    --tol       判定一致的容差 (默认 0.01 吨)
输出: output/库存比对/库存差异核对_yyyymmddhhmmss.xlsx, 结束时打印完整路径
  汇总 / 税局变动明细 / GainLoss明细(窗口内) / TMP库存快照
"""
import argparse
import os
import sys
from datetime import datetime, date
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # .../energy
OUT = os.path.join(BASE, 'output')
OUT_DIR = os.path.join(OUT, '库存比对')   # 报告输出目录

# 默认文件名 (均可用命令行参数覆盖)
TAX_FILE = '税局Energy增值税发票7.1-10.23.xlsx'
TMP_STOCK_FILE = 'TMP成品油库存.xlsx'
GL_FILE = 'TMPGainLoss明细.xlsx'
INIT_FILE = 'Energy初始数据.xlsx'
DEFAULT_START = '2026-07-16'
DEFAULT_END = '2026-10-23'

TAX_DIR = {'增值税发票勾选': 1, '销售红冲': 1, '增值税发票开具': -1, '已勾选红冲': -1}
GL_NAME2CAT = {'Gasoline': '汽油', 'Gas oil': '柴油', 'Naphtha': '石脑油', 'Diesel': '柴油'}
# 初始数据中的大类行标签 (品名列含"合计"); 不在表中的大类初始按0计
CAT2INIT_ROW = {'汽油': '汽油合计', '柴油': '柴油合计', '石脑油': '石脑油合计',
                '车用乙醇汽油调和组分油': '组分油合计', '燃料油': '燃料油合计'}

HDR_FONT = Font(bold=True, color='FFFFFF')
HDR_FILL = PatternFill('solid', fgColor='4472C4')
DIFF_FILL = PatternFill('solid', fgColor='FCE4EC')   # 不一致 淡红
OK_FILL = PatternFill('solid', fgColor='E2EFDA')     # 一致 淡绿


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


def s(v):
    return '' if v is None else str(v)


# ---------------- 数据读取 ----------------
def load_init(path):
    """初始数据 -> {大类: {'G':x, 'H':y, 'I':z}} (按 合计 行)"""
    wb = openpyxl.load_workbook(path, data_only=True)
    rows = list(wb[wb.sheetnames[0]].iter_rows(values_only=True))
    init = {}
    for r in rows[1:]:
        name = s(r[0]).strip()
        if name.endswith('合计'):
            init[name] = {'G': f(r[6]), 'H': f(r[7]), 'I': f(r[8])}
    return init


def load_tax_movements(path, start, end):
    """税局文件 -> (按大类的变动明细, 大类清单); 变动 = 开具日期在 [start, end] 内的行"""
    wb = openpyxl.load_workbook(path, data_only=True)
    rows = list(wb['增值税发票'].iter_rows(values_only=True))
    h = {c: i for i, c in enumerate(rows[0])}
    detail = []
    cats = set()
    for r in rows[1:]:
        cat = s(r[h['商品和服务名称']]).strip()
        if not cat:
            continue
        d = to_dt(r[h['开具日期']])
        if d < datetime.combine(start, datetime.min.time()) or d > datetime.combine(end, datetime.max.time()):
            continue
        way = s(r[h['变动方式']]).strip()
        sign = TAX_DIR.get(way)
        if sign is None:
            sign = 0   # 未识别的变动方式: 计0并在明细中标注
        qty = f(r[h['数量（吨）']]) or 0.0
        no = s(r[h['数电发票号码']] or r[h['发票号码']]).strip()
        detail.append({'cat': cat, 'date': d.date(), 'no': no, 'way': way,
                       'qty': round(qty * sign, 6), 'dir': '入库' if sign > 0 else ('出库' if sign < 0 else '未识别')})
        cats.add(cat)
    return detail, cats


def load_tmp_stock(path):
    """TMP成品油库存 -> (快照行列表, {大类: 合计剩余库存})"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows, agg = [], defaultdict(float)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[2] is None and r[5] is None:
            continue
        cat = s(r[2]).strip()
        qty = f(r[5]) or 0.0
        rows.append({'sub': s(r[0]).strip(), 'cat': cat, 'qty': qty, 'upd': r[6]})
        agg[cat] += qty
    return rows, dict(agg)


def load_gl(path, start, end):
    """GainLoss明细 -> (窗口内明细, {大类: 合计数量}), 取创建时间在 [start, end] 内"""
    wb = openpyxl.load_workbook(path, data_only=True)
    rows = list(wb[wb.sheetnames[0]].iter_rows(min_row=2, values_only=True))
    detail, agg = [], defaultdict(float)
    for r in rows:
        if s(r[15]).strip() == 'Y':   # 删除标志
            continue
        d = to_dt(r[13])   # 创建时间
        if d < datetime.combine(start, datetime.min.time()) or d > datetime.combine(end, datetime.max.time()):
            continue
        cat = GL_NAME2CAT.get(s(r[3]).strip(), s(r[3]).strip())
        qty = f(r[10]) or 0.0
        detail.append({'sn': s(r[0]), 'cat': cat, 'sub': s(r[4]).strip(), 'qty': qty,
                       'party': s(r[7]).strip(), 'create': d})
        agg[cat] += qty
    return detail, dict(agg)


# ---------------- 报告 ----------------
def _hdr(ws, header, widths):
    ws.append(header)
    for c in ws[1]:
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(horizontal='center')
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'


def write_report(dst, cats, init, tax_detail, tmp_rows, tmp_agg, gl_detail, gl_agg, start, end, tol):
    wb = openpyxl.Workbook()

    # --- 汇总 ---
    ws = wb.active
    ws.title = '汇总'
    ws['A1'] = '税局 vs TMP 剩余库存核对 (取数窗口 %s ~ %s 含当天; TMP库存剔除Gain/Loss)' % (start, end)
    ws['A1'].font = Font(bold=True)
    header = ['大类', 'TMP初始(G列)', '税局初始(H列)', '初始差异(I列=G-H)', 'I列文件值',
              '税局入库', '税局出库', '税局净变动', '税局剩余库存(H+净变动)',
              'TMP快照剩余库存', 'GainLoss合计(窗口内)', 'TMP库存(剔除GL)',
              '库存差异(TMP剔GL-税局)', '库存差异-初始差异', '结论(差异=初始差异?)']
    _hdr(ws, header, [16, 12, 12, 14, 12, 11, 11, 11, 16, 13, 14, 13, 16, 14, 16])
    total = defaultdict(float)
    for cat in sorted(cats):
        g = init.get(CAT2INIT_ROW.get(cat, ''), {}).get('G')
        hh = init.get(CAT2INIT_ROW.get(cat, ''), {}).get('H')
        i_file = init.get(CAT2INIT_ROW.get(cat, ''), {}).get('I')
        h_init = hh if hh is not None else 0.0
        g_init = g if g is not None else 0.0
        tds = [d for d in tax_detail if d['cat'] == cat]
        t_in = round(sum(d['qty'] for d in tds if d['qty'] > 0), 3)
        t_out = round(sum(d['qty'] for d in tds if d['qty'] < 0), 3)
        t_net = round(t_in + t_out, 3)
        tax_stock = round(h_init + t_net, 3)
        tmp_snap = round(tmp_agg.get(cat, 0.0), 3)
        gl = round(gl_agg.get(cat, 0.0), 3)
        tmp_ex = round(tmp_snap - gl, 3)
        i_calc = round(g_init - h_init, 3)
        stock_diff = round(tmp_ex - tax_stock, 3)
        resid = round(stock_diff - i_calc, 3)
        ok = '一致' if abs(resid) <= tol else '不一致'
        ws.append([cat, g if g is not None else '空(按0)', h_init, i_calc,
                    i_file if i_file is not None else '空', t_in, t_out, t_net, tax_stock,
                    tmp_snap, gl, tmp_ex, stock_diff, resid, ok])
        row = ws[ws.max_row]
        for c in row:
            c.fill = OK_FILL if ok == '一致' else DIFF_FILL
        for k, v in [('税局入库', t_in), ('税局出库', t_out), ('税局净变动', t_net),
                     ('税局剩余库存', tax_stock), ('TMP快照', tmp_snap), ('GL', gl),
                     ('TMP剔GL', tmp_ex), ('库存差异', stock_diff), ('残差', resid)]:
            total[k] += v
    # 合计行
    ws.append(['成品油合计', '', '', '', '',
               round(total['税局入库'], 3), round(total['税局出库'], 3), round(total['税局净变动'], 3),
               '', round(sum(v for k, v in tmp_agg.items()), 3), round(sum(gl_agg.values()), 3),
               round(sum(tmp_agg.values()) - sum(gl_agg.values()), 3),
               round(total['库存差异'], 3), round(total['残差'], 3), ''])
    ws.cell(ws.max_row, 1).font = Font(bold=True)
    note = ('口径: 税局库存=H列初始(取不到按0)+税局变动(开具日期在窗口内); '
            'TMP库存=快照按税收分类名称合计-GainLoss(创建时间在窗口内); '
            '初始差异=G列TMP初始-H列税局初始; 库存差异-初始差异=隐含变动差异(TMP非GL变动-税局变动), '
            '不为0即两边期间变动存在差异')
    ws.append([note])
    ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=15)

    # --- 税局变动明细 ---
    ws2 = wb.create_sheet('税局变动明细')
    _hdr(ws2, ['大类', '开具日期', '发票号码', '变动方式', '数量(方向后)', '方向'],
         [14, 12, 24, 14, 13, 8])
    for d in sorted(tax_detail, key=lambda x: (x['cat'], x['date'], x['no'])):
        ws2.append([d['cat'], d['date'], d['no'], d['way'], d['qty'], d['dir']])
        if d['dir'] == '未识别':
            for c in ws2[ws2.max_row]:
                c.fill = DIFF_FILL

    # --- GainLoss明细 ---
    ws3 = wb.create_sheet('GainLoss明细(窗口内)')
    _hdr(ws3, ['序列号', '大类', '油品细类', '数量', '对手方合同号', '创建时间'], [14, 10, 12, 11, 18, 20])
    for d in sorted(gl_detail, key=lambda x: x['create']):
        ws3.append([d['sn'], d['cat'], d['sub'], d['qty'], d['party'], d['create']])

    # --- TMP库存快照 ---
    ws4 = wb.create_sheet('TMP库存快照')
    _hdr(ws4, ['油品分类(细类)', '税收分类名称(大类)', '剩余库存(吨)', '更新时间'], [28, 20, 13, 20])
    for r in sorted(tmp_rows, key=lambda x: (x['cat'], x['sub'])):
        ws4.append([r['sub'], r['cat'], r['qty'], r['upd']])

    wb.save(dst)


# ---------------- 主流程 ----------------
def resolve(path):
    """入参文件路径: 绝对路径直接用, 相对路径基于 energy/ 目录解析"""
    return path if os.path.isabs(path) else os.path.join(BASE, path)


def run(args):
    start = date.fromisoformat(args.start)
    end = date.fromisoformat(args.end)
    init = load_init(resolve(args.init))
    tax_detail, tax_cats = load_tax_movements(resolve(args.tax), start, end)
    tmp_rows, tmp_agg = load_tmp_stock(resolve(args.tmp_stock))
    gl_detail, gl_agg = load_gl(resolve(args.gl), start, end)
    cats = sorted(set(tax_cats) | set(tmp_agg) | set(gl_agg))

    dst = os.path.join(OUT_DIR, '库存差异核对_%s.xlsx' % datetime.now().strftime('%Y%m%d%H%M%S'))
    os.makedirs(OUT_DIR, exist_ok=True)
    write_report(dst, cats, init, tax_detail, tmp_rows, tmp_agg, gl_detail, gl_agg, start, end, args.tol)

    p('取数窗口: %s ~ %s (含当天)' % (start, end))
    p('输入文件: 税局=%s | TMP库存=%s | GL=%s | 初始数据=%s' % (
        resolve(args.tax), resolve(args.tmp_stock), resolve(args.gl), resolve(args.init)))
    p('大类: %s' % ', '.join(cats))
    for cat in cats:
        row = init.get(CAT2INIT_ROW.get(cat, ''), {})
        g, hh, i = row.get('G'), row.get('H'), row.get('I')
        h_init = hh if hh is not None else 0.0
        tds = [d for d in tax_detail if d['cat'] == cat]
        t_net = round(sum(d['qty'] for d in tds), 3)
        tmp_ex = round(tmp_agg.get(cat, 0.0) - gl_agg.get(cat, 0.0), 3)
        stock_diff = round(tmp_ex - (h_init + t_net), 3)
        i_calc = round((g if g is not None else 0.0) - h_init, 3)
        p('  %s: 税局库存=%.3f | TMP剔GL=%.3f | 库存差异=%s | 初始差异(I=G-H)=%s | 残差=%s%s' % (
            cat, h_init + t_net, tmp_ex, stock_diff, i_calc, round(stock_diff - i_calc, 3),
            '  [一致]' if abs(stock_diff - i_calc) <= args.tol else '  [不一致] <<<<'))
    p('输出文件完整路径: %s' % dst)


# ---------------- 启动时交互输入 ----------------
def ask(prompt, default):
    """询问一项, 回车取默认值"""
    try:
        v = input('%s (回车=%s): ' % (prompt, default)).strip()
    except EOFError:
        v = ''
    return v or default


def ask_file(prompt, default):
    """询问文件名, 回车取默认值; 不存在则重问"""
    while True:
        v = ask(prompt, default)
        fp = resolve(v)
        if os.path.exists(fp):
            return v
        p('!! 文件不存在: %s , 请重新输入' % fp)


def ask_date(prompt, default):
    """询问日期 (yyyy-mm-dd), 回车取默认值; 格式错误则重问"""
    while True:
        v = ask(prompt + ' (yyyy-mm-dd)', default)
        try:
            date.fromisoformat(v)
            return v
        except ValueError:
            p('!! 日期格式不正确: %s , 请输入 yyyy-mm-dd' % v)


def prompt_inputs(args):
    """启动时逐项询问输入文件与取数窗口 (覆盖到 args 上)"""
    p('=' * 64)
    p('税局 vs TMP 剩余库存核对 - 请输入参数 (直接回车使用括号内默认值)')
    p('=' * 64)
    args.tax = ask_file('税局发票文件', args.tax)
    args.tmp_stock = ask_file('TMP库存快照文件', args.tmp_stock)
    args.gl = ask_file('GainLoss明细文件', args.gl)
    args.init = ask_file('初始数据文件', args.init)
    args.start = ask_date('取数起始日', args.start)
    args.end = ask_date('取数截止日', args.end)


def main():
    ap = argparse.ArgumentParser(description='税局 vs TMP 剩余库存核对 (独立工具)')
    ap.add_argument('--tax', default=TAX_FILE, help='税局发票文件 (相对 energy/ 或绝对路径)')
    ap.add_argument('--tmp-stock', default=TMP_STOCK_FILE, help='TMP库存快照文件')
    ap.add_argument('--gl', default=GL_FILE, help='TMP GainLoss明细文件')
    ap.add_argument('--init', default=INIT_FILE, help='初始数据文件')
    ap.add_argument('--start', default=DEFAULT_START, help='取数起始日 (含当天)')
    ap.add_argument('--end', default=DEFAULT_END, help='取数截止日 (含当天)')
    ap.add_argument('--tol', type=float, default=0.01, help='判定"库存差异=初始差异"的容差(吨)')
    ap.add_argument('--batch', action='store_true', help='不询问, 直接用默认值/命令行参数')
    args = ap.parse_args()
    if not args.batch:
        prompt_inputs(args)
    run(args)


if __name__ == '__main__':
    main()
