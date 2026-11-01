# -*- coding: utf-8 -*-
"""
剩余可用库存计算 (按细类), 输出到 energy/output/库存计算/
- 范围: 指定大类 + 指定细类列表; 进项/销项数量均取 判定文件明细行 (红冲明细为负数自动抵减, 销项含已预制票)
- 初始数据: Energy初始数据.xlsx G列, 按 CONFIG 关键词匹配行 (如 '92#小计'); 取不到或为空 -> 0, 报告注明
- 未识别细类的发票 (大类命中但细类=NA/空): 不剔除, 在报告中逐票注明; 其他未纳入计算的细类也一并列示
可被流程调用: from stock_report import run; run(level='细类', ...)
"""
import io
import os
import sys
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # .../energy
OUT = os.path.join(BASE, 'output')
STOCK_DIR = os.path.join(OUT, '库存计算')

TOL = 1e-6
NA_SUBS = ('NA', '', 'None', 'none')   # 视为"未识别细类"

HDR_FONT = Font(bold=True, color='FFFFFF')
HDR_FILL = PatternFill('solid', fgColor='4472C4')
NEG_FILL = PatternFill('solid', fgColor='FCE4EC')
INIT_FILL = PatternFill('solid', fgColor='E2EFDA')
FLAG_FILL = PatternFill('solid', fgColor='FFF3E0')


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


def load_details(path):
    """判定文件 -> [(日期, 发票号码, 大类, 细类, 数量, 商品名称, 是否已预制, 销方名称, 购方名称), ...]
    销方名称取明细行; 购方名称明细表没有, 按发票号码联 发票汇总判定 (销方缺失时同样回退汇总)"""
    wb = openpyxl.load_workbook(path, data_only=True)
    rows = list(wb['发票明细判定'].iter_rows(values_only=True))
    h = {c: i for i, c in enumerate(rows[0])}
    sum_rows = list(wb['发票汇总判定'].iter_rows(values_only=True))
    hs = {c: i for i, c in enumerate(sum_rows[0])}
    party = {}   # 发票号码 -> (销方, 购方)
    for r in sum_rows[1:]:
        no = str(r[hs['发票号码']])
        if no and no not in party:
            party[no] = (str(r[hs['销方名称']] or ''), str(r[hs['购方名称']] or ''))
    out = []
    for r in rows[1:]:
        if r[h['是否成品油']] != '是':
            continue
        qty = f(r[h['数量']])
        if qty is None:
            continue
        premade = '已预制' in str(r[h['备注']] or '')
        no = str(r[h['发票号码']])
        s_sell, s_buy = party.get(no, ('', ''))
        seller = str(r[h['销方名称']] or '') or s_sell
        out.append((to_date(r[h['开票日期']]), no,
                    str(r[h['油品分类（大类）']]), str(r[h['油品细分（细类）']]),
                    qty, str(r[h['商品名称']] or ''), premade, seller, s_buy))
    return out


def load_initial(init_path, keywords):
    """按关键词匹配初始行(品名+规格包含关键词)求和, 返回(数值或None, 来源描述)
    关键词带 '!' 前缀表示减项 (用于 'XX小计' 内含子项时做差, 如 0#小计-0#京标)"""
    wb = openpyxl.load_workbook(init_path, data_only=True)
    rows = list(wb.active.iter_rows(values_only=True))
    if not keywords:
        return None, '未配置匹配关键词'
    texts = [(str(r[0] or '') + '/' + str(r[1] or '')).strip('/') for r in rows[1:]]
    total, notes, found = 0.0, [], False
    for kw in keywords:
        neg, k = (True, kw[1:]) if kw.startswith('!') else (False, kw)
        hits = [i for i, t in enumerate(texts) if k in t]
        if not hits:
            notes.append(f"'{k}'未找到行")
            continue
        found = True
        for i in hits:
            v = f(rows[1:][i][6]) or 0.0   # G列
            total += -v if neg else v
            notes.append(f"{'-' if neg else '+'}'{texts[i]}'={round(v, 4)}")
    if not found:
        return None, '未找到匹配行, 记0'
    return total, 'G列(' + ' '.join(notes) + ')'


def run(level, categories, subcategories, result_in_path, result_out_path,
        init_path, dst, start_date, init_config):
    """剩余可用库存计算, 返回统计 dict
    level: '细类'=每个细类一个桶 (subcategories 为桶); '大类'=每个大类一个桶 (含 NA/未识别细类行)
    init_config: {桶名(细类或大类): [关键词,...]}, 匹配初始数据 品名/规格
    """
    det_in = load_details(result_in_path)
    det_out = load_details(result_out_path)

    by_big = (level == '大类')
    subs = list(categories) if by_big else list(subcategories)
    sub_set = set(subs)
    cat_set = set(categories)

    def bucket(big, s_):
        """明细行归属哪个桶: 大类口径按大类(细类不限, 含NA); 细类口径按细类"""
        return big if by_big else s_

    # 范围外/未识别 明细 (细类口径: 大类命中但细类未纳入); 已预制票单独剔除, 不重复标记
    flagged = []   # (方向, 日期, 发票号码, 大类, 细类, 数量, 商品名称, 类型)
    if not by_big:
        for src, direction in ((det_in, '进项'), (det_out, '销项')):
            for d, no, big, sub, qty, name, premade, seller, buyer in src:
                if big not in cat_set or premade:
                    continue
                if sub in NA_SUBS:
                    flagged.append((direction, d, no, big, sub, qty, name, '未识别细类(NA)'))
                elif sub not in sub_set:
                    flagged.append((direction, d, no, big, sub, qty, name, '细类未纳入本次计算'))

    # 销项剔除已预制票 (未开票, 税局无号, 与大类台账税局口径一致)
    excluded_sales = sorted(
        [(d, no, big, sub, qty, name, seller, buyer)
         for d, no, big, sub, qty, name, premade, seller, buyer in det_out
         if premade and big in cat_set and d >= start_date],
        key=lambda x: (x[0], x[1]))

    # 每个桶独立台账 (只统计 初始日期及之后的进销变动, 之前的发生额已含在初始库存中)
    results = []
    for sub in subs:
        init_v, init_note = load_initial(init_path, init_config.get(sub, []))
        events = []   # (日期, 方向, 发票号码, 销方, 购方, 大类, 细类, 商品名称, 数量)
        for d, no, big, s_, qty, name, premade, seller, buyer in det_in:
            if big in cat_set and bucket(big, s_) == sub and d >= start_date:
                events.append((d, '进项', no, seller, buyer, big, s_, name, qty))
        for d, no, big, s_, qty, name, premade, seller, buyer in det_out:
            if big in cat_set and bucket(big, s_) == sub and d >= start_date and not premade:
                events.append((d, '销项', no, seller, buyer, big, s_, name, qty))
        events.sort(key=lambda e: (e[0], e[1], e[2]))
        ledger = [(start_date, '初始', '(初始数据)', '', '', '', '', '',
                   round(init_v or 0.0, 6), round(init_v or 0.0, 6), '')]
        balance = init_v or 0.0
        first_neg = None
        for d, direction, no, seller, buyer, big, s_, name, qty in events:
            delta = qty if direction == '进项' else -qty
            balance += delta
            ledger.append((d, direction, no, seller, buyer, big, s_, name,
                           round(delta, 6), round(balance, 6), ''))
            if first_neg is None and balance < -TOL:
                first_neg = (d, direction, no, round(delta, 4), round(balance, 6))
        total_in = round(sum(e[8] for e in events if e[1] == '进项'), 4)
        total_out = round(sum(e[8] for e in events if e[1] == '销项'), 4)
        results.append({'细类': sub, '初始': init_v, '初始说明': init_note,
                        '进项合计': total_in, '销项合计': total_out,
                        '明细行数': len(events), '最终': round(balance, 4),
                        '首次转负': first_neg, '流水': ledger,
                        '剔除': [(d, no, big, s_, qty, name, seller, buyer)
                                 for d, no, b, s_, qty, name, seller, buyer in excluded_sales
                                 if bucket(b, s_) == sub]})

    # 按发票归并未识别明细 (需求1: 提示到发票)
    inv_flag = {}
    for direction, d, no, big, sub, qty, name, typ in flagged:
        rec = inv_flag.setdefault((no, str(d), typ), {'cats': set(), 'qty': 0.0, 'lines': 0, 'name': name})
        rec['cats'].add(f'{big}/{sub or "空"}')
        rec['qty'] += qty
        rec['lines'] += 1

    _write(dst, level, categories, subs, results, inv_flag, flagged, excluded_sales, start_date)

    na_in_win = round(sum(q for direction, d, no, big, sub, q, name, typ in flagged
                          if typ == '未识别细类(NA)' and d >= start_date), 4)
    return {'输出': dst, '范围': {'大类': categories, '细类': subs},
            '细类结果': [{k: v for k, v in r.items() if k != '流水'} for r in results],
            '未识别/未纳入明细行数': len(flagged),
            '窗口内未识别细类(NA)数量合计': na_in_win,
            '销项剔除已预制': excluded_sales,
            '涉及发票数': len(inv_flag)}


def _write(dst, level, categories, subs, results, inv_flag, flagged, excluded_sales, start_date):
    out = openpyxl.Workbook()
    ws1 = out.active
    ws1.title = '汇总'
    p_row = lambda row: ws1.append(list(row))

    unit = '大类' if level == '大类' else '细类'
    p_row([f'剩余可用库存计算 ({"/".join(categories)} 按{unit})'])
    ws1[ws1.max_row][0].font = Font(bold=True, size=12)
    p_row([f'口径: 进项取判定文件明细数量; 销项取已开票明细数量(红冲为负, 已预制票剔除, 与大类台账税局口径一致); '
           f'取数窗口自 {start_date} 起(之前的发生额已含在初始库存中); '
           f'初始数据取 {start_date} (Energy初始数据.xlsx G列, 取不到记0)'])
    p_row([])
    for r in results:
        p_row([f"{unit}: {r['细类']}"])
        ws1[ws1.max_row][0].font = Font(bold=True)
        p_row([f"初始数量({start_date})", round(r['初始'] or 0.0, 4), '吨', r['初始说明']])
        p_row([f"进项合计(结果文件明细, {r['明细行数']}行)", r['进项合计'], '吨'])
        p_row(['销项合计(结果文件明细)', r['销项合计'], '吨'])
        p_row(['剩余可用库存 = 初始 + 进项 - 销项', r['最终'], '吨'])
        p_row(['是否出现负数', '是' if r['首次转负'] else '否'])
        if r['首次转负']:
            d, direction, no, delta, bal = r['首次转负']
            p_row(['首次转负日期', str(d)])
            p_row(['首次转负发票', no])
            p_row(['该票方向/数量变动', f'{direction} / {delta}'])
            p_row(['该票结余(转负时)', bal])
        p_row([])

    if inv_flag:
        p_row(['!! 以下发票的大类命中但细类未识别/未纳入本次计算, 其数量未计入上方任何细类, 请核实:'])
        ws1[ws1.max_row][0].font = Font(bold=True, color='C00000')
        na_win_qty = 0.0
        for (no, d, typ), rec in sorted(inv_flag.items(), key=lambda x: (x[0][1], x[0][0])):
            p_row([f'  {no} {d} {"、".join(sorted(rec["cats"]))} 明细{rec["lines"]}行 '
                   f'数量合计={round(rec["qty"], 4)} [{typ}]'])
            if d >= str(start_date):
                na_win_qty += rec['qty']
        if na_win_qty:
            p_row([f'窗口内未识别/未纳入数量合计(未计入任何细类, 对账时需另行考虑):',
                   round(na_win_qty, 4), '吨'])
    if excluded_sales:
        p_row(['销项中已剔除的已预制发票(未开票, 税局无号, 与大类台账口径一致):'])
        for d, no, big, sub, qty, name, seller, buyer in excluded_sales:
            p_row([f'  {no} {d} {big}/{sub} 数量={qty}'])
    for c in ws1['A']:
        if c.value and not c.font.bold:   # 不重赋 StyleProxy, 只在未加粗时赋新 Font
            c.font = Font(bold=True)
    ws1.column_dimensions['A'].width = 56
    ws1.column_dimensions['B'].width = 22
    ws1.column_dimensions['C'].width = 8
    ws1.column_dimensions['D'].width = 40

    # 未识别/未纳入 明细行
    if inv_flag:
        ws3 = out.create_sheet('未识别细类明细')
        ws3.append(['方向', '开票日期', '发票号码', '大类', '细类', '数量', '商品名称', '类型'])
        for c in ws3[1]:
            c.font, c.fill = HDR_FONT, HDR_FILL
            c.alignment = Alignment(horizontal='center')
        for direction, d, no, big, sub, qty, name, typ in flagged:
            ws3.append([direction, str(d), no, big, sub, qty, name, typ])
        for i, w in enumerate([8, 12, 24, 18, 26, 12, 40, 22], 1):
            ws3.column_dimensions[get_column_letter(i)].width = w
        ws3.freeze_panes = 'A2'

    for r in results:
        ws = out.create_sheet(f"流水_{r['细类']}"[:31])
        ws.append(['日期', '方向', '发票号码', '销方名称', '购方名称', '油品大类', '油品细类',
                   '商品名称', '数量变动(+进/-销)', '累计结余', '备注'])
        for c in ws[1]:
            c.font, c.fill = HDR_FONT, HDR_FILL
            c.alignment = Alignment(horizontal='center')
        for d, direction, no, seller, buyer, big, s_, name, change, bal, remark in r['流水']:
            ws.append([str(d), direction, no, seller, buyer, big, s_, name, change, bal, remark])
            row = ws[ws.max_row]
            if direction == '初始':
                for c in row:
                    c.fill = INIT_FILL
            elif bal < -TOL:
                for c in row:
                    c.fill = NEG_FILL
        if r.get('剔除'):
            ws.append([])
            ws.append(['以下为已剔除数据(未开票, 税局无号), 未计入上方累计结余:'])
            for c in ws[ws.max_row]:
                if c.value:
                    c.font = Font(bold=True)
                    c.fill = FLAG_FILL
            for d, no, big, s_, qty, name, seller, buyer in r['剔除']:
                ws.append([str(d), '销项(已预制剔除)', no, seller, buyer, big, s_, name,
                           qty, '', '已剔除, 未计入结余'])
                for c in ws[ws.max_row]:
                    c.fill = FLAG_FILL
        for i, w in enumerate([12, 14, 24, 26, 26, 14, 22, 36, 17, 14, 22], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions

    out.save(dst)
