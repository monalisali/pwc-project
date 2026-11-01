# -*- coding: utf-8 -*-
"""
进项发票成品油判定 (流程步骤1-进项)
逻辑: 用【规格型号+商品名称】与 oil_rules.RULES 条目做"包含"比较
可独立运行:  python -I oil_classify.py
可被流程调用: from oil_classify import run; run(src, dst)
"""
import os
import io
import sys
from collections import Counter
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from oil_rules import write_sheet, classify_detail, s

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # .../energy
DEFAULT_SRC = os.path.join(BASE, 'Energy进项20260501-20261026.xlsx')
DEFAULT_DST = os.path.join(BASE, 'output', '成品油发票判定_Energy进项20260501-20261026.xlsx')

SUMMARY_HEADER = ['发票号码', '开票日期', '发票类型', '发票状态', '认证期间', '分摊期间',
                  '销方名称', '购方名称', '单位', '数量', '不含税金额', '税额', '含税金额',
                  '是否成品油发票', '油品分类（大类）', '油品细分（细类）', '备注']
SUMMARY_WIDTHS = [24, 12, 16, 11, 10, 10, 30, 26, 8, 12, 12, 10, 12, 13, 16, 34, 16]
DETAIL_HEADER = ['发票号码', '开票日期', '发票类型', '销方名称', '商品名称', '规格型号',
                 '单位', '数量', '规格型号+商品名称', '是否成品油', '油品分类（大类）',
                 '油品细分（细类）', '命中关键字', '备注']
DETAIL_WIDTHS = [24, 12, 16, 30, 36, 20, 8, 10, 52, 10, 16, 26, 22, 18]


def run(src, dst):
    """判定进项发票, 输出结果文件, 返回统计 dict"""
    wb = openpyxl.load_workbook(src, data_only=True)
    ws_info, ws_det = wb.worksheets[0], wb.worksheets[1]

    rows_info = list(ws_info.iter_rows(values_only=True))
    rows_det = list(ws_det.iter_rows(values_only=True))
    hi = {h: i for i, h in enumerate(rows_info[0])}
    hd = {h: i for i, h in enumerate(rows_det[0])}

    # 发票信息按发票号码建索引
    info_by_no = {}
    for r in rows_info[1:]:
        no = s(r[hi['发票号码']]).strip()
        if no:
            info_by_no.setdefault(no, r)

    # 逐明细行判定
    detail_out = []
    inv_flag = {}            # 发票号码 -> {'cats': set, 'units': [..], 'qty': float}
    for r in rows_det[1:]:
        no = s(r[hd['发票号码']]).strip()
        spec = s(r[hd['规格型号']]).strip()
        name = s(r[hd['商品名称']]).strip()
        unit = s(r[hd['单位']]).strip()
        qty = r[hd['数量']]
        row, hit = classify_detail(
            no, s(r[hd['开票日期']]), s(r[hd['发票类型']]),
            s(r[hd['销方名称（发票）']]), spec, name, unit, qty, spec + name)
        detail_out.append(row)
        if hit:
            rec = inv_flag.setdefault(no, {'cats': set(), 'units': [], 'qty': 0.0})
            rec['cats'].add(hit)
            if unit and unit not in rec['units']:
                rec['units'].append(unit)
            try:
                rec['qty'] += float(qty or 0)
            except (TypeError, ValueError):
                pass

    # 汇总到发票级
    inv_out = []
    for no, rec in inv_flag.items():
        cats = rec['cats']
        info = info_by_no.get(no)
        inv_out.append([
            no,
            s(info[hi['开票日期']])[:10] if info else '',
            s(info[hi['发票类型']]) if info else '',
            s(info[hi['发票状态']]) if info else '',
            s(info[hi['认证期间']]) if info else '',
            s(info[hi['分摊期间']]) if info else '',
            s(info[hi['销方名称']]) if info else '',
            s(info[hi['购方名称']]) if info else '',
            ';'.join(rec['units']),
            round(rec['qty'], 6) if rec['qty'] else None,
            info[hi['不含税金额']] if info else None,
            info[hi['税额']] if info else None,
            info[hi['含税金额']] if info else None,
            '是',
            '、'.join(sorted({c[0] for c in cats})),
            '、'.join(sorted({c[1] for c in cats})),
            '含归一化匹配行' if any(c[2] for c in cats) else '',
        ])

    inv_out.sort(key=lambda x: (x[1], x[0]))
    detail_out.sort(key=lambda x: (x[1], x[0], x[4]))

    out = openpyxl.Workbook()
    ws1 = out.active
    ws1.title = '发票汇总判定'
    write_sheet(ws1, SUMMARY_HEADER, inv_out, SUMMARY_WIDTHS, flag_col=13)
    ws2 = out.create_sheet('发票明细判定')
    write_sheet(ws2, DETAIL_HEADER, detail_out, DETAIL_WIDTHS, flag_col=9)
    out.save(dst)

    norm_rows = [d for d in detail_out if d[13]]
    return {
        '源文件': src,
        '输出文件': dst,
        '明细行总数': len(detail_out),
        '成品油明细行': sum(1 for d in detail_out if d[9] == '是'),
        '成品油发票数': len(inv_out),
        '按细类统计': sorted(Counter((d[10], d[11]) for d in detail_out if d[9] == '是').items()),
        '归一化匹配行数': len(norm_rows),
    }


def format_stats(st):
    buf = io.StringIO()
    p = lambda *a: print(*a, file=buf)
    p('输出文件:', st['输出文件'])
    p('明细行总数:', st['明细行总数'])
    p('成品油明细行:', st['成品油明细行'])
    p('成品油发票数:', st['成品油发票数'])
    p('按细类统计(明细行):')
    for (big, small), n in st['按细类统计']:
        p(f'  {big} / {small}: {n} 行')
    p('归一化匹配行数:', st['归一化匹配行数'])
    return buf.getvalue()


if __name__ == '__main__':
    st = run(DEFAULT_SRC, DEFAULT_DST)
    with open(os.path.join(os.path.dirname(DEFAULT_DST), '_stats_进项.txt'), 'w', encoding='utf-8') as f:
        f.write(format_stats(st))
    print('saved')
