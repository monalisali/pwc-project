# -*- coding: utf-8 -*-
"""
销项发票成品油判定 (流程步骤1-销项)
逻辑/字段与进项一致, 规则共用 oil_rules.py; 匹配文本 = 规格型号(空) + 商品名称
销项源文件差异: 明细无"规格型号""单位"列(数量列为"商品数量"), 发票信息无"认证期间/分摊期间"(留空)
已预制发票不剔除: 明细空号码行按 购方名称+不含税金额 匹配信息行的单据号作为标识
可独立运行:  python -I oil_classify_sales.py
可被流程调用: from oil_classify_sales import run; run(src, dst)
"""
import os
import io
import sys
from collections import Counter
import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from oil_rules import write_sheet, classify_detail, s

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # .../energy
DEFAULT_SRC = os.path.join(BASE, 'Energy销项.xlsx')
DEFAULT_DST = os.path.join(BASE, 'output', '成品油发票判定_Energy销项.xlsx')

SUMMARY_HEADER = ['发票号码', '开票日期', '发票类型', '发票状态', '认证期间', '分摊期间',
                  '销方名称', '购方名称', '单位', '数量', '不含税金额', '税额', '含税金额',
                  '是否成品油发票', '油品分类（大类）', '油品细分（细类）', '备注']
SUMMARY_WIDTHS = [24, 12, 16, 11, 10, 10, 30, 26, 8, 12, 12, 10, 12, 13, 16, 34, 16]
DETAIL_HEADER = ['发票号码', '开票日期', '发票类型', '销方名称', '商品名称', '规格型号',
                 '单位', '数量', '规格型号+商品名称', '是否成品油', '油品分类（大类）',
                 '油品细分（细类）', '命中关键字', '备注']
DETAIL_WIDTHS = [24, 12, 16, 30, 36, 20, 8, 10, 52, 10, 16, 26, 22, 18]


def run(src, dst):
    """判定销项发票, 输出结果文件, 返回统计 dict"""
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

    # 已预制信息行(发票号码为空): 供明细空号码行按 购方名称+不含税金额 匹配单据号
    pre_info = [r for r in rows_info[1:] if not s(r[hi['发票号码']]).strip()]
    pre_used = set()

    def match_pre(r):
        buyer = s(r[hd['客户名称']]).strip()
        amt = r[hd['开票金额']]
        cands = [p for i, p in enumerate(pre_info)
                 if i not in pre_used
                 and s(p[hi['购方名称']]).strip() == buyer and p[hi['不含税金额']] == amt]
        if len(cands) == 1:
            pre_used.add(pre_info.index(cands[0]))
            return cands[0]
        return None

    # 逐明细行判定 (销项明细: 机构名称=销方, 商品数量=数量, 发票日期=开票日期)
    detail_out = []
    inv_flag = {}            # 发票号码/单据号 -> {'cats': set, 'units': [], 'qty': float, 'first': row, 'pre': bool}
    n_pre_matched = 0
    n_pre_unmatched = 0
    for r in rows_det[1:]:
        no = s(r[hd['发票号码']]).strip()
        name = s(r[hd['商品名称']]).strip()
        qty = r[hd['商品数量']]
        pre = False
        if not no:                        # 已预制: 单据日期作开票日期, 尝试匹配单据号
            date_s = s(r[hd['单据日期']])[:10]
            pinfo = match_pre(r)
            if pinfo:
                no = s(pinfo[hi['单据号']]).strip()
                n_pre_matched += 1
            else:
                n_pre_unmatched += 1
            pre = True
        else:
            date_s = s(r[hd['发票日期']])[:10]
        row, hit = classify_detail(
            no, date_s, s(r[hd['发票类型']]),
            s(r[hd['机构名称']]), '', name, '', qty, name)
        if pre:
            pre_note = '已预制(无发票号码, 以单据号标识)' if no else '已预制(未能匹配单据号)'
            row[13] = (row[13] + '；' if row[13] else '') + pre_note
        detail_out.append(row)
        if hit:
            rec = inv_flag.setdefault(no, {'cats': set(), 'units': [], 'qty': 0.0, 'first': r, 'pre': pre})
            rec['cats'].add(hit)
            rec['pre'] = rec['pre'] or pre
            try:
                rec['qty'] += float(qty or 0)
            except (TypeError, ValueError):
                pass

    # 汇总到发票级 (属性优先取发票信息sheet, 无对应信息时回退明细行; 认证/分摊期间销项无, 留空)
    pre_by_danhao = {s(p[hi['单据号']]).strip(): p for p in pre_info}
    inv_out = []
    for no, rec in inv_flag.items():
        cats = rec['cats']
        info = info_by_no.get(no) or pre_by_danhao.get(no)
        first = rec['first']

        def g(key_info, key_det, cut10=False):
            v = info[hi[key_info]] if info else first[hd[key_det]]
            v = s(v)
            if key_info == '发票日期' and not v:      # 预制行无发票日期, 回退单据日期
                v = s(first[hd['单据日期']])
            return v[:10] if cut10 else v

        inv_out.append([
            no,
            g('发票日期', '发票日期', cut10=True),
            g('发票类型', '发票类型'),
            g('发票状态', '发票状态'),
            '',   # 认证期间(销项无)
            '',   # 分摊期间(销项无)
            g('销方名称', '机构名称'),
            g('购方名称', '客户名称'),
            ';'.join(rec['units']),
            round(rec['qty'], 6) if rec['qty'] else None,
            info[hi['不含税金额']] if info else first[hd['开票金额']],
            info[hi['税额']] if info else first[hd['开票税额']],
            info[hi['含税金额']] if info else first[hd['含税金额']],
            '是',
            '、'.join(sorted({c[0] for c in cats})),
            '、'.join(sorted({c[1] for c in cats})),
            ('已预制(发票号码为空, 以单据号标识)' if rec['pre'] else '')
            + ('；' if rec['pre'] and any(c[2] for c in cats) else '')
            + ('含归一化匹配行' if any(c[2] for c in cats) else ''),
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

    norm_rows = [d for d in detail_out if d[13] and '归一化' in d[13]]
    return {
        '源文件': src,
        '输出文件': dst,
        '明细行总数': len(detail_out),
        '已预制明细行': (n_pre_matched, n_pre_unmatched),
        '成品油明细行': sum(1 for d in detail_out if d[9] == '是'),
        '成品油发票数(含预制)': len(inv_out),
        '按细类统计': sorted(Counter((d[10], d[11]) for d in detail_out if d[9] == '是').items()),
        '归一化匹配行数': len(norm_rows),
    }


def format_stats(st):
    buf = io.StringIO()
    p = lambda *a: print(*a, file=buf)
    p('输出文件:', st['输出文件'])
    p('明细行总数:', st['明细行总数'])
    p('其中已预制明细行:', st['已预制明细行'][0] + st['已预制明细行'][1],
      f"(匹配到单据号 {st['已预制明细行'][0]}, 未匹配 {st['已预制明细行'][1]})")
    p('成品油明细行:', st['成品油明细行'])
    p('成品油发票数(含预制):', st['成品油发票数(含预制)'])
    p('按细类统计(明细行):')
    for (big, small), n in st['按细类统计']:
        p(f'  {big} / {small}: {n} 行')
    p('归一化匹配行数:', st['归一化匹配行数'])
    return buf.getvalue()


if __name__ == '__main__':
    st = run(DEFAULT_SRC, DEFAULT_DST)
    with open(os.path.join(os.path.dirname(DEFAULT_DST), '_stats_销项.txt'), 'w', encoding='utf-8') as f:
        f.write(format_stats(st))
    print('saved')
