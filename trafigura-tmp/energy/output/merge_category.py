# -*- coding: utf-8 -*-
"""
把原始进/销项文件中"按规则判定为指定大类、但现有判定文件里未识别"的行, 合并进判定文件:
- 已存在的明细行(是否成品油=否): 就地更新为 是 + 大类/细类/命中关键字, 并补成品油底色
- 不存在的明细行: 追加到明细表尾
- 发票汇总: 无此发票则按 oil_classify 同构逻辑追加; 已有则合并大类/细类、累加数量
- 不改动、不重排任何现有行(保留手动修改); 匹配键: 发票号码+商品名称+数量(空号码行用 日期+商品名称+数量)
用法: from merge_category import run; run('石脑油', kind='进项', src=..., result_path=...)
"""
from collections import defaultdict
import openpyxl
from oil_rules import classify_detail, s, OIL_FILL


def _qkey(v):
    try:
        return round(float(v), 6)
    except (TypeError, ValueError):
        return v


def run(category, kind, src, result_path):
    assert kind in ('进项', '销项')
    wb = openpyxl.load_workbook(result_path)   # 非 data_only, 保留现有样式与值
    ws_sum, ws_det = wb['发票汇总判定'], wb['发票明细判定']
    det_rows = list(ws_det.iter_rows(values_only=True))
    sum_nos = {s(r[0]).strip() for r in list(ws_sum.iter_rows(values_only=True))[1:]}

    src_wb = openpyxl.load_workbook(src, data_only=True)
    ws_info, ws_raw = src_wb.worksheets[0], src_wb.worksheets[1]
    rows_info = list(ws_info.iter_rows(values_only=True))
    rows_raw = list(ws_raw.iter_rows(values_only=True))
    hi = {h: i for i, h in enumerate(rows_info[0])}
    hr = {h: i for i, h in enumerate(rows_raw[0])}

    # 现有明细行索引
    idx_full, idx_pre = defaultdict(list), defaultdict(list)
    for i, r in enumerate(det_rows[1:], start=2):
        no, name, qty, date_s = s(r[0]).strip(), s(r[4]).strip(), _qkey(r[7]), s(r[1])[:10]
        (idx_full if no else idx_pre)[(no, name, qty) if no else (date_s, name, qty)].append(i)
    info_by_no = {}
    for r in rows_info[1:]:
        no = s(r[hi['发票号码']]).strip()
        if no and no not in info_by_no:
            info_by_no[no] = r
    pre_info = [r for r in rows_info[1:] if not s(r[hi['发票号码']]).strip()] if kind == '销项' else []
    pre_used = set()

    def match_pre(r):
        buyer, amt = s(r[hr['客户名称']]).strip(), r[hr['开票金额']]
        cands = [p for i, p in enumerate(pre_info) if i not in pre_used
                 and s(p[hi['购方名称']]).strip() == buyer and p[hi['不含税金额']] == amt]
        if len(cands) == 1:
            pre_used.add(pre_info.index(cands[0]))
            return cands[0]
        return None

    n_updated = n_appended = 0
    inv = defaultdict(lambda: {'rows': [], 'units': [], 'qty': 0.0, 'pre': False, 'norm': False})
    for r in rows_raw[1:]:
        if kind == '进项':
            no = s(r[hr['发票号码']]).strip()
            spec, name = s(r[hr['规格型号']]).strip(), s(r[hr['商品名称']]).strip()
            unit, qty, date_s = s(r[hr['单位']]).strip(), r[hr['数量']], s(r[hr['开票日期']])[:10]
            seller, inv_type, pre = s(r[hr['销方名称（发票）']]), s(r[hr['发票类型']]), False
            text = spec + name
        else:
            no = s(r[hr['发票号码']]).strip()
            name = s(r[hr['商品名称']]).strip()
            qty, unit, spec = r[hr['商品数量']], '', ''
            inv_type, pre = s(r[hr['发票类型']]), not no
            if pre:
                date_s = s(r[hr['单据日期']])[:10]
                pinfo = match_pre(r)
                if pinfo:
                    no = s(pinfo[hi['单据号']]).strip()
            else:
                date_s = s(r[hr['发票日期']]) or s(r[hr['单据日期']])
                date_s = date_s[:10]
            seller = s(r[hr['机构名称']])
            text = name
        row, hit = classify_detail(no, date_s, inv_type, seller, spec, name, unit, qty, text)
        if not hit or hit[0] != category:
            continue
        if pre:
            row[13] = (row[13] + '；' if row[13] else '') + (
                '已预制(无发票号码, 以单据号标识)' if no else '已预制(未能匹配单据号)')
        # 合并进明细: 已存在(原为否)则更新, 否则追加
        key = (no, name, _qkey(qty)) if no else (date_s, name, _qkey(qty))
        pool = idx_full if no else idx_pre
        if pool.get(key):
            ridx = pool[key].pop(0)
            for col, val in ((10, row[9]), (11, row[10]), (12, row[11]), (13, row[12])):
                ws_det.cell(ridx, col, val)
            for c in ws_det[ridx]:
                c.fill = OIL_FILL
            n_updated += 1
        else:
            ws_det.append(row)
            for c in ws_det[ws_det.max_row]:
                c.fill = OIL_FILL
            n_appended += 1
        rec = inv[no]
        rec['rows'].append((row, r))
        rec['pre'] = rec['pre'] or pre
        rec['norm'] = rec['norm'] or bool(hit[2])
        if unit and unit not in rec['units']:
            rec['units'].append(unit)
        try:
            rec['qty'] += float(qty or 0)
        except (TypeError, ValueError):
            pass

    # 发票汇总: 不存在的追加, 存在的合并
    n_sum_new = n_sum_merged = 0
    for no, rec in inv.items():
        bigs, smalls = set(), set()
        for row, r in rec['rows']:
            bigs.add(row[10]); smalls.add(row[11])
        if no and no not in sum_nos:
            info = info_by_no.get(no)
            first_row, first_raw = rec['rows'][0]
            if kind == '进项':
                ws_sum.append([
                    no,
                    s(info[hi['开票日期']])[:10] if info else s(first_raw[hr['开票日期']])[:10],
                    s(info[hi['发票类型']]) if info else s(first_raw[hr['发票类型']]),
                    s(info[hi['发票状态']]) if info else '',
                    s(info[hi['认证期间']]) if info else '',
                    s(info[hi['分摊期间']]) if info else '',
                    s(info[hi['销方名称']]) if info else s(first_raw[hr['销方名称（发票）']]),
                    s(info[hi['购方名称']]) if info else '',
                    ';'.join(rec['units']),
                    round(rec['qty'], 6) if rec['qty'] else None,
                    info[hi['不含税金额']] if info else None,
                    info[hi['税额']] if info else None,
                    info[hi['含税金额']] if info else None,
                    '是', '、'.join(sorted(bigs)), '、'.join(sorted(smalls)),
                    '含归一化匹配行' if rec['norm'] else '',
                ])
            else:
                def g(key_info, key_raw):
                    v = info[hi[key_info]] if info else first_raw[hr[key_raw]]
                    v = s(v)
                    if key_info == '发票日期' and not v:
                        v = s(first_raw[hr['单据日期']])
                    return v[:10] if key_info == '发票日期' else v
                ws_sum.append([
                    no, g('发票日期', '发票日期'), g('发票类型', '发票类型'), g('发票状态', '发票状态'),
                    '', '',
                    g('销方名称', '机构名称'), g('购方名称', '客户名称'),
                    ';'.join(rec['units']),
                    round(rec['qty'], 6) if rec['qty'] else None,
                    info[hi['不含税金额']] if info else first_raw[hr['开票金额']],
                    info[hi['税额']] if info else first_raw[hr['开票税额']],
                    info[hi['含税金额']] if info else first_raw[hr['含税金额']],
                    '是', '、'.join(sorted(bigs)), '、'.join(sorted(smalls)),
                    ('已预制(发票号码为空, 以单据号标识)；' if rec['pre'] else '')
                    + ('含归一化匹配行' if rec['norm'] else ''),
                ])
            new_row = ws_sum[ws_sum.max_row]
            if new_row[13].value == '是':
                for c in new_row:
                    c.fill = OIL_FILL
            n_sum_new += 1
        elif no:
            # 已有该发票(其他油品行): 合并大类/细类/数量
            target = None
            for r_ in ws_sum.iter_rows(min_row=2):
                if s(r_[0].value).strip() == no:
                    target = r_
                    break
            if target is None:
                continue
            for b in sorted(bigs):
                if b not in s(target[14].value):
                    target[14].value = (s(target[14].value) + '、' if s(target[14].value) else '') + b
            for sm in sorted(smalls):
                if sm not in s(target[15].value):
                    target[15].value = (s(target[15].value) + '、' if s(target[15].value) else '') + sm
            try:
                target[9].value = round(float(target[9].value or 0) + rec['qty'], 6)
            except (TypeError, ValueError):
                pass
            n_sum_merged += 1

    for ws in (ws_sum, ws_det):
        ws.auto_filter.ref = ws.dimensions
    wb.save(result_path)
    return {'类别': category, '方向': kind, '明细更新行': n_updated, '明细追加行': n_appended,
            '汇总新增发票': n_sum_new, '汇总合并发票': n_sum_merged}
