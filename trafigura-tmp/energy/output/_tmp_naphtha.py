# -*- coding: utf-8 -*-
import io, os
from collections import Counter
import openpyxl
BASE = r'c:\D\AIWork\pwc-project\trafigura-tmp\energy'
buf = io.StringIO()
p = lambda *a: print(*a, file=buf)

# 税局文件: 商品和服务名称 分布
wb = openpyxl.load_workbook(os.path.join(BASE, '税局Energy增值税发票7.1-10.23.xlsx'), data_only=True)
rows = list(wb['增值税发票'].iter_rows(values_only=True))
h = {c: i for i, c in enumerate(rows[0])}
p('税局 商品和服务名称 分布:')
for k, n in Counter(str(r[h['商品和服务名称']]).strip() for r in rows[1:]).most_common():
    p('  ', k, n)

# TMP购进: 油品类型 x 油品细类 分布
wb2 = openpyxl.load_workbook(os.path.join(BASE, 'output', '购进数据选择.xlsx'), data_only=True)
rows2 = list(wb2.worksheets[0].iter_rows(values_only=True))
h2 = {c: i for i, c in enumerate(rows2[0])}
p('TMP 油品类型x细类 分布:')
for k, n in Counter((r[h2['油品类型']], r[h2['油品细类']]) for r in rows2[1:]).most_common():
    p('  ', k, n)
p('TMP 表头:', rows2[0])

# 判定文件里 石脑油 相关
for fn in ['成品油发票判定_Energy进项20260501-20261026.xlsx', '成品油发票判定_Energy销项.xlsx']:
    wb3 = openpyxl.load_workbook(os.path.join(BASE, 'output', fn), data_only=True)
    rows3 = list(wb3['发票明细判定'].iter_rows(values_only=True))
    h3 = {c: i for i, c in enumerate(rows3[0])}
    cnt = Counter((r[h3['油品分类（大类）']], r[h3['油品细分（细类）']]) for r in rows3[1:]
                  if r[h3['是否成品油']] == '是' and '石脑' in str(r[h3['油品分类（大类）']]) + str(r[h3['油品细分（细类）']]))
    p(fn, '石脑油相关:', dict(cnt))
    cnt2 = Counter((r[h3['油品分类（大类）']], r[h3['油品细分（细类）']]) for r in rows3[1:]
                   if r[h3['是否成品油']] == '是')
    p(fn, '全部大类分布:')
    for k, n in sorted(cnt2.items()):
        p('   ', k, n)

with open(os.path.join(BASE, 'output', '_naphtha_check.txt'), 'w', encoding='utf-8') as f:
    f.write(buf.getvalue())
print('saved')
