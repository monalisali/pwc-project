# -*- coding: utf-8 -*-
import io, os
from collections import Counter
import openpyxl
BASE = r'c:\D\AIWork\pwc-project\trafigura-tmp\energy'
buf = io.StringIO()
p = lambda *a: print(*a, file=buf)

# 原始进项/销项 明细中 商品名称含 石脑油 的行
for fn, sheet, name_col, qty_col in [('Energy进项20260501-20261026.xlsx', '发票明细', '商品名称', '数量'),
                                     ('Energy销项.xlsx', '发票明细信息', '商品名称', '商品数量')]:
    wb = openpyxl.load_workbook(os.path.join(BASE, fn), data_only=True)
    rows = list(wb[sheet].iter_rows(values_only=True))
    h = {c: i for i, c in enumerate(rows[0])}
    hits = [r for r in rows[1:] if '石脑' in str(r[h[name_col]] or '')]
    p(f'{fn}: 含石脑油明细 {len(hits)} 行')
    tot = 0.0
    for r in hits[:8]:
        p('   ', str(r[h['发票号码']])[:24], str(r[h[name_col]])[:40], r[h[qty_col]])
    tot = sum(float(r[h[qty_col]] or 0) for r in hits if isinstance(r[h[qty_col]], (int, float)))
    p('    数量合计:', round(tot, 3))

# 税局 石脑油行: 变动方式 x 数量, 发票号样例
wb = openpyxl.load_workbook(os.path.join(BASE, '税局Energy增值税发票7.1-10.23.xlsx'), data_only=True)
rows = list(wb['增值税发票'].iter_rows(values_only=True))
h = {c: i for i, c in enumerate(rows[0])}
nh = [r for r in rows[1:] if str(r[h['商品和服务名称']]).strip() == '石脑油']
p('税局石脑油:', len(nh), '行, 变动方式分布:',
  dict(Counter(str(r[h['变动方式']]) for r in nh)),
  '数量合计:', round(sum(float(r[h['数量（吨）']] or 0) for r in nh), 3))
for r in nh[:5]:
    p('   ', str(r[h['数电发票号码']] or r[h['发票号码']]), str(r[h['开具日期']])[:10],
      r[h['数量（吨）']], r[h['变动方式']])

# TMP 类型3 的行
wb2 = openpyxl.load_workbook(os.path.join(BASE, 'output', '购进数据选择.xlsx'), data_only=True)
rows2 = list(wb2.worksheets[0].iter_rows(values_only=True))
h2 = {c: i for i, c in enumerate(rows2[0])}
t3 = [r for r in rows2[1:] if r[h2['油品类型']] == 3]
p('TMP油品类型3:', len(t3), '行')
for r in t3:
    p('   ', str(r[h2['发票号码']]), str(r[h2['开票日期']])[:10], r[h2['油品数量（吨）']],
      '细类:', r[h2['油品细类']])

with open(os.path.join(BASE, 'output', '_naphtha2.txt'), 'w', encoding='utf-8') as f:
    f.write(buf.getvalue())
print('saved')
