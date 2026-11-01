# -*- coding: utf-8 -*-
"""
成品油发票数量比对 - 完整流程总控
================================
流程:
  步骤1  原始进项/销项发票 -> 成品油判定文件        (oil_classify.py / oil_classify_sales.py)
  步骤2  判定文件 vs 税局 vs TMP购进 -> 比对文件     (compare_category.py, 按大类)
  步骤3  初始 + 进项比对 - 销项比对 -> 台账结果文件  (ledger_category.py, 按大类)
  库存   剩余可用库存计算 (按大类/细类)
全部输出文件 (含时间戳后缀 _yyyymmddhhmmss) 统一写入: energy/output/库存计算/

用法 (在 energy/output 目录下):
  python -I run_pipeline.py            # 跑完整流程 (默认)
  python -I run_pipeline.py snapshot   # 把当前输出快照到 archive/baseline, 作为校验基准
  python -I run_pipeline.py validate   # 当前输出 vs 基准, 逐单元格比对, 报告写入 _validate_report.txt
  python -I run_pipeline.py stock      # 计算剩余可用库存 (缺参数时交互式询问)
    可选参数:
      --level 大类|细类        计算口径 (默认细类)
      --cats 汽油,柴油         大类名称, 逗号分隔, 可多个 (默认汽油)
      --subs 92号车用汽油（VIB）,95号车用汽油（VIB）   细类名称 (level=细类时必填, 交互输入)
      --no-run                 不重跑流程, 直接用已有输出文件计算
    示例:
      python -I run_pipeline.py stock --level 细类 --cats 汽油 --subs 92号车用汽油（VIB）,95号车用汽油（VIB）,98号车用汽油（VIB）,95号车用汽油（京VIB）,92号车用汽油（京VIB）

换数据时只需: 1) 替换 energy/ 下的原始进项/销项文件  2) 改下方 CONFIG 里的文件名与窗口
"""
import argparse
import os
import re
import shutil
import sys
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import oil_classify
import oil_classify_sales
import compare_category
import ledger_category
import stock_report

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # .../energy
OUT = os.path.join(BASE, 'output')
OUT_DIR = os.path.join(OUT, '库存计算')   # run_pipeline 全部输出 (判定/比对/台账/库存) 统一放这里
SNAP_DIR = os.path.join(OUT, 'archive', 'baseline')

TS = None   # 本次 run 的时间戳 (yyyymmddhhmmss), run() 开始时赋值

# ============================ 配置 ============================
CONFIG = {
    # --- 源文件 (相对 energy/ 目录; 换数据改这里) ---
    '进项源文件': 'Energy进项20260501-20261026.xlsx',
    '销项源文件': 'Energy销项.xlsx',
    '税局成品油发票': '税局Energy增值税发票7.1-10.23.xlsx',
    'TMP购进数据': 'output/购进数据选择.xlsx',
    '初始数据': 'Energy初始数据.xlsx',
    # --- 名义取数窗口 (即使窗口起点当天无数据也按名义窗口记录) ---
    '进项源文件窗口': (date(2026, 5, 1), date(2026, 10, 26)),
    '税局取数起始日': date(2026, 7, 16),
    '税局取数截止日': date(2026, 10, 23),
    # --- 比对的大类列表 (以后扩展柴油等, 加到这里并在"大类配置"补配置) ---
    '比对大类': ['汽油'],
    # --- 每个大类的比对配置 ---
    '大类配置': {
        '汽油': {
            '税局商品名称': '汽油',        # 税局文件"商品和服务名称"的取值
            'TMP油品类型': 1,              # TMP购进"油品类型"编码 (1=汽油, 2=柴油, 3=其他)
            'TMP剔除细类': (10, 11),       # TMP购进剔除的"油品细类"，原来是（10，11，）10：95号车用乙醇汽油；11：92号车用乙醇汽油；
        },
        '石脑油': {
            '税局商品名称': '石脑油',
            'TMP油品类型': 3,
            'TMP剔除细类': (),
        },
    },
    # --- 台账初始数据: 行标签按 f'{大类}合计' 取; 每个初始列生成一个场景 ---
    '初始数据列': [
        ('G', 'VAT库存as of 0716 Ops, H=D-A+B+C'),
        ('H', 'VAT库存as of 0716 税局, F'),
    ],
    # --- 细类库存计算: 初始数据(G列)行的匹配关键词 (匹配 品名/规格 包含全部关键词);
    #     取不到或G列为空按0计, 报告中注明来源 ---
    # --- 库存计算(细类)初始数据: 每个细类的匹配关键词 ('!'前缀=减项); 大类口径用 '大类初始匹配' ---
    '大类初始匹配': {
        '石脑油': ['石脑油合计'],
    },
    '细类初始匹配': {
        '92号车用汽油（VIB）': ['92#小计'],
        '95号车用汽油（VIB）': ['95#小计'],
        '98号车用汽油（VIB）': [],
        '95号车用汽油（京VIB）': ['95#京标小计'],
        '92号车用汽油（京VIB）': ['92#京标小计'],
        '95号车用乙醇汽油调合组分': ['95#组分油小计'],
        '92号车用乙醇汽油调合组分': ['92#组分油小计'],
        # '!' 前缀 = 减项: 0#小计 含 0#京标, 0号VI 取 0#小计-0#京标, 避免与 0号京VI 重复计 2001
        '0号车用柴油（VI）': ['0#小计', '!0#京标'],
        '-35号车用柴油（VI）': ['-35#小计'],
        '0号车用柴油（京VI）': ['0#京标'],
        'DMX船用燃料油': ['DMX'],
        'DMA船用燃料油': ['DMA'],
    },
}
# ==============================================================


def p(*a):
    print(*a, file=sys.stderr)


def join(name):
    return os.path.join(BASE, name)


def stamp():
    return datetime.now().strftime('%Y%m%d%H%M%S')


def strip_ts(name):
    """去掉文件名中的时间戳后缀 (_yyyymmddhhmmss), 用于快照/校验的基准名"""
    return re.sub(r'_\d{14}(?=\.xlsx$)', '', name)


def latest_run_ts():
    """输出目录中最新一次 run 的时间戳 (无则 None)"""
    tss = [m.group(1) for d in (OUT_DIR, OUT) if os.path.isdir(d)
           for fn in os.listdir(d)
           for m in [re.search(r'_(\d{14})\.xlsx$', fn)] if m]
    return max(tss) if tss else None


def out_ts():
    """当前应引用的时间戳: 本次运行 > 输出目录最新一次 > None(旧命名)"""
    return TS or latest_run_ts()


def show_inputs():
    """启动时显示输入文件名称与各自取数窗口 (口径与比对/台账实际用法一致)"""
    cfg = CONFIG
    tax_win = f"{cfg['税局取数起始日']} ~ {cfg['税局取数截止日']}"
    as_of = cfg['税局取数起始日']
    rows = [
        ('进项源文件', cfg['进项源文件'],
         f"{cfg['进项源文件窗口'][0]} ~ {cfg['进项源文件窗口'][1]} (名义窗口)"),
        ('销项源文件', cfg['销项源文件'],
         '未单独配置 (比对窗口由各方实际数据范围取交集)'),
        ('税局成品油发票', cfg['税局成品油发票'],
         f'{tax_win} (自起始日起含当天)'),
        ('TMP购进数据', cfg['TMP购进数据'],
         '不设窗口 (实际数据范围参与比对窗口计算)'),
        ('初始数据', cfg['初始数据'],
         f'as of {as_of} (取 G/H 列期初库存)'),
    ]
    p('=' * 64)
    p('输入文件及取数窗口:')
    for label, fname, win in rows:
        mark = 'OK' if os.path.exists(join(fname)) else '!! 文件不存在'
        p(f'  {label}: {fname}  [{mark}]')
        p(f'    取数窗口: {win}')
    p('=' * 64)


def run():
    global TS
    TS = stamp()
    os.makedirs(OUT_DIR, exist_ok=True)
    cfg = CONFIG
    # 步骤1: 成品油判定
    src_in = join(cfg['进项源文件'])
    dst_in = os.path.join(OUT_DIR, f'成品油发票判定_{os.path.basename(src_in)[:-5]}_{TS}.xlsx')
    p('步骤1a 进项判定:', src_in, '->', os.path.basename(dst_in))
    st1 = oil_classify.run(src_in, dst_in)

    src_out = join(cfg['销项源文件'])
    dst_out = os.path.join(OUT_DIR, f'成品油发票判定_{os.path.basename(src_out)[:-5]}_{TS}.xlsx')
    p('步骤1b 销项判定:', src_out, '->', os.path.basename(dst_out))
    st2 = oil_classify_sales.run(src_out, dst_out)

    tax_path = join(cfg['税局成品油发票'])
    tmp_path = join(cfg['TMP购进数据'])
    init_path = join(cfg['初始数据'])
    tax_from = cfg['税局取数起始日']
    tax_window = (tax_from, cfg['税局取数截止日'])

    for cat in cfg['比对大类']:
        cc = cfg['大类配置'][cat]
        # 步骤2: 比对
        cmp_in = os.path.join(OUT_DIR, f'{cat}数量比对_进项_三方_{TS}.xlsx')
        p(f'步骤2a {cat} 进项三方比对 ->', os.path.basename(cmp_in))
        compare_category.run_input(cat, dst_in, tax_path, tmp_path, cmp_in,
                                   tax_from=tax_from, res_window=cfg['进项源文件窗口'],
                                   tax_window=tax_window, tmp_type=cc['TMP油品类型'],
                                   tmp_exclude_sub=cc['TMP剔除细类'])
        cmp_out = os.path.join(OUT_DIR, f'{cat}数量比对_销项vs税局_{TS}.xlsx')
        p(f'步骤2b {cat} 销项vs税局比对 ->', os.path.basename(cmp_out))
        compare_category.run_output(cat, dst_out, tax_path, cmp_out,
                                    tax_from=tax_from, tax_window=tax_window)
        # 步骤3: 台账
        ledger = os.path.join(OUT_DIR, f'{cat}数量结果_初始加进项减销项_{TS}.xlsx')
        p(f'步骤3 {cat} 台账 ->', os.path.basename(ledger))
        st3 = ledger_category.run(cat, init_path, cmp_in, cmp_out, ledger,
                                  start_date=tax_from, init_cols=cfg['初始数据列'])
        for s_ in st3['场景']:
            p(f"   {s_['名称']}: 最终={s_['最终']}, 首次转负={s_['首次转负']}")
    p('流程完成')


def expected_outputs(ts=None):
    """当前配置下应生成的输出文件路径; ts 缺省取 本次运行/输出目录最新一次 的时间戳"""
    ts = ts or out_ts()
    suf = f'_{ts}' if ts else ''
    src_in = join(CONFIG['进项源文件'])
    src_out = join(CONFIG['销项源文件'])
    files = [os.path.join(OUT_DIR, f'成品油发票判定_{os.path.basename(src_in)[:-5]}{suf}.xlsx'),
             os.path.join(OUT_DIR, f'成品油发票判定_{os.path.basename(src_out)[:-5]}{suf}.xlsx')]
    for cat in CONFIG['比对大类']:
        files += [os.path.join(OUT_DIR, f'{cat}数量比对_进项_三方{suf}.xlsx'),
                  os.path.join(OUT_DIR, f'{cat}数量比对_销项vs税局{suf}.xlsx'),
                  os.path.join(OUT_DIR, f'{cat}数量结果_初始加进项减销项{suf}.xlsx')]
    return files


def snapshot():
    os.makedirs(SNAP_DIR, exist_ok=True)
    for fp in expected_outputs():
        # 基准文件去掉时间戳后缀存放, 保证不同时间戳的输出可与同一基准校验
        base_name = strip_ts(os.path.basename(fp))
        shutil.copy2(fp, os.path.join(SNAP_DIR, base_name))
        p('快照:', base_name)
    p('基准目录:', SNAP_DIR)


def _eq(a, b):
    """单元格值比较: 数字按容差, None 等价空串"""
    if a is None and (b is None or b == ''):
        return True
    if b is None and a == '':
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 1e-9
    return a == b


def validate():
    lines = []
    n_diff_files = 0
    for fp in expected_outputs():
        name = strip_ts(os.path.basename(fp))
        base = os.path.join(SNAP_DIR, name)
        if not os.path.exists(base):
            lines.append(f'[缺基准] {name}')
            n_diff_files += 1
            continue
        if not os.path.exists(fp):
            lines.append(f'[缺当前文件] {name}')
            n_diff_files += 1
            continue
        wb_a = _load(fp)
        wb_b = _load(base)
        if wb_a is None or wb_b is None:
            lines.append(f'[无法打开] {name}: 当前={wb_a is not None}, 基准={wb_b is not None}')
            n_diff_files += 1
            continue
        diffs = _diff_workbook(wb_a, wb_b)
        if diffs:
            n_diff_files += 1
            lines.append(f'[不一致] {name}: {len(diffs)} 处')
            lines.extend('    ' + d for d in diffs[:50])
        else:
            lines.append(f'[一致] {name}')
    header = f'校验完成: {len(expected_outputs())} 个文件, {n_diff_files} 个不一致'
    lines.insert(0, header)
    rpt = os.path.join(OUT, '_validate_report.txt')
    with open(rpt, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(lines) + '\n')
    p(header, '| 报告:', rpt)


def _load(path):
    try:
        import openpyxl
        return openpyxl.load_workbook(path, data_only=True)
    except Exception as e:   # noqa: BLE001
        return None


def _diff_workbook(wa, wb):
    diffs = []
    if wa.sheetnames != wb.sheetnames:
        diffs.append(f'sheet不一致: {wa.sheetnames} vs {wb.sheetnames}')
    for sn in wa.sheetnames:
        if sn not in wb.sheetnames:
            continue
        sa, sb = wa[sn], wb[sn]
        max_row = max(sa.max_row, sb.max_row)
        max_col = max(sa.max_column, sb.max_column)
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                va = sa.cell(r, c).value
                vb = sb.cell(r, c).value
                if not _eq(va, vb):
                    diffs.append(f'{sn}!{sa.cell(r, c).coordinate}: {va!r} vs {vb!r}')
                    if len(diffs) > 200:
                        diffs.append('...(超过200处, 截断)')
                        return diffs
    return diffs


def stock(level, cats, subs, rerun=True):
    """剩余可用库存计算, 报告输出到 库存计算/"""
    cfg = CONFIG
    if rerun:
        run()
    # 判定文件路径 (stock 需要; 若不存在先跑流程); 引用最近一次 run 的时间戳
    src_in = join(cfg['进项源文件'])
    src_out = join(cfg['销项源文件'])
    suf = f'_{out_ts()}' if out_ts() else ''
    det_in = os.path.join(OUT_DIR, f'成品油发票判定_{os.path.basename(src_in)[:-5]}{suf}.xlsx')
    det_out = os.path.join(OUT_DIR, f'成品油发票判定_{os.path.basename(src_out)[:-5]}{suf}.xlsx')
    if not (os.path.exists(det_in) and os.path.exists(det_out)):
        p('判定文件不存在, 先跑完整流程...')
        run()
    os.makedirs(stock_report.STOCK_DIR, exist_ok=True)
    sts = TS or stamp()   # 库存报告时间戳: 沿用本次 run 的, 独立运行时取当前时间

    if level == '大类':
        # 大类口径: 同一套报告引擎, 每个大类一个桶 (含NA细类行), 报告放 库存计算/
        dst = os.path.join(stock_report.STOCK_DIR, f"库存计算_大类_{'_'.join(cats)}_{sts}.xlsx")
        p(f'库存计算(大类) 大类={cats} ->', dst)
        st = stock_report.run('大类', cats, [], det_in, det_out,
                              join(cfg['初始数据']), dst,
                              start_date=cfg['税局取数起始日'],
                              init_config=cfg.get('大类初始匹配', {}))
        for r in st['细类结果']:
            p(f"   {r['细类']}: 初始={r['初始']} ({r['初始说明']}), "
              f"进项={r['进项合计']}, 销项={r['销项合计']}, 剩余可用库存={r['最终']}, "
              f"首次转负={r['首次转负']}")
    else:
        # 细类口径: 进销项均取判定文件明细, 未识别细类的发票在报告中注明
        dst = os.path.join(stock_report.STOCK_DIR, f"库存计算_细类_{'_'.join(cats)}_{sts}.xlsx")
        p(f'库存计算(细类) 大类={cats} 细类={subs} ->', dst)
        st = stock_report.run('细类', cats, subs, det_in, det_out,
                              join(cfg['初始数据']), dst,
                              start_date=cfg['税局取数起始日'],
                              init_config=cfg['细类初始匹配'])
        for r in st['细类结果']:
            p(f"   {r['细类']}: 初始={r['初始']} ({r['初始说明']}), "
              f"进项={r['进项合计']}, 销项={r['销项合计']}, 剩余可用库存={r['最终']}, "
              f"首次转负={r['首次转负']}")
        if st['未识别/未纳入明细行数']:
            p(f"   !! 有 {st['未识别/未纳入明细行数']} 行明细未识别/未纳入细类 "
              f"(涉及 {st['涉及发票数']} 张发票), 详见报告")
    p('库存计算完成, 报告目录:', stock_report.STOCK_DIR)


def main():
    ap = argparse.ArgumentParser(description='成品油发票数量比对流程')
    ap.add_argument('mode', nargs='?', default='run',
                    choices=['run', 'snapshot', 'validate', 'stock'])
    ap.add_argument('--level', default=None, help='库存计算口径: 大类 / 细类 (默认细类)')
    ap.add_argument('--cats', default=None, help='大类名称, 逗号分隔, 可多个 (默认 汽油)')
    ap.add_argument('--subs', default=None, help='细类名称, 逗号分隔, 可多个')
    ap.add_argument('--no-run', action='store_true', help='stock 模式: 不重跑流程, 直接用已有输出')
    args = ap.parse_args()

    if args.mode in ('run', 'stock'):
        show_inputs()

    if args.mode == 'stock':
        # 缺参数时交互式询问
        level = args.level
        if level not in ('大类', '细类'):
            level = input('按大类还是细类计算剩余库存? (大类/细类, 回车=细类): ').strip() or '细类'
        cats = [c.strip() for c in args.cats.split(',')] if args.cats else None
        if not cats:
            raw = input('输入大类名称, 多个用逗号分隔 (回车=汽油): ').strip()
            cats = [c.strip() for c in raw.replace('，', ',').split(',') if c.strip()] or ['汽油']
        subs = [c.strip() for c in args.subs.split(',')] if args.subs else None
        if level == '细类' and not subs:
            raw = input('输入细类名称, 多个用逗号分隔: ').strip()
            subs = [c.strip() for c in raw.replace('，', ',').split(',') if c.strip()]
            if not subs:
                p('未输入细类, 退出')
                sys.exit(2)
        stock(level, cats, subs or [], rerun=not args.no_run)
        return

    if args.mode == 'run':
        run()
    elif args.mode == 'snapshot':
        snapshot()
    elif args.mode == 'validate':
        validate()


if __name__ == '__main__':
    main()
