# -*- coding: utf-8 -*-
"""扫描 发票信息 sheet：备注中含 Titan SI No、但 Titan SI No 栏位为空的发票。
行数取自原始 sheet XML（openpyxl 的 max_row 可能偏小）。"""
import json
import re
import sys
import zipfile

import openpyxl

SRC = r"titan-si-nunber/发票总览 (5).xlsx"
DST = r"titan-si-nunber/备注含TitanSI号但栏位为空.xlsx"
PROBE = "_scan_result.json"


def true_row_count(xlsx_path, sheet_name):
    """从 XML 层面取该 sheet 的真实最大行号。"""
    with zipfile.ZipFile(xlsx_path) as z:
        wbxml = z.read("xl/workbook.xml").decode("utf-8")
        m = re.search(r'name="%s"[^>]*r:id="(rId\d+)"' % sheet_name, wbxml)
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        target = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % m.group(1), rels).group(1)
        xml = z.read("xl/" + target).decode("utf-8")
        rows = re.findall(r'<row r="(\d+)"', xml)
        return max(int(r) for r in rows)


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["发票信息"]

    # 表头定位（不写死列号，防列顺序变化）
    hdr = {}
    for j in range(1, ws.max_column + 1):
        v = ws.cell(1, j).value
        if v is not None:
            hdr[str(v).strip()] = j
    need = ["发票号码", "销方名称", "购方名称", "发票状态", "发票类型",
            "发票日期", "含税金额", "备注", "Titan SI No"]
    missing = [h for h in need if h not in hdr]
    if missing:
        sys.exit("缺少列: %s; 实际表头: %s" % (missing, sorted(hdr)))

    last = true_row_count(SRC, "发票信息")

    # SI 格式：实体前缀(2-6字母)+SI/SS/SC+5位以上数字，前后不能是字母数字
    pat = re.compile(r"(?i)(?:^|[^A-Za-z0-9])([A-Za-z]{2,6}(?:SI|SS|SC)\s*[-:]?\s*\d{5,})")

    out = openpyxl.Workbook()
    o = out.active
    o.title = "备注含TitanSI但栏位为空"
    o.append(["发票号码", "销方名称", "购方名称", "发票状态", "发票类型",
              "发票日期", "含税金额", "备注", "找到的Titan SI No"])

    matched_rows, empty_si, total = [], 0, 0
    for i in range(2, last + 1):
        total += 1
        si = ws.cell(i, hdr["Titan SI No"]).value
        if si in (None, ""):
            empty_si += 1
            note = ws.cell(i, hdr["备注"]).value
            if note:
                m = pat.search(str(note))
                if m:
                    matched_rows.append([
                        ws.cell(i, hdr["发票号码"]).value,
                        ws.cell(i, hdr["销方名称"]).value,
                        ws.cell(i, hdr["购方名称"]).value,
                        ws.cell(i, hdr["发票状态"]).value,
                        ws.cell(i, hdr["发票类型"]).value,
                        ws.cell(i, hdr["发票日期"]).value,
                        ws.cell(i, hdr["含税金额"]).value,
                        ws.cell(i, hdr["备注"]).value,
                        m.group(1),
                    ])

    for r in matched_rows:
        o.append(r)
    out.save(DST)

    with open(PROBE, "w", encoding="utf-8") as f:
        json.dump({"true_last_row": last, "total": total, "empty_si": empty_si,
                   "matched": len(matched_rows), "rows": matched_rows},
                  f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()
