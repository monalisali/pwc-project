# -*- coding: utf-8 -*-
"""Compare commodity_code in 托克生产税收分类编码.xlsx against sphfwssflhbbm in
乐企销项开票税收分类编码.json (data.ssbmList). Output missing codes to Excel."""
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

XLSX_PATH = r"tax-code\托克生产税收分类编码.xlsx"
JSON_PATH = r"tax-code\乐企销项开票税收分类编码.json"
OUT_PATH = r"tax-code\托克税收分类编码比对结果.xlsx"

# --- Load JSON codes (sphfwssflhbbm under data.ssbmList) ---
with open(JSON_PATH, encoding="utf-8") as f:
    data = json.load(f)
ssbm_list = data.get("data", {}).get("ssbmList", [])
json_codes = set()
for item in ssbm_list:
    code = item.get("sphfwssflhbbm")
    if code:
        json_codes.add(str(code).strip())
print(f"JSON ssbmList entries: {len(ssbm_list)}, unique codes: {len(json_codes)}")

# --- Load Excel rows ---
wb = openpyxl.load_workbook(XLSX_PATH, read_only=True)
ws = wb.worksheets[0]
rows = list(ws.iter_rows(values_only=True))
header = list(rows[0])
idx = {name: header.index(name) for name in
       ("commodity_code", "commodity_name", "commodity_name_abbr",
        "commodity_description", "tax_rate")}
print(f"Excel data rows: {len(rows) - 1}")

missing_rows = []
excel_codes = set()
for r in rows[1:]:
    code = r[idx["commodity_code"]]
    if code is None:
        continue
    code = str(code).strip()
    excel_codes.add(code)
    if code not in json_codes:
        missing_rows.append([
            code,
            r[idx["commodity_name"]],
            r[idx["commodity_name_abbr"]],
            r[idx["commodity_description"]],
            r[idx["tax_rate"]],
        ])
wb.close()
print(f"Excel unique codes: {len(excel_codes)}")
print(f"Missing codes (rows): {len(missing_rows)}, unique missing codes: {len(set(m[0] for m in missing_rows))}")

# --- Write output Excel ---
out = openpyxl.Workbook()
ows = out.active
ows.title = "缺失税收分类编码"
headers = ["税收分类编码", "commodity_name", "commodity_name_abbr",
           "commodity_description", "tax_rate"]
ows.append(headers)
bold = Font(bold=True)
fill = PatternFill("solid", start_color="DDEBF7")
for c in ows[1]:
    c.font = bold
    c.fill = fill
    c.alignment = Alignment(horizontal="center")
for row in missing_rows:
    ows.append(row)
widths = [22, 40, 30, 80, 10]
for i, w in enumerate(widths, 1):
    ows.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ows.freeze_panes = "A2"
out.save(OUT_PATH)
print(f"Saved: {OUT_PATH}")
