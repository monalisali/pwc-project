# -*- coding: utf-8 -*-
"""Compare codes in 托克生产实际开票税收分类编码.txt against sphfwssflhbbm in
乐企销项开票税收分类编码.json. For missing codes, enrich columns from
托克生产税收分类编码.xlsx."""
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

TXT_PATH = r"tax-code\托克生产实际开票税收分类编码.txt"
JSON_PATH = r"tax-code\乐企真实_销项开票税收分类编码.json"
XLSX_PATH = r"tax-code\托克生产税收分类编码.xlsx"
OUT_PATH = r"tax-code\实际开票税收分类编码比对结果.xlsx"

# --- Actual invoiced codes from txt ---
with open(TXT_PATH, encoding="utf-8") as f:
    txt_codes = [line.strip() for line in f if line.strip()]
print(f"txt codes (lines): {len(txt_codes)}, unique: {len(set(txt_codes))}")

# --- JSON codes (sphfwssflhbbm under data.ssbmList) ---
with open(JSON_PATH, encoding="utf-8") as f:
    data = json.load(f)
ssbm_list = data.get("data", {}).get("ssbmList", [])
json_codes = set()
for item in ssbm_list:
    code = item.get("sphfwssflhbbm")
    if code:
        json_codes.add(str(code).strip())
print(f"JSON ssbmList entries: {len(ssbm_list)}, unique codes: {len(json_codes)}")

# --- Excel lookup for enrichment ---
wb = openpyxl.load_workbook(XLSX_PATH, read_only=True)
ws = wb.worksheets[0]
rows = list(ws.iter_rows(values_only=True))
header = list(rows[0])
idx = {name: header.index(name) for name in
       ("commodity_code", "commodity_name", "commodity_name_abbr",
        "commodity_description", "tax_rate")}
xlsx_info = {}
for r in rows[1:]:
    code = r[idx["commodity_code"]]
    if code is None:
        continue
    code = str(code).strip()
    if code not in xlsx_info:  # keep first occurrence
        xlsx_info[code] = (r[idx["commodity_name"]],
                           r[idx["commodity_name_abbr"]],
                           r[idx["commodity_description"]],
                           r[idx["tax_rate"]])
wb.close()
print(f"Excel unique codes: {len(xlsx_info)}")

# --- Compare: txt codes missing from JSON ---
missing_rows = []
seen = set()
not_in_json_not_in_xlsx = []
for code in txt_codes:
    if code in seen:
        continue
    seen.add(code)
    if code not in json_codes:
        info = xlsx_info.get(code, ("", "", "", ""))
        if code not in xlsx_info:
            not_in_json_not_in_xlsx.append(code)
        missing_rows.append([code, *info])

print(f"Missing codes: {len(missing_rows)}")
print(f"  of which also NOT found in Excel: {len(not_in_json_not_in_xlsx)}")
if not_in_json_not_in_xlsx:
    print("  codes not in Excel either:", not_in_json_not_in_xlsx)

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
