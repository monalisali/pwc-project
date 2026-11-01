# -*- coding: utf-8 -*-
"""Generate Energy初始数据.xlsx from the transcribed table image."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()
ws = wb.active
ws.title = "Energy初始数据"

# ---------- styles ----------
thin = Side(style="thin")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
right = Alignment(horizontal="right", vertical="center")

bold = Font(bold=True)
italic = Font(italic=True)

yellow = PatternFill("solid", fgColor="FFD966")   # highlighted cells
green = PatternFill("solid", fgColor="C6E0B4")    # 合计 rows
blue = PatternFill("solid", fgColor="9DC3E6")     # 成品油合计 row

NUM3 = '#,##0.000;(#,##0.000);"-"'

# ---------- header ----------
headers = [
    "品名", "规格",
    "采购未收票\nA", "采购预收票\nB", "销售未开票\nC",
    "WDMS stock as of 0716\nD",
    "VAT库存as of 0716 Ops\nH=D-A+B+C",
    "VAT库存as of 0716 税局\nF",
    "Diff\nE-F", "Ops",
]
for col, h in enumerate(headers, 1):
    c = ws.cell(row=1, column=col, value=h)
    c.font = bold
    c.alignment = center
    c.border = border
ws.row_dimensions[1].height = 42

# ---------- data ----------
rows = {
    2:  {"A": "组分油", "B": "95#组分油小计"},
    3:  {"A": "组分油合计", "C": 5363.589, "D": 619.867, "G": 5983.456, "J": "Iris",
         "G_yellow": True},
    4:  {"A": "柴油", "B": "0#", "C": 4500.000, "D": 2332.532, "G": 6832.532, "J": "candice",
         "yellow": True},
    5:  {"B": "0#京标", "C": 499.880, "D": 1501.120, "G": 2001.000, "J": "candice",
         "yellow": True},
    6:  {"B": "-35#", "C": 319.384, "G": 319.384, "J": "Iris"},
    7:  {"B": "0#小计", "A1": 0, "Bv": 0, "C": 10363.469, "D": 4453.519, "G": 14816.988,
         "italic": True},
    8:  {"B": "燃料油*DMX"},
    9:  {"B": "燃料油*DMA"},
    10: {"A": "柴油合计", "A1": 0, "Bv": 0, "C": 10682.853, "D": 4453.519, "G": 15136.372,
         "H": 16359.4695, "I": -1223.098, "green": True, "bold": True},
    11: {"A": "汽油", "B": "92#", "A1": 6833.868, "C": 6629.020, "D": 6330.169, "G": 6125.321,
         "J": "sherry"},
    12: {"C": 3986.498, "G": 3986.498, "J": "Iris"},
    13: {"C": 2995.150, "G": 2995.150, "J": "candice"},
    14: {"B": "92#小计", "A1": 6833.868, "Bv": 0, "C": 13610.668, "D": 6330.169,
         "G": 13106.969, "italic": True},
    15: {"B": "95#", "C": 1447.540, "D": 1759.463, "G": 3207.003, "J": "sherry"},
    16: {"A1": 1405.040, "C": 23224.257, "G": 21819.217, "J": "Iris"},
    17: {"A1": 13930.949, "D": 4880.600, "G": -9050.349, "J": "candice"},
    18: {"B": "95#小计", "A1": 15335.989, "Bv": 0, "C": 24671.797, "D": 6640.063,
         "G": 15975.871, "italic": True},
    19: {"B": "92#京标小计"},
    20: {"B": "95#京标小计", "G": 6053.260, "J": "candice", "G_yellow": True},
    21: {"A": "汽油合计", "A1": 22169.857, "Bv": 0, "C": 38282.465, "D": 12970.232,
         "G": 35136.100, "H": 38254.93399, "I": -3118.834, "green": True, "bold": True},
    22: {"A": "石脑油", "A1": 4788.833, "D": 4788.833, "G": 0, "J": "Iris",
         "G_yellow": True},
    23: {"A": "石脑油合计", "A1": 4788.833, "D": 4788.833, "G": 0, "H": 185.169,
         "I": -185.169, "green": True, "bold": True},
    24: {"A": "燃料油合计", "bold": True},
    25: {"A": "成品油合计", "D": 22212.584, "G": 50272.472, "H": 54799.57249,
         "I": -4527.100, "blue": True, "bold": True},
}

# column map in sheet: C=采购未收票A, D=采购预收票B, E=销售未开票C,
# F=WDMS D, G=VAT Ops H, H=VAT 税局 F, I=Diff, J=Ops
key2col = {"A1": 3, "Bv": 4, "C": 5, "D": 6, "G": 7, "H": 8, "I": 9}

for r, spec in rows.items():
    if "A" in spec:
        cell = ws.cell(row=r, column=1, value=spec["A"])
        cell.alignment = center
    if "B" in spec:
        cell = ws.cell(row=r, column=2, value=spec["B"])
        cell.alignment = center
    for key, col in key2col.items():
        if key in spec:
            cell = ws.cell(row=r, column=col, value=spec[key])
            cell.number_format = NUM3
            cell.alignment = right
    # 税局 F / Diff with more than 3 decimals
    if "H" in spec and abs(spec["H"] * 1000 - round(spec["H"] * 1000)) > 1e-9:
        ws.cell(row=r, column=8).number_format = '#,##0.00000;(#,##0.00000)'
    if "J" in spec:
        cell = ws.cell(row=r, column=10, value=spec["J"])
        cell.alignment = center
    for col in range(1, 11):
        cell = ws.cell(row=r, column=col)
        cell.border = border
        if spec.get("yellow") and col in (2, 7):
            cell.fill = yellow
        if spec.get("G_yellow") and col == 7:
            cell.fill = yellow
        if spec.get("green"):
            cell.fill = green
        if spec.get("blue"):
            cell.fill = blue
        if spec.get("italic") and col in range(3, 8):
            cell.font = italic
        if spec.get("bold"):
            f = cell.font
            cell.font = Font(bold=True, italic=f.italic)

# ---------- merges ----------
ws.merge_cells("A4:A9")    # 柴油
ws.merge_cells("B11:B13")  # 92#
ws.merge_cells("B15:B17")  # 95#

# ---------- column widths ----------
widths = [10, 14, 12, 12, 12, 14, 16, 14, 11, 9]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws.freeze_panes = "A2"

out = r"C:\D\AIWork\pwc-project\trafigura-tmp\energy\Energy初始数据.xlsx"
wb.save(out)
print("saved:", out)
