# -*- coding: utf-8 -*-
"""Compare energy_output_1014.xlsx (original) vs energy_output_1016.xlsx (updated) by quota."""
import pandas as pd

F1014 = "coce/energy_output_1014.xlsx"
F1016 = "coce/energy_output_1016.xlsx"
OUT = "coce/energy_output_compare_1014_vs_1016.xlsx"

KEY = "quota"
LEAD_COLS = ["quota", "business_line", "group_company", "commodity"]
CMP_FIELDS = [
    "this_month_actual_cargo_value",
    "previous_month_cargo_value_upstream",
    "this_month_actual_qty_upstream",
    "previous_month_qty",
    "unit_price_upstream",
]

df14 = pd.read_excel(F1014)
df16 = pd.read_excel(F1016)

# drop junk records: quota containing "P"
before = len(df14), len(df16)
df14 = df14[~df14[KEY].astype(str).str.contains("P", case=False)]
df16 = df16[~df16[KEY].astype(str).str.contains("P", case=False)]
print(f"dropped quota with P: 1014 {before[0]}->{len(df14)}, 1016 {before[1]}->{len(df16)}")

# outer join on quota, suffixes mark source file
m = df14.merge(df16, on=KEY, how="outer", suffixes=("_1014", "_1016"), indicator=True)

# leading columns: take from 1016, fall back to 1014 when missing there
for c in LEAD_COLS[1:]:
    m[c] = m[f"{c}_1016"].combine_first(m[f"{c}_1014"])

# value + diff columns
for f in CMP_FIELDS:
    v14, v16 = m[f"{f}_1014"], m[f"{f}_1016"]
    m[f"diff_{f}"] = v16 - v14  # diff = 1016(updated) - 1014(original)

out = m[[KEY] + LEAD_COLS[1:]].copy()
for f in CMP_FIELDS:
    out[f"1014_{f}"] = m[f"{f}_1014"]
    out[f"1016_{f}"] = m[f"{f}_1016"]
    out[f"diff_{f}"] = m[f"diff_{f}"]
out["remark"] = m["_merge"].map(
    {"left_only": "only in 1014", "right_only": "only in 1016", "both": ""}
)

with pd.ExcelWriter(OUT, engine="openpyxl") as w:
    out.to_excel(w, index=False, sheet_name="diff")
    ws = w.sheets["diff"]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    # red font on non-zero diffs
    from openpyxl.styles import Font
    red = Font(color="FFCC0000")
    for col_idx, name in enumerate(out.columns, start=1):
        if name.startswith("diff_"):
            for r in range(2, ws.max_row + 1):
                v = ws.cell(row=r, column=col_idx).value
                if isinstance(v, (int, float)) and abs(v) > 1e-9:
                    ws.cell(row=r, column=col_idx).font = red

n_only14 = (out["remark"] == "only in 1014").sum()
n_only16 = (out["remark"] == "only in 1016").sum()
n_changed = (out[[f"diff_{f}" for f in CMP_FIELDS]].abs().sum(axis=1) > 1e-9).sum()
print(f"rows: {len(out)} (both={len(out)-n_only14-n_only16}, only1014={n_only14}, only1016={n_only16})")
print(f"rows with any non-zero diff: {n_changed}")
print("saved ->", OUT)
