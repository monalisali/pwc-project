"""Dump FR tables from both docs into readable per-item blocks for comparison."""
import json
import re
import difflib

d = json.load(open("_pi_tables.json", encoding="utf-8"))


def n(s):
    return re.sub(r"\s+", " ", s.replace("\n", " ")).strip()


def dump(doc_key, out_path):
    tables = d[doc_key]
    fr = tables[5]
    with open(out_path, "w", encoding="utf-8") as f:
        for row in fr[1:]:
            # pad row to 6 columns
            cells = (row + [""] * 6)[:6]
            fid, item, desc, system, pwc, phase = cells
            f.write(f"=== {n(fid)} | Item: {n(item)} | System: {n(system)} | 第一期: {n(phase) or '(空)'}\n")
            f.write(f"DESC: {n(desc)}\n")
            f.write(f"PWC: {n(pwc) or '(空)'}\n\n")
    return fr


fr21 = dump("0921", "_fr_0921.txt")
fr29 = dump("0929", "_fr_0929.txt")

# quick similarity matching preview
def key(row):
    cells = (row + [""] * 6)[:6]
    return n(cells[1]).lower(), n(cells[2]).lower()

pairs = []
used = set()
for r1 in fr21[1:]:
    k1 = key(r1)
    best, bs = None, 0
    for j, r2 in enumerate(fr29[1:]):
        if j in used:
            continue
        k2 = key(r2)
        s = difflib.SequenceMatcher(None, k1[0] + " " + k1[1], k2[0] + " " + k2[1]).ratio()
        if s > bs:
            bs, best = s, j
    if best is not None and bs > 0.55:
        used.add(best)
        pairs.append((n(r1[0]), n(fr29[1:][best][0]), round(bs, 2)))
    else:
        pairs.append((n(r1[0]), None, round(bs, 2)))

print("0921 -> 0929 matches (ID, ID, desc+item sim):")
for p in pairs:
    print("  ", p)
print("\nUnmatched 0929 rows:", [n(r[0]) for j, r in enumerate(fr29[1:]) if j not in used])
