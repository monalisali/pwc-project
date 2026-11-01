import io
import json
import re
import sys
from difflib import SequenceMatcher, unified_diff

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
d = json.load(open("_fr_compare.json", encoding="utf-8"))
A = {a["id"]: a for a in d["A"]}
B = {b["id"]: b for b in d["B"]}

print("========== matched pairs with description changes ==========")
for m in d["matched"]:
    if not m["b_id"]:
        continue
    a, b = A[m["a_id"]], B[m["b_id"]]
    r = SequenceMatcher(None, re.sub(r"\s+", " ", a["desc"]), re.sub(r"\s+", " ", b["desc"])).ratio()
    if r < 0.99:
        print(f"--- {m['a_id']} -> {m['b_id']} (desc sim {r:.2f}) item: {a['item'][:60]}")
        print("  0929:", re.sub(r"\s+", " ", a["desc"])[:600])
        print("  1008:", re.sub(r"\s+", " ", b["desc"])[:600])
        print()

print("========== section headings diff ==========")


def headings(path):
    hs = []
    for line in open(path, encoding="utf-8"):
        s = line.strip()
        m = re.match(r"^(\d+(?:\.\d+)*)\.?\s+\S", s)
        if m and len(s) < 90 and not s.startswith("@"):
            hs.append(s)
    return hs


h1, h2 = headings("_pi_0929.txt"), headings("_pi_1008.txt")
print("--- 0929 only:")
for h in h1:
    if h not in h2:
        print("  ", h)
print("--- 1008 only:")
for h in h2:
    if h not in h1:
        print("  ", h)
