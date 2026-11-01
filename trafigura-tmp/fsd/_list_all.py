import io
import json
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
d = json.load(open("_fr_compare.json", encoding="utf-8"))

print("========== 0929 FR list (30) ==========")
for a in d["A"]:
    print(f"{a['id']:7} | {a['phase1']:1} | {a['system'][:25]:25} | {a['item'][:70]}")
print()
print("========== 1008 FR list (37) ==========")
for b in d["B"]:
    print(f"{b['id']:7} | {b['last']:6} | {b['system'][:25]:25} | {b['item'][:70]}")
print()
print("========== matched pairs ==========")
for m in d["matched"]:
    flag = "" if m["a_id"] == m["b_id"] else "  <<< ID CHANGED"
    print(f"{m['a_id']} -> {m['b_id']:7} score={m['score']:5} phase1={m['a_phase1']:1} fix={m['b_fix']:6} | {m['a_item'][:50]}{flag}")
