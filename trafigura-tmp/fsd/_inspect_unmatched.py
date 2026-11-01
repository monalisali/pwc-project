import io
import json
import re
import sys
from difflib import SequenceMatcher

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
d = json.load(open("_fr_compare.json", encoding="utf-8"))

print("0929 ids:", sorted((a["id"] for a in d["A"]), key=lambda x: int(x.split("-")[1])))
print("1008 ids:", sorted((b["id"] for b in d["B"]), key=lambda x: int(x.split("-")[1])))
print()


def norm(s):
    s = re.sub(r"\s+", " ", s).strip().lower()
    return re.sub(r"[^\w一-鿿]+", "", s)


for u in d["unmatched_b"]:
    b = next(x for x in d["B"] if x["id"] == u["b_id"])
    print("=" * 70)
    print(u["b_id"], "|", b["item"], "| Fix:", b["last"])
    print(b["desc"][:400].replace("\n", " / "))
    nb = norm(b["item"] + " " + b["desc"])
    scored = sorted(((SequenceMatcher(None, norm(a["item"] + " " + a["desc"]), nb).ratio(), a["id"]) for a in d["A"]), reverse=True)[:3]
    print("  best 0929 candidates:", [(i, round(s, 2)) for s, i in scored])
