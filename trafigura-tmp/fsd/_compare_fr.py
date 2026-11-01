"""Compare Functional Requirements between FSD 0929 and 1008 versions.

0929 FR table (T15): ID | Item | Description | System | PWC Comments | 第一期是否做
1008 FR table (T7):  ID | Item | Description | System | Fix Version
Match by Item + Description content similarity (IDs may have shifted).
"""
import json
import re
from difflib import SequenceMatcher

# ---------------------------------------------------------------- load grids
def load_tables(dump_path):
    tables = []
    cur = None
    with open(dump_path, encoding="utf-8") as f:
        for line in f:
            m = re.match(r"===== TABLE (\d+) \((\d+) rows x (\d+) cols\) =====", line)
            if m:
                cur = {"idx": int(m.group(1)), "rows": []}
                tables.append(cur)
                continue
            if cur is not None and line.strip():
                cur["rows"].append([c.replace("\\n", "\n") for c in line.rstrip("\n").split(" || ")])
    return tables


def get_fr_table(tables, want_cols):
    for t in tables:
        if not t["rows"]:
            continue
        hdr = [re.sub(r"\s+", " ", c).strip() for c in t["rows"][0]]
        if all(any(w in h for h in hdr) for w in want_cols):
            return t
    raise SystemExit(f"FR table with {want_cols} not found")


def norm(s):
    """Normalize text for fuzzy comparison: collapse whitespace, drop punctuation/case."""
    s = re.sub(r"\s+", " ", s).strip().lower()
    s = re.sub(r"[^\w一-鿿]+", "", s)
    return s


t0929 = load_tables("_tables_0929.txt")
t1008 = load_tables("_tables_1008.txt")

fr_a = get_fr_table(t0929, ["ID", "Item", "Description", "System", "Comments"])
fr_b = get_fr_table(t1008, ["ID", "Item", "Description", "System", "Fix Version"])


def to_records(t, ncol, phase_col=None):
    recs = []
    for r in t["rows"][1:]:
        if not re.match(r"FR-\d+", r[0].strip()):
            continue
        rec = {
            "id": r[0].strip(),
            "item": re.sub(r"\s+", " ", r[1]).strip(),
            "desc": r[2].strip(),
            "system": r[3].strip(),
            "last": re.sub(r"\s+", " ", r[ncol - 1]).strip(),
        }
        if phase_col:
            rec[phase_col] = r[5].strip() if len(r) > 5 else ""
        recs.append(rec)
    return recs


A = to_records(fr_a, 6, "phase1")
B = to_records(fr_b, 5)

print(f"0929 FR count: {len(A)} (max id {max(int(a['id'].split('-')[1]) for a in A)})")
print(f"1008 FR count: {len(B)} (max id {max(int(b['id'].split('-')[1]) for b in B)})")

# ---------------------------------------------------------------- matching
def sim(a, b):
    return SequenceMatcher(None, a, b).ratio()


matches = []  # (a_rec, b_rec, score)
used_b = set()
for a in A:
    na_i, na_d = norm(a["item"]), norm(a["desc"])
    best, best_score = None, 0.0
    for j, b in enumerate(B):
        if j in used_b:
            continue
        nb_i, nb_d = norm(b["item"]), norm(b["desc"])
        # item similarity dominates; desc provides confirmation
        s_item = sim(na_i, nb_i)
        s_desc = sim(na_d, nb_d)
        # desc texts can be long; also try prefix (first 300 chars)
        s_head = sim(na_d[:300], nb_d[:300])
        score = 0.5 * s_item + 0.3 * s_desc + 0.2 * s_head
        if score > best_score:
            best, best_score = j, score
    if best is not None and best_score >= 0.55:
        used_b.add(best)
        matches.append((a, B[best], round(best_score, 3)))
    else:
        matches.append((a, None, round(best_score, 3)))

unmatched_b = [B[j] for j in range(len(B)) if j not in used_b]

out = {
    "matched": [
        {
            "a_id": a["id"], "a_item": a["item"], "a_phase1": a["phase1"],
            "b_id": b["id"] if b else "", "b_item": b["item"] if b else "",
            "b_fix": b["last"] if b else "",
            "score": s,
        }
        for a, b, s in matches
    ],
    "unmatched_b": [
        {"b_id": b["id"], "b_item": b["item"], "b_fix": b["last"]} for b in unmatched_b
    ],
}
with open("_fr_compare.json", "w", encoding="utf-8") as f:
    json.dump({"A": A, "B": B, **out}, f, ensure_ascii=False, indent=1)

low = [m for m in out["matched"] if m["score"] < 0.75]
print(f"matched: {len(out['matched'])}, unmatched in 1008: {len(out['unmatched_b'])}, low-confidence(<0.75): {len(low)}")
for m in low:
    print(f"  LOW {m['a_id']} <-> {m['b_id'] or '?'} score={m['score']} | {m['a_item'][:40]}")
for u in out["unmatched_b"]:
    print(f"  1008-only {u['b_id']}: {u['b_item'][:60]} (Fix={u['b_fix']})")
