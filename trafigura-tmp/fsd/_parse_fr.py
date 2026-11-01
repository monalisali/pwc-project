"""Parse the Functional Requirement tables from _pi_0929.txt and _pi_1008.txt into structured rows."""
import json
import re


def load_rows(path):
    with open(path, encoding="utf-8") as f:
        text = f.read()
    rows = []
    for chunk in text.split("@ROW@")[1:]:
        # collapse whitespace, then split cells on the pipe markers
        chunk = re.sub(r"\s+", " ", chunk)
        cells = [c.strip() for c in chunk.split("|")]
        cells = [c for c in cells if c != ""]
        rows.append(cells)
    return rows


def find_fr_table(rows):
    """Return index of the row whose first cell == 'ID' and contains 'Item'."""
    for i, r in enumerate(rows):
        if r and r[0] == "ID" and any("Item" in c for c in r):
            return i
    return -1


def parse_fr(path, label):
    rows = load_rows(path)
    start = find_fr_table(rows)
    assert start >= 0, f"FR header row not found in {path}"
    header = rows[start]
    print(f"[{label}] header: {header}")
    frs = []
    for r in rows[start + 1:]:
        if not r:
            continue
        if not re.match(r"FR-\d+", r[0]):
            # end of table (next section header etc.)
            if re.match(r"^\d+\.\s", r[0]) or r[0].startswith("Non-functional"):
                break
            continue
        frs.append(r)
    print(f"[{label}] {len(frs)} FR rows; cell counts: {sorted(set(len(r) for r in frs))}")
    return frs


fr0929 = parse_fr("_pi_0929.txt", "0929")
fr1008 = parse_fr("_pi_1008.txt", "1008")

with open("_fr_0929.json", "w", encoding="utf-8") as f:
    json.dump(fr0929, f, ensure_ascii=False, indent=1)
with open("_fr_1008.json", "w", encoding="utf-8") as f:
    json.dump(fr1008, f, ensure_ascii=False, indent=1)

print("--- 0929 sample (first 2) ---")
for r in fr0929[:2]:
    print([c[:120] for c in r])
print("--- 1008 sample (first 2) ---")
for r in fr1008[:2]:
    print([c[:120] for c in r])
