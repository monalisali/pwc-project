# -*- coding: utf-8 -*-
"""Extract tables from Word-exported HTML into structured text, dropping ignored columns."""
import re, sys
from bs4 import BeautifulSoup

IGNORE = ("PWC Comments", "第一期是否做")
# 0921: FR table has no header row; cols 4,5 = PWC Comments / 第一期是否做
#       NFR table (T18) & Questions table (T20): col 3 = PWC Comments
DROP = {'_cmp_0921.txt': {5: {4, 5}, 18: {3}, 20: {3}}, '_cmp_0929.txt': {}}

def norm(s):
    return re.sub(r'\s+', ' ', s.replace('\xa0', ' ').replace('\u3000', ' ')).strip()

def extract(path, out):
    soup = BeautifulSoup(open(path, 'rb').read(), 'html.parser')
    forced = DROP.get(out.split('\\')[-1].split('/')[-1], {})
    lines = []
    for ti, tbl in enumerate(soup.find_all('table')):
        rows = tbl.find_all('tr')
        if not rows:
            continue
        grid = []
        for tr in rows:
            cells = tr.find_all(['td', 'th'], recursive=False)  # direct children only
            grid.append([(norm(c.get_text(' ')), int(c.get('colspan', 1)) or 1) for c in cells])
        # detect ignore columns via header cells anywhere
        skip_idx = set()
        for row_cells in grid:
            idx = 0
            for text, span in row_cells:
                for k in range(span):
                    if any(key.lower() in text.lower() for key in IGNORE):
                        skip_idx.add(idx + k)
                    idx += 1
        skip_idx |= forced.get(ti, set())
        lines.append(f"===== TABLE {ti} =====")
        for ri, row_cells in enumerate(grid):
            idx = 0
            kept = []
            for text, span in row_cells:
                for k in range(span):
                    if idx + k not in skip_idx and text:
                        kept.append(text)
                idx += span
            if kept:
                lines.append(f"T{ti}R{ri}: " + " | ".join(kept))
    open(out, 'w', encoding='utf-8').write("\n".join(lines))
    print(f"{out}: done")

extract(sys.argv[1], sys.argv[2])
