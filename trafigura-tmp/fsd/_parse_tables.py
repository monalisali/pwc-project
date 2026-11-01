"""Parse all tables from the FSD HTML files with rowspan/colspan support; emit a grid dump."""
import re
import sys
from html.parser import HTMLParser


class Table:
    def __init__(self):
        self.rows = []  # list of list of cell text

    def add(self, grid):
        self.rows.append(grid)


class TableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self.stack = []  # stack of Table-ish state
        self.cur = None
        self.row = None
        self.cell = None
        self.cell_attrs = None

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.stack.append(self.cur)
            self.cur = Table()
        elif tag == "tr" and self.cur is not None:
            self.row = []
        elif tag in ("td", "th") and self.row is not None:
            self.cell = []
            self.cell_attrs = dict(attrs)
        elif tag == "br" and self.cell is not None:
            self.cell.append("\n")

    def handle_endtag(self, tag):
        if tag == "table" and self.cur is not None:
            done = self.cur
            self.cur = self.stack.pop() if self.stack else None
            if done.rows:
                self.tables.append(done)
        elif tag == "tr" and self.row is not None:
            if self.cur is not None:
                self.cur.rows.append(self.row)
            self.row = None
        elif tag in ("td", "th") and self.cell is not None:
            text = "".join(self.cell)
            text = re.sub(r"[ \t\xa0]+", " ", text)
            text = re.sub(r" *\n *", "\n", text).strip()
            self.row.append({
                "text": text,
                "rowspan": int(self.cell_attrs.get("rowspan", 1) or 1),
                "colspan": int(self.cell_attrs.get("colspan", 1) or 1),
            })
            self.cell = None
            self.cell_attrs = None

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)


def build_grid(table):
    """Expand rowspans/colspans into a rectangular grid."""
    nrows = len(table.rows)
    grid = []
    pending = {}  # (row_idx) -> list of cells carried over
    for r, row in enumerate(table.rows):
        out = []
        c = 0
        for cell in row:
            while (r, c) in pending:
                out.append(pending.pop((r, c)))
                c += 1
            for dr in range(cell["rowspan"]):
                for dc in range(cell["colspan"]):
                    if dr == 0 and dc == 0:
                        out.append(cell["text"])
                    else:
                        pending[(r + dr, c + dc)] = cell["text"] if dr == 0 else ""
            c += cell["colspan"]
        # trailing pending cells
        while (r, c) in pending:
            out.append(pending.pop((r, c)))
            c += 1
        grid.append(out)
    # normalize width
    width = max(len(r) for r in grid)
    for r in grid:
        while len(r) < width:
            r.append("")
    return grid


def main(path, out):
    with open(path, encoding="utf-8") as f:
        html = f.read()
    p = TableParser()
    p.feed(html)
    with open(out, "w", encoding="utf-8") as f:
        for ti, t in enumerate(p.tables):
            grid = build_grid(t)
            f.write(f"===== TABLE {ti} ({len(grid)} rows x {len(grid[0]) if grid else 0} cols) =====\n")
            for row in grid:
                cells = [c.replace("\n", "\\n") for c in row]
                f.write(" || ".join(cells) + "\n")
            f.write("\n")
    print(f"{path}: {len(p.tables)} tables -> {out}")


main(sys.argv[1], sys.argv[2])
