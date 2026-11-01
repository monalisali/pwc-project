"""Parse all tables from the two PI FSD MHT .doc files into JSON (nested-table aware)."""
import email
import json
import re
from email import policy
from html.parser import HTMLParser


def decode_payload(payload: bytes) -> str:
    if payload.startswith((b"\xff\xfe", b"\xfe\xff")):
        return payload.decode("utf-16")
    sample = payload[:2000]
    if sample and sample.count(b"\x00") > len(sample) * 0.4:
        return payload.decode("utf-16-le", errors="replace")
    return payload.decode("utf-8", errors="replace")


def get_html(path):
    with open(path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            return decode_payload(part.get_payload(decode=True))
    raise RuntimeError(f"no html in {path}")


def norm(text):
    text = re.sub(r"[ \t\xa0]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    return text.strip()


class TableExtractor(HTMLParser):
    """Stack-based table parser. Nested tables are flattened into the parent
    cell as 'a | b | c' lines."""

    def __init__(self):
        super().__init__()
        self.tables = []          # completed top-level tables
        self.stack = []           # active table contexts (innermost last)
        self.cell = None          # list of text chunks for current cell (innermost)
        self.buf = []

    # -- helpers -----------------------------------------------------
    @property
    def top(self):
        return self.stack[-1] if self.stack else None

    def cellbuf(self):
        # the buffer list belonging to the innermost open cell
        for ctx in reversed(self.stack):
            if ctx.get("cell") is not None:
                return ctx["cell"]
        return self.buf if self.cell is None else self.cell

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            ctx = {"rows": [], "row": None, "cell": None}
            self.stack.append(ctx)
        elif tag == "tr":
            if self.stack:
                self.top["row"] = []
        elif tag in ("td", "th"):
            if self.stack and self.top["row"] is not None:
                self.top["cell"] = []
        elif tag == "br":
            self._append("\n")

    def _append(self, text):
        target = self.buf if self.cell is None else self.cell
        target.append(text)

    def handle_endtag(self, tag):
        if tag == "table":
            if not self.stack:
                return
            ctx = self.stack.pop()
            tbl = [r for r in ctx["rows"] if any(c.strip() for c in r)]
            if not self.stack:
                if tbl:
                    self.tables.append(tbl)
            else:
                # nested table: flatten into parent cell text
                parent = self.stack[-1]
                lines = []
                for row in tbl:
                    lines.append(" | ".join(c.replace("\n", " ") for c in row if c.strip()))
                parent_cell = parent.get("cell")
                if parent_cell is not None:
                    parent_cell.extend(["\n[表] " + ln + "\n" for ln in lines if ln])
                # re-open parent cell so following text continues
                if parent_cell is not None:
                    self.cell = parent_cell
                    self.buf = parent_cell
                    return
                self.cell = None
                self.buf = []
        elif tag == "tr":
            if not self.stack:
                return
            row = self.top["row"]
            if row is not None:
                self._close_cell()
                if row and any(c.strip() for c in row):
                    self.top["rows"].append(row)
            self.top["row"] = None
        elif tag in ("td", "th"):
            self._close_cell()

    def _close_cell(self):
        if not self.stack:
            return
        ctx = self.top
        cell = ctx.get("cell")
        if cell is None:
            return
        text = norm("".join(cell))
        if ctx["row"] is not None:
            ctx["row"].append(text)
        ctx["cell"] = None
        self.cell = None
        self.buf = []

    def handle_data(self, data):
        if self.stack and self.top.get("cell") is not None:
            self.top["cell"].append(data)


FILES = {
    "0921": "FSD-TMP+Input+VAT+API+Integration+with+Titan+PI_20260921.doc",
    "0929": "FSD-TMP+Input+VAT+API+Integration+with+Titan+PI_20260929.doc",
}

result = {}
for key, fname in FILES.items():
    p = TableExtractor()
    p.feed(get_html(fname))
    tables = p.tables
    result[key] = tables
    print(f"{key}: {len(tables)} non-empty tables")
    for i, t in enumerate(tables):
        header = " | ".join(c.replace("\n", " ")[:25] for c in t[0])
        print(f"  [{i}] rows={len(t)}: {header[:140]}")

with open("_pi_tables.json", "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=1)
