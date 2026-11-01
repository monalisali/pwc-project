"""Re-extract text from the MHT-format .doc FSD files with correct UTF-16 decoding."""
import email
import html as htmllib
import re
import sys
from email import policy

FILES = [
    ("TMP+Output+VAT+API+Integration+with+Titan+SI_20260921.doc", "_output_vat_clean.txt"),
    ("FSD-TMP+Input+VAT+API+Integration+with+Titan+PI_20260921.doc", "_input_vat_clean.txt"),
]


def decode_payload(payload: bytes) -> str:
    """Try UTF-16 (charset 'unicode' in MHT usually means UTF-16LE), fall back to UTF-8."""
    if payload.startswith(b"\xff\xfe") or payload.startswith(b"\xfe\xff"):
        return payload.decode("utf-16")
    # Heuristic: if roughly every other byte is NUL, treat as UTF-16LE
    sample = payload[:2000]
    if sample and sample.count(b"\x00") > len(sample) * 0.4:
        return payload.decode("utf-16-le", errors="replace")
    return payload.decode("utf-8", errors="replace")


def mht_to_text(path: str) -> str:
    with open(path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    html_content = None
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            html_content = decode_payload(part.get_payload(decode=True))
            break
    if html_content is None:
        raise RuntimeError(f"no text/html part in {path}")

    html_content = re.sub(r"<(style|script)[^>]*>.*?</\1>", " ", html_content, flags=re.S | re.I)
    # Drop VML/word artifacts often left outside style tags
    html_content = re.sub(r"v\\:\*.*?;", " ", html_content)
    html_content = re.sub(r"</(td|th)>", " | ", html_content, flags=re.I)
    html_content = re.sub(r"</tr>", "\n", html_content, flags=re.I)
    html_content = re.sub(r"</(p|div|h[1-6]|li|table)>", "\n", html_content, flags=re.I)
    html_content = re.sub(r"<br\s*/?>", "\n", html_content, flags=re.I)
    text = re.sub(r"<[^>]+>", "", html_content)
    text = htmllib.unescape(text)
    text = text.replace("\x00", "")
    text = re.sub(r"[ \t\xa0]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text


for src, dst in FILES:
    text = mht_to_text(src)
    with open(dst, "w", encoding="utf-8") as f:
        f.write(text)
    bad = text.count("�")
    print(f"{src} -> {dst}: {len(text)} chars, {bad} replacement chars")
