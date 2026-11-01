"""Extract the HTML part of each MHT .doc to a standalone .html file."""
import email
from email import policy


def decode_payload(payload: bytes) -> str:
    if payload.startswith(b"\xff\xfe") or payload.startswith(b"\xfe\xff"):
        return payload.decode("utf-16")
    sample = payload[:2000]
    if sample and sample.count(b"\x00") > len(sample) * 0.4:
        return payload.decode("utf-16-le", errors="replace")
    return payload.decode("utf-8", errors="replace")


for src, dst in [
    ("FSD-TMP+Input+VAT+API+Integration+with+Titan+PI_20260929.doc", "_pi_0929.html"),
    ("FSD-TMP+Input+VAT+API+Integration+with+Titan+PI_20261008.doc", "_pi_1008.html"),
]:
    with open(src, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            html = decode_payload(part.get_payload(decode=True))
            with open(dst, "w", encoding="utf-8") as f:
                f.write(html)
            print(f"{src} -> {dst}: {len(html)} chars")
            break
