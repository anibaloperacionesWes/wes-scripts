"""Lee el correo completo de Diego (UID 105)."""
import os
import imaplib
import email
from email.header import decode_header
from pathlib import Path

ROOT = Path(__file__).resolve().parent
USER = "agente.ia@wes.cl"


def _password() -> str:
    p = os.environ.get("WES_SMTP_PASSWORD", "").strip() or os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
    if p:
        return p.replace(" ", "")
    f = ROOT / "gmail_oauth" / "app_password.txt"
    if f.is_file():
        for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                return line.replace(" ", "")
    return ""


def _dec(s) -> str:
    if s is None:
        return ""
    if isinstance(s, bytes):
        return s.decode("utf-8", errors="replace")
    parts = decode_header(s)
    out = []
    for frag, enc in parts:
        out.append(frag.decode(enc or "utf-8", errors="replace") if isinstance(frag, bytes) else frag)
    return "".join(out)


def _body(msg) -> str:
    parts = []
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct in ("text/plain", "text/html"):
                pl = part.get_payload(decode=True)
                if pl:
                    parts.append(f"--- {ct} ---\n" + pl.decode("utf-8", errors="replace"))
    else:
        pl = msg.get_payload(decode=True)
        if pl:
            parts.append(pl.decode("utf-8", errors="replace"))
    return "\n\n".join(parts)


def main() -> None:
    mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    mail.login(USER, _password())
    mail.select("INBOX")
    typ, msg_data = mail.uid("fetch", b"105", "(RFC822)")
    raw = msg_data[0][1]
    msg = email.message_from_bytes(bytes(raw))
    print("ASUNTO:", _dec(msg.get("Subject")))
    print("DE:", _dec(msg.get("From")))
    print("FECHA:", msg.get("Date"))
    print("\n" + "=" * 70 + "\n")
    print(_body(msg))
    mail.logout()


if __name__ == "__main__":
    main()
