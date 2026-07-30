"""Busca correos de Diego en inbox agente.ia@wes.cl."""
import os
import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
USER = "agente.ia@wes.cl"
DIEGO = "diegocarrasco@wes.cl"


def _password() -> str:
    p = (
        os.environ.get("WES_SMTP_PASSWORD", "").strip()
        or os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
    )
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
        if isinstance(frag, bytes):
            out.append(frag.decode(enc or "utf-8", errors="replace"))
        else:
            out.append(frag)
    return "".join(out)


def _body_snippet(msg, limit: int = 350) -> str:
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                pl = part.get_payload(decode=True)
                if pl:
                    return pl.decode("utf-8", errors="replace")[:limit].strip()
        return ""
    pl = msg.get_payload(decode=True)
    return pl.decode("utf-8", errors="replace")[:limit].strip() if pl else ""


def main() -> None:
    pw = _password()
    if not pw:
        print("[ERROR] Sin contraseña IMAP")
        return

    mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    mail.login(USER, pw)
    mail.select("INBOX")

    queries = [
        f'(FROM "{DIEGO}")',
        f'(FROM "{DIEGO}" UNSEEN)',
    ]
    seen_uids: set[bytes] = set()
    rows: list[tuple] = []

    for q in queries:
        st, data = mail.uid("search", None, q)
        if st != "OK" or not data[0]:
            continue
        uids = sorted(data[0].split(), key=lambda u: int(u), reverse=True)[:20]
        for uid in uids:
            if uid in seen_uids:
                continue
            seen_uids.add(uid)
            typ, msg_data = mail.uid("fetch", uid, "(RFC822 FLAGS)")
            if typ != "OK" or not msg_data or not msg_data[0]:
                continue
            raw = msg_data[0][1] if isinstance(msg_data[0], tuple) else msg_data[0]
            msg = email.message_from_bytes(bytes(raw))
            meta = b""
            if isinstance(msg_data[0], tuple):
                meta = msg_data[0][0] if isinstance(msg_data[0][0], bytes) else str(msg_data[0][0]).encode()
            unseen = b"\\Seen" not in meta
            try:
                dt = parsedate_to_datetime(msg.get("Date"))
            except Exception:
                dt = None
            rows.append(
                (
                    uid,
                    dt,
                    unseen,
                    _dec(msg.get("Subject")),
                    _dec(msg.get("From")),
                    _body_snippet(msg),
                )
            )

    mail.logout()

    rows.sort(key=lambda r: r[1] or __import__("datetime").datetime.min.replace(tzinfo=__import__("datetime").timezone.utc), reverse=True)

    print("=== Correos de Diego en agente.ia@wes.cl (INBOX) ===")
    print(f"Total: {len(rows)}")
    if not rows:
        print("No se encontraron correos de diegocarrasco@wes.cl.")
        return
    for uid, dt, unseen, subj, frm, body in rows:
        uid_s = uid.decode() if isinstance(uid, bytes) else str(uid)
        fecha = dt.strftime("%d-%m-%Y %H:%M") if dt else "?"
        print("-" * 60)
        print(f"UID {uid_s} | {fecha} | {'NO LEÍDO' if unseen else 'leído'}")
        print(f"Asunto: {subj}")
        print(f"De: {frm}")
        if body:
            print(f"Vista previa: {body.replace(chr(10), ' ')[:280]}")


if __name__ == "__main__":
    main()
