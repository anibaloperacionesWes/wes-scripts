"""
Envía informe Genchi CDP Puente Alto (Word + PDF) a Juan, Diego y Aníbal.
José recibe copia oculta (solo en sobre SMTP, sin cabecera Bcc en el mensaje).

  python enviar_informe_genchi_cdp_puente_alto_juan_diego_anibal.py
"""

from __future__ import annotations

import os
import smtplib
import sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "reports" / "Genchi" / "CDP Puente Alto" / "informe"
TS = "20260605_1427"
STEM = f"Informe_Genchi_CDP_Puente_Alto_facturaciones_{TS}"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_VISIBLE = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
    "anibal.aoperaciones@wes.cl",
]
BCC_ENVELOPE_ONLY = ["joseotarola@wes.cl"]


def _smtp_password() -> str:
    p = (
        os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
        or os.environ.get("WES_SMTP_PASSWORD", "").strip()
    )
    if p:
        return p.replace(" ", "").strip()
    f = ROOT / "gmail_oauth" / "app_password.txt"
    if f.is_file():
        try:
            for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
                line = line.strip()
                if line and not line.startswith("#"):
                    return line.replace(" ", "").strip()
        except OSError:
            pass
    return ""


def main() -> int:
    pw = _smtp_password()
    if not pw:
        print(
            "[ERROR] Falta contraseña SMTP (WES_GMAIL_APP_PASSWORD, WES_SMTP_PASSWORD o gmail_oauth/app_password.txt).",
            file=sys.stderr,
        )
        return 1

    docx = OUT_DIR / f"{STEM}.docx"
    pdf = OUT_DIR / f"{STEM}.pdf"
    for p in (docx, pdf):
        if not p.is_file():
            print(f"[ERROR] Falta archivo: {p}", file=sys.stderr)
            return 1

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_VISIBLE)
    msg["Subject"] = (
        "Gendarmería Genchi — CDP Puente Alto — Informe comparativo facturación Aguas Andinas (Word + PDF)"
    )
    msg.attach(
        MIMEText(
            "Estimados Juan, Diego y Aníbal,\n\n"
            "Adjunto el informe comparativo de facturación Aguas Andinas para la unidad "
            "Gendarmería CDP Detención Preventiva Puente Alto (cuenta 1008941-7), "
            "con análisis con y sin monitoreo WES (Word y PDF).\n\n"
            "Continuaremos con las demás unidades carcelarias que tengan facturación de Aguas Andinas.\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n",
            "plain",
            "utf-8",
        )
    )
    for p, subtype in (
        (docx, "vnd.openxmlformats-officedocument.wordprocessingml.document"),
        (pdf, "pdf"),
    ):
        with open(p, "rb") as f:
            part = MIMEApplication(f.read(), _subtype=subtype)
            part.add_header("Content-Disposition", "attachment", filename=p.name)
            msg.attach(part)
        print(f"[OK] Adjunto: {p.name}")

    envelope = list(TO_VISIBLE) + list(BCC_ENVELOPE_ONLY)
    print("[INFO] Enviando correo…")
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=envelope)

    print("[OK] Enviado.")
    print(f"  Para: {', '.join(TO_VISIBLE)}")
    print(f"  Copia oculta (sobre): {', '.join(BCC_ENVELOPE_ONLY)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
