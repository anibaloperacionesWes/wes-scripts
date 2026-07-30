"""Envía por correo el PDF AGUNSA Intermodal (abril 2026) a Juan y Aníbal."""

from __future__ import annotations

import os
import smtplib
import sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent
PDF_PATH = (
    ROOT
    / "reports"
    / "AGUNSA"
    / "REPORTE"
    / "Intermodal-San_Antonio_20260507_1849"
    / "Reporte_AGUNSA_Intermodal_San_Antonio_20260401_20260430.pdf"
)

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
TO_RECIPIENTS = [
    "juanlopez@wes.cl",
    "anibal.aoperaciones@wes.cl",
]


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


def main() -> None:
    pw = _smtp_password()
    if not pw:
        print(
            "[ERROR] Falta contraseña SMTP (WES_GMAIL_APP_PASSWORD o gmail_oauth/app_password.txt).",
            file=sys.stderr,
        )
        raise SystemExit(1)

    if not PDF_PATH.is_file():
        print(f"[ERROR] No existe el PDF: {PDF_PATH}", file=sys.stderr)
        raise SystemExit(1)

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = "AGUNSA — Intermodal San Antonio — Reporte PDF — Abril 2026"
    msg.attach(
        MIMEText(
            "Estimados Juan y Aníbal,\n\n"
            "Adjunto el reporte (PDF) de AGUNSA — Intermodal San Antonio para el período 01/04/2026 al 30/04/2026.\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n",
            "plain",
            "utf-8",
        )
    )

    with open(PDF_PATH, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=PDF_PATH.name)
        msg.attach(part)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg)

    print(f"[OK] Correo enviado a: {', '.join(TO_RECIPIENTS)}")


if __name__ == "__main__":
    main()

