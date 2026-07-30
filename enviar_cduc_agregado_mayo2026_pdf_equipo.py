"""
PDF agregado CDUC mayo 2026 → Juan, Diego, Aníbal; BCC José (mismos destinatarios CORMUP/PA).
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
DOCX_GLOB = "Reporte_Agregado_CDUC_20260501_20260531.docx"

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
        for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                return line.replace(" ", "").strip()
    return ""


def _ultimo_docx() -> Path:
    base = ROOT / "reports" / "CDUC" / "ABREGADO"
    candidatos = list(base.rglob(DOCX_GLOB)) if base.is_dir() else []
    if not candidatos:
        raise FileNotFoundError(f"No se encontró {DOCX_GLOB} bajo {base}")
    return max(candidatos, key=lambda p: p.stat().st_mtime)


def main() -> int:
    pw = _smtp_password()
    if not pw:
        print("[ERROR] Falta contraseña SMTP.", file=sys.stderr)
        return 1

    docx = _ultimo_docx()
    print(f"[OK] Word: {docx}")

    from control_nocturno import convertir_docx_a_pdf

    pdf = convertir_docx_a_pdf(docx)
    if not pdf or not pdf.is_file():
        print("[ERROR] No se pudo convertir a PDF.", file=sys.stderr)
        return 1
    print(f"[OK] PDF: {pdf}")

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_VISIBLE)
    msg["Subject"] = "CDUC — Reporte agregado PDF — Mayo 2026"
    msg.attach(
        MIMEText(
            "Estimados Juan, Diego y Aníbal,\n\n"
            "Adjunto el reporte agregado (PDF) de CDUC para el periodo 01/05/2026 al 31/05/2026.\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n",
            "plain",
            "utf-8",
        )
    )
    with open(pdf, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=pdf.name)
        msg.attach(part)

    envelope = list(TO_VISIBLE) + list(BCC_ENVELOPE_ONLY)
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=envelope)

    print(f"[OK] Enviado a: {', '.join(TO_VISIBLE)}")
    print(f"[OK] Copia oculta: {', '.join(BCC_ENVELOPE_ONLY)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
