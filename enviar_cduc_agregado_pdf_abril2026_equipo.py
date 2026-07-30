"""Convierte a PDF el agregado CDUC abril 2026 (más reciente) y lo envía a Aníbal, Juan y Diego."""

from __future__ import annotations

import os
import smtplib
import sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from generar_reporte_word import convertir_word_a_pdf

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent
ABREGADO = ROOT / "reports" / "CDUC" / "ABREGADO"
AGG_PATTERN = "Reporte_Agregado_CDUC_20260401_20260430.docx"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
TO_RECIPIENTS = [
    "anibal.aoperaciones@wes.cl",
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
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


def _latest_docx() -> Path | None:
    if not ABREGADO.is_dir():
        return None
    cands = [p for p in ABREGADO.rglob(AGG_PATTERN) if p.is_file()]
    if not cands:
        return None
    return max(cands, key=lambda p: p.stat().st_mtime)


def main() -> None:
    pw = _smtp_password()
    if not pw:
        print(
            "[ERROR] Falta contraseña SMTP (WES_GMAIL_APP_PASSWORD o gmail_oauth/app_password.txt).",
            file=sys.stderr,
        )
        raise SystemExit(1)

    docx = _latest_docx()
    if not docx:
        print(f"[ERROR] No hay {AGG_PATTERN} en {ABREGADO}", file=sys.stderr)
        raise SystemExit(1)

    print(f"[INFO] Word agregado: {docx}")
    pdf_path = convertir_word_a_pdf(docx)
    if not pdf_path or not Path(pdf_path).is_file():
        print("[ERROR] No se pudo generar el PDF.", file=sys.stderr)
        raise SystemExit(1)
    pdf_path = Path(pdf_path)
    print(f"[INFO] PDF: {pdf_path}")

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = "CDUC — Reporte agregado PDF — Abril 2026"
    msg.attach(
        MIMEText(
            "Estimados Aníbal, Juan y Diego,\n\n"
            "Adjunto el reporte agregado (PDF) de CDUC para el período del 01/04/2026 al 30/04/2026.\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n",
            "plain",
            "utf-8",
        )
    )
    with open(pdf_path, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=pdf_path.name)
        msg.attach(part)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg)

    print(f"[OK] Correo enviado a: {', '.join(TO_RECIPIENTS)}")


if __name__ == "__main__":
    main()
