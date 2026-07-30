"""Envía el Word de respaldo facturación Juan Pablo II (Las Condes) a Juan y Aníbal."""

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
DOCX_PATH = (
    ROOT
    / "reports"
    / "Las_Condes"
    / "Facturacion"
    / "Juan_Pablo_II"
    / "20260401_20260412"
    / "Respaldo_Facturacion_Juan_Pablo_II_20260401_20260412.docx"
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

    if not DOCX_PATH.is_file():
        print(f"[ERROR] No existe el Word: {DOCX_PATH}", file=sys.stderr)
        raise SystemExit(1)

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = (
        "Las Condes — Juan Pablo II — Respaldo facturación WES — 01/04/2026 al 12/04/2026"
    )
    msg.attach(
        MIMEText(
            "Estimados Juan y Aníbal,\n\n"
            "Adjunto el informe agregado WES (Word) del Colegio Juan Pablo II (Las Condes), "
            "nodo 000022-01, como respaldo para el cobro de la última factura.\n\n"
            "Periodo de consumo en aplicación: 01/04/2026 al 12/04/2026.\n"
            "Consumo total registrado en app WES: 223,7 m³ (11 días con registro).\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n",
            "plain",
            "utf-8",
        )
    )

    with open(DOCX_PATH, "rb") as f:
        part = MIMEApplication(
            f.read(),
            _subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        part.add_header("Content-Disposition", "attachment", filename=DOCX_PATH.name)
        msg.attach(part)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg)

    print(f"[OK] Correo enviado a: {', '.join(TO_RECIPIENTS)}")
    print(f"     Adjunto: {DOCX_PATH.name}")


if __name__ == "__main__":
    main()
