"""
Envía por correo el informe de comparación ICCO Renca (DOCX + PDF)
a Aníbal y Diego.

Uso:
  python enviar_comparacion_icco_renca_anibal_diego.py
"""

from __future__ import annotations

import os
import smtplib
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from config_correos_equipo import obtener_correos_por_rol


ROOT = Path(__file__).resolve().parent
OUT_DIR = (
    ROOT
    / "reports"
    / "Renca"
    / "Coparacion App con Aguas Andinas"
    / "reporte_comparacion_Icco"
)
DOCX_PATH = OUT_DIR / "Comparacion_App_vs_Facturaciones_ICCO_Renca.docx"
PDF_PATH = OUT_DIR / "Comparacion_App_vs_Facturaciones_ICCO_Renca.pdf"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587


def _leer_smtp_password() -> str:
    pw = (
        os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
        or os.environ.get("WES_SMTP_PASSWORD", "").strip()
        or os.environ.get("SMTP_PASSWORD", "").strip()
    )
    if pw:
        return pw.replace(" ", "").strip()
    # Fallback local (ignorado por git en este repo)
    p = ROOT / "gmail_oauth" / "app_password.txt"
    if p.exists():
        try:
            v = p.read_text(encoding="utf-8").strip()
            return v.replace(" ", "").strip()
        except Exception:
            return ""
    return ""


def main() -> None:
    if not DOCX_PATH.exists():
        raise SystemExit(f"[ERROR] No existe DOCX: {DOCX_PATH}")
    if not PDF_PATH.exists():
        raise SystemExit(f"[ERROR] No existe PDF: {PDF_PATH}")

    destinatarios = obtener_correos_por_rol("anibal", "diego")
    if not destinatarios:
        raise SystemExit("[ERROR] No se encontraron correos para anibal/diego en config_correos_equipo.py")

    pw = _leer_smtp_password()
    if not pw:
        raise SystemExit(
            "[ERROR] Falta contraseña SMTP. Define WES_GMAIL_APP_PASSWORD o WES_SMTP_PASSWORD "
            "o crea gmail_oauth/app_password.txt"
        )

    msg = MIMEMultipart()
    msg["From"] = f"Agente WES <{SMTP_USUARIO}>"
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = f"ICCO Renca — Comparación App vs Facturación — {datetime.now():%d-%m-%Y}"

    cuerpo = (
        "Equipo,\n\n"
        "Adjunto el informe de comparación App WES vs Facturación Aguas Andinas (ICCO Renca).\n\n"
        f"- Word: {DOCX_PATH.name}\n"
        f"- PDF: {PDF_PATH.name}\n\n"
        "Saludos,\n"
        "Agente WES\n"
    )
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for pth, subtype in [(DOCX_PATH, "vnd.openxmlformats-officedocument.wordprocessingml.document"), (PDF_PATH, "pdf")]:
        with open(pth, "rb") as f:
            adj = MIMEApplication(f.read(), _subtype=subtype)
        adj.add_header("Content-Disposition", "attachment", filename=pth.name)
        msg.attach(adj)

    print(f"[INFO] Enviando a: {', '.join(destinatarios)}")
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=destinatarios)
    print("[OK] Correo enviado.")


if __name__ == "__main__":
    main()

