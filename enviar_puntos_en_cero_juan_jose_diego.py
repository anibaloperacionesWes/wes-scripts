"""Envía el último reporte de puntos en cero a Juan, José y Diego.

Adjunta PDF si se puede convertir; si no, adjunta el DOCX.
"""

from __future__ import annotations

import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from generar_reporte_word import convertir_word_a_pdf


SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
SMTP_PASSWORD = os.environ.get("WES_SMTP_PASSWORD") or "vxbynfpoehbweelj"

TO_RECIPIENTS = ["juanlopez@wes.cl", "joseotarola@wes.cl", "diegocarrasco@wes.cl"]


def _latest_docx() -> Path:
    root = Path(__file__).resolve().parent
    out_dir = root / "reporte en cero"
    candidatos = [p for p in out_dir.glob("Reporte_Puntos_En_Cero_*.docx") if not p.name.startswith("~")]
    if not candidatos:
        raise FileNotFoundError(f"No se encontró Reporte_Puntos_En_Cero_*.docx en {out_dir}")
    candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0]


def main() -> int:
    docx_path = _latest_docx()
    pdf_path = convertir_word_a_pdf(docx_path)
    adj_path = pdf_path if pdf_path and pdf_path.exists() else docx_path
    subtype = "pdf" if adj_path.suffix.lower() == ".pdf" else "docx"

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = f"Reporte puntos en cero — {docx_path.stem.replace('Reporte_Puntos_En_Cero_', '')}"

    cuerpo = (
        "Estimados,\n\n"
        "Adjunto el reporte de puntos en cero solicitado.\n\n"
        f"Archivo: {adj_path.name}\n"
        "\nSaludos,\nAgente IA WES\n"
    )
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    with open(adj_path, "rb") as f:
        adj = MIMEApplication(f.read(), _subtype=subtype)
        adj.add_header("Content-Disposition", "attachment", filename=adj_path.name)
        msg.attach(adj)

    try:
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.sendmail(SMTP_USUARIO, TO_RECIPIENTS, msg.as_string())
        print(f"[OK] Enviado a: {', '.join(TO_RECIPIENTS)}")
        print(f"[OK] Adjunto: {adj_path}")
        return 0
    except Exception as e:
        print(f"[ERROR] Envío: {e}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

