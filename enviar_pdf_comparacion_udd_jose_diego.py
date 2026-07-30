"""Envía el PDF de comparación UDD a José y Diego."""

from __future__ import annotations

import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
SMTP_PASSWORD = (
    os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
    or os.environ.get("WES_SMTP_PASSWORD", "").strip()
    or "vxbynfpoehbweelj"
).replace(" ", "")

TO_RECIPIENTS = ["joseotarola@wes.cl", "diegocarrasco@wes.cl", "juanlopez@wes.cl"]

PDF_PATH = (
    Path(__file__).resolve().parent
    / "reports"
    / "comparacion_udd_cuentas_vs_wes"
    / "Comparacion_cuentas_vs_WES_Honduras_UDD.pdf"
)


def main() -> int:
    if not PDF_PATH.exists():
        print(f"[ERROR] No existe PDF: {PDF_PATH}")
        return 1

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = "Comparación cuentas vs WES — UDD Honduras"
    msg.attach(
        MIMEText(
            (
                "Estimados José, Diego y Juan,\n\n"
                "Adjunto el informe en PDF de comparación entre cuentas de agua y registro WES "
                "para UDD Honduras.\n\n"
                "Saludos,\nSistema WES\n"
            ),
            "plain",
            "utf-8",
        )
    )

    with open(PDF_PATH, "rb") as f:
        adj = MIMEApplication(f.read(), _subtype="pdf")
        adj.add_header("Content-Disposition", "attachment", filename=PDF_PATH.name)
        msg.attach(adj)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, SMTP_PASSWORD)
        server.sendmail(SMTP_USUARIO, TO_RECIPIENTS, msg.as_string())

    print(f"[OK] Enviado a: {', '.join(TO_RECIPIENTS)}")
    print(f"[OK] Adjunto: {PDF_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
