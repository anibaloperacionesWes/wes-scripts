"""
Envía por correo el informe Puente Alto (PDF + Word + consolidado CSV y Excel checkpoint)
a Aníbal, Juan, José y Diego.

  python enviar_informe_pa_pdf_word_csv.py

Opcional: ``WES_GMAIL_APP_PASSWORD`` o ``WES_SMTP_PASSWORD``; si faltan, usa el mismo
fallback que otros ``enviar_*.py`` de este repo.
"""

from __future__ import annotations

import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPORTS = ROOT / "reports" / "proyeccion ahorre puente 2025"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
SMTP_PASSWORD = (
    os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
    or os.environ.get("WES_SMTP_PASSWORD", "").strip()
    or os.environ.get("SMTP_PASSWORD", "").strip()
    or "vxbynfpoehbweelj"
)

TO_RECIPIENTS = [
    "anibal.aoperaciones@wes.cl",
    "juanlopez@wes.cl",
    "joseotarola@wes.cl",
    "diegocarrasco@wes.cl",
]

CONSOLIDADO_PREFIX = (
    "consolidado_m3_mensual_colegios_puente_alto_2025_20260427_1701_desde_checkpoint"
)
CONSOLIDADO_CSV = f"{CONSOLIDADO_PREFIX}.csv"
CONSOLIDADO_XLSX = f"{CONSOLIDADO_PREFIX}.xlsx"


def main() -> int:
    pdf = REPORTS / "Informe_PA_generado_compaginacion.pdf"
    docx = REPORTS / "Informe_PA_generado_compaginacion.docx"
    csv_path = REPORTS / CONSOLIDADO_CSV
    xlsx_path = REPORTS / CONSOLIDADO_XLSX

    for p in (pdf, docx, csv_path, xlsx_path):
        if not p.is_file():
            print(f"[ERROR] Falta archivo: {p}")
            return 1

    if not SMTP_PASSWORD:
        print(
            "[ERROR] Defina la variable de entorno WES_SMTP_PASSWORD "
            "(contraseña de aplicación Gmail para agente.ia@wes.cl)."
        )
        return 1

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = (
        "Informe Puente Alto — PDF, Word y consolidado (CSV + Excel checkpoint)"
    )

    body = (
        "Estimados Aníbal, Juan, José y Diego,\n\n"
        "Adjunto:\n"
        f"  • {pdf.name}\n"
        f"  • {docx.name}\n"
        f"  • {csv_path.name} (consolidado mensual, checkpoint)\n"
        f"  • {xlsx_path.name} (consolidado mensual, checkpoint)\n\n"
        "Saludos,\nAgente IA WES\n"
    )
    msg.attach(MIMEText(body, "plain", "utf-8"))

    def _attach(path: Path) -> None:
        with open(path, "rb") as f:
            data = f.read()
        ext = path.suffix.lower()
        subtype = "pdf" if ext == ".pdf" else "octet-stream"
        part = MIMEApplication(data, _subtype=subtype)
        part.add_header("Content-Disposition", "attachment", filename=path.name)
        msg.attach(part)

    _attach(pdf)
    _attach(docx)
    _attach(csv_path)
    _attach(xlsx_path)

    try:
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.sendmail(SMTP_USUARIO, TO_RECIPIENTS, msg.as_string())
    except Exception as e:
        print(f"[ERROR] Envío: {e}")
        return 1

    print("[OK] Enviado a:", ", ".join(TO_RECIPIENTS))
    print(
        "     Adjuntos:",
        f"{pdf.name}, {docx.name}, {csv_path.name}, {xlsx_path.name}",
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
