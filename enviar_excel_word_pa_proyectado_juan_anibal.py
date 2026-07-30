from __future__ import annotations

import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPORTS = ROOT / "reports" / "proyeccion ahorre puente 2025"

SMTP_USER = "agente.ia@wes.cl"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_PASSWORD = (
    os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
    or os.environ.get("WES_SMTP_PASSWORD", "").strip()
    or os.environ.get("SMTP_PASSWORD", "").strip()
    or "vxbynfpoehbweelj"
)

TO_RECIPIENTS = [
    "anibal.aoperaciones@wes.cl",
    "juanlopez@wes.cl",
]

XLSX_NAME = "consolidado_m3_mensual_colegios_puente_alto_2025_20260428_1224_desde_checkpoint.xlsx"
DOCX_NAME = "Informe_PA_generado_compaginacion_proyectado_30dias.docx"


def attach_file(msg: MIMEMultipart, path: Path) -> None:
    with open(path, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="octet-stream")
    part.add_header("Content-Disposition", "attachment", filename=path.name)
    msg.attach(part)


def main() -> int:
    xlsx = REPORTS / XLSX_NAME
    docx = REPORTS / DOCX_NAME
    for p in (xlsx, docx):
        if not p.is_file():
            print(f"[ERROR] Falta archivo: {p}")
            return 1

    msg = MIMEMultipart()
    msg["From"] = SMTP_USER
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = "Puente Alto - Excel y Word proyectados a 30 dias"
    body = (
        "Estimados Anibal y Juan,\n\n"
        "Adjunto el ultimo Excel consolidado y el Word del informe con la proyeccion a 30 dias.\n\n"
        "Resumen de lo realizado:\n"
        "- Sheet1: se mantiene igual (base original de consumos medidos).\n"
        "- Sheet2: se proyecta por punto/mes a 30 dias usando promedio diario con datos:\n"
        "  valor proyectado = (m3 del mes / dias con dato) * 30.\n"
        "- En Sheet2, cuando aplica proyeccion, la columna de dias del mes queda en 30.\n"
        "- Sheet3: totalizados por periodo comparando hoja base (Sheet1) vs hoja proyectada (Sheet2)\n"
        "  para Con WES, Sin WES y Ahorro.\n"
        "- El Word adjunto fue regenerado usando los consumos proyectados (Sheet2),\n"
        "  manteniendo misma estructura de tablas, graficas y secciones.\n\n"
        "Saludos,\n"
        "Agente IA WES\n"
    )
    msg.attach(MIMEText(body, "plain", "utf-8"))
    attach_file(msg, xlsx)
    attach_file(msg, docx)

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, TO_RECIPIENTS, msg.as_string())
    except Exception as e:
        print(f"[ERROR] Envio: {e}")
        return 1

    print("[OK] Enviado a:", ", ".join(TO_RECIPIENTS))
    print(f"     Adjuntos: {xlsx.name}, {docx.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
