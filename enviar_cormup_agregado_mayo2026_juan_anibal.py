"""Envía agregado CORMUP mayo 2026 (PDF) a Juan con copia a Aníbal."""

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
PERIODO_SUFIJO = "20260501_20260531"
DOCX_GLOB = f"Reporte_Agregado_CORMUP_{PERIODO_SUFIJO}.docx"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
TO_RECIPIENTS = ["juanlopez@wes.cl"]
CC_RECIPIENTS = ["anibal.aoperaciones@wes.cl"]


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
    base = ROOT / "reports" / "CORMUP" / "ABREGADO"
    candidatos = list(base.rglob(DOCX_GLOB)) if base.is_dir() else []
    if not candidatos:
        raise FileNotFoundError(f"No se encontró {DOCX_GLOB} bajo {base}")
    return max(candidatos, key=lambda p: p.stat().st_mtime)


def _asegurar_pdf() -> Path:
    pdf_candidatos = list((ROOT / "reports" / "CORMUP" / "ABREGADO").rglob(
        f"Reporte_Agregado_CORMUP_{PERIODO_SUFIJO}.pdf"
    ))
    if pdf_candidatos:
        pdf = max(pdf_candidatos, key=lambda p: p.stat().st_mtime)
        if pdf.stat().st_size > 10_000:
            return pdf
    docx = _ultimo_docx()
    from control_nocturno import convertir_docx_a_pdf

    pdf = convertir_docx_a_pdf(docx)
    if not pdf or not pdf.is_file():
        raise RuntimeError("No se pudo generar el PDF")
    return pdf


def main() -> int:
    pw = _smtp_password()
    if not pw:
        print("[ERROR] Falta contraseña SMTP.", file=sys.stderr)
        return 1

    pdf = _asegurar_pdf()
    print(f"[OK] PDF: {pdf} ({pdf.stat().st_size // 1024} KB)")

    cuerpo = (
        "Estimado Juan,\n\n"
        "Con copia a Aníbal.\n\n"
        "Adjunto el reporte agregado (PDF) de los colegios CORMUP / Peñalolén "
        "para el periodo 01/05/2026 al 31/05/2026.\n\n"
        "Sobre el consumo nocturno del informe: la diferencia respecto a versiones "
        "anteriores se debe a que antes el cálculo tomaba como inicio del consumo "
        "nocturno desde las 20:00 horas, y en esta versión el consumo nocturno se "
        "calcula únicamente sumando el consumo entre las 00:00 y las 07:00 horas "
        "(según el CSV horario de cada día), por lo que los valores son menores y "
        "más acordes a lo que se puede validar en la app.\n\n"
        "Saludos cordiales,\n"
        "Sistema WES\n"
    )

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Cc"] = ", ".join(CC_RECIPIENTS)
    msg["Subject"] = "CORMUP Peñalolén — Reporte agregado PDF — Mayo 2026 (reenvío)"
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    with open(pdf, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=pdf.name)
        msg.attach(part)

    envelope = TO_RECIPIENTS + CC_RECIPIENTS
    try:
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO, timeout=60) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(SMTP_USUARIO, pw)
            refused = server.sendmail(SMTP_USUARIO, envelope, msg.as_string())
            if refused:
                print(f"[ERROR] Destinatarios rechazados: {refused}", file=sys.stderr)
                return 1
    except Exception as exc:
        print(f"[ERROR] Envío SMTP: {exc}", file=sys.stderr)
        return 1

    print(f"[OK] Enviado a: {', '.join(TO_RECIPIENTS)}")
    print(f"[OK] CC: {', '.join(CC_RECIPIENTS)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
