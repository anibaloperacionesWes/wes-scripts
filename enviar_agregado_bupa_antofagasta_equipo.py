"""
Convierte a PDF y envía el agregado Bupa Antofagasta (23–27/07/2026)
a Juan, Diego y Aníbal.
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

from generar_reporte_word import convertir_word_a_pdf

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "reports" / "Bupa_Antofagasta" / "ABREGADO" / "AGREGADO_20260728_1950"
DOCX_PATH = OUT / "Reporte_Agregado_BUPA_20260723_20260727.docx"
PDF_PATH = OUT / "Reporte_Agregado_BUPA_20260723_20260727.pdf"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_RECIPIENTS = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
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
        for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                return line.replace(" ", "").strip()
    return ""


def main() -> int:
    if not DOCX_PATH.is_file():
        print(f"[ERROR] No existe: {DOCX_PATH}", file=sys.stderr)
        return 1

    print(f"[INFO] Word: {DOCX_PATH} ({DOCX_PATH.stat().st_size // 1024} KB)")
    print("[INFO] Convirtiendo a PDF...")
    pdf = convertir_word_a_pdf(DOCX_PATH)
    if not pdf or not pdf.is_file():
        print("[ERROR] No se pudo generar el PDF.", file=sys.stderr)
        return 1
    print(f"[OK] PDF: {pdf} ({pdf.stat().st_size // 1024} KB)")

    pw = _smtp_password()
    if not pw:
        print("[ERROR] Falta contraseña SMTP.", file=sys.stderr)
        return 1

    cuerpo = """Estimados Juan, Diego y Aníbal,

Adjunto el reporte agregado de Clínica Bupa Antofagasta (nodos 000029-07, 000029-08, 000029-09 y 000029-10).

Resumen
-------
• Periodo: 23/07/2026 – 27/07/2026 (5 días civiles completos).
• Proyección mensual = promedio diario WES × 30 (sanitaria incluida).
• Factura julio (6.696 m³ / $18.538.860) solo como referencia histórica y tarifa (~$2.769 CLP/m³).
• Sanitaria (5 días): 1.249,8 m³ → ~250 m³/día → proyección ~7.498,7 m³ (~$20,8 M).
• Salas de bomba proyectadas: ~19,1% de la cuenta; no monitoreado: ~80,9%.

Adjunto: PDF (y Word).

Saludos,
Agente IA WES
"""

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = (
        "Bupa Antofagasta — Reporte agregado 23–27/07/2026 (proyección WES + participación)"
    )
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for path, subtype in (
        (pdf, "pdf"),
        (DOCX_PATH, "vnd.openxmlformats-officedocument.wordprocessingml.document"),
    ):
        with open(path, "rb") as f:
            part = MIMEApplication(f.read(), _subtype=subtype)
            part.add_header("Content-Disposition", "attachment", filename=path.name)
            msg.attach(part)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg)

    print(f"[OK] Enviado a: {', '.join(TO_RECIPIENTS)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
