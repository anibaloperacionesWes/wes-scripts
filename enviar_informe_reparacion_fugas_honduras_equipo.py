"""
Envía a Juan, Diego y Anibal el informe comparativo de reparación de fugas
en matriz — Sala impulsión Honduras (UDD).

Adjunta PDF + Word desde:
  reports/udd_reparacion_fugas_honduras/

Uso:
  python enviar_informe_reparacion_fugas_honduras_equipo.py
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
OUT = ROOT / "reports" / "udd_reparacion_fugas_honduras"
DOCX_PATH = OUT / "Informe_Reparacion_Fugas_Honduras_UDD.docx"
PDF_PATH = OUT / "Informe_Reparacion_Fugas_Honduras_UDD.pdf"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_RECIPIENTS = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
    "anibal.aoperaciones@wes.cl",
]
CC_RECIPIENTS: list[str] = []


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
    pw = _smtp_password()
    if not pw:
        print("[ERROR] Falta contraseña SMTP.", file=sys.stderr)
        return 1
    if not PDF_PATH.is_file():
        print(f"[ERROR] No existe: {PDF_PATH}", file=sys.stderr)
        return 1
    if not DOCX_PATH.is_file():
        print(f"[ERROR] No existe: {DOCX_PATH}", file=sys.stderr)
        return 1

    cuerpo = """Estimados Juan, Diego y Anibal,

Adjunto el informe comparativo de consumo del punto WES Sala impulsión Honduras (000026-01 — UDD), asociado a las reparaciones de fugas en la matriz realizadas el sábado 11/07/2026.

Resumen
-------
• Comparación: 7 días antes (04/07–10/07) vs 7 días después (12/07–18/07).
• El 11/07/2026 queda fuera del análisis y del gráfico comparativo (día de corte de agua).
• Gráfico de consumo diario en formato de reportes agregados: rojo = antes; azul = después.
• Resultado: 411,86 m³ antes vs 128,45 m³ después (−283,41 m³ / −68,8 %).

Adjuntos: PDF y Word.

Saludos,
Agente IA WES
"""

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    if CC_RECIPIENTS:
        msg["Cc"] = ", ".join(CC_RECIPIENTS)
    msg["Subject"] = (
        "UDD Honduras — Informe comparativo reparación fugas matriz (11/07/2026)"
    )
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for path, subtype in (
        (PDF_PATH, "pdf"),
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

    print(f"[OK] Para: {', '.join(TO_RECIPIENTS)}")
    print(f"     PDF:  {PDF_PATH}")
    print(f"     DOCX: {DOCX_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
