"""
Envía 3 correos separados (PDF) — Intermodal, AGUNSA Lampa, COPEC junio 2026.
Destinatarios: Juan, Diego, Aníbal.

Uso:
  python enviar_agregados_junio2026_intermodal_lampa_copec.py
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
SUFIJO = "20260601_20260630"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_VISIBLE = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
    "anibal.aoperaciones@wes.cl",
]

ENVIOS = [
    (
        ROOT / "reports" / "AGUNSA" / "ABREGADO" / "AGREGADO_20260701_1340",
        f"Reporte_Agregado_AGUNSA_{SUFIJO}.docx",
        "AGUNSA Intermodal San Antonio — Reporte agregado PDF — Junio 2026",
        "Adjunto el reporte agregado (PDF) de AGUNSA — Intermodal San Antonio "
        "para el periodo 01/06/2026 al 30/06/2026.\n",
    ),
    (
        ROOT / "reports" / "AGUNSA" / "ABREGADO" / "AGREGADO_20260701_1342",
        f"Reporte_Agregado_AGUNSA_{SUFIJO}.docx",
        "AGUNSA Lampa — Reporte agregado PDF — Junio 2026",
        "Adjunto el reporte agregado (PDF) de AGUNSA Lampa (Depósito + módulos D, ABC y E) "
        "para el periodo 01/06/2026 al 30/06/2026.\n",
    ),
    (
        ROOT / "reports" / "COPEC" / "ABREGADO" / "AGREGADO_20260701_1348",
        f"Reporte_Agregado_COPEC_{SUFIJO}.docx",
        "COPEC — Reporte agregado PDF — Junio 2026",
        "Adjunto el reporte agregado (PDF) de COPEC (11 puntos de monitoreo) "
        "para el periodo 01/06/2026 al 30/06/2026.\n",
    ),
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
    pw = _smtp_password()
    if not pw:
        print("[ERROR] Falta contraseña SMTP.", file=sys.stderr)
        return 1

    from control_nocturno import convertir_docx_a_pdf

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)

        for carpeta, docx_name, asunto, cuerpo_detalle in ENVIOS:
            docx = carpeta / docx_name
            if not docx.is_file():
                print(f"[ERROR] No existe: {docx}", file=sys.stderr)
                return 1

            print(f"\n[INFO] {asunto}")
            print(f"  Word: {docx}")
            pdf = convertir_docx_a_pdf(docx)
            if not pdf or not pdf.is_file():
                print(f"[ERROR] No se pudo convertir a PDF: {docx}", file=sys.stderr)
                return 1
            print(f"  PDF: {pdf}")

            msg = MIMEMultipart()
            msg["From"] = SMTP_USUARIO
            msg["To"] = ", ".join(TO_VISIBLE)
            msg["Subject"] = asunto
            msg.attach(
                MIMEText(
                    "Estimados Juan, Diego y Aníbal,\n\n"
                    f"{cuerpo_detalle}\n"
                    "Saludos cordiales,\n"
                    "Sistema WES\n",
                    "plain",
                    "utf-8",
                )
            )
            with open(pdf, "rb") as f:
                part = MIMEApplication(f.read(), _subtype="pdf")
                part.add_header("Content-Disposition", "attachment", filename=pdf.name)
                msg.attach(part)

            server.send_message(msg, to_addrs=TO_VISIBLE)
            print("  [OK] Correo enviado")

    print(f"\n[OK] 3 correos enviados. Para: {', '.join(TO_VISIBLE)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
