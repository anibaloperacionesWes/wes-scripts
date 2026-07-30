"""
Envía 3 correos separados (PDF) — La Florida, La Reina, Providencia mayo 2026.
Destinatarios: Juan, Diego, Aníbal; BCC José.
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
SUFIJO = "20260501_20260531"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_VISIBLE = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
    "anibal.aoperaciones@wes.cl",
]
BCC_ENVELOPE_ONLY = ["joseotarola@wes.cl"]

ENVIOS = [
    (
        ROOT / "reports" / "La_Florida" / "ABREGADO",
        f"Reporte_Agregado_La_Florida_{SUFIJO}.docx",
        "La Florida — Reporte agregado PDF — Mayo 2026",
        "Adjunto el reporte agregado (PDF) de La Florida (Liceo Alto Cordillera) "
        "para el periodo 01/05/2026 al 31/05/2026.\n",
    ),
    (
        ROOT / "reports" / "La_Reina" / "ABREGADO",
        f"Reporte_Agregado_La_Reina_{SUFIJO}.docx",
        "La Reina — Reporte agregado PDF — Mayo 2026",
        "Adjunto el reporte agregado (PDF) de La Reina (Eugenio María De Hostos) "
        "para el periodo 01/05/2026 al 31/05/2026.\n",
    ),
    (
        ROOT / "reports" / "Providencia" / "ABREGADO",
        f"Reporte_Agregado_Providencia_{SUFIJO}.docx",
        "Providencia — Reporte agregado PDF — Mayo 2026",
        "Adjunto el reporte agregado (PDF) de los colegios de Providencia "
        "(Lastarria, Carmela Carvajal, Liceo 7, Juan Pablo Duarte) "
        "para el periodo 01/05/2026 al 31/05/2026.\n",
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


def _ultimo_docx(base: Path, nombre: str) -> Path:
    candidatos = list(base.rglob(nombre)) if base.is_dir() else []
    if not candidatos:
        raise FileNotFoundError(f"No se encontró {nombre} en {base}")
    return max(candidatos, key=lambda p: p.stat().st_mtime)


def main() -> int:
    pw = _smtp_password()
    if not pw:
        print("[ERROR] Falta contraseña SMTP.", file=sys.stderr)
        return 1

    from control_nocturno import convertir_docx_a_pdf

    envelope = list(TO_VISIBLE) + list(BCC_ENVELOPE_ONLY)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)

        for base, docx_name, asunto, cuerpo_detalle in ENVIOS:
            docx = _ultimo_docx(base, docx_name)
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

            server.send_message(msg, to_addrs=envelope)
            print(f"  [OK] Correo enviado")

    print(f"\n[OK] 3 correos enviados. Para: {', '.join(TO_VISIBLE)}")
    print(f"[OK] Copia oculta: {', '.join(BCC_ENVELOPE_ONLY)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
