"""
Envía 7 correos separados (PDF) — agregados mayo 2026 (clientes batch).
Para: Juan, Diego, Aníbal | BCC: José.
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

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_VISIBLE = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
    "anibal.aoperaciones@wes.cl",
]
BCC_ENVELOPE_ONLY = ["joseotarola@wes.cl"]

DOCX_MAYO = "Reporte_Agregado_*_20260501_20260531.docx"

ENVIOS = [
    (
        ROOT / "reports" / "Fundo_Zapallar" / "ABREGADO" / "AGREGADO_20260601_1430",
        "Reporte_Agregado_Fundo_Zapallar_20260501_20260531.docx",
        "Fundo Zapallar — Reporte agregado PDF — Mayo 2026",
        "Adjunto el reporte agregado (PDF) de **Fundo Zapallar** (8 puntos) "
        "para el periodo 01/05/2026 al 31/05/2026.\n",
    ),
    (
        ROOT / "reports" / "DERCO" / "ABREGADO" / "AGREGADO_20260601_1439",
        "Reporte_Agregado_DERCO_20260501_20260531.docx",
        "DERCO — Reporte agregado PDF — Mayo 2026",
        "Adjunto el reporte agregado (PDF) de **DERCO** (Quilicura) "
        "para el periodo 01/05/2026 al 31/05/2026.\n",
    ),
    (
        ROOT / "reports" / "Nido_de_Aguilas" / "ABREGADO" / "AGREGADO_20260601_1448",
        "Reporte_Agregado_Nido_de_Aguilas_20260501_20260531.docx",
        "Nido de Aguilas — Reporte agregado PDF — Mayo 2026",
        "Adjunto el reporte agregado (PDF) de **Nido de Aguilas** "
        "para el periodo 01/05/2026 al 31/05/2026.\n",
    ),
    (
        ROOT / "reports" / "Lo_Valledor" / "Lo_valledor" / "ABREGADO" / "AGREGADO_20260601_1453",
        "Reporte_Agregado_Lo_valledor_20260501_20260531.docx",
        "Lo Valledor P1 — Reporte agregado PDF — Mayo 2026",
        "Adjunto el reporte agregado (PDF) de **Lo Valledor — P1** "
        "para el periodo 01/05/2026 al 31/05/2026.\n",
    ),
    (
        ROOT / "reports" / "Barrio_Norte" / "Lo_valledor" / "ABREGADO" / "AGREGADO_20260601_1454",
        "Reporte_Agregado_Lo_valledor_20260501_20260531.docx",
        "Barrio Norte — Reporte agregado PDF — Mayo 2026",
        "Adjunto el reporte agregado (PDF) de **Lo Valledor — Barrio Norte** "
        "para el periodo 01/05/2026 al 31/05/2026.\n",
    ),
    (
        ROOT / "reports" / "UDD" / "ABREGADO" / "AGREGADO_20260601_1456",
        "Reporte_Agregado_UDD_20260501_20260531.docx",
        "UDD — Reporte agregado PDF — Mayo 2026",
        "Adjunto el reporte agregado (PDF) de **UDD** (Honduras y Aula Magna) "
        "para el periodo 01/05/2026 al 31/05/2026.\n",
    ),
    (
        ROOT / "reports" / "Club_Providencia" / "ABREGADO" / "AGREGADO_20260601_1459",
        "Reporte_Agregado_Club_Providencia_20260501_20260531.docx",
        "Club Providencia — Reporte agregado PDF — Mayo 2026",
        "Adjunto el reporte agregado (PDF) de **Club Providencia** "
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

        for carpeta, docx_name, asunto, cuerpo_detalle in ENVIOS:
            docx = carpeta / docx_name
            if not docx.is_file():
                print(f"[ERROR] No existe: {docx}", file=sys.stderr)
                return 1

            print(f"\n[INFO] {asunto}")
            print(f"  Word: {docx}")
            pdf = docx.with_suffix(".pdf")
            if not (pdf.is_file() and pdf.stat().st_mtime >= docx.stat().st_mtime):
                pdf = convertir_docx_a_pdf(docx)
            if not pdf or not pdf.is_file():
                print(f"[ERROR] No se pudo convertir a PDF: {docx}", file=sys.stderr)
                return 1
            print(f"  PDF: {pdf}")

            msg = MIMEMultipart()
            msg["From"] = SMTP_USUARIO
            msg["To"] = ", ".join(TO_VISIBLE)
            msg["Subject"] = asunto
            cuerpo_limpio = cuerpo_detalle.replace("**", "")
            msg.attach(
                MIMEText(
                    "Estimados Juan, Diego y Aníbal,\n\n"
                    f"{cuerpo_limpio}\n"
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
            print("  [OK] Correo enviado")

    print(f"\n[OK] {len(ENVIOS)} correos enviados. Para: {', '.join(TO_VISIBLE)}")
    print(f"[OK] Copia oculta: {', '.join(BCC_ENVELOPE_ONLY)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
