"""
Envía correos separados (PDF) — agregados julio 2026 (lote solicitado).
Destinatarios: Juan, Diego, Aníbal.

Uso:
  python enviar_agregados_julio2026_lote_pdf_equipo.py
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
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
SUFIJO = "20260701_20260728"
PERIODO_TXT = "01/07/2026 al 28/07/2026"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_VISIBLE = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
    "anibal.aoperaciones@wes.cl",
]

# (carpeta ABREGADO, nombre docx, asunto, detalle cuerpo)
ENVIOS = [
    (
        ROOT / "reports" / "Fundo_Zapallar" / "ABREGADO",
        f"Reporte_Agregado_Fundo_Zapallar_{SUFIJO}.docx",
        "Fundo Zapallar — Reporte agregado PDF — Julio 2026",
        f"Adjunto el reporte agregado (PDF) de Fundo Zapallar para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "DERCO" / "ABREGADO",
        f"Reporte_Agregado_DERCO_{SUFIJO}.docx",
        "DERCO — Reporte agregado PDF — Julio 2026",
        f"Adjunto el reporte agregado (PDF) de DERCO para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "Nido_de_Aguilas" / "ABREGADO",
        f"Reporte_Agregado_Nido_de_Aguilas_{SUFIJO}.docx",
        "Nido de Águilas — Reporte agregado PDF — Julio 2026",
        f"Adjunto el reporte agregado (PDF) de Nido de Águilas para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "Lo_Valledor" / "ABREGADO",
        f"Reporte_Agregado_Lo_valledor_{SUFIJO}.docx",
        "Lo Valledor — Reporte agregado PDF — Julio 2026",
        "Adjunto el reporte agregado (PDF) de Lo Valledor (P1 + Barrio Norte) "
        f"para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "UDD" / "ABREGADO",
        f"Reporte_Agregado_UDD_{SUFIJO}.docx",
        "UDD — Reporte agregado PDF — Julio 2026",
        f"Adjunto el reporte agregado (PDF) de UDD para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "Club_Providencia" / "ABREGADO",
        f"Reporte_Agregado_Club_Providencia_{SUFIJO}.docx",
        "Club Providencia — Reporte agregado PDF — Julio 2026",
        f"Adjunto el reporte agregado (PDF) de Club Providencia para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "AGUNSA_Lampa" / "ABREGADO",
        f"Reporte_Agregado_AGUNSA_{SUFIJO}.docx",
        "AGUNSA Lampa — Reporte agregado PDF — Julio 2026",
        "Adjunto el reporte agregado (PDF) de AGUNSA Lampa (Depósito + módulos) "
        f"para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "Agunsa_Intermodal" / "ABREGADO",
        f"Reporte_Agregado_AGUNSA_{SUFIJO}.docx",
        "AGUNSA Intermodal San Antonio — Reporte agregado PDF — Julio 2026",
        "Adjunto el reporte agregado (PDF) de AGUNSA Intermodal San Antonio "
        f"para el periodo {PERIODO_TXT}.\n",
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


def _docx_mas_reciente(base: Path, nombre: str) -> Path:
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

    ok = 0
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)

        for base, docx_name, asunto, cuerpo_detalle in ENVIOS:
            try:
                docx = _docx_mas_reciente(base, docx_name)
            except FileNotFoundError as e:
                print(f"[ERROR] {e}", file=sys.stderr)
                continue

            print(f"\n[INFO] {asunto}")
            print(f"  Word: {docx}")
            pdf = convertir_docx_a_pdf(docx)
            if not pdf or not Path(pdf).is_file():
                print(f"[ERROR] No se pudo convertir a PDF: {docx}", file=sys.stderr)
                continue
            pdf = Path(pdf)
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
            ok += 1

    print(f"\n[OK] {ok}/{len(ENVIOS)} correos enviados. Para: {', '.join(TO_VISIBLE)}")
    return 0 if ok == len(ENVIOS) else 1


if __name__ == "__main__":
    raise SystemExit(main())
