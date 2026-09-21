"""
Envía un correo consolidado con los 8 PDFs del lote agosto 2026.
Destinatarios: Juan y Aníbal (un solo correo, todos los adjuntos).

Uso:
  python enviar_consolidado_agosto2026_juan_anibal.py
  python enviar_consolidado_agosto2026_juan_anibal.py --dry-run
"""

from __future__ import annotations

import argparse
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
SUFIJO = "20260801_20260828"
PERIODO_TXT = "01/08/2026 al 28/08/2026"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_VISIBLE = [
    "juanlopez@wes.cl",
    "anibal.aoperaciones@wes.cl",
]

CLIENTES = [
    ("Fundo Zapallar", ROOT / "reports" / "Fundo_Zapallar" / "ABREGADO"),
    ("Inchcape (ex DERCO)", ROOT / "reports" / "Inchcape" / "ABREGADO"),
    ("Nido de Águilas", ROOT / "reports" / "Nido_de_Aguilas" / "ABREGADO"),
    ("Lo Valledor", ROOT / "reports" / "Lo_Valledor" / "ABREGADO"),
    ("UDD", ROOT / "reports" / "UDD" / "ABREGADO"),
    ("Club Providencia", ROOT / "reports" / "Club_Providencia" / "ABREGADO"),
    ("AGUNSA Lampa", ROOT / "reports" / "AGUNSA_Lampa" / "ABREGADO"),
    ("AGUNSA Intermodal San Antonio", ROOT / "reports" / "Agunsa_Intermodal" / "ABREGADO"),
]

DRIVE_LINKS = [
    ("Fundo Zapallar", "https://drive.google.com/file/d/10hP-cnAMFJC8ny1YMH-22TEDsyeMFiNf/view?usp=drivesdk"),
    ("Inchcape", "https://drive.google.com/file/d/1Ajx17aVXqNjGcqtYQZuePsBeCV7zRANK/view?usp=drivesdk"),
    ("Nido de Águilas", "https://drive.google.com/file/d/19De85VA-Y5SqUSri9Opidtqb70l3SqJL/view?usp=drivesdk"),
    ("Lo Valledor", "https://drive.google.com/file/d/13v2FQ1ozgscpSOogpUQ9KzZg6PFUPMQg/view?usp=drivesdk"),
    ("UDD", "https://drive.google.com/file/d/1i5d6U35qywysnSjL_Rwb0P2tx4ZYcgjR/view?usp=drivesdk"),
    ("Club Providencia", "https://drive.google.com/file/d/17qnLaliECs8m0ndO-IMQdOLksWoVHCl_/view?usp=drivesdk"),
    ("AGUNSA Lampa", "https://drive.google.com/file/d/1Js6hbETsjrGu5f4PVpwa6BUxqPDEtibO/view?usp=drivesdk"),
    ("AGUNSA Intermodal", "https://drive.google.com/file/d/1KXTWXFhU9yTee9vx7su97mxW32idZiQP/view?usp=drivesdk"),
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


def _pdf_mas_reciente(base: Path) -> Path:
    pattern = f"Reporte_Agregado_*_{SUFIJO}.pdf"
    candidatos = list(base.rglob(pattern)) if base.is_dir() else []
    if not candidatos:
        raise FileNotFoundError(f"No se encontró {pattern} en {base}")
    return max(candidatos, key=lambda p: p.stat().st_mtime)


def enviar_consolidado(*, dry_run: bool = False) -> None:
    pdfs: list[tuple[str, Path]] = []
    for label, base in CLIENTES:
        pdf = _pdf_mas_reciente(base)
        print(f"[OK] {label}: {pdf}")
        pdfs.append((pdf.name, pdf))

    lista = "\n".join(f"  • {label}" for label, _ in CLIENTES)
    links = "\n".join(f"  • {label}: {url}" for label, url in DRIVE_LINKS)
    cuerpo = (
        "Estimados Juan y Aníbal,\n\n"
        f"Adjunto el consolidado de reportes agregados en PDF para el periodo {PERIODO_TXT}:\n\n"
        f"{lista}\n\n"
        "Son 8 archivos PDF adjuntos en este mismo correo.\n\n"
        "También disponibles en Drive:\n"
        f"{links}\n\n"
        "Saludos cordiales,\n"
        "Sistema WES\n"
    )

    if dry_run:
        print("\n[DRY-RUN] Asunto: Consolidado — Reportes agregados Agosto 2026 (8 clientes)")
        print(f"[DRY-RUN] Para: {', '.join(TO_VISIBLE)}")
        print(cuerpo)
        return

    pw = _smtp_password()
    if not pw:
        raise RuntimeError(
            "Falta contraseña SMTP. Configurá WES_GMAIL_APP_PASSWORD o WES_SMTP_PASSWORD "
            "en Cloud Agents → Secrets, o gmail_oauth/app_password.txt"
        )

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_VISIBLE)
    msg["Subject"] = "Consolidado — Reportes agregados Agosto 2026 (8 clientes)"
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for nombre, pdf in pdfs:
        with open(pdf, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", "attachment", filename=nombre)
            msg.attach(part)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=TO_VISIBLE)

    print(f"\n[OK] Correo consolidado enviado a {', '.join(TO_VISIBLE)}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    try:
        enviar_consolidado(dry_run=args.dry_run)
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
