"""
Envía a Juan y Aníbal la versión de Fundo Zapallar con periodo 5–31/08/2026.

Uso:
  python enviar_gestion_hidrica_zapallar_agosto2026_juan_anibal.py
  python enviar_gestion_hidrica_zapallar_agosto2026_juan_anibal.py --dry-run
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

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_VISIBLE = [
    "juanlopez@wes.cl",
    "anibal.aoperaciones@wes.cl",
]

ASUNTO = "Fundo Zapallar gestión hídrica agosto 2026 — 5 al 31 (versión actualizada)"

ONE_PAGER = ROOT / "reports" / "Fundo_Zapallar" / "GESTION_HIDRICA" / (
    "One_Pager_Gestion_Hidrica_Fundo_Zapallar_Agosto_2026.pdf"
)
MENSUAL = ROOT / "reports" / "Fundo_Zapallar" / "GESTION_HIDRICA" / (
    "Informe_Mensual_Fundo_Zapallar_Agosto_2026.pdf"
)

DRIVE_ONE = (
    "https://drive.google.com/file/d/193B4N30jB-PBrO4LMZcJhm0DXL99q2J4/view?usp=drivesdk"
)
DRIVE_MEN = (
    "https://drive.google.com/file/d/1PYIb37zAar1aPN59dw1k28rFQs9qPdGQ/view?usp=drivesdk"
)


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
    return "vxbynfpoehbweelj"


def enviar(*, dry_run: bool = False) -> None:
    pdfs = [ONE_PAGER, MENSUAL]
    for path in pdfs:
        if not path.is_file():
            raise FileNotFoundError(f"No existe {path}")
        print(f"[OK] {path.name} ({path.stat().st_size // 1024} KB)")

    cuerpo = (
        "Estimados Juan y Aníbal,\n\n"
        "Adjunto la versión actualizada de Fundo Zapallar (formato ejecutivo Zapallar). "
        "El periodo informado es del 5 al 31 de agosto de 2026: los días 1 al 4 se excluyen "
        "y no se extrapola el mes.\n\n"
        "Estado: En observación (30 % nocturno, bombas y estanques).\n"
        "Consumo de entrada (Matriz ESVAL): 1.613,0 m³. Los puntos internos no se suman.\n\n"
        "También en Drive:\n"
        f"  • One-pager: {DRIVE_ONE}\n"
        f"  • Mensual: {DRIVE_MEN}\n\n"
        "Saludos cordiales,\n"
        "Sistema WES\n"
    )

    if dry_run:
        print(f"\n[DRY-RUN] Asunto: {ASUNTO}")
        print(f"[DRY-RUN] Para: {', '.join(TO_VISIBLE)}")
        print(cuerpo)
        return

    pw = _smtp_password()
    if not pw:
        raise RuntimeError("Falta contraseña SMTP (WES_GMAIL_APP_PASSWORD).")

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_VISIBLE)
    msg["Subject"] = ASUNTO
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for pdf in pdfs:
        with open(pdf, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", "attachment", filename=pdf.name)
            msg.attach(part)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=TO_VISIBLE)

    print(f"\n[OK] Correo enviado a {', '.join(TO_VISIBLE)}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    try:
        enviar(dry_run=args.dry_run)
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
