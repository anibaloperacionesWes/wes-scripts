"""
Envía a Juan y Aníbal el one-pager semanal de Fundo Zapallar.

Uso:
  python enviar_gestion_hidrica_semanal_zapallar_juan_anibal.py
  python enviar_gestion_hidrica_semanal_zapallar_juan_anibal.py --dry-run
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

ASUNTO = "Fundo Zapallar — seguimiento semanal 24 al 30 de agosto de 2026"

PDF = ROOT / "reports" / "Fundo_Zapallar" / "GESTION_HIDRICA" / "SEMANAL" / (
    "One_Pager_Semanal_Fundo_Zapallar_20260824_20260830.pdf"
)
DRIVE = (
    "https://drive.google.com/file/d/10oaaNtXfv-bcUF_9Z3jpSP88EqciIMpl/view?usp=drivesdk"
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


def enviar(*, dry_run: bool = False, drive: str = DRIVE) -> None:
    if not PDF.is_file():
        raise FileNotFoundError(f"No existe {PDF}")
    print(f"[OK] {PDF.name} ({PDF.stat().st_size // 1024} KB)")

    cuerpo = (
        "Estimados Juan y Aníbal,\n\n"
        "Adjunto el one-pager de seguimiento SEMANAL de Fundo Zapallar "
        "(24 al 30 de agosto de 2026), comparado con la semana previa (17 al 23).\n\n"
        "Este informe no reemplaza el de fin de mes: sirve para ver cambios en la semana "
        "y atacar alzas, picos o nocturno anómalo antes del cierre.\n\n"
        "Estado de la semana: En observación (33 % nocturno, bombas y estanques).\n"
        "Consumo de entrada (Matriz ESVAL): 358,3 m³ (−11 % vs la semana previa).\n\n"
        f"También en Drive:\n  • {drive}\n\n"
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
    with open(PDF, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=PDF.name)
        msg.attach(part)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=TO_VISIBLE)

    print(f"\n[OK] Correo enviado a {', '.join(TO_VISIBLE)}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--drive", default=DRIVE)
    args = parser.parse_args()
    try:
        enviar(dry_run=args.dry_run, drive=args.drive)
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
