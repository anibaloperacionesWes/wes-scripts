"""
Envía a Juan los informes de gestión hídrica (one-pager + mensual)
de Liceo 7 Luisa Saavedra — agosto 2026.

Uso:
  python enviar_gestion_hidrica_liceo7_agosto2026_juan.py
  python enviar_gestion_hidrica_liceo7_agosto2026_juan.py --dry-run
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
]

ASUNTO = "Informes de gestión hídrica — Liceo 7 — Agosto 2026"

ONE_PAGER_DRIVE = (
    "https://drive.google.com/file/d/1TA_mPjLa7_JRbwImdtSthqndxQ0LNDaB/view?usp=drivesdk"
)
MENSUAL_DRIVE = (
    "https://drive.google.com/file/d/1GMMs9ldNCHwv9IpEMQVuq0inM_61ICl_/view?usp=drivesdk"
)

ESTADO = (
    "Requiere atención — 9.465,8 m³ (35,3× mediana previa de 268 m³); "
    "305 m³/día; 1.115,6 m³ nocturnos (12 %); pico 7/08 con 602,4 m³"
)

BASE = ROOT / "reports" / "Providencia" / "Liceo_7" / "GESTION_HIDRICA"


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


def _adjuntos() -> list[tuple[str, Path]]:
    ones = sorted(BASE.glob("One_Pager_*.pdf"))
    mens = sorted(BASE.glob("Informe_Mensual_*.pdf"))
    if not ones or not mens:
        raise FileNotFoundError(f"Faltan PDF en {BASE}")
    return [(ones[-1].name, ones[-1]), (mens[-1].name, mens[-1])]


def enviar(*, dry_run: bool = False) -> None:
    pdfs = _adjuntos()
    for nombre, path in pdfs:
        print(f"[OK] {nombre} ({path.stat().st_size // 1024} KB)")

    cuerpo = (
        "Estimado Juan,\n\n"
        "Adjunto los informes de gestión hídrica (formato ejecutivo Zapallar) "
        "del Liceo 7 Luisa Saavedra (000006-04), periodo 1 al 31 de agosto de 2026.\n\n"
        "Van dos PDF: one-pager e informe mensual, aislados del agregado de "
        "Colegios Providencia para diagnosticar el salto de consumo "
        "(concentró ~84 % del total del mes).\n\n"
        f"  • Liceo 7: {ESTADO}\n\n"
        "También en Drive:\n"
        f"  • One-pager: {ONE_PAGER_DRIVE}\n"
        f"  • Mensual: {MENSUAL_DRIVE}\n\n"
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

    for nombre, pdf in pdfs:
        with open(pdf, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", "attachment", filename=nombre)
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
