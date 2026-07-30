# -*- coding: utf-8 -*-
"""
Envía el Informe WES Parque Arauco 7 Malls a Juan, Diego y Aníbal.

  python enviar_informe_pa_7malls_diego_anibal.py
"""

from __future__ import annotations

import os
import shutil
import smtplib
import sys
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
PPT = ROOT / "reports" / "_tmp_pa_7malls_charts" / (
    "Informe WES __ Parque Arauco 7 Malls (07.07.2026).pptx"
)
ENTREGA_DIR = ROOT / "reports" / "_tmp_pa_7malls_charts" / "entrega_diego_anibal"

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


def _copiar_entrega(ppt: Path) -> Path:
    ENTREGA_DIR.mkdir(parents=True, exist_ok=True)
    dest = ENTREGA_DIR / ppt.name
    shutil.copy2(ppt, dest)
    return dest


def _enviar(ppt: Path) -> None:
    pw = _smtp_password()
    if not pw:
        raise RuntimeError(
            "Falta contraseña SMTP (WES_GMAIL_APP_PASSWORD o gmail_oauth/app_password.txt)"
        )

    fecha = datetime.now().strftime("%d-%m-%Y")
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = (
        f"Informe WES Parque Arauco 7 Malls — MAE/MAM/MAQ/BOM/AEB/CUR/PAK ({fecha})"
    )
    msg.attach(
        MIMEText(
            "Estimados Juan, Diego y Aníbal,\n\n"
            "Adjunto el PowerPoint del informe WES Parque Arauco 7 Malls "
            "(período 01/05/2026 – 07/07/2026).\n\n"
            "Contenido actualizado:\n"
            "- MAE / MAM / MAQ / BOM / AEB / CUR\n"
            "- PAK: ranking, cadena abastecimiento DL, consumo diario por puntos "
            "y patrón nocturno 0–8 h\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n",
            "plain",
            "utf-8",
        )
    )
    with open(ppt, "rb") as f:
        part = MIMEApplication(
            f.read(),
            _subtype="vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        part.add_header("Content-Disposition", "attachment", filename=ppt.name)
        msg.attach(part)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=TO_RECIPIENTS)


def main() -> int:
    if not PPT.is_file():
        print(f"[ERROR] No existe: {PPT}")
        return 1

    copia = _copiar_entrega(PPT)
    print(f"[OK] Copia en entrega: {copia}")

    print(f"[INFO] Enviando a: {', '.join(TO_RECIPIENTS)}")
    _enviar(PPT)
    print("[OK] Correo enviado con PPT adjunto.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
