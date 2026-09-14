# -*- coding: utf-8 -*-
"""Envía el consolidado PPT Parque Arauco 7 malls a Diego y Aníbal.

  python enviar_consolidado_pa_diego_anibal.py
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
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
PPT = ROOT / "reports" / "Parque_Arauco" / "CONSOLIDADO" / (
    "Consolidado_PA_7malls_20260908.pptx"
)
DRIVE = (
    "https://docs.google.com/presentation/d/13Hc1b559uwiTeq_QRtV5eROAfFd6Ngq9/"
    "edit?usp=drivesdk"
)

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_RECIPIENTS = [
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
    return (os.environ.get("WES_GMAIL_APP_PASSWORD") or "vxbynfpoehbweelj").replace(
        " ", ""
    ).strip()


def main() -> int:
    if not PPT.is_file():
        print(f"[ERROR] No existe: {PPT}")
        return 1

    pw = _smtp_password()
    if not pw:
        print("[ERROR] Falta contraseña SMTP")
        return 1

    asunto = (
        "Consolidado Parque Arauco 7 malls — PPT con antecedentes al 08/09/2026"
    )
    cuerpo = (
        "Estimados Diego y Aníbal,\n\n"
        "Adjunto el consolidado ejecutivo de Parque Arauco (MAE / MAM / MAQ / BOM / "
        "AEB / CUR / PAK), con la misma lectura del PPT de presentación y datos "
        "actualizados al 08/09/2026.\n\n"
        "También está en Drive:\n"
        f"{DRIVE}\n\n"
        "Nota MAM: el mall se alimenta por Placa Bancaria o por Falabella; cuando "
        "se inyecta una, se corta la otra. El 08/09 Placa baja a 4,5 m³ y Falabella "
        "sube a 92 m³: es cambio de alimentación, no dos fallas.\n\n"
        "Saludos,\n"
        "Agente IA WES\n"
    )

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = asunto
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    with open(PPT, "rb") as f:
        part = MIMEApplication(
            f.read(),
            _subtype="vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        part.add_header("Content-Disposition", "attachment", filename=PPT.name)
        msg.attach(part)

    print(f"[INFO] Enviando a: {', '.join(TO_RECIPIENTS)}")
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=TO_RECIPIENTS)
    print("[OK] Correo enviado con PPT adjunto.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
