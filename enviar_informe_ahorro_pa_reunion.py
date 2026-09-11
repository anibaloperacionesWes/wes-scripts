# -*- coding: utf-8 -*-
"""Envía el informe de ahorro PA (propuesta n.° 1) a Juan, Diego y Aníbal.

  python enviar_informe_ahorro_pa_reunion.py
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
DOCX = ROOT / "reports" / "Parque_Arauco" / "CONSOLIDADO" / (
    "Informe_ahorro_PA_reunion_20260908.docx"
)
DRIVE = (
    "https://docs.google.com/document/d/1MjWiw0ZKbTjlZFTOX7iOo82F_Tl-Ff8G/"
    "edit?usp=drivesdk"
)

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
    return (os.environ.get("WES_GMAIL_APP_PASSWORD") or "vxbynfpoehbweelj").replace(
        " ", ""
    ).strip()


def main() -> int:
    if not DOCX.is_file():
        print(f"[ERROR] No existe: {DOCX}")
        return 1

    pw = _smtp_password()
    if not pw:
        print("[ERROR] Falta contraseña SMTP")
        return 1

    asunto = (
        "Parque Arauco — propuesta n.° 1 de informe de ahorro para reunión"
    )
    cuerpo = (
        "Estimados Juan, Diego y Aníbal,\n\n"
        "Les enviamos la propuesta n.° 1 del informe de ahorro de Parque Arauco "
        "para la reunión (MAE, Buenaventura, Quilicura, Kennedy y pendiente de Maipú).\n\n"
        "Va adjunto en Word y también está en Drive:\n"
        f"{DRIVE}\n\n"
        "Es una primera versión para revisar fondo y números. Si lo desean, "
        "podemos generar mejoras de apariencia (formato, gráficos y diagramación) "
        "sobre esta misma base.\n\n"
        "Quedamos atentos a sus comentarios.\n\n"
        "Saludos,\n"
        "Agente IA WES\n"
    )

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = asunto
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    with open(DOCX, "rb") as f:
        part = MIMEApplication(
            f.read(),
            _subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        part.add_header("Content-Disposition", "attachment", filename=DOCX.name)
        msg.attach(part)

    print(f"[INFO] Enviando a: {', '.join(TO_RECIPIENTS)}")
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=TO_RECIPIENTS)
    print("[OK] Correo enviado con Word adjunto.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
