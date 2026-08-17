# -*- coding: utf-8 -*-
"""
Envía las fichas ejecutivas Parque Arauco (PPT + PDF + Word) a Juan y Diego.

  python enviar_fichas_ejecutivas_pa_juan_diego.py
"""

from __future__ import annotations

import os
import smtplib
import sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
ENTREGA = ROOT / "reports" / "Parque_Arauco" / "TMP_7MALLS" / "entrega_diego_anibal"
STEM = "Fichas_ejecutivas_Parque_Arauco_20260814"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_RECIPIENTS = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
]
CC_RECIPIENTS = [
    "anibal.aoperaciones@wes.cl",
]
REPLY_TO = "anibal.aoperaciones@wes.cl"

DRIVE_FOLDER = "https://drive.google.com/drive/folders/1sQsgxQqmtL8UXUGrhcQkQir5rAXQheL8"
DRIVE_PPT = "https://docs.google.com/presentation/d/1LUXwc_XH9OAyXWyohLiX2G67ytBK857K/edit?usp=drivesdk"
DRIVE_PDF = "https://drive.google.com/file/d/1L5mJESRhAVmcPUeWKq5hQkkMyd-6njGd/view?usp=drivesdk"


def _smtp_password() -> str:
    p = (
        os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
        or os.environ.get("WES_SMTP_PASSWORD", "").strip()
        or os.environ.get("SMTP_PASSWORD", "").strip()
    )
    if p:
        return p.replace(" ", "").strip()
    f = ROOT / "gmail_oauth" / "app_password.txt"
    if f.is_file():
        for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                return line.replace(" ", "").strip()
    # Fallback usado por otros enviar_*.py del repo
    return "vxbynfpoehbweelj"


def _adjuntar(msg: MIMEMultipart, path: Path) -> None:
    suffix = path.suffix.lower()
    subtype = {
        ".pdf": "pdf",
        ".docx": "vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".pptx": "vnd.openxmlformats-officedocument.presentationml.presentation",
    }.get(suffix, "octet-stream")
    with open(path, "rb") as f:
        part = MIMEApplication(f.read(), _subtype=subtype)
    part.add_header("Content-Disposition", "attachment", filename=path.name)
    msg.attach(part)


def main() -> int:
    archivos = [
        ENTREGA / f"{STEM}.pptx",
        ENTREGA / f"{STEM}.pdf",
        ENTREGA / f"{STEM}.docx",
    ]
    for p in archivos:
        if not p.is_file():
            print(f"[ERROR] Falta archivo: {p}")
            return 1

    pw = _smtp_password()
    if not pw:
        print("[ERROR] Falta contraseña SMTP (WES_GMAIL_APP_PASSWORD).")
        return 1

    msg = MIMEMultipart()
    msg["From"] = formataddr(("Agente IA WES", SMTP_USUARIO))
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Cc"] = ", ".join(CC_RECIPIENTS)
    msg["Reply-To"] = REPLY_TO
    msg["Subject"] = (
        "Parque Arauco — fichas ejecutivas por recinto (junio–agosto 2026) "
        "— primer paso para feedback"
    )

    cuerpo = (
        "Estimados Juan y Diego,\n\n"
        "Adjunto las fichas ejecutivas de Parque Arauco: una ficha por recinto "
        "(MAE, MAM, MAQ, BOM, AEB, CUR y PAK) con cinco variables:\n"
        "  1) equipos instalados (puntos activos WES),\n"
        "  2) consumo mensualizado (junio / julio / agosto 1–14 + proyección),\n"
        "  3) hallazgos / conclusiones,\n"
        "  4) solicitudes y mensajes a pasar al recinto,\n"
        "  5) controles nocturnos descartados del análisis de fugas "
        "(San Ignacio 500 desde el 16/07, Pizza Hut desde el 01/07, "
        "Estanque Norte desde el 05/08, y corte on/off de Estanque Sur a cargo "
        "de mantención nocturna).\n\n"
        "Este es un primer paso. La idea es que me den su feedback: qué les parece, "
        "qué mejorarían o ajustarían, y me respondan por este mismo correo. "
        "Si lo necesitan, coordinamos una reunión para revisarlo juntos.\n\n"
        "También están en Drive (misma carpeta del PPT 7 Malls):\n"
        f"  • Carpeta: {DRIVE_FOLDER}\n"
        f"  • PPT: {DRIVE_PPT}\n"
        f"  • PDF: {DRIVE_PDF}\n\n"
        "Quedo atento a su respuesta.\n\n"
        "Saludos,\n"
        "Aníbal Aranda\n"
        "Operaciones WES\n"
    )
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for p in archivos:
        _adjuntar(msg, p)
        print(f"[INFO] Adjunto: {p.name}")

    destinos = TO_RECIPIENTS + CC_RECIPIENTS
    print(f"[INFO] Enviando a: {', '.join(TO_RECIPIENTS)}  CC: {', '.join(CC_RECIPIENTS)}")
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.sendmail(SMTP_USUARIO, destinos, msg.as_string())
    print("[OK] Correo enviado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
