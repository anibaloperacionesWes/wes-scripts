# -*- coding: utf-8 -*-
"""
Envía el PPT de recorrido ejecutivo Parque Arauco a Diego y Aníbal.

  python3 enviar_recorrido_ejecutivo_pa.py
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

ROOT = Path(__file__).resolve().parent
PPT = (
    ROOT
    / "reports"
    / "Parque_Arauco"
    / "TMP_7MALLS"
    / "entrega_diego_anibal"
    / "Recorrido_ejecutivo_PA_MAE_20260817.pptx"
)

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
TO_RECIPIENTS = [
    "diegocarrasco@wes.cl",
    "anibal.aoperaciones@wes.cl",
]
REPLY_TO = "anibal.aoperaciones@wes.cl"
DRIVE_FOLDER = "https://drive.google.com/drive/folders/1sQsgxQqmtL8UXUGrhcQkQir5rAXQheL8"
DRIVE_PPT = "https://docs.google.com/presentation/d/1O-Ahgs3ycLQNjFRu5uFye9KJfqtuxNfw/edit?usp=drivesdk"


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
    return "vxbynfpoehbweelj"


def main() -> int:
    if not PPT.is_file():
        print(f"[ERROR] Falta archivo: {PPT}")
        return 1
    pw = _smtp_password()
    if not pw:
        print("[ERROR] Falta contraseña SMTP (WES_GMAIL_APP_PASSWORD).")
        return 1

    msg = MIMEMultipart()
    msg["From"] = formataddr(("Agente IA WES", SMTP_USUARIO))
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Reply-To"] = REPLY_TO
    msg["Subject"] = (
        "Parque Arauco — recorrido ejecutivo por recinto (mayo–agosto 2026)"
    )
    cuerpo = (
        "Estimados Diego y Aníbal,\n\n"
        "Adjunto la última versión del PPT de recorrido ejecutivo Parque Arauco "
        "(01/05/2026 – 17/08/2026).\n\n"
        "Una lámina de presentación por recinto y lámina 2 de hallazgos en "
        "MAE, MAM, MAQ, BOM, AEB, CUR y PAK.\n\n"
        "En corto:\n"
        "  • MAE: presostatos Sur 10/06 y controles nocturnos Norte / Pizza Hut.\n"
        "  • MAM: Placa 68 % de julio; desde el 15/08 el mall sale por Falabella. Umbrales.\n"
        "  • MAQ: Matriz 99 %. Alza desde el 22/06. Umbral diario y noche; propuesta de control.\n"
        "  • BOM: San Ignacio 500 con control desde el 17/07 (noche 32 → 2 m³). "
        "Ahorro ~$1.247.000/mes. San Ignacio 300: alerta de umbral.\n"
        "  • AEB: Matriz A.A. desactivada 15/05. Alza Anillo Plaza (mayo 7 → julio 21 m³/día). "
        "Umbrales del total del día: Anillo 22 · Matriz 75 m³/día.\n"
        "  • CUR: Sur un poco sobre Norte (julio 54/46). Noche 00–06 ~0,8 vs 0,4 m³. "
        "Umbrales del día: Sur 14 · Norte 13 m³/día.\n"
        "  • PAK: cadena Sandía → Distrito de Lujo → Bazar / DL Kennedy. "
        "Hueco nocturno en el tronco (~13 m³/noche).\n\n"
        "También está en Drive:\n"
        f"  • PPT: {DRIVE_PPT}\n"
        f"  • Carpeta: {DRIVE_FOLDER}\n\n"
        "Saludos,\n"
        "Aníbal Aranda\n"
        "Operaciones WES\n"
    )
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))
    with open(PPT, "rb") as f:
        part = MIMEApplication(
            f.read(),
            _subtype="vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    part.add_header("Content-Disposition", "attachment", filename=PPT.name)
    msg.attach(part)

    print(f"[INFO] Adjunto: {PPT.name} ({PPT.stat().st_size} bytes)")
    print(f"[INFO] Enviando a: {', '.join(TO_RECIPIENTS)}")
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.sendmail(SMTP_USUARIO, TO_RECIPIENTS, msg.as_string())
    print("[OK] Correo enviado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
