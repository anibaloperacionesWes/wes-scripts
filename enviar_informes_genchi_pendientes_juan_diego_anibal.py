"""
Envía informes Genchi pendientes (Word + PDF), uno por correo.
Destinatarios: Juan, Diego y Aníbal. José en copia oculta (sobre SMTP).

No reenvía CDP Puente Alto (ya enviado).

  python enviar_informes_genchi_pendientes_juan_diego_anibal.py
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
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
GENCHI = ROOT / "reports" / "Genchi"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_VISIBLE = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
    "anibal.aoperaciones@wes.cl",
]
BCC_ENVELOPE_ONLY = ["joseotarola@wes.cl"]

INFORMES = [
    {
        "dir": GENCHI / "CPF San Miguel" / "informe",
        "stem": "Informe_Genchi_CPF_San_Miguel_facturaciones_20260605_1501",
        "subject": "Gendarmería Genchi — CPF San Miguel — Informe comparativo facturación Aguas Andinas (Word + PDF)",
        "body": (
            "Estimados Juan, Diego y Aníbal,\n\n"
            "Adjunto el informe comparativo de facturación Aguas Andinas para "
            "CPF San Miguel (cuenta 1008398-2), con análisis con y sin monitoreo WES "
            "(retiro WES 02-07-2025). Word y PDF adjuntos.\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n"
        ),
    },
    {
        "dir": GENCHI / "CCP Santiago Sur" / "informe",
        "stem": "Informe_Genchi_CCP_Santiago_Sur_facturaciones_20260605_1503",
        "subject": "Gendarmería Genchi — CCP Santiago Sur — Informe comparativo facturación Aguas Andinas (Word + PDF)",
        "body": (
            "Estimados Juan, Diego y Aníbal,\n\n"
            "Adjunto el informe comparativo de facturación Aguas Andinas para "
            "CCP Santiago Sur (cuenta 1007968-3), con análisis con y sin monitoreo WES "
            "(retiro WES 30-06-2025). Word y PDF adjuntos.\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n"
        ),
    },
    {
        "dir": GENCHI / "informe",
        "stem": "Informe_Genchi_Agregado_facturaciones_20260605_1615",
        "subject": "Gendarmería Genchi — Informe agregado facturación Aguas Andinas — 3 unidades (Word + PDF)",
        "body": (
            "Estimados Juan, Diego y Aníbal,\n\n"
            "Adjunto el informe agregado de facturación Aguas Andinas que consolida "
            "CDP Puente Alto, CPF San Miguel y CCP Santiago Sur (Word y PDF).\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n"
        ),
    },
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
        try:
            for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
                line = line.strip()
                if line and not line.startswith("#"):
                    return line.replace(" ", "").strip()
        except OSError:
            pass
    return ""


def _enviar_uno(server: smtplib.SMTP, item: dict) -> None:
    docx = item["dir"] / f"{item['stem']}.docx"
    pdf = item["dir"] / f"{item['stem']}.pdf"
    for p in (docx, pdf):
        if not p.is_file():
            raise FileNotFoundError(p)

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_VISIBLE)
    msg["Subject"] = item["subject"]
    msg.attach(MIMEText(item["body"], "plain", "utf-8"))

    for p, subtype in (
        (docx, "vnd.openxmlformats-officedocument.wordprocessingml.document"),
        (pdf, "pdf"),
    ):
        with open(p, "rb") as f:
            part = MIMEApplication(f.read(), _subtype=subtype)
            part.add_header("Content-Disposition", "attachment", filename=p.name)
            msg.attach(part)

    envelope = list(TO_VISIBLE) + list(BCC_ENVELOPE_ONLY)
    server.send_message(msg, to_addrs=envelope)
    print(f"[OK] Enviado: {item['stem']}")


def main() -> int:
    pw = _smtp_password()
    if not pw:
        print("[ERROR] Falta contraseña SMTP.", file=sys.stderr)
        return 1

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        for i, item in enumerate(INFORMES, 1):
            print(f"[{i}/{len(INFORMES)}] {item['stem']}…")
            _enviar_uno(server, item)

    print("[OK] Todos los informes pendientes enviados.")
    print(f"  Para: {', '.join(TO_VISIBLE)}")
    print(f"  Copia oculta (sobre): {', '.join(BCC_ENVELOPE_ONLY)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
