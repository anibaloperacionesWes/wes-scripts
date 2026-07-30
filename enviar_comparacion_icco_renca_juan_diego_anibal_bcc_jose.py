"""
Envía comparación ICCO Renca (Word + PDF + Excel) a Juan, Diego y Aníbal.
José recibe copia oculta (solo en sobre SMTP, sin cabecera Bcc en el mensaje).
"""

from __future__ import annotations

import os
import smtplib
import subprocess
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
OUT_DIR = (
    ROOT
    / "reports"
    / "Renca"
    / "Coparacion App con Aguas Andinas"
    / "reporte_comparacion_Icco"
)

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_VISIBLE = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
    "anibal.aoperaciones@wes.cl",
]
BCC_ENVELOPE_ONLY = ["joseotarola@wes.cl"]


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
    return os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip() or ""


def main() -> int:
    pw = _smtp_password()
    if not pw:
        print(
            "[ERROR] Falta contraseña SMTP (WES_GMAIL_APP_PASSWORD, WES_SMTP_PASSWORD o gmail_oauth/app_password.txt).",
            file=sys.stderr,
        )
        return 1

    print("[1/3] Regenerando Excel ICCO…")
    r = subprocess.run(
        [sys.executable, str(ROOT / "exportar_excel_comparacion_facturacion_vs_wes.py"), "--site", "icco"],
        cwd=str(ROOT),
    )
    if r.returncode != 0:
        print("[ERROR] Falló exportación Excel.", file=sys.stderr)
        return 1

    docx = OUT_DIR / "Comparacion_App_vs_Facturaciones_ICCO_Renca.docx"
    pdf = OUT_DIR / "Comparacion_App_vs_Facturaciones_ICCO_Renca.pdf"
    xlsx = OUT_DIR / "Comparacion_Facturacion_vs_WES_ICCO_Renca.xlsx"
    for p in (docx, pdf, xlsx):
        if not p.is_file():
            print(f"[ERROR] Falta archivo: {p}", file=sys.stderr)
            return 1

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_VISIBLE)
    msg["Subject"] = "ICCO Renca — Comparación App WES vs facturaciones (Word + PDF + Excel)"
    msg.attach(
        MIMEText(
            "Estimados Juan, Diego y Aníbal,\n\n"
            "Adjunto el informe comparativo ICCO Renca (Word y PDF) y el Excel facturación vs WES "
            "(por defecto, boletas con ambas lecturas desde el 28-09-2025, alineado al criterio del informe Word).\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n",
            "plain",
            "utf-8",
        )
    )
    for p, subtype in (
        (docx, "vnd.openxmlformats-officedocument.wordprocessingml.document"),
        (pdf, "pdf"),
        (xlsx, "vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ):
        with open(p, "rb") as f:
            part = MIMEApplication(f.read(), _subtype=subtype)
            part.add_header("Content-Disposition", "attachment", filename=p.name)
            msg.attach(part)
        print(f"[OK] Adjunto: {p.name}")

    envelope = list(TO_VISIBLE) + list(BCC_ENVELOPE_ONLY)
    print("[2/3] Enviando correo…")
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=envelope)

    print("[3/3] Enviado.")
    print(f"  Para: {', '.join(TO_VISIBLE)}")
    print(f"  Copia oculta (sobre): {', '.join(BCC_ENVELOPE_ONLY)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
