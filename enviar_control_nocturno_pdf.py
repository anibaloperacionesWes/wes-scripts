"""
Envía por correo el PDF del reporte de control nocturno.

Por defecto adjunta el ``control_nocturno_*.pdf`` más reciente en
``reports/control_nocturno/`` (excluye variantes ``*_sin_PA_*`` y ``alerta_*``).
También considera PDF antiguos ``reporte_control_nocturno_*.pdf`` si no hay ninguno nuevo.
"""

from __future__ import annotations

import argparse
import os
import smtplib
import sys
from datetime import datetime
from pathlib import Path

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
SMTP_PASSWORD = (
    os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
    or os.environ.get("WES_SMTP_PASSWORD", "").strip()
    or "vxbynfpoehbweelj"
)

DEFAULT_DIR = Path("reports") / "control_nocturno"


def _ultimo_pdf_control_nocturno(dir_path: Path) -> Path:
    if not dir_path.is_dir():
        raise FileNotFoundError(f"No existe carpeta: {dir_path}")
    cands: list[Path] = []
    for pattern in ("control_nocturno_*.pdf", "reporte_control_nocturno_*.pdf"):
        for p in dir_path.glob(pattern):
            if (
                p.is_file()
                and "sin_PA" not in p.name
                and not p.name.startswith("alerta_")
                and not p.name.startswith("~$")
            ):
                cands.append(p)
    if not cands:
        raise FileNotFoundError(
            f"No hay control_nocturno_*.pdf (ni legado reporte_control_nocturno_*.pdf) en {dir_path} "
            "(excl. sin_PA y alerta_*)."
        )
    cands.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return cands[0]


def enviar(pdf_path: Path, *, to_email: str, cc_emails: list[str]) -> bool:
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = to_email
    if cc_emails:
        msg["Cc"] = ", ".join(cc_emails)
    fecha = datetime.now().strftime("%d-%m-%Y")
    msg["Subject"] = f"Reporte control nocturno — {fecha}"
    msg.attach(
        MIMEText(
            "Estimado equipo,\n\n"
            "Se adjunta el reporte de control nocturno (PDF).\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n",
            "plain",
            "utf-8",
        )
    )
    with open(pdf_path, "rb") as f:
        adj = MIMEApplication(f.read(), _subtype="pdf")
        adj.add_header("Content-Disposition", "attachment", filename=pdf_path.name)
        msg.attach(adj)
    dest = [to_email] + list(cc_emails)
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, SMTP_PASSWORD)
        server.send_message(msg, to_addrs=dest)
    return True


def main() -> int:
    ap = argparse.ArgumentParser(description="Enviar PDF control nocturno por correo")
    ap.add_argument("--to", default="anibal.aoperaciones@wes.cl")
    ap.add_argument("--cc", action="append", default=[])
    ap.add_argument("--pdf", type=Path, default=None, help="Ruta al PDF (si no: el más reciente)")
    ap.add_argument("--dir", type=Path, default=DEFAULT_DIR, help="Carpeta de reportes")
    args = ap.parse_args()

    pdf = args.pdf.expanduser().resolve() if args.pdf else _ultimo_pdf_control_nocturno(args.dir)
    if not pdf.is_file():
        print(f"[ERROR] No existe PDF: {pdf}")
        return 1
    print(f"[INFO] Adjunto: {pdf}")
    enviar(pdf, to_email=args.to, cc_emails=args.cc)
    print(f"[OK] Enviado a {args.to}")
    for c in args.cc:
        print(f"     CC {c}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
