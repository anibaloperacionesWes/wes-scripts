"""
Convierte el último reporte de puntos en cero a PDF y lo envía a las 09:00 (hora Chile).

Destinatarios por defecto: Mauricio, Juan, Diego, Aníbal.
"""

from __future__ import annotations

import argparse
import os
import smtplib
import sys
import time
from datetime import datetime, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from zoneinfo import ZoneInfo

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from wes_paths import reporte_cero_dir

CERO_DIR = reporte_cero_dir()
TZ_CHILE = ZoneInfo("America/Santiago")

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

DESTINATARIOS = [
    "mauricioorellana@wes.cl",
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


def _ultimo_docx() -> Path:
    files = [
        f
        for f in CERO_DIR.glob("Reporte_Puntos_En_Cero_*.docx")
        if f.is_file() and not f.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError(f"No hay reporte .docx en {CERO_DIR}")
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0]


def _esperar_hasta(hora: int, minuto: int) -> None:
    ahora = datetime.now(TZ_CHILE)
    objetivo = ahora.replace(hour=hora, minute=minuto, second=0, microsecond=0)
    if ahora >= objetivo:
        print(f"[INFO] Ya pasaron las {hora:02d}:{minuto:02d} — envío inmediato.")
        return
    delta = (objetivo - ahora).total_seconds()
    print(
        f"[INFO] Esperando hasta {objetivo.strftime('%d/%m/%Y %H:%M')} (Chile) "
        f"— {int(delta // 60)} min {int(delta % 60)} s...",
        flush=True,
    )
    while True:
        restante = (objetivo - datetime.now(TZ_CHILE)).total_seconds()
        if restante <= 0:
            break
        time.sleep(min(restante, 30))


def _enviar(pdf: Path) -> None:
    fecha = datetime.now(TZ_CHILE).strftime("%d-%m-%Y")
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(DESTINATARIOS)
    msg["Subject"] = f"Reporte de Puntos en Cero — {fecha}"
    msg.attach(
        MIMEText(
            "Estimados Mauricio, Juan, Diego y Aníbal,\n\n"
            "Se adjunta el reporte de puntos en cero (PDF).\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n",
            "plain",
            "utf-8",
        )
    )
    with open(pdf, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=pdf.name)
        msg.attach(part)

    pw = _smtp_password()
    if not pw:
        raise RuntimeError("Falta contraseña SMTP (WES_SMTP_PASSWORD o gmail_oauth/app_password.txt).")

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=DESTINATARIOS)

    print("[OK] Correo enviado a:")
    for e in DESTINATARIOS:
        print(f"  - {e}")


def main() -> int:
    ap = argparse.ArgumentParser(description="PDF puntos en cero + envío programado")
    ap.add_argument("--hora", type=int, default=9, help="Hora Chile (default 9)")
    ap.add_argument("--minuto", type=int, default=0, help="Minuto (default 0)")
    ap.add_argument("--ahora", action="store_true", help="Enviar de inmediato sin esperar")
    args = ap.parse_args()

    print("=" * 70)
    print("  PUNTOS EN CERO — PDF + ENVÍO PROGRAMADO")
    print("=" * 70)

    docx = _ultimo_docx()
    print(f"[OK] Word: {docx}")

    pdf = docx.with_suffix(".pdf")
    if not (pdf.is_file() and pdf.stat().st_mtime >= docx.stat().st_mtime):
        from generar_reporte_word import convertir_word_a_pdf

        pdf = convertir_word_a_pdf(docx)
        if not pdf or not pdf.is_file():
            print("[ERROR] No se pudo generar el PDF.", file=sys.stderr)
            return 1
    print(f"[OK] PDF: {pdf}", flush=True)

    if not args.ahora:
        _esperar_hasta(args.hora, args.minuto)

    print(f"[INFO] Enviando correo ({datetime.now(TZ_CHILE).strftime('%H:%M:%S')} Chile)...")
    _enviar(pdf)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
