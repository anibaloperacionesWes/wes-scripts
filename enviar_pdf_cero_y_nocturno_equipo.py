"""
Convierte el último reporte de puntos en cero a PDF y envía ambos PDF
(puntos en cero + control nocturno) a Aníbal, Mauricio, Juan y Diego.

Uso:
  python enviar_pdf_cero_y_nocturno_equipo.py
  python enviar_pdf_cero_y_nocturno_equipo.py --solo-enviar
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

from generar_reporte_word import convertir_word_a_pdf

from wes_paths import reporte_cero_dir, wes_scripts_root

CERO_DIR = reporte_cero_dir()
NOCTURNO_DIR = wes_scripts_root() / "reports" / "control_nocturno"

DESTINATARIOS = [
    "anibal.aoperaciones@wes.cl",
    "mauricioorellana@wes.cl",
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
]

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587


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


def _ultimo_docx_cero() -> Path:
    files = [
        f
        for f in CERO_DIR.glob("Reporte_Puntos_En_Cero_*.docx")
        if f.is_file() and not f.name.startswith("~$")
    ]
    if not files:
        files = [
            f for f in CERO_DIR.glob("*.docx") if f.is_file() and not f.name.startswith("~$")
        ]
    if not files:
        raise FileNotFoundError(f"No hay .docx en {CERO_DIR}")
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0]


def _ultimo_pdf_nocturno() -> Path:
    cands: list[Path] = []
    for pattern in ("control_nocturno_*.pdf", "reporte_control_nocturno_*.pdf"):
        for p in NOCTURNO_DIR.glob(pattern):
            if (
                p.is_file()
                and "sin_PA" not in p.name
                and not p.name.startswith("alerta_")
                and not p.name.startswith("~$")
            ):
                cands.append(p)
    if not cands:
        raise FileNotFoundError(f"No hay PDF de control nocturno en {NOCTURNO_DIR}")
    cands.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return cands[0]


def _asegurar_pdf_cero() -> Path:
    docx = _ultimo_docx_cero()
    pdf = docx.with_suffix(".pdf")
    if pdf.is_file() and pdf.stat().st_mtime >= docx.stat().st_mtime:
        return pdf
    out = convertir_word_a_pdf(docx)
    if not out or not out.is_file():
        raise RuntimeError(f"No se pudo convertir a PDF: {docx}")
    return out


def _enviar_pdfs(pdfs: list[Path]) -> None:
    fecha = datetime.now().strftime("%d-%m-%Y")
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(DESTINATARIOS)
    msg["Subject"] = f"Reportes WES: Puntos en cero + Control nocturno — {fecha}"
    nombres = "\n".join(f"- {p.name}" for p in pdfs)
    msg.attach(
        MIMEText(
            "Estimados Aníbal, Mauricio, Juan y Diego,\n\n"
            "Se adjuntan los reportes del día en PDF:\n"
            f"{nombres}\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n",
            "plain",
            "utf-8",
        )
    )
    for p in pdfs:
        with open(p, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", "attachment", filename=p.name)
            msg.attach(part)
    pw = _smtp_password()
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=DESTINATARIOS)


def main() -> int:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    ap = argparse.ArgumentParser(description="PDF y envío cero + control nocturno al equipo")
    ap.add_argument(
        "--solo-enviar",
        action="store_true",
        help="No reconvierte puntos en cero; usa PDF ya existente junto al .docx",
    )
    args = ap.parse_args()

    print("=" * 70)
    print("  PDF + ENVÍO: PUNTOS EN CERO Y CONTROL NOCTURNO")
    print("=" * 70)

    if args.solo_enviar:
        docx = _ultimo_docx_cero()
        pdf_cero = docx.with_suffix(".pdf")
        if not pdf_cero.is_file():
            print(f"[ERROR] No existe PDF: {pdf_cero}")
            return 1
    else:
        print("[1/3] Convirtiendo puntos en cero a PDF...")
        pdf_cero = _asegurar_pdf_cero()
    print(f"      {pdf_cero}")

    pdf_noct = _ultimo_pdf_nocturno()
    print(f"[2/3] Control nocturno PDF: {pdf_noct}")

    print(f"[3/3] Enviando a: {', '.join(DESTINATARIOS)}")
    _enviar_pdfs([pdf_cero, pdf_noct])
    print("[OK] Correo enviado con ambos PDF adjuntos.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
