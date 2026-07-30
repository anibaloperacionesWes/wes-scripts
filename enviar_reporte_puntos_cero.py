"""
Convierte el último reporte de puntos en cero a PDF y lo envía por correo.

- No se conecta a la API: solo toma el .docx existente en ``reporte en cero/``.
- Por defecto envía a Aníbal (operaciones).
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

# Configuración SMTP (Gmail app password).
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
SMTP_PASSWORD = (
    os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
    or os.environ.get("WES_SMTP_PASSWORD", "").strip()
    or "vxbynfpoehbweelj"
)

from wes_paths import reporte_cero_dir

REPORTE_DIR = reporte_cero_dir()


def _find_ultimo_docx(reporte_dir: Path) -> Path:
    if not reporte_dir.exists():
        raise FileNotFoundError(f"No existe carpeta: {reporte_dir}")
    archivos_docx = [f for f in reporte_dir.glob("*.docx") if f.is_file() and not f.name.startswith("~$")]
    if not archivos_docx:
        raise FileNotFoundError(f"No se encontraron .docx en {reporte_dir}")
    archivos_docx.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return archivos_docx[0]


def enviar_correo(pdf_path: Path, *, to_emails: list[str], cc_emails: list[str] | None = None) -> bool:
    """
    Envía el reporte PDF por correo electrónico.
    """
    if not pdf_path or not pdf_path.exists():
        print(f"[ERROR] El archivo PDF no existe: {pdf_path}")
        return False

    to_emails = [e.strip() for e in to_emails if e and str(e).strip()]
    if not to_emails:
        print("[ERROR] No hay destinatarios en Para (To).")
        return False

    cc_emails = cc_emails or []
    print(f"[2/3] Preparando envío de correo...")
    
    try:
        # Crear mensaje
        msg = MIMEMultipart()
        msg["From"] = SMTP_USUARIO
        msg["To"] = ", ".join(to_emails)
        if cc_emails:
            msg["Cc"] = ", ".join(cc_emails)
        
        # Crear asunto
        fecha = datetime.now().strftime("%d-%m-%Y")
        asunto = f"Reporte de Puntos en Cero — {fecha}"
        msg["Subject"] = asunto
        
        # Crear cuerpo del mensaje
        cuerpo = (
            "Estimado equipo,\n\n"
            "Se adjunta el reporte de puntos en cero (PDF).\n\n"
            "Saludos cordiales,\n"
            "Sistema WES\n"
        )
        
        msg.attach(MIMEText(cuerpo, "plain", "utf-8"))
        
        # Adjuntar PDF
        with open(pdf_path, "rb") as f:
            adjunto = MIMEApplication(f.read(), _subtype="pdf")
            adjunto.add_header(
                "Content-Disposition",
                "attachment",
                filename=pdf_path.name,
            )
            msg.attach(adjunto)
        
        # Enviar correo
        print(f"[3/3] Enviando correo...")
        destinatarios_todos = list(to_emails) + list(cc_emails)
        
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg, to_addrs=destinatarios_todos)
        
        print("[OK] Correo enviado exitosamente a:")
        for t in to_emails:
            print(f"  - {t}")
        for cc in cc_emails:
            print(f"  - {cc} (copia)")
        
        return True
    except Exception as e:
        print(f"[ERROR] Error al enviar correo: {e}")
        import traceback
        traceback.print_exc()
        return False


def main() -> int:
    ap = argparse.ArgumentParser(description="Enviar último Reporte Puntos en Cero (PDF)")
    ap.add_argument(
        "--to",
        action="append",
        default=[],
        dest="to_emails",
        metavar="EMAIL",
        help="Destinatario en Para (repetible). Por defecto: anibal.aoperaciones@wes.cl",
    )
    ap.add_argument("--cc", action="append", default=[], help="Agregar destinatario en copia (repetible)")
    ap.add_argument("--reporte-dir", type=Path, default=REPORTE_DIR, help="Carpeta donde buscar el .docx")
    args = ap.parse_args()

    print("=" * 70)
    print("  ENVÍO DE REPORTE PUNTOS EN CERO")
    print("=" * 70)
    print("")

    try:
        reporte_path = _find_ultimo_docx(args.reporte_dir)
    except FileNotFoundError as e:
        print(f"[ERROR] {e}")
        return 1

    print(f"[OK] Reporte encontrado: {reporte_path}")
    pdf_path = convertir_word_a_pdf(reporte_path)
    if not pdf_path or not pdf_path.exists():
        print("[ERROR] No se pudo convertir el reporte a PDF.")
        return 1

    to_list = args.to_emails if args.to_emails else ["anibal.aoperaciones@wes.cl"]
    ok = enviar_correo(pdf_path, to_emails=to_list, cc_emails=args.cc)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

