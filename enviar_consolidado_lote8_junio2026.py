"""
Envía un correo consolidado con los 8 PDFs del lote junio 2026 (formato extendido).
Destinatarios: Juan, Diego, Aníbal.

Uso:
  python enviar_consolidado_lote8_junio2026.py
  python enviar_consolidado_lote8_junio2026.py --limpiar   # solo elimina duplicados
  python enviar_consolidado_lote8_junio2026.py --solo-enviar
"""

from __future__ import annotations

import argparse
import os
import shutil
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
SUFIJO = "20260601_20260630"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_VISIBLE = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
    "anibal.aoperaciones@wes.cl",
]

# Carpetas buenas (formato extendido + tabla WES) — junio 2026
CLIENTES = [
    {
        "label": "Nido de Águilas",
        "base": ROOT / "reports" / "Nido_de_Aguilas" / "ABREGADO",
        "keeper": "AGREGADO_20260703_0051",
        "docx": f"Reporte_Agregado_Nido_de_Aguilas_{SUFIJO}.docx",
        "pdf_name": f"Reporte_Agregado_Nido_de_Aguilas_{SUFIJO}.pdf",
    },
    {
        "label": "AGUNSA Lampa",
        "base": ROOT / "reports" / "AGUNSA_Lampa" / "ABREGADO",
        "keeper": "AGREGADO_20260703_0338",
        "docx": f"Reporte_Agregado_AGUNSA_{SUFIJO}.docx",
        "pdf_name": f"Reporte_Agregado_AGUNSA_Lampa_{SUFIJO}.pdf",
    },
    {
        "label": "CDUC",
        "base": ROOT / "reports" / "CDUC" / "ABREGADO",
        "keeper": "AGREGADO_20260703_0120",
        "docx": f"Reporte_Agregado_CDUC_{SUFIJO}.docx",
        "pdf_name": f"Reporte_Agregado_CDUC_{SUFIJO}.pdf",
    },
    {
        "label": "Fundo Zapallar",
        "base": ROOT / "reports" / "Fundo_Zapallar" / "ABREGADO",
        "keeper": "AGREGADO_20260703_0211",
        "docx": f"Reporte_Agregado_Fundo_Zapallar_{SUFIJO}.docx",
        "pdf_name": f"Reporte_Agregado_Fundo_Zapallar_{SUFIJO}.pdf",
    },
    {
        "label": "La Florida",
        "base": ROOT / "reports" / "La_Florida" / "ABREGADO",
        "keeper": "AGREGADO_20260703_0244",
        "docx": f"Reporte_Agregado_La_Florida_{SUFIJO}.docx",
        "pdf_name": f"Reporte_Agregado_La_Florida_{SUFIJO}.pdf",
    },
    {
        "label": "Corporación Providencia (colegios)",
        "base": ROOT / "reports" / "Providencia" / "ABREGADO",
        "keeper": "AGREGADO_20260703_0303",
        "docx": f"Reporte_Agregado_Providencia_{SUFIJO}.docx",
        "pdf_name": f"Reporte_Agregado_Providencia_{SUFIJO}.pdf",
    },
    {
        "label": "DERCO",
        "base": ROOT / "reports" / "DERCO" / "ABREGADO",
        "keeper": "AGREGADO_20260703_0312",
        "docx": f"Reporte_Agregado_DERCO_{SUFIJO}.docx",
        "pdf_name": f"Reporte_Agregado_DERCO_{SUFIJO}.pdf",
    },
    {
        "label": "AGUNSA Intermodal San Antonio",
        "base": ROOT / "reports" / "Agunsa_Intermodal" / "ABREGADO",
        "keeper": "AGREGADO_20260703_1502",
        "docx": f"Reporte_Agregado_AGUNSA_{SUFIJO}.docx",
        "pdf_name": f"Reporte_Agregado_AGUNSA_Intermodal_{SUFIJO}.pdf",
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
        for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                return line.replace(" ", "").strip()
    return ""


def _resolver_pdf(cliente: dict) -> Path:
    from control_nocturno import convertir_docx_a_pdf

    carpeta = cliente["base"] / cliente["keeper"]
    docx = carpeta / cliente["docx"]
    if not docx.is_file():
        raise FileNotFoundError(f"No se encontró {docx}")
    pdf = carpeta / cliente["docx"].replace(".docx", ".pdf")
    if not pdf.is_file():
        pdf = Path(convertir_docx_a_pdf(docx))
    if not pdf.is_file():
        raise FileNotFoundError(f"No se pudo obtener PDF para {cliente['label']}")
    return pdf


def limpiar_duplicados_junio2026() -> int:
    eliminados = 0
    for c in CLIENTES:
        base = c["base"]
        keeper = c["keeper"]
        docx_name = c["docx"]
        if not base.is_dir():
            print(f"[WARN] Sin carpeta: {base}")
            continue
        for sub in sorted(base.iterdir()):
            if not sub.is_dir() or not sub.name.startswith("AGREGADO_"):
                continue
            if sub.name == keeper:
                continue
            docx = sub / docx_name
            if docx.is_file():
                print(f"[DEL] {c['label']}: {sub.name}")
                shutil.rmtree(sub)
                eliminados += 1
    return eliminados


def enviar_consolidado() -> None:
    pdfs: list[tuple[str, Path]] = []
    for c in CLIENTES:
        pdf = _resolver_pdf(c)
        print(f"[OK] {c['label']}: {pdf}")
        pdfs.append((c["pdf_name"], pdf))

    lista = "\n".join(f"  • {c['label']}" for c in CLIENTES)
    cuerpo = (
        "Estimados Juan, Diego y Aníbal,\n\n"
        "Adjunto el consolidado de reportes agregados en PDF para el periodo "
        "01/06/2026 al 30/06/2026 (formato extendido WES):\n\n"
        f"{lista}\n\n"
        "Son 8 archivos PDF adjuntos en este mismo correo.\n\n"
        "Saludos cordiales,\n"
        "Sistema WES\n"
    )

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_VISIBLE)
    msg["Subject"] = "Consolidado — Reportes agregados Junio 2026 (8 clientes)"
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for nombre_adj, pdf in pdfs:
        with open(pdf, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", "attachment", filename=nombre_adj)
            msg.attach(part)

    pw = _smtp_password()
    if not pw:
        raise RuntimeError("Falta contraseña SMTP.")

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=TO_VISIBLE)

    print(f"\n[OK] Correo consolidado enviado a: {', '.join(TO_VISIBLE)}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--limpiar", action="store_true", help="Solo eliminar duplicados junio 2026")
    parser.add_argument("--solo-enviar", action="store_true", help="Solo enviar correo")
    args = parser.parse_args()

    if args.limpiar and not args.solo_enviar:
        n = limpiar_duplicados_junio2026()
        print(f"\n[OK] {n} carpeta(s) duplicada(s) eliminada(s).")
        return 0

    if args.solo_enviar:
        enviar_consolidado()
        return 0

    enviar_consolidado()
    n = limpiar_duplicados_junio2026()
    print(f"\n[OK] {n} carpeta(s) duplicada(s) eliminada(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
