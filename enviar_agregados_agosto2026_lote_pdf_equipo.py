"""
Convierte a PDF, sube a Drive y envía correos — agregados agosto 2026.

Destinatarios correo: Juan, Diego, Aníbal.
Drive: Agente WES / wes-scripts / reports / <cliente> / ABREGADO

Uso:
  python enviar_agregados_agosto2026_lote_pdf_equipo.py
  python enviar_agregados_agosto2026_lote_pdf_equipo.py --sin-correo
"""

from __future__ import annotations

import argparse
import os
import shutil
import smtplib
import subprocess
import sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Optional

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
SUFIJO = "20260801_20260828"
PERIODO_TXT = "01/08/2026 al 28/08/2026"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_VISIBLE = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
    "anibal.aoperaciones@wes.cl",
]

# (carpeta ABREGADO, glob docx, subcarpeta Drive, asunto, detalle cuerpo)
ENVIOS = [
    (
        ROOT / "reports" / "Fundo_Zapallar" / "ABREGADO",
        f"Reporte_Agregado_*_{SUFIJO}.docx",
        "Fundo_Zapallar/ABREGADO",
        "Fundo Zapallar — Reporte agregado PDF — Agosto 2026",
        f"Adjunto el reporte agregado (PDF) de Fundo Zapallar para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "Inchcape" / "ABREGADO",
        f"Reporte_Agregado_*_{SUFIJO}.docx",
        "Inchcape/ABREGADO",
        "Inchcape — Reporte agregado PDF — Agosto 2026",
        f"Adjunto el reporte agregado (PDF) de Inchcape (ex DERCO) para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "Nido_de_Aguilas" / "ABREGADO",
        f"Reporte_Agregado_*_{SUFIJO}.docx",
        "Nido_de_Aguilas/ABREGADO",
        "Nido de Águilas — Reporte agregado PDF — Agosto 2026",
        f"Adjunto el reporte agregado (PDF) de Nido de Águilas para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "Lo_Valledor" / "ABREGADO",
        f"Reporte_Agregado_*_{SUFIJO}.docx",
        "Lo_Valledor/ABREGADO",
        "Lo Valledor — Reporte agregado PDF — Agosto 2026",
        "Adjunto el reporte agregado (PDF) de Lo Valledor (P1 + Barrio Norte) "
        f"para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "UDD" / "ABREGADO",
        f"Reporte_Agregado_*_{SUFIJO}.docx",
        "UDD/ABREGADO",
        "UDD — Reporte agregado PDF — Agosto 2026",
        f"Adjunto el reporte agregado (PDF) de UDD para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "Club_Providencia" / "ABREGADO",
        f"Reporte_Agregado_*_{SUFIJO}.docx",
        "Club_Providencia/ABREGADO",
        "Club Providencia — Reporte agregado PDF — Agosto 2026",
        f"Adjunto el reporte agregado (PDF) de Club Providencia para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "AGUNSA_Lampa" / "ABREGADO",
        f"Reporte_Agregado_*_{SUFIJO}.docx",
        "AGUNSA_Lampa/ABREGADO",
        "AGUNSA Lampa — Reporte agregado PDF — Agosto 2026",
        "Adjunto el reporte agregado (PDF) de AGUNSA Lampa (Depósito + módulos) "
        f"para el periodo {PERIODO_TXT}.\n",
    ),
    (
        ROOT / "reports" / "Agunsa_Intermodal" / "ABREGADO",
        f"Reporte_Agregado_*_{SUFIJO}.docx",
        "Agunsa_Intermodal/ABREGADO",
        "AGUNSA Intermodal San Antonio — Reporte agregado PDF — Agosto 2026",
        "Adjunto el reporte agregado (PDF) de AGUNSA Intermodal San Antonio "
        f"para el periodo {PERIODO_TXT}.\n",
    ),
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


def _docx_mas_reciente(base: Path, pattern: str) -> Path:
    candidatos = list(base.rglob(pattern)) if base.is_dir() else []
    if not candidatos:
        raise FileNotFoundError(f"No se encontró {pattern} en {base}")
    return max(candidatos, key=lambda p: p.stat().st_mtime)


def convertir_docx_a_pdf(docx_path: Path) -> Optional[Path]:
    pdf_path = docx_path.with_suffix(".pdf")
    if pdf_path.is_file() and pdf_path.stat().st_mtime >= docx_path.stat().st_mtime:
        return pdf_path

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run(
            [
                soffice,
                "--headless",
                "--nologo",
                "--nofirststartwizard",
                "--convert-to",
                "pdf",
                "--outdir",
                str(docx_path.parent),
                str(docx_path),
            ],
            check=True,
            capture_output=True,
            text=True,
        )
        return pdf_path if pdf_path.is_file() else None

    try:
        from docx2pdf import convert  # type: ignore

        convert(str(docx_path), str(pdf_path))
        return pdf_path if pdf_path.is_file() else None
    except Exception:
        return None


def _subir_drive(pdf: Path, docx: Path, subcarpeta: str) -> list[str]:
    links: list[str] = []
    try:
        from wes_google_drive import credenciales_configuradas, subir_a_drive
    except Exception as e:
        print(f"  [WARN] Drive no disponible: {e}")
        return links
    if not credenciales_configuradas():
        print("  [WARN] Faltan secretos GOOGLE_DRIVE_*")
        return links
    for path in (pdf, docx):
        try:
            info = subir_a_drive(path, subcarpeta=subcarpeta)
            link = info.get("web_view_link") or ""
            print(f"  [OK] Drive {path.name}: {link}")
            if link:
                links.append(f"{path.name}: {link}")
        except Exception as e:
            print(f"  [WARN] No se pudo subir {path.name}: {e}")
    return links


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sin-correo", action="store_true")
    args = parser.parse_args()

    pw = "" if args.sin_correo else _smtp_password()
    if not args.sin_correo and not pw:
        print("[WARN] Falta contraseña SMTP; se sube a Drive y no se envía correo.")

    ok = 0
    server = None
    if pw:
        server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO)
        server.starttls()
        server.login(SMTP_USUARIO, pw)

    try:
        for base, pattern, subcarpeta, asunto, cuerpo_detalle in ENVIOS:
            try:
                docx = _docx_mas_reciente(base, pattern)
            except FileNotFoundError as e:
                print(f"[ERROR] {e}", file=sys.stderr)
                continue

            print(f"\n[INFO] {asunto}")
            print(f"  Word: {docx}")
            pdf = convertir_docx_a_pdf(docx)
            if not pdf or not Path(pdf).is_file():
                print(f"[ERROR] No se pudo convertir a PDF: {docx}", file=sys.stderr)
                continue
            pdf = Path(pdf)
            print(f"  PDF: {pdf}")

            links = _subir_drive(pdf, docx, subcarpeta)

            if server:
                msg = MIMEMultipart()
                msg["From"] = SMTP_USUARIO
                msg["To"] = ", ".join(TO_VISIBLE)
                msg["Subject"] = asunto
                extra_drive = ""
                if links:
                    extra_drive = "También en Drive:\n" + "\n".join(links) + "\n\n"
                msg.attach(
                    MIMEText(
                        "Estimados Juan, Diego y Aníbal,\n\n"
                        f"{cuerpo_detalle}\n"
                        f"{extra_drive}"
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
                server.send_message(msg, to_addrs=TO_VISIBLE)
                print("  [OK] Correo enviado")

            ok += 1
    finally:
        if server:
            server.quit()

    print(f"\n[OK] {ok}/{len(ENVIOS)} reportes procesados.")
    return 0 if ok == len(ENVIOS) else 1


if __name__ == "__main__":
    raise SystemExit(main())
