"""
Envía el agregado Bupa / UPA Antofagasta (PDF + Word) a Juan, Diego y Aníbal.

Uso:
  python enviar_agregado_bupa_antofagasta_equipo.py
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
ABREGADO = ROOT / "reports" / "Bupa_Antofagasta" / "ABREGADO"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_RECIPIENTS = [
    "anibal.aoperaciones@wes.cl",
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
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


def _latest_pair() -> tuple[Path, Path]:
    dirs = sorted(
        [d for d in ABREGADO.glob("AGREGADO_*") if d.is_dir()],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    for d in dirs:
        pdfs = list(d.glob("Reporte_Agregado_BUPA_*.pdf"))
        docs = list(d.glob("Reporte_Agregado_BUPA_*.docx"))
        # Preferir el par con mismo stem
        for pdf in sorted(pdfs, key=lambda p: p.stat().st_mtime, reverse=True):
            docx = pdf.with_suffix(".docx")
            if docx.is_file():
                return docx, pdf
        if docs and pdfs:
            return docs[0], pdfs[0]
    raise FileNotFoundError(f"No hay PDF/Word en {ABREGADO}")


def main() -> int:
    try:
        docx_path, pdf_path = _latest_pair()
    except FileNotFoundError as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        return 1

    print(f"[INFO] Word: {docx_path}")
    print(f"[INFO] PDF:  {pdf_path}")

    pw = _smtp_password()
    if not pw:
        print(
            "[ERROR] Falta contraseña SMTP. Configurá el secreto "
            "WES_GMAIL_APP_PASSWORD (o WES_SMTP_PASSWORD) en Cursor → "
            "Cloud Agents → Secrets, o gmail_oauth/app_password.txt",
            file=sys.stderr,
        )
        return 1

    cuerpo = """Estimados Aníbal, Juan y Diego,

Adjunto el reporte agregado de Clínica Bupa / UPA Antofagasta
(nodos 000029-07, 000029-08, 000029-09 y 000029-10).

Resumen
-------
• Periodo: 23/07/2026 – 11/08/2026 (20 días civiles completos).
• Formato clásico: curvas diarias por punto (sin alertas).
• Punto verde el 29/07 solo en Medidor Principal Sanitaria (bajada por mejora de mantención).
• Comparativo semanal del periodo (4 semanas) — sin historial de 6 meses.
• Proyección cierre agosto vs ritmo previo:
  - Sin mejora: ~7.601 m³
  - Con mejora: ~3.037 m³
  - Ahorro proyectado: ~4.565 m³ / ~$12,6 M (tarifa factura).

También está en Drive:
Agente WES → wes-scripts → reports → Bupa_Antofagasta → ABREGADO
https://drive.google.com/drive/folders/1d0NWa3cpG-AWju3mV0Lw7LbW29nB6qd5

Adjunto: PDF y Word.

Saludos,
Agente IA WES
"""

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = (
        "Bupa / UPA Antofagasta — Reporte agregado 23/07–11/08/2026 "
        "(mejora 29/07 + proyección agosto)"
    )
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for path, subtype in (
        (pdf_path, "pdf"),
        (docx_path, "vnd.openxmlformats-officedocument.wordprocessingml.document"),
    ):
        with open(path, "rb") as f:
            part = MIMEApplication(f.read(), _subtype=subtype)
            part.add_header("Content-Disposition", "attachment", filename=path.name)
            msg.attach(part)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg)

    print(f"[OK] Enviado a: {', '.join(TO_RECIPIENTS)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
