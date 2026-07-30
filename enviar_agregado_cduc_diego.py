"""Envía el reporte Word agregado a Diego Carrasco (diegocarrasco@wes.cl).

Busca el último Reporte_Agregado_*.docx en reports/<empresa>/ABREGADO; si no hay, genera
el agregado para el periodo indicado (por defecto 01/01/2026 – 06/04/2026).

Uso:
  python enviar_agregado_cduc_diego.py
  python enviar_agregado_cduc_diego.py --empresa agunsa
  python enviar_agregado_cduc_diego.py --regenerar
"""

from __future__ import annotations

import argparse
import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from generar_reporte_word import (
    convertir_word_a_pdf,
    generate_aggregated_report,
    get_company_name,
)

NODOS_CDUC_TODOS = [
    "000021-01",
    "000021-02",
    "000021-03",
    "000021-04",
    "000021-05",
    "000021-07",
    "000021-08",
]
NODOS_AGUNSA = [
    "000020-01",
    "000020-02",
    "000020-03",
    "000020-04",
    "000020-05",
]

EMPRESAS = {
    "cduc": ("000021", NODOS_CDUC_TODOS, "CDUC"),
    "agunsa": ("000020", NODOS_AGUNSA, "AGUNSA"),
}

DESTINATARIO = "diegocarrasco@wes.cl"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
SMTP_PASSWORD = os.environ.get("WES_SMTP_PASSWORD") or "vxbynfpoehbweelj"

DEFAULT_START = "01/01/2026"
DEFAULT_END = "06/04/2026"


def _latest_agregado(reports_dir: Path, carpeta_empresa: str) -> Path | None:
    base = reports_dir / carpeta_empresa / "ABREGADO"
    if not base.exists():
        return None
    docs = sorted(base.rglob("Reporte_Agregado_*.docx"))
    return docs[-1] if docs else None


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument(
        "--empresa",
        choices=("cduc", "agunsa"),
        default="cduc",
        help="Cliente cuyo agregado se envía (default: cduc).",
    )
    p.add_argument("--regenerar", action="store_true", help="Fuerza generación del agregado antes de enviar.")
    p.add_argument("--start-date", default=DEFAULT_START)
    p.add_argument("--end-date", default=DEFAULT_END)
    args = p.parse_args()

    company_id, nodos, carpeta_empresa = EMPRESAS[args.empresa]

    root = Path(__file__).resolve().parent
    reports_dir = root / "reports"

    reporte_path: Path | None = None
    if not args.regenerar:
        reporte_path = _latest_agregado(reports_dir, carpeta_empresa)

    if reporte_path is None or args.regenerar:
        print(f"[INFO] Generando reporte agregado {carpeta_empresa}...")
        reporte_path = generate_aggregated_report(
            company_id,
            list(nodos),
            args.start_date,
            args.end_date,
            output_dir=str(reports_dir),
            generate_ppt=False,
            parallel_node_fetch=True,
        )

    if not reporte_path or not reporte_path.exists():
        print("[ERROR] No hay archivo agregado para enviar.")
        return 1

    company_name = get_company_name(company_id)
    pdf_path = convertir_word_a_pdf(reporte_path)
    adj_path = pdf_path if pdf_path and pdf_path.exists() else reporte_path
    es_pdf = adj_path.suffix.lower() == ".pdf"

    print(f"[INFO] Adjunto: {adj_path}")

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = DESTINATARIO
    msg["Subject"] = f"Reporte agregado {company_name} — {args.start_date} a {args.end_date}"

    cuerpo = f"""Estimado Diego,

Adjunto el reporte agregado de consumo y alertas de {company_name} para el periodo
{args.start_date} al {args.end_date}.

Saludos,
Agente IA WES
"""
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    with open(adj_path, "rb") as f:
        subtype = "pdf" if es_pdf else "docx"
        adj = MIMEApplication(f.read(), _subtype=subtype)
        adj.add_header("Content-Disposition", "attachment", filename=adj_path.name)
        msg.attach(adj)

    try:
        print(f"[INFO] Enviando a {DESTINATARIO}...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.sendmail(SMTP_USUARIO, [DESTINATARIO], msg.as_string())
        print("[OK] Correo enviado.")
        return 0
    except Exception as e:
        print(f"[ERROR] Envío: {e}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
