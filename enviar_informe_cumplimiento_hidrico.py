"""
Envía el informe de cumplimiento hídrico (PDF + Word + CSV) a Mauricio y Aníbal.

Uso:
  python enviar_informe_cumplimiento_hidrico.py
  python enviar_informe_cumplimiento_hidrico.py --dir reports/control_nocturno/cumplimiento_hidrico_YYYYMMDD_...
  python enviar_informe_cumplimiento_hidrico.py --drive-pdf URL --drive-docx URL --drive-csv URL
"""

from __future__ import annotations

import argparse
import csv
import os
import smtplib
import sys
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parent
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
CHILE = ZoneInfo("America/Santiago")

DESTINATARIOS = [
    "mauricioorellana@wes.cl",
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
    return "vxbynfpoehbweelj"


def _ultimo_dir() -> Path:
    base = ROOT / "reports" / "control_nocturno"
    dirs = sorted(
        (p for p in base.glob("cumplimiento_hidrico_*") if p.is_dir()),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not dirs:
        raise FileNotFoundError(f"No hay carpetas cumplimiento_hidrico_* en {base}")
    return dirs[0]


def _archivos(out_dir: Path) -> list[Path]:
    pdfs = list(out_dir.glob("Informe_Cumplimiento_Hidrico_*.pdf"))
    docs = list(out_dir.glob("Informe_Cumplimiento_Hidrico_*.docx"))
    csvs = list(out_dir.glob("cumplimiento_hidrico_*.csv"))
    if not pdfs:
        raise FileNotFoundError(f"No hay PDF en {out_dir}")
    files = [pdfs[0]]
    if docs:
        files.append(docs[0])
    if csvs:
        files.append(csvs[0])
    return files


def _leer_filas(out_dir: Path) -> list[dict[str, str]]:
    csvs = list(out_dir.glob("cumplimiento_hidrico_*.csv"))
    if not csvs:
        return []
    with csvs[0].open(encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f, delimiter=";"))


def _fila(rows: list[dict[str, str]], node_id: str) -> dict[str, str] | None:
    for row in rows:
        if row.get("node_id") == node_id:
            return row
    return None


def _linea_especial(row: dict[str, str] | None, etiqueta: str, node_id: str) -> str:
    if not row:
        return f"- {etiqueta} ({node_id}): sin datos en el CSV."
    cumple = "Cumple" if row.get("cumple") == "SI" else "No cumple"
    obs = (row.get("observacion") or "").strip().rstrip(".")
    cerrado = (row.get("tiempo_cerrado") or "").strip()
    retraso = (row.get("retraso_corte") or "").strip()
    bits: list[str] = []
    if obs:
        bits.append(obs)
    if retraso and retraso not in {"—", "-"} and retraso not in obs:
        bits.append(retraso)
    if (
        cerrado
        and cerrado not in {"—", "-"}
        and "Abierto" not in cerrado
        and cerrado not in obs
    ):
        bits.append(cerrado)
    if cumple.lower() not in (obs or "").lower():
        bits.append(cumple)
    return f"- {etiqueta} ({node_id}): {'; '.join(bits)}."


def _cuerpo(
    fecha: str,
    rows: list[dict[str, str]],
    drive_pdf: str,
    drive_docx: str,
    drive_csv: str,
) -> str:
    total = len(rows)
    cumplen = sum(1 for r in rows if r.get("cumple") == "SI")
    no_cumplen = total - cumplen
    nombres_no = [r.get("punto") or r.get("node_id") or "" for r in rows if r.get("cumple") != "SI"]
    lista_no = ", ".join(nombres_no) if nombres_no else "ninguno"

    derco = _linea_especial(
        _fila(rows, "000012-06"), "Derco Matriz Principal", "000012-06"
    )
    gym = _linea_especial(_fila(rows, "000017-05"), "GYM Renca", "000017-05")
    icco = _linea_especial(_fila(rows, "000017-08"), "ICCO Renca", "000017-08")

    drive_lines = []
    if drive_pdf:
        drive_lines.append(f"- PDF: {drive_pdf}")
    if drive_docx:
        drive_lines.append(f"- Word: {drive_docx}")
    if drive_csv:
        drive_lines.append(f"- CSV: {drive_csv}")
    bloque_drive = ""
    if drive_lines:
        bloque_drive = "\nTambién está en Drive:\n" + "\n".join(drive_lines) + "\n"

    return f"""Estimados Mauricio y Aníbal,

Adjunto el informe de cumplimiento hídrico del {fecha}, cruzando la planilla de
horarios de habilitación de agua con el caudal horario WES.

Resumen:
- {total} puntos revisados. {cumplen} cumplen y {no_cumplen} no cumplen (se toleran hasta 2 h de retraso de válvula).
{derco}
{gym}
{icco}

Puntos que no cumplen: {lista_no}.
{bloque_drive}
Saludos,
Sistema WES
"""


def main() -> int:
    if sys.platform == "win32":
        for s in (sys.stdout, sys.stderr):
            try:
                s.reconfigure(encoding="utf-8")
            except Exception:
                pass

    ap = argparse.ArgumentParser(description="Enviar informe de cumplimiento hídrico")
    ap.add_argument("--dir", type=Path, default=None)
    ap.add_argument("--drive-pdf", default="")
    ap.add_argument("--drive-docx", default="")
    ap.add_argument("--drive-csv", default="")
    args = ap.parse_args()

    out_dir = args.dir.resolve() if args.dir else _ultimo_dir()
    adjuntos = _archivos(out_dir)
    rows = _leer_filas(out_dir)
    fecha = datetime.now(CHILE).strftime("%d-%m-%Y")

    asunto = f"Informe de cumplimiento hídrico (habilitación / corte) — {fecha}"
    cuerpo = _cuerpo(fecha, rows, args.drive_pdf, args.drive_docx, args.drive_csv)

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(DESTINATARIOS)
    msg["Subject"] = asunto
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for path in adjuntos:
        subtype = {
            ".pdf": "pdf",
            ".docx": "vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".csv": "csv",
        }.get(path.suffix.lower(), "octet-stream")
        with path.open("rb") as f:
            part = MIMEApplication(f.read(), _subtype=subtype)
            part.add_header("Content-Disposition", "attachment", filename=path.name)
            msg.attach(part)

    print(f"[INFO] Enviando desde {SMTP_USUARIO}")
    print(f"[INFO] Para: {', '.join(DESTINATARIOS)}")
    print(f"[INFO] Adjuntos: {', '.join(p.name for p in adjuntos)}")
    print("--- cuerpo ---")
    print(cuerpo)
    print("---")
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, _smtp_password())
        server.send_message(msg, to_addrs=DESTINATARIOS)
    print("[OK] Correo enviado.")
    for d in DESTINATARIOS:
        print(f"  - {d}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
