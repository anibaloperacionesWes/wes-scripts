"""
Envía el informe de cumplimiento hídrico (PDF + Word + CSV) a Mauricio y Aníbal.

Uso:
  python enviar_informe_cumplimiento_hidrico.py
  python enviar_informe_cumplimiento_hidrico.py --dir reports/control_nocturno/cumplimiento_hidrico_YYYYMMDD_...
"""

from __future__ import annotations

import argparse
import os
import smtplib
import sys
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

DESTINATARIOS = [
    "mauricioorellana@wes.cl",
    "anibal.aoperaciones@wes.cl",
]

DRIVE_PDF = "https://drive.google.com/file/d/1lb7-vo_ECkDhWx3tP7IzOOiSJdJFPWeG/view?usp=drivesdk"
DRIVE_DOCX = "https://docs.google.com/document/d/1_3OH9SC6YD2_8zeGgz1SJKQfVao_qFvq/edit?usp=drivesdk"
DRIVE_CSV = "https://drive.google.com/file/d/1L7PhK7Ls0Ivpb1LtX-q9a1DCaOXdWgwa/view?usp=drivesdk"


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


def main() -> int:
    if sys.platform == "win32":
        for s in (sys.stdout, sys.stderr):
            try:
                s.reconfigure(encoding="utf-8")
            except Exception:
                pass

    ap = argparse.ArgumentParser(description="Enviar informe de cumplimiento hídrico")
    ap.add_argument("--dir", type=Path, default=None)
    args = ap.parse_args()

    out_dir = args.dir.resolve() if args.dir else _ultimo_dir()
    adjuntos = _archivos(out_dir)
    fecha = datetime.now().strftime("%d-%m-%Y")

    asunto = f"Informe de cumplimiento hídrico (habilitación / corte) — {fecha}"
    cuerpo = f"""Estimados Mauricio y Aníbal,

Adjunto el informe de cumplimiento hídrico del {fecha}, cruzando la planilla de
horarios de habilitación de agua con el caudal horario WES.

Resumen:
- 32 puntos revisados. 22 cumplen y 10 no cumplen (se toleran hasta 2 h de retraso de válvula).
- Derco Matriz Principal (000012-06): mínimo nocturno de guardias 0,40 m³/h vs histórico 0,43
  (límite +25 % = 0,54). Cumple.
- GYM Renca (000017-05): mínimo nocturno 0,21 m³/h vs histórico 0,16 (límite +25 % = 0,20).
  No cumple: supera el 25 % del histórico (máx. 0,28 m³/h).
- ICCO Renca (000017-08): corte programado 20:00, cerró 22:00 (retraso 2 h) y se mantuvo
  cerrado 8 h (22:00–06:00). Hoy está habilitado (ALTA hasta 20:00). Cumple.

Puntos que no cumplen: Raimundo Tupper, Antonio Hermida, Carlos Fernández Peña,
Erasmo Escala, Matilde Huici Navas, Lo Valledor P1, Carmela Carvajal, Lastarria,
ICCP y GYM Renca.

También está en Drive:
- PDF: {DRIVE_PDF}
- Word: {DRIVE_DOCX}
- CSV: {DRIVE_CSV}

Saludos,
Sistema WES
"""

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
