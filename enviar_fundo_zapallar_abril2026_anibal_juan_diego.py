"""
Un solo correo: PDFs individuales + agregado Fundo Zapallar (abr-2026),
generados hoy (carpetas bajo REPORTE/ABREGADO con fecha 20260505).
Destinatarios: Aníbal, Juan y Diego.
"""

from __future__ import annotations

import os
import smtplib
import sys
from pathlib import Path
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent
BASE = ROOT / "reports" / "Fundo_Zapallar"
PERIOD_TAG = "20260401_20260430"
RUN_DATE_TAG = "20260505"  # envíos / corridas del 5-may-2026

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
        try:
            for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
                line = line.strip()
                if line and not line.startswith("#"):
                    return line.replace(" ", "").strip()
        except OSError:
            pass
    return ""


def _coleccionar_pdfs() -> list[Path]:
    reporte = BASE / "REPORTE"
    if not reporte.is_dir():
        raise SystemExit(f"No existe {reporte}")

    ind = sorted(
        p
        for p in reporte.rglob(f"*{PERIOD_TAG}.pdf")
        if RUN_DATE_TAG in p.as_posix()
    )
    if not ind:
        raise SystemExit(
            f"No se encontraron PDF individuales con {PERIOD_TAG} en rutas con {RUN_DATE_TAG}."
        )

    ag = BASE / "ABREGADO"
    agg_candidates = sorted(
        p
        for p in ag.rglob(f"Reporte_Agregado_Fundo_Zapallar_{PERIOD_TAG}.pdf")
        if p.is_file()
    )
    # Preferir carpeta de corrida de hoy; si no hay, usar el más reciente
    agg_today = [p for p in agg_candidates if RUN_DATE_TAG in p.as_posix()]
    agg = max(
        agg_today or agg_candidates,
        key=lambda p: p.stat().st_mtime,
    )

    seen: set[str] = set()
    out: list[Path] = []
    for p in sorted(ind, key=lambda x: x.name) + [agg]:
        key = str(p.resolve())
        if key not in seen:
            seen.add(key)
            out.append(p)
    return out


def main() -> None:
    pw = _smtp_password()
    if not pw:
        print(
            "[ERROR] Falta contraseña SMTP: define WES_GMAIL_APP_PASSWORD "
            "o crea gmail_oauth/app_password.txt",
            file=sys.stderr,
        )
        raise SystemExit(1)

    paths = _coleccionar_pdfs()
    total_bytes = sum(p.stat().st_size for p in paths)
    print(f"[INFO] Adjuntos: {len(paths)} archivo(s), ~{total_bytes // 1024} KB total")

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = (
        f"Fundo Zapallar — reportes PDF abril 2026 ({PERIOD_TAG.replace('_', '–')})"
    )

    nombres = "\n".join(f"- {p.name}" for p in paths)
    cuerpo = f"""Estimados Aníbal, Juan y Diego,

Se adjuntan en un solo correo los reportes PDF de Fundo Zapallar correspondientes al período 01/04/2026 al 30/04/2026 (corrida del {RUN_DATE_TAG[:4]}-{RUN_DATE_TAG[4:6]}-{RUN_DATE_TAG[6:8]}), incluyendo el reporte agregado.

Archivos:
{nombres}

Saludos cordiales,
Sistema WES
"""
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for p in paths:
        with open(p, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", "attachment", filename=p.name)
            msg.attach(part)
        print(f"[OK] Adjuntado: {p.name}")

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg)

    print(f"[OK] Correo enviado a: {', '.join(TO_RECIPIENTS)}")


if __name__ == "__main__":
    main()
