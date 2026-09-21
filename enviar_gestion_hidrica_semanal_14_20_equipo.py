"""
Envía a Juan, Diego y Aníbal el listado con links del reporte SEMANAL
de gestión hídrica (one-pager + extendido), 14–20 Sep 2026.

Uso:
  python enviar_gestion_hidrica_semanal_14_20_equipo.py
  python enviar_gestion_hidrica_semanal_14_20_equipo.py --dry-run
"""

from __future__ import annotations

import argparse
import json
import os
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
RESUMEN = (
    ROOT
    / "reports"
    / "CONSOLIDADO"
    / "SEMANAL"
    / "One_Pagers_Semanal_20260914_20260920.json"
)

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_VISIBLE = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
    "anibal.aoperaciones@wes.cl",
]

ASUNTO = "WES · Reporte semanal de gestión hídrica · 14 al 20 de septiembre de 2026"

PRETTY = {
    "BAJO CONTROL": "Bajo control",
    "EN OBSERVACIÓN": "En observación",
    "REQUIERE ATENCIÓN": "Requiere atención",
    "CRÍTICO": "Crítico",
}


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


def _cuerpo(filas: list[dict]) -> str:
    bloques = []
    for r in filas:
        estado = PRETTY.get(r.get("clasificacion") or "", r.get("clasificacion") or "—")
        extra = []
        if r.get("kpi"):
            extra.append(r["kpi"])
        if r.get("noct"):
            extra.append(f"nocturno {r['noct']}")
        detalle = f" · {' · '.join(extra)}" if extra else ""
        bloques.append(
            f"  • {r['cliente']} — {estado}{detalle}\n"
            f"      One-pager: {r.get('drive_one') or '—'}\n"
            f"      Extendido: {r.get('drive_ext') or '—'}"
        )
    lista = "\n".join(bloques)
    return (
        "Estimados Juan, Diego y Aníbal,\n\n"
        "Les comparto el REPORTE SEMANAL de gestión hídrica "
        "(lunes 14 al domingo 20 de septiembre de 2026), comparado con la "
        "semana previa (7 al 13).\n\n"
        "Es el mismo formato de fin de mes (one-pager + informe extendido), "
        "pero de esta semana. No reemplaza el informe de cierre mensual.\n\n"
        "Por cada cliente van dos PDF en Drive:\n"
        f"{lista}\n\n"
        "Los archivos están en Agente WES / wes-scripts / reports / "
        "<cliente> / GESTION_HIDRICA / SEMANAL.\n\n"
        "Saludos cordiales,\n"
        "Agente IA WES\n"
    )


def enviar(*, dry_run: bool = False) -> None:
    data = json.loads(RESUMEN.read_text(encoding="utf-8"))
    filas = data.get("ok") or []
    if not filas:
        raise RuntimeError(f"Sin clientes en {RESUMEN}")
    cuerpo = _cuerpo(filas)
    print(f"[INFO] {len(filas)} cliente(s) en el listado")

    if dry_run:
        print(f"\n[DRY-RUN] Asunto: {ASUNTO}")
        print(f"[DRY-RUN] Para: {', '.join(TO_VISIBLE)}")
        print(cuerpo)
        return

    pw = _smtp_password()
    if not pw:
        raise RuntimeError("Falta contraseña SMTP (WES_GMAIL_APP_PASSWORD).")

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_VISIBLE)
    msg["Subject"] = ASUNTO
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=TO_VISIBLE)

    print(f"\n[OK] Correo enviado a {', '.join(TO_VISIBLE)}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    try:
        enviar(dry_run=args.dry_run)
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
