"""
Envía el consolidado semanal a Juan y Aníbal.

No enviar a Diego. Uso habitual: correr_consolidado_semanal_lunes.py
"""

from __future__ import annotations

import argparse
import os
import smtplib
import sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Sequence

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
TO_VISIBLE = [
    "juanlopez@wes.cl",
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


def enviar(
    pdf: Path,
    *,
    periodo: str,
    filas: Sequence[dict],
    sin_alerta: Sequence[str],
    drive: str = "",
    dry_run: bool = False,
) -> None:
    if not pdf.is_file():
        raise FileNotFoundError(f"No existe {pdf}")
    print(f"[OK] {pdf.name} ({pdf.stat().st_size // 1024} KB)")

    atacar = [r for r in filas if r.get("prio") == "ATENCIÓN"]
    n_seg = sum(1 for r in filas if r.get("prio") == "SEGUIMIENTO")
    mostrar = atacar[:6]
    if mostrar:
        lista = "\n".join(
            f"  • {r['cliente']} — {r['punto']} [{r.get('control') or 'SIN CONTROL'}]: "
            f"{r['revisar']} ({r['m3']} m³, noct {r['noct']}, {r['wow']} vs prev.)"
            for r in mostrar
        )
        if len(atacar) > 6:
            lista += f"\n  • … y {len(atacar) - 6} más en el PDF."
    else:
        lista = "  • Ningún punto para atacar esta semana."
    ok = ", ".join(sin_alerta) if sin_alerta else "—"
    cuerpo = (
        "Estimados Juan y Aníbal,\n\n"
        f"Semana {periodo}. Para atacar ahora — no reemplaza el cierre mensual.\n\n"
        f"Atacar esta semana ({len(atacar)}). "
        "CON CONTROL = tiene CPA/WES (si sube, el equipo no evitó el aumento). "
        "SIN CONTROL = no hay corte automático.\n\n"
        f"{lista}\n\n"
        f"El PDF adjunto trae el detalle y {n_seg} punto(s) en seguimiento "
        "(no urgentes; conviene mirarlos si hay tiempo).\n"
        f"Clientes sin alerta: {ok}\n\n"
        + (f"También en Drive:\n  • {drive}\n\n" if drive else "")
        + "Saludos cordiales,\nAgente IA WES\n"
    )
    asunto = f"WES · Semana {periodo} · {len(atacar)} punto(s) a atacar"

    if dry_run:
        print(f"\n[DRY-RUN] Asunto: {asunto}")
        print(f"[DRY-RUN] Para: {', '.join(TO_VISIBLE)}")
        print(cuerpo)
        return

    pw = _smtp_password()
    if not pw:
        raise RuntimeError("Falta contraseña SMTP (WES_GMAIL_APP_PASSWORD).")
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_VISIBLE)
    msg["Subject"] = asunto
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))
    with open(pdf, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=pdf.name)
        msg.attach(part)
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg, to_addrs=TO_VISIBLE)
    print(f"\n[OK] Correo enviado a {', '.join(TO_VISIBLE)}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--periodo", default="")
    parser.add_argument("--drive", default="")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    enviar(args.pdf, periodo=args.periodo, filas=[], sin_alerta=[], drive=args.drive, dry_run=args.dry_run)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
