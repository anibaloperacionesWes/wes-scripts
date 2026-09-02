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

    if filas:
        lista = "\n".join(
            f"  • {r['cliente']} — {r['punto']}: {r['revisar']} "
            f"({r['m3']} m³, noct {r['noct']}, {r['wow']} vs prev.)"
            for r in filas
        )
    else:
        lista = "  • Ningún punto cumple criterio de revisión esta semana."
    ok = ", ".join(sin_alerta) if sin_alerta else "—"
    cuerpo = (
        "Estimados Juan y Aníbal,\n\n"
        f"Consolidado de seguimiento SEMANAL ({periodo}).\n"
        "No reemplaza el informe de fin de mes: indica qué puntos revisar "
        "esta semana para atacar alzas, picos o nocturno anómalo antes del cierre.\n\n"
        "Puntos a revisar:\n"
        f"{lista}\n\n"
        f"Sin alerta: {ok}\n\n"
        + (f"También en Drive:\n  • {drive}\n\n" if drive else "")
        + "Saludos cordiales,\nSistema WES\n"
    )
    asunto = f"Consolidado semanal gestión hídrica — {periodo} — puntos a revisar"

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
