"""
Envía un correo a Aníbal, Juan y Diego con los PDF existentes de COPEC
(01/04/2026–30/04/2026): todos los reportes por punto bajo REPORTE y,
si existe, el PDF agregado bajo ABREGADO. No borra carpetas ni regenera reportes.
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
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent
BASE = ROOT / "reports" / "COPEC"
REPORTE = BASE / "REPORTE"
ABREGADO = BASE / "ABREGADO"

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


def _pdf_por_punto() -> list[Path]:
    if not REPORTE.is_dir():
        return []
    return sorted(
        p for p in REPORTE.rglob("Reporte_*20260401_20260430.pdf") if p.is_file()
    )


def _pdf_agregado_mas_reciente() -> Path | None:
    if not ABREGADO.is_dir():
        return None
    cand = sorted(
        p
        for p in ABREGADO.rglob("*.pdf")
        if p.is_file() and "20260401_20260430" in p.name and "Agregado" in p.name
    )
    if not cand:
        cand = sorted(
            p for p in ABREGADO.rglob("Reporte_Agregado*.pdf") if p.is_file()
        )
        cand = [p for p in cand if "20260401_20260430" in p.name]
    if not cand:
        return None
    return max(cand, key=lambda p: p.stat().st_mtime)


def main() -> None:
    pw = _smtp_password()
    if not pw:
        print(
            "[ERROR] Falta contraseña SMTP (WES_GMAIL_APP_PASSWORD o gmail_oauth/app_password.txt).",
            file=sys.stderr,
        )
        sys.exit(1)

    individuales = _pdf_por_punto()
    if not individuales:
        print(f"[ERROR] No hay PDF por punto en {REPORTE}", file=sys.stderr)
        sys.exit(1)

    pdf_agg = _pdf_agregado_mas_reciente()
    adjuntos = sorted(individuales, key=lambda p: p.name)
    if pdf_agg:
        adjuntos = adjuntos + [pdf_agg]
        extra = " y el reporte agregado (PDF)."
    else:
        extra = (
            " No se encontró PDF de reporte agregado en ABREGADO para el mismo período; "
            "solo se adjuntan los puntos."
        )

    total = sum(p.stat().st_size for p in adjuntos)
    print(f"[INFO] Adjuntos: {len(adjuntos)} (~{total // 1024} KB)")

    nombres = "\n".join(f"- {p.name}" for p in adjuntos)
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = "COPEC — PDF abril 2026 (por punto" + (
        " + agregado)" if pdf_agg else ")"
    )
    msg.attach(
        MIMEText(
            f"""Estimados Aníbal, Juan y Diego,

Adjunto los reportes PDF de COPEC por punto de monitoreo (01/04/2026 al 30/04/2026){extra}

Archivos ({len(adjuntos)}):
{nombres}

Saludos cordiales,
Sistema WES
""",
            "plain",
            "utf-8",
        )
    )
    for p in adjuntos:
        # Evita colisiones (p. ej. dos puntos "Copec Costanera" con el mismo nombre de archivo)
        safe_folder = "".join(
            c if c.isalnum() or c in ("-", "_") else "_" for c in p.parent.name
        )
        fname = f"{safe_folder}__{p.name}"
        with open(p, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", "attachment", filename=fname)
            msg.attach(part)
        print(f"[OK] Adjunto: {fname}")

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg)

    print(f"[OK] Correo enviado a: {', '.join(TO_RECIPIENTS)}")


if __name__ == "__main__":
    main()
