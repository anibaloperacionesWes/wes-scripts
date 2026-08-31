"""
Envía a Juan y Aníbal los informes de gestión hídrica (formato Zapallar)
del lote comercial de 8, periodo 01/08/2026–31/08/2026.

Uso:
  python enviar_gestion_hidrica_agosto2026_juan_anibal.py
  python enviar_gestion_hidrica_agosto2026_juan_anibal.py --dry-run
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

ASUNTO = "Informes de gestión hídrica — Agosto 2026 (1 al 31) — 8 clientes"

# (cliente, carpeta, estado, one-pager Drive, mensual Drive)
CLIENTES = [
    (
        "Fundo Zapallar",
        "Fundo_Zapallar",
        "En observación (32 % nocturno)",
        "https://drive.google.com/file/d/193B4N30jB-PBrO4LMZcJhm0DXL99q2J4/view?usp=drivesdk",
        "https://drive.google.com/file/d/1PYIb37zAar1aPN59dw1k28rFQs9qPdGQ/view?usp=drivesdk",
    ),
    (
        "Inchcape Quilicura",
        "Inchcape",
        "Bajo control (11 % nocturno WES)",
        "https://drive.google.com/file/d/1VmIZ3vj7qsZnLLtdGAITNZrQ1A66VE4m/view?usp=drivesdk",
        "https://drive.google.com/file/d/1JZm-pjIvnlX2vzFtS91ZJZwW6RjkELzY/view?usp=drivesdk",
    ),
    (
        "Nido de Águilas",
        "Nido_de_Aguilas",
        "En observación (20 % nocturno)",
        "https://drive.google.com/file/d/1wW7OkX3FEtunWR3NDZQDGsV1MRdocpnS/view?usp=drivesdk",
        "https://drive.google.com/file/d/1LF1zy-uFwW5AtLw0zofKqi1E5uma6Jpl/view?usp=drivesdk",
    ),
    (
        "Lo Valledor",
        "Lo_Valledor",
        "En observación (30 % nocturno)",
        "https://drive.google.com/file/d/17koaZKFdbKLJDbS28hINsQ8F_0ffQash/view?usp=drivesdk",
        "https://drive.google.com/file/d/1UgAvHzVxjpTT95IoxMGBOsEEZD0OpBre/view?usp=drivesdk",
    ),
    (
        "UDD",
        "UDD",
        "Bajo control (2 % nocturno)",
        "https://drive.google.com/file/d/1RY7Uy0Bk8I0BgilZNpBTibeN6AmWqArH/view?usp=drivesdk",
        "https://drive.google.com/file/d/1IywBjbFW3EdsYmOLBdPG-u6ncrPmOIjX/view?usp=drivesdk",
    ),
    (
        "Club Providencia",
        "Club_Providencia",
        "En observación (25 % nocturno)",
        "https://drive.google.com/file/d/1qfYn_7KQwDzhTlBieCmc79tijeGX-mO0/view?usp=drivesdk",
        "https://drive.google.com/file/d/1rTqYW3spT33not_-NnRlG-3CVL_eyrC8/view?usp=drivesdk",
    ),
    (
        "AGUNSA Lampa",
        "AGUNSA_Lampa",
        "En observación (20 % nocturno)",
        "https://drive.google.com/file/d/12APHY8HPdgcRWxRyQIxJ80rH07y22vMT/view?usp=drivesdk",
        "https://drive.google.com/file/d/1e83gOXStazek_zTQXB4eZQJDVEJcJCyQ/view?usp=drivesdk",
    ),
    (
        "AGUNSA Intermodal",
        "Agunsa_Intermodal",
        "En observación (24 % nocturno)",
        "https://drive.google.com/file/d/1fa4mLMjQopXqncb_BN53LxYl7DhCDiM1/view?usp=drivesdk",
        "https://drive.google.com/file/d/13wYH50xuPtKBAYLSHLzTsPpPB1L03XHW/view?usp=drivesdk",
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
    return "vxbynfpoehbweelj"


def _adjuntos() -> list[tuple[str, Path]]:
    out: list[tuple[str, Path]] = []
    for _label, folder, _estado, _op, _men in CLIENTES:
        base = ROOT / "reports" / folder / "GESTION_HIDRICA"
        ones = sorted(base.glob("One_Pager_*.pdf"))
        mens = sorted(base.glob("Informe_Mensual_*.pdf"))
        if not ones or not mens:
            raise FileNotFoundError(f"Faltan PDF en {base}")
        out.append((ones[-1].name, ones[-1]))
        out.append((mens[-1].name, mens[-1]))
    return out


def enviar(*, dry_run: bool = False) -> None:
    pdfs = _adjuntos()
    for nombre, path in pdfs:
        print(f"[OK] {nombre} ({path.stat().st_size // 1024} KB)")

    lista = "\n".join(f"  • {label}: {estado}" for label, _f, estado, _a, _b in CLIENTES)
    links = "\n".join(
        f"  • {label}\n      One-pager: {op}\n      Mensual: {men}"
        for label, _f, _e, op, men in CLIENTES
    )
    cuerpo = (
        "Estimados Juan y Aníbal,\n\n"
        "Adjunto los informes de gestión hídrica (formato ejecutivo Zapallar) "
        "del lote comercial, periodo 1 al 31 de agosto de 2026.\n\n"
        "Por cada cliente van dos PDF: one-pager y informe mensual.\n\n"
        f"{lista}\n\n"
        "Son 16 archivos PDF adjuntos en este mismo correo.\n\n"
        "También en Drive:\n"
        f"{links}\n\n"
        "Saludos cordiales,\n"
        "Sistema WES\n"
    )

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

    for nombre, pdf in pdfs:
        with open(pdf, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", "attachment", filename=nombre)
            msg.attach(part)

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
