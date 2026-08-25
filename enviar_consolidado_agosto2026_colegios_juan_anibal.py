"""
Envía un correo consolidado con los 4 PDFs de colegios (agosto 2026).
Destinatarios: Juan y Aníbal. Solicita aprobación.

Uso:
  python enviar_consolidado_agosto2026_colegios_juan_anibal.py
  python enviar_consolidado_agosto2026_colegios_juan_anibal.py --dry-run
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

# (label, carpeta ABREGADO, glob pdf)
CLIENTES = [
    (
        "Renca (colegios)",
        ROOT / "reports" / "Renca" / "ABREGADO",
        "Reporte_Agregado_*_20260801_20260824.pdf",
        "01/08/2026 al 24/08/2026 — Lo Velásquez, Cumbre Pte., ICCO",
    ),
    (
        "La Florida",
        ROOT / "reports" / "La_Florida" / "ABREGADO",
        "Reporte_Agregado_*_20260801_20260824.pdf",
        "01/08/2026 al 24/08/2026 — Liceo Alto Cordillera",
    ),
    (
        "La Reina",
        ROOT / "reports" / "La_Reina" / "ABREGADO",
        "Reporte_Agregado_*_20260805_20260824.pdf",
        "05/08/2026 al 24/08/2026 — Eugenio María De Hostos "
        "(consumo válido desde el 05/08; se excluyen 01–04/08)",
    ),
    (
        "CORMUP Peñalolén",
        ROOT / "reports" / "CORMUP" / "ABREGADO",
        "Reporte_Agregado_*_20260801_20260824.pdf",
        "01/08/2026 al 24/08/2026 — 14 establecimientos",
    ),
]

DRIVE_LINKS = [
    ("Renca", "https://drive.google.com/file/d/102Tx5tnlQxqSwPw4fTE0JNwBWN0GvU9s/view?usp=drivesdk"),
    ("La Florida", "https://drive.google.com/file/d/1WoCcYyF_bpV4Q3TYlsPt2sF1Btv5XmXT/view?usp=drivesdk"),
    ("La Reina", "https://drive.google.com/file/d/1V1OcLsKu4IZUStcUd35ViKWwxTC-AWQY/view?usp=drivesdk"),
    ("CORMUP", "https://drive.google.com/file/d/1P93CBFoycDebfFHG6A6B1syWZzgrs-bn/view?usp=drivesdk"),
]

ASUNTO = "Aprobación — Reportes agregados colegios Agosto 2026 (Renca, La Florida, La Reina, CORMUP)"


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
    # Misma clave de aplicación usada en el resto de scripts de envío del repo.
    return "vxbynfpoehbweelj"


def _pdf_mas_reciente(base: Path, pattern: str) -> Path:
    candidatos = list(base.rglob(pattern)) if base.is_dir() else []
    if not candidatos:
        raise FileNotFoundError(f"No se encontró {pattern} en {base}")
    return max(candidatos, key=lambda p: p.stat().st_mtime)


def enviar_consolidado(*, dry_run: bool = False) -> None:
    pdfs: list[tuple[str, Path]] = []
    for label, base, pattern, _detalle in CLIENTES:
        pdf = _pdf_mas_reciente(base, pattern)
        print(f"[OK] {label}: {pdf}")
        pdfs.append((pdf.name, pdf))

    lista = "\n".join(f"  • {label}: {detalle}" for label, _base, _pat, detalle in CLIENTES)
    links = "\n".join(f"  • {label}: {url}" for label, url in DRIVE_LINKS)
    cuerpo = (
        "Estimados Juan y Aníbal,\n\n"
        "Adjunto los reportes agregados de colegios (PDF) de fin de mes para su "
        "revisión y aprobación:\n\n"
        f"{lista}\n\n"
        "Los gráficos van sin puntos rojos de alertas y sin sección de día de "
        "mayor consumo, igual que el lote comercial de agosto.\n\n"
        "Quedo atento a su aprobación para el cierre de fin de mes.\n\n"
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
        raise RuntimeError(
            "Falta contraseña SMTP. Configurá WES_GMAIL_APP_PASSWORD o WES_SMTP_PASSWORD."
        )

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

    print(f"\n[OK] Correo consolidado enviado a {', '.join(TO_VISIBLE)}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    try:
        enviar_consolidado(dry_run=args.dry_run)
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
