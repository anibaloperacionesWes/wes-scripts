"""
Envía a Juan los informes de gestión hídrica (formato Zapallar)
del lote de colegios, periodo agosto 2026.

Clientes: Renca, La Florida, La Reina, CORMUP y Alexander Fleming.

Uso:
  python enviar_gestion_hidrica_colegios_agosto2026_juan.py
  python enviar_gestion_hidrica_colegios_agosto2026_juan.py --dry-run
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
]

ASUNTO = "Informes de gestión hídrica — Colegios — Agosto 2026"

# (cliente, carpeta, estado, one-pager Drive, mensual Drive)
CLIENTES = [
    (
        "Renca (Lo Velásquez, Cumbre Pte., ICCO)",
        "Renca",
        "Bajo control (13 % nocturno) — 2.044,6 m³",
        "https://drive.google.com/file/d/1l98gUOCr1vMudovBhkdiakK5XUUnarbU/view?usp=drivesdk",
        "https://drive.google.com/file/d/1CIo1OAGYw2aLoKGIf9UndO31r3o8tYmr/view?usp=drivesdk",
    ),
    (
        "La Florida (Liceo Alto Cordillera)",
        "La_Florida",
        "En observación (24 % nocturno) — 575,6 m³",
        "https://drive.google.com/file/d/1CSppysamO-1CQGs2l_Py_J_4ArK-537n/view?usp=drivesdk",
        "https://drive.google.com/file/d/1nIrg0-e8ltKbk-o6UhltYdKA-vgKpWzg/view?usp=drivesdk",
    ),
    (
        "La Reina (Eugenio María De Hostos)",
        "La_Reina",
        "Bajo control (0 % nocturno) — 541,3 m³ · 5 al 31/08; julio excluido por falla del sensor de pulso",
        "https://drive.google.com/file/d/19nUp-2PzreetV6OHIKlNLxsLJUneILjY/view?usp=drivesdk",
        "https://drive.google.com/file/d/133EZnEB5yNR7ZGTbpUNwru1wNYGrRRqJ/view?usp=drivesdk",
    ),
    (
        "CORMUP Peñalolén (14 colegios)",
        "CORMUP",
        "Bajo control (10 % nocturno) — 14.337,0 m³",
        "https://drive.google.com/file/d/1M4NZJbRtRINVLPtv0-P28YuQBvWBw2Il/view?usp=drivesdk",
        "https://drive.google.com/file/d/1NwZSVhY-clNpBVrk6ifrYimOX7IkHoEj/view?usp=drivesdk",
    ),
    (
        "Alexander Fleming",
        "Alexander_Fleming",
        "Requiere atención — 41.269 m³ (4,4× la mediana de mar–jul); 24 % nocturno",
        "https://drive.google.com/file/d/1x4NQthEWLWNhH73_8S-JH_qEWnBG5XB2/view?usp=drivesdk",
        "https://drive.google.com/file/d/1zZF2PBzrgc1gT-HE5w0OoW_QE-wE8aYw/view?usp=drivesdk",
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
        "Estimado Juan,\n\n"
        "Adjunto los informes de gestión hídrica (formato ejecutivo Zapallar) "
        "de colegios, periodo agosto de 2026.\n\n"
        "Por cada cliente van dos PDF: one-pager y informe mensual.\n\n"
        f"{lista}\n\n"
        "Notas:\n"
        "  • La Reina: se excluye julio por falla del sensor de pulso; agosto usa "
        "solo los días válidos (5 al 31).\n"
        "  • Alexander Fleming: solo el medidor 000022-00; Juan Pablo II no entra.\n\n"
        f"Son {len(pdfs)} archivos PDF adjuntos en este mismo correo.\n\n"
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
