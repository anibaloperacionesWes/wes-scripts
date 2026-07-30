"""
Genera y envia diariamente (un solo correo):
1) Alerta control nocturno (madrugada de hoy)
2) Reporte de puntos en cero

Destinatarios: Mauricio y Anibal (mismo correo con ambos adjuntos).

Uso manual:
  python enviar_reportes_diarios.py
"""

from __future__ import annotations

import os
import smtplib
import subprocess
import sys
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import List

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
SMTP_PASSWORD = os.environ.get("WES_SMTP_PASSWORD") or "vxbynfpoehbweelj"

# Un solo envío: ambos informes a Mauricio y Anibal
DESTINATARIOS_PRINCIPALES = [
    "mauricioorellana@wes.cl",
    "anibal.aoperaciones@wes.cl",
]


def _workspace_dir() -> Path:
    return Path(__file__).resolve().parent


def _excel_horarios_path() -> Path:
    """
    Ruta oficial del Excel para reportes nocturnos (incluye Parque Arauco si está en el libro).
    Prioriza HORARIOS CONTROL NOCTURNO.xlsx; respaldo HORARIOS COLEGIOS.xlsx.
    """
    reports = _workspace_dir() / "reports"
    reports.mkdir(parents=True, exist_ok=True)
    primary = reports / "HORARIOS CONTROL NOCTURNO.xlsx"
    legacy = reports / "HORARIOS COLEGIOS.xlsx"
    if primary.exists():
        return primary
    if legacy.exists():
        return legacy
    return primary


def _generar_control_nocturno_hoy() -> Path:
    """
    Genera alerta control nocturno para la fecha de hoy (madrugada hoy).
    Retorna ruta del PDF generado.
    """
    from control_nocturno import generar_reporte_control_nocturno

    hoy = datetime.now()
    desde = datetime(hoy.year, hoy.month, hoy.day)
    hasta = desde
    excel_path = _excel_horarios_path()

    rows, out_csv, out_docx, out_pdf = generar_reporte_control_nocturno(
        desde=desde,
        hasta=hasta,
        umbral=0.0,
        excel_path=excel_path,
    )
    print(f"[OK] Control nocturno generado: {out_pdf} (alertas={len(rows)})")
    return out_pdf


def _convertir_docx_a_pdf(docx_path: Path) -> Path:
    """
    Convierte DOCX a PDF si es posible. Si falla, devuelve DOCX.
    """
    pdf_path = docx_path.with_suffix(".pdf")
    try:
        from docx2pdf import convert  # type: ignore

        convert(str(docx_path), str(pdf_path))
        if pdf_path.exists():
            return pdf_path
    except Exception:
        pass
    return docx_path


def _generar_puntos_en_cero() -> Path:
    """
    Ejecuta reporte_puntos_en_cero.py y retorna el ultimo archivo (PDF si se pudo, DOCX si no).
    """
    ws = _workspace_dir()
    script = ws / "reporte_puntos_en_cero.py"
    if not script.exists():
        raise FileNotFoundError(f"No se encontro: {script}")

    cmd = [sys.executable, str(script)]
    print(f"[RUN] {' '.join(cmd)}")
    subprocess.run(cmd, cwd=str(ws), check=True)

    out_dir = ws / "reporte en cero"
    docx_files = [p for p in out_dir.glob("Reporte_Puntos_En_Cero_*.docx") if not p.name.startswith("~")]
    if not docx_files:
        raise FileNotFoundError("No se encontro DOCX de puntos en cero.")
    docx_files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    ultimo_docx = docx_files[0]

    archivo_envio = _convertir_docx_a_pdf(ultimo_docx)
    print(f"[OK] Puntos en cero generado: {archivo_envio}")
    return archivo_envio


def _enviar_correo(destinatarios: List[str], asunto: str, cuerpo: str, adjuntos: List[Path]) -> None:
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = asunto
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for path in adjuntos:
        if not path.exists():
            raise FileNotFoundError(f"No existe adjunto: {path}")
        subtype = "pdf" if path.suffix.lower() == ".pdf" else "octet-stream"
        with path.open("rb") as f:
            part = MIMEApplication(f.read(), _subtype=subtype)
            part.add_header("Content-Disposition", "attachment", filename=path.name)
            msg.attach(part)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, SMTP_PASSWORD)
        server.send_message(msg, to_addrs=destinatarios)

    print(f"[OK] Correo enviado ({asunto}) a:")
    for d in destinatarios:
        print(f"  - {d}")


def main() -> int:
    ws = _workspace_dir()
    os.chdir(ws)

    print("=" * 70)
    print("ENVIO DIARIO DE REPORTES WES")
    print("=" * 70)

    pdf_control_nocturno = _generar_control_nocturno_hoy()
    archivo_puntos_cero = _generar_puntos_en_cero()

    fecha = datetime.now().strftime("%d-%m-%Y")
    asunto = f"Reportes WES diarios — Puntos en cero + Control nocturno — {fecha}"
    cuerpo = f"""Estimados Mauricio y Anibal,

Adjuntos del día {fecha}:

1) Reporte de puntos en cero: {archivo_puntos_cero.name}
2) Alerta control nocturno (madrugada de hoy): {pdf_control_nocturno.name}

Saludos,
Sistema WES
"""
    _enviar_correo(
        destinatarios=DESTINATARIOS_PRINCIPALES,
        asunto=asunto,
        cuerpo=cuerpo,
        adjuntos=[archivo_puntos_cero, pdf_control_nocturno],
    )

    print("[OK] Proceso diario finalizado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

