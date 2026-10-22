"""
Acta de visita técnica — CDUC / Calle de Servicio (000021-05).

Genera Word + PDF, sube a Drive (reports/CDUC/Visitas_Tecnicas) y envía
correo a los destinatarios del hilo de Aníbal (22/10/2026).

Uso:
  python3 generar_y_enviar_acta_visita_calle_servicio_cduc.py
  python3 generar_y_enviar_acta_visita_calle_servicio_cduc.py --solo-generar
  python3 generar_y_enviar_acta_visita_calle_servicio_cduc.py --sin-correo
"""

from __future__ import annotations

import argparse
import os
import smtplib
import subprocess
import sys
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, make_msgid
from pathlib import Path

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt, RGBColor

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
FECHA_VISITA = datetime(2026, 10, 22)
STAMP = datetime.now().strftime("%Y%m%d_%H%M")
OUT_DIR = ROOT / "reports" / "CDUC" / "Visitas_Tecnicas" / f"Calle_de_Servicio_{STAMP}"
NOMBRE_BASE = "Acta_Visita_Tecnica_Calle_de_Servicio_CDUC"

NODE_ID = "000021-05"
NODE_NOMBRE = "Calle de Servicio (camino de servicios)"
CLIENTE = "CDUC — Club Deportivo Universidad Católica"
COMPANY_ID = "000021"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

# Reply-all al hilo de Aníbal (22/10/2026 12:18)
TO_RECIPIENTS = [
    "felipecuevas.mancilla@gmail.com",
    "anibal.aoperaciones@wes.cl",
    "diegocarrasco@wes.cl",
    "broman@cduc.cl",
    "eleal@cduc.cl",
]
CC_RECIPIENTS: list[str] = []

COLOR_TITULO = RGBColor(31, 71, 136)
COLOR_WES = RGBColor(0, 80, 179)


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
    # Fallback histórico del repo (agente.ia@wes.cl)
    return "vxbynfpoehbweelj"


def _set_run_font(run, size_pt: float = 11, bold: bool = False, color: RGBColor | None = None) -> None:
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def _celda(cell, texto: str, *, bold: bool = False, fill: str | None = None) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(texto)
    _set_run_font(run, 10, bold=bold, color=RGBColor(255, 255, 255) if fill else None)
    if fill:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), fill)
        shd.set(qn("w:val"), "clear")
        tcPr.append(shd)


def _bullet(doc: Document, texto: str) -> None:
    p = doc.add_paragraph(texto, style="List Bullet")
    for run in p.runs:
        _set_run_font(run, 11)


def _parrafo(doc: Document, texto: str, *, bold_prefix: str | None = None) -> None:
    p = doc.add_paragraph()
    if bold_prefix:
        r0 = p.add_run(bold_prefix)
        _set_run_font(r0, 11, bold=True)
        r1 = p.add_run(texto)
        _set_run_font(r1, 11)
    else:
        r = p.add_run(texto)
        _set_run_font(r, 11)


def convertir_docx_a_pdf(docx_path: Path) -> Path | None:
    """LibreOffice headless (Linux) o fallback generar_reporte_word."""
    pdf_path = docx_path.with_suffix(".pdf")
    for bin_name in ("soffice", "libreoffice"):
        try:
            subprocess.run(
                [
                    bin_name,
                    "--headless",
                    "--nologo",
                    "--nofirststartwizard",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(docx_path.parent),
                    str(docx_path),
                ],
                check=True,
                capture_output=True,
                timeout=120,
            )
            if pdf_path.is_file():
                return pdf_path
        except (FileNotFoundError, subprocess.CalledProcessError, subprocess.TimeoutExpired):
            continue
    try:
        from generar_reporte_word import convertir_word_a_pdf

        p = convertir_word_a_pdf(docx_path)
        if p and Path(p).is_file():
            return Path(p)
    except Exception as exc:
        print(f"[ADVERTENCIA] Fallback PDF falló: {exc}")
    return None


def generar_acta() -> tuple[Path, Path | None]:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    try:
        from generar_reporte_word import add_logo_to_header

        add_logo_to_header(doc)
    except Exception as exc:
        print(f"[ADVERTENCIA] Logo no agregado: {exc}")

    titulo = doc.add_heading("Acta de visita técnica", level=0)
    if titulo.runs:
        titulo.runs[0].font.color.rgb = COLOR_TITULO

    sub = doc.add_paragraph()
    sub.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r = sub.add_run("Punto Calle de Servicio — Visita presencial")
    _set_run_font(r, 14, bold=True, color=COLOR_WES)

    doc.add_paragraph()

    # Datos generales
    tbl = doc.add_table(rows=7, cols=2)
    tbl.style = "Table Grid"
    filas = [
        ("Cliente", CLIENTE),
        ("Punto WES", f"{NODE_ID} — {NODE_NOMBRE}"),
        ("Tipo de visita", "Presencial / intervención en terreno"),
        ("Fecha de visita", FECHA_VISITA.strftime("%d/%m/%Y")),
        ("Motivo", "Caída abrupta de caudal medido desde ~25–26/08/2026 (comunicación OK)"),
        ("Estado al cierre", "Medición normalizada — en seguimiento"),
        ("Documento generado", datetime.now().strftime("%d/%m/%Y %H:%M")),
    ]
    for i, (k, v) in enumerate(filas):
        _celda(tbl.rows[i].cells[0], k, bold=True, fill="1F4788")
        _celda(tbl.rows[i].cells[1], v)

    doc.add_paragraph()

    h = doc.add_heading("1. Antecedentes", level=1)
    if h.runs:
        h.runs[0].font.color.rgb = COLOR_TITULO
    _parrafo(
        doc,
        "La plataforma WES recibía datos horarios de forma continua, por lo que se descartó "
        "una falla de comunicación. Se observó una caída significativa del caudal medido a "
        "partir del 25–26 de agosto de 2026: en julio el orden de magnitud era ~15 m³/día y, "
        "desde el 27 de agosto, valores del orden de ~0,07 m³/día (septiembre cerró en "
        "aproximadamente 3,1 m³ en el mes). Se programó visita presencial para determinar la causa.",
    )

    h = doc.add_heading("2. Hallazgos en terreno", level=1)
    if h.runs:
        h.runs[0].font.color.rgb = COLOR_TITULO
    _bullet(
        doc,
        "Sensores ultrasónicos con presencia de sulfato en su interior, producto del agua "
        "que ingresó / cayó en el lugar de instalación.",
    )
    _bullet(
        doc,
        "Cables afectados por sulfato en su interior; se cortó el tramo deteriorado hasta "
        "donde el sobrante de cable lo permitió.",
    )
    _bullet(
        doc,
        "Parámetros del ultrasonido desconfigurados, lo que también distorsionaba la medición.",
    )

    h = doc.add_heading("3. Trabajos realizados", level=1)
    if h.runs:
        h.runs[0].font.color.rgb = COLOR_TITULO
    _bullet(doc, "Limpieza y remoción del sulfato en los sensores ultrasónicos.")
    _bullet(doc, "Reacondicionamiento y conexionado correcto de los cables.")
    _bullet(
        doc,
        "Corte del cable afectado por sulfato, aprovechando el sobrante disponible en campo.",
    )
    _bullet(doc, "Reconfiguración de los parámetros del ultrasonido.")
    _bullet(doc, "Verificación de medición: el punto quedó midiendo de forma correcta / normal.")

    h = doc.add_heading("4. Estado actual y seguimiento", level=1)
    if h.runs:
        h.runs[0].font.color.rgb = COLOR_TITULO
    _parrafo(
        doc,
        "El punto se normalizó y está midiendo en rango normal. Queda en seguimiento para "
        "mantener el correcto funcionamiento y detectar a tiempo cualquier nueva anomalía.",
    )

    h = doc.add_heading("5. Observaciones / próxima intervención", level=1)
    if h.runs:
        h.runs[0].font.color.rgb = COLOR_TITULO
    _parrafo(
        doc,
        "En esta visita se cortó el cable hasta el máximo posible: habitualmente existe un "
        "pequeño sobrante, pero en esta ocasión se sobrepasó ese margen. El cable quedó al "
        "límite; no es posible cortar más sin reemplazo.",
    )
    _parrafo(
        doc,
        "Si los sensores/cables vuelven a fallar por sulfato u otra causa asociada, se "
        "requerirá una intervención más larga, con cambio completo de cables (ya no hay "
        "margen de corte adicional). La próxima falla de este tipo implica cambio de cable.",
        bold_prefix="Importante: ",
    )

    h = doc.add_heading("6. Conclusión", level=1)
    if h.runs:
        h.runs[0].font.color.rgb = COLOR_TITULO
    _parrafo(
        doc,
        "La caída de caudal no respondía a un problema de comunicación sino a deterioro "
        "por sulfato en sensores ultrasónicos y cables, sumado a parámetros de ultrasonido "
        "desconfigurados. Tras la intervención presencial, la medición quedó normalizada. "
        "Se mantiene seguimiento operativo; ante una nueva falla se planificará reemplazo de cables.",
    )

    doc.add_paragraph()
    firma = doc.add_paragraph()
    r = firma.add_run("WES — Operaciones / Agente IA")
    _set_run_font(r, 10, color=RGBColor(120, 120, 120))
    p2 = doc.add_paragraph()
    r2 = p2.add_run(
        "Documento generado a partir de la visita presencial al punto Calle de Servicio "
        f"({NODE_ID}), en respuesta al seguimiento de caudal reportado el 22/10/2026."
    )
    _set_run_font(r2, 9, color=RGBColor(120, 120, 120))

    docx_path = OUT_DIR / f"{NOMBRE_BASE}.docx"
    doc.save(docx_path)
    print(f"[OK] Word: {docx_path}")

    pdf_path = convertir_docx_a_pdf(docx_path)
    if pdf_path:
        print(f"[OK] PDF:  {pdf_path}")
    else:
        print("[ADVERTENCIA] No se pudo generar PDF.")
    return docx_path, pdf_path


def subir_drive(paths: list[Path]) -> list[dict]:
    from wes_google_drive import credenciales_configuradas, subir_a_drive

    if not credenciales_configuradas():
        print("[ADVERTENCIA] Sin credenciales Google Drive; se omite subida.")
        return []
    out = []
    sub = "CDUC/Visitas_Tecnicas"
    for p in paths:
        if not p or not p.is_file():
            continue
        info = subir_a_drive(p, subcarpeta=sub)
        print(f"[OK] Drive: {info.get('name')} → {info.get('web_view_link')}")
        out.append(info)
    return out


def enviar_correo(docx_path: Path, pdf_path: Path | None, drive_links: list[dict]) -> None:
    pw = _smtp_password()
    if not pw:
        raise RuntimeError("Falta contraseña SMTP (WES_GMAIL_APP_PASSWORD).")

    links_txt = ""
    if drive_links:
        links_txt = "\nEnlaces Google Drive:\n"
        for info in drive_links:
            links_txt += f"  • {info.get('name')}: {info.get('web_view_link')}\n"

    cuerpo = f"""Estimados,

Adjunto el acta de visita técnica presencial del punto Calle de Servicio ({NODE_ID} — CDUC), en seguimiento al correo del 22/10/2026 sobre la caída de caudal medido.

Resumen
-------
• Visita presencial al punto Calle de Servicio (camino de servicios).
• Causa: sensores ultrasónicos con sulfato en su interior, producto del agua que cayó / ingresó en el lugar; cables también afectados por sulfato.
• Además, los parámetros del ultrasonido se encontraban desconfigurados.
• Se limpió el sulfato, se reacondicionaron los cables y se reconfiguró el ultrasonido.
• Quedó midiendo de forma normal. Se deja en seguimiento.

Observación importante
----------------------
El cable se cortó hasta el límite del sobrante disponible; no es posible cortar más. Si vuelve a fallar, la próxima intervención será más larga e incluirá cambio de cables.

Adjuntos: PDF y Word del acta.
{links_txt}
Saludos cordiales,
Agente IA WES
Operaciones WES
"""

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    if CC_RECIPIENTS:
        msg["Cc"] = ", ".join(CC_RECIPIENTS)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="wes.cl")
    msg["Subject"] = (
        "CDUC — Acta de visita técnica · Calle de Servicio (000021-05) · "
        f"{FECHA_VISITA.strftime('%d/%m/%Y')}"
    )
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    adjuntos: list[tuple[Path, str]] = [
        (docx_path, "vnd.openxmlformats-officedocument.wordprocessingml.document"),
    ]
    if pdf_path and pdf_path.is_file():
        adjuntos.insert(0, (pdf_path, "pdf"))

    for path, subtype in adjuntos:
        with open(path, "rb") as f:
            part = MIMEApplication(f.read(), _subtype=subtype)
            part.add_header("Content-Disposition", "attachment", filename=path.name)
            msg.attach(part)

    destinatarios = TO_RECIPIENTS + CC_RECIPIENTS
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.sendmail(SMTP_USUARIO, destinatarios, msg.as_string())

    print(f"[OK] Correo enviado a: {', '.join(destinatarios)}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Acta visita Calle de Servicio CDUC")
    parser.add_argument("--solo-generar", action="store_true", help="Solo Word/PDF")
    parser.add_argument("--sin-correo", action="store_true", help="Genera y sube Drive, sin correo")
    args = parser.parse_args()

    docx_path, pdf_path = generar_acta()
    if args.solo_generar:
        return 0

    paths = [p for p in (pdf_path, docx_path) if p and p.is_file()]
    drive_links = subir_drive(paths)

    if args.sin_correo:
        return 0

    enviar_correo(docx_path, pdf_path, drive_links)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
