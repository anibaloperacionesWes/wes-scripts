"""Envía por correo a Juan, José y Diego los PDF de cada auditoría Renca, el reporte agregado (PDF) y el PPT."""

from __future__ import annotations

import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from generar_reporte_word import convertir_word_a_pdf

ROOT = Path(__file__).resolve().parent
BASE_AUDIT = (
    ROOT
    / "reports"
    / "reporte de auditoria"
    / "auditoria_puntos_renca_abril_2026"
)

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
SMTP_PASSWORD = (
    os.environ.get("WES_SMTP_PASSWORD", "").strip()
    or os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
    or "vxbynfpoehbweelj"
)

TO_RECIPIENTS = ["juanlopez@wes.cl", "joseotarola@wes.cl", "diegocarrasco@wes.cl"]


def _pdf_principal_carpeta_auditoria(carpeta: Path) -> Path | None:
    pdfs = sorted(carpeta.glob("*.pdf"))
    if not pdfs:
        return None
    con_nodo = [p for p in pdfs if "000017-" in p.name]
    if con_nodo:
        return sorted(con_nodo, key=lambda p: len(p.name), reverse=True)[0]
    return pdfs[0]


def _coleccionar_pdfs_puntos() -> list[Path]:
    out: list[Path] = []
    for d in sorted(BASE_AUDIT.iterdir()):
        if not d.is_dir():
            continue
        if not d.name.startswith("Auditoria "):
            continue
        p = _pdf_principal_carpeta_auditoria(d)
        if p and p.is_file():
            out.append(p)
    return out


def _ultimo_docx_agregado() -> Path | None:
    dirs = sorted(BASE_AUDIT.glob("Reporte_agregado_5_auditorias_*"), key=lambda x: x.stat().st_mtime, reverse=True)
    if not dirs:
        return None
    docx = dirs[0] / "Reporte_agregado_5_auditorias.docx"
    return docx if docx.is_file() else None


def _ultimo_ppt() -> Path | None:
    ppts = sorted(BASE_AUDIT.glob("Informe_auditorias_WES_Renca_*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return ppts[0] if ppts else None


def _adjuntar(msg: MIMEMultipart, path: Path) -> None:
    ext = path.suffix.lower()
    if ext == ".pdf":
        subtype = "pdf"
    elif ext == ".pptx":
        subtype = "vnd.openxmlformats-officedocument.presentationml.presentation"
    else:
        subtype = "octet-stream"
    with open(path, "rb") as f:
        part = MIMEApplication(f.read(), _subtype=subtype)
        part.add_header("Content-Disposition", "attachment", filename=path.name)
        msg.attach(part)


def main() -> int:
    pdfs_pts = _coleccionar_pdfs_puntos()
    if len(pdfs_pts) < 5:
        print(f"[WARN] Se esperaban 5 carpetas Auditoria* con PDF; encontrados {len(pdfs_pts)} PDF.")

    docx_agg = _ultimo_docx_agregado()
    if docx_agg is None:
        print("[ERROR] No hay Reporte_agregado_5_auditorias.docx en Reporte_agregado_*")
        return 1

    pdf_agg = convertir_word_a_pdf(docx_agg)
    if pdf_agg is None or not Path(pdf_agg).exists():
        print("[ERROR] No se pudo convertir el agregado a PDF.")
        return 1
    pdf_agg = Path(pdf_agg)

    ppt = _ultimo_ppt()
    if ppt is None or not ppt.is_file():
        print("[ERROR] No hay Informe_auditorias_WES_Renca_*.pptx en la carpeta de auditorías.")
        return 1

    adjuntos: list[Path] = list(pdfs_pts) + [pdf_agg, ppt]

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Subject"] = "Auditorías WES Renca — PDF por punto, reporte agregado e informe PPT"

    lista = "\n".join(f" - {p.name}" for p in adjuntos)
    cuerpo = (
        "Estimados Juan, José y Diego,\n\n"
        "Adjunto:\n"
        " • PDF de auditoría por cada punto Renca (5 archivos)\n"
        " • Reporte agregado de las 5 auditorías (PDF)\n"
        " • Informe PowerPoint consolidado\n\n"
        f"Archivos ({len(adjuntos)}):\n{lista}\n\n"
        "Saludos,\nAgente IA WES\n"
    )
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for p in adjuntos:
        _adjuntar(msg, p)

    try:
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.sendmail(SMTP_USUARIO, TO_RECIPIENTS, msg.as_string())
        print(f"[OK] Enviado a: {', '.join(TO_RECIPIENTS)}")
        print(f"[OK] Adjuntos: {len(adjuntos)}")
        for p in adjuntos:
            print(f"     - {p}")
        return 0
    except Exception as e:
        print(f"[ERROR] Envío: {e}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
