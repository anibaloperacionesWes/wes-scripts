# -*- coding: utf-8 -*-
"""Genera PDF de acta de visita técnica WES a partir de un dict/JSON."""

from __future__ import annotations

import base64
import io
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parent
LOGO = ROOT / "logo_wes.png"
OUT_DIR = ROOT / "salidas"

BLUE = colors.HexColor("#1F4E79")
BLUE2 = colors.HexColor("#2E75B6")
LIGHT = colors.HexColor("#E7F0F8")
AMBER = colors.HexColor("#FFF4D6")
LINE = colors.HexColor("#C5D5E6")


def _slug(text: str) -> str:
    text = (text or "sin-dato").strip()
    text = (
        text.encode("ascii", "ignore").decode("ascii")
        if False
        else text
    )
    text = re.sub(r"[^\w\-]+", "_", text, flags=re.UNICODE)
    return (text.strip("_") or "sin-dato")[:40]


def _esc(text: Any) -> str:
    s = "" if text is None else str(text)
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _p(text: Any, style: ParagraphStyle, *, raw_html: bool = False) -> Paragraph:
    """Si raw_html=False, escapa el texto. Si True, deja <b>/<br/> de ReportLab."""
    s = "" if text is None else str(text)
    if not raw_html:
        s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        s = s.replace("\n", "<br/>")
    else:
        s = s.replace("\n", "<br/>")
    return Paragraph(s or "—", style)


def _firma_image(data_url: str, max_w: float = 90 * mm, max_h: float = 28 * mm) -> Optional[Image]:
    if not data_url or "," not in data_url:
        return None
    try:
        raw = base64.b64decode(data_url.split(",", 1)[1])
        img = Image(io.BytesIO(raw))
        img._restrictSize(max_w, max_h)
        return img
    except Exception:
        return None


def _checklist_table(items: List[Dict[str, str]], styles) -> Table:
    data = [[
        _p("<b>Elemento</b>", styles["cell"], raw_html=True),
        _p("<b>Estado</b>", styles["cell"], raw_html=True),
        _p("<b>Obs. / medición</b>", styles["cell"], raw_html=True),
    ]]
    for it in items or []:
        data.append([
            _p(it.get("elemento", ""), styles["cell"]),
            _p(it.get("estado", ""), styles["cell"]),
            _p(it.get("obs", "") or "—", styles["cell"]),
        ])
    if len(data) == 1:
        data.append([_p("—", styles["cell"]), _p("—", styles["cell"]), _p("—", styles["cell"])])
    t = Table(data, colWidths=[70 * mm, 28 * mm, 72 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), LIGHT),
        ("TEXTCOLOR", (0, 0), (-1, 0), BLUE),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFD")]),
    ]))
    return t


def generar_pdf_acta(data: Dict[str, Any], out_path: Optional[Path] = None) -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    cliente = _slug(str(data.get("cliente") or "cliente"))
    maquina = _slug(str(data.get("maquina") or "sitio"))
    if out_path is None:
        out_path = OUT_DIR / f"Acta_visita_WES_{cliente}_{maquina}_{stamp}.pdf"

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="title_wes", fontName="Helvetica-Bold", fontSize=16,
        textColor=colors.white, alignment=TA_CENTER, leading=20,
    ))
    styles.add(ParagraphStyle(
        name="sub_wes", fontName="Helvetica", fontSize=9,
        textColor=colors.white, alignment=TA_CENTER, leading=12,
    ))
    styles.add(ParagraphStyle(
        name="hsec", fontName="Helvetica-Bold", fontSize=11,
        textColor=BLUE, spaceBefore=8, spaceAfter=4, leading=14,
    ))
    styles.add(ParagraphStyle(
        name="body", fontName="Helvetica", fontSize=9.5,
        textColor=colors.HexColor("#1a1a1a"), leading=13, alignment=TA_LEFT,
    ))
    styles.add(ParagraphStyle(
        name="cell", fontName="Helvetica", fontSize=8.5,
        textColor=colors.HexColor("#1a1a1a"), leading=11,
    ))
    styles.add(ParagraphStyle(
        name="small", fontName="Helvetica", fontSize=8,
        textColor=colors.HexColor("#445566"), leading=10,
    ))
    styles.add(ParagraphStyle(
        name="acusar", fontName="Helvetica-Bold", fontSize=10,
        textColor=BLUE, leading=13, alignment=TA_CENTER,
    ))

    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=f"Acta visita WES — {data.get('cliente')} / {data.get('maquina')}",
        author="WES",
    )

    story = []

    # Header
    logo_cell: Any = ""
    if LOGO.is_file():
        logo = Image(str(LOGO))
        logo._restrictSize(42 * mm, 12 * mm)
        logo_cell = logo
    header = Table(
        [[
            logo_cell,
            [
                _p("ACTA DE VISITA TÉCNICA WES", styles["title_wes"]),
                _p("Sociedad Tecnológica WES SpA · www.wes.cl", styles["sub_wes"]),
            ],
        ]],
        colWidths=[48 * mm, 122 * mm],
    )
    header.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), BLUE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(header)
    story.append(Spacer(1, 6 * mm))

    meta = [
        ["Folio / OT", str(data.get("folio") or data.get("ot") or "—"), "Estado", str(data.get("estado_visita") or "—")],
        ["Cliente", str(data.get("cliente") or "—"), "Máquina / sitio", str(data.get("maquina") or "—")],
        ["Comuna", str(data.get("comuna") or "—"), "Fecha", str(data.get("fecha") or "—")],
        ["Hora", str(data.get("hora") or "—"), "Técnico WES", str(data.get("tecnico") or "—")],
        ["Lectura medidor", str(data.get("lectura_medidor") or "—"), "", ""],
        [
            "Motivo",
            ", ".join(data.get("motivos") or []) or "—",
            "Tecnología",
            ", ".join(data.get("tecnologias") or []) or "—",
        ],
        [
            "Tipo mtto",
            str(data.get("tipo_mtto") or "—"),
            "Tipo falla",
            str(data.get("tipo_falla") or "—"),
        ],
        [
            "Falla específica",
            str(data.get("falla_especifica") or "—"),
            "",
            "",
        ],
    ]
    meta_cells = []
    for row in meta:
        meta_cells.append([
            _p(f"<b>{row[0]}</b>", styles["cell"], raw_html=True),
            _p(row[1], styles["cell"]),
            _p(f"<b>{row[2]}</b>", styles["cell"], raw_html=True) if row[2] else _p("", styles["cell"]),
            _p(row[3], styles["cell"]),
        ])
    mt = Table(meta_cells, colWidths=[32 * mm, 53 * mm, 32 * mm, 53 * mm])
    mt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), LIGHT),
        ("BACKGROUND", (2, 0), (2, -1), LIGHT),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(mt)

    story.append(_p("Solución y/o diagnóstico", styles["hsec"]))
    story.append(_p(data.get("solucion") or "—", styles["body"]))
    if data.get("observaciones"):
        story.append(_p("Observaciones", styles["hsec"]))
        story.append(_p(data.get("observaciones"), styles["body"]))

    story.append(_p("Checklist CIR — eléctrico / electrónico", styles["hsec"]))
    story.append(_checklist_table(data.get("checklist_cir") or [], styles))
    story.append(_p("Checklist CPA — hídrico / cámara", styles["hsec"]))
    story.append(_checklist_table(data.get("checklist_cpa") or [], styles))
    sab = data.get("checklist_sab") or []
    if any((x.get("estado") not in (None, "", "OK", "N/A")) or x.get("obs") for x in sab) or any(
        (x.get("estado") or "") not in ("", "OK") for x in sab
    ):
        # Always include SAB section briefly
        pass
    story.append(_p("Checklist SAB (si aplica)", styles["hsec"]))
    story.append(_checklist_table(sab, styles))

    story.append(Spacer(1, 4 * mm))
    story.append(_p("Recepción del cliente · solicitud de acuse de recibo", styles["hsec"]))
    story.append(_p(
        "Se solicita al cliente <b>acusar recibo</b> de esta acta respondiendo el correo "
        "con la frase «Acuso recibo» o firmando digitalmente abajo. "
        "La constancia queda registrada en el sistema de análisis WES.",
        styles["body"],
        raw_html=True,
    ))

    firma_img = _firma_image(str(data.get("firma_png") or ""))
    firma_block = [
        [
            _p(
                f"<b>Recibido por:</b> {_esc(data.get('recibido_por') or '—')}",
                styles["body"],
                raw_html=True,
            ),
            _p(
                f"<b>Cargo:</b> {_esc(data.get('cargo') or '—')}",
                styles["body"],
                raw_html=True,
            ),
        ],
        [
            firma_img or _p("(sin firma)", styles["small"]),
            _p(
                f"Correo cliente: {_esc(data.get('email_cliente') or '—')}<br/>"
                f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                styles["small"],
                raw_html=True,
            ),
        ],
    ]
    ft = Table(firma_block, colWidths=[95 * mm, 75 * mm])
    ft.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1, BLUE2),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FBFE")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(ft)
    story.append(Spacer(1, 5 * mm))
    story.append(_p("POR FAVOR ACUSAR RECIBO DE ESTA ACTA POR CORREO", styles["acusar"]))
    story.append(_p(
        "WES · Estrecho de Magallanes 1481, Renca · +569 7559 5695 / +569 8198 1426",
        styles["small"],
    ))

    doc.build(story)
    return out_path


if __name__ == "__main__":
    demo = {
        "cliente": "CORMUP",
        "maquina": "TOBALABA",
        "comuna": "Providencia",
        "fecha": datetime.now().strftime("%Y-%m-%d"),
        "hora": "10:30",
        "tecnico": "Anibal Aranda",
        "ot": "OT-DEMO-001",
        "motivos": ["Mantenimiento"],
        "tecnologias": ["CPA y CIR"],
        "tipo_mtto": "Mtto Preventivo",
        "tipo_falla": "Auditoría",
        "falla_especifica": "Validación Data",
        "solucion": "Revisión preventiva CIR/CPA. Sistemas operativos.",
        "observaciones": "Demo de acta digital.",
        "estado_visita": "cerrada",
        "recibido_por": "Juan Pérez",
        "cargo": "Encargado",
        "email_cliente": "demo@cliente.cl",
        "checklist_cir": [{"elemento": "Conectividad App", "estado": "OK", "obs": ""}],
        "checklist_cpa": [{"elemento": "Func. medidor", "estado": "OK", "obs": ""}],
        "checklist_sab": [],
        "firma_png": "",
    }
    path = generar_pdf_acta(demo)
    print(path)
