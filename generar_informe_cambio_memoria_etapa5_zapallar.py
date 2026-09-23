"""
Informe corto — cambio de memoria placa / Etapa N°5 Fundo Zapallar.

Documenta la revisión en terreno (sensor, pulsos, voltajes), el criterio hidráulico
DN90 (~60 m³/h) y la normalización tras reemplazo de memoria de la placa.

Uso:
  python generar_informe_cambio_memoria_etapa5_zapallar.py
"""

from __future__ import annotations

import sys
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt, RGBColor

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "reports" / "Fundo_Zapallar" / "Informes_Tecnicos"
COMPANY = "Fundo Zapallar"
COMPANY_ID = "000027"
NODE_ID = "000027-03"
NODE_NAME = "Etapa N°5"
FECHA_INTERVENCION = date(2026, 9, 22)
DIAMETRO = "DN90 fierro dúctil"
CAUDAL_MAX_REF_M3H = 60.0
COLOR_TITULO = RGBColor(31, 71, 136)


def _set_run_font(run, *, bold: bool = False, size: int = 11, color: RGBColor | None = None) -> None:
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    if color is not None:
        run.font.color.rgb = color


def _add_bullet(doc: Document, text: str) -> None:
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(text)
    _set_run_font(run, size=11)


def _shade_cell(cell, hex_color: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def build_doc() -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    ahora = datetime.now(ZoneInfo("America/Santiago"))
    stamp = ahora.strftime("%Y%m%d_%H%M")
    out = OUT_DIR / f"Informe_Cambio_Memoria_Etapa5_Zapallar_{stamp}.docx"

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("INFORME TÉCNICO CORTO")
    _set_run_font(r, bold=True, size=16, color=COLOR_TITULO)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run(
        f"Cambio de memoria de placa — {NODE_NAME} ({NODE_ID})\n"
        f"{COMPANY} ({COMPANY_ID})"
    )
    _set_run_font(r, bold=True, size=12, color=COLOR_TITULO)

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = meta.add_run(
        f"Fecha intervención: {FECHA_INTERVENCION.strftime('%d/%m/%Y')}  ·  "
        f"Generado: {ahora.strftime('%d/%m/%Y %H:%M')} (Chile)"
    )
    _set_run_font(r, size=10, color=RGBColor(90, 90, 90))

    h = doc.add_heading("1. Resumen", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO

    p = doc.add_paragraph()
    r = p.add_run(
        "En terreno se detectaron lecturas de caudal anómalas (pulsos del orden de "
        "92 y 200 m³/h) en la matriz de Etapa N°5. Tras descartar falla del sensor "
        "y de la cadena de pulsos, y verificar voltajes correctos, se concluyó que "
        "el error provenía de la memoria de la placa. Se reemplazó la memoria para "
        "normalizar el punto y evitar recurrencia."
    )
    _set_run_font(r, size=11)

    h = doc.add_heading("2. Revisión en terreno", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO

    _add_bullet(
        doc,
        "Sensor Census / sensor inductivo: probado y en buen estado.",
    )
    _add_bullet(
        doc,
        "Pulsos de revisión: realizados; respuesta correcta.",
    )
    _add_bullet(
        doc,
        "Voltajes de alimentación / alimentación de placa: correctos.",
    )
    _add_bullet(
        doc,
        "Diagnóstico: memoria de la placa arrojaba el error (lecturas irreales).",
    )
    _add_bullet(
        doc,
        "Acción correctiva: cambio de memoria de la placa; punto reparado/normalizado.",
    )

    h = doc.add_heading("3. Criterio hidráulico (DN90)", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO

    p = doc.add_paragraph()
    r = p.add_run(
        f"La red del punto es tubería de {DIAMETRO}. Como referencia de capacidad "
        f"máxima razonable de red se adopta ≈ {CAUDAL_MAX_REF_M3H:.0f} m³/h "
        f"(equivalente a ~2,5 m/s en diámetro interior ≈ 90 mm)."
    )
    _set_run_font(r, size=11)

    table = doc.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    filas = [
        ("Velocidad de referencia", "Caudal teórico DN90"),
        ("1,5 m/s (diseño habitual)", "≈ 34 m³/h"),
        ("2,0 m/s", "≈ 46 m³/h"),
        ("2,5 m/s (techo práctico de red)", "≈ 57–60 m³/h"),
    ]
    for i, (a, b) in enumerate(filas):
        table.rows[i].cells[0].text = a
        table.rows[i].cells[1].text = b
        for cell in table.rows[i].cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    _set_run_font(run, bold=(i == 0), size=10)
        if i == 0:
            for cell in table.rows[i].cells:
                _shade_cell(cell, "1F4788")
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(255, 255, 255)

    p = doc.add_paragraph()
    r = p.add_run(
        "Los pulsos observados de ~92 m³/h y ~200 m³/h superan de forma clara el "
        "techo hidráulico de la tubería (implicarían velocidades del orden de "
        "4–9 m/s), por lo que no corresponden a caudal real de la red sino a "
        "error de registro en memoria."
    )
    _set_run_font(r, size=11)

    h = doc.add_heading("4. Conclusión", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO

    p = doc.add_paragraph()
    r = p.add_run(
        "Se confirma falla de memoria de placa (no del sensor ni de la red). "
        "Con el reemplazo de memoria el punto queda normalizado. Se deja "
        "seguimiento diario matutino del nodo 000027-03 con umbral de alerta "
        f"en {CAUDAL_MAX_REF_M3H:.0f} m³/h para detectar cualquier recurrencia "
        "de lecturas imposibles."
    )
    _set_run_font(r, size=11)

    p = doc.add_paragraph()
    r = p.add_run(
        "Nota: este informe documenta el hallazgo y la intervención en terreno "
        "según revisión de campo del equipo WES."
    )
    _set_run_font(r, size=9, color=RGBColor(100, 100, 100))

    doc.save(out)
    return out


def main() -> int:
    path = build_doc()
    print(f"[OK] Informe generado: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
