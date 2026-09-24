"""
Propuesta de monitoreo hídrico — Weir Minerals, San Bernardo.

La vista satelital está en fichas/Weir_San_Bernardo/ubicacion/vista_satelital.png.
Los sitios de la visita se agregan en fichas/Weir_San_Bernardo/sitios/<carpeta>/
con fotos (jpg/png) y descripcion.txt.

Uso:
  python generar_propuesta_monitoreo_weir_san_bernardo.py
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Inches, Pt, RGBColor
from PIL import Image

ROOT = Path(__file__).resolve().parent
BASE = ROOT / "fichas" / "Weir_San_Bernardo"
VISTA = BASE / "ubicacion" / "baquedano_1215.png"
LOGO = BASE / "ubicacion" / "logo_wes.png"
FOTO_MEDIDOR = BASE / "sitios" / "01_baquedano_1215" / "01_medidor.png"
SITIOS_DIR = BASE / "sitios"
OUT_DIR = ROOT / "reports" / "Weir_San_Bernardo" / "PROPUESTA"
OUT_DOCX = OUT_DIR / "Propuesta_monitoreo_Weir_San_Bernardo.docx"

AZUL = RGBColor(0, 51, 102)
NEGRO = RGBColor(0, 0, 0)
EXT_FOTO = {".png", ".jpg", ".jpeg", ".webp"}


def _shade(cell, hex_color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tc_pr.append(shd)


def _set_run(run, *, size=11, bold=False, color=NEGRO, name="Calibri") -> None:
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color


def _p(doc, text, *, bold=False, size=11, center=False, space_after=8, color=NEGRO):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    run = para.add_run(text)
    _set_run(run, size=size, bold=bold, color=color)
    return para


def _heading(doc, text: str) -> None:
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(12)
    para.paragraph_format.space_after = Pt(6)
    para.paragraph_format.keep_with_next = True
    run = para.add_run(text.upper())
    _set_run(run, size=13, bold=True, color=AZUL)


def _kv_table(doc, rows: list[tuple[str, str]]) -> None:
    table = doc.add_table(rows=len(rows), cols=2)
    table.autofit = True
    for i, (k, v) in enumerate(rows):
        c0, c1 = table.rows[i].cells
        c0.text = ""
        c1.text = ""
        _set_run(c0.paragraphs[0].add_run(k), size=10, bold=True, color=AZUL)
        _set_run(c1.paragraphs[0].add_run(v), size=10)
        _shade(c0, "E8EEF6")
        if i % 2 == 1:
            _shade(c1, "F7F9FC")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def _footer(section) -> None:
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("WES  ·  Propuesta de monitoreo hídrico  ·  Weir Minerals, San Bernardo  ·  ")
    _set_run(r, size=8, color=RGBColor(90, 90, 90))
    run = p.add_run()
    _set_run(run, size=8, color=RGBColor(90, 90, 90))
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


def _ancho_foto(path: Path, max_w_cm: float, max_h_cm: float) -> Emu:
    with Image.open(path) as im:
        w_px, h_px = im.size
    if w_px <= 0 or h_px <= 0:
        return Cm(max_w_cm)
    ratio = w_px / h_px
    ancho = max_w_cm
    if ancho / ratio > max_h_cm:
        ancho = max_h_cm * ratio
    return Cm(ancho)


def construir() -> Document:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.6)
    _footer(section)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.font.color.rgb = NEGRO

    if LOGO.is_file():
        logo_p = doc.add_paragraph()
        logo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo_p.add_run().add_picture(str(LOGO), width=Inches(1.55))

    _p(doc, "PROPUESTA DE MONITOREO HÍDRICO", bold=True, size=18, center=True, space_after=2, color=AZUL)
    _p(doc, "Weir Minerals Chile", bold=True, size=16, center=True, space_after=2, color=AZUL)
    _p(doc, "Av. San José 0815, San Bernardo", size=12, center=True, space_after=10)

    _heading(doc, "1. Ubicación de la empresa")
    _kv_table(
        doc,
        [
            ("Empresa", "Weir Minerals Chile (Vulco S.A.)"),
            ("Dirección", "Av. San José 0815"),
            ("Comuna", "San Bernardo"),
            ("Región", "Metropolitana"),
            ("País", "Chile"),
        ],
    )

    _heading(doc, "2. Punto 1 — Baquedano 1215")
    _kv_table(
        doc,
        [
            ("Dirección del punto", "Baquedano 1215, San Bernardo"),
            ("Qué se ve en terreno", "Medidor con dos salidas"),
        ],
    )
    _p(
        doc,
        "En el medidor se separan dos líneas. La que sigue hacia abajo es la Línea A. "
        "La que sale hacia la derecha, por el costado del camino, es la Línea B. "
        "Cada una tiene su propia descripción.",
    )

    if VISTA.is_file():
        doc.add_page_break()
        _p(doc, "Vista satelital del punto", bold=True, size=12, center=True, space_after=4, color=AZUL)
        foto = doc.add_paragraph()
        foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
        foto.paragraph_format.space_before = Pt(2)
        foto.paragraph_format.space_after = Pt(2)
        foto.add_run().add_picture(str(VISTA), width=_ancho_foto(VISTA, 16.2, 9.2))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_after = Pt(8)
        r = cap.add_run("Vista satelital — Baquedano 1215, San Bernardo. Logo WES en la esquina superior derecha.")
        _set_run(r, size=9, color=RGBColor(80, 80, 80))

    if FOTO_MEDIDOR.is_file():
        doc.add_page_break()
        _p(doc, "Medidor y las dos salidas", bold=True, size=12, center=True, space_after=4, color=AZUL)
        pic = doc.add_paragraph()
        pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic.paragraph_format.space_after = Pt(2)
        pic.add_run().add_picture(str(FOTO_MEDIDOR), width=_ancho_foto(FOTO_MEDIDOR, 10.5, 9.2))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_after = Pt(8)
        r = cap.add_run("Foto 1. Medidor en Baquedano 1215. Hacia abajo: Línea A. Hacia la derecha: Línea B.")
        _set_run(r, size=9, color=RGBColor(80, 80, 80))

    def _caja(titulo: str, texto: str) -> None:
        caja = doc.add_table(rows=1, cols=1)
        cell = caja.cell(0, 0)
        cell.text = ""
        _shade(cell, "F4F7FB")
        titulo_p = cell.paragraphs[0]
        titulo_p.paragraph_format.space_after = Pt(2)
        rt = titulo_p.add_run(titulo)
        _set_run(rt, size=11, bold=True, color=AZUL)
        cuerpo = cell.add_paragraph()
        cuerpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        cuerpo.paragraph_format.space_after = Pt(2)
        _set_run(cuerpo.add_run(texto), size=11)
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    _caja(
        "Línea A — sale del medidor hacia abajo",
        "Va desde la alimentación de camión aljibe, el llenado de estanque, el riego y "
        "camarines frontera, hasta la alimentación de riego de las canchas de tenis y "
        "el baño del estadio.",
    )
    _caja(
        "Línea B — sale del medidor hacia la derecha",
        "Sigue por el costado del camino hasta los baños de logística. Solo alimenta baños.",
    )
    return doc


def main() -> None:
    if not VISTA.is_file():
        raise SystemExit(f"Falta la vista satelital: {VISTA}")
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    construir().save(OUT_DOCX)
    print(f"[OK] {OUT_DOCX}")


if __name__ == "__main__":
    main()
