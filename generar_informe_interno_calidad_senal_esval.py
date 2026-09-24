"""
Informe final — Validación Matriz ESVAL (Fundo Zapallar).

Uso:
  python generar_informe_interno_calidad_senal_esval.py
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

try:
    from PIL import Image
except ImportError:
    Image = None  # type: ignore

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent

# Paleta WES
_NAVY = RGBColor(31, 71, 136)
_NAVY_HEX = "1F4788"
_NAVY_LIGHT_HEX = "D6E3F0"
_ROW_ALT_HEX = "F5F8FB"
_OK_HEX = "E8F5E9"
_OK_TEXT = RGBColor(27, 94, 32)
_MUTED = RGBColor(100, 110, 120)
_LINE = RGBColor(31, 71, 136)
_WHITE = RGBColor(255, 255, 255)
_BLACK = RGBColor(40, 40, 40)

NODE_ID = "000027-01"
NODO_NOMBRE = "Matriz ESVAL"
EMPRESA = "Fundo Zapallar"
COMPANY_ID = "000027"

UMBRAL_FABRICANTE = 60
DN_UP_INICIAL, DN_UP_POST = 76.7, 80
Q_INICIAL, Q_POST = 93, 96

LITROS_ESVAL_ITRON_1 = 230
LITROS_ULTRASONIDO_1 = 100
DIFERENCIA_LITROS_POST_CONFIG = 40
DIFERENCIA_LITROS_POST_M45 = 30

DIAMETRO_CONFIG_ERRONEO_MM = 110
DIAMETRO_ESVAL_MM = 118
DIAMETRO_FINAL_MM = 127
INCREMENTO_DIAMETRO_PCT = 8
ESPESOR_MORTERO_MM = 3
MATERIAL = "Fierro dúctil (EN545)"
MARCADO_FISICO = "EN545 / DN100 / PN16"

M45_FACTOR_ACTUAL = 0.5934
M45_CAUDAL_TURBINA_LPM = 460.0
M45_CAUDAL_ULTRASONIDO_LPM = 87.05
M45_CALCULADO = 3.1362
M45_MAXIMO_EQUIPO = 1.5

VAL_FECHA = "22-09-2026"
HORA_INI, HORA_FIN = "11:50", "15:00"
ITRON_INI, ITRON_FIN = 383956.27, 383978.01
ITRON_DELTA = round(ITRON_FIN - ITRON_INI, 2)
US_NET_CRUDO_INI, US_NET_CRUDO_FIN = 1811906, 1814145
US_INI = round(US_NET_CRUDO_INI * 0.01, 2)
US_FIN = round(US_NET_CRUDO_FIN * 0.01, 2)
US_DELTA = round(US_FIN - US_INI, 2)
ERROR_PCT = round((1 - ITRON_DELTA / US_DELTA) * 100, 1)
TOL_US_PCT = 1
TOL_ITRON_PCT = 5

VAL_FECHA_B_INI = "22-09-2026"
VAL_FECHA_B_FIN = "23-09-2026"
HORA_B_INI, HORA_B_FIN = "15:00", "17:00"
ITRON_B_INI = ITRON_FIN
ITRON_B_FIN = 384148.9
ITRON_B_DELTA = round(ITRON_B_FIN - ITRON_B_INI, 2)
APP_B_DELTA = 175.41
ERROR_B_PCT = round((1 - ITRON_B_DELTA / APP_B_DELTA) * 100, 1)

EVIDENCIAS = ROOT / "reports" / "Fundo_Zapallar" / "Informes_Tecnicos" / "_evidencias"
FOTO_ITRON_1500 = EVIDENCIAS / "itron_1500.jpg"
FOTO_ITRON_1700 = EVIDENCIAS / "itron_1700_2309.jpg"
# Rotación por foto para odómetro horizontal (PIL: 90=CCW, 270=CW)
FOTO_ROTACION = {
    FOTO_ITRON_1500.name: Image.ROTATE_90 if Image else None,    # CCW
    FOTO_ITRON_1700.name: Image.ROTATE_270 if Image else None,   # CW
}
FOTO_ANCHO_CM = 5.8


def _fmt(x: float, dec: int = 2) -> str:
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _font(run, *, size: int = 11, bold: bool = False, color: RGBColor | None = None) -> None:
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def _shade(cell, hex_color: str) -> None:
    shading = parse_xml(
        f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{hex_color}"/>'
    )
    cell._tc.get_or_add_tcPr().append(shading)


def _set_cell_borders(cell, color: str = "B0BEC5", sz: str = "4") -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    # quitar bordes previos si existen
    for child in list(tcPr):
        if child.tag == qn("w:tcBorders"):
            tcPr.remove(child)
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)


def _row_cant_split(row) -> None:
    """Impide que la fila de tabla se parta entre páginas."""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    # evitar duplicados
    for child in list(trPr):
        if child.tag == qn("w:cantSplit"):
            return
    cant = OxmlElement("w:cantSplit")
    trPr.append(cant)


def _keep_table_on_one_page(tbl) -> None:
    """Evita cortes de tabla entre páginas (todas las filas cantSplit)."""
    for row in tbl.rows:
        _row_cant_split(row)


def _para_keep_with_next(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag == qn("w:keepNext"):
            return
    kn = OxmlElement("w:keepNext")
    pPr.append(kn)


def _foto_horizontal(src: Path, out_dir: Path) -> Path:
    """Copia la foto rotada para que los números del medidor se lean en horizontal."""
    out_dir.mkdir(parents=True, exist_ok=True)
    rot = FOTO_ROTACION.get(src.name)
    dst = out_dir / f"{src.stem}_horiz.jpg"
    if Image is None or rot is None:
        shutil.copy2(src, dst if not dst.exists() else dst)
        if not dst.exists():
            shutil.copy2(src, dst)
        return dst if dst.exists() else src
    im = Image.open(src).convert("RGB")
    im = im.transpose(rot)
    im.save(dst, quality=92, optimize=True)
    return dst


def _cell_text(
    cell,
    text: str,
    *,
    bold: bool = False,
    size: int = 10,
    color: RGBColor | None = None,
    align=WD_ALIGN_PARAGRAPH.LEFT,
) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    _font(r, size=size, bold=bold, color=color or _BLACK)


def _style_table(tbl) -> None:
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = True
    for row in tbl.rows:
        for cell in row.cells:
            _set_cell_borders(cell)
    _keep_table_on_one_page(tbl)


def _h(doc: Document, text: str, level: int = 1) -> None:
    p = doc.add_heading(text, level=level)
    p.paragraph_format.space_before = Pt(14 if level == 1 else 10)
    p.paragraph_format.space_after = Pt(6)
    _para_keep_with_next(p)
    for r in p.runs:
        r.font.color.rgb = _NAVY
        r.font.name = "Calibri"
        r.font.size = Pt(14 if level == 1 else 12)


def _p(doc: Document, text: str, *, size: int = 11, keep_next: bool = False) -> None:
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(6)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if keep_next:
        _para_keep_with_next(para)
    r = para.add_run(text)
    _font(r, size=size, color=_BLACK)


def _bullet(doc: Document, text: str) -> None:
    para = doc.add_paragraph(style="List Bullet")
    para.paragraph_format.space_after = Pt(2)
    para.paragraph_format.space_before = Pt(0)
    para.clear()
    r = para.add_run(text)
    _font(r, size=10.5, color=_BLACK)


def _hr(doc: Document) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), _NAVY_HEX)
    pBdr.append(bottom)
    pPr.append(pBdr)


def _kpi_strip(doc: Document) -> None:
    """Franja de resultados clave al inicio."""
    tbl = doc.add_table(rows=2, cols=3)
    _style_table(tbl)
    headers = ["Itron vs ultrasónico", "Itron vs app WES", "Config. final"]
    values = [
        f"Error {_fmt(ERROR_PCT, 1)} %",
        f"Error {_fmt(ERROR_B_PCT, 1)} %",
        f"Ø {DIAMETRO_FINAL_MM} mm · mortero {ESPESOR_MORTERO_MM} mm",
    ]
    subs = [
        f"{_fmt(ITRON_DELTA)} / {_fmt(US_DELTA)} m³",
        f"{_fmt(ITRON_B_DELTA)} / {_fmt(APP_B_DELTA)} m³",
        f"M45 tope {_fmt(M45_MAXIMO_EQUIPO, 1)}",
    ]
    for j, h in enumerate(headers):
        _shade(tbl.rows[0].cells[j], _NAVY_HEX)
        _cell_text(tbl.rows[0].cells[j], h, bold=True, size=9, color=_WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        _shade(tbl.rows[1].cells[j], _OK_HEX)
        cell = tbl.rows[1].cells[j]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p.add_run(values[j] + "\n")
        _font(r1, size=12, bold=True, color=_OK_TEXT)
        r2 = p.add_run(subs[j])
        _font(r2, size=8, color=_MUTED)
    doc.add_paragraph()


def _fill_header_styled(tbl, headers: list[str]) -> None:
    for j, h in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        _shade(cell, _NAVY_HEX)
        _cell_text(cell, h, bold=True, size=9, color=_WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)


def _fill_row_styled(tbl, i: int, values: list[str], *, bold: bool = False, accent: bool = False) -> None:
    for j, val in enumerate(values):
        cell = tbl.rows[i].cells[j]
        if accent:
            _shade(cell, _OK_HEX)
        elif i % 2 == 0:
            _shade(cell, _ROW_ALT_HEX)
        align = WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
        _cell_text(
            cell,
            val,
            bold=bold or accent,
            size=9.5,
            color=_OK_TEXT if accent else _BLACK,
            align=align,
        )


def _tabla_simple(doc: Document, headers: list[str], rows: list[list[str]], *, highlight_last: bool = False) -> None:
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.style = "Table Grid"
    _style_table(tbl)
    _fill_header_styled(tbl, headers)
    for i, row in enumerate(rows, start=1):
        last = highlight_last and i == len(rows)
        _fill_row_styled(tbl, i, row, bold=last, accent=last)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(4)


def _merge_title_styled(tbl, text: str, cols: int) -> None:
    cell = tbl.rows[0].cells[0]
    for j in range(1, cols):
        cell.merge(tbl.rows[0].cells[j])
    _shade(cell, _NAVY_HEX)
    _cell_text(cell, text, bold=True, size=10, color=_WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)


def _fotos_pequenas(doc: Document, pares: list[tuple[Path, str]], out_dir: Path) -> None:
    n = len(pares)
    if n == 0:
        return
    tbl = doc.add_table(rows=2, cols=n)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, (src, caption) in enumerate(pares):
        if not src.is_file():
            continue
        foto = _foto_horizontal(src, out_dir)
        cell = tbl.rows[0].cells[j]
        _shade(cell, "FAFBFC")
        _set_cell_borders(cell, "CFD8DC", "6")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        # Ancho acotado; fotos ya rotadas con odómetro horizontal
        p.add_run().add_picture(str(foto), width=Cm(FOTO_ANCHO_CM))
        cap = tbl.rows[1].cells[j]
        _shade(cap, _NAVY_LIGHT_HEX)
        _set_cell_borders(cap, "CFD8DC", "6")
        _cell_text(cap, caption, size=8, color=_NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    _keep_table_on_one_page(tbl)
    doc.add_paragraph()


def _tabla_validacion(
    doc: Document,
    titulo: str,
    analisis: str,
    fecha_ini: str,
    lectura_ini: float,
    fecha_fin: str,
    lectura_fin: float,
    consumo: float,
    horas_label: str,
) -> None:
    headers = [
        "ANÁLISIS",
        "FECHA INICIAL",
        "LECTURA (m³)",
        "FECHA FINAL",
        "LECTURA (m³)",
        "CONSUMO m³",
        "HORAS",
    ]
    tbl = doc.add_table(rows=3, cols=7)
    tbl.style = "Table Grid"
    _style_table(tbl)
    _merge_title_styled(tbl, titulo, 7)
    for j, h in enumerate(headers):
        cell = tbl.rows[1].cells[j]
        _shade(cell, _NAVY_LIGHT_HEX)
        _cell_text(cell, h, bold=True, size=8, color=_NAVY, align=WD_ALIGN_PARAGRAPH.CENTER)
    vals = [
        analisis,
        fecha_ini,
        _fmt(lectura_ini),
        fecha_fin,
        _fmt(lectura_fin),
        _fmt(consumo),
        horas_label,
    ]
    for j, v in enumerate(vals):
        align = WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
        _cell_text(tbl.rows[2].cells[j], v, size=9, align=align, bold=(j == 5))
        if j == 5:
            _shade(tbl.rows[2].cells[j], _OK_HEX)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(2)


def _tabla_error(doc: Document, filas: list[tuple[str, str]]) -> None:
    tbl = doc.add_table(rows=len(filas), cols=2)
    tbl.style = "Table Grid"
    _style_table(tbl)
    for i, (k, v) in enumerate(filas):
        last = i == len(filas) - 1
        if last:
            _shade(tbl.rows[i].cells[0], _OK_HEX)
            _shade(tbl.rows[i].cells[1], _OK_HEX)
            _cell_text(tbl.rows[i].cells[0], k, bold=True, size=10, color=_OK_TEXT)
            _cell_text(tbl.rows[i].cells[1], v, bold=True, size=11, color=_OK_TEXT, align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            if i % 2 == 1:
                _shade(tbl.rows[i].cells[0], _ROW_ALT_HEX)
                _shade(tbl.rows[i].cells[1], _ROW_ALT_HEX)
            _cell_text(tbl.rows[i].cells[0], k, size=10)
            _cell_text(tbl.rows[i].cells[1], v, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()


def generar_informe(out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_docx = out_dir / f"Informe_Validacion_Matriz_ESVAL_Fundo_Zapallar_{stamp}.docx"

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.6)
    section.left_margin = Cm(1.9)
    section.right_margin = Cm(1.9)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)

    try:
        from generar_reporte_word import add_logo_to_header

        add_logo_to_header(doc)
    except Exception as exc:
        print(f"[WARN] Logo: {exc}")

    # Portada compacta
    title = doc.add_heading("Informe de validación — Matriz ESVAL · Fundo Zapallar", level=0)
    title.paragraph_format.space_after = Pt(2)
    for run in title.runs:
        run.font.color.rgb = _NAVY
        run.font.name = "Calibri"
        run.font.size = Pt(18)

    meta = doc.add_paragraph()
    meta.paragraph_format.space_after = Pt(4)
    r = meta.add_run(
        f"{EMPRESA}  ·  {NODO_NOMBRE} ({NODE_ID})  ·  Informe final\n"
        f"Validaciones: {VAL_FECHA} {HORA_INI}–{HORA_FIN}  |  "
        f"{VAL_FECHA_B_INI} {HORA_B_INI} → {VAL_FECHA_B_FIN} {HORA_B_FIN}"
    )
    _font(r, size=9, color=_MUTED)
    _hr(doc)

    _kpi_strip(doc)

    _h(doc, "1. Objetivo", level=1)
    _p(
        doc,
        "Documentar las acciones realizadas en terreno sobre el medidor ultrasónico de "
        f"{NODO_NOMBRE} ante un desvío respecto del medidor de turbina Itron de ESVAL, "
        "y presentar el cálculo de validación del periodo revisado.",
    )

    _h(doc, "2. Resumen", level=1)
    _p(
        doc,
        "La calidad de señal (DN/UP y Q) estaba dentro del umbral del fabricante. "
        f"La primera comparación volumétrica mostró un desvío importante "
        f"({LITROS_ESVAL_ITRON_1} L Itron vs {LITROS_ULTRASONIDO_1} L ultrasónico). "
        "Se verificaron cables, se limpió la tubería y se renovó la silicona; "
        "luego se corrigió la configuración de cañería y el factor de escala. "
        f"Las validaciones posteriores arrojan errores de {_fmt(ERROR_PCT, 1)} % "
        f"(Itron vs ultrasónico) y {_fmt(ERROR_B_PCT, 1)} % (Itron vs app), "
        "aceptables según tolerancias de fabricante.",
    )

    _h(doc, "3. Actividades realizadas", level=1)

    _h(doc, "3.1 Calidad de señal", level=2)
    _p(
        doc,
        "En el menú de diagnóstico se evaluó la calidad de señal entre transductores (DN/UP) "
        "y la calidad de sonido (Q). Ambos valores iniciales superaban el umbral del fabricante "
        f"(> {UMBRAL_FABRICANTE}).",
    )
    _tabla_simple(
        doc,
        ["Parámetro", "Inicial", "Tras limpieza + silicona", "Recomendación"],
        [
            ["DN / UP", f"{DN_UP_INICIAL}", f"{DN_UP_POST}", f"> {UMBRAL_FABRICANTE}"],
            ["Q", f"{Q_INICIAL}", f"{Q_POST}", f"> {UMBRAL_FABRICANTE}"],
        ],
    )

    _h(doc, "3.2 Comparación inicial por pulso", level=2)
    _p(
        doc,
        "Se contrastó el volumen del ultrasónico WES contra el medidor de turbina Itron "
        "(referencia ESVAL) en la misma ventana de prueba.",
    )
    delta0 = LITROS_ESVAL_ITRON_1 - LITROS_ULTRASONIDO_1
    _tabla_simple(
        doc,
        ["Equipo", "Volumen", "Observación"],
        [
            ["Itron / ESVAL", f"{LITROS_ESVAL_ITRON_1} L", "Referencia"],
            ["Ultrasónico WES", f"{LITROS_ULTRASONIDO_1} L", "Misma ventana"],
            ["Diferencia", f"{delta0} L", "Desvío relevante"],
        ],
        highlight_last=True,
    )

    _h(doc, "3.3 Continuidad de cables", level=2)
    _p(
        doc,
        "Se verificó la continuidad eléctrica de los cables de los transductores. "
        "Resultado: correcta. No explica el desvío de volumen.",
    )

    _h(doc, "3.4 Limpieza de tubería y silicona", level=2)
    _p(
        doc,
        "Antes de renovar la silicona se limpió la superficie de la tubería en la zona de "
        "los transductores. Tras el cambio de silicona mejoraron DN/UP y Q, pero no se resolvió "
        "la discrepancia volumétrica.",
    )
    _bullet(doc, f"DN / UP: {DN_UP_INICIAL} → {DN_UP_POST}")
    _bullet(doc, f"Q: {Q_INICIAL} → {Q_POST}")

    _h(doc, "3.5 Configuración de cañería", level=2)
    _p(
        doc,
        "Se revisó el material en obra y se contrastó con la información de la sanitaria ESVAL.",
    )
    _bullet(doc, f"Marcado físico: {MARCADO_FISICO} ({MATERIAL}).")
    _bullet(
        doc,
        f"Diámetro según ESVAL: {_fmt(DIAMETRO_ESVAL_MM, 0)} mm "
        f"(estaba en {_fmt(DIAMETRO_CONFIG_ERRONEO_MM, 0)} mm).",
    )
    _bullet(
        doc,
        f"Lining: mortero inyectado — menú de cobertura con espesor {ESPESOR_MORTERO_MM} mm.",
    )
    _p(
        doc,
        f"Tras aplicar {DIAMETRO_ESVAL_MM} mm + mortero {ESPESOR_MORTERO_MM} mm, "
        f"la diferencia bajó a aproximadamente {DIFERENCIA_LITROS_POST_CONFIG} L.",
    )

    _h(doc, "3.6 Factor de escala M45", level=2)
    _p(
        doc,
        "Con la prueba de 1 minuto (caudal turbina vs caudal ultrasónico menú 00) se calculó "
        "el nuevo factor de escala:",
    )
    _p(
        doc,
        f"Nuevo M45 = {_fmt(M45_FACTOR_ACTUAL, 4)} × "
        f"({_fmt(M45_CAUDAL_TURBINA_LPM, 0)} / {_fmt(M45_CAUDAL_ULTRASONIDO_LPM, 2)}) = "
        f"{_fmt(M45_CALCULADO, 4)}.",
    )
    _p(
        doc,
        f"El valor calculado supera el máximo del equipo ({_fmt(M45_MAXIMO_EQUIPO, 1)}). "
        f"Se seteó el tope. La diferencia bajó a ~{DIFERENCIA_LITROS_POST_M45} L.",
    )

    _h(doc, "3.7 Ajuste final de diámetro", level=2)
    _p(
        doc,
        f"Según recomendación del manual, se aumentó el diámetro un {INCREMENTO_DIAMETRO_PCT} %: "
        f"{DIAMETRO_ESVAL_MM} → {DIAMETRO_FINAL_MM} mm. Tras ese ajuste ambos medidores "
        "quedaron en el mismo ciclo de lectura en terreno.",
    )
    _tabla_simple(
        doc,
        ["Parámetro", "Antes / intermedio", "Final"],
        [
            ["Material", "Sin validar", MATERIAL],
            ["Diámetro (mm)", f"{DIAMETRO_CONFIG_ERRONEO_MM} → {DIAMETRO_ESVAL_MM}", f"{DIAMETRO_FINAL_MM}"],
            ["Cobertura", "Sin mortero", f"Mortero {ESPESOR_MORTERO_MM} mm"],
            [
                "Factor M45",
                f"{_fmt(M45_FACTOR_ACTUAL, 4)} → calc. {_fmt(M45_CALCULADO, 4)}",
                f"Tope {_fmt(M45_MAXIMO_EQUIPO, 1)}",
            ],
        ],
    )

    _h(doc, "4. Cálculo de validación", level=1)

    _h(doc, "4.1 Itron vs ultrasónico WES (22-09, 11:50 → 15:00)", level=2)
    _p(
        doc,
        "Lecturas del medidor Itron y del ultrasónico WES en la misma ventana. "
        "En el ultrasónico el NET se convierte a m³ × 0,01.",
    )
    _tabla_validacion(
        doc,
        "Validación medidor Itron (ESVAL)",
        NODO_NOMBRE,
        f"{VAL_FECHA} {HORA_INI}",
        ITRON_INI,
        f"{VAL_FECHA} {HORA_FIN}",
        ITRON_FIN,
        ITRON_DELTA,
        "3,2",
    )
    _tabla_validacion(
        doc,
        "Validación medidor ultrasónico WES (NET × 0,01)",
        NODO_NOMBRE,
        f"{VAL_FECHA} {HORA_INI}",
        US_INI,
        f"{VAL_FECHA} {HORA_FIN}",
        US_FIN,
        US_DELTA,
        "3,2",
    )
    _tabla_error(
        doc,
        [
            ("Total Itron (lectura)", f"{_fmt(ITRON_DELTA)} m³"),
            ("Total ultrasónico WES", f"{_fmt(US_DELTA)} m³"),
            ("% Error", f"{_fmt(ERROR_PCT, 1)} %"),
        ],
    )
    _p(
        doc,
        f"Error = 1 − ({_fmt(ITRON_DELTA)} / {_fmt(US_DELTA)}) = {_fmt(ERROR_PCT, 1)} %. "
        f"Aceptable frente a tolerancias (ultrasónico ±{TOL_US_PCT} %, Itron ±{TOL_ITRON_PCT} %).",
    )

    _h(doc, "4.2 Itron vs app WES (22-09 15:00 → 23-09 17:00)", level=2)
    _p(
        doc,
        "Validación con lecturas fotográficas del medidor Itron y el consumo de la app WES. "
        "App: horas 16→23 del 22-09 + horas 00→16 del 23-09.",
        keep_next=True,
    )
    _fotos_pequenas(
        doc,
        [
            (FOTO_ITRON_1500, f"Itron · {VAL_FECHA_B_INI} {HORA_B_INI}"),
            (FOTO_ITRON_1700, f"Itron · {VAL_FECHA_B_FIN} {HORA_B_FIN}"),
        ],
        out_dir,
    )
    _tabla_validacion(
        doc,
        "Validación Matriz ESVAL — medidor Itron",
        NODO_NOMBRE,
        f"{VAL_FECHA_B_INI} {HORA_B_INI}",
        ITRON_B_INI,
        f"{VAL_FECHA_B_FIN} {HORA_B_FIN}",
        ITRON_B_FIN,
        ITRON_B_DELTA,
        "26",
    )
    _tabla_error(
        doc,
        [
            ("Total App WES", f"{_fmt(APP_B_DELTA)} m³"),
            ("Total Lectura Itron", f"{_fmt(ITRON_B_DELTA)} m³"),
            ("% Error", f"{_fmt(ERROR_B_PCT, 1)} %"),
        ],
    )
    _p(
        doc,
        f"Error entre lectura Itron y app WES: {_fmt(ERROR_B_PCT, 1)} % "
        f"(1 − {_fmt(ITRON_B_DELTA)}/{_fmt(APP_B_DELTA)}).",
    )

    _h(doc, "5. Conclusión", level=1)
    _p(
        doc,
        "El desvío inicial no se debió a calidad de señal ni a cableado. La limpieza y la silicona "
        "mejoraron los indicadores de señal, pero no el volumen. La corrección de diámetro, mortero "
        f"y factor de escala (tope {_fmt(M45_MAXIMO_EQUIPO, 1)}), más el ajuste final a "
        f"{DIAMETRO_FINAL_MM} mm, alineó ambos medidores en terreno. "
        f"Validación Itron vs ultrasónico: {_fmt(ERROR_PCT, 1)} %. "
        f"Validación Itron vs app: {_fmt(ERROR_B_PCT, 1)} %. "
        "Ambos resultados son aceptables según las tolerancias de fabricante.",
    )

    _hr(doc)
    pie = doc.add_paragraph()
    pie.paragraph_format.space_before = Pt(2)
    r = pie.add_run(
        f"WES · Informe final · {EMPRESA} · {datetime.now().strftime('%d-%m-%Y %H:%M')}"
    )
    _font(r, size=8, color=_MUTED)

    doc.save(out_docx)
    print(f"[OK] Word: {out_docx}")
    print(
        f"[CIFRAS] A: {ITRON_DELTA}/{US_DELTA}={ERROR_PCT}% | "
        f"B: {ITRON_B_DELTA}/{APP_B_DELTA}={ERROR_B_PCT}%"
    )
    return out_docx


def _intentar_pdf(docx_path: Path) -> Path | None:
    pdf_path = docx_path.with_suffix(".pdf")
    for cmd in ("soffice", "libreoffice"):
        bin_path = shutil.which(cmd)
        if not bin_path:
            continue
        try:
            subprocess.run(
                [
                    bin_path,
                    "--headless",
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
        except Exception:
            pass
    return None


def main() -> int:
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = (
        ROOT
        / "reports"
        / "Fundo_Zapallar"
        / "Informes_Tecnicos"
        / f"calidad_senal_ultrasonido_{stamp}"
    )
    docx_path = generar_informe(out_dir)
    _intentar_pdf(docx_path)
    print(f"[OK] Carpeta: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
