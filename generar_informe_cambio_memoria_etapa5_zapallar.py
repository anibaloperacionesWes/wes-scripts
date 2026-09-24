"""
Informe de validación — Etapa N°5 Fundo Zapallar.

Estilo alineado a Informe_Validacion_Matriz_ESVAL_Fundo_Zapallar_FINAL:
  - Título + meta + franja KPI (verde #E8F5E9 / #1B5E20)
  - Tablas: header azul, secundario #D6E3F0/#1F4788, filas alt #F5F8FB
  - Fotos relojería ~1,55" lado a lado; §4 compacto en la 2ª hoja
  - cantSplit + tblHeader + keepNext (tablas no se cortan entre hojas)
  - Validación: hora 14 completa → hasta 17:00; §4 y §5 en hojas propias

Uso:
  python generar_informe_cambio_memoria_etapa5_zapallar.py
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from PIL import Image

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "reports" / "Fundo_Zapallar" / "Informes_Tecnicos"
FOTOS_DIR = OUT_DIR / "fotos_etapa5"

try:
    from generar_reporte_word import CHILE_TZ, acl_node_base_url
except Exception:
    CHILE_TZ = ZoneInfo("America/Santiago")

    def acl_node_base_url() -> str:
        return "http://104.248.53.141:7003/wes/api/acl-node/v1"

COMPANY = "Fundo Zapallar"
COMPANY_ID = "000027"
NODE_ID = "000027-03"
NODE_NAME = "Etapa N°5"
FECHA_INTERVENCION = date(2026, 9, 22)
DIAMETRO = "DN90 fierro dúctil"
CAUDAL_MAX_REF_M3H = 60.0
COLOR_TITULO = RGBColor(0x1F, 0x47, 0x88)
COLOR_META = RGBColor(0x64, 0x6E, 0x78)
COLOR_TEXTO = RGBColor(0x28, 0x28, 0x28)
COLOR_KPI = RGBColor(0x1B, 0x5E, 0x20)
COLOR_HEADER_FILL = "1F4788"
COLOR_KPI_FILL = "E8F5E9"
COLOR_HEADER2_FILL = "D6E3F0"  # encabezado secundario (letras azules)
COLOR_ALT_FILL = "F5F8FB"
COLOR_FOTO_FILL = "FAFBFC"
FOTO_ANCHO = Inches(1.55)  # compacto: §4 completo en la 2ª hoja


# Lecturas mecánicas (fotos terreno)
LECTURA_AYER_M3 = 5144.0
LECTURA_AYER_DT = datetime(2026, 9, 22, 14, 30, tzinfo=CHILE_TZ)
LECTURA_HOY_M3 = 5177.0
LECTURA_HOY_DT = datetime(2026, 9, 23, 16, 54, tzinfo=CHILE_TZ)
# Suma WES hasta las 17:00 (incluye hora 16)
WES_HASTA_DT = datetime(2026, 9, 23, 17, 0, tzinfo=CHILE_TZ)

# Hueco 22/09 leído directo de la placa (m³/h).
# Extracción listó H:15:00 dos veces (0,60 y 2,90): se interpreta 0,60 como hora 14
# (coherente con lectura a las 14:30) y 2,90 como hora 15.
PLACA_HUECO_22: Dict[int, float] = {
    14: 0.60,
    15: 2.90,
    16: 5.50,
    17: 2.60,
    18: 0.90,
    19: 0.00,
    20: 0.00,
    21: 0.00,
    22: 0.00,
    23: 0.00,
}

MEDIDOR_MARCA = "Sensus"
MEDIDOR_MODELO = "MeiStream Plus 100"
MEDIDOR_SERIE = "8 SEN01 2370 9030"
MEDIDOR_Q3_M3H = 100.0
HRI_MODELO = "HRI-Mei B4 500 ms"
HRI_SERIE = "31730485"
HRI_PULSO_DN40_125_L = 100

FOTO_AYER = FOTOS_DIR / "lectura_20260922_1430_sensus_5144.png"
FOTO_HOY = FOTOS_DIR / "lectura_20260923_1654_sensus_5177.png"
FOTO_AYER_RELOJ = FOTOS_DIR / "lectura_20260922_1430_sensus_5144_reloj.jpg"
FOTO_HOY_RELOJ = FOTOS_DIR / "lectura_20260923_1654_sensus_5177_reloj.jpg"
# Encuadre vertical tipo Informe Validación ESVAL FINAL (cara circular centrada)
# (cx, cy, half_w, half_h) sobre original 960x1280
CROP_AYER = (480, 580, 320, 380)
CROP_HOY = (480, 540, 320, 380)
FOTO_SIZE = (768, 900)


@dataclass
class Validacion:
    lectura_ayer: float
    lectura_hoy: float
    delta_mecanico: float
    wes_m3: float
    detalle: List[Tuple[str, float, str]]
    hueco_horas_api: List[str]
    diferencia_m3: float
    diferencia_pct: float
    estado: str


def _set_run_font(run, *, bold: bool = False, size: int = 11, color: RGBColor | None = None) -> None:
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    if color is not None:
        run.font.color.rgb = color


def _add_title(doc: Document, text: str) -> None:
    p = doc.add_paragraph(style="Title")
    for child in list(p._element):
        if child.tag.endswith("}r"):
            p._element.remove(child)
    r = p.add_run(text)
    _set_run_font(r, bold=True, size=18, color=COLOR_TITULO)


def _add_meta(doc: Document, text: str, *, size: int = 9) -> None:
    p = doc.add_paragraph()
    r = p.add_run(text)
    _set_run_font(r, size=size, color=COLOR_META)


def _add_h(doc: Document, text: str, level: int = 1, *, compact: bool = False) -> None:
    h = doc.add_heading(text, level=level)
    size = 14 if level == 1 else 12
    if compact:
        size = 12 if level == 1 else 11
        h.paragraph_format.space_before = Pt(6)
        h.paragraph_format.space_after = Pt(2)
    for run in h.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(size)
        run.font.color.rgb = COLOR_TITULO
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    # Mantener título junto al contenido / tabla siguiente
    _p_keep(h, with_next=True, lines=True)


def _add_body(doc: Document, text: str, *, size: int = 11, compact: bool = False) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if compact:
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        size = min(size, 10)
    r = p.add_run(text)
    _set_run_font(r, size=size, color=COLOR_TEXTO)


def _add_line(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    r = p.add_run(text)
    _set_run_font(r, size=10.5, color=COLOR_TEXTO)


def _shade_cell(cell, hex_color: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def _p_keep(paragraph, *, with_next: bool = True, lines: bool = True) -> None:
    """keepNext / keepLines en un párrafo (no partir bloques)."""
    pPr = paragraph._p.get_or_add_pPr()
    if with_next and pPr.find(qn("w:keepNext")) is None:
        pPr.append(OxmlElement("w:keepNext"))
    if lines and pPr.find(qn("w:keepLines")) is None:
        pPr.append(OxmlElement("w:keepLines"))


def _row_no_partir(row, *, as_header: bool = True) -> None:
    """Fila indivisible: cantSplit + tblHeader (igual que FINAL ESVAL)."""
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))
    if as_header and trPr.find(qn("w:tblHeader")) is None:
        hdr = OxmlElement("w:tblHeader")
        hdr.set(qn("w:val"), "true")
        trPr.append(hdr)


def _table_no_partir(table) -> None:
    """
    Evita que la tabla se corte entre hojas:
      - cantSplit + tblHeader en todas las filas (estilo FINAL)
      - keepNext en párrafos de filas intermedias (encadena filas)
      - keepLines en todos los párrafos de celdas
    """
    n = len(table.rows)
    for i, row in enumerate(table.rows):
        _row_no_partir(row, as_header=True)
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(1)
                # Encadena filas: todas menos la última quedan unidas a la siguiente
                _p_keep(p, with_next=(i < n - 1), lines=True)


def _keep_prev_with_table(doc: Document) -> None:
    """El párrafo/título previo a la tabla no queda huérfano en la hoja anterior."""
    if doc.paragraphs:
        _p_keep(doc.paragraphs[-1], with_next=True, lines=True)


def _set_cell_text(
    cell,
    text: str,
    *,
    bold: bool = False,
    size: int = 9,
    color: RGBColor | None = None,
    fill: str | None = None,
) -> None:
    p = cell.paragraphs[0]
    for child in list(p._element):
        if child.tag.endswith("}r") or child.tag.endswith("}hyperlink"):
            p._element.remove(child)
    r = p.add_run(text)
    _set_run_font(r, bold=bold, size=size, color=color)
    if fill:
        _shade_cell(cell, fill)


def _set_cell_kpi_value(cell, linea1: str, linea2: str) -> None:
    """Valor KPI: línea 1 verde bold + línea 2 gris (como FINAL)."""
    p = cell.paragraphs[0]
    for child in list(p._element):
        if child.tag.endswith("}r") or child.tag.endswith("}hyperlink"):
            p._element.remove(child)
    r1 = p.add_run(f"{linea1}\n")
    _set_run_font(r1, bold=True, size=12, color=COLOR_KPI)
    r2 = p.add_run(linea2)
    _set_run_font(r2, bold=False, size=8, color=COLOR_META)
    _shade_cell(cell, COLOR_KPI_FILL)


def _set_table_borders(table, color: str = "D0D5DD") -> None:
    """Bordes suaves estilo Informe Validación ESVAL FINAL."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is not None:
        tblPr.remove(borders)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def _add_kpi_banner(doc: Document, cards: List[Tuple[str, str, str]]) -> None:
    """Franja KPI: header azul + valor verde sobre fondo #E8F5E9 (FINAL)."""
    _keep_prev_with_table(doc)
    t = doc.add_table(rows=2, cols=len(cards))
    _set_table_borders(t, color="1F4788")
    for j, (titulo, linea1, linea2) in enumerate(cards):
        _set_cell_text(
            t.rows[0].cells[j],
            titulo,
            bold=True,
            size=9,
            color=RGBColor(255, 255, 255),
            fill=COLOR_HEADER_FILL,
        )
        _set_cell_kpi_value(t.rows[1].cells[j], linea1, linea2)
    _table_no_partir(t)
    doc.add_paragraph()


def _add_tabla_simple(doc: Document, filas: List[Tuple[str, ...]]) -> None:
    cols = len(filas[0])
    _keep_prev_with_table(doc)
    t = doc.add_table(rows=len(filas), cols=cols)
    _set_table_borders(t)
    for i, row in enumerate(filas):
        for j, val in enumerate(row):
            if i == 0:
                _set_cell_text(
                    t.rows[i].cells[j],
                    val,
                    bold=True,
                    size=9,
                    color=RGBColor(255, 255, 255),
                    fill=COLOR_HEADER_FILL,
                )
            else:
                fill = COLOR_ALT_FILL if i % 2 == 0 else None
                _set_cell_text(
                    t.rows[i].cells[j],
                    val,
                    bold=False,
                    size=10,
                    color=COLOR_TEXTO,
                    fill=fill,
                )
    _table_no_partir(t)


def _crop_relojeria(
    src: Path,
    dst: Path,
    crop: Tuple[int, int, int, int],
) -> Path:
    """Recorte vertical centrado en la relojería (mismo encuadre que ESVAL FINAL)."""
    from PIL import ImageEnhance, ImageFilter, ImageOps

    FOTOS_DIR.mkdir(parents=True, exist_ok=True)
    if not src.is_file():
        return dst
    im = Image.open(src).convert("RGB")
    cx, cy, half_w, half_h = crop
    box = (
        max(0, cx - half_w),
        max(0, cy - half_h),
        min(im.width, cx + half_w),
        min(im.height, cy + half_h),
    )
    c = im.crop(box)
    c = ImageOps.fit(
        c,
        FOTO_SIZE,
        method=Image.Resampling.LANCZOS,
        centering=(0.5, 0.42),
    )
    c = c.filter(ImageFilter.UnsharpMask(radius=1.2, percent=110, threshold=2))
    c = ImageEnhance.Contrast(c).enhance(1.08)
    c = ImageEnhance.Color(c).enhance(1.05)
    c.save(dst, quality=95, optimize=True)
    return dst


def _add_foto_tabla(doc: Document, path: Path, caption: str) -> None:
    """Una foto por tabla (fallback); caption azul sobre #D6E3F0 como FINAL."""
    if not path.is_file():
        return
    _keep_prev_with_table(doc)
    tbl = doc.add_table(rows=2, cols=1)
    _set_table_borders(tbl, color="E5E7EB")
    cell_img = tbl.rows[0].cells[0]
    _shade_cell(cell_img, COLOR_FOTO_FILL)
    cell_img.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cell_img.paragraphs[0].add_run()
    run.add_picture(str(path), width=FOTO_ANCHO)
    cell_cap = tbl.rows[1].cells[0]
    _shade_cell(cell_cap, COLOR_HEADER2_FILL)
    cell_cap.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for child in list(cell_cap.paragraphs[0]._element):
        if child.tag.endswith("}r"):
            cell_cap.paragraphs[0]._element.remove(child)
    r = cell_cap.paragraphs[0].add_run(caption)
    _set_run_font(r, bold=True, size=8, color=COLOR_TITULO)
    _table_no_partir(tbl)
    # sin párrafo vacío: compactar §4 en una sola hoja


def _add_fotos_lado_a_lado(
    doc: Document,
    path_izq: Path,
    caption_izq: str,
    path_der: Path,
    caption_der: str,
) -> None:
    """Dos fotos en una sola tabla 2×2 (estilo FINAL: ~2,28\" + caption azul)."""
    if not path_izq.is_file() and not path_der.is_file():
        return
    if not path_izq.is_file() or not path_der.is_file():
        # fallback a tablas separadas si falta una
        if path_izq.is_file():
            _add_foto_tabla(doc, path_izq, caption_izq)
        if path_der.is_file():
            _add_foto_tabla(doc, path_der, caption_der)
        return

    _keep_prev_with_table(doc)
    tbl = doc.add_table(rows=2, cols=2)
    _set_table_borders(tbl, color="E5E7EB")
    for j, (path, caption) in enumerate(
        ((path_izq, caption_izq), (path_der, caption_der))
    ):
        cell_img = tbl.rows[0].cells[j]
        _shade_cell(cell_img, COLOR_FOTO_FILL)
        cell_img.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell_img.paragraphs[0].add_run()
        run.add_picture(str(path), width=FOTO_ANCHO)

        cell_cap = tbl.rows[1].cells[j]
        _shade_cell(cell_cap, COLOR_HEADER2_FILL)
        cell_cap.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for child in list(cell_cap.paragraphs[0]._element):
            if child.tag.endswith("}r"):
                cell_cap.paragraphs[0]._element.remove(child)
        r = cell_cap.paragraphs[0].add_run(caption)
        _set_run_font(r, bold=True, size=8, color=COLOR_TITULO)

    _table_no_partir(tbl)
    # sin párrafo vacío tras fotos (compactar §4)


def _add_tabla_validacion_7(
    doc: Document,
    titulo: str,
    analisis: str,
    fecha_ini: str,
    lectura_ini: str,
    fecha_fin: str,
    lectura_fin: str,
    consumo: str,
    horas: str,
) -> None:
    """Tabla 7 columnas estilo FINAL: título azul/blanco + headers #D6E3F0/#1F4788."""
    _keep_prev_with_table(doc)
    t = doc.add_table(rows=3, cols=7)
    _set_table_borders(t)
    t.rows[0].cells[0].merge(t.rows[0].cells[6])
    _set_cell_text(
        t.rows[0].cells[0],
        titulo,
        bold=True,
        size=9,
        color=RGBColor(255, 255, 255),
        fill=COLOR_HEADER_FILL,
    )
    headers = [
        "ANÁLISIS",
        "FECHA INICIAL",
        "LECTURA (m³)",
        "FECHA FINAL",
        "LECTURA (m³)",
        "CONSUMO m³",
        "HORAS",
    ]
    for j, htxt in enumerate(headers):
        _set_cell_text(
            t.rows[1].cells[j],
            htxt,
            bold=True,
            size=8,
            color=COLOR_TITULO,
            fill=COLOR_HEADER2_FILL,
        )
    vals = [analisis, fecha_ini, lectura_ini, fecha_fin, lectura_fin, consumo, horas]
    for j, v in enumerate(vals):
        _set_cell_text(t.rows[2].cells[j], v, bold=False, size=8, color=COLOR_TEXTO)
    _table_no_partir(t)


def _add_tabla_error(doc: Document, filas: List[Tuple[str, str]]) -> None:
    """Fila final % Error en verde #E8F5E9 / #1B5E20 (FINAL)."""
    _keep_prev_with_table(doc)
    t = doc.add_table(rows=len(filas), cols=2)
    _set_table_borders(t)
    last = len(filas) - 1
    for i, (a, b) in enumerate(filas):
        if i == last:
            _set_cell_text(
                t.rows[i].cells[0], a, bold=True, size=10, color=COLOR_KPI, fill=COLOR_KPI_FILL
            )
            _set_cell_text(
                t.rows[i].cells[1], b, bold=True, size=10, color=COLOR_KPI, fill=COLOR_KPI_FILL
            )
        else:
            fill = COLOR_ALT_FILL if i % 2 == 1 else None
            _set_cell_text(
                t.rows[i].cells[0], a, bold=False, size=10, color=COLOR_TEXTO, fill=fill
            )
            _set_cell_text(
                t.rows[i].cells[1], b, bold=False, size=10, color=COLOR_TEXTO, fill=fill
            )
    _table_no_partir(t)


def _fmt(n: float, dec: int = 1) -> str:
    return f"{n:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def obtener_api_por_hora() -> Dict[datetime, float]:
    base = acl_node_base_url()
    out: Dict[datetime, float] = {}
    for d in (date(2026, 9, 21), date(2026, 9, 22), date(2026, 9, 23)):
        ds = d.strftime("%d%m%Y")
        url = f"{base}/nodes/{NODE_ID}/dates.measures.csv"
        r = requests.get(url, params=[("start", ds), ("end", ds)], timeout=45)
        r.raise_for_status()
        lines = [ln for ln in r.text.strip().splitlines() if ln.strip()]
        start = 1 if lines and "TIME" in lines[0].upper() else 0
        for line in lines[start:]:
            parts = line.split(",")
            if len(parts) < 2:
                continue
            raw = parts[0].strip().replace("Z", "+00:00")
            dt = datetime.fromisoformat(raw)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            out[dt.astimezone(CHILE_TZ)] = float(parts[1])
    return out


def valor_hora(dt: datetime, api: Dict[datetime, float]) -> Tuple[float, str]:
    """Prioriza valores de placa en el hueco del 22/09; si no, API."""
    if dt.date() == date(2026, 9, 22) and dt.hour in PLACA_HUECO_22:
        return PLACA_HUECO_22[dt.hour], "placa (hueco)"
    if dt in api:
        return api[dt], "API WES"
    return 0.0, "sin dato (0)"


def calcular_validacion(lectura_ayer: float, lectura_hoy: float) -> Validacion:
    api = obtener_api_por_hora()
    detalle: List[Tuple[str, float, str]] = []
    total = 0.0
    hueco_api: List[str] = []

    inicio_hora = LECTURA_AYER_DT.replace(minute=0, second=0, microsecond=0)
    v14, src14 = valor_hora(inicio_hora, api)
    detalle.append(
        (
            inicio_hora.strftime("%d/%m/%Y %H:%M"),
            v14,
            f"hora 14 completa; {_fmt(v14, 2)} m³/h [{src14}]",
        )
    )
    total += v14
    if "API" not in src14 and inicio_hora not in api:
        hueco_api.append(inicio_hora.strftime("%d/%m/%Y %H:%M"))

    cur = inicio_hora + timedelta(hours=1)
    while cur < WES_HASTA_DT:
        v, src = valor_hora(cur, api)
        detalle.append((cur.strftime("%d/%m/%Y %H:%M"), v, src))
        total += v
        if src.startswith("placa") or src.startswith("sin"):
            if cur not in api:
                hueco_api.append(cur.strftime("%d/%m/%Y %H:%M"))
        cur += timedelta(hours=1)

    delta = lectura_hoy - lectura_ayer
    dif = delta - total
    # Estilo ESVAL: % Error = |1 − (lectura / serie)| o |1 − (serie / lectura)|
    # Se reporta el valor absoluto respecto al mayor de ambos.
    if delta and total:
        pct = abs(1.0 - (min(delta, total) / max(delta, total))) * 100.0
    else:
        pct = 0.0
    if pct <= 5:
        estado = "Aceptable"
    elif pct <= 15:
        estado = "Aceptable (diferencia moderada)"
    else:
        estado = "Revisar — diferencia relevante"

    return Validacion(
        lectura_ayer=lectura_ayer,
        lectura_hoy=lectura_hoy,
        delta_mecanico=round(delta, 2),
        wes_m3=round(total, 2),
        detalle=detalle,
        hueco_horas_api=hueco_api,
        diferencia_m3=round(dif, 2),
        diferencia_pct=round(pct, 1),
        estado=estado,
    )


def build_doc(lectura_ayer: float, lectura_hoy: float) -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    ahora = datetime.now(CHILE_TZ)
    stamp = ahora.strftime("%Y%m%d_%H%M")
    out = OUT_DIR / f"Informe_Validacion_Etapa5_Zapallar_{stamp}.docx"
    val = calcular_validacion(lectura_ayer, lectura_hoy)

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.6)
    section.left_margin = Cm(1.9)
    section.right_margin = Cm(1.9)

    _add_title(doc, f"Informe de validación — {NODE_NAME} · {COMPANY}")
    _add_meta(
        doc,
        f"{COMPANY}  ·  {NODE_NAME} ({NODE_ID})  ·  Informe final\n"
        f"Validaciones: {LECTURA_AYER_DT.strftime('%d-%m-%Y')} 14:00 → "
        f"{WES_HASTA_DT.strftime('%d-%m-%Y %H:%M')}  ·  "
        f"Intervención placa {FECHA_INTERVENCION.strftime('%d-%m-%Y')}",
    )

    dens = max(val.delta_mecanico, val.wes_m3)
    num = min(val.delta_mecanico, val.wes_m3)
    _add_kpi_banner(
        doc,
        [
            (
                "Sensus vs app WES",
                f"Error {_fmt(val.diferencia_pct, 1)} %",
                f"{_fmt(val.delta_mecanico, 2)} / {_fmt(val.wes_m3, 2)} m³",
            ),
            (
                "Lecturas Sensus",
                f"{_fmt(lectura_ayer, 0)} → {_fmt(lectura_hoy, 0)} m³",
                f"Δ {_fmt(val.delta_mecanico, 0)} m³",
            ),
            (
                "Red / medidor",
                f"DN90 · techo ≈ {CAUDAL_MAX_REF_M3H:.0f} m³/h",
                f"Q3 medidor {_fmt(MEDIDOR_Q3_M3H, 0)} m³/h",
            ),
        ],
    )

    _add_h(doc, "1. Objetivo", 1)
    _add_body(
        doc,
        "Documentar la intervención en terreno sobre la placa de Etapa N°5 ante caudales "
        "anómalos incompatibles con la red DN90, y presentar el cálculo de validación "
        "entre lecturas mecánicas del medidor Sensus y el consumo de la app WES.",
    )

    _add_h(doc, "2. Resumen", 1)
    _add_body(
        doc,
        "Se detectaron pulsos del orden de 92 y 200 m³/h. El sensor inductivo, los pulsos "
        "de revisión y los voltajes estaban correctos; el error provenía de la memoria de "
        "la placa. Se reemplazó la memoria y se normalizó el punto. La validación posterior "
        f"entre lectura Sensus y app WES arroja un error de {_fmt(val.diferencia_pct, 1)} % "
        f"({val.estado.lower()}).",
    )

    _add_h(doc, "3. Actividades realizadas", 1)

    _add_h(doc, "3.1 Revisión en terreno", 2)
    _add_body(
        doc,
        "Se verificó sensor Census / inductivo (OK), pulsos de revisión (OK) y voltajes de "
        "alimentación (correctos). El diagnóstico apuntó a falla de memoria de la placa.",
    )
    _add_line(doc, "Acción: cambio de memoria de la placa; punto reparado/normalizado.")

    _add_h(doc, "3.2 Medidor instalado", 2)
    _add_body(
        doc,
        f"Medidor {MEDIDOR_MARCA} {MEDIDOR_MODELO}, serie {MEDIDOR_SERIE} (2023), "
        f"Q3 = {_fmt(MEDIDOR_Q3_M3H, 0)} m³/h. Módulo {HRI_MODELO} (serie {HRI_SERIE}), "
        f"peso de pulso DN 40–125 = {HRI_PULSO_DN40_125_L} L/pulso.",
    )
    _add_line(doc, f"Lectura inicial: {_fmt(lectura_ayer, 0)} m³ ({LECTURA_AYER_DT.strftime('%d-%m-%Y')}).")
    _add_line(doc, f"Lectura final: {_fmt(lectura_hoy, 0)} m³ ({LECTURA_HOY_DT.strftime('%d-%m-%Y')}).")
    _add_line(doc, f"Δ mecánico: {_fmt(val.delta_mecanico, 0)} m³.")

    _add_h(doc, "3.3 Criterio hidráulico (DN90)", 2)
    _add_body(
        doc,
        f"La red es {DIAMETRO}. Techo práctico de red ≈ {CAUDAL_MAX_REF_M3H:.0f} m³/h "
        f"(~2,5 m/s). El medidor admite Q3 = {_fmt(MEDIDOR_Q3_M3H, 0)} m³/h, pero la red "
        "no puede entregar de forma realista 92–200 m³/h (error de memoria).",
    )
    _add_tabla_simple(
        doc,
        [
            ("Velocidad de referencia", "Caudal teórico DN90"),
            ("1,5 m/s", "≈ 34 m³/h"),
            ("2,0 m/s", "≈ 46 m³/h"),
            ("2,5 m/s (techo práctico)", "≈ 57–60 m³/h"),
        ],
    )

    _add_h(doc, "4. Cálculo de validación", 1, compact=True)
    # Empezar validación + fotos en hoja nueva para que las tablas no se partan
    doc.paragraphs[-1].paragraph_format.page_break_before = True
    _add_h(
        doc,
        f"4.1 Sensus vs app WES ({LECTURA_AYER_DT.strftime('%d-%m')} 14:00 → "
        f"{WES_HASTA_DT.strftime('%d-%m %H:%M')})",
        2,
        compact=True,
    )
    _add_body(
        doc,
        "Lecturas fotográficas Sensus vs consumo app WES "
        "(hora 14 completa del 22-09 → 17:00 del 23-09).",
        compact=True,
    )

    # Tabla de fotos lado a lado (compacta para caber en hoja 2)
    foto_ayer = _crop_relojeria(FOTO_AYER, FOTO_AYER_RELOJ, CROP_AYER)
    foto_hoy = _crop_relojeria(FOTO_HOY, FOTO_HOY_RELOJ, CROP_HOY)
    if foto_ayer.is_file() or foto_hoy.is_file():
        _add_fotos_lado_a_lado(
            doc,
            foto_ayer,
            f"Sensus · {LECTURA_AYER_DT.strftime('%d-%m-%Y %H:%M')}",
            foto_hoy,
            f"Sensus · {LECTURA_HOY_DT.strftime('%d-%m-%Y %H:%M')}",
        )

    # Detalle del cálculo en tablas propias (después de las fotos)
    inicio_wes = LECTURA_AYER_DT.replace(minute=0, second=0, microsecond=0)
    horas_ventana = (WES_HASTA_DT - inicio_wes).total_seconds() / 3600.0
    _add_tabla_validacion_7(
        doc,
        "Validación Etapa N°5 — medidor Sensus (lectura mecánica)",
        NODE_NAME,
        inicio_wes.strftime("%d-%m-%Y %H:%M"),
        _fmt(val.lectura_ayer, 0),
        WES_HASTA_DT.strftime("%d-%m-%Y %H:%M"),
        _fmt(val.lectura_hoy, 0),
        _fmt(val.delta_mecanico, 2),
        _fmt(horas_ventana, 0),
    )

    _add_tabla_error(
        doc,
        [
            ("Total App WES", f"{_fmt(val.wes_m3, 2)} m³"),
            ("Total Lectura Sensus", f"{_fmt(val.delta_mecanico, 2)} m³"),
            ("% Error", f"{_fmt(val.diferencia_pct, 1)} %"),
        ],
    )

    _add_body(
        doc,
        f"Error Sensus vs app WES: {_fmt(val.diferencia_pct, 1)} % "
        f"(1 − {_fmt(num, 2)}/{_fmt(dens, 2)}).",
        compact=True,
    )

    _add_h(doc, "5. Conclusión", 1)
    # Conclusión siempre en la última hoja
    doc.paragraphs[-1].paragraph_format.page_break_before = True
    _add_body(
        doc,
        "Se confirma falla de memoria de placa (sensor Sensus/HRI y voltajes OK). "
        f"Validación post-cambio: lectura mecánica {_fmt(val.delta_mecanico, 2)} m³ vs "
        f"app WES {_fmt(val.wes_m3, 2)} m³ (% error {_fmt(val.diferencia_pct, 1)} %). "
        f"{val.estado}. El agente IA de WES revisará de forma diaria que el caudal "
        "esté de acuerdo al consumo histórico del punto y al máximo de consumo "
        f"compatible con la matriz DN90 (umbral {_fmt(CAUDAL_MAX_REF_M3H, 0)} m³/h).",
    )

    _add_meta(
        doc,
        f"WES · Informe final · {COMPANY} · {ahora.strftime('%d-%m-%Y %H:%M')}",
        size=8,
    )

    doc.save(out)
    print(
        f"[VALIDACION] Δ mec={val.delta_mecanico} | serie={val.wes_m3} | "
        f"dif={val.diferencia_m3} ({val.diferencia_pct}%) | {val.estado}"
    )
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--lectura-ayer", type=float, default=LECTURA_AYER_M3)
    parser.add_argument("--lectura-hoy", type=float, default=LECTURA_HOY_M3)
    args = parser.parse_args()
    path = build_doc(args.lectura_ayer, args.lectura_hoy)
    print(f"[OK] Informe generado: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
