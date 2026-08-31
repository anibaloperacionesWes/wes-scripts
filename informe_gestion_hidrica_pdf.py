"""
Informes de gestión hídrica en formato Zapallar (one-pager + mensual), PDF.

Ajustes respecto de la muestra:
- Cuadro de hallazgos: texto de celdas en blanco sobre fondo de color.
- Columna PRIORIDAD (SEGUIMIENTO / INFORMATIVA): letra más chica para que no se parta.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from PIL import Image as PILImage
from reportlab.lib.colors import Color, HexColor, white
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

NAVY = HexColor("#003B64")
TEAL = HexColor("#087EAE")
GRAY = HexColor("#677681")
BODY = HexColor("#20313D")
LINE = HexColor("#C5D3DA")
KPI_BG = HexColor("#EAF4F8")
ROW_ALT = HexColor("#F7FAFB")
ORANGE = HexColor("#E67E22")
GREEN_BG = HexColor("#DCEFEB")
OBS_BG = HexColor("#087EAE")
ATT_BG = HexColor("#FFF1DE")
CRIT_BG = HexColor("#F5E3E3")
SCALE_L = HexColor("#EAF4F8")
SCALE_R = HexColor("#E2F0F6")
HALLAZGO_BG = HexColor("#003B64")
HALLAZGO_BG_ALT = HexColor("#0A4A70")
PRIO_SEGUIMIENTO = HexColor("#087EAE")
PRIO_INFORMATIVA = HexColor("#5D6D7E")
PRIO_ATENCION = HexColor("#C47A12")
PRIO_CRITICO = HexColor("#A93226")

PAGE_W, PAGE_H = LETTER
ML = 48.5
MR = 48.5
CONTENT_W = PAGE_W - ML - MR
TABLE_X = 53.7
TABLE_W = 504.56

# idle_bg, idle_fg, active_bg, active_fg, header_fg
STATUS_META = {
    "BAJO CONTROL": (GREEN_BG, GRAY, HexColor("#1E8449"), white, HexColor("#1E8449")),
    "EN OBSERVACIÓN": (HexColor("#D6EAF8"), GRAY, OBS_BG, white, TEAL),
    "REQUIERE ATENCIÓN": (ATT_BG, GRAY, HexColor("#D68910"), white, HexColor("#B9770E")),
    "CRÍTICO": (CRIT_BG, GRAY, HexColor("#C0392B"), white, HexColor("#922B21")),
}

PRIO_FILL = {
    "SEGUIMIENTO": PRIO_SEGUIMIENTO,
    "INFORMATIVA": PRIO_INFORMATIVA,
    "ATENCIÓN": PRIO_ATENCION,
    "CRÍTICO": PRIO_CRITICO,
}


@dataclass
class Hallazgo:
    prioridad: str
    titulo: str
    detalle: str
    lectura: str


@dataclass
class Accion:
    accion: str
    plazo: str
    objetivo: str
    responsable: str = ""


@dataclass
class PuntoIndicador:
    nombre: str
    total: float
    promedio: float
    max_m3: float
    max_fecha: str
    nocturno: float
    cobertura: int
    es_matriz: bool = False


@dataclass
class SerieDiaria:
    nombre: str
    fechas: List[datetime]
    valores: List[Optional[float]]
    lectura: str
    es_matriz: bool = False


@dataclass
class VisitaTecnicaSpec:
    fecha: str
    tecnico: str
    punto: str
    motivo: str
    diagnostico: str


@dataclass
class InformeSpec:
    cliente: str
    sitio: str
    periodo_corto: str
    footer: str
    titulo_onepager: str
    titulo_mensual: str
    clasificacion: str
    motivo: str
    kpi_entrada: str
    kpi_promedio: str
    kpi_nocturno: str
    kpi_pct: str
    panorama: List[Tuple[str, bool]]
    panorama_nota: str
    hallazgos: List[Hallazgo]
    acciones: List[Accion]
    conclusion: List[List[Tuple[str, bool]]]
    lectura_ejecutiva: List[List[Tuple[str, bool]]]
    nota_agosto: str
    kpi_consumo_label: str = "Consumo de entrada"
    chart_6m: Optional[Path] = None
    chart_puntos: Optional[Path] = None
    chart_puntos_nota: str = ""
    max_entrada_txt: str = ""
    chart_nocturno: Optional[Path] = None
    chart_nocturno_nota: str = ""
    indicadores: List[PuntoIndicador] = field(default_factory=list)
    criterio_nocturno: List[List[Tuple[str, bool]]] = field(default_factory=list)
    nota_cobertura: str = ""
    series_diarias: List[SerieDiaria] = field(default_factory=list)
    logo_path: Optional[Path] = None
    visitas: List[VisitaTecnicaSpec] = field(default_factory=list)


def _fmt(value: float, decimals: int = 1) -> str:
    formatted = f"{value:.{decimals}f}"
    integer_part, _, decimal_part = formatted.partition(".")
    neg = integer_part.startswith("-")
    if neg:
        integer_part = integer_part[1:]
    grouped = ""
    for i, digit in enumerate(reversed(integer_part)):
        if i and i % 3 == 0:
            grouped = "." + grouped
        grouped = digit + grouped
    if decimals <= 0:
        return ("-" if neg else "") + grouped
    decimal_part = decimal_part.rstrip("0")
    if not decimal_part:
        decimal_part = "0"
    return ("-" if neg else "") + grouped + "," + decimal_part


def _fmt_clp(value: float) -> str:
    return f"${_fmt(value, 0)}"


def _fecha_es(iso: str) -> str:
    dt = datetime.strptime(iso, "%Y-%m-%d")
    meses = {
        1: "enero",
        2: "febrero",
        3: "marzo",
        4: "abril",
        5: "mayo",
        6: "junio",
        7: "julio",
        8: "agosto",
        9: "septiembre",
        10: "octubre",
        11: "noviembre",
        12: "diciembre",
    }
    return f"{dt.day} de {meses[dt.month]}"


def resolve_logo() -> Optional[Path]:
    root = Path(__file__).resolve().parent
    for cand in (
        root / "logo wes.bmp",
        root / "logo_wes.bmp",
        root / "logo wes.png",
        root / "logo_wes.png",
    ):
        if cand.is_file():
            return cand
    return None


def _logo_reader(path: Optional[Path]) -> Optional[ImageReader]:
    if path is None or not path.is_file():
        return None
    img = PILImage.open(path).convert("RGBA")
    return ImageReader(img)


def _wrap(text: str, font: str, size: float, width: float) -> List[str]:
    text = (text or "").strip()
    if not text:
        return [""]
    words = text.split()
    lines: List[str] = []
    cur = ""
    for word in words:
        trial = (cur + " " + word).strip()
        if stringWidth(trial, font, size) <= width:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines or [""]


def _draw_header(c: canvas.Canvas, logo: Optional[ImageReader]) -> None:
    c.setFillColor(NAVY)
    c.rect(0, PAGE_H - 8, PAGE_W, 8, fill=1, stroke=0)
    c.setFillColor(GRAY)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(42.5, PAGE_H - 28.8, "WATER EFFICIENCY SERVICES")
    if logo is not None:
        c.drawImage(logo, 454.7, PAGE_H - 37.1, width=114.8, height=25.45, mask="auto")


def _draw_footer(c: canvas.Canvas, footer: str, page: int) -> None:
    c.setStrokeColor(HexColor("#D7E1E6"))
    c.setLineWidth(1)
    c.line(42.5, 32.6, 569.5, 32.6)
    c.setFillColor(GRAY)
    c.setFont("Helvetica", 7)
    c.drawString(42.5, 20.9, footer)
    c.drawRightString(569.5, 20.9, str(page))


def _draw_title_block(c: canvas.Canvas, title: str, subtitle: str, y: float) -> float:
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 24)
    c.drawString(ML, y, title)
    c.setFillColor(GRAY)
    c.setFont("Helvetica", 10.5)
    c.drawString(ML, y - 33.9, subtitle)
    return y - 52


def _draw_status_scale(c: canvas.Canvas, clasificacion: str, motivo: str, top: float) -> float:
    x0, y0 = TABLE_X, top - 79
    w, h_top, h_mid, h_bot = TABLE_W, 25, 32, 22
    c.setFillColor(SCALE_L)
    c.rect(x0, y0 + h_mid + h_bot, w / 2, h_top, fill=1, stroke=0)
    c.setFillColor(SCALE_R)
    c.rect(x0 + w / 2, y0 + h_mid + h_bot, w / 2, h_top, fill=1, stroke=0)
    labels = ["BAJO CONTROL", "EN OBSERVACIÓN", "REQUIERE ATENCIÓN", "CRÍTICO"]
    subtitles = {
        "BAJO CONTROL": "Consumo estable",
        "EN OBSERVACIÓN": "Requiere seguimiento",
        "REQUIERE ATENCIÓN": "Revisión y decisión",
        "CRÍTICO": "Acción inmediata",
    }
    col_w = w / 4
    for i, lab in enumerate(labels):
        idle_bg, idle_fg, active_bg, active_fg, _hdr = STATUS_META[lab]
        active = lab == clasificacion
        c.setFillColor(active_bg if active else idle_bg)
        c.rect(x0 + i * col_w, y0 + h_bot, col_w, h_mid, fill=1, stroke=0)
        c.setFillColor(active_fg if active else idle_fg)
        c.setFont("Helvetica-Bold", 7.5)
        c.drawCentredString(x0 + i * col_w + col_w / 2, y0 + h_bot + 17.2, lab)
        c.setFont("Helvetica", 6.5)
        c.drawCentredString(x0 + i * col_w + col_w / 2, y0 + h_bot + 7.1, subtitles[lab])
        if i:
            c.setStrokeColor(white)
            c.setLineWidth(0.7)
            c.line(x0 + i * col_w, y0 + h_bot, x0 + i * col_w, y0 + h_bot + h_mid)
    c.setFillColor(ROW_ALT)
    c.rect(x0, y0, w, h_bot, fill=1, stroke=0)
    c.setStrokeColor(HexColor("#BFCFE0"))
    c.setLineWidth(0.7)
    c.rect(x0, y0, w, h_top + h_mid + h_bot, fill=0, stroke=1)
    c.setFillColor(GRAY)
    c.setFont("Helvetica", 7.5)
    c.drawCentredString(x0 + w / 4, y0 + h_bot + h_mid + 7.4, "ESCALA DE ESTADO HÍDRICO")
    c.setFillColor(STATUS_META[clasificacion][4])
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(x0 + w / 2 + 8, y0 + h_bot + h_mid + 5.2, f"CLASIFICACIÓN ACTUAL: {clasificacion}")
    c.setFillColor(GRAY)
    c.setFont("Helvetica", 7.2)
    motivo_txt = f"Motivo del periodo: {motivo}"
    lines = _wrap(motivo_txt, "Helvetica", 7.2, w - 16)
    c.drawString(x0 + 8, y0 + 8.8 if len(lines) == 1 else y0 + 12.2, lines[0])
    if len(lines) > 1:
        c.drawString(x0 + 8, y0 + 3.2, lines[1])
    return y0 - 8


def _draw_kpis(c: canvas.Canvas, spec: InformeSpec, top: float) -> float:
    h = 48
    y = top - h
    col_w = TABLE_W / 4
    c.setFillColor(KPI_BG)
    c.roundRect(TABLE_X, y, TABLE_W, h, 0, fill=1, stroke=0)
    c.setStrokeColor(HexColor("#BFD7E3"))
    c.setLineWidth(0.5)
    c.rect(TABLE_X, y, TABLE_W, h, fill=0, stroke=1)
    items = [
        (spec.kpi_entrada, spec.kpi_consumo_label or "Consumo de entrada"),
        (spec.kpi_promedio, "Promedio diario"),
        (spec.kpi_nocturno, "Consumo nocturno"),
        (spec.kpi_pct, "Participación nocturna"),
    ]
    for i, (val, lab) in enumerate(items):
        cx = TABLE_X + i * col_w + col_w / 2
        if i:
            c.setStrokeColor(white)
            c.line(TABLE_X + i * col_w, y, TABLE_X + i * col_w, y + h)
        c.setFillColor(NAVY)
        c.setFont("Helvetica-Bold", 17)
        c.drawCentredString(cx, y + 24.5, val)
        c.setFillColor(GRAY)
        c.setFont("Helvetica", 7.5)
        c.drawCentredString(cx, y + 8.9, lab)
    return y - 12


def _section(c: canvas.Canvas, title: str, y: float, size: float = 15) -> float:
    c.setFillColor(NAVY if size >= 14 else TEAL)
    c.setFont("Helvetica-Bold", size)
    c.drawString(ML, y, title)
    return y - 16


def _draw_runs(
    c: canvas.Canvas,
    runs: Sequence[Tuple[str, bool]],
    x: float,
    y: float,
    width: float,
    size: float = 9,
    color: Color = BODY,
    leading: float = 13,
    align: str = "left",
) -> float:
    tokens: List[Tuple[str, str]] = []
    for text, bold in runs:
        font = "Helvetica-Bold" if bold else "Helvetica"
        for i, word in enumerate(text.split(" ")):
            piece = word if i == 0 else " " + word
            tokens.append((piece, font))
    lines: List[List[Tuple[str, str]]] = [[]]
    acc = 0.0
    for piece, font in tokens:
        w = stringWidth(piece, font, size)
        if acc + w > width and lines[-1]:
            lines.append([])
            acc = 0.0
            if piece.startswith(" "):
                piece = piece.lstrip(" ")
                w = stringWidth(piece, font, size)
        lines[-1].append((piece, font))
        acc += w
    for line in lines:
        total = sum(stringWidth(p, f, size) for p, f in line)
        if align == "center":
            cx = x + (width - total) / 2
        else:
            cx = x
        c.setFillColor(color)
        for piece, font in line:
            c.setFont(font, size)
            c.drawString(cx, y, piece)
            cx += stringWidth(piece, font, size)
        y -= leading
    return y


def _draw_numbered_hallazgos(c: canvas.Canvas, hallazgos: Sequence[Hallazgo], top: float) -> float:
    y = top
    x0 = TABLE_X
    num_w = 35.4
    row_gap = 0
    for i, h in enumerate(hallazgos):
        title_lines = _wrap(h.titulo, "Helvetica-Bold", 9, TABLE_W - num_w - 16)
        body_lines = _wrap(h.detalle, "Helvetica", 9, TABLE_W - num_w - 16)
        inner_h = 8 + 12 * len(title_lines) + 12 * len(body_lines) + 6
        inner_h = max(inner_h, 40)
        y0 = y - inner_h
        bg = HALLAZGO_BG if i % 2 == 0 else HALLAZGO_BG_ALT
        c.setFillColor(bg)
        c.rect(x0, y0, TABLE_W, inner_h, fill=1, stroke=0)
        c.setStrokeColor(HexColor("#C7D3DA"))
        c.setLineWidth(0.4)
        c.line(x0 + num_w, y0, x0 + num_w, y0 + inner_h)
        c.setFillColor(white)
        c.setFont("Helvetica-Bold", 17)
        c.drawCentredString(x0 + num_w / 2, y0 + inner_h / 2 - 5, str(i + 1))
        ty = y0 + inner_h - 16
        c.setFont("Helvetica-Bold", 9)
        for line in title_lines:
            c.drawString(x0 + num_w + 8, ty, line)
            ty -= 13
        c.setFont("Helvetica", 9)
        for line in body_lines:
            c.drawString(x0 + num_w + 8, ty, line)
            ty -= 13
        y = y0 - row_gap
    c.setStrokeColor(HexColor("#C7D3DA"))
    c.rect(x0, y, TABLE_W, top - y, fill=0, stroke=1)
    return y - 14


def _draw_table(
    c: canvas.Canvas,
    headers: Sequence[str],
    rows: Sequence[Sequence[str]],
    col_ws: Sequence[float],
    top: float,
    header_h: float = 19,
    font_size: float = 9,
    header_white: bool = True,
) -> float:
    x = TABLE_X
    y = top - header_h
    c.setFillColor(NAVY)
    c.rect(x, y, sum(col_ws), header_h, fill=1, stroke=0)
    c.setFillColor(white if header_white else GRAY)
    c.setFont("Helvetica", 7.5)
    acc = x
    for i, h in enumerate(headers):
        c.drawCentredString(acc + col_ws[i] / 2, y + 6.5, h)
        acc += col_ws[i]
    wrapped_rows = []
    for row in rows:
        wrapped = [_wrap(cell, "Helvetica", font_size, col_ws[i] - 14) for i, cell in enumerate(row)]
        wrapped_rows.append(wrapped)
    for r_i, wrapped in enumerate(wrapped_rows):
        nlines = max(len(w) for w in wrapped)
        rh = max(19, 8 + nlines * 13)
        y -= rh
        c.setFillColor(white if r_i % 2 == 0 else ROW_ALT)
        c.rect(x, y, sum(col_ws), rh, fill=1, stroke=0)
        c.setFillColor(BODY)
        acc = x
        for i, lines in enumerate(wrapped):
            ty = y + rh - 14
            for line in lines:
                c.setFont("Helvetica", font_size)
                c.drawString(acc + 7, ty, line)
                ty -= 13
            acc += col_ws[i]
    total_h = header_h + sum(max(19, 8 + max(len(w) for w in wr) * 13) for wr in wrapped_rows)
    c.setStrokeColor(HexColor("#C7D3DA"))
    c.setLineWidth(0.4)
    c.rect(x, y, sum(col_ws), total_h, fill=0, stroke=1)
    yy = top
    c.line(x, yy - header_h, x + sum(col_ws), yy - header_h)
    acc = x
    for w in col_ws[:-1]:
        acc += w
        c.line(acc, y, acc, top)
    cursor = top - header_h
    for wrapped in wrapped_rows:
        nlines = max(len(w) for w in wrapped)
        rh = max(19, 8 + nlines * 13)
        cursor -= rh
        c.line(x, cursor, x + sum(col_ws), cursor)
    return y - 12


def _draw_hallazgos_table(c: canvas.Canvas, hallazgos: Sequence[Hallazgo], top: float) -> float:
    """PRIORIDAD en una línea (letra chica) y texto de celdas en blanco."""
    col_prio, col_hall, col_lect = 78.0, 172.3, TABLE_W - 78.0 - 172.3
    header_h = 23
    x = TABLE_X
    y = top - header_h
    c.setFillColor(NAVY)
    c.rect(x, y, TABLE_W, header_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Helvetica", 7.5)
    c.drawCentredString(x + col_prio / 2, y + 8, "PRIORIDAD")
    c.drawCentredString(x + col_prio + col_hall / 2, y + 8, "HALLAZGO")
    c.drawCentredString(x + col_prio + col_hall + col_lect / 2, y + 8, "LECTURA Y DECISIÓN")

    row_specs = []
    for h in hallazgos:
        title_lines = _wrap(h.titulo, "Helvetica-Bold", 8.5, col_hall - 14)
        det_lines = _wrap(h.detalle, "Helvetica", 8.5, col_hall - 14)
        lect_lines = _wrap(h.lectura, "Helvetica", 8.5, col_lect - 14)
        n = max(len(title_lines) + len(det_lines), len(lect_lines), 2)
        rh = max(40, 10 + n * 12)
        row_specs.append((h, title_lines, det_lines, lect_lines, rh))

    for i, (h, title_lines, det_lines, lect_lines, rh) in enumerate(row_specs):
        y -= rh
        row_bg = HALLAZGO_BG if i % 2 == 0 else HALLAZGO_BG_ALT
        c.setFillColor(row_bg)
        c.rect(x, y, TABLE_W, rh, fill=1, stroke=0)
        prio_bg = PRIO_FILL.get(h.prioridad, PRIO_INFORMATIVA)
        c.setFillColor(prio_bg)
        c.rect(x, y, col_prio, rh, fill=1, stroke=0)
        c.setFillColor(white)
        # Letra chica para que SEGUIMIENTO / INFORMATIVA queden en una sola línea.
        c.setFont("Helvetica-Bold", 6.4)
        c.drawCentredString(x + col_prio / 2, y + rh / 2 - 2.2, h.prioridad)
        c.setFillColor(white)
        ty = y + rh - 14
        c.setFont("Helvetica-Bold", 8.5)
        for line in title_lines:
            c.drawString(x + col_prio + 7, ty, line)
            ty -= 12
        c.setFont("Helvetica", 8.5)
        for line in det_lines:
            c.drawString(x + col_prio + 7, ty, line)
            ty -= 12
        ty = y + rh - 14
        c.setFont("Helvetica", 8.5)
        for line in lect_lines:
            c.drawString(x + col_prio + col_hall + 7, ty, line)
            ty -= 12

    total_h = header_h + sum(r[-1] for r in row_specs)
    c.setStrokeColor(HexColor("#C7D3DA"))
    c.setLineWidth(0.4)
    c.rect(x, y, TABLE_W, total_h, fill=0, stroke=1)
    c.line(x + col_prio, y, x + col_prio, top)
    c.line(x + col_prio + col_hall, y, x + col_prio + col_hall, top)
    c.line(x, top - header_h, x + TABLE_W, top - header_h)
    cursor = top - header_h
    for *_, rh in row_specs:
        cursor -= rh
        c.line(x, cursor, x + TABLE_W, cursor)
    return y - 10


def _draw_image(c: canvas.Canvas, path: Optional[Path], top: float, height: float) -> float:
    if path is None or not path.is_file():
        return top
    y = top - height
    c.drawImage(str(path), TABLE_X, y, width=TABLE_W, height=height, preserveAspectRatio=True, mask="auto")
    return y - 6


def _style_axes(ax) -> None:
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#C5D3DA")
    ax.spines["bottom"].set_color("#C5D3DA")
    ax.tick_params(colors="#677681", labelsize=7.5)
    ax.yaxis.label.set_color("#20313D")
    ax.grid(axis="y", color="#E6EEF2", linewidth=0.6)
    ax.set_axisbelow(True)


def build_chart_6_meses(path: Path, labels: Sequence[str], values: Sequence[float]) -> Path:
    fig, ax = plt.subplots(figsize=(8.3, 2.55), dpi=160)
    colors = ["#E67E22" if i == len(values) - 1 else "#003B64" for i in range(len(values))]
    bars = ax.bar(labels, values, width=0.62, color=colors, edgecolor="none")
    ax.set_ylabel("m³", fontsize=8, fontweight="bold")
    ax.set_title("Consumo de entrada - últimos 6 meses", fontsize=10, fontweight="bold", color="#003B64", pad=8)
    _style_axes(ax)
    ymax = max(values) * 1.18 if values else 1
    ax.set_ylim(0, ymax)
    for bar, v in zip(bars, values):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height(),
            _fmt(v, 0),
            ha="center",
            va="bottom",
            fontsize=7.5,
            fontweight="bold",
            color="#20313D",
        )
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def build_chart_puntos(path: Path, names: Sequence[str], values: Sequence[float], matriz_name: str) -> Path:
    fig, ax = plt.subplots(figsize=(8.3, 2.7), dpi=160)
    colors = ["#E67E22" if n == matriz_name else "#5B9BD5" for n in names]
    bars = ax.bar(names, values, width=0.64, color=colors, edgecolor="none")
    ax.set_ylabel("m³", fontsize=8, fontweight="bold")
    ax.set_title("Consumo registrado por punto", fontsize=10, fontweight="bold", color="#003B64", pad=8)
    _style_axes(ax)
    rot = 35 if len(names) >= 8 else 18
    fs = 6.0 if len(names) >= 10 else 7.5
    plt.setp(ax.get_xticklabels(), rotation=rot, ha="right", fontsize=fs)
    ax.set_ylim(0, max(values) * 1.22 if values else 1)
    for bar, v in zip(bars, values):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height(),
            _fmt(v, 0) if v >= 100 else _fmt(v, 1),
            ha="center",
            va="bottom",
            fontsize=7,
            fontweight="bold",
            color="#20313D",
        )
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def build_chart_nocturno(
    path: Path,
    names: Sequence[str],
    values: Sequence[float],
    matriz_name: str = "",
    leyenda: Optional[Tuple[str, str]] = None,
) -> Path:
    fig, ax = plt.subplots(figsize=(8.3, 3.15), dpi=160)
    order = sorted(zip(names, values), key=lambda x: x[1], reverse=True)
    names_s = [n for n, _ in order]
    vals_s = [v for _, v in order]
    colors = ["#E67E22" if matriz_name and n == matriz_name else "#7FB3D5" for n in names_s]
    y = range(len(names_s))
    ax.barh(list(y), vals_s, color=colors, height=0.62, edgecolor="none")
    ax.invert_yaxis()
    ax.set_yticks(list(y), labels=names_s, fontsize=8)
    ax.set_xlabel("m³", fontsize=8, fontweight="bold")
    ax.set_title("Consumo nocturno por punto", fontsize=10, fontweight="bold", color="#003B64", pad=8)
    _style_axes(ax)
    ax.grid(axis="x", color="#E6EEF2", linewidth=0.6)
    ax.grid(axis="y", visible=False)
    xmax = max(vals_s) * 1.18 if vals_s else 1
    ax.set_xlim(0, xmax)
    for yi, v in zip(y, vals_s):
        ax.text(v + xmax * 0.012, yi, _fmt(v, 1), va="center", fontsize=7.5, color="#20313D", fontweight="bold")
    if matriz_name and leyenda:
        ax.legend(
            handles=[
                Patch(facecolor="#E67E22", label=leyenda[0]),
                Patch(facecolor="#7FB3D5", label=leyenda[1]),
            ],
            loc="lower right",
            fontsize=7,
            frameon=False,
        )
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def build_chart_diario(path: Path, serie: SerieDiaria) -> Path:
    fig, ax = plt.subplots(figsize=(8.4, 1.55), dpi=160)
    xs = serie.fechas
    ys = [v if v is not None else float("nan") for v in serie.valores]
    ax.plot(xs, ys, color="#003B64", linewidth=1.6, marker="o", markersize=3.2, markerfacecolor="#003B64")
    ax.fill_between(xs, ys, color="#D6EAF8", alpha=0.85)
    nums = [(x, v) for x, v in zip(xs, serie.valores) if v is not None]
    if nums:
        xmax, vmax = max(nums, key=lambda t: t[1])
        ax.scatter([xmax], [vmax], color="#E67E22", s=28, zorder=5)
        ax.annotate(
            f"{_fmt(vmax, 1)} m³",
            xy=(xmax, vmax),
            xytext=(8, 6),
            textcoords="offset points",
            fontsize=7,
            color="#E67E22",
            fontweight="bold",
        )
    ax.set_ylabel("m³", fontsize=7.5, fontweight="bold")
    _style_axes(ax)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d/%m"))
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=3))
    plt.setp(ax.get_xticklabels(), fontsize=6.5)
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def _new_page(c: canvas.Canvas, spec: InformeSpec, logo: Optional[ImageReader], page: int) -> None:
    _draw_header(c, logo)
    _draw_footer(c, spec.footer, page)


def render_one_pager(spec: InformeSpec, out_path: Path) -> Path:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(out_path), pagesize=LETTER)
    logo = _logo_reader(spec.logo_path or resolve_logo())
    _new_page(c, spec, logo, 1)
    y = _draw_title_block(c, spec.titulo_onepager, spec.periodo_corto, PAGE_H - 61)
    y = _draw_status_scale(c, spec.clasificacion, spec.motivo, y)
    y = _draw_kpis(c, spec, y)
    y = _section(c, "Panorama del periodo", y - 4, 15)
    y = _draw_runs(c, spec.panorama, ML, y, CONTENT_W, 9, BODY, 13) - 2
    y = _draw_runs(c, [(spec.panorama_nota, False)], ML, y, CONTENT_W, 9, BODY, 13)
    y = _section(c, "Hallazgos que requieren decisión", y - 6, 10.5)
    y = _draw_numbered_hallazgos(c, spec.hallazgos, y + 4)
    y = _section(c, "Decisiones recomendadas", y + 2, 10.5)
    rows = [[a.accion, a.plazo, a.objetivo] for a in spec.acciones]
    y = _draw_table(c, ["ACCIÓN", "PLAZO", "OBJETIVO"], rows, [198.4, 87.9, 218.26], y + 4)
    y = _section(c, "Conclusión", y + 2, 10.5)
    for para in spec.conclusion:
        y = _draw_runs(c, para, ML, y, CONTENT_W, 9, BODY, 13) - 2
    c.save()
    return out_path


def _draw_visitas_section(c: canvas.Canvas, spec: InformeSpec, y: float) -> float:
    """Solo se llama si hay visitas del formulario en el periodo."""
    y = _section(c, "Visitas técnicas del periodo", y - 6, 10.5)
    y = _draw_runs(
        c,
        [
            (
                "Registro del formulario de técnicos WES en este recinto. "
                "Se copian fecha, punto, tipo y diagnóstico de cada visita.",
                False,
            )
        ],
        ML,
        y + 2,
        CONTENT_W,
        8,
        GRAY,
        11,
    )
    rows = [
        [v.fecha, v.tecnico, v.punto, v.motivo, v.diagnostico] for v in spec.visitas
    ]
    col_ws = [58.0, 92.0, 78.0, 118.0, TABLE_W - 58.0 - 92.0 - 78.0 - 118.0]
    return _draw_table(
        c,
        ["FECHA", "TÉCNICO", "PUNTO", "MOTIVO / TIPO", "DIAGNÓSTICO"],
        rows,
        col_ws,
        y + 4,
        font_size=7.5,
    )


def render_mensual(spec: InformeSpec, out_path: Path, chart_dir: Path) -> Path:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    chart_dir.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(out_path), pagesize=LETTER)
    logo = _logo_reader(spec.logo_path or resolve_logo())

    # Página 1
    _new_page(c, spec, logo, 1)
    y = _draw_title_block(c, spec.titulo_mensual, spec.periodo_corto, PAGE_H - 61)
    y = _draw_status_scale(c, spec.clasificacion, spec.motivo, y)
    y = _draw_kpis(c, spec, y)
    y = _section(c, "Lectura ejecutiva", y - 2, 15)
    for para in spec.lectura_ejecutiva:
        y = _draw_runs(c, para, ML, y, CONTENT_W, 9, BODY, 13) - 4
    y = _draw_image(c, spec.chart_6m, y - 2, 168)
    c.setFillColor(GRAY)
    c.setFont("Helvetica", 7.2)
    for line in _wrap(spec.nota_agosto, "Helvetica", 7.2, CONTENT_W):
        c.drawString(ML, y, line)
        y -= 10
    c.showPage()

    # Página 2
    _new_page(c, spec, logo, 2)
    y = _section(c, "Hallazgos prioritarios", PAGE_H - 58, 15)
    y = _draw_hallazgos_table(c, spec.hallazgos, y + 6)
    y = _draw_image(c, spec.chart_puntos, y, 175)
    y = _draw_runs(
        c,
        [("Cómo leer este gráfico: ", True), (spec.chart_puntos_nota, False)],
        ML,
        y,
        CONTENT_W,
        7.2,
        GRAY,
        10,
    )
    y = _draw_runs(c, [(spec.max_entrada_txt, False)], ML, y - 4, CONTENT_W, 9, BODY, 13)
    c.showPage()

    # Página 3
    _new_page(c, spec, logo, 3)
    y = _section(c, "Consumo nocturno y plan de acción", PAGE_H - 58, 15)
    chart_h = 158 if spec.visitas else 210
    y = _draw_image(c, spec.chart_nocturno, y + 6, chart_h)
    y = _draw_runs(c, [(spec.chart_nocturno_nota, False)], ML, y, CONTENT_W, 7.2, GRAY, 10)
    y = _section(c, "Acciones recomendadas", y - 8, 10.5)
    rows = [[a.accion, a.responsable, a.plazo, a.objetivo] for a in spec.acciones]
    y = _draw_table(
        c,
        ["ACCIÓN", "RESPONSABLE SUGERIDO", "PLAZO", "RESULTADO ESPERADO"],
        rows,
        [150.0, 110.0, 72.0, TABLE_W - 332.0],
        y + 4,
        font_size=8.5,
    )
    if spec.visitas:
        if y < 210:
            c.showPage()
            _new_page(c, spec, logo, 3)
            y = PAGE_H - 58
        y = _draw_visitas_section(c, spec, y)
    if y < 88:
        c.showPage()
        _new_page(c, spec, logo, 3)
        y = PAGE_H - 58
    y = _section(c, "Conclusión", y + 2, 10.5)
    for para in spec.conclusion:
        y = _draw_runs(c, para, ML, y, CONTENT_W, 9, BODY, 13) - 4
    c.showPage()

    # Página 4
    _new_page(c, spec, logo, 4)
    y = _section(c, "Anexo técnico - indicadores por punto", PAGE_H - 58, 15)
    headers = ["PUNTO", "TOTAL", "PROMEDIO", "MÁXIMO DIARIO", "NOCTURNO", "COBERTURA NOCHE"]
    col_ws = [110, 72, 78, 100, 72, TABLE_W - 110 - 72 - 78 - 100 - 72]
    rows = []
    for p in spec.indicadores:
        rows.append(
            [
                p.nombre,
                f"{_fmt(p.total, 1)} m³",
                f"{_fmt(p.promedio, 1)} m³/día",
                f"{_fmt(p.max_m3, 1)} m³ · {p.max_fecha}",
                f"{_fmt(p.nocturno, 1)} m³",
                f"{p.cobertura} días",
            ]
        )
    y = _draw_table(c, headers, rows, col_ws, y + 6, font_size=7 if len(rows) >= 10 else 8)
    y = _section(c, "Criterio nocturno y costo referencial", y - 2, 10.5)
    for para in spec.criterio_nocturno:
        y = _draw_runs(c, para, ML, y, CONTENT_W, 9, BODY, 13) - 4
    if spec.nota_cobertura:
        y = _draw_runs(c, [(spec.nota_cobertura, False)], ML, y - 2, CONTENT_W, 7.2, GRAY, 10)
    c.showPage()

    # Página 5
    _new_page(c, spec, logo, 5)
    y = _section(c, "Anexo técnico - gráficos relevantes", PAGE_H - 58, 15)
    y = _draw_runs(
        c,
        [
            (
                "Se presentan únicamente las series que aportan información para interpretar el periodo y respaldar decisiones.",
                False,
            )
        ],
        ML,
        y + 4,
        CONTENT_W,
        8,
        GRAY,
        11,
    )
    for i, serie in enumerate(spec.series_diarias[:3]):
        chart = chart_dir / f"diario_{i + 1}.png"
        build_chart_diario(chart, serie)
        y = _section(c, serie.nombre, y - 6, 10.5)
        y = _draw_image(c, chart, y + 8, 112)
        y = _draw_runs(
            c,
            [("Lectura: ", True), (serie.lectura, False)],
            ML,
            y,
            CONTENT_W,
            8,
            GRAY,
            11,
        )
        y -= 2
    c.save()
    return out_path


# re-export helpers used by the Inchcape runner
__all__ = [
    "Hallazgo",
    "Accion",
    "PuntoIndicador",
    "SerieDiaria",
    "VisitaTecnicaSpec",
    "InformeSpec",
    "_fmt",
    "_fmt_clp",
    "_fecha_es",
    "resolve_logo",
    "build_chart_6_meses",
    "build_chart_puntos",
    "build_chart_nocturno",
    "render_one_pager",
    "render_mensual",
]
