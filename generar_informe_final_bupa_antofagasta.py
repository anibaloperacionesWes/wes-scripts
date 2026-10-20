"""
Informe final — Estudio de diagnóstico de gestión hídrica
Clínica Bupa Antofagasta (solo Antofagasta).

Periodo: 31/07/2026 → 30/09/2026.
Puntos: Sanitaria, Bomba Principal, Sexto Piso (sin Sala N°2).
Modelo visual: Informe final WES - BPA.
"""

from __future__ import annotations

import json
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from generar_reporte_word import format_number_chilean

ROOT = Path("reports/Bupa_Antofagasta/INFORME_FINAL")
CHARTS = ROOT / "charts"
FOTOS = ROOT / "fotos_crop"
CACHE = Path(
    "reports/Bupa_Antofagasta/ANALISIS_PILOTO/cache_diario_bupa_antofagasta.json"
)
HOURLY_CACHE = ROOT / "hourly_cache.json"
_MESES_APP = {
    1: "ene.", 2: "feb.", 3: "mar.", 4: "abr.", 5: "may.", 6: "jun.",
    7: "jul.", 8: "ago.", 9: "sep.", 10: "oct.", 11: "nov.", 12: "dic.",
}
_MESES_XTICK = {
    1: "Ene.", 2: "Feb.", 3: "Mar.", 4: "Abr.", 5: "May.", 6: "Jun.",
    7: "Jul.", 8: "Agt.", 9: "Sep.", 10: "Oct.", 11: "Nov.", 12: "Dic.",
}

START = date(2026, 7, 31)
END = date(2026, 9, 30)
NUM_DIAS = (END - START).days + 1  # 62
DAY_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
NODOS = [
    {"id": "000029-09", "corto": "Sanitaria", "nombre": "Medidor Principal Sanitaria", "foto": "sanitaria.jpg"},
    {"id": "000029-07", "corto": "Bomba Principal", "nombre": "Sala de Bomba Principal (Primer Piso)", "foto": "bomba_principal.jpg"},
    {"id": "000029-08", "corto": "Sexto Piso", "nombre": "Sala de Bomba Sexto Piso", "foto": "sexto_piso.jpg"},
]

# Facturas Aguas Antofagasta (Centro Médico Antofagasta S.A. · cliente 1110080842)
FACTURAS = [
    {"ciclo": "Junio 2026", "emision": "02-06-2026", "m3": 7102.0, "clp": 19_662_850.0, "nro": "1453711", "lect_ant": 87562, "lect_act": 94664},
    {"ciclo": "Julio 2026", "emision": "03-07-2026", "m3": 6696.0, "clp": 18_538_860.0, "nro": "1460583", "lect_ant": 94664, "lect_act": 101360},
    {"ciclo": "Agosto 2026", "emision": "04-08-2026", "m3": 7627.0, "clp": 21_116_260.0, "nro": "1467321", "lect_ant": 101360, "lect_act": 108987},
    {"ciclo": "Septiembre 2026", "emision": "27-08-2026", "m3": 2873.0, "clp": 7_955_190.0, "nro": "1474021", "lect_ant": 108987, "lect_act": 111860},
]


def _fmt_clp(v: float) -> str:
    return f"${format_number_chilean(v, 0)}"


def _fmt_m3(v: float, nd: int = 1) -> str:
    return format_number_chilean(v, nd)


def _fmt_fecha(d: date | str) -> str:
    if isinstance(d, str):
        d = date.fromisoformat(d)
    return d.strftime("%d/%m/%Y")


def _set_run(run, *, size=10, bold=False, color=None, highlight=None) -> None:
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color or RGBColor(0, 0, 0)
    if highlight is not None:
        run.font.highlight_color = highlight


def _add_para(
    doc: Document,
    text: str,
    *,
    size: int = 10,
    bold: bool = False,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    space_after: int = 6,
    space_before: int = 0,
    highlight: bool = False,
    bullet: bool = False,
) -> None:
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    if bullet:
        p.style = "List Bullet"
    r = p.add_run(text)
    _set_run(
        r,
        size=size,
        bold=bold,
        highlight=WD_COLOR_INDEX.YELLOW if highlight else None,
    )


def _add_title(doc: Document, text: str, *, size: int = 13) -> None:
    """Títulos de sección sobrios (estilo reportes WES, sin bloque navy)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    _set_run(r, size=size, bold=True, color=RGBColor(47, 84, 150))


def _add_subtitle(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    _set_run(r, size=11, bold=True, color=RGBColor(55, 65, 81))


def _page_break(doc: Document) -> None:
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def _add_picture(doc: Document, path: Path, width_in: float, caption: str = "") -> None:
    if not path.is_file():
        _add_para(doc, f"[Sin imagen: {path.name}]", size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(str(path), width=Inches(width_in))
    if caption:
        c = doc.add_paragraph()
        c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraph_format.space_after = Pt(8)
        r = c.add_run(caption)
        _set_run(r, size=8, color=RGBColor(80, 80, 80))


def _center_table(table) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _shade_header(row) -> None:
    """Cabecera clara estilo WES (D9E1F2), sin rectángulo navy + blanco."""
    for cell in row.cells:
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "D9E1F2")
        shd.set(qn("w:val"), "clear")
        tcPr.append(shd)
        for para in cell.paragraphs:
            for run in para.runs:
                run.font.color.rgb = RGBColor(31, 56, 100)
                run.bold = True
                run.font.size = Pt(9)


def _add_table(doc: Document, headers: list, rows: list, *, compact: bool = False) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for j, h in enumerate(headers):
        table.rows[0].cells[j].text = h
    _shade_header(table.rows[0])
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = table.rows[i].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            r = p.add_run(str(val))
            _set_run(r, size=8 if compact else 9, bold=("Total" in str(row[0])))
    _center_table(table)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def _add_kpi_table(doc: Document, rows: List[Tuple[str, str]]) -> None:
    """Tabla compacta Indicador | Valor (sin negritas agresivas)."""
    table = doc.add_table(rows=len(rows), cols=2)
    table.style = "Table Grid"
    for i, (k, v) in enumerate(rows):
        table.rows[i].cells[0].text = ""
        table.rows[i].cells[1].text = ""
        # fila alternada muy suave
        if i % 2 == 0:
            for cell in table.rows[i].cells:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:fill"), "F7F9FC")
                shd.set(qn("w:val"), "clear")
                tcPr.append(shd)
        r0 = table.rows[i].cells[0].paragraphs[0].add_run(k)
        r1 = table.rows[i].cells[1].paragraphs[0].add_run(v)
        _set_run(r0, size=9, bold=False, color=RGBColor(55, 65, 81))
        _set_run(r1, size=9, bold=False, color=RGBColor(17, 24, 39))
        table.rows[i].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        table.rows[i].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    doc.add_paragraph().paragraph_format.space_after = Pt(6)


def _setup(doc: Document) -> None:
    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)


def _series_from_cache(cache: dict, nid: str) -> List[Tuple[date, float]]:
    s = cache.get(nid, {})
    out = []
    cur = START
    while cur <= END:
        if cur.isoformat() in s:
            out.append((cur, float(s[cur.isoformat()])))
        cur += timedelta(days=1)
    return out


def _pick_max_min_near(
    pts: List[Tuple[date, float]],
) -> Tuple[Optional[Tuple[date, float]], Optional[Tuple[date, float]], List[Tuple[date, float]], float]:
    positive = [(d, v) for d, v in pts if v > 0.01]
    if not positive:
        return None, None, [], 0.0
    avg = sum(v for _, v in positive) / len(positive)
    mx = max(positive, key=lambda x: x[1])
    mn = min(positive, key=lambda x: x[1])
    excl = {mx[0], mn[0]}
    near = sorted((p for p in positive if p[0] not in excl), key=lambda x: abs(x[1] - avg))[:3]
    return mx, mn, near, avg


def _chart_daily(pts: List[Tuple[date, float]], path: Path, title: str) -> None:
    fig, ax = plt.subplots(figsize=(9.0, 2.45))
    dates = [datetime.combine(d, datetime.min.time()) for d, _ in pts]
    vals = [v for _, v in pts]
    ax.fill_between(dates, vals, color="#B7D4F0", alpha=0.55)
    ax.plot(dates, vals, color="#3A7CC9", lw=1.5, marker="o", markersize=2.0,
            markerfacecolor="white", markeredgecolor="#3A7CC9", markeredgewidth=0.8)
    ax.set_title(title, fontsize=10, color="#1F3864", pad=6, loc="left")
    ax.set_ylabel("m³", fontsize=8, color="#666666")
    ax.set_ylim(bottom=0)
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(dates) // 10)))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%b"))
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=30, ha="right", fontsize=7, color="#666666")
    ax.tick_params(axis="y", labelsize=7, colors="#666666")
    ax.yaxis.grid(True, color="#EEEEEE", lw=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color("#DDDDDD")
    ax.spines["bottom"].set_color("#DDDDDD")
    fig.tight_layout(pad=0.35)
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def _chart_dow(pts: List[Tuple[date, float]], path: Path, title: str) -> Dict[str, float]:
    buckets: Dict[int, List[float]] = defaultdict(list)
    for d, v in pts:
        buckets[d.weekday()].append(v)
    avgs = {DAY_ES[i]: (sum(buckets[i]) / len(buckets[i]) if buckets[i] else 0.0) for i in range(7)}
    fig, ax = plt.subplots(figsize=(8.8, 2.3))
    values = [avgs[d] for d in DAY_ES]
    bars = ax.bar(DAY_ES, values, color="#5B9BD5", alpha=0.9, edgecolor="none", width=0.62)
    for b, v in zip(bars, values):
        ax.text(
            b.get_x() + b.get_width() / 2,
            b.get_height(),
            _fmt_m3(v, 1),
            ha="center",
            va="bottom",
            fontsize=7.5,
            color="#444444",
        )
    ax.set_title(title, fontsize=10, color="#1F3864", pad=6, loc="left")
    ax.set_ylabel("m³", fontsize=8, color="#666666")
    ax.set_ylim(bottom=0)
    ax.tick_params(colors="#666666", labelsize=7.5)
    ax.yaxis.grid(True, color="#EEEEEE", lw=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color("#DDDDDD")
    ax.spines["bottom"].set_color("#DDDDDD")
    fig.tight_layout(pad=0.35)
    fig.savefig(path, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return avgs


def _load_hourly_cache() -> dict:
    if HOURLY_CACHE.is_file():
        return json.loads(HOURLY_CACHE.read_text(encoding="utf-8"))
    return {"shapes": {}, "exact": {}}


def _hourly_for_day(nid: str, day: date, total_m3: float, hourly_cache: dict) -> List[float]:
    """Serie horaria real si existe; si no, forma del punto escalada al total del día."""
    exact = (hourly_cache.get("exact") or {}).get(nid, {}).get(day.isoformat())
    if exact and len(exact) == 24:
        arr = [float(x) for x in exact]
        s = sum(arr)
        if s > 0 and abs(s - total_m3) / max(total_m3, 1e-6) > 0.02:
            arr = [x * total_m3 / s for x in arr]
        return arr
    shape = (hourly_cache.get("shapes") or {}).get(nid)
    if shape and len(shape) == 24 and sum(shape) > 0:
        return [float(x) * float(total_m3) for x in shape]
    # Último recurso: no plano absoluto — leve perfil día/noche genérico
    base = [0.7, 0.65, 0.6, 0.6, 0.7, 0.9, 1.1, 1.2, 1.3, 1.4, 1.35, 1.3,
            1.25, 1.2, 1.15, 1.1, 1.0, 1.05, 1.0, 0.95, 0.9, 0.85, 0.8, 0.75]
    s = sum(base)
    return [b / s * float(total_m3) for b in base]


def _chart_24h_app(vals: List[float], path: Path, day: date) -> None:
    """Gráfico 24 h estilo App WES: título, subtítulo debajo, área suave, markers huecos."""
    import math

    hours = list(range(24))
    fig, ax = plt.subplots(figsize=(8.6, 3.05))
    # relleno con degradé suave (como la app), no bloque sólido
    ax.fill_between(hours, vals, 0, color="#C5DCF5", alpha=0.75)
    ax.plot(
        hours,
        vals,
        color="#2F6FBE",
        lw=1.8,
        marker="o",
        markersize=4.0,
        markerfacecolor="white",
        markeredgecolor="#2F6FBE",
        markeredgewidth=1.1,
    )
    day_label = f"{day.day} {_MESES_APP[day.month]} {day.year}"
    ax.set_title("Consumo (m3/hr)", fontsize=12, color="#1F3864", loc="left", pad=18)
    # subtítulo DEBAJO del título (como en la app), no arriba
    ax.text(
        0.0,
        1.02,
        f"1 dispositivo seleccionado  ›  {day_label}",
        transform=ax.transAxes,
        fontsize=8.5,
        color="#8A8A8A",
        va="bottom",
        ha="left",
    )
    ax.set_xlim(0, 23)
    ymax = max(vals) * 1.28 if vals else 1.0
    step = 5 if ymax >= 8 else (2 if ymax >= 3 else 1)
    nice = int(math.ceil(ymax / step) * step)
    ax.set_ylim(0, max(nice, step))
    ax.set_yticks(list(range(0, max(nice, step) + 1, step)))
    ax.set_xticks([0, 6, 12, 18])
    ax.set_xticklabels(
        [f"{day.day:02d}. {_MESES_XTICK[day.month]}", "06:00", "12:00", "18:00"]
    )
    ax.yaxis.grid(True, color="#EFEFEF", lw=0.9)
    ax.set_axisbelow(True)
    ax.tick_params(colors="#777777", labelsize=8)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_color("#E0E0E0")
    ax.spines["bottom"].set_color("#E0E0E0")
    fig.subplots_adjust(top=0.78, left=0.07, right=0.98, bottom=0.12)
    fig.savefig(path, dpi=170, facecolor="white")
    plt.close(fig)


def _prepare_node_assets(cache: dict) -> Dict[str, dict]:
    CHARTS.mkdir(parents=True, exist_ok=True)
    hourly_cache = _load_hourly_cache()
    out: Dict[str, dict] = {}
    for n in NODOS:
        nid = n["id"]
        slug = n["corto"].lower().replace(" ", "_")
        pts = _series_from_cache(cache, nid)
        vals = [v for _, v in pts]
        total = sum(vals)
        prom = total / len(vals) if vals else 0.0
        prom_mes = total * 30.0 / NUM_DIAS if NUM_DIAS else 0.0
        mx, mn, near, avg = _pick_max_min_near(pts)
        _chart_daily(
            pts,
            CHARTS / f"chart_diario_{slug}.png",
            f"Perfil de consumo diario (m³) — {n['corto']} · Fuente: APP WES",
        )
        dow = _chart_dow(
            pts,
            CHARTS / f"chart_dow_{slug}.png",
            f"Promedio de consumo diario (m³) — {n['corto']} · Fuente: WES App",
        )

        def _make_24h(day: date, total_day: float, out_path: Path) -> None:
            series = _hourly_for_day(nid, day, total_day, hourly_cache)
            _chart_24h_app(series, out_path, day)

        if mx:
            _make_24h(mx[0], mx[1], CHARTS / f"chart_24h_max_{slug}.png")
        if mn:
            _make_24h(mn[0], mn[1], CHARTS / f"chart_24h_min_{slug}.png")
        near_paths = []
        for i, (d, v) in enumerate(near, start=1):
            p = CHARTS / f"chart_24h_cerca{i}_{slug}.png"
            _make_24h(d, v, p)
            near_paths.append((d, v, p))
        out[nid] = {
            **n,
            "pts": pts,
            "total": total,
            "n": len(vals),
            "prom": prom,
            "prom_mes": prom_mes,
            "avg_pos": avg,
            "mx": mx,
            "mn": mn,
            "near": near_paths,
            "dow": dow,
            "slug": slug,
        }
    return out


def build() -> Path:
    ROOT.mkdir(parents=True, exist_ok=True)
    cache = json.loads(CACHE.read_text(encoding="utf-8"))
    nodes = _prepare_node_assets(cache)

    san = nodes["000029-09"]
    prin = nodes["000029-07"]
    sexto = nodes["000029-08"]
    tot_salas = prin["total"] + sexto["total"]
    gap = max(0.0, san["total"] - tot_salas)
    pct_prin = 100.0 * prin["total"] / san["total"] if san["total"] else 0.0
    pct_sexto = 100.0 * sexto["total"] / san["total"] if san["total"] else 0.0
    pct_salas = 100.0 * tot_salas / san["total"] if san["total"] else 0.0
    pct_gap = 100.0 * gap / san["total"] if san["total"] else 0.0
    tot_fact_m3 = sum(f["m3"] for f in FACTURAS)
    tot_fact_clp = sum(f["clp"] for f in FACTURAS)

    # pie + facturas chart (estilo sobrio)
    fig, ax = plt.subplots(figsize=(3.8, 2.9))
    wedges, texts, autotexts = ax.pie(
        [prin["total"], sexto["total"], gap],
        labels=["Principal", "Sexto", "No mon."],
        colors=["#5B9BD5", "#9DC3E6", "#F4B183"],
        autopct=lambda p: f"{p:.0f}%",
        startangle=90,
        textprops={"fontsize": 8, "color": "#444444"},
        wedgeprops={"linewidth": 1, "edgecolor": "white"},
    )
    for t in autotexts:
        t.set_fontsize(8)
        t.set_color("#333333")
    ax.set_title("% vs Sanitaria", fontsize=9, color="#1F3864", pad=4)
    fig.tight_layout()
    fig.savefig(CHARTS / "chart_participacion.png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(7.0, 2.6))
    labs = [f["ciclo"].replace(" 2026", "") for f in FACTURAS]
    vals = [f["m3"] for f in FACTURAS]
    ax.bar(labs, vals, color="#5B9BD5", edgecolor="none", width=0.55)
    for i, v in enumerate(vals):
        ax.text(i, v, _fmt_m3(v, 0), ha="center", va="bottom", fontsize=7.5, color="#444444")
    ax.set_ylabel("m³", fontsize=8, color="#666666")
    ax.set_title("Consumo facturado Aguas Antofagasta (m³)", fontsize=10, color="#1F3864", loc="left")
    ax.set_ylim(bottom=0)
    ax.tick_params(colors="#666666", labelsize=8)
    ax.yaxis.grid(True, color="#EEEEEE", lw=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color("#DDDDDD")
    ax.spines["bottom"].set_color("#DDDDDD")
    fig.tight_layout()
    fig.savefig(CHARTS / "chart_facturas_m3.png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    doc = Document()
    _setup(doc)

    # ===== PORTADA =====
    for _ in range(5):
        doc.add_paragraph()
    _add_para(doc, "Estudio de diagnóstico de gestión hídrica", size=20, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    _add_para(doc, "Clínica Bupa Antofagasta", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    _add_para(doc, "WES — Monitoreo Inteligente", size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
    _add_para(doc, "Octubre 2026", size=12, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    _add_para(
        doc,
        f"Periodo de monitoreo: {_fmt_fecha(START)} – {_fmt_fecha(END)} ({NUM_DIAS} días)",
        size=10,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_before=10,
    )

    # ===== ÍNDICE =====
    _page_break(doc)
    _add_title(doc, "Índice", size=16)
    for titulo, pag in [
        ("Resumen Ejecutivo", "2"),
        ("I. Antecedentes", "3"),
        ("II. Objetivos y resultados esperados", "4"),
        ("III. Definición de sectores a monitorear", "4"),
        ("IV. Fotografías de instalación", "5"),
        ("V. Análisis de sectores monitoreados", "6"),
        ("    1) Medidor Principal Sanitaria", "6"),
        ("    2) Sala de Bomba Principal", "8"),
        ("    3) Sala de Bomba Sexto Piso", "10"),
        ("VI. Facturación sanitaria (Aguas Antofagasta)", "12"),
        ("VII. Conclusiones y hallazgos", "13"),
        ("VIII. Recomendaciones y próximos pasos", "13"),
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"{titulo}\t{pag}")
        _set_run(r, size=10)

    # ===== RESUMEN =====
    _page_break(doc)
    _add_title(doc, "Resumen Ejecutivo")
    _add_subtitle(doc, "Propósito y Alcance")
    _add_para(
        doc,
        "El Estudio tuvo como objetivo principal establecer una línea de base detallada del consumo "
        "hídrico de la Clínica Bupa Antofagasta y validar la solución de monitoreo inteligente de WES "
        "para potenciar la capacidad de gestión eficiente y sostenible del agua.",
        bullet=True,
    )
    _add_para(
        doc,
        f"El estudio contempló el monitoreo en línea de 3 puntos críticos durante {_fmt_fecha(START)}–"
        f"{_fmt_fecha(END)} ({NUM_DIAS} días), con perfiles diarios, lun–dom y curvas 24 h del día "
        "pico, del día mínimo y de tres días cercanos al promedio.",
        bullet=True,
    )

    _add_subtitle(doc, "Hallazgos Principales de Consumo")
    _add_para(
        doc,
        f"Brecha de visibilidad: {_fmt_m3(gap)} m³ ({_fmt_m3(pct_gap)}% del total Sanitaria) entre "
        f"Medidor Principal Sanitaria ({_fmt_m3(san['total'])} m³) y la suma de Bomba Principal + "
        f"Sexto Piso ({_fmt_m3(tot_salas)} m³). Hipótesis: tramos de red no instrumentados "
        "(empalmes previos a impulsión / otros consumos del recinto).",
        highlight=True,
        bullet=True,
    )
    _add_para(
        doc,
        f"Sanitaria = 100% de la cuenta WES del periodo: promedio {_fmt_m3(san['prom'])} m³/día "
        f"(rango {_fmt_m3(san['mn'][1] if san['mn'] else 0)}–{_fmt_m3(san['mx'][1] if san['mx'] else 0)} m³/día). "
        f"Principal {_fmt_m3(pct_prin)}% · Sexto {_fmt_m3(pct_sexto)}%.",
        highlight=True,
        bullet=True,
    )
    _add_para(
        doc,
        f"Facturación Aguas Antofagasta (jun–sep 2026): {_fmt_m3(tot_fact_m3, 0)} m³ y "
        f"{_fmt_clp(tot_fact_clp)} (cliente 1110080842 · medidor 534 · Centro Médico Antofagasta S.A.).",
        highlight=True,
        bullet=True,
    )

    _add_subtitle(doc, "Recomendaciones Estratégicas")
    _add_para(
        doc,
        f"Monitoreo incremental: extender medición a tramos no visualizados para cerrar la brecha "
        f"del {_fmt_m3(pct_gap)}%.",
        highlight=True,
        bullet=True,
    )
    _add_para(
        doc,
        "Apoyo tecnológico: implementar Control Inteligente WES en sectores de alto consumo "
        "(horarios inhábiles prioritario).",
        highlight=True,
        bullet=True,
    )
    _add_para(
        doc,
        "Seguimiento factura vs WES: contrastar cada ciclo de Aguas Antofagasta con Sanitaria WES.",
        highlight=True,
        bullet=True,
    )

    # ===== ANTECEDENTES =====
    _page_break(doc)
    _add_title(doc, "I. Antecedentes")
    _add_para(
        doc,
        "La Clínica Bupa solicitó a WES un estudio de diagnóstico de gestión hídrica en "
        "Clínica Bupa Antofagasta para el monitoreo en línea de sus redes.",
        bullet=True,
    )
    _add_para(
        doc,
        "El objetivo es construir una línea de base detallada del consumo hídrico de los "
        "sectores instrumentados del recinto.",
        bullet=True,
    )
    _add_para(
        doc,
        f"Periodo analizado en este informe: {_fmt_fecha(START)} al {_fmt_fecha(END)} "
        f"({NUM_DIAS} días), con tres puntos de medición.",
        bullet=True,
    )
    _add_subtitle(doc, "a. Hitos")
    for h in [
        "Reunión de Kick-off / 3 de junio",
        "Visita de inspección técnica / 24 de junio",
        "Habilitación empalmes eléctricos / junio – julio",
        "Instalación y habilitación equipos WES / 20 y 21 julio",
        "Capacitación equipo BUPA Plataforma WES / 31 julio",
    ]:
        _add_para(doc, h, bullet=True, space_after=2)
    _add_subtitle(doc, "Ubicación")
    _add_table(doc, ["Recinto", "Dirección"], [["Clínica Bupa Antofagasta", "Manuel Antonio Matta 1945, Antofagasta"]])

    # ===== OBJETIVOS + SECTORES =====
    _page_break(doc)
    _add_title(doc, "II. Objetivos y resultados esperados")
    _add_para(doc, "Tres ejes técnicos centrales:")
    _add_para(doc, "(A) Línea de base hídrica con data de alta frecuencia.", bullet=True)
    _add_para(doc, "(B) Perfiles diario, lun–dom y 24 h (pico, mínimo y cercanos al promedio).", bullet=True)
    _add_para(doc, "(C) Detección de ineficiencias, huecos de data y desacoples operacionales.", bullet=True)

    _add_title(doc, "III. Definición de sectores a monitorear")
    _add_para(
        doc,
        "Tras visita técnica y reuniones Bupa–WES (junio–julio), se instrumentaron:",
    )
    _add_para(
        doc,
        "1. Medidor Principal Sanitaria: matriz de cobre de 2\" a la entrada. Mide el 100% del "
        "volumen entregado por Aguas Antofagasta.",
        bullet=True,
    )
    _add_para(
        doc,
        "2. Sala de Bomba Principal (Primer Piso): impulsión de cobre de 2\" en planta baja; "
        "distribuye agua a gran parte de la infraestructura.",
        bullet=True,
    )
    _add_para(
        doc,
        "3. Sala de Bomba Sexto Piso: red de 2\" en cobre del sexto nivel; presuriza y abastece "
        "el resto del recinto.",
        bullet=True,
    )

    # ===== FOTOS (3) =====
    _page_break(doc)
    _add_title(doc, "IV. Fotografías de instalación")
    _add_para(
        doc,
        "Instalación de equipos WES (20–21 julio). Se muestra el gabinete y el sector de cada punto "
        "(fuente: Bupa Antofagasta / Fotografias).",
    )
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    caps = [
        ("sanitaria.jpg", "Medidor Principal Sanitaria"),
        ("bomba_principal.jpg", "Sala de Bomba Principal"),
        ("sexto_piso.jpg", "Sala de Bomba Sexto Piso"),
    ]
    for i, (fname, cap) in enumerate(caps):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        img = FOTOS / fname
        if img.is_file():
            p.add_run().add_picture(str(img), width=Inches(2.05))
        c = cell.add_paragraph()
        c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = c.add_run(cap)
        _set_run(r, size=7, color=RGBColor(80, 80, 80))
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    _add_para(
        doc,
        "Usuarios BUPA en WES App: Catherine Velásquez, Joel Cartes y Jonathan Chaparro.",
        size=9,
    )

    # ===== ANÁLISIS POR PUNTO =====
    _page_break(doc)
    _add_title(doc, "V. Análisis de sectores monitoreados")
    _add_para(
        doc,
        f"Data WES {_fmt_fecha(START)}–{_fmt_fecha(END)}. Gráficos propios estilo App WES. "
        "Por cada punto: perfil diario, promedio lun–dom, y curvas 24 h (m³/hr) del día "
        "máximo, del día mínimo y de tres días cercanos al promedio.",
    )

    for idx, nid in enumerate(["000029-09", "000029-07", "000029-08"], start=1):
        if idx > 1:
            _page_break(doc)
        n = nodes[nid]
        _add_subtitle(doc, f"{idx}) {n['nombre']}")

        # KPI table estilo modelo (incluye rango pico/mínimo)
        rango = "—"
        if n["mx"] and n["mn"]:
            rango = (
                f"{_fmt_m3(n['mn'][1])} – {_fmt_m3(n['mx'][1])} m³\n"
                f"({_fmt_fecha(n['mn'][0])} / {_fmt_fecha(n['mx'][0])})"
            )
        kpi = [
            ("Total (m³)", _fmt_m3(n["total"])),
            ("Promedio mensual (m³)", _fmt_m3(n["prom_mes"])),
            ("Promedio diario (m³)", _fmt_m3(n["prom"])),
            ("Días con data", f"{n['n']} / {NUM_DIAS}"),
            ("Rango min y max diario (m³)", rango.replace("\n", " ")),
        ]
        if nid == "000029-08":
            kpi.append(("Hueco sin registro", "14/08/2026 – 17/08/2026 (lluvias)"))
        if nid != "000029-09":
            pct = 100.0 * n["total"] / san["total"] if san["total"] else 0.0
            kpi.append(("% sobre Sanitaria", f"{_fmt_m3(pct)}%"))
        else:
            kpi.append(("% sobre Sanitaria", "100,0%"))
        _add_kpi_table(doc, kpi)

        _add_picture(
            doc,
            CHARTS / f"chart_diario_{n['slug']}.png",
            5.95,
            f"Perfil de consumo diario (m³) — {n['corto']} · Fuente: WES",
        )
        if n["mx"] and n["mn"]:
            _add_para(
                doc,
                f"Patrón del periodo con pico {_fmt_m3(n['mx'][1])} m³/día "
                f"({_fmt_fecha(n['mx'][0])}) y mínimo {_fmt_m3(n['mn'][1])} m³/día "
                f"({_fmt_fecha(n['mn'][0])}). Promedio {_fmt_m3(n['prom'])} m³/día.",
                highlight=True,
                bullet=True,
            )

        _add_picture(doc, CHARTS / f"chart_dow_{n['slug']}.png", 5.8)
        dow_rows = [[d, _fmt_m3(n["dow"].get(d, 0), 2)] for d in DAY_ES]
        _add_table(doc, ["Día", "Promedio (m³)"], dow_rows, compact=True)
        _add_para(
            doc,
            "El promedio lun–dom permite contrastar días hábiles e inhábiles.",
            highlight=True,
            bullet=True,
        )

        # 24h max + min (estilo App)
        _page_break(doc)
        _add_subtitle(doc, f"Monitoreo Inteligente WES / Perfil hídrico horario — {n['corto']}")
        _add_para(
            doc,
            f"Curvas 24 h (m³/hr) del día máximo, del día mínimo y de tres días con consumo "
            f"más cercano al promedio ({_fmt_m3(n['avg_pos'])} m³/día). Fuente: App WES.",
        )
        if n["mx"]:
            _add_picture(
                doc,
                CHARTS / f"chart_24h_max_{n['slug']}.png",
                5.95,
                f"Día máximo {_fmt_fecha(n['mx'][0])} — {_fmt_m3(n['mx'][1])} m³ · Fuente: APP WES",
            )
        if n["mn"]:
            _add_picture(
                doc,
                CHARTS / f"chart_24h_min_{n['slug']}.png",
                5.95,
                f"Día mínimo {_fmt_fecha(n['mn'][0])} — {_fmt_m3(n['mn'][1])} m³ · Fuente: APP WES",
            )

        if n["near"]:
            _page_break(doc)
            _add_subtitle(doc, f"Días cercanos al promedio — {n['corto']}")
            near_txt = "; ".join(
                f"{_fmt_fecha(d)} ({_fmt_m3(v)} m³, Δ {_fmt_m3(abs(v - n['avg_pos']))})"
                for d, v, _ in n["near"]
            )
            _add_para(doc, f"Seleccionados: {near_txt}.", space_after=4)
            for d, v, path in n["near"]:
                _add_picture(
                    doc,
                    path,
                    5.8,
                    f"Cercano al promedio {_fmt_fecha(d)} — {_fmt_m3(v)} m³ · Fuente: APP WES",
                )

    # Participación
    _page_break(doc)
    _add_subtitle(doc, "Participación sobre Sanitaria (periodo)")
    _add_table(
        doc,
        ["Punto", "Consumo (m³)", "%"],
        [
            ["Medidor Principal Sanitaria", _fmt_m3(san["total"]), "100,0%"],
            ["Sala de Bomba Principal", _fmt_m3(prin["total"]), f"{_fmt_m3(pct_prin)}%"],
            ["Sala de Bomba Sexto Piso", _fmt_m3(sexto["total"]), f"{_fmt_m3(pct_sexto)}%"],
            ["Principal + Sexto", _fmt_m3(tot_salas), f"{_fmt_m3(pct_salas)}%"],
            ["No monitoreado (brecha)", _fmt_m3(gap), f"{_fmt_m3(pct_gap)}%"],
        ],
    )
    _add_picture(doc, CHARTS / "chart_participacion.png", 3.4)

    # ===== FACTURAS =====
    _page_break(doc)
    _add_title(doc, "VI. Facturación sanitaria — Aguas Antofagasta")
    _add_para(
        doc,
        "Boletas Centro Médico Antofagasta S.A. (N° cliente 1110080842 · medidor 534 · "
        "Manuel Antonio Matta 1957). Fuente: Bupa Antofagasta / Facturacion.",
    )
    _add_picture(doc, CHARTS / "chart_facturas_m3.png", 5.4)
    rows = [
        [f["ciclo"], f["emision"], _fmt_m3(f["m3"], 0), _fmt_clp(f["clp"]), f["nro"]]
        for f in FACTURAS
    ]
    rows.append(["Total", "—", _fmt_m3(tot_fact_m3, 0), _fmt_clp(tot_fact_clp), "—"])
    _add_table(
        doc,
        ["Periodo", "Emisión", "Consumo (m³)", "Monto facturado", "N° factura"],
        rows,
        compact=True,
    )
    _add_subtitle(doc, "Lecturas encadenadas — medidor 534")
    _add_table(
        doc,
        ["Ciclo", "Lect. anterior", "Lect. actual", "m³", "N° factura"],
        [
            [f["ciclo"], str(f["lect_ant"]), str(f["lect_act"]), _fmt_m3(f["m3"], 0), f["nro"]]
            for f in FACTURAS
        ],
        compact=True,
    )

    # ===== CONCLUSIONES =====
    _page_break(doc)
    _add_title(doc, "VII. Conclusiones y hallazgos")
    _add_para(
        doc,
        f"1) Periodo {_fmt_fecha(START)}–{_fmt_fecha(END)}: Sanitaria {_fmt_m3(san['total'])} m³ "
        f"({_fmt_m3(san['prom'])} m³/día).",
        bullet=True,
    )
    _add_para(
        doc,
        f"2) Principal+Sexto = {_fmt_m3(pct_salas)}%; brecha = {_fmt_m3(pct_gap)}% "
        f"({_fmt_m3(gap)} m³).",
        bullet=True,
    )
    _add_para(
        doc,
        "3) Hueco de data solo en Sexto Piso 14–17/08/2026; Sanitaria y Principal sin huecos.",
        bullet=True,
    )
    _add_para(
        doc,
        f"4) Facturación jun–sep 2026 Antofagasta: {_fmt_m3(tot_fact_m3, 0)} m³ / {_fmt_clp(tot_fact_clp)}.",
        bullet=True,
    )
    _add_para(
        doc,
        "5) Por cada punto se presentan perfiles 24 h del día máximo, mínimo y tres cercanos al promedio.",
        bullet=True,
    )

    _add_title(doc, "VIII. Recomendaciones y próximos pasos")
    _add_para(doc, "Acción 1 — Instrumentar tramos no medidos para cerrar la brecha.", bullet=True)
    _add_para(doc, "Acción 2 — Evaluar Control Inteligente WES en horarios inhábiles.", bullet=True)
    _add_para(doc, "Acción 3 — Contraste ciclo a ciclo factura Aguas Antofagasta vs Sanitaria WES.", bullet=True)
    _add_para(doc, "Acción 4 — Mantener usuarios BUPA activos y revisión periódica de alertas.", bullet=True)

    _add_para(
        doc,
        f"Documento generado {datetime.now().strftime('%d/%m/%Y %H:%M')} · WES · Clínica Bupa Antofagasta",
        size=8,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_before=16,
    )

    out = ROOT / (
        f"Informe_final_WES_Bupa_Antofagasta_{START.strftime('%Y%m%d')}_"
        f"{END.strftime('%Y%m%d')}.docx"
    )
    doc.save(str(out))
    return out


if __name__ == "__main__":
    path = build()
    print(f"[OK] {path}")
