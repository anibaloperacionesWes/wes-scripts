"""
Análisis piloto — Clínica Bupa Antofagasta.

Periodo limpio: 23/07/2026 → domingo recién pasado (sin semana en curso).
Incluye:
  - Huecos sin data por punto (exactos)
  - 4 puntos: día a día + promedio lun–dom + tabla (una hoja por punto)
  - Participación % vs Sanitaria (mensual/semanal) incluyendo Sala N°2
  - Validación de no-duplicados en la serie diaria

Uso:
  python generar_analisis_piloto_bupa_antofagasta.py
"""

from __future__ import annotations

import calendar
import sys
import time
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.text import WD_BREAK, WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from generar_reporte_word import (
    MeasurePoint,
    acl_node_base_url,
    add_formatted_title,
    add_table,
    calculate_weekly_averages,
    fetch_json,
    flatten_measures,
    format_number_chilean,
    normalize_measures_payload,
    summarize_consumption,
)

# ---------------------------------------------------------------------------
# Alcance
# ---------------------------------------------------------------------------
FOLDER = "Bupa_Antofagasta"
START = date(2026, 7, 23)  # desde aquí la serie está correcta
END = date(2026, 10, 11)  # domingo pasado; sin esta semana
LLUVIAS_SIN_DATA = (date(2026, 8, 14), date(2026, 8, 17))  # Sexto Piso
SALA2_CAMBIO = date(2026, 9, 10)

# Los cuatro puntos monitoreados (orden de presentación).
NODOS = [
    {"id": "000029-09", "nombre": "Medidor Principal Sanitaria", "corto": "Sanitaria", "rol": "matriz"},
    {"id": "000029-07", "nombre": "Sala de Bomba Principal", "corto": "Bomba Principal", "rol": "sala"},
    {"id": "000029-08", "nombre": "Sala de Bomba Sexto Piso", "corto": "Sexto Piso", "rol": "sala"},
    {"id": "000029-10", "nombre": "Sala de Bomba N°2", "corto": "Bomba N°2", "rol": "sala2"},
]

MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}


def _fmt_fecha(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def _set_run_black(run) -> None:
    run.font.color.rgb = RGBColor(0, 0, 0)


def _add_body(doc: Document, text: str, bold_prefix: str = "", space_after: int = 4) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    if bold_prefix:
        r = p.add_run(bold_prefix)
        r.bold = True
        _set_run_black(r)
    r2 = p.add_run(text)
    _set_run_black(r2)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)


def _add_compact_title(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    r = p.add_run(text.upper())
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(0, 51, 102)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True


def _add_picture_compact(doc: Document, image_path: str, width_in: float) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.add_run().add_picture(image_path, width=Inches(width_in))
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True


def _page_break(doc: Document) -> None:
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def _compact_last_table(doc: Document) -> None:
    """Reduce tipografía/espaciado de la última tabla para que no se parta entre páginas."""
    if not doc.tables:
        return
    table = doc.tables[-1]
    from docx.oxml import OxmlElement

    # Evitar que filas individuales se partan; encabezado se repite si igual hubiera overflow.
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl.insert(0, tbl_pr)

    for i, row in enumerate(table.rows):
        tr_pr = row._tr.get_or_add_trPr()
        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))
        if i == 0 and tr_pr.find(qn("w:tblHeader")) is None:
            tr_pr.append(OxmlElement("w:tblHeader"))
        for cell in row.cells:
            for para in cell.paragraphs:
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(0)
                for run in para.runs:
                    run.font.size = Pt(7 if i > 0 else 8)


def _fetch_series(node_id: str) -> Tuple[List[MeasurePoint], dict]:
    """Retorna medidas del periodo + meta de validación (duplicados)."""
    raw = fetch_json(
        f"{acl_node_base_url()}/nodes/measures/dates",
        params=[
            ("id", node_id),
            ("start", START.strftime("%d%m%Y")),
            ("end", END.strftime("%d%m%Y")),
        ],
    )
    # Validar duplicados en payload crudo
    months = []
    if isinstance(raw, list) and raw:
        months = raw[0].get("month") or []
    elif isinstance(raw, dict):
        months = raw.get("month") or []
    dates_raw = [m.get("date") for m in months if m.get("date")]
    counts = Counter(dates_raw)
    dups = {d: n for d, n in counts.items() if n > 1}

    measures = flatten_measures(normalize_measures_payload(raw, node_id))
    out = []
    for m in measures:
        d = m.date.date() if hasattr(m.date, "date") else m.date
        if START <= d <= END:
            out.append(m)
    out.sort(key=lambda x: x.date)
    dates_flat = [(m.date.date() if hasattr(m.date, "date") else m.date) for m in out]
    flat_dups = {d: n for d, n in Counter(dates_flat).items() if n > 1}

    meta = {
        "raw_entries": len(dates_raw),
        "raw_unique": len(counts),
        "raw_dups": dups,
        "flat_dups": flat_dups,
    }
    return out, meta


def _by_date(measures: List[MeasurePoint]) -> Dict[date, float]:
    # Si hubiera duplicados en flatten, conservar el último (y se reporta en meta).
    out: Dict[date, float] = {}
    for m in measures:
        d = m.date.date() if hasattr(m.date, "date") else m.date
        out[d] = float(m.total_m3 or 0.0)
    return out


def _gap_ranges(by: Dict[date, float], start: date, end: date) -> List[Tuple[date, date, int]]:
    missing = []
    cur = start
    while cur <= end:
        if cur not in by:
            missing.append(cur)
        cur += timedelta(days=1)
    if not missing:
        return []
    missing.sort()
    ranges = []
    s = e = missing[0]
    for d in missing[1:]:
        if d == e + timedelta(days=1):
            e = d
        else:
            ranges.append((s, e, (e - s).days + 1))
            s = e = d
    ranges.append((s, e, (e - s).days + 1))
    return ranges


def _sum_range(by: Dict[date, float], d0: date, d1: date) -> float:
    total = 0.0
    cur = d0
    while cur <= d1:
        total += by.get(cur, 0.0)
        cur += timedelta(days=1)
    return total


def _month_windows(start: date, end: date) -> List[Tuple[date, date, str]]:
    windows = []
    y, m = start.year, start.month
    while True:
        first = date(y, m, 1)
        last = date(y, m, calendar.monthrange(y, m)[1])
        d0 = max(first, start)
        d1 = min(last, end)
        if d0 <= d1:
            label = f"{MESES_ES[m].capitalize()} {y}"
            if d0 != first or d1 != last:
                label += f" ({_fmt_fecha(d0)}–{_fmt_fecha(d1)})"
            windows.append((d0, d1, label))
        if y == end.year and m == end.month:
            break
        if m == 12:
            y, m = y + 1, 1
        else:
            m += 1
    return windows


def _week_windows(start: date, end: date) -> List[Tuple[date, date, str]]:
    first_monday = start - timedelta(days=start.weekday())
    windows = []
    cur = first_monday
    n = 1
    while cur <= end:
        week_end = cur + timedelta(days=6)
        d0 = max(cur, start)
        d1 = min(week_end, end)
        if d0 <= d1:
            windows.append((d0, d1, f"Semana {n} ({_fmt_fecha(d0)}–{_fmt_fecha(d1)})"))
            n += 1
        cur += timedelta(days=7)
    return windows


def _chart_daily_compact(measures: List[MeasurePoint], path: Path, title: str) -> Optional[Path]:
    if not measures:
        return None
    dates = [m.date for m in measures]
    totals = [m.total_m3 for m in measures]
    fig, ax = plt.subplots(figsize=(9.0, 2.25))
    ax.plot(dates, totals, marker="o", markersize=2.2, linestyle="-", color="#0050b3", linewidth=1.3)
    ax.fill_between(dates, totals, alpha=0.22, color="#0050b3")
    ax.set_title(title, fontsize=11, fontweight="bold", pad=4)
    ax.set_ylabel("m³", fontsize=9)
    ax.set_ylim(bottom=0)
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(dates) // 10)))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%b"))
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=35, ha="right", fontsize=7)
    ax.tick_params(axis="y", labelsize=8)
    ax.grid(True, linestyle="--", alpha=0.3, axis="y")
    fig.tight_layout(pad=0.3)
    fig.savefig(path, dpi=140, bbox_inches="tight")
    plt.close(fig)
    return path


def _chart_weekly_compact(weekly_data: dict, path: Path, title: str) -> Optional[Path]:
    if not weekly_data or "averages" not in weekly_data:
        return None
    day_names = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    values = [weekly_data["averages"].get(d, 0.0) for d in day_names]
    fig, ax = plt.subplots(figsize=(9.0, 2.15))
    bars = ax.bar(day_names, values, color="#0050b3", alpha=0.85, edgecolor="#003d8c", linewidth=1.0)
    for bar, value in zip(bars, values):
        ax.text(
            bar.get_x() + bar.get_width() / 2.0,
            bar.get_height(),
            f"{format_number_chilean(value, 1)}",
            ha="center",
            va="bottom",
            fontsize=8,
            fontweight="bold",
        )
    ax.set_title(title, fontsize=11, fontweight="bold", pad=4)
    ax.set_ylabel("m³", fontsize=9)
    ax.set_ylim(bottom=0)
    ax.tick_params(axis="x", labelsize=8)
    ax.tick_params(axis="y", labelsize=8)
    ax.grid(True, linestyle="--", alpha=0.3, axis="y")
    fig.tight_layout(pad=0.3)
    fig.savefig(path, dpi=140, bbox_inches="tight")
    plt.close(fig)
    return path


def _chart_participacion_barras(
    path: Path,
    labels: List[str],
    pct_principal: List[float],
    pct_sexto: List[float],
    pct_n2: List[float],
    title: str,
) -> Path:
    """Barras agrupadas (no apiladas a 100%) para no ocultar N°2 cuando supera cuenta."""
    fig, ax = plt.subplots(figsize=(10.5, 4.2))
    x = list(range(len(labels)))
    w = 0.22
    ax.bar([i - 1.5 * w for i in x], pct_principal, w, label="Principal", color="#0050b3")
    ax.bar([i - 0.5 * w for i in x], pct_sexto, w, label="Sexto Piso", color="#40a9ff")
    ax.bar([i + 0.5 * w for i in x], pct_n2, w, label="Bomba N°2", color="#fa8c16")
    # gap = 100 - principal - sexto (balance de los dos puntos de abajo; N2 aparte)
    gap = [max(0.0, 100.0 - a - b) for a, b in zip(pct_principal, pct_sexto)]
    ax.bar([i + 1.5 * w for i in x], gap, w, label="No mon. (vs Princ.+Sexto)", color="#c0392b", alpha=0.85)
    ax.axhline(100, color="#333", linestyle="--", linewidth=1, alpha=0.45)
    ax.set_ylabel("% sobre Sanitaria", fontsize=10, fontweight="bold")
    ax.set_title(title, fontsize=12, fontweight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=35, ha="right", fontsize=7)
    ax.set_ylim(0, max(110, max(pct_n2 + pct_principal + [0]) + 10))
    ax.legend(loc="upper left", fontsize=8, ncol=2)
    ax.grid(True, axis="y", linestyle="--", alpha=0.3)
    plt.tight_layout()
    fig.savefig(path, dpi=140, bbox_inches="tight")
    plt.close(fig)
    return path


def _chart_pie_small(path: Path, salas: float, gap: float, sanit: float) -> Path:
    fig, ax = plt.subplots(figsize=(3.4, 2.6))
    vals = [max(salas, 0), max(gap, 0)]
    if sum(vals) <= 0:
        vals = [1, 0]
    wedges, texts, autotexts = ax.pie(
        vals,
        labels=["Princ.+Sexto", "No mon."],
        colors=["#0050b3", "#c0392b"],
        autopct=lambda p: f"{p:.0f}%",
        startangle=90,
        textprops={"fontsize": 8},
    )
    for t in autotexts:
        t.set_fontsize(8)
    for w in wedges:
        w.set_edgecolor("white")
        w.set_linewidth(1.0)
    ax.set_title("Periodo vs Sanitaria", fontsize=9, fontweight="bold", pad=2)
    plt.tight_layout(pad=0.2)
    fig.savefig(path, dpi=130, bbox_inches="tight")
    plt.close(fig)
    return path


def _convertir_pdf(docx_path: Path) -> Optional[Path]:
    pdf_path = docx_path.with_suffix(".pdf")
    try:
        import subprocess

        subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(docx_path.parent),
                str(docx_path),
            ],
            check=True,
            capture_output=True,
            timeout=180,
        )
        if pdf_path.exists():
            return pdf_path
    except Exception as e:
        print(f"[ADVERTENCIA] Conversión PDF falló: {e}")
    return None


def _setup_margins(doc: Document) -> None:
    for section in doc.sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)


def main() -> None:
    print("=" * 70)
    print("[INFO] Bupa Antofagasta — análisis piloto (v3)")
    print(f"       Periodo: {_fmt_fecha(START)} → {_fmt_fecha(END)}")
    print("=" * 70)
    t0 = time.perf_counter()

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = Path("reports") / FOLDER / "ANALISIS_PILOTO" / f"PILOTO_{stamp}"
    out_dir.mkdir(parents=True, exist_ok=True)

    series: Dict[str, List[MeasurePoint]] = {}
    by_date: Dict[str, Dict[date, float]] = {}
    meta_dup: Dict[str, dict] = {}
    for n in NODOS:
        print(f"[INFO] Descargando {n['nombre']}...")
        measures, meta = _fetch_series(n["id"])
        series[n["id"]] = measures
        by_date[n["id"]] = _by_date(measures)
        meta_dup[n["id"]] = meta
        print(
            f"       {len(measures)} días · {sum(by_date[n['id']].values()):.1f} m³ · "
            f"dups_raw={meta['raw_dups'] or 'no'} dups_flat={meta['flat_dups'] or 'no'}"
        )

    sanit_id = "000029-09"
    sal_07 = "000029-07"
    sal_08 = "000029-08"
    sal_10 = "000029-10"

    tot_san = sum(by_date[sanit_id].values())
    tot_07 = sum(by_date[sal_07].values())
    tot_08 = sum(by_date[sal_08].values())
    tot_10 = sum(by_date[sal_10].values())
    tot_salas = tot_07 + tot_08
    tot_gap = max(0.0, tot_san - tot_salas)
    pct_07 = 100.0 * tot_07 / tot_san if tot_san else 0.0
    pct_08 = 100.0 * tot_08 / tot_san if tot_san else 0.0
    pct_10 = 100.0 * tot_10 / tot_san if tot_san else 0.0
    pct_salas = 100.0 * tot_salas / tot_san if tot_san else 0.0
    pct_gap = 100.0 * tot_gap / tot_san if tot_san else 0.0

    num_dias = (END - START).days + 1
    prom_san = tot_san / max(len(by_date[sanit_id]), 1)
    prom_07 = tot_07 / max(len(by_date[sal_07]), 1)
    prom_08 = tot_08 / max(len(by_date[sal_08]), 1)
    prom_10 = tot_10 / max(len(by_date[sal_10]), 1)

    # Huecos por punto
    gaps_por_punto = {n["id"]: _gap_ranges(by_date[n["id"]], START, END) for n in NODOS}

    doc = Document()
    _setup_margins(doc)

    title = doc.add_paragraph()
    tr = title.add_run(
        "ANÁLISIS PILOTO — CLÍNICA BUPA ANTOFAGASTA\n"
        f"Consumo {_fmt_fecha(START)} al {_fmt_fecha(END)}"
    )
    tr.bold = True
    tr.font.size = Pt(15)
    tr.font.color.rgb = RGBColor(0, 51, 102)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    title.paragraph_format.space_after = Pt(8)

    # --- 1. Huecos ---
    add_formatted_title(doc, "1. Disponibilidad de data por punto (rangos exactos)")
    _add_body(
        doc,
        "Se revisó día a día cada punto entre el inicio y el fin del periodo. "
        "Donde falta registro se indica el rango exacto.",
        space_after=4,
    )
    gap_rows = [("Punto", "Días con data", "Huecos sin registro", "Observación")]
    for n in NODOS:
        nid = n["id"]
        n_ok = len(by_date[nid])
        gaps = gaps_por_punto[nid]
        if gaps:
            huecos = "; ".join(
                f"{_fmt_fecha(a)}–{_fmt_fecha(b)} ({nd} día{'s' if nd > 1 else ''})"
                for a, b, nd in gaps
            )
        else:
            huecos = "Ninguno"
        obs = "Serie continua en el periodo."
        if n["corto"] == "Sexto Piso" and gaps:
            obs = "Hueco asociado a lluvias en Antofagasta (afectó conectividad/data)."
        if n["corto"] == "Bomba N°2":
            obs = (
                f"Sin huecos de registro. Cambio de régimen desde "
                f"{_fmt_fecha(SALA2_CAMBIO)} (ver sección 6)."
            )
        gap_rows.append((n["nombre"], f"{n_ok} / {num_dias}", huecos, obs))
    add_table(doc, "Cobertura de data", gap_rows, wes_style=True, has_total_row=False)

    # --- 2. Resumen ---
    add_formatted_title(doc, "2. Resumen ejecutivo")
    table_resumen = [
        ("Punto", "Consumo periodo (m³)", "Promedio diario (m³)", "% sobre Sanitaria"),
        (
            "Medidor Principal Sanitaria (100%)",
            format_number_chilean(tot_san, 1),
            format_number_chilean(prom_san, 1),
            "100,0%",
        ),
        (
            "Sala de Bomba Principal",
            format_number_chilean(tot_07, 1),
            format_number_chilean(prom_07, 1),
            f"{format_number_chilean(pct_07, 1)}%",
        ),
        (
            "Sala de Bomba Sexto Piso",
            format_number_chilean(tot_08, 1),
            format_number_chilean(prom_08, 1),
            f"{format_number_chilean(pct_08, 1)}%",
        ),
        (
            "Sala de Bomba N°2",
            format_number_chilean(tot_10, 1),
            format_number_chilean(prom_10, 1),
            f"{format_number_chilean(pct_10, 1)}%",
        ),
        (
            "TOTAL Principal + Sexto",
            format_number_chilean(tot_salas, 1),
            format_number_chilean(tot_salas / num_dias, 1),
            f"{format_number_chilean(pct_salas, 1)}%",
        ),
        (
            "NO MONITOREADO (Sanitaria − Princ. − Sexto)",
            format_number_chilean(tot_gap, 1),
            format_number_chilean(tot_gap / num_dias, 1),
            f"{format_number_chilean(pct_gap, 1)}%",
        ),
    ]
    add_table(doc, "Resumen del periodo", table_resumen, wes_style=True)

    _add_body(
        doc,
        f"Las dos salas de abajo (Principal + Sexto) suman "
        f"{format_number_chilean(pct_salas, 1)}% de Sanitaria; queda "
        f"{format_number_chilean(pct_gap, 1)}% no monitoreado "
        f"({format_number_chilean(tot_gap, 1)} m³). "
        f"Sala N°2 se muestra aparte en el % (ver secciones 5–6): no se suma al "
        f"balance de “abajo” porque en la segunda etapa deja de ser coherente "
        f"con Sanitaria = 100%.",
        bold_prefix="Lectura: ",
        space_after=4,
    )

    pie_path = _chart_pie_small(
        out_dir / "chart_participacion_periodo.png",
        tot_salas,
        tot_gap,
        tot_san,
    )
    # torta chica
    _add_picture_compact(doc, str(pie_path), 2.4)

    # --- 3. Cuatro puntos: intro + 3.1 en la misma hoja; luego una hoja por punto ---
    _page_break(doc)
    add_formatted_title(doc, "3. Consumo monitoreado — cuatro puntos")
    _add_body(
        doc,
        "Por cada punto: consumo día a día, promedio por día de la semana (lun–dom) "
        "y tabla de promedios. Cada punto está en una sola hoja.",
        space_after=2,
    )

    for idx, n in enumerate(NODOS, start=1):
        if idx > 1:
            _page_break(doc)
        nid = n["id"]
        measures = series[nid]
        summary = summarize_consumption(measures)
        slug = n["corto"].lower().replace(" ", "_").replace("°", "").replace(".", "")

        _add_compact_title(doc, f"3.{idx} {n['nombre']}")
        line = (
            f"Total: {format_number_chilean(summary['total'], 1)} m³ · "
            f"Promedio: {format_number_chilean(summary['promedio_diario'], 1)} m³/día"
        )
        if summary.get("max"):
            line += (
                f" · Pico: {summary['max'].date.strftime('%d-%m-%y')} "
                f"({format_number_chilean(summary['max'].total_m3, 1)} m³)"
            )
        gaps = gaps_por_punto[nid]
        if gaps:
            line += " · Sin data: " + "; ".join(
                f"{_fmt_fecha(a)}–{_fmt_fecha(b)}" for a, b, _ in gaps
            )
        else:
            line += " · Sin huecos de registro en el periodo."
        _add_body(doc, line, space_after=2)

        chart_d = _chart_daily_compact(
            measures,
            out_dir / f"chart_diario_{slug}.png",
            f"Consumo diario — {n['corto']} (m³)",
        )
        if chart_d:
            _add_picture_compact(doc, str(chart_d), 5.9)

        weekly = calculate_weekly_averages(measures)
        if weekly:
            chart_w = _chart_weekly_compact(
                weekly,
                out_dir / f"chart_promedio_dow_{slug}.png",
                f"Promedio lun–dom — {n['corto']} ({len(weekly['weeks'])} semanas)",
            )
            if chart_w:
                _add_picture_compact(doc, str(chart_w), 5.9)
            avg = weekly["averages"]
            orden = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
            rows = [("Día", "Promedio (m³)")] + [
                (d, format_number_chilean(avg.get(d, 0), 2)) for d in orden
            ]
            # Tabla compacta: sin párrafo vacío extra — add_table agrega uno; lo aceptamos
            add_table(
                doc,
                f"Promedio lun–dom — {n['corto']}",
                rows,
                wes_style=True,
                has_total_row=False,
            )
        else:
            _add_body(doc, "Sin suficientes semanas completas para promedio lun–dom.")

    # --- 4. Participación mensual (con N°2) ---
    _page_break(doc)
    add_formatted_title(doc, "4. Participación % vs Sanitaria — por mes")
    _add_body(
        doc,
        "Sanitaria = 100%. Se incluye Sala de Bomba N°2 en todas las filas "
        "(aunque en julio–agosto aporte ~0–1%).",
    )

    month_wins = _month_windows(START, END)
    table_m = [
        (
            "Mes",
            "Sanitaria m³",
            "Principal %",
            "Sexto %",
            "Bomba N°2 %",
        )
    ]
    labels_m, p07_m, p08_m, p10_m = [], [], [], []
    for d0, d1, label in month_wins:
        s = _sum_range(by_date[sanit_id], d0, d1)
        a = _sum_range(by_date[sal_07], d0, d1)
        b = _sum_range(by_date[sal_08], d0, d1)
        c = _sum_range(by_date[sal_10], d0, d1)
        pa = 100.0 * a / s if s else 0.0
        pb = 100.0 * b / s if s else 0.0
        pc = 100.0 * c / s if s else 0.0
        table_m.append(
            (
                label,
                format_number_chilean(s, 1),
                f"{format_number_chilean(pa, 1)}%",
                f"{format_number_chilean(pb, 1)}%",
                f"{format_number_chilean(pc, 1)}%",
            )
        )
        labels_m.append(label.split("(")[0].strip())
        p07_m.append(pa)
        p08_m.append(pb)
        p10_m.append(pc)

    add_table(doc, "Participación mensual sobre Sanitaria", table_m, wes_style=True, has_total_row=False)
    chart_m = _chart_participacion_barras(
        out_dir / "chart_participacion_mensual.png",
        labels_m,
        p07_m,
        p08_m,
        p10_m,
        "Participación mensual (% sobre Sanitaria)",
    )
    _add_picture_compact(doc, str(chart_m), 6.2)

    # --- 5. Participación semanal (con N°2): tabla completa en una hoja ---
    _page_break(doc)
    add_formatted_title(doc, "5. Participación % vs Sanitaria — por semana")
    _add_body(
        doc,
        "Misma lógica semanal (lun–dom), incluyendo Bomba N°2 en cada semana.",
        space_after=2,
    )
    week_wins = _week_windows(START, END)
    table_w = [
        (
            "Semana",
            "Sanitaria m³",
            "Principal %",
            "Sexto %",
            "Bomba N°2 %",
            "Princ.+Sexto %",
            "No mon. %",
        )
    ]
    labels_w, p07_w, p08_w, p10_w = [], [], [], []
    for d0, d1, label in week_wins:
        s = _sum_range(by_date[sanit_id], d0, d1)
        a = _sum_range(by_date[sal_07], d0, d1)
        b = _sum_range(by_date[sal_08], d0, d1)
        c = _sum_range(by_date[sal_10], d0, d1)
        salas = a + b
        gap = max(0.0, s - salas)
        pa = 100.0 * a / s if s else 0.0
        pb = 100.0 * b / s if s else 0.0
        pc = 100.0 * c / s if s else 0.0
        ps = 100.0 * salas / s if s else 0.0
        pg = 100.0 * gap / s if s else 0.0
        # Etiqueta corta para que la tabla quepa en una hoja
        short_label = f"{d0.strftime('%d/%m')}–{d1.strftime('%d/%m')}"
        table_w.append(
            (
                short_label,
                format_number_chilean(s, 1),
                f"{format_number_chilean(pa, 1)}%",
                f"{format_number_chilean(pb, 1)}%",
                f"{format_number_chilean(pc, 1)}%",
                f"{format_number_chilean(ps, 1)}%",
                f"{format_number_chilean(pg, 1)}%",
            )
        )
        labels_w.append(d0.strftime("%d/%m"))
        p07_w.append(pa)
        p08_w.append(pb)
        p10_w.append(pc)

    # Tabla primero (completa en esta hoja); el gráfico va después / puede seguir en la siguiente.
    add_table(doc, "Participación semanal sobre Sanitaria", table_w, wes_style=True, has_total_row=False)
    _compact_last_table(doc)
    chart_w = _chart_participacion_barras(
        out_dir / "chart_participacion_semanal.png",
        labels_w,
        p07_w,
        p08_w,
        p10_w,
        "Participación semanal (% sobre Sanitaria)",
    )
    _add_picture_compact(doc, str(chart_w), 6.0)

    # --- 6. Sala N°2 explicación ---
    add_formatted_title(doc, "6. Qué está pasando en Sala de Bomba N°2")
    by10 = by_date[sal_10]
    pre = [v for d, v in by10.items() if d < SALA2_CAMBIO]
    post = [v for d, v in by10.items() if d >= SALA2_CAMBIO]
    avg_pre = sum(pre) / len(pre) if pre else 0.0
    avg_post = sum(post) / len(post) if post else 0.0
    tot_pre = sum(pre)
    tot_post = sum(post)
    san_pre = sum(v for d, v in by_date[sanit_id].items() if d < SALA2_CAMBIO)
    san_post = sum(v for d, v in by_date[sanit_id].items() if d >= SALA2_CAMBIO)
    pct_pre = 100.0 * tot_pre / san_pre if san_pre else 0.0
    pct_post = 100.0 * tot_post / san_post if san_post else 0.0

    _add_body(
        doc,
        "En Sala de Bomba N°2 se ven dos etapas claras:",
    )
    _add_body(
        doc,
        f"del {_fmt_fecha(START)} al {_fmt_fecha(SALA2_CAMBIO - timedelta(days=1))} "
        f"promedia {format_number_chilean(avg_pre, 1)} m³/día "
        f"({format_number_chilean(pct_pre, 1)}% de Sanitaria) — casi no aporta.",
        bold_prefix="Etapa 1: ",
    )
    _add_body(
        doc,
        f"desde el {_fmt_fecha(SALA2_CAMBIO)} promedia "
        f"{format_number_chilean(avg_post, 1)} m³/día "
        f"({format_number_chilean(pct_post, 1)}% aparente de Sanitaria). "
        f"En varios días supera el 100% de Sanitaria, lo que no cuadra si "
        f"Sanitaria mide toda el agua de entrada. Por eso no entra al balance "
        f"de Principal+Sexto hasta validar instalación/factor/empalme.",
        bold_prefix="Etapa 2: ",
    )
    add_table(
        doc,
        "Sala de Bomba N°2 — dos etapas",
        [
            ("Etapa", "Días", "Promedio (m³/día)", "% aparente vs Sanitaria"),
            (
                f"Antes del {_fmt_fecha(SALA2_CAMBIO)}",
                str(len(pre)),
                format_number_chilean(avg_pre, 1),
                f"{format_number_chilean(pct_pre, 1)}%",
            ),
            (
                f"Desde el {_fmt_fecha(SALA2_CAMBIO)}",
                str(len(post)),
                format_number_chilean(avg_post, 1),
                f"{format_number_chilean(pct_post, 1)}%",
            ),
        ],
        wes_style=True,
        has_total_row=False,
    )

    # --- 7. Conclusiones ---
    add_formatted_title(doc, "7. Conclusiones para cierre")
    ll_a, ll_b = LLUVIAS_SIN_DATA
    _add_body(
        doc,
        f"1) Periodo limpio 23/07–{_fmt_fecha(END)}.\n"
        f"2) Huecos: solo Sexto Piso {_fmt_fecha(ll_a)}–{_fmt_fecha(ll_b)} (lluvias); "
        f"Sanitaria, Principal y N°2 sin huecos de registro.\n"
        f"3) Balance: Principal+Sexto = {format_number_chilean(pct_salas, 1)}%; "
        f"no monitoreado = {format_number_chilean(pct_gap, 1)}% "
        f"({format_number_chilean(tot_gap, 1)} m³).\n"
        f"4) N°2: dos etapas; fuera del balance hasta validar medición.\n"
        f"5) Hipótesis de gap: empalme desde matriz principal antes de impulsión "
        f"2° piso.",
    )

    docx_path = out_dir / (
        f"Analisis_Piloto_Bupa_Antofagasta_{START.strftime('%Y%m%d')}_"
        f"{END.strftime('%Y%m%d')}.docx"
    )
    doc.save(str(docx_path))
    print(f"[OK] Word: {docx_path}")

    pdf_path = _convertir_pdf(docx_path)
    if pdf_path:
        print(f"[OK] PDF:  {pdf_path}")

    print("-" * 70)
    print(f"Sanitaria {tot_san:.1f} | Princ {tot_07:.1f} ({pct_07:.1f}%) | "
          f"Sexto {tot_08:.1f} ({pct_08:.1f}%) | N2 {tot_10:.1f} ({pct_10:.1f}%)")
    print(f"Salas {tot_salas:.1f} ({pct_salas:.1f}%) | Gap {tot_gap:.1f} ({pct_gap:.1f}%)")
    print(f"OUT_DIR={out_dir}")
    print(f"DOCX_PATH={docx_path}")
    if pdf_path:
        print(f"PDF_PATH={pdf_path}")
    print(f"Tiempo: {time.perf_counter()-t0:.1f}s")


if __name__ == "__main__":
    main()
