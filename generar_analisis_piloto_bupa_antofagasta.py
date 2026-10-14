"""
Análisis piloto — Clínica Bupa Antofagasta (nodos 000029-07..10).

Periodo: desde habilitación operativa del piloto hasta el domingo recién pasado
(excluye la semana en curso). Incluye:
  - Consumo día a día por punto
  - Promedio por día de la semana (lun–dom), estilo Parque Arauco
  - Participación % vs Medidor Principal Sanitaria (100%), mensual y semanal
  - Gap no monitoreado (posible empalme antes de bomba 2° piso)

Uso:
  python generar_analisis_piloto_bupa_antofagasta.py
"""

from __future__ import annotations

import calendar
import sys
import time
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
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
    add_picture_with_pagination,
    add_table,
    build_consumption_chart,
    build_weekly_averages_chart,
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
COMPANY_ID = "000029"
FOLDER = "Bupa_Antofagasta"

# Desde el 23/07 la serie ya está correcta (días previos de julio = caos de puesta en marcha).
START = date(2026, 7, 23)
# Domingo recién pasado (hoy mié 14/10 → domingo 11/10); sin esta semana.
END = date(2026, 10, 11)

# Hueco sin data en Sexto Piso por lluvias en Antofagasta.
LLUVIAS_SIN_DATA = (date(2026, 8, 14), date(2026, 8, 17))  # inclusive
# Sala N°2 cambia de régimen el 10/09 (antes ~1 m³/día; después ~58 m³/día).
SALA2_CAMBIO = date(2026, 9, 10)

# Tres puntos principales del pedido (matriz + dos salas abajo).
# N°2 se reporta aparte por lecturas anómalas desde septiembre.
NODOS_PRINCIPALES = [
    {
        "id": "000029-09",
        "nombre": "Medidor Principal Sanitaria",
        "rol": "matriz",
        "corto": "Sanitaria",
    },
    {
        "id": "000029-07",
        "nombre": "Sala de Bomba Principal",
        "rol": "sala",
        "corto": "Bomba Principal",
    },
    {
        "id": "000029-08",
        "nombre": "Sala de Bomba Sexto Piso",
        "rol": "sala",
        "corto": "Sexto Piso",
    },
]

NODO_EXTRA = {
    "id": "000029-10",
    "nombre": "Sala de Bomba N°2",
    "rol": "sala",
    "corto": "Bomba N°2",
}

MESES_ES = {
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


def _fmt_fecha(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def _add_body(doc: Document, text: str, bold_prefix: str = "") -> None:
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    if bold_prefix:
        r = p.add_run(bold_prefix)
        r.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)
    r2 = p.add_run(text)
    r2.font.color.rgb = RGBColor(0, 0, 0)
    p.paragraph_format.space_after = Pt(6)


def _fetch_series(node_id: str) -> List[MeasurePoint]:
    raw = fetch_json(
        f"{acl_node_base_url()}/nodes/measures/dates",
        params=[
            ("id", node_id),
            ("start", START.strftime("%d%m%Y")),
            ("end", END.strftime("%d%m%Y")),
        ],
    )
    measures = flatten_measures(normalize_measures_payload(raw, node_id))
    # Solo días del periodo (API a veces trae bordes)
    out = []
    for m in measures:
        d = m.date.date() if hasattr(m.date, "date") else m.date
        if START <= d <= END:
            out.append(m)
    out.sort(key=lambda x: x.date)
    return out


def _by_date(measures: List[MeasurePoint]) -> Dict[date, float]:
    return {
        (m.date.date() if hasattr(m.date, "date") else m.date): float(m.total_m3 or 0.0)
        for m in measures
    }


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
    """Semanas lun–dom que intersectan el periodo (recortadas al rango)."""
    # primer lunes <= start
    first_monday = start - timedelta(days=start.weekday())
    windows = []
    cur = first_monday
    n = 1
    while cur <= end:
        week_end = cur + timedelta(days=6)
        d0 = max(cur, start)
        d1 = min(week_end, end)
        if d0 <= d1:
            label = f"Semana {n} ({_fmt_fecha(d0)}–{_fmt_fecha(d1)})"
            windows.append((d0, d1, label))
            n += 1
        cur += timedelta(days=7)
    return windows


def _chart_participacion_barras(
    path: Path,
    labels: List[str],
    pct_principal: List[float],
    pct_sexto: List[float],
    pct_gap: List[float],
    title: str,
) -> Path:
    fig, ax = plt.subplots(figsize=(11, 5.5))
    x = list(range(len(labels)))
    w = 0.55
    b1 = ax.bar(x, pct_principal, w, label="Bomba Principal", color="#0050b3")
    b2 = ax.bar(
        x, pct_sexto, w, bottom=pct_principal, label="Sexto Piso", color="#40a9ff"
    )
    bottom2 = [a + b for a, b in zip(pct_principal, pct_sexto)]
    ax.bar(x, pct_gap, w, bottom=bottom2, label="No monitoreado", color="#c0392b", alpha=0.85)
    ax.axhline(100, color="#333", linestyle="--", linewidth=1, alpha=0.5)
    ax.set_ylabel("% sobre Sanitaria (100%)", fontsize=11, fontweight="bold")
    ax.set_title(title, fontsize=13, fontweight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=35, ha="right", fontsize=8)
    ax.set_ylim(0, 110)
    ax.legend(loc="upper right", fontsize=9)
    ax.grid(True, axis="y", linestyle="--", alpha=0.3)
    # Etiquetas de suma salas en cada barra
    for i, (p1, p2) in enumerate(zip(pct_principal, pct_sexto)):
        s = p1 + p2
        ax.text(i, s + 1.5, f"{s:.0f}%", ha="center", va="bottom", fontsize=8, fontweight="bold")
    plt.tight_layout()
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def _chart_pie_periodo(path: Path, salas: float, gap: float, sanit: float) -> Path:
    fig, ax = plt.subplots(figsize=(6.2, 4.8))
    vals = [max(salas, 0), max(gap, 0)]
    if sum(vals) <= 0:
        vals = [1, 0]
    wedges, *_ = ax.pie(
        vals,
        labels=[
            f"Principal + Sexto\n({format_number_chilean(salas / sanit * 100 if sanit else 0, 1)}%)",
            f"No monitoreado\n({format_number_chilean(gap / sanit * 100 if sanit else 0, 1)}%)",
        ],
        colors=["#0050b3", "#c0392b"],
        autopct=lambda p: f"{p:.1f}%",
        startangle=90,
        textprops={"fontsize": 10},
    )
    for w in wedges:
        w.set_edgecolor("white")
        w.set_linewidth(1.5)
    ax.set_title(
        "Participación del periodo\nsobre Medidor Principal Sanitaria",
        fontsize=12,
        fontweight="bold",
    )
    plt.tight_layout()
    fig.savefig(path, dpi=150, bbox_inches="tight")
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


def main() -> None:
    print("=" * 70)
    print("[INFO] Bupa Antofagasta — análisis piloto")
    print(f"       Periodo: {_fmt_fecha(START)} → {_fmt_fecha(END)}")
    print(f"       Días: {(END - START).days + 1}")
    print("=" * 70)
    t0 = time.perf_counter()

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = Path("reports") / FOLDER / "ANALISIS_PILOTO" / f"PILOTO_{stamp}"
    out_dir.mkdir(parents=True, exist_ok=True)

    series: Dict[str, List[MeasurePoint]] = {}
    by_date: Dict[str, Dict[date, float]] = {}
    for n in NODOS_PRINCIPALES + [NODO_EXTRA]:
        print(f"[INFO] Descargando {n['id']} {n['nombre']}...")
        series[n["id"]] = _fetch_series(n["id"])
        by_date[n["id"]] = _by_date(series[n["id"]])
        tot = sum(by_date[n["id"]].values())
        print(f"       {len(series[n['id']])} días · {tot:.1f} m³")

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
    pct_salas = 100.0 * tot_salas / tot_san if tot_san else 0.0
    pct_gap = 100.0 * tot_gap / tot_san if tot_san else 0.0
    pct_10 = 100.0 * tot_10 / tot_san if tot_san else 0.0

    num_dias = (END - START).days + 1
    # Promedios sobre días con dato (Sexto tiene hueco por lluvias).
    n_san = max(len(by_date[sanit_id]), 1)
    n_07 = max(len(by_date[sal_07]), 1)
    n_08 = max(len(by_date[sal_08]), 1)
    prom_san = tot_san / n_san
    prom_07 = tot_07 / n_07
    prom_08 = tot_08 / n_08

    # --- Documento ---
    doc = Document()
    title = doc.add_paragraph()
    tr = title.add_run(
        "ANÁLISIS PILOTO — CLÍNICA BUPA ANTOFAGASTA\n"
        f"Consumo {_fmt_fecha(START)} al {_fmt_fecha(END)}"
    )
    tr.bold = True
    tr.font.size = Pt(16)
    tr.font.color.rgb = RGBColor(0, 51, 102)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    _add_body(
        doc,
        f"Periodo analizado: {_fmt_fecha(START)} al {_fmt_fecha(END)} "
        f"({num_dias} días civiles). El inicio es el 23/07/2026: desde esa fecha "
        f"la medición ya está correcta; los días previos de julio corresponden a "
        f"la puesta en marcha (caos de habilitación) y se dejan fuera. El cierre "
        f"es el domingo recién pasado, sin la semana en curso. "
        f"El Medidor Principal Sanitaria se trata como el 100% del agua de entrada "
        f"(medición WES del punto sanitaria, no la factura de la sanitaria). "
        f"Sala de Bomba Principal y Sala de Bomba Sexto Piso son las dos "
        f"submediciones “de abajo”. La diferencia respecto de Sanitaria es volumen "
        f"no monitoreado punto a punto (candidato: empalme desde la matriz "
        f"principal antes de la bomba del segundo piso).",
    )

    # Huecos / lluvias
    add_formatted_title(doc, "1. Periodos sin data (lluvias Antofagasta)")
    ll_a, ll_b = LLUVIAS_SIN_DATA
    n_ll = (ll_b - ll_a).days + 1
    _add_body(
        doc,
        f"En el periodo hay un tramo sin registro en Sala de Bomba Sexto Piso: "
        f"del {_fmt_fecha(ll_a)} al {_fmt_fecha(ll_b)} ({n_ll} días). "
        f"Ese hueco se asocia al evento de lluvias en Antofagasta que afectó la "
        f"conectividad / disponibilidad de data de ese punto. "
        f"Sanitaria y Sala de Bomba Principal sí tienen serie en esas fechas; "
        f"al calcular promedios y % se trabaja con los días disponibles "
        f"(no se inventa consumo en el hueco).",
    )

    # Resumen ejecutivo
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
            "TOTAL salas Principal + Sexto",
            format_number_chilean(tot_salas, 1),
            format_number_chilean(tot_salas / num_dias, 1),
            f"{format_number_chilean(pct_salas, 1)}%",
        ),
        (
            "NO MONITOREADO (Sanitaria − salas)",
            format_number_chilean(tot_gap, 1),
            format_number_chilean(tot_gap / num_dias, 1),
            f"{format_number_chilean(pct_gap, 1)}%",
        ),
    ]
    add_table(doc, "Resumen del periodo", table_resumen, wes_style=True)

    _add_body(
        doc,
        f"En el periodo completo, las dos salas de abajo suman el "
        f"{format_number_chilean(pct_salas, 1)}% de Sanitaria. "
        f"Queda un diferencial no monitoreado de "
        f"{format_number_chilean(pct_gap, 1)}% "
        f"({format_number_chilean(tot_gap, 1)} m³). "
        f"Ejemplo de lectura: si Sanitaria fuera 100 m³, Principal aporta "
        f"~{format_number_chilean(pct_07, 0)} m³ y Sexto Piso ~{format_number_chilean(pct_08, 0)} m³; "
        f"el resto (~{format_number_chilean(pct_gap, 0)} m³) no está medido en esos puntos.",
        bold_prefix="Lectura: ",
    )

    pie_path = _chart_pie_periodo(
        out_dir / "chart_participacion_periodo.png",
        tot_salas,
        tot_gap,
        tot_san,
    )
    add_picture_with_pagination(doc, str(pie_path), Inches(4.8), keep_with_next=False)

    # --- Por cada punto: día a día + promedio día semana ---
    add_formatted_title(doc, "3. Consumo día a día y promedio por día de la semana")
    _add_body(
        doc,
        "Para cada uno de los tres puntos se grafica el consumo de todos los días "
        "del periodo y, a continuación, el promedio de cómo se comporta cada día "
        "de la semana (lunes a domingo) sobre las semanas completas del periodo "
        "(misma lógica usada en Parque Arauco).",
    )

    for n in NODOS_PRINCIPALES:
        nid = n["id"]
        measures = series[nid]
        summary = summarize_consumption(measures)
        add_formatted_title(doc, f"3.{NODOS_PRINCIPALES.index(n)+1} {n['nombre']}")
        nota_lluvia = ""
        if n["corto"] == "Sexto Piso":
            nota_lluvia = (
                f" Incluye el hueco sin data del {_fmt_fecha(LLUVIAS_SIN_DATA[0])} "
                f"al {_fmt_fecha(LLUVIAS_SIN_DATA[1])} por lluvias en Antofagasta."
            )
        _add_body(
            doc,
            f"Consumo total: {format_number_chilean(summary['total'], 1)} m³. "
            f"Promedio diario: {format_number_chilean(summary['promedio_diario'], 1)} m³. "
            + (
                f"Día pico: {summary['max'].date.strftime('%d-%m-%y')} "
                f"({format_number_chilean(summary['max'].total_m3, 1)} m³)."
                if summary.get("max")
                else ""
            )
            + nota_lluvia,
        )

        # Día a día
        chart_daily = out_dir / f"chart_diario_{nid.replace('-', '_')}.png"
        start_dt = datetime.combine(START, datetime.min.time())
        end_dt = datetime.combine(END, datetime.max.time())
        built = build_consumption_chart(measures, chart_daily, start_dt, end_dt)
        if built:
            add_formatted_title(doc, f"Consumo diario — {n['corto']}")
            add_picture_with_pagination(doc, str(built), Inches(6.2), keep_with_next=True)
            _add_body(
                doc,
                f"Serie día a día desde {_fmt_fecha(START)} hasta {_fmt_fecha(END)}.",
            )

        # Promedio por día de semana
        weekly = calculate_weekly_averages(measures)
        if weekly:
            chart_w = out_dir / f"chart_promedio_dow_{nid.replace('-', '_')}.png"
            built_w = build_weekly_averages_chart(weekly, chart_w)
            if built_w:
                n_weeks = len(weekly["weeks"])
                add_formatted_title(
                    doc,
                    f"Promedio por día de la semana — {n['corto']} "
                    f"({n_weeks} semanas completas)",
                )
                add_picture_with_pagination(doc, str(built_w), Inches(5.8), keep_with_next=True)
                avg = weekly["averages"]
                orden = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                detalle = ", ".join(
                    f"{d} {format_number_chilean(avg.get(d, 0), 1)} m³" for d in orden
                )
                _add_body(
                    doc,
                    f"Semanas completas consideradas: "
                    f"{weekly['first_week_start'].strftime('%d/%m/%Y')} a "
                    f"{weekly['last_week_end'].strftime('%d/%m/%Y')} "
                    f"({n_weeks} semanas). Promedios: {detalle}.",
                )
                # Tabla compacta
                rows = [("Día", "Promedio (m³)")]
                for d in orden:
                    rows.append((d, format_number_chilean(avg.get(d, 0), 2)))
                add_table(
                    doc,
                    f"Promedio lun–dom — {n['corto']}",
                    rows,
                    wes_style=True,
                )
        else:
            _add_body(
                doc,
                "No hay suficientes semanas completas (lun–dom) para calcular el promedio por día.",
            )

    # --- Participación mensual ---
    add_formatted_title(doc, "4. Participación % vs Sanitaria — por mes")
    _add_body(
        doc,
        "Sanitaria = 100%. Para cada mes del piloto se muestra qué porcentaje "
        "representan Bomba Principal y Sexto Piso, y el remanente no monitoreado.",
    )

    month_wins = _month_windows(START, END)
    table_m = [
        (
            "Mes",
            "Sanitaria m³",
            "Principal m³ (%)",
            "Sexto m³ (%)",
            "Principal + Sexto (%)",
            "No monitoreado (%)",
        )
    ]
    labels_m, p07_m, p08_m, gap_m = [], [], [], []
    for d0, d1, label in month_wins:
        s = _sum_range(by_date[sanit_id], d0, d1)
        a = _sum_range(by_date[sal_07], d0, d1)
        b = _sum_range(by_date[sal_08], d0, d1)
        salas = a + b
        gap = max(0.0, s - salas)
        pa = 100.0 * a / s if s else 0.0
        pb = 100.0 * b / s if s else 0.0
        ps = 100.0 * salas / s if s else 0.0
        pg = 100.0 * gap / s if s else 0.0
        table_m.append(
            (
                label,
                format_number_chilean(s, 1),
                f"{format_number_chilean(a, 1)} ({format_number_chilean(pa, 1)}%)",
                f"{format_number_chilean(b, 1)} ({format_number_chilean(pb, 1)}%)",
                f"{format_number_chilean(ps, 1)}%",
                f"{format_number_chilean(pg, 1)}%",
            )
        )
        # etiqueta corta para gráfico
        short = label.split("(")[0].strip()
        labels_m.append(short)
        p07_m.append(pa)
        p08_m.append(pb)
        gap_m.append(pg)

    add_table(doc, "Participación mensual sobre Sanitaria", table_m, wes_style=True)

    chart_m = _chart_participacion_barras(
        out_dir / "chart_participacion_mensual.png",
        labels_m,
        p07_m,
        p08_m,
        gap_m,
        "Participación mensual (% sobre Sanitaria = 100%)",
    )
    add_picture_with_pagination(doc, str(chart_m), Inches(6.2), keep_with_next=False)

    # --- Participación semanal ---
    add_formatted_title(doc, "5. Participación % vs Sanitaria — por semana")
    _add_body(
        doc,
        "Misma lógica, desagregada por semana (lunes a domingo, recortada al periodo del piloto).",
    )

    week_wins = _week_windows(START, END)
    table_w = [
        (
            "Semana",
            "Sanitaria m³",
            "Principal %",
            "Sexto %",
            "Principal + Sexto %",
            "No mon. %",
        )
    ]
    labels_w, p07_w, p08_w, gap_w = [], [], [], []
    for d0, d1, label in week_wins:
        s = _sum_range(by_date[sanit_id], d0, d1)
        a = _sum_range(by_date[sal_07], d0, d1)
        b = _sum_range(by_date[sal_08], d0, d1)
        salas = a + b
        gap = max(0.0, s - salas)
        pa = 100.0 * a / s if s else 0.0
        pb = 100.0 * b / s if s else 0.0
        ps = 100.0 * salas / s if s else 0.0
        pg = 100.0 * gap / s if s else 0.0
        table_w.append(
            (
                label,
                format_number_chilean(s, 1),
                f"{format_number_chilean(pa, 1)}%",
                f"{format_number_chilean(pb, 1)}%",
                f"{format_number_chilean(ps, 1)}%",
                f"{format_number_chilean(pg, 1)}%",
            )
        )
        labels_w.append(f"{d0.strftime('%d/%m')}")
        p07_w.append(pa)
        p08_w.append(pb)
        gap_w.append(pg)

    add_table(doc, "Participación semanal sobre Sanitaria", table_w, wes_style=True)

    chart_w = _chart_participacion_barras(
        out_dir / "chart_participacion_semanal.png",
        labels_w,
        p07_w,
        p08_w,
        gap_w,
        "Participación semanal (% sobre Sanitaria = 100%)",
    )
    add_picture_with_pagination(doc, str(chart_w), Inches(6.3), keep_with_next=False)

    # --- Sala N°2: explicación clara en dos etapas ---
    add_formatted_title(doc, "6. Qué está pasando en Sala de Bomba N°2")
    by10 = by_date[sal_10]
    pre = [v for d, v in by10.items() if d < SALA2_CAMBIO]
    post = [v for d, v in by10.items() if d >= SALA2_CAMBIO]
    avg_pre = sum(pre) / len(pre) if pre else 0.0
    avg_post = sum(post) / len(post) if post else 0.0
    tot_pre = sum(pre)
    tot_post = sum(post)
    # % vs sanitaria in each phase
    san_pre = sum(v for d, v in by_date[sanit_id].items() if d < SALA2_CAMBIO)
    san_post = sum(v for d, v in by_date[sanit_id].items() if d >= SALA2_CAMBIO)
    pct_pre = 100.0 * tot_pre / san_pre if san_pre else 0.0
    pct_post = 100.0 * tot_post / san_post if san_post else 0.0

    _add_body(
        doc,
        "Sala de Bomba N°2 no entra en el balance de “los dos puntos de abajo” "
        "(Principal + Sexto). La razón es simple: su comportamiento cambia de "
        "golpe a mitad del piloto y, en la segunda etapa, deja de ser coherente "
        "con Sanitaria como 100% del agua que entra.",
    )
    _add_body(
        doc,
        f"Etapa 1 — del {_fmt_fecha(START)} al {_fmt_fecha(SALA2_CAMBIO - timedelta(days=1))}: "
        f"el punto casi no mueve agua. Promedio "
        f"{format_number_chilean(avg_pre, 1)} m³/día "
        f"({format_number_chilean(pct_pre, 1)}% de Sanitaria en ese tramo). "
        f"En la práctica aporta muy poco al consumo interno medido.",
        bold_prefix="En fácil: ",
    )
    _add_body(
        doc,
        f"Etapa 2 — desde el {_fmt_fecha(SALA2_CAMBIO)} hasta {_fmt_fecha(END)}: "
        f"la lectura salta a un promedio de "
        f"{format_number_chilean(avg_post, 1)} m³/día "
        f"({format_number_chilean(pct_post, 1)}% aparente de Sanitaria). "
        f"En varios días e incluso en el acumulado de octubre esa cifra supera "
        f"el 100% de Sanitaria: eso no cuadra si Sanitaria mide toda el agua de "
        f"entrada. Por eso se deja fuera del % de participación hasta revisar "
        f"instalación, factor de pulso o el empalme de ese tramo.",
        bold_prefix="En fácil: ",
    )

    table_s2 = [
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
        (
            "Periodo completo",
            str(len(by10)),
            format_number_chilean(tot_10 / num_dias, 1),
            f"{format_number_chilean(pct_10, 1)}%",
        ),
    ]
    add_table(doc, "Sala de Bomba N°2 — dos etapas", table_s2, wes_style=True)

    measures_10 = series[sal_10]
    chart_10 = out_dir / "chart_diario_sala_n2.png"
    built_10 = build_consumption_chart(
        measures_10,
        chart_10,
        datetime.combine(START, datetime.min.time()),
        datetime.combine(END, datetime.max.time()),
    )
    if built_10:
        add_formatted_title(doc, "Consumo diario — Sala de Bomba N°2")
        add_picture_with_pagination(doc, str(built_10), Inches(6.0), keep_with_next=True)
        _add_body(
            doc,
            f"Se ve el salto a partir del {_fmt_fecha(SALA2_CAMBIO)}: "
            f"antes casi plano (~1 m³/día) y después en un nivel mucho más alto.",
        )

    # --- Conclusiones ---
    add_formatted_title(doc, "7. Conclusiones para cierre")
    _add_body(
        doc,
        f"1) Periodo limpio: se usa desde el 23/07 (fuera el caos de los últimos "
        f"días de puesta en marcha de julio) hasta el {_fmt_fecha(END)}.\n"
        f"2) Lluvias: Sexto Piso sin data del {_fmt_fecha(ll_a)} al {_fmt_fecha(ll_b)}; "
        f"queda identificado y no se rellena.\n"
        f"3) Balance: Sanitaria = 100%. Principal + Sexto cubren el "
        f"{format_number_chilean(pct_salas, 1)}%; queda "
        f"{format_number_chilean(pct_gap, 1)}% no monitoreado "
        f"({format_number_chilean(tot_gap, 1)} m³).\n"
        f"4) Hipótesis de gap: empalme desde la matriz principal antes de la "
        f"impulsión del segundo piso; confirmar en terreno.\n"
        f"5) Sala N°2: dos etapas claras (casi nula hasta el 09/09; salto desde "
        f"el 10/09). No usar en el % hasta validar la medición.",
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
    else:
        print("[ADVERTENCIA] Sin PDF; se subirá solo el Word.")

    # Resumen consola
    print("-" * 70)
    print(f"Sanitaria: {tot_san:.1f} m³ (100%) · prom {prom_san:.1f}/día")
    print(f"Principal: {tot_07:.1f} m³ ({pct_07:.1f}%)")
    print(f"Sexto:     {tot_08:.1f} m³ ({pct_08:.1f}%)")
    print(f"Salas:     {tot_salas:.1f} m³ ({pct_salas:.1f}%)")
    print(f"Gap:       {tot_gap:.1f} m³ ({pct_gap:.1f}%)")
    print(f"N°2:       {tot_10:.1f} m³ ({pct_10:.1f}% aparente) — NO usado en balance")
    print(f"Tiempo: {time.perf_counter() - t0:.1f}s")
    print(f"OUT_DIR={out_dir}")
    if pdf_path:
        print(f"PDF_PATH={pdf_path}")
    print(f"DOCX_PATH={docx_path}")


if __name__ == "__main__":
    main()
