"""
Bupa / UPA Antofagasta — reporte agregado formato clásico extendido
(mismo estilo Club Providencia / Fundo Zapallar de fin de mes).

Periodo por defecto: 23/07/2026 – 11/08/2026 (días civiles completos).

Uso:
  python generar_agregado_bupa_antofagasta_clasico.py
"""

from __future__ import annotations

import sys
import time
from datetime import date, datetime
from pathlib import Path
from typing import List, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import requests
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches

if sys.platform == "win32":
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", line_buffering=True)
        except Exception:
            pass

from generar_reporte_word import (
    add_formatted_heading,
    add_formatted_title,
    add_picture_with_pagination,
    add_table,
    build_monthly_comparison_chart,
    calculate_nocturnal_metrics,
    estilizar_tabla_wes,
    format_currency_chilean,
    format_number_chilean,
    generate_aggregated_report,
    get_water_price_per_m3,
    parse_date,
    acl_node_base_url,
    fetch_json,
    flatten_measures,
    normalize_measures_payload,
    summarize_consumption,
)
from generar_consolidado_m3_mensual_puente_alto import consumo_mes_un_nodo

COMPANY_ID = "000029"
# Puntos UPA Antofagasta (clínica)
NODE_IDS = [
    "000029-07",  # Sala de Bomba Principal
    "000029-08",  # Sala de Bomba Sexto Piso
    "000029-09",  # Medidor Principal Sanitaria (cuenta)
    "000029-10",  # Sala de Bomba N°2
]
# Para el comparativo de 6 meses usamos la cuenta (evita doble conteo con salas)
CUENTA_ID = "000029-09"
START = "23/07/2026"
END = "11/08/2026"
FOLDER = "Bupa_Antofagasta"
COLOR_BARRA = "#0050b3"
PRECIO_DEFAULT = 2768.65  # tarifa ref. factura julio

# Hito: bajada de consumo por mejora del equipo de mantención
MEJORA_DATE = date(2026, 7, 29)
BEFORE_START = "23/07/2026"
BEFORE_END = "28/07/2026"  # ritmo previo a la mejora
AFTER_START = "29/07/2026"
AFTER_END = "11/08/2026"  # ritmo post-mejora (días completos)
AUGUST_YEAR = 2026
AUGUST_DAYS = 31
AUGUST_OBS_END = "11/08/2026"  # último día observado de agosto


def _meses_ultimos_6(end_d: date) -> List[Tuple[int, int]]:
    y, m = end_d.year, end_d.month
    out: List[Tuple[int, int]] = []
    for _ in range(6):
        out.append((y, m))
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    out.reverse()
    return out


def _plot_6_meses(series: List[Tuple[str, float]], out: Path) -> Path:
    labels = [s[0] for s in series]
    vals = [s[1] for s in series]
    fig, ax = plt.subplots(figsize=(10, 5.5))
    bars = ax.bar(labels, vals, color=COLOR_BARRA, width=0.65)
    ax.set_ylabel("Consumo mensual cuenta (m³)", fontsize=11, fontweight="bold")
    ax.set_xlabel("Mes", fontsize=11, fontweight="bold")
    ax.set_title(
        "Comparativo últimos 6 meses — Medidor Principal Sanitaria",
        fontsize=12,
        fontweight="bold",
    )
    ax.grid(axis="y", linestyle="--", alpha=0.35)
    for bar, v in zip(bars, vals):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height(),
            f"{format_number_chilean(v, 0)}",
            ha="center",
            va="bottom",
            fontsize=9,
            fontweight="bold",
        )
    fig.tight_layout()
    fig.savefig(out, dpi=160, bbox_inches="tight")
    plt.close(fig)
    return out


def _totales_periodo(start_dt: datetime, end_dt: datetime) -> Tuple[float, float, float]:
    """Retorna (total_m3 todos los puntos, nocturno_m3, price)."""
    total = 0.0
    noct = 0.0
    prices: List[float] = []
    for nid in NODE_IDS:
        raw = fetch_json(
            f"{acl_node_base_url()}/nodes/measures/dates",
            params=[
                ("id", nid),
                ("start", start_dt.strftime("%d%m%Y")),
                ("end", end_dt.strftime("%d%m%Y")),
            ],
        )
        payload = normalize_measures_payload(raw, nid)
        measures = flatten_measures(payload)
        summary = summarize_consumption(measures)
        total += float(summary["total"])
        nm = calculate_nocturnal_metrics(nid, start_dt, end_dt)
        noct += float(nm.get("consumo_nocturno_total") or 0.0)
        prices.append(get_water_price_per_m3(COMPANY_ID, nid, payload) or PRECIO_DEFAULT)
    price = sum(prices) / len(prices) if prices else PRECIO_DEFAULT
    return total, noct, price


def _consumo_cuenta_periodo(start_dt: datetime, end_dt: datetime) -> float:
    raw = fetch_json(
        f"{acl_node_base_url()}/nodes/measures/dates",
        params=[
            ("id", CUENTA_ID),
            ("start", start_dt.strftime("%d%m%Y")),
            ("end", end_dt.strftime("%d%m%Y")),
        ],
    )
    payload = normalize_measures_payload(raw, CUENTA_ID)
    return float(summarize_consumption(flatten_measures(payload))["total"])


def _promedio_diario_cuenta(start_s: str, end_s: str) -> Tuple[float, float, int]:
    """Retorna (total_m3, promedio_diario, num_dias)."""
    start_dt = parse_date(start_s)
    end_dt = parse_date(end_s, end_of_day=True)
    total = _consumo_cuenta_periodo(start_dt, end_dt)
    n = (end_dt.date() - start_dt.date()).days + 1
    return total, (total / n if n else 0.0), n


def _plot_proyeccion_agosto(
    path: Path,
    sin_mejora: float,
    con_mejora: float,
    observado: float,
) -> Path:
    labels = [
        "Agosto si se\nmantenía el ritmo\nprevio",
        "Agosto proyectado\ncon la mejora\n(cierre de mes)",
        "Ya consumido\nen agosto\n(01–11)",
    ]
    vals = [sin_mejora, con_mejora, observado]
    colors = ["#c0392b", "#2ecc71", "#0050b3"]
    fig, ax = plt.subplots(figsize=(8.5, 5.2))
    bars = ax.bar(labels, vals, color=colors, width=0.62, edgecolor="#333333", linewidth=0.8)
    ax.set_ylabel("m³ en agosto", fontsize=11, fontweight="bold")
    ax.set_title(
        "Proyección de cierre de agosto — Medidor Principal Sanitaria",
        fontsize=12,
        fontweight="bold",
    )
    ax.set_ylim(bottom=0)
    ax.grid(axis="y", linestyle="--", alpha=0.35)
    for bar, v in zip(bars, vals):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height(),
            f"{format_number_chilean(v, 0)} m³",
            ha="center",
            va="bottom",
            fontsize=10,
            fontweight="bold",
        )
    fig.tight_layout()
    fig.savefig(path, dpi=160, bbox_inches="tight")
    plt.close(fig)
    return path


def _append_proyeccion_agosto_y_mejora(doc: Document, out_dir: Path, price: float) -> None:
    """Proyección de cierre de agosto + narrativa WES / mantención."""
    before_tot, before_avg, before_n = _promedio_diario_cuenta(BEFORE_START, BEFORE_END)
    after_tot, after_avg, after_n = _promedio_diario_cuenta(AFTER_START, AFTER_END)
    ago_obs = _consumo_cuenta_periodo(
        parse_date("01/08/2026"),
        parse_date(AUGUST_OBS_END, end_of_day=True),
    )
    dias_obs_ago = (parse_date(AUGUST_OBS_END).date() - date(AUGUST_YEAR, 8, 1)).days + 1
    dias_restantes = AUGUST_DAYS - dias_obs_ago
    proy_cierre = ago_obs + after_avg * dias_restantes
    proy_sin_mejora = before_avg * AUGUST_DAYS
    ahorro_m3 = max(0.0, proy_sin_mejora - proy_cierre)
    ahorro_clp = ahorro_m3 * price
    pct_baja = (1.0 - after_avg / before_avg) * 100.0 if before_avg > 0 else 0.0

    chart = _plot_proyeccion_agosto(
        out_dir / "chart_proyeccion_cierre_agosto.png",
        proy_sin_mejora,
        proy_cierre,
        ago_obs,
    )

    add_formatted_heading(doc, "Proyección de cierre de agosto y efecto de la mejora", level=1)

    hito = doc.add_paragraph()
    hito.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    r0 = hito.add_run("Hito 29/07/2026: ")
    r0.bold = True
    hito.add_run(
        "en los gráficos diarios (punto verde) se marca la fecha en que el consumo de la "
        "cuenta bajó de forma nítida. Ese día el Medidor Principal Sanitaria pasó de un "
        f"ritmo cercano a {format_number_chilean(before_avg, 1)} m³/día "
        f"(promedio {BEFORE_START}–{BEFORE_END}) a unos "
        f"{format_number_chilean(after_avg, 1)} m³/día "
        f"(promedio {AFTER_START}–{AFTER_END}), una reducción del orden del "
        f"{format_number_chilean(pct_baja, 0)} %. Esta bajada se asocia a una mejora "
        "ejecutada por el equipo de mantención, a partir de la información de consumo "
        "que entrega WES."
    )

    add_formatted_title(doc, "Proyección mensual — cierre de agosto 2026")
    intro = doc.add_paragraph(
        "Para cerrar la cuenta de agosto se usa el consumo ya registrado "
        f"(01/08–{AUGUST_OBS_END}) y se proyectan los "
        f"{dias_restantes} días restantes al ritmo posterior a la mejora. "
        "En paralelo se muestra qué habría sido el mes si se hubiera mantenido "
        "el consumo previo a la intervención."
    )
    intro.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    rows = [
        (
            "Concepto",
            "m³ / día",
            "Base de cálculo",
            "Total agosto (m³)",
            "Valorizado (CLP)",
        ),
        (
            "Ritmo antes de la mejora",
            format_number_chilean(before_avg, 1),
            f"{BEFORE_START}–{BEFORE_END} ({before_n} días)",
            format_number_chilean(proy_sin_mejora, 0),
            format_currency_chilean(proy_sin_mejora * price),
        ),
        (
            "Ritmo después de la mejora",
            format_number_chilean(after_avg, 1),
            f"{AFTER_START}–{AFTER_END} ({after_n} días)",
            "—",
            "—",
        ),
        (
            "Ya consumido en agosto",
            format_number_chilean(ago_obs / dias_obs_ago, 1),
            f"01/08–{AUGUST_OBS_END} ({dias_obs_ago} días)",
            format_number_chilean(ago_obs, 0),
            format_currency_chilean(ago_obs * price),
        ),
        (
            "Proyección cierre agosto (con mejora)",
            format_number_chilean(proy_cierre / AUGUST_DAYS, 1),
            f"Observado + {dias_restantes} días × ritmo post-mejora",
            format_number_chilean(proy_cierre, 0),
            format_currency_chilean(proy_cierre * price),
        ),
        (
            "Ahorro proyectado vs ritmo previo",
            "—",
            "Diferencia de cierre de mes",
            format_number_chilean(ahorro_m3, 0),
            format_currency_chilean(ahorro_clp),
        ),
    ]
    add_table(doc, "Proyección / ahorro agosto", rows, wes_style=True)

    lectura = doc.add_paragraph()
    lectura.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    r1 = lectura.add_run("Lectura para el cliente: ")
    r1.bold = True
    lectura.add_run(
        f"si el consumo se hubiera mantenido en ~{format_number_chilean(before_avg, 0)} m³/día, "
        f"agosto habría cerrado cerca de {format_number_chilean(proy_sin_mejora, 0)} m³ "
        f"({format_currency_chilean(proy_sin_mejora * price)}). Con el ritmo actual post-mejora, "
        f"el cierre proyectado es de ~{format_number_chilean(proy_cierre, 0)} m³ "
        f"({format_currency_chilean(proy_cierre * price)}), es decir un ahorro del orden de "
        f"{format_number_chilean(ahorro_m3, 0)} m³ "
        f"({format_currency_chilean(ahorro_clp)}) solo en este mes."
    )

    doc.add_paragraph("")
    add_picture_with_pagination(doc, str(chart), Inches(5.8), keep_with_next=True)

    add_formatted_title(doc, "Rol de WES y del equipo de mantención")
    wes = doc.add_paragraph()
    wes.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    wes.add_run(
        "La información que entrega WES es clave para el equipo de mantención: permite ver "
        "en qué momento el consumo se disparó o se mantuvo alto, actuar con precisión y "
        "verificar de inmediato el efecto de las mejoras. En este caso, la bajada marcada "
        f"el {MEJORA_DATE.strftime('%d/%m/%Y')} muestra que, cuando el equipo interviene "
        "apoyado en los datos de monitoreo, la clínica puede reducir de forma sustancial "
        "el volumen facturado y, con ello, la cuenta de agua. Mantener este seguimiento "
        "permite sostener el ahorro y detectar a tiempo cualquier nuevo desvío."
    )

    print(
        f"[OK] Proyección agosto: sin mejora {proy_sin_mejora:.0f} m³ | "
        f"cierre con mejora {proy_cierre:.0f} m³ | ahorro {ahorro_m3:.0f} m³ "
        f"({format_currency_chilean(ahorro_clp)})"
    )


def _append_comparativos(docx_path: Path) -> Path:
    start_dt = parse_date(START)
    end_dt = parse_date(END, end_of_day=True)
    out_dir = docx_path.parent

    print("[INFO] Calculando totales periodo para comparativos...", flush=True)
    total_m3, noct_total, price = _totales_periodo(start_dt, end_dt)
    num_dias = (end_dt.date() - start_dt.date()).days + 1
    factor_30 = 30.0 / max(num_dias, 1)
    leak_monthly = noct_total * factor_30
    efectivo_monthly = max(0.0, (total_m3 - noct_total) * factor_30)

    chart_mes = out_dir / "chart_comparacion_mensual_extra.png"
    built_mes = build_monthly_comparison_chart(
        leak_monthly, efectivo_monthly, price, chart_mes
    )

    print("[INFO] Descargando últimos 6 meses (cuenta Sanitaria)...", flush=True)
    sess = requests.Session()
    cuenta_periodo = _consumo_cuenta_periodo(start_dt, end_dt)
    series_6: List[Tuple[str, float]] = []
    for y, m in _meses_ultimos_6(end_dt.date()):
        # Primer y último día civil del mes
        if m == 12:
            last_day = date(y, 12, 31)
        else:
            last_day = date(y, m + 1, 1).fromordinal(date(y, m + 1, 1).toordinal() - 1)
        first_day = date(y, m, 1)
        # Tramo del periodo que cae dentro de este mes
        seg_start = max(start_dt.date(), first_day)
        seg_end = min(end_dt.date(), last_day)
        overlaps_periodo = seg_start <= seg_end and (
            (y == start_dt.year and m == start_dt.month)
            or (y == end_dt.year and m == end_dt.month)
        )
        if overlaps_periodo and (seg_start > first_day or seg_end < last_day):
            m3 = _consumo_cuenta_periodo(
                parse_date(seg_start.strftime("%d/%m/%Y")),
                parse_date(seg_end.strftime("%d/%m/%Y"), end_of_day=True),
            )
            label = f"{y}-{m:02d}*"
        else:
            m3, _, _ = consumo_mes_un_nodo(sess, CUENTA_ID, y, m)
            label = f"{y}-{m:02d}"
        series_6.append((label, float(m3)))
        print(f"  {label}: {m3:.1f} m³", flush=True)
    chart_6m = _plot_6_meses(series_6, out_dir / "chart_ultimos_6_meses.png")

    doc = Document(str(docx_path))
    add_formatted_heading(doc, "Comparativo del mes (nocturno vs efectivo)", level=1)
    p5 = doc.add_paragraph(
        f"Proyección a 30 días a partir del periodo ({num_dias} días): "
        f"nocturno proyectado {format_number_chilean(leak_monthly, 1)} m³ "
        f"({format_currency_chilean(leak_monthly * price)}) y consumo efectivo "
        f"{format_number_chilean(efectivo_monthly, 1)} m³ "
        f"({format_currency_chilean(efectivo_monthly * price)})."
    )
    p5.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    if built_mes and Path(built_mes).exists():
        add_picture_with_pagination(doc, str(chart_mes), Inches(4.5), keep_with_next=True)

    add_formatted_heading(doc, "Comparativo últimos 6 meses", level=1)
    p6 = doc.add_paragraph(
        "Consumo mensual del Medidor Principal Sanitaria (cuenta de agua de la clínica). "
        "Los meses marcados con * corresponden a tramos parciales del periodo de este reporte "
        f"(desde {start_dt.strftime('%d/%m/%Y')} hasta {end_dt.strftime('%d/%m/%Y')}). "
        f"En el periodo, la cuenta registró {format_number_chilean(cuenta_periodo, 1)} m³."
    )
    p6.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    add_picture_with_pagination(doc, str(chart_6m), Inches(6), keep_with_next=True)

    table = doc.add_table(rows=1 + len(series_6), cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Mes"
    table.rows[0].cells[1].text = "Cuenta Sanitaria (m³)"
    for i, (lab, v) in enumerate(series_6):
        table.rows[i + 1].cells[0].text = lab
        table.rows[i + 1].cells[1].text = format_number_chilean(v, 1)
    estilizar_tabla_wes(table, has_total_row=False)

    # Proyección cierre agosto + narrativa mejora mantención / WES
    # Valorización con tarifa de factura julio (ref. cuenta), no el default API.
    _append_proyeccion_agosto_y_mejora(doc, out_dir, PRECIO_DEFAULT)

    doc.save(str(docx_path))
    print(f"[OK] Comparativos agregados a {docx_path}")
    return docx_path


def main() -> int:
    print("=" * 70)
    print("BUPA / UPA ANTOFAGASTA — agregado clásico extendido + comparativos")
    print(f"Periodo: {START} → {END}")
    print(f"Nodos: {', '.join(NODE_IDS)}")
    print("=" * 70)
    t0 = time.perf_counter()
    out = generate_aggregated_report(
        company_id=COMPANY_ID,
        node_ids=list(NODE_IDS),
        start_date=START,
        end_date=END,
        output_dir="reports",
        apply_exclusions=False,
        generate_ppt=False,
        parallel_node_fetch=True,
        max_parallel_workers=4,
        fuente_agua_id=None,
        company_folder_override=FOLDER,
    )
    docx = Path(out)
    print(f"[OK] Base: {docx}")
    _append_comparativos(docx)
    print(f"[INFO] Tiempo total: {time.perf_counter() - t0:.1f} s")
    print(f"[INFO] Word: {docx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
