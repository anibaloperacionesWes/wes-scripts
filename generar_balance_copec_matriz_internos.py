#!/usr/bin/env python3
"""Balance anual COPEC: matriz 100% vs puntos internos.

Evalúa ene-2025 a ago-2026 para detectar desde qué mes los internos
dejan de calzar con la Matriz Principal (000009-06).
"""
from __future__ import annotations

import io
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as pdfcanvas

from agregado_extendido_extra import _meses_ultimos_n, _serie_mensual_nodo
from generar_reporte_word import (
    acl_node_base_url,
    fetch_json,
    flatten_measures,
    normalize_measures_payload,
)
from informe_gestion_hidrica_pdf import (
    BODY,
    GRAY,
    NAVY,
    ORANGE,
    _logo_reader,
    _style_axes,
    resolve_logo,
)

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "reports" / "COPEC" / "Balance_Matriz_Internos"
NAVY_HEX = "#003B64"
ORANGE_HEX = "#E67E22"
RED_HEX = "#C0392B"
GREEN_HEX = "#1E8449"

NODOS = [
    ("000009-06", "Matriz Principal"),
    ("000009-00", "Costanera"),
    ("000009-01", "Oficina"),
    ("000009-03", "Lav. automático N"),
    ("000009-04", "Lav. automático S"),
    ("000009-05", "Riego"),
    ("000009-08", "Pronto Baños"),
    ("000009-09", "Lav. autoservicio N"),
    ("000009-10", "Lav. autoservicio S"),
    ("000009-11", "Pronto Tienda"),
    ("000009-02", "Estanque reutilización"),
]
INTERNOS_IDS = [
    "000009-00",
    "000009-01",
    "000009-03",
    "000009-04",
    "000009-05",
    "000009-08",
    "000009-09",
    "000009-10",
    "000009-11",
]
MESES_ES = {
    1: "Ene",
    2: "Feb",
    3: "Mar",
    4: "Abr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Ago",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dic",
}
END = datetime(2026, 8, 31)
N_MESES = 20


def _fmt(n: Optional[float]) -> str:
    if n is None:
        return "—"
    formatted = f"{n:,.0f}"
    return formatted.replace(",", ".")


def fetch_mensual() -> Dict[str, List[Tuple[str, float]]]:
    out: Dict[str, List[Tuple[str, float]]] = {}
    for nid, nombre in NODOS:
        print(f"  mensual {nid} {nombre}…")
        out[nid] = _serie_mensual_nodo(nid, END, N_MESES)
    return out


def fetch_diario(node_id: str, start: datetime, end: datetime):
    payload_raw = fetch_json(
        f"{acl_node_base_url()}/nodes/measures/dates",
        params=[
            ("id", node_id),
            ("start", start.strftime("%d%m%Y")),
            ("end", end.strftime("%d%m%Y")),
        ],
    )
    payload = normalize_measures_payload(payload_raw, node_id)
    return flatten_measures(payload)


def monthly_table(series: Dict[str, List[Tuple[str, float]]]) -> pd.DataFrame:
    meses = _meses_ultimos_n(END, N_MESES)
    rows = []
    for i, (y, m) in enumerate(meses):
        parts = {nid: float(series[nid][i][1] or 0.0) for nid, _ in NODOS}
        matriz = parts["000009-06"]
        internos = sum(parts[nid] for nid in INTERNOS_IDS)
        ratio = (internos / matriz * 100) if matriz > 1 else None
        if y == 2025 and m in (10, 11) and matriz < 1000:
            nota = "Matriz colapsada (~15 m³/día)"
        elif ratio is not None and ratio > 80:
            nota = "Internos superan matriz (no calzan)"
        elif y == 2026 and m >= 6 and parts["000009-00"] < 1:
            if m >= 8 and parts["000009-03"] < 1:
                nota = "Costanera y Lav. auto N en 0"
            else:
                nota = "Costanera en 0 (posible corte)"
        else:
            nota = "Relación coherente (internos < matriz)"
        rows.append(
            {
                "periodo": datetime(y, m, 1),
                "etiqueta": f"{MESES_ES[m]} {y}",
                "matriz": matriz,
                "internos": internos,
                "ratio_pct": ratio,
                "delta": internos - matriz,
                "oficina": parts["000009-01"],
                "riego": parts["000009-05"],
                "costanera": parts["000009-00"],
                "lav_n": parts["000009-03"],
                "lav_s": parts["000009-04"],
                "pronto_b": parts["000009-08"],
                "pronto_t": parts["000009-11"],
                "lav_as_n": parts["000009-09"],
                "lav_as_s": parts["000009-10"],
                "estanque": parts["000009-02"],
                "nota": nota,
            }
        )
    return pd.DataFrame(rows)


def oficina_daily_facts(measures) -> dict:
    facts: dict = {
        "first_spike": None,
        "first_spike_m3": None,
        "last_spike": None,
        "max_day": None,
        "max_m3": None,
        "days_over_100": 0,
        "series": [],
    }
    max_m3 = -1.0
    for mp in measures:
        v = float(mp.total_m3)
        d = mp.date.date()
        facts["series"].append((d, v))
        if v >= 100:
            facts["days_over_100"] += 1
            if facts["first_spike"] is None:
                facts["first_spike"] = d
                facts["first_spike_m3"] = v
            facts["last_spike"] = d
        if v > max_m3:
            max_m3 = v
            facts["max_day"] = d
            facts["max_m3"] = v
    return facts


def fig_to_reader(fig) -> ImageReader:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return ImageReader(buf)


def chart_balance(df: pd.DataFrame) -> ImageReader:
    fig, ax = plt.subplots(figsize=(11.4, 4.0))
    x = list(range(len(df)))
    labels = list(df["etiqueta"])
    ax.bar(
        [i - 0.18 for i in x],
        df["matriz"],
        width=0.36,
        color=ORANGE_HEX,
        label="Matriz Principal (100%)",
        zorder=3,
    )
    ax.bar(
        [i + 0.18 for i in x],
        df["internos"],
        width=0.36,
        color=NAVY_HEX,
        label="Suma internos (potable)",
        zorder=3,
    )
    ax.set_yscale("log")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=50, ha="right", fontsize=8)
    ax.set_ylabel("m³ / mes (escala log)")
    ax.set_title(
        "Matriz vs suma de puntos internos — COPEC (ene-2025 a ago-2026)",
        color=NAVY_HEX,
        fontweight="bold",
        fontsize=11,
    )
    _style_axes(ax)
    ax.legend(loc="upper left", fontsize=8, frameon=False)
    ax.axvline(7.5, color=RED_HEX, ls="--", lw=1.1, zorder=2)
    ax.text(7.65, ax.get_ylim()[1] * 0.45, "Desde Ago-2025\nno calzan", color=RED_HEX, fontsize=8)
    ax.axvline(14.5, color=GREEN_HEX, ls="--", lw=1.1, zorder=2)
    ax.text(14.65, ax.get_ylim()[1] * 0.12, "Abr-2026\nse recupera", color=GREEN_HEX, fontsize=8)
    fig.tight_layout()
    return fig_to_reader(fig)


def chart_oficina_riego(df: pd.DataFrame) -> ImageReader:
    fig, ax = plt.subplots(figsize=(11.4, 3.5))
    x = list(range(len(df)))
    labels = list(df["etiqueta"])
    ax.plot(x, df["oficina"], color="#8E44AD", marker="o", lw=2, label="Oficina 000009-01")
    ax.plot(x, df["riego"], color="#16A085", marker="s", lw=2, label="Riego 000009-05")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=50, ha="right", fontsize=8)
    ax.set_ylabel("m³ / mes")
    ax.set_title("Puntos que inflan el descalce: Oficina y Riego", color=NAVY_HEX, fontweight="bold", fontsize=11)
    _style_axes(ax)
    ax.legend(fontsize=8, frameon=False)
    fig.tight_layout()
    return fig_to_reader(fig)


def chart_oficina_diario(facts: dict) -> Optional[ImageReader]:
    series = facts.get("series") or []
    if not series:
        return None
    # Recortar a jun-2025 → abr-2026 para ver el salto y la recuperación.
    pts = [(d, v) for d, v in series if d >= datetime(2025, 6, 1).date() and d <= datetime(2026, 4, 30).date()]
    if not pts:
        pts = series
    fig, ax = plt.subplots(figsize=(11.4, 3.2))
    ax.plot([d for d, _ in pts], [v for _, v in pts], color="#8E44AD", lw=1.2)
    ax.set_ylabel("m³ / día")
    ax.set_title("Oficina (000009-01) — consumo diario jun-2025 a abr-2026", color=NAVY_HEX, fontweight="bold", fontsize=11)
    _style_axes(ax)
    if facts.get("first_spike"):
        ax.axvline(facts["first_spike"], color=RED_HEX, ls="--", lw=1.0)
        ax.text(
            facts["first_spike"],
            ax.get_ylim()[1] * 0.85,
            f"  Primer salto\n  {facts['first_spike']:%d/%m/%Y}\n  {_fmt(facts['first_spike_m3'])} m³",
            color=RED_HEX,
            fontsize=7.5,
        )
    fig.autofmt_xdate(rotation=30)
    fig.tight_layout()
    return fig_to_reader(fig)


def wrap(c, text: str, x: float, y: float, w: float, font: str, size: float, leading: float, color=BODY) -> float:
    from reportlab.pdfbase.pdfmetrics import stringWidth

    words = text.split()
    line = ""
    c.setFillColor(color)
    c.setFont(font, size)
    for word in words:
        trial = (line + " " + word).strip()
        if stringWidth(trial, font, size) <= w:
            line = trial
        else:
            c.drawString(x, y, line)
            y -= leading
            line = word
    if line:
        c.drawString(x, y, line)
        y -= leading
    return y


def _header(c, pw, ph, title: str, subtitle: str, logo) -> None:
    c.setFillColor(NAVY)
    c.rect(0, ph - 6, pw, 6, fill=1, stroke=0)
    c.setFillColor(GRAY)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(12 * mm, ph - 14 * mm, "WATER EFFICIENCY SERVICES")
    if logo is not None:
        c.drawImage(logo, pw - 12 * mm - 115, ph - 16 * mm, width=115, height=25, mask="auto")
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(12 * mm, ph - 24 * mm, title)
    c.setFillColor(GRAY)
    c.setFont("Helvetica", 9.5)
    c.drawString(12 * mm, ph - 30 * mm, subtitle)


def _footer(c, pw, page: int, total: int) -> None:
    c.setStrokeColor(HexColor("#D7E1E6"))
    c.setLineWidth(1)
    c.line(12 * mm, 10 * mm, pw - 12 * mm, 10 * mm)
    c.setFillColor(GRAY)
    c.setFont("Helvetica", 7)
    c.drawString(12 * mm, 5.5 * mm, "COPEC · Balance Matriz Principal vs puntos internos · Ene 2025 – Ago 2026")
    c.drawRightString(pw - 12 * mm, 5.5 * mm, f"{page} / {total}")


def build_pdf(df: pd.DataFrame, facts: dict, path: Path) -> None:
    img1 = chart_balance(df)
    img2 = chart_oficina_riego(df)
    img3 = chart_oficina_diario(facts)
    logo = _logo_reader(resolve_logo())

    c = pdfcanvas.Canvas(str(path), pagesize=landscape(A4))
    pw, ph = landscape(A4)
    ml, mr, mb = 12 * mm, 12 * mm, 14 * mm

    _header(
        c,
        pw,
        ph,
        "COPEC · Balance matriz vs puntos internos",
        "Enero 2025 – Agosto 2026 · Matriz Principal = 100% del recinto",
        logo,
    )

    y = ph - 38 * mm
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(ml, y, "Conclusión")
    y -= 6 * mm

    first = facts.get("first_spike")
    first_txt = (
        f"El {first.strftime('%d/%m/%Y')} Oficina (000009-01) registra {_fmt(facts.get('first_spike_m3'))} m³ "
        f"en un día —antes estaba en ~0—"
        if first
        else "En agosto 2025 Oficina (000009-01) deja de estar en ~0"
    )
    last = facts.get("last_spike")
    last_txt = f"hasta {last.strftime('%d/%m/%Y')}" if last else "hasta marzo 2026"
    max_txt = ""
    if facts.get("max_day") is not None:
        max_txt = (
            f" El máximo diario de Oficina es {_fmt(facts.get('max_m3'))} m³ "
            f"el {facts['max_day'].strftime('%d/%m/%Y')}."
        )

    conclusion = (
        "Los consumos internos dejan de calzar con la Matriz Principal desde agosto 2025. "
        "Hasta julio 2025 la suma de ramales era el 17–38% de la matriz (coherente: no todos "
        f"los usos están medidos). {first_txt} y desde ahí entrega cientos a miles de m³/día "
        f"{last_txt}. Riego se infla en el mismo tramo.{max_txt} "
        "En octubre–noviembre 2025 la propia matriz cae a cientos de m³/mes (antes ~3.000–4.000), "
        "así que ese bimestre ni siquiera el total del recinto es creíble. En abril 2026 la "
        "relación vuelve al patrón sano (internos ≈ 23–45% de la matriz). Costanera queda en 0 "
        "desde junio 2026 y Lav. automático N en 0 en agosto 2026; el estanque de reutilización "
        "nunca midió."
    )
    y = wrap(c, conclusion, ml, y, pw - ml - mr, "Helvetica", 8.2, 11)

    y -= 2 * mm
    c.setFont("Helvetica-Bold", 10)
    c.setFillColor(NAVY)
    c.drawString(ml, y, "Serie mensual (m³)")
    y -= 5 * mm

    headers = [
        (18, "Mes"),
        (22, "Matriz"),
        (24, "Internos"),
        (16, "% int."),
        (22, "Oficina"),
        (20, "Riego"),
        (22, "Costanera"),
        (52, "Diagnóstico"),
    ]
    row_h = 5.0 * mm
    table_w = sum(h[0] for h in headers) * mm
    x0 = ml

    def draw_row(yy, vals, fill=None, bold=False, fg=NAVY):
        if fill:
            c.setFillColor(fill)
            c.rect(x0, yy - 1.4 * mm, table_w, row_h, fill=1, stroke=0)
        c.setFillColor(fg)
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 6.6 if not bold else 7)
        x = x0
        for wmm, val in zip([h[0] for h in headers], vals):
            c.drawString(x + 1.2 * mm, yy, str(val))
            x += wmm * mm

    draw_row(y, [h[1] for h in headers], fill=NAVY, bold=True, fg=white)
    y -= row_h
    for _, r in df.iterrows():
        ratio = r["ratio_pct"]
        fill = None
        if ratio is not None and ratio > 80:
            fill = HexColor("#F9EBEA")
        elif r["periodo"].month in (10, 11) and r["periodo"].year == 2025:
            fill = HexColor("#FEF5E7")
        elif str(r["nota"]).startswith("Relación coherente"):
            fill = HexColor("#E8F6F3")
        vals = [
            r["etiqueta"],
            _fmt(r["matriz"]),
            _fmt(r["internos"]),
            f"{ratio:.0f}%" if ratio is not None else "—",
            _fmt(r["oficina"]),
            _fmt(r["riego"]),
            _fmt(r["costanera"]),
            r["nota"][:44],
        ]
        draw_row(y, vals, fill=fill)
        y -= row_h

    chart_h = max(38 * mm, y - mb - 6 * mm)
    c.drawImage(img1, ml, mb + 2 * mm, width=pw - ml - mr, height=chart_h, preserveAspectRatio=True, mask="auto")
    _footer(c, pw, 1, 2)
    c.showPage()

    _header(
        c,
        pw,
        ph,
        "COPEC · Detalle del descalce",
        "Oficina y Riego · implicancia para la gestión hídrica",
        logo,
    )
    y = ph - 38 * mm
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(ml, y, "Qué pasó, mes a mes")
    y -= 6 * mm
    bullets = [
        "Ene–jul 2025: matriz 3.300–3.900 m³/mes. Internos 17–38%. Oficina ≈ 0. Relación usable.",
        "Ago 2025: Oficina salta a miles de m³. Internos superan el 200% de la matriz. Aquí empieza el descalce.",
        "Sep 2025: Oficina y Riego se inflan. Internos varias veces la matriz.",
        "Oct–nov 2025: la matriz misma se cae (cientos de m³/mes). Oficina sigue en miles. Doble falla: total del recinto e internos.",
        "Dic 2025–mar 2026: matriz vuelve a ~4.000 m³, pero Oficina y Riego siguen orders of magnitude por encima. Peso de pulso / calibración, no uso real.",
        "Abr 2026: Oficina vuelve a valores residuales. Internos ≈ 40% de la matriz. Se recupera el calce relativo.",
        "Jun–ago 2026: Costanera en 0 todos los días. Ago: Lav. automático N también en 0. Estanque reutilización sigue en 0 todo el año.",
    ]
    # Completar bullets con cifras reales del df
    def row_of(y_, m_):
        hit = df[(df["periodo"].dt.year == y_) & (df["periodo"].dt.month == m_)]
        return None if hit.empty else hit.iloc[0]

    ago = row_of(2025, 8)
    sep = row_of(2025, 9)
    oct_ = row_of(2025, 10)
    nov = row_of(2025, 11)
    abr = row_of(2026, 4)
    bullets = [
        "Ene–jul 2025: matriz ~3.300–3.900 m³/mes. Internos 17–38%. Oficina ≈ 0. Relación usable.",
        (
            f"Ago 2025: matriz {_fmt(ago['matriz'])} m³, internos {_fmt(ago['internos'])} "
            f"({ago['ratio_pct']:.0f}%). Oficina {_fmt(ago['oficina'])} m³. Aquí empieza el descalce."
            if ago is not None
            else bullets[1]
        ),
        (
            f"Sep 2025: matriz {_fmt(sep['matriz'])} m³, internos {_fmt(sep['internos'])} "
            f"({sep['ratio_pct']:.0f}%). Oficina {_fmt(sep['oficina'])} + Riego {_fmt(sep['riego'])}."
            if sep is not None
            else bullets[2]
        ),
        (
            f"Oct–nov 2025: matriz {_fmt(oct_['matriz'])} y {_fmt(nov['matriz'])} m³. "
            f"Oficina {_fmt(oct_['oficina'])} / {_fmt(nov['oficina'])}. Doble falla: total del recinto e internos."
            if oct_ is not None and nov is not None
            else bullets[3]
        ),
        "Dic 2025–mar 2026: matriz ~3.800–4.400 m³, pero Oficina llega a 37–103 mil m³/mes y Riego a 4.600–20.500. Peso de pulso / calibración, no uso real.",
        (
            f"Abr 2026: Oficina {_fmt(abr['oficina'])} m³. Internos = {abr['ratio_pct']:.0f}% de la matriz. Se recupera el calce relativo."
            if abr is not None
            else bullets[5]
        ),
        "Jun–ago 2026: Costanera en 0 todos los días. Ago: Lav. automático N también en 0. Estanque reutilización sigue en 0 todo el año.",
    ]
    for b in bullets:
        y = wrap(c, "• " + b, ml, y, pw - ml - mr, "Helvetica", 8.0, 11)
        y -= 1.0 * mm

    y -= 2 * mm
    c.setFont("Helvetica-Bold", 11)
    c.setFillColor(NAVY)
    c.drawString(ml, y, "Implicancia para reportes")
    y -= 6 * mm
    impl = (
        "La Matriz Principal sigue siendo el 100% del recinto y el único total potable creíble "
        "cuando está sana (ene–sep 2025 y dic 2025 en adelante, con la salvedad oct–nov 2025). "
        "No se deben sumar los ramales para armar el consumo del sitio: entre ago-2025 y mar-2026 "
        "esa suma es varias veces la matriz. Oficina y Riego en ese tramo hay que marcarlos como "
        "dato no usable. Costanera desde jun-2026 y Lav. automático N en ago-2026 están en silencio "
        "de sonda, no en consumo cero real. El estanque de reutilización no aporta serie."
    )
    y = wrap(c, impl, ml, y, pw - ml - mr, "Helvetica", 8.2, 11)

    y -= 3 * mm
    remaining = y - mb - 4 * mm
    if img3 is not None:
        h2 = remaining * 0.48
        h3 = remaining * 0.48
        c.drawImage(img2, ml, y - h2, width=pw - ml - mr, height=h2, preserveAspectRatio=True, mask="auto")
        c.drawImage(img3, ml, mb + 2 * mm, width=pw - ml - mr, height=h3, preserveAspectRatio=True, mask="auto")
    else:
        c.drawImage(img2, ml, mb + 8 * mm, width=pw - ml - mr, height=max(40 * mm, remaining), preserveAspectRatio=True, mask="auto")
    _footer(c, pw, 2, 2)
    c.save()


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    print("Descargando series mensuales COPEC (20 meses)…")
    series = fetch_mensual()
    df = monthly_table(series)

    print("Descargando diario Oficina (ene-2025 a ago-2026)…")
    oficina_daily = fetch_diario("000009-01", datetime(2025, 1, 1), END)
    facts = oficina_daily_facts(oficina_daily)
    print(
        "  primer salto:",
        facts.get("first_spike"),
        facts.get("first_spike_m3"),
        "| último >100:",
        facts.get("last_spike"),
        "| max:",
        facts.get("max_day"),
        facts.get("max_m3"),
        "| días >100:",
        facts.get("days_over_100"),
    )

    xlsx = OUT_DIR / "COPEC_Balance_Matriz_vs_Internos_2025-2026.xlsx"
    pdf = OUT_DIR / "COPEC_Balance_Matriz_vs_Internos_2025-2026.pdf"

    export = df.copy()
    export["periodo"] = export["periodo"].dt.strftime("%Y-%m")
    export.to_excel(xlsx, index=False)
    print("Excel:", xlsx)

    build_pdf(df, facts, pdf)
    print("PDF:", pdf)

    print("\n=== Resumen ratio internos/matriz ===")
    for _, r in df.iterrows():
        ratio = r["ratio_pct"]
        rs = f"{ratio:7.0f}%" if ratio is not None else "      —"
        print(f"  {r['etiqueta']:12} matriz={r['matriz']:8.0f}  int={r['internos']:9.0f}  {rs}  {r['nota']}")


if __name__ == "__main__":
    main()
