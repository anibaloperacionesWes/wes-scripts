# -*- coding: utf-8 -*-
"""
Recorrido ejecutivo Parque Arauco — PPT aparte de las fichas.

Misma estética (navy / gold / cards / fondo PA). Se arma mall por mall.
Esta versión: solo MAE (equipos + consumo mensualizado).

  1) Equipos instalados — mismos 4 puntos del deck (sin 000025-02)
  2) Consumo mensualizado — gráfico mayo → fecha (agosto a la fecha vs proyección)
     Texto: peso de cada equipo en julio (último mes cerrado)

Uso:
  python3 generar_ppt_recorrido_ejecutivo_pa.py
  python3 generar_ppt_recorrido_ejecutivo_pa.py --skip-refresh
"""

from __future__ import annotations

import json
import sys
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from pptx import Presentation
from pptx.dml.color import RGBColor as PptRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches as PptInches, Pt as PptPt

from generar_reporte_word import format_number_chilean

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "reports" / "Parque_Arauco" / "TMP_7MALLS" / "entrega_diego_anibal"
CHARTS = OUT_DIR / "charts_recorrido_mae"
JSON_DATOS = OUT_DIR / "datos_mae_may_ago.json"
LOGO = ROOT / "logo wes.bmp"
FONDO = ROOT / "Parque arauco fondo.jpg"

# MAE: 4 puntos del deck. 000025-02 no entra en este recorrido.
MAE_NODOS = ["000025-01", "000025-04", "000025-07", "000025-19"]
NOMBRE_LARGO = {
    "000025-01": "Estanque Norte Locales Mall",
    "000025-04": "Baños Públicos",
    "000025-07": "Pizza Hut",
    "000025-19": "Sala de Bomba Estanque Sur",
}
NOMBRE_CORTO = {
    "000025-01": "Estanque Norte",
    "000025-04": "Baños Públicos",
    "000025-07": "Pizza Hut",
    "000025-19": "Estanque Sur",
}
COLOR_NODO = {
    "000025-07": (196, 92, 38),
    "000025-01": (13, 59, 102),
    "000025-19": (31, 119, 180),
    "000025-04": (123, 163, 201),
}

DESDE = date(2026, 5, 1)
HASTA = date(2026, 8, 17)
AGO_MES = 31
PERIODO = f"{DESDE.strftime('%d/%m/%Y')} – {HASTA.strftime('%d/%m/%Y')}"
FECHA_EMISION = "18 agosto 2026"
AGO_ETQ = f"1–{HASTA.day}"

NAVY = (13, 59, 102)
GOLD = (201, 162, 39)
TEAL = (31, 119, 180)
GRAY = (90, 90, 90)
LIGHT = (245, 247, 250)
WHITE = (255, 255, 255)
GOLD_SOFT = (232, 213, 163)


def fn(v: float, dec: int = 1) -> str:
    return format_number_chilean(float(v), dec)


def _rgb(t: Tuple[int, int, int]) -> PptRGB:
    return PptRGB(*t)


def _hex(t: Tuple[int, int, int]) -> str:
    return f"#{t[0]:02X}{t[1]:02X}{t[2]:02X}"


def refrescar_datos() -> None:
    from generar_reportes_y_ppt_mall_maipu import guardar_datos_json, obtener_datos_agregados

    print(f"[INFO] Descargando MAE {PERIODO}…", flush=True)
    datos = obtener_datos_agregados(
        MAE_NODOS,
        DESDE.strftime("%d/%m/%Y"),
        HASTA.strftime("%d/%m/%Y"),
    )
    datos["all_measures"] = []
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    guardar_datos_json(datos, JSON_DATOS)


def cargar_mae() -> Tuple[Dict[str, str], Dict[str, Dict[str, Any]], Dict[str, float]]:
    raw = json.loads(JSON_DATOS.read_text(encoding="utf-8"))
    names: Dict[str, str] = {}
    by: Dict[str, Dict[str, Any]] = {}
    for ns in raw["nodes_summary"]:
        nid = ns["node_id"]
        names[nid] = ns["node_name"]
        row: Dict[str, Any] = {
            "may": 0.0,
            "jun": 0.0,
            "jul": 0.0,
            "ago": 0.0,
            "daily": {},
        }
        for m in ns["measures"]:
            d = str(m["date"])[:10]
            v = float(m["total_m3"])
            row["daily"][d] = row["daily"].get(d, 0.0) + v
            month = d[5:7]
            if month == "05":
                row["may"] += v
            elif month == "06":
                row["jun"] += v
            elif month == "07":
                row["jul"] += v
            elif month == "08":
                row["ago"] += v
        by[nid] = row
    tot = {"may": 0.0, "jun": 0.0, "jul": 0.0, "ago": 0.0}
    for nid in MAE_NODOS:
        for k in tot:
            tot[k] += float((by.get(nid) or {}).get(k) or 0)
    tot["ago_d"] = tot["ago"] / HASTA.day if HASTA.day else 0.0
    tot["ago_proy"] = tot["ago_d"] * AGO_MES
    tot["jul_d"] = tot["jul"] / 31.0
    tot["may_d"] = tot["may"] / 31.0
    tot["jun_d"] = tot["jun"] / 30.0
    return names, by, tot


def chart_mensual_mae(path: Path, tot: Dict[str, float]) -> None:
    """Mayo–julio cerrados; agosto apilado (a la fecha + proyección al 31)."""
    fig, ax = plt.subplots(figsize=(9.4, 4.55), dpi=160)
    x = np.arange(4)
    w = 0.62
    may, jun, jul = tot["may"], tot["jun"], tot["jul"]
    ago, proy = tot["ago"], tot["ago_proy"]
    resto = max(proy - ago, 0.0)

    ax.bar(
        x[:3],
        [may, jun, jul],
        width=w,
        color=_hex(NAVY),
        zorder=3,
        label="Mes cerrado",
    )
    ax.bar(
        [3],
        [ago],
        width=w,
        color=_hex(GOLD),
        zorder=3,
        label=f"Agosto {AGO_ETQ} (a la fecha)",
    )
    ax.bar(
        [3],
        [resto],
        width=w,
        bottom=[ago],
        color=_hex(GOLD_SOFT),
        edgecolor=_hex(GOLD),
        linewidth=0.8,
        zorder=3,
        label="Proyección resto de agosto",
    )

    ax.set_xticks(x)
    ax.set_xticklabels(
        ["Mayo", "Junio", "Julio", f"Agosto\n({AGO_ETQ} + proy.)"],
        fontsize=11,
        color=_hex(NAVY),
        fontweight="bold",
    )
    ax.set_ylabel("m³", fontsize=11, color=_hex(NAVY))
    ax.tick_params(axis="y", labelsize=10, colors=_hex(NAVY))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#C5CDD6")
    ax.spines["bottom"].set_color("#C5CDD6")
    ax.yaxis.grid(True, linestyle=":", alpha=0.55, zorder=0)
    ax.set_axisbelow(True)

    ymax = max(may, jun, jul, proy) * 1.18 if max(may, jun, jul, proy) > 0 else 1
    ax.set_ylim(0, ymax)

    for i, v in enumerate([may, jun, jul]):
        ax.text(
            i,
            v + ymax * 0.02,
            fn(v, 0),
            ha="center",
            va="bottom",
            fontsize=11,
            color=_hex(NAVY),
            fontweight="bold",
        )
    # Etiqueta del tramo real de agosto (dentro de la barra dorada)
    ax.text(
        3,
        ago * 0.50,
        fn(ago, 0),
        ha="center",
        va="center",
        fontsize=11,
        color="white",
        fontweight="bold",
    )
    ax.text(
        3,
        proy + ymax * 0.02,
        f"proy. {fn(proy, 0)}",
        ha="center",
        va="bottom",
        fontsize=11,
        color=_hex(NAVY),
        fontweight="bold",
    )

    leg = ax.legend(
        loc="upper right",
        frameon=False,
        fontsize=9,
        labelcolor=_hex(NAVY),
    )
    for t in leg.get_texts():
        t.set_color(_hex(NAVY))

    fig.tight_layout(pad=0.35)
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def _set_run(run, text: str, size: int, bold: bool = False, color=NAVY) -> None:
    run.text = text
    run.font.size = PptPt(size)
    run.font.bold = bold
    run.font.color.rgb = _rgb(color)
    run.font.name = "Calibri"


def _caja(slide, l, t, w, h, fill=LIGHT, line=None):
    sh = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        PptInches(l),
        PptInches(t),
        PptInches(w),
        PptInches(h),
    )
    sh.fill.solid()
    sh.fill.fore_color.rgb = _rgb(fill)
    sh.line.fill.background()
    if line:
        sh.line.fill.solid()
        sh.line.color.rgb = _rgb(line)
        sh.line.width = Emu(6350)
    try:
        sh.adjustments[0] = 0.08
    except Exception:
        pass
    return sh


def _tb(slide, l, t, w, h, lines: List[Tuple[str, int, bool, Tuple[int, int, int]]], align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(PptInches(l), PptInches(t), PptInches(w), PptInches(h))
    tf = box.text_frame
    tf.word_wrap = True
    for i, (text, size, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = PptPt(3)
        run = p.add_run()
        _set_run(run, text, size, bold, color)
    return box


def _header_bar(slide, prs, titulo: str, sub: str) -> None:
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, PptInches(0.92)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = _rgb(NAVY)
    bar.line.fill.background()
    gold = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, PptInches(0.92), prs.slide_width, PptInches(0.06)
    )
    gold.fill.solid()
    gold.fill.fore_color.rgb = _rgb(GOLD)
    gold.line.fill.background()
    if LOGO.is_file():
        slide.shapes.add_picture(str(LOGO), PptInches(11.85), PptInches(0.16), width=PptInches(1.25))
    _tb(slide, 0.28, 0.12, 11.6, 0.42, [(titulo, 20, True, WHITE)])
    _tb(slide, 0.28, 0.50, 11.6, 0.36, [(sub, 11, False, (220, 230, 240))])


def _portada(prs) -> None:
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    if FONDO.is_file():
        pic = sl.shapes.add_picture(str(FONDO), 0, 0, width=prs.slide_width, height=prs.slide_height)
        spTree = sl.shapes._spTree
        spTree.remove(pic.element)
        spTree.insert(2, pic.element)
    veil = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, PptInches(3.55), prs.slide_width, PptInches(3.95))
    veil.fill.solid()
    veil.fill.fore_color.rgb = _rgb(NAVY)
    veil.line.fill.background()
    _tb(sl, 0.6, 3.75, 12, 0.5, [("WES  ·  Parque Arauco", 16, True, GOLD)])
    _tb(sl, 0.6, 4.18, 12, 0.7, [("Recorrido ejecutivo por recinto", 32, True, WHITE)])
    _tb(
        sl,
        0.6,
        4.95,
        12,
        1.2,
        [
            ("Empezamos por MAE  ·  Mall Arauco Estación", 16, False, WHITE),
            (f"Período {PERIODO}   |   Emisión {FECHA_EMISION}", 15, False, (220, 230, 240)),
            ("PPT aparte de las fichas  ·  un tema por lámina", 14, False, GOLD),
        ],
    )


def _slide_equipos(prs) -> None:
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _header_bar(
        sl,
        prs,
        "MAE  ·  1. Equipos instalados",
        "4 puntos activos WES   |   Mall Arauco Estación   |   Recepción 20/10/2025",
    )
    # 2x2 cards
    cards = [
        ("000025-01", "Estanque Norte — locales mall"),
        ("000025-04", "Baños públicos del recinto"),
        ("000025-07", "Local Pizza Hut"),
        ("000025-19", "Sala de bomba — estanque sur"),
    ]
    positions = [(0.28, 1.18), (6.78, 1.18), (0.28, 3.55), (6.78, 3.55)]
    for (nid, nota), (x, y) in zip(cards, positions):
        _caja(sl, x, y, 6.28, 2.18, fill=LIGHT, line=TEAL)
        _tb(sl, x + 0.22, y + 0.18, 5.85, 0.32, [(nid, 12, True, GOLD)])
        _tb(sl, x + 0.22, y + 0.52, 5.85, 0.70, [(NOMBRE_LARGO[nid], 22, True, NAVY)])
        _tb(sl, x + 0.22, y + 1.32, 5.85, 0.55, [(nota, 14, False, GRAY)])

    _caja(sl, 0.28, 5.90, 12.78, 1.28, fill=(255, 249, 235), line=GOLD)
    _tb(sl, 0.48, 6.02, 12.4, 0.28, [("RECINTO", 11, True, GOLD)])
    _tb(
        sl,
        0.48,
        6.32,
        12.4,
        0.72,
        [
            (
                "Recepción 20/10/2025  ·  Capacitación 18/02/2025  ·  "
                "Usuarios: medioambiente.dcl@parauco.com  ·  Sala de monitores MAE  ·  "
                "Sergio Fuenzalida — Analista Gestión Ambiental",
                13,
                False,
                NAVY,
            )
        ],
    )


def _slide_consumo(prs, by: Dict[str, Dict[str, Any]], tot: Dict[str, float]) -> None:
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _header_bar(
        sl,
        prs,
        "MAE  ·  2. Consumo mensualizado",
        f"Suma de los 4 puntos WES   |   {PERIODO}   |   Agosto: a la fecha vs proyección al 31",
    )

    ch_mes = CHARTS / "mae_mensual_may_ago.png"
    chart_mensual_mae(ch_mes, tot)

    _caja(sl, 0.22, 1.12, 8.72, 5.05)
    sl.shapes.add_picture(str(ch_mes), PptInches(0.38), PptInches(1.22), width=PptInches(8.40))

    _caja(sl, 9.08, 1.12, 4.02, 5.05, fill=WHITE, line=TEAL)
    _tb(sl, 9.22, 1.22, 3.74, 0.28, [("JULIO · último mes cerrado", 11, True, TEAL)])
    _tb(
        sl,
        9.22,
        1.50,
        3.74,
        0.62,
        [
            (
                f"El recinto sumó {fn(tot['jul'], 0)} m³. Peso de cada equipo en el total:",
                13,
                False,
                GRAY,
            )
        ],
    )
    ranked = sorted(MAE_NODOS, key=lambda n: -float((by.get(n) or {}).get("jul") or 0))
    y = 2.20
    for nid in ranked:
        v = float((by.get(nid) or {}).get("jul") or 0)
        pct = (v / tot["jul"] * 100) if tot["jul"] else 0
        _caja(sl, 9.22, y, 3.74, 0.88, fill=LIGHT, line=COLOR_NODO[nid])
        _tb(sl, 9.36, y + 0.08, 3.46, 0.28, [(NOMBRE_CORTO[nid], 13, True, COLOR_NODO[nid])])
        _tb(
            sl,
            9.36,
            y + 0.38,
            3.46,
            0.40,
            [(f"{fn(v, 0)} m³    {fn(pct, 0)} %", 18, True, NAVY)],
        )
        y += 0.96

    _tb(
        sl,
        0.28,
        6.28,
        13.0,
        1.05,
        [
            (
                f"Mayo {fn(tot['may'], 0)} m³   ·   Junio {fn(tot['jun'], 0)} m³   ·   "
                f"Julio {fn(tot['jul'], 0)} m³   ·   Agosto {AGO_ETQ}: {fn(tot['ago'], 0)} m³   ·   "
                f"Proyección agosto: {fn(tot['ago_proy'], 0)} m³ ({fn(tot['ago_d'])} m³/día × 31).",
                13,
                False,
                NAVY,
            ),
            (
                "Dorado = lo que agosto ya lleva. Tramo claro = proyección del resto del mes. "
                "Mayo cierra más alto por Estanque Sur, antes de la reparación del 10/06.",
                12,
                False,
                GRAY,
            ),
        ],
    )


def build_ppt(by: Dict[str, Dict[str, Any]], tot: Dict[str, float]) -> Path:
    prs = Presentation()
    prs.slide_width = PptInches(13.333)
    prs.slide_height = PptInches(7.5)
    _portada(prs)
    _slide_equipos(prs)
    _slide_consumo(prs, by, tot)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"Recorrido_ejecutivo_PA_MAE_{HASTA.strftime('%Y%m%d')}.pptx"
    prs.save(str(path))
    print(f"[OK] PPT {path}")
    return path


def main() -> int:
    skip = "--skip-refresh" in sys.argv
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHARTS.mkdir(parents=True, exist_ok=True)
    if not skip or not JSON_DATOS.is_file():
        refrescar_datos()
    _names, by, tot = cargar_mae()
    ppt = build_ppt(by, tot)
    print("\n=== SALIDA ===")
    print(ppt)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
