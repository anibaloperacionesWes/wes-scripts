# -*- coding: utf-8 -*-
"""
Recorrido ejecutivo Parque Arauco — PPT aparte de las fichas.

Misma estética (navy / gold / cards / fondo PA). Se arma mall por mall.
Esta versión: MAE.

  1) Equipos instalados — 4 puntos del deck (sin 000025-02)
  2) Consumo mensualizado — gráfico mayo → fecha (agosto a la fecha vs proyección)
  3) Hallazgos — Estanque Sur (presostatos, diario) + horas en cero del control nocturno

Uso:
  python3 generar_ppt_recorrido_ejecutivo_pa.py
  python3 generar_ppt_recorrido_ejecutivo_pa.py --skip-refresh
"""

from __future__ import annotations

import json
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import ListedColormap
from matplotlib.patches import Patch
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
JSON_NOCHES = OUT_DIR / "noches_control_mae.json"
JSON_PERFILES = OUT_DIR / "perfiles_horarios_control_mae.json"
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

# Hallazgos MAE
SUR_REPARACION = date(2026, 6, 10)  # presostatos Estanque Sur
CTRL_PIZZA = date(2026, 7, 1)
CTRL_NORTE = date(2026, 8, 5)
# 26–29/07 Pizza Hut: la noche volvió; no entra a la mediana post-control.
PIZZA_NOCHES_ATIPICAS = {date(2026, 7, 26), date(2026, 7, 27), date(2026, 7, 28), date(2026, 7, 29)}

NAVY = (13, 59, 102)
GOLD = (201, 162, 39)
TEAL = (31, 119, 180)
GRAY = (90, 90, 90)
LIGHT = (245, 247, 250)
WHITE = (255, 255, 255)
GOLD_SOFT = (232, 213, 163)
TARIFA_CLP_M3 = 1400.0  # misma tarifa del PPT 7 Malls


def fn(v: float, dec: int = 1) -> str:
    return format_number_chilean(float(v), dec)


def _clp(m3: float) -> str:
    return f"${fn(float(m3) * TARIFA_CLP_M3, 0)}"


def _rango_horas(horas: List[int]) -> str:
    if not horas:
        return "—"
    h = sorted(horas)
    return f"{h[0]:02d}:00–{h[-1]:02d}:00"


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


def _n06_de_serie(serie) -> float:
    return sum(float(v) for h, v in (serie or []) if int(h) < 6)


def _rango_dias(d0: date, d1: date) -> List[date]:
    out = []
    d = d0
    while d <= d1:
        out.append(d)
        d += timedelta(days=1)
    return out


def refrescar_noches() -> Dict[str, Dict[str, float]]:
    """m³ 00:00–06:00 por día para Pizza Hut y Estanque Norte (controles MAE)."""
    from generar_reporte_word import get_hourly_measures_for_day

    hourly: Dict[str, Dict[str, float]] = {"000025-07": {}, "000025-01": {}}
    if JSON_NOCHES.is_file():
        try:
            hourly = json.loads(JSON_NOCHES.read_text(encoding="utf-8")).get("hourly") or hourly
        except Exception:
            pass
    pendientes: List[Tuple[str, date]] = []
    for nid, d0, d1 in (
        ("000025-07", date(2026, 6, 20), HASTA),
        ("000025-01", date(2026, 7, 25), HASTA),
    ):
        hourly.setdefault(nid, {})
        for d in _rango_dias(d0, d1):
            if d.isoformat() not in hourly[nid]:
                pendientes.append((nid, d))
    print(f"[INFO] Noches MAE a descargar: {len(pendientes)}", flush=True)

    def _uno(nid: str, d: date) -> Tuple[str, str, float]:
        serie = get_hourly_measures_for_day(nid, datetime(d.year, d.month, d.day)) or []
        return nid, d.isoformat(), round(_n06_de_serie(serie), 2)

    if pendientes:
        with ThreadPoolExecutor(max_workers=6) as pool:
            futs = [pool.submit(_uno, nid, d) for nid, d in pendientes]
            for i, fut in enumerate(as_completed(futs), 1):
                nid, iso, n06 = fut.result()
                hourly[nid][iso] = n06
                if i % 15 == 0 or i == len(pendientes):
                    print(f"  {i}/{len(pendientes)} noches", flush=True)
    JSON_NOCHES.write_text(
        json.dumps({"hourly": hourly}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return hourly


def cargar_noches() -> Dict[str, Dict[str, float]]:
    if not JSON_NOCHES.is_file():
        return {}
    return json.loads(JSON_NOCHES.read_text(encoding="utf-8")).get("hourly") or {}


def cargar_perfiles() -> Dict[str, Dict[str, Dict[str, float]]]:
    if not JSON_PERFILES.is_file():
        return {}
    return json.loads(JSON_PERFILES.read_text(encoding="utf-8")).get("by_h") or {}


def refrescar_perfiles() -> Dict[str, Dict[str, Dict[str, float]]]:
    """Perfil horario 0–7 h en días con control (para marcar horas en cero)."""
    from generar_reporte_word import get_hourly_measures_for_day

    by_h: Dict[str, Dict[str, Dict[str, float]]] = cargar_perfiles() or {
        "000025-01": {},
        "000025-07": {},
        "000025-19": {},
    }
    pendientes: List[Tuple[str, date]] = []
    for nid, d0, d1 in (
        ("000025-01", CTRL_NORTE, HASTA),
        ("000025-07", CTRL_PIZZA, HASTA),
    ):
        by_h.setdefault(nid, {})
        for d in _rango_dias(d0, d1):
            if d.isoformat() not in by_h[nid]:
                pendientes.append((nid, d))
    print(f"[INFO] Perfiles horarios a descargar: {len(pendientes)}", flush=True)

    def _uno(nid: str, d: date) -> Tuple[str, str, Dict[str, float]]:
        serie = get_hourly_measures_for_day(nid, datetime(d.year, d.month, d.day)) or []
        rec = {str(int(h)): round(float(v), 3) for h, v in serie}
        return nid, d.isoformat(), rec

    if pendientes:
        with ThreadPoolExecutor(max_workers=6) as pool:
            futs = [pool.submit(_uno, nid, d) for nid, d in pendientes]
            for i, fut in enumerate(as_completed(futs), 1):
                nid, iso, rec = fut.result()
                by_h[nid][iso] = rec
                if i % 20 == 0 or i == len(pendientes):
                    print(f"  {i}/{len(pendientes)} perfiles", flush=True)
    JSON_PERFILES.write_text(
        json.dumps({"by_h": by_h}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return by_h


def _mediana(vals: List[float]) -> float:
    if not vals:
        return 0.0
    s = sorted(vals)
    n = len(s)
    if n % 2:
        return float(s[n // 2])
    return 0.5 * (s[n // 2 - 1] + s[n // 2])


def _stats_noche(hourly: Dict[str, Dict[str, float]], nid: str, ctrl: date) -> Dict[str, float]:
    serie = hourly.get(nid) or {}
    pre, post = [], []
    for iso, v in serie.items():
        d = date.fromisoformat(iso)
        if nid == "000025-07" and d in PIZZA_NOCHES_ATIPICAS:
            continue
        if d < ctrl:
            pre.append(float(v))
        else:
            post.append(float(v))
    pre_m = _mediana(pre)
    post_m = _mediana(post)
    ahorro = max(0.0, pre_m - post_m)
    return {
        "pre": pre_m,
        "post": post_m,
        "ahorro_noche": ahorro,
        "n_pre": float(len(pre)),
        "n_post": float(len(post)),
        "ahorro_acum": ahorro * float(len(post)),
    }


def _avg_daily(daily: Dict[str, float], d0: date, d1: date) -> float:
    vals = [
        float(v)
        for iso, v in daily.items()
        if d0 <= date.fromisoformat(iso) <= d1
    ]
    return (sum(vals) / len(vals)) if vals else 0.0


def chart_sur_diario(path: Path, daily: Dict[str, float]) -> Dict[str, float]:
    """Estanque Sur m³/día. Marca la reparación de presostatos (10/06)."""
    dias = sorted(date.fromisoformat(k) for k in daily)
    vals = [float(daily[d.isoformat()]) for d in dias]
    pre = _avg_daily(daily, DESDE, SUR_REPARACION - timedelta(days=1))
    post = _avg_daily(daily, SUR_REPARACION + timedelta(days=1), HASTA)
    dia10 = float(daily.get(SUR_REPARACION.isoformat()) or 0.0)
    n_post = sum(
        1
        for iso in daily
        if SUR_REPARACION + timedelta(days=1) <= date.fromisoformat(iso) <= HASTA
    )
    baja = max(0.0, pre - post)
    m3_acum = baja * float(n_post)

    fig, ax = plt.subplots(figsize=(7.15, 3.55), dpi=160)
    ax.axvspan(datetime(2026, 5, 1), datetime(2026, 6, 10), color="#F4E6C8", alpha=0.55, zorder=0)
    ax.axvspan(datetime(2026, 6, 10), datetime(2026, 8, 18), color="#E7F1F8", alpha=0.55, zorder=0)
    ax.plot(dias, vals, color=_hex(TEAL), linewidth=1.7, zorder=3)
    ax.fill_between(dias, vals, 0, color=_hex(TEAL), alpha=0.18, zorder=2)
    ax.axvline(SUR_REPARACION, color="#C04545", linestyle="--", linewidth=1.3, zorder=4)
    ax.axhline(pre, color="#C9A227", linestyle=":", linewidth=1.2, zorder=4)
    ax.axhline(post, color=_hex(NAVY), linestyle=":", linewidth=1.2, zorder=4)
    ax.plot([SUR_REPARACION], [dia10], marker="o", markersize=7, color="#C04545", zorder=5)
    ax.annotate(
        f"10/06  {fn(dia10, 0)} m³",
        xy=(SUR_REPARACION, dia10),
        xytext=(15, 18),
        textcoords="offset points",
        fontsize=9,
        color="#C04545",
        fontweight="bold",
        arrowprops=dict(arrowstyle="->", color="#C04545", lw=0.9),
    )
    ax.text(
        datetime(2026, 5, 12),
        pre + 4,
        f"antes  {fn(pre, 0)} m³/día",
        fontsize=8.5,
        color="#8A6A12",
        fontweight="bold",
    )
    ax.text(
        datetime(2026, 7, 5),
        post + 6,
        f"después  {fn(post, 0)} m³/día",
        fontsize=8.5,
        color=_hex(NAVY),
        fontweight="bold",
    )
    ax.text(
        datetime(2026, 7, 8),
        max(vals) * 0.78 if vals else 90,
        f"−{fn(baja, 0)} m³/día\n{_clp(baja)} / día\n{_clp(baja * 30)} / mes",
        fontsize=9,
        color=_hex(NAVY),
        fontweight="bold",
        va="top",
        ha="left",
        bbox=dict(boxstyle="round,pad=0.35", facecolor="white", edgecolor=_hex(GOLD), linewidth=1.2),
    )
    ax.set_ylabel("m³/día", fontsize=10, color=_hex(NAVY))
    ax.set_xlim(datetime(2026, 5, 1), datetime(2026, 8, 18))
    ax.set_ylim(0, max(vals) * 1.18 if vals else 1)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d/%m"))
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=2))
    ax.tick_params(labelsize=8, colors=_hex(NAVY))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#C5CDD6")
    ax.spines["bottom"].set_color("#C5CDD6")
    ax.yaxis.grid(True, linestyle=":", alpha=0.5, zorder=1)
    ax.set_axisbelow(True)
    fig.autofmt_xdate(rotation=30, ha="right")
    fig.tight_layout(pad=0.25)
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return {
        "pre": pre,
        "post": post,
        "dia10": dia10,
        "baja": baja,
        "n_post": n_post,
        "m3_acum": m3_acum,
    }


def chart_horas_en_cero(
    path: Path,
    perfiles: Dict[str, Dict[str, Dict[str, float]]],
) -> Dict[str, Any]:
    """Mapa día × hora: dorado = en cero. Estanque Norte desde el 05/08."""
    cero = 0.05
    horas = list(range(7))  # 00–06
    norte = perfiles.get("000025-01") or {}
    dias = sorted(
        date.fromisoformat(k) for k in norte if date.fromisoformat(k) >= CTRL_NORTE
    )
    if not dias:
        dias = []
    mat = np.zeros((len(dias), len(horas)))
    for i, d in enumerate(dias):
        rec = norte.get(d.isoformat()) or {}
        for j, h in enumerate(horas):
            v = float(rec.get(str(h), 0.0) or 0.0)
            mat[i, j] = 1.0 if v < cero else 0.0  # 1 = en cero

    fig, ax = plt.subplots(figsize=(5.55, 3.55), dpi=160)
    cmap = ListedColormap(["#0D3B66", "#C9A227"])
    ax.imshow(mat, aspect="auto", cmap=cmap, vmin=0, vmax=1, interpolation="nearest")
    ax.set_xticks(range(len(horas)))
    ax.set_xticklabels([f"{h:02d}" for h in horas], fontsize=9, color=_hex(NAVY))
    ax.set_yticks(range(len(dias)))
    ax.set_yticklabels([d.strftime("%d/%m") for d in dias], fontsize=8, color=_hex(NAVY))
    ax.set_xlabel("Hora", fontsize=9, color=_hex(NAVY))
    ax.tick_params(colors=_hex(NAVY))
    for spine in ax.spines.values():
        spine.set_color("#C5CDD6")
    ax.set_xticks(np.arange(-0.5, len(horas), 1), minor=True)
    ax.set_yticks(np.arange(-0.5, len(dias), 1), minor=True)
    ax.grid(which="minor", color="white", linestyle="-", linewidth=1.1)
    ax.tick_params(which="minor", bottom=False, left=False)
    ax.legend(
        handles=[
            Patch(facecolor="#C9A227", edgecolor="none", label="En cero"),
            Patch(facecolor="#0D3B66", edgecolor="none", label="Con caudal"),
        ],
        loc="lower right",
        frameon=False,
        fontsize=8,
        bbox_to_anchor=(1.02, -0.22),
        ncol=2,
        labelcolor=_hex(NAVY),
    )
    fig.tight_layout(pad=0.35)
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)

    # Horas que quedan en cero casi todas las noches (≥ 80 %).
    n = len(dias) or 1
    horas_cero = []
    for j, h in enumerate(horas):
        pct = 100.0 * float(mat[:, j].sum()) / n
        if pct >= 80:
            horas_cero.append(h)
    pizza = perfiles.get("000025-07") or {}
    pizza_dias = [
        date.fromisoformat(k)
        for k in pizza
        if date.fromisoformat(k) >= CTRL_PIZZA and date.fromisoformat(k) not in PIZZA_NOCHES_ATIPICAS
    ]
    pizza_cero = []
    if pizza_dias:
        for h in horas:
            c = 0
            for d in pizza_dias:
                v = float((pizza.get(d.isoformat()) or {}).get(str(h), 0) or 0)
                if v < cero:
                    c += 1
            if 100.0 * c / len(pizza_dias) >= 80:
                pizza_cero.append(h)
    return {"norte_horas_cero": horas_cero, "pizza_horas_cero": pizza_cero, "n_norte": len(dias)}


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


def _slide_hallazgos(
    prs,
    by: Dict[str, Dict[str, Any]],
    perfiles: Dict[str, Dict[str, Dict[str, float]]],
) -> None:
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _header_bar(
        sl,
        prs,
        "MAE  ·  3. Hallazgos",
        "Estanque Sur: reparación de presostatos (10/06)  ·  Control nocturno: horas en cero",
    )

    daily_sur = (by.get("000025-19") or {}).get("daily") or {}
    ch_sur = CHARTS / "mae_sur_diario_presostatos.png"
    ch_noc = CHARTS / "mae_horas_en_cero.png"
    st_sur = chart_sur_diario(ch_sur, daily_sur)
    st_h = chart_horas_en_cero(ch_noc, perfiles)
    baja = st_sur["baja"]
    pct = (baja / st_sur["pre"] * 100) if st_sur["pre"] else 0.0

    _caja(sl, 0.22, 1.08, 12.88, 1.28, fill=(255, 249, 235), line=GOLD)
    _tb(sl, 0.40, 1.14, 12.55, 0.26, [("ESTANQUE SUR  ·  PRESÓSTATOS", 11, True, GOLD)])
    _tb(
        sl,
        0.40,
        1.40,
        12.55,
        0.88,
        [
            (
                f"El 10/06 se repararon los presostatos. El estanque venía en "
                f"{fn(st_sur['pre'], 0)} m³/día (mayo–9/jun) y ese día ya marca "
                f"{fn(st_sur['dia10'], 0)} m³. Desde el 11/06 se sostiene en "
                f"{fn(st_sur['post'], 0)} m³/día (−{fn(baja, 0)} m³/día, {fn(pct, 0)}%). "
                f"A ${fn(TARIFA_CLP_M3, 0)}/m³: {_clp(baja)}/día  ·  {_clp(baja * 30)}/mes  ·  "
                f"acumulado 11/06–{HASTA.strftime('%d/%m')}: {_clp(st_sur['m3_acum'])} "
                f"({fn(st_sur['m3_acum'], 0)} m³ en {int(st_sur['n_post'])} días).",
                12,
                False,
                NAVY,
            )
        ],
    )

    _caja(sl, 0.22, 2.46, 5.78, 3.78)
    _tb(sl, 0.36, 2.52, 5.50, 0.26, [("ESTANQUE NORTE — horas en cero (desde 05/08)", 11, True, TEAL)])
    sl.shapes.add_picture(
        str(ch_noc), PptInches(0.36), PptInches(2.80), width=PptInches(5.48), height=PptInches(3.20)
    )

    _caja(sl, 6.14, 2.46, 6.96, 3.78)
    _tb(sl, 6.28, 2.52, 6.68, 0.26, [("ESTANQUE SUR — m³/día y costo evitado", 11, True, TEAL)])
    sl.shapes.add_picture(
        str(ch_sur), PptInches(6.28), PptInches(2.80), width=PptInches(6.68), height=PptInches(3.20)
    )

    norte_r = _rango_horas(st_h.get("norte_horas_cero") or [])
    pizza_r = _rango_horas(st_h.get("pizza_horas_cero") or [])
    _tb(
        sl,
        0.28,
        6.32,
        12.8,
        1.02,
        [
            (
                f"Estanque Norte (control 05/08): en cero {norte_r} "
                f"({int(st_h.get('n_norte') or 0)} noches). "
                "Algunos días el cero parte desde las 00:00 (p. ej. 05/08).",
                12,
                False,
                NAVY,
            ),
            (
                f"Pizza Hut (control 01/07): en cero {pizza_r}. "
                "Estanque Sur: corte on/off de mantención nocturna.",
                12,
                False,
                GRAY,
            ),
        ],
    )


def build_ppt(
    by: Dict[str, Dict[str, Any]],
    tot: Dict[str, float],
    perfiles: Dict[str, Dict[str, Dict[str, float]]],
) -> Path:
    prs = Presentation()
    prs.slide_width = PptInches(13.333)
    prs.slide_height = PptInches(7.5)
    _portada(prs)
    _slide_equipos(prs)
    _slide_consumo(prs, by, tot)
    _slide_hallazgos(prs, by, perfiles)
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
    if not skip or not JSON_NOCHES.is_file():
        refrescar_noches()
    if not skip or not JSON_PERFILES.is_file():
        refrescar_perfiles()
    _names, by, tot = cargar_mae()
    perfiles = cargar_perfiles()
    if not (perfiles.get("000025-01") and perfiles.get("000025-07")):
        perfiles = refrescar_perfiles()
    ppt = build_ppt(by, tot, perfiles)
    print("\n=== SALIDA ===")
    print(ppt)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
