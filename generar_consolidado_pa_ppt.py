# -*- coding: utf-8 -*-
"""
Consolidado Parque Arauco (7 malls) — misma estética del PPT de presentación
(recorrido ejecutivo: navy / gold / fondo PA) con datos y antecedentes al día.

Cubre MAE · MAM · MAQ · BOM · AEB · CUR · PAK.

  python3 generar_consolidado_pa_ppt.py
  python3 generar_consolidado_pa_ppt.py --hasta 08/09/2026

Salida:
  reports/Parque_Arauco/CONSOLIDADO/Consolidado_PA_7malls_<hasta>.pptx
"""

from __future__ import annotations

import argparse
import json
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches as PptInches

from generar_ppt_recorrido_ejecutivo_pa import (
    CHIP_NOTA,
    COLOR_NODO,
    FONDO,
    GOLD,
    GRAY,
    LIGHT,
    LOGO,
    MALLS,
    NAVY,
    NOMBRE_CORTO,
    TEAL,
    WHITE,
    _caja,
    _header_bar,
    _hex,
    _rgb,
    _tb,
    fn,
    nodos_todos,
    nodos_totales,
)
from generar_reporte_word import get_hourly_measures_for_day
from generar_reportes_y_ppt_mall_maipu import obtener_datos_agregados
from reporte_puntos_en_cero import (
    HORAS_UMBRAL_CONEXION_APP,
    _fmt_antiguedad,
    _horas_desde,
    obtener_estado_conexion_nodo,
)

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "reports" / "Parque_Arauco" / "CONSOLIDADO"
CHARTS = OUT_DIR / "charts_consolidado"
JSON_DAILY = OUT_DIR / "datos_consolidado_diario.json"
JSON_NIGHT = OUT_DIR / "datos_consolidado_noches.json"
JSON_CONN = OUT_DIR / "datos_consolidado_conexion.json"

DESDE_HIST = date(2026, 5, 1)
NODOS_CLAVE_NOCHE = [
    "000025-01",
    "000025-07",
    "000025-19",
    "000025-08",
    "000025-09",
    "000025-10",
    "000025-13",
    "000025-17",
    "000025-18",
    "000025-11",
    "000025-12",
    "000025-37",
    "000025-38",
    "000025-27",
    "000025-35",
    "000025-36",
]
ANTECEDENTES_FIJOS: List[Tuple[str, str]] = [
    (
        "MAE",
        "Estanque Sur: reparación de presostatos 10/06. Pizza Hut con control nocturno desde 01/07. Estanque Norte con control desde 05/08 (00–05 h).",
    ),
    (
        "MAM",
        "Placa Bancaria sube desde 18/06. Falabella reactivado 11/08 (antes sala de bombas en reparación / fuera de puntos en cero).",
    ),
    (
        "MAQ",
        "Matriz Principal concentra el recinto. El día se duplicó desde el 22/06; Alim. Baños es uso hábil.",
    ),
    (
        "BOM",
        "San Ignacio 500 con on/off nocturno operativo desde 17/07. San Ignacio 300 queda en monitoreo.",
    ),
    (
        "AEB",
        "Matriz A.A. desactivada 15/05; el recinto se lee en Matriz 1° piso y Anillo Plaza.",
    ),
    (
        "CUR",
        "Anillo Norte y Sur desde el 19/05 (mayo no se grafica). Noche chica; no se propone on/off.",
    ),
    (
        "PAK",
        "000025-25/26 retirados 30/03 (reemplazo 35/36). Distrito de Lujo, Bazar Gourmet y DL Kennedy son cadena: no se suman a la cabecera.",
    ),
]


def _parse_ddmmyyyy(s: str) -> date:
    return datetime.strptime(s.strip(), "%d/%m/%Y").date()


def _rango(d0: date, d1: date) -> List[date]:
    out: List[date] = []
    d = d0
    while d <= d1:
        out.append(d)
        d += timedelta(days=1)
    return out


def _medida_fecha(m: Any) -> Optional[date]:
    raw = getattr(m, "date", None)
    if raw is None and isinstance(m, dict):
        raw = m.get("date")
    if raw is None:
        return None
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    s = str(raw)[:10]
    try:
        return date.fromisoformat(s)
    except ValueError:
        return None


def _medida_m3(m: Any) -> float:
    if hasattr(m, "total_m3"):
        return float(m.total_m3 or 0)
    if isinstance(m, dict):
        return float(m.get("total_m3") or m.get("totalM3") or 0)
    return 0.0


def _n06(serie) -> float:
    return sum(float(v) for h, v in (serie or []) if int(h) < 6)


def _nombre(nid: str) -> str:
    return NOMBRE_CORTO.get(nid, nid)


def sumar_rango(daily: Dict[str, float], d0: date, d1: date) -> float:
    tot = 0.0
    for d in _rango(d0, d1):
        tot += float(daily.get(d.isoformat(), 0.0) or 0.0)
    return tot


def prom_rango(daily: Dict[str, float], d0: date, d1: date) -> float:
    dias = _rango(d0, d1)
    if not dias:
        return 0.0
    return sumar_rango(daily, d0, d1) / len(dias)


def mes_parcial(daily: Dict[str, float], year: int, month: int, hasta: date) -> Tuple[float, int]:
    tot = 0.0
    n = 0
    for iso, v in daily.items():
        d = date.fromisoformat(iso)
        if d.year == year and d.month == month and d <= hasta:
            tot += float(v or 0)
            n += 1
    return tot, n


def daily_de(by: Dict[str, Dict[str, Any]], nids: List[str]) -> Dict[str, float]:
    acc: Dict[str, float] = {}
    for nid in nids:
        for iso, v in ((by.get(nid) or {}).get("daily") or {}).items():
            acc[iso] = acc.get(iso, 0.0) + float(v or 0)
    return acc


def refrescar_diario(nodos: List[str], desde: date, hasta: date) -> Dict[str, Dict[str, Any]]:
    print(f"[INFO] Descargando diario {desde:%d/%m/%Y}–{hasta:%d/%m/%Y} ({len(nodos)} puntos)…", flush=True)
    datos = obtener_datos_agregados(
        nodos,
        desde.strftime("%d/%m/%Y"),
        hasta.strftime("%d/%m/%Y"),
    )
    by: Dict[str, Dict[str, Any]] = {}
    for ns in datos.get("nodes_summary") or []:
        nid = str(ns.get("node_id") or "")
        daily: Dict[str, float] = {}
        for m in ns.get("measures") or []:
            d = _medida_fecha(m)
            if d is None:
                continue
            daily[d.isoformat()] = daily.get(d.isoformat(), 0.0) + _medida_m3(m)
        by[nid] = {
            "name": str(ns.get("node_name") or _nombre(nid)),
            "daily": daily,
            "total": float((ns.get("summary") or {}).get("total") or sum(daily.values())),
        }
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    serial = {
        nid: {"name": row["name"], "daily": row["daily"], "total": row["total"]}
        for nid, row in by.items()
    }
    JSON_DAILY.write_text(json.dumps(serial, ensure_ascii=False, indent=2), encoding="utf-8")
    return by


def cargar_diario() -> Dict[str, Dict[str, Any]]:
    if not JSON_DAILY.is_file():
        return {}
    return json.loads(JSON_DAILY.read_text(encoding="utf-8"))


def refrescar_noches(nodos: List[str], d0: date, d1: date) -> Dict[str, Dict[str, float]]:
    hourly: Dict[str, Dict[str, float]] = {n: {} for n in nodos}
    if JSON_NIGHT.is_file():
        try:
            prev = json.loads(JSON_NIGHT.read_text(encoding="utf-8")).get("hourly") or {}
            for n in nodos:
                hourly[n] = dict(prev.get(n) or {})
        except Exception:
            pass
    pendientes: List[Tuple[str, date]] = []
    for nid in nodos:
        for d in _rango(d0, d1):
            if d.isoformat() not in hourly[nid]:
                pendientes.append((nid, d))
    print(f"[INFO] Noches 00–06 a descargar: {len(pendientes)}", flush=True)

    def _uno(nid: str, d: date) -> Tuple[str, str, float]:
        serie = get_hourly_measures_for_day(nid, datetime(d.year, d.month, d.day)) or []
        return nid, d.isoformat(), round(_n06(serie), 2)

    if pendientes:
        with ThreadPoolExecutor(max_workers=8) as pool:
            futs = [pool.submit(_uno, nid, d) for nid, d in pendientes]
            for i, fut in enumerate(as_completed(futs), 1):
                nid, iso, n06 = fut.result()
                hourly.setdefault(nid, {})[iso] = n06
                if i % 20 == 0 or i == len(pendientes):
                    print(f"  {i}/{len(pendientes)} noches", flush=True)
    JSON_NIGHT.write_text(json.dumps({"hourly": hourly}, ensure_ascii=False, indent=2), encoding="utf-8")
    return hourly


def refrescar_conexion(nodos: List[str]) -> Dict[str, Dict[str, Any]]:
    print(f"[INFO] Estado de conexión ({len(nodos)} puntos)…", flush=True)
    out: Dict[str, Dict[str, Any]] = {}

    def _uno(nid: str) -> Tuple[str, Dict[str, Any]]:
        st = obtener_estado_conexion_nodo(nid)
        horas = _horas_desde(st.get("lastUpdate"))
        desconectado = bool(horas is not None and horas >= HORAS_UMBRAL_CONEXION_APP)
        return nid, {
            "wesStatus": st.get("wesStatus"),
            "lastUpdate": st["lastUpdate"].isoformat() if st.get("lastUpdate") else None,
            "horas": horas,
            "desconectado": desconectado,
            "antiguedad": _fmt_antiguedad(horas) if horas is not None else "—",
        }

    with ThreadPoolExecutor(max_workers=10) as pool:
        futs = [pool.submit(_uno, n) for n in nodos]
        for fut in as_completed(futs):
            nid, rec = fut.result()
            out[nid] = rec
    JSON_CONN.write_text(json.dumps(out, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    return out


def chart_mensual(path: Path, daily: Dict[str, float], hasta: date) -> None:
    meses = [
        (2026, 5, "Mayo"),
        (2026, 6, "Junio"),
        (2026, 7, "Julio"),
        (2026, 8, "Agosto"),
        (2026, 9, f"Sep\n1–{hasta.day}" if hasta.month == 9 else "Septiembre"),
    ]
    vals = []
    labels = []
    for y, m, lab in meses:
        tot, n = mes_parcial(daily, y, m, hasta)
        if m == 9 and hasta < date(2026, 9, 1):
            continue
        vals.append(tot)
        labels.append(lab)
    x = np.arange(len(vals))
    fig, ax = plt.subplots(figsize=(6.4, 2.55), dpi=150)
    ax.bar(x, vals, color=_hex(TEAL), zorder=3)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=8, color=_hex(NAVY))
    ax.set_ylabel("m³", fontsize=8, color=_hex(NAVY))
    ax.tick_params(axis="y", labelsize=7, colors=_hex(NAVY))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#C5CDD6")
    ax.spines["bottom"].set_color("#C5CDD6")
    ax.yaxis.grid(True, linestyle=":", alpha=0.5, zorder=0)
    ax.set_axisbelow(True)
    ymax = max(vals + [1.0]) * 1.22
    ax.set_ylim(0, ymax)
    for xi, v in zip(x, vals):
        ax.text(xi, v + ymax * 0.02, fn(v, 0), ha="center", va="bottom", fontsize=7.5, color=_hex(NAVY), fontweight="bold")
    fig.tight_layout(pad=0.2)
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def chart_ultimos_dias(path: Path, daily: Dict[str, float], d0: date, d1: date, titulo: str) -> None:
    dias = _rango(d0, d1)
    ys = [float(daily.get(d.isoformat(), 0.0) or 0.0) for d in dias]
    fig, ax = plt.subplots(figsize=(6.4, 2.35), dpi=150)
    ax.plot(dias, ys, color=_hex(NAVY), linewidth=1.8, marker="o", markersize=3.2, zorder=3)
    ax.fill_between(dias, ys, color=_hex(TEAL), alpha=0.18)
    ax.set_ylabel("m³ / día", fontsize=8, color=_hex(NAVY))
    ax.tick_params(axis="y", labelsize=7, colors=_hex(NAVY))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d/%m"))
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(dias) // 7)))
    plt.setp(ax.get_xticklabels(), fontsize=7, color=_hex(NAVY), rotation=0)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#C5CDD6")
    ax.spines["bottom"].set_color("#C5CDD6")
    ax.yaxis.grid(True, linestyle=":", alpha=0.5, zorder=0)
    ax.set_axisbelow(True)
    ax.set_title(titulo, fontsize=9, color=_hex(NAVY), loc="left", pad=4)
    fig.tight_layout(pad=0.2)
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def _portada(prs, periodo: str, emision: str, ante: str) -> None:
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    if FONDO.is_file():
        pic = sl.shapes.add_picture(str(FONDO), 0, 0, width=prs.slide_width, height=prs.slide_height)
        spTree = sl.shapes._spTree
        spTree.remove(pic.element)
        spTree.insert(2, pic.element)
    veil = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, PptInches(3.40), prs.slide_width, PptInches(4.10))
    veil.fill.solid()
    veil.fill.fore_color.rgb = _rgb(NAVY)
    veil.line.fill.background()
    _tb(sl, 0.6, 3.55, 12, 0.4, [("WES  ·  Parque Arauco", 16, True, GOLD)])
    _tb(sl, 0.6, 3.95, 12, 0.7, [("Consolidado ejecutivo — 7 malls", 30, True, WHITE)])
    _tb(
        sl,
        0.6,
        4.70,
        12,
        1.5,
        [
            ("Misma lectura del PPT de presentación, con antecedentes al día", 16, False, WHITE),
            (f"Período {periodo}   |   Emisión {emision}", 15, False, (220, 230, 240)),
            (f"Antecedentes recientes: {ante}", 14, False, GOLD),
            ("MAE  ·  MAM  ·  MAQ  ·  BOM  ·  AEB  ·  CUR  ·  PAK", 16, False, GOLD),
        ],
    )
    if LOGO.is_file():
        sl.shapes.add_picture(str(LOGO), PptInches(11.70), PptInches(6.85), width=PptInches(1.35))


def _delta_txt(act: float, prev: float) -> str:
    if prev <= 0.05:
        return "n/d" if act <= 0.05 else "↑ n/d"
    pct = (act - prev) / prev * 100.0
    signo = "↑" if pct > 1 else ("↓" if pct < -1 else "≈")
    return f"{signo} {fn(abs(pct), 0)} %"


def _slide_resumen_malls(
    prs,
    by: Dict[str, Dict[str, Any]],
    s_act: Tuple[date, date],
    s_prev: Tuple[date, date],
    hasta: date,
) -> None:
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _header_bar(
        sl,
        prs,
        "Resumen por recinto",
        f"Semana {s_act[0]:%d/%m}–{s_act[1]:%d/%m} vs previa {s_prev[0]:%d/%m}–{s_prev[1]:%d/%m}   ·   agosto completo y septiembre a {hasta:%d/%m}",
    )
    headers = ["Mall", "Puntos", "Agosto m³", f"Sep 1–{hasta.day} m³", "Sem. previa", "Sem. actual", "Δ semana", "Top punto (sem.)"]
    xs = [0.22, 1.55, 2.55, 4.25, 5.95, 7.65, 9.35, 10.55]
    ws = [1.28, 0.95, 1.65, 1.65, 1.65, 1.65, 1.15, 2.55]
    y0 = 1.15
    h_h = 0.38
    for x, w, lab in zip(xs, ws, headers):
        _caja(sl, x, y0, w, h_h, fill=NAVY)
        _tb(sl, x + 0.04, y0 + 0.06, w - 0.08, 0.28, [(lab, 10, True, WHITE)], align=PP_ALIGN.CENTER)
    y = y0 + h_h + 0.06
    row_h = 0.72
    for i, mall in enumerate(MALLS):
        ids = nodos_totales(mall)
        daily = daily_de(by, ids)
        ago, _ = mes_parcial(daily, 2026, 8, hasta)
        sep, _ = mes_parcial(daily, 2026, 9, hasta)
        prev = sumar_rango(daily, *s_prev)
        act = sumar_rango(daily, *s_act)
        ranked = sorted(
            ids,
            key=lambda n: -sumar_rango((by.get(n) or {}).get("daily") or {}, *s_act),
        )
        top = ranked[0] if ranked else "—"
        top_v = sumar_rango((by.get(top) or {}).get("daily") or {}, *s_act) if ranked else 0
        fill = LIGHT if i % 2 == 0 else WHITE
        vals = [
            f"{mall['code']}  {mall['titulo'].replace('Mall Arauco ', '').replace('Arauco ', '').replace('Parque Arauco ', '')}",
            str(len(ids)),
            fn(ago, 0),
            fn(sep, 0),
            fn(prev, 0),
            fn(act, 0),
            _delta_txt(act, prev),
            f"{_nombre(top)}  {fn(top_v, 0)}",
        ]
        for x, w, lab in zip(xs, ws, vals):
            _caja(sl, x, y, w, row_h, fill=fill, line=(210, 216, 222))
            _tb(sl, x + 0.06, y + 0.18, w - 0.10, 0.40, [(lab, 11, True if x == xs[0] else False, NAVY)])
        y += row_h + 0.04
    _tb(
        sl,
        0.28,
        7.10,
        12.7,
        0.28,
        [("PAK: totales de cabecera (sin cadena DL / Bazar / DL Kennedy). Cifras API WES.", 11, False, GRAY)],
    )


RESIDUALES_CERO = {"000025-32", "000025-33"}  # Pasillo técnico / ARROW: caudal residual


def _puntos_cero_recientes(by: Dict[str, Dict[str, Any]], d0: date, d1: date) -> List[str]:
    lineas: List[str] = []
    dias = _rango(d0, d1)
    for mall in MALLS:
        ceros: List[str] = []
        for nid in mall["nodes"]:
            daily = (by.get(nid) or {}).get("daily") or {}
            vals = [float(daily.get(d.isoformat(), 0.0) or 0.0) for d in dias]
            if vals and all(v <= 0.05 for v in vals):
                extra = " · residual, no es alerta" if nid in RESIDUALES_CERO else ""
                ceros.append(f"{_nombre(nid)} ({nid}){extra}")
        if ceros:
            lineas.append(f"{mall['code']}: {', '.join(ceros)}")
    return lineas


def _hechos_estos_dias(
    by: Dict[str, Dict[str, Any]],
    n06: Dict[str, Dict[str, float]],
    d0: date,
    d1: date,
) -> List[str]:
    """Picos/caídas diarias materiales en la ventana reciente (no ruido de 1 m³)."""
    hechos: List[str] = []
    dias = _rango(d0, d1)
    for mall in MALLS:
        for nid in mall["nodes"]:
            if nid in RESIDUALES_CERO:
                continue
            daily = (by.get(nid) or {}).get("daily") or {}
            vals = [(d, float(daily.get(d.isoformat(), 0.0) or 0.0)) for d in dias]
            xs = [v for _, v in vals]
            if not xs:
                continue
            med = sorted(xs)[len(xs) // 2]
            if med < 8:
                continue
            dmax, vmax = max(vals, key=lambda t: t[1])
            dmin, vmin = min(vals, key=lambda t: t[1])
            if vmax >= med * 1.7 and vmax >= 40:
                hechos.append(
                    f"{mall['code']} {_nombre(nid)}: pico {dmax:%d/%m} {fn(vmax, 0)} m³ "
                    f"(mediana ventana {fn(med, 0)})."
                )
            if vmin <= max(med * 0.15, 8) and med >= 40:
                hechos.append(
                    f"{mall['code']} {_nombre(nid)}: caída {dmin:%d/%m} {fn(vmin, 1)} m³ "
                    f"(mediana ventana {fn(med, 0)}). Revisar dato/corte."
                )
            nserie = n06.get(nid) or {}
            nvals = [(d, float(nserie.get(d.isoformat(), 0.0) or 0.0)) for d in dias]
            if nvals:
                nmed = sorted(v for _, v in nvals)[len(nvals) // 2]
                dn, nv = max(nvals, key=lambda t: t[1])
                if nmed >= 4 and nv >= nmed * 1.8 and nv >= 12:
                    hechos.append(
                        f"{mall['code']} {_nombre(nid)}: noche 00–06 {dn:%d/%m} {fn(nv, 1)} m³ "
                        f"(mediana {fn(nmed, 1)})."
                    )
    cae = [h for h in hechos if "caída" in h]
    pico = [h for h in hechos if "pico" in h]
    noc = [h for h in hechos if "noche" in h]
    out: List[str] = []
    for grupo in (cae, pico, noc):
        for h in grupo:
            if h not in out:
                out.append(h)
        if len(out) >= 6:
            break
    return out[:6]


def _slide_antecedentes(
    prs,
    by: Dict[str, Dict[str, Any]],
    conn: Dict[str, Dict[str, Any]],
    n06: Dict[str, Dict[str, float]],
    ante0: date,
    ante1: date,
    s_act: Tuple[date, date],
    s_prev: Tuple[date, date],
) -> None:
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _header_bar(
        sl,
        prs,
        "Antecedentes de estos días",
        f"Ventana {ante0:%d/%m}–{ante1:%d/%m}  ·  conexión al corte de emisión  ·  noche 00–06 en puntos clave",
    )
    desc = []
    for nid, rec in sorted(conn.items()):
        if rec.get("desconectado"):
            mall = next((m["code"] for m in MALLS if nid in m["nodes"]), "?")
            desc.append(f"{mall} {_nombre(nid)} · lastUpdate {rec.get('antiguedad')}")
    ceros = _puntos_cero_recientes(by, ante0, ante1)
    noc_alta: List[str] = []
    for nid in NODOS_CLAVE_NOCHE:
        act = sum((n06.get(nid) or {}).get(d.isoformat(), 0.0) for d in _rango(*s_act))
        prev = sum((n06.get(nid) or {}).get(d.isoformat(), 0.0) for d in _rango(*s_prev))
        if prev > 4 and act > 8 and act > prev * 1.25:
            noc_alta.append(f"{_nombre(nid)} noche {fn(act, 1)} vs {fn(prev, 1)} m³ ({_delta_txt(act, prev)})")
    hechos = _hechos_estos_dias(by, n06, ante0, ante1)

    _caja(sl, 0.22, 1.12, 6.35, 2.55, fill=(255, 249, 235), line=GOLD)
    _tb(sl, 0.38, 1.18, 6.05, 0.26, [("AHORA  ·  CONEXIÓN", 12, True, GOLD)])
    if desc:
        _tb(sl, 0.38, 1.48, 6.05, 2.05, [(f"•  {t}", 12, False, NAVY) for t in desc[:8]])
    else:
        _tb(sl, 0.38, 1.48, 6.05, 2.05, [("Ningún punto del consolidado figura desconectado (lastUpdate < 2 h).", 13, False, NAVY)])

    _caja(sl, 6.75, 1.12, 6.35, 2.55, fill=WHITE, line=TEAL)
    _tb(sl, 6.91, 1.18, 6.05, 0.26, [("EN CERO  ·  TODA LA VENTANA", 12, True, TEAL)])
    if ceros:
        _tb(sl, 6.91, 1.48, 6.05, 2.05, [(f"•  {t}", 12, False, NAVY) for t in ceros[:8]])
    else:
        _tb(
            sl,
            6.91,
            1.48,
            6.05,
            2.05,
            [("Ningún punto operativo del deck acumula cero todos los días de la ventana.", 13, False, NAVY)],
        )

    _caja(sl, 0.22, 3.80, 12.88, 1.55, fill=LIGHT, line=TEAL)
    _tb(sl, 0.38, 3.86, 12.55, 0.24, [("HECHOS DE ESTOS DÍAS  ·  PICOS, CAÍDAS Y NOCHE", 12, True, TEAL)])
    mix = hechos + [t for t in noc_alta if all(t.split(" noche")[0] not in h for h in hechos)]
    if mix:
        _tb(sl, 0.38, 4.14, 12.55, 1.10, [(f"•  {t}", 13, False, NAVY) for t in mix[:4]])
    else:
        _tb(
            sl,
            0.38,
            4.14,
            12.55,
            1.10,
            [("Sin picos/caídas materiales ni alzas de madrugada > 25 % en puntos clave.", 13, False, NAVY)],
        )

    _caja(sl, 0.22, 5.48, 12.88, 1.78, fill=(255, 249, 235), line=GOLD)
    _tb(sl, 0.38, 5.54, 12.55, 0.24, [("CONTEXTO DEL PPT (SIGUE VIGENTE)", 12, True, GOLD)])
    _tb(
        sl,
        0.38,
        5.82,
        12.55,
        1.32,
        [(f"•  {cod} — {txt}", 11, False, NAVY) for cod, txt in ANTECEDENTES_FIJOS],
        space_after=2,
    )


def _hallazgos_mall(
    mall: Dict[str, Any],
    by: Dict[str, Dict[str, Any]],
    n06: Dict[str, Dict[str, float]],
    s_act: Tuple[date, date],
    s_prev: Tuple[date, date],
    hasta: date,
) -> List[str]:
    ids = nodos_totales(mall)
    daily = daily_de(by, ids)
    act = sumar_rango(daily, *s_act)
    prev = sumar_rango(daily, *s_prev)
    dias_act = max((s_act[1] - s_act[0]).days + 1, 1)
    ranked = sorted(
        mall["nodes"],
        key=lambda n: -sumar_rango((by.get(n) or {}).get("daily") or {}, *s_act),
    )
    top = ranked[0]
    top_v = sumar_rango((by.get(top) or {}).get("daily") or {}, *s_act)
    lines = [
        f"Semana {s_act[0]:%d/%m}–{s_act[1]:%d/%m}: {fn(act, 0)} m³ (prom. {fn(act / dias_act, 1)} m³/día) · Δ vs previa {_delta_txt(act, prev)}.",
        f"Mayor aporte de la semana: {_nombre(top)} ({top}) con {fn(top_v, 0)} m³.",
    ]
    noche_ids = [n for n in mall["nodes"] if n in NODOS_CLAVE_NOCHE]
    if noche_ids:
        partes = []
        for nid in noche_ids[:3]:
            nact = sum((n06.get(nid) or {}).get(d.isoformat(), 0.0) for d in _rango(*s_act))
            partes.append(f"{_nombre(nid)} {fn(nact, 1)}")
        lines.append("Noche 00–06 (suma semana): " + "  ·  ".join(partes) + " m³.")
    nota = CHIP_NOTA
    extras = [f"{_nombre(n)} ({nota[n]})" for n in mall["nodes"] if n in nota]
    if extras:
        lines.append("Notas de puntos: " + " · ".join(extras[:4]) + ".")
    ago, _ = mes_parcial(daily, 2026, 8, hasta)
    sep, nd = mes_parcial(daily, 2026, 9, hasta)
    if nd:
        lines.append(f"Agosto {fn(ago, 0)} m³  ·  septiembre 1–{hasta.day} {fn(sep, 0)} m³ ({fn(sep / nd, 1)} m³/día).")
    cap = str(mall.get("caption") or "").strip()
    if cap:
        lines.append(cap)
    return lines[:6]


def _slide_mall(
    prs,
    mall: Dict[str, Any],
    by: Dict[str, Dict[str, Any]],
    n06: Dict[str, Dict[str, float]],
    s_act: Tuple[date, date],
    s_prev: Tuple[date, date],
    ante0: date,
    hasta: date,
) -> None:
    ids = nodos_totales(mall)
    daily = daily_de(by, ids)
    p_mes = CHARTS / f"mensual_{mall['code']}.png"
    p_dias = CHARTS / f"dias_{mall['code']}.png"
    chart_mensual(p_mes, daily, hasta)
    chart_ultimos_dias(
        p_dias,
        daily,
        ante0,
        hasta,
        f"{mall['code']}  ·  m³/día  {ante0:%d/%m}–{hasta:%d/%m}",
    )
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    n_pts = len(mall["nodes"])
    _header_bar(
        sl,
        prs,
        f"{mall['code']}  ·  {mall['titulo']}",
        f"{n_pts} puntos WES   ·   recepción {mall.get('recepcion', '—')}   ·   {mall.get('usuarios', '')}",
    )
    hall = _hallazgos_mall(mall, by, n06, s_act, s_prev, hasta)
    _caja(sl, 0.22, 1.08, 12.88, 1.55, fill=(255, 249, 235), line=GOLD)
    _tb(sl, 0.38, 1.12, 12.55, 0.22, [("HALLAZGOS  ·  ESTOS DÍAS", 11, True, GOLD)])
    _tb(sl, 0.38, 1.34, 12.55, 1.22, [(f"•  {t}", 12, False, NAVY) for t in hall[:4]], space_after=2)

    if p_mes.is_file():
        sl.shapes.add_picture(str(p_mes), PptInches(0.28), PptInches(2.78), width=PptInches(6.35))
    if p_dias.is_file():
        sl.shapes.add_picture(str(p_dias), PptInches(6.75), PptInches(2.78), width=PptInches(6.35))

    y = 5.55
    chips = mall.get("chip_order") or mall["nodes"]
    n = len(chips)
    gap = 0.08
    w = (12.88 - gap * (n - 1)) / max(n, 1)
    x = 0.22
    for nid in chips:
        v = sumar_rango((by.get(nid) or {}).get("daily") or {}, *s_act)
        _caja(sl, x, y, w, 1.55, fill=WHITE, line=COLOR_NODO.get(nid, TEAL))
        _tb(sl, x + 0.06, y + 0.06, w - 0.10, 0.42, [(_nombre(nid), 11, True, NAVY)])
        _tb(sl, x + 0.06, y + 0.48, w - 0.10, 0.36, [(fn(v, 0) + " m³", 16, True, GOLD)])
        extra = CHIP_NOTA.get(nid, nid)
        _tb(sl, x + 0.06, y + 0.90, w - 0.10, 0.50, [(extra, 10, False, GRAY)])
        x += w + gap


def _slide_propuestas(prs, hasta: date) -> None:
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _header_bar(
        sl,
        prs,
        "Propuestas vigentes (PPT de presentación)",
        f"On/off = corte 00:00–06:00  ·  umbral = total 24 h (promedio operativo × 1,25)  ·  corte {hasta:%d/%m/%Y}",
    )
    filas = [
        ("MAE", "Mantener controles Norte (desde 05/08) y Pizza Hut (desde 01/07). Umbrales 24 h: Norte 40 · Sur 35 · Pizza 50 · Baños 10 m³/día."),
        ("MAM", "On/off a revisar en Placa. Falabella activo desde 11/08: umbral 140 m³/día. Placa 290 · Ripley según operativo."),
        ("MAQ", "Proponer on/off 00–06 en Matriz Principal. Umbral 24 h Matriz 240 m³/día. Alim. Baños no es el punto de control."),
        ("BOM", "Mantener on/off SI500 (desde 17/07). Umbrales: SI500 145 · SI300 45 m³/día."),
        ("AEB", "Proponer on/off 00–06 en Anillo Plaza y Matriz 1° piso. Umbrales: Anillo 22 · Matriz 75 m³/día."),
        ("CUR", "Sin on/off. Umbrales: Anillo Sur 14 · Anillo Norte 13 m³/día."),
        ("PAK", "Proponer on/off 00–06 en el eslabón de mayor noche (DL / Bazar). Umbrales: DL 390 · Bazar 250 · DL Kennedy 20 m³/día. Cadena no se suma a cabecera."),
    ]
    y0, gap, bottom = 1.10, 0.05, 7.32
    n = len(filas)
    h = (bottom - y0 - gap * (n - 1)) / n
    y = y0
    for code, txt in filas:
        _caja(sl, 0.22, y, 12.88, h, fill=(255, 249, 235), line=GOLD)
        _tb(sl, 0.36, y + 0.04, 12.60, 0.20, [(code, 12, True, GOLD)])
        _tb(sl, 0.36, y + 0.24, 12.60, h - 0.30, [(txt, 12, False, NAVY)])
        y += h + gap


def build_ppt(
    by: Dict[str, Dict[str, Any]],
    n06: Dict[str, Dict[str, float]],
    conn: Dict[str, Dict[str, Any]],
    desde: date,
    hasta: date,
    ante0: date,
    s_prev: Tuple[date, date],
    s_act: Tuple[date, date],
    emision: str,
) -> Path:
    prs = Presentation()
    prs.slide_width = PptInches(13.333)
    prs.slide_height = PptInches(7.5)
    periodo = f"{desde:%d/%m/%Y} – {hasta:%d/%m/%Y}"
    ante = f"{ante0:%d/%m/%Y} – {hasta:%d/%m/%Y}"
    _portada(prs, periodo, emision, ante)
    _slide_resumen_malls(prs, by, s_act, s_prev, hasta)
    _slide_antecedentes(prs, by, conn, n06, ante0, hasta, s_act, s_prev)
    for mall in MALLS:
        print(f"[INFO] Lámina {mall['code']}…", flush=True)
        _slide_mall(prs, mall, by, n06, s_act, s_prev, ante0, hasta)
    _slide_propuestas(prs, hasta)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"Consolidado_PA_7malls_{hasta.strftime('%Y%m%d')}.pptx"
    prs.save(str(path))
    print(f"[OK] PPT {path}", flush=True)
    return path


def _chile_today() -> date:
    try:
        from zoneinfo import ZoneInfo

        return datetime.now(ZoneInfo("America/Santiago")).date()
    except Exception:
        return date.today()


def main() -> int:
    hoy = _chile_today()
    default_hasta = hoy - timedelta(days=1)  # último día civil cerrado
    parser = argparse.ArgumentParser(description="Consolidado PA 7 malls (PPT presentación + antecedentes al día)")
    parser.add_argument("--desde", default=DESDE_HIST.strftime("%d/%m/%Y"))
    parser.add_argument("--hasta", default=default_hasta.strftime("%d/%m/%Y"))
    parser.add_argument("--skip-refresh", action="store_true")
    args = parser.parse_args()

    desde = _parse_ddmmyyyy(args.desde)
    hasta = _parse_ddmmyyyy(args.hasta)
    if hasta >= hoy:
        hasta = hoy - timedelta(days=1)
    ante0 = max(desde, hasta - timedelta(days=13))
    s_act = (max(desde, hasta - timedelta(days=6)), hasta)
    s_prev = (s_act[0] - timedelta(days=7), s_act[0] - timedelta(days=1))
    emision = f"{hoy.day} septiembre {hoy.year}" if hoy.month == 9 else hoy.strftime("%d/%m/%Y")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHARTS.mkdir(parents=True, exist_ok=True)

    nodos = nodos_todos()
    if args.skip_refresh and JSON_DAILY.is_file():
        by = cargar_diario()
    else:
        by = refrescar_diario(nodos, desde, hasta)
    if args.skip_refresh and JSON_NIGHT.is_file():
        n06 = json.loads(JSON_NIGHT.read_text(encoding="utf-8")).get("hourly") or {}
    else:
        n06 = refrescar_noches(NODOS_CLAVE_NOCHE, ante0, hasta)
    if args.skip_refresh and JSON_CONN.is_file():
        conn = json.loads(JSON_CONN.read_text(encoding="utf-8"))
    else:
        conn = refrescar_conexion(nodos)

    ppt = build_ppt(by, n06, conn, desde, hasta, ante0, s_prev, s_act, emision)
    print("\n=== SALIDA ===")
    print(ppt)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
