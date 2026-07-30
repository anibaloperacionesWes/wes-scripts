"""
Auditoría CPA (equipo hídrico WES): comparación de consumo horario en jornada completa.

Colegio ICCO Renca (carpeta «Auditoria ICCO abril») = nodo **000017-08**. En API/config el punto puede figurar con otro nombre; gráficos/CSV/informe usan «Colegio ICCO Renca».

Periodos (hora Chile, jornada [00:00, 24:00) = horas 0..23 en datos horarios):
  - Con control (Con WES): 13-04-2026 a 19-04-2026 (7 días; total_ref en API interna)
  - Sin control (Sin WES): 06-04-2026 a 12-04-2026 (7 días; total_aud en API interna)

Genera PNG (barras + perfiles 24 h) y CSV de totales y ahorro en reports/auditoria_cpa_icco/
(no se genera el PNG de líneas comparativas de serie larga; la serie sigue en el CSV).

PNG 24 h «Con WES vs línea base (Sin WES)»: un archivo por día homólogo y uno de promedio (m³/h, hora Chile / fusión CSV como la app).

Uso:
  python auditoria_cpa_icco_renca_grafico.py
  python auditoria_cpa_icco_renca_grafico.py --node-id 000017-08
  python auditoria_cpa_icco_renca_grafico.py --prueba-martes-homologo
    → 3 PNG: solo 24-03 Con WES, solo 07-04 Sin WES, doble panel; CSV por API en csv_api_descarga/.
  python auditoria_cpa_icco_renca_grafico.py --prueba-martes-homologo --prueba-superpuesto
    → además un PNG con ambas curvas en el mismo eje (solo comparación visual).
"""

from __future__ import annotations

import argparse
import csv
import os
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
import numpy as np

from generar_reporte_word import (
    BASE_URL,
    _dt_to_chile,
    _requests_session,
    _utc_calendar_dates_for_chile_day,
    get_hourly_measures_for_day,
)

NODE_DEFAULT = "000017-08"
# Texto en gráficos, CSV e informe (permanente; no usar nombre histórico API tipo «Esc. Lo Velásquez»).
NOMBRE_PUNTO = "Colegio ICCO Renca"

# Jornada completa: horas Chile 0..23 (N días × 24 h puntos en la serie comparativa).
HORA_INICIO = 0
HORA_FIN_EXCL = 24


@dataclass(frozen=True)
class Periodo:
    nombre: str
    dias: Tuple[date, ...]


# Periodos fijos auditoría CPA ICCO Renca (2026): 7+7 días (Excel / referencia consumos).
# total_ref = PERIODO_REFERENCIA = periodo **con control** (Con WES).
# total_aud = PERIODO_AUDITORIA = periodo **sin control** (línea base).
# Ahorro (m³) = total_ref − total_aud = con − sin. Rendimiento % sobre periodo con control.
PERIODO_REFERENCIA = Periodo(
    "Con control (13-04 a 19-04-2026)",
    tuple(date(2026, 4, d) for d in range(13, 20)),
)
PERIODO_AUDITORIA = Periodo(
    "Sin control (06-04 a 12-04-2026)",
    tuple(date(2026, 4, d) for d in range(6, 13)),
)


@dataclass(frozen=True)
class ComparacionDia24h:
    """Totales diarios (suma de 24 valores m³/h ≈ m³/día) para el par homólogo con vs sin control."""

    nombre_dia: str
    fecha_con: date
    fecha_sin: date
    total_con_m3: float
    total_sin_m3: float


@dataclass
class ResultadoAuditoriaCpa:
    node_id: str
    etiquetas: List[str]
    y_referencia: List[float]
    y_auditoria: List[float]
    total_ref_m3: float
    total_aud_m3: float
    ahorro_m3: float
    rendimiento_pct: float
    png_path: Optional[Path] = None
    csv_path: Optional[Path] = None
    png_barras_path: Optional[Path] = None
    png_paths_24h: Optional[List[Path]] = None
    comparaciones_diarias_24h: Optional[List[ComparacionDia24h]] = None
    # Vectores 24 h usados para totales y gráficos (misma lectura por día; sin triple llamada API).
    vecs_ref: Optional[List[List[float]]] = None
    vecs_aud: Optional[List[List[float]]] = None


def list_comparaciones_diarias_24h(
    node_id: str,
    ref: Periodo = PERIODO_REFERENCIA,
    aud: Periodo = PERIODO_AUDITORIA,
) -> List[ComparacionDia24h]:
    """
    Un registro por cada día homólogo: total del día sumando horas 0–23 (m³/h → m³).
    """
    if len(ref.dias) != len(aud.dias):
        raise ValueError("Periodos con distinto número de días.")
    vecs_con = _vectores_m3h_por_dias(node_id, ref.dias)
    vecs_sin = _vectores_m3h_por_dias(node_id, aud.dias)
    return _comparaciones_desde_vectores(ref, aud, vecs_con, vecs_sin)


_NOMBRE_DIA_SEMANA_ES: Tuple[str, ...] = (
    "Lunes",
    "Martes",
    "Miércoles",
    "Jueves",
    "Viernes",
    "Sábado",
    "Domingo",
)


def _nombre_dia_semana_es(d: date) -> str:
    """Nombre del día según calendario Chile (lunes=0)."""
    return _NOMBRE_DIA_SEMANA_ES[d.weekday()]


def _vectores_m3h_por_dias(node_id: str, dias: Sequence[date]) -> List[List[float]]:
    """Un vector 0..23 por cada fecha (una sola lectura API por día)."""
    return [_vector_m3h_24_desde_api(node_id, d) for d in dias]


def _flatten_grilla_desde_vectores(
    dias: Sequence[date],
    vecs: Sequence[Sequence[float]],
) -> Tuple[List[str], List[float], float]:
    """Misma grilla que antes (_serie_grilla), pero sin volver a llamar a la API."""
    if len(dias) != len(vecs):
        raise ValueError("Cantidad de fechas y vectores no coincide.")
    labels: List[str] = []
    valores: List[float] = []
    total = 0.0
    for i, dia in enumerate(dias, start=1):
        vec = vecs[i - 1]
        if len(vec) != HORA_FIN_EXCL - HORA_INICIO:
            raise ValueError(f"Vector horario debe tener 24 valores; día {dia}.")
        for h in range(HORA_INICIO, HORA_FIN_EXCL):
            v = float(vec[h])
            labels.append(f"D{i} {h:02d}:00\n{dia.strftime('%d-%m-%y')}")
            valores.append(v)
            total += v
    return labels, valores, total


def _comparaciones_desde_vectores(
    ref: Periodo,
    aud: Periodo,
    vecs_con: Sequence[Sequence[float]],
    vecs_sin: Sequence[Sequence[float]],
) -> List[ComparacionDia24h]:
    if len(ref.dias) != len(aud.dias) or len(vecs_con) != len(ref.dias) or len(vecs_sin) != len(aud.dias):
        raise ValueError("Periodos o vectores de distinta longitud.")
    out: List[ComparacionDia24h] = []
    for d_con, d_sin, vc, vs in zip(ref.dias, aud.dias, vecs_con, vecs_sin):
        out.append(
            ComparacionDia24h(
                nombre_dia=_nombre_dia_semana_es(d_con),
                fecha_con=d_con,
                fecha_sin=d_sin,
                total_con_m3=float(sum(vc)),
                total_sin_m3=float(sum(vs)),
            )
        )
    return out


def _vector_m3h_24_desde_api(node_id: str, dia: date) -> List[float]:
    """Un valor por hora Chile 0..23 (m³/h); 0 si no hay dato (vía ``get_hourly_measures_for_day``)."""
    target = datetime.combine(dia, datetime.min.time())
    hourly_list = get_hourly_measures_for_day(node_id, target) or []
    por_hora: Dict[int, float] = defaultdict(float)
    for h, v in hourly_list:
        hi = int(h)
        if 0 <= hi < 24:
            por_hora[hi] += float(v)
    return [float(por_hora.get(h, 0.0)) for h in range(24)]


def descargar_csv_dia_chile_api(node_id: str, dia_chile: date, dest_subdir: Path) -> List[Path]:
    """
    Descarga ``GET .../nodes/{id}/dates.measures.csv`` por cada día UTC que cubre ``dia_chile``
    (misma fuente que la app). Guarda un archivo por fecha UTC en ``dest_subdir``.
    """
    dest_subdir.mkdir(parents=True, exist_ok=True)
    url = f"{BASE_URL}/nodes/{node_id}/dates.measures.csv"
    sess = _requests_session()
    out: List[Path] = []
    for ud in _utc_calendar_dates_for_chile_day(dia_chile):
        date_str = ud.strftime("%d%m%Y")
        r = sess.get(url, params=[("start", date_str), ("end", date_str)], timeout=60)
        r.raise_for_status()
        fn = dest_subdir / f"{node_id}_dia{dia_chile:%Y%m%d}_utc{ud:%Y%m%d}.csv"
        fn.write_text(r.text, encoding="utf-8")
        out.append(fn)
    return out


def _parse_float_valor_csv(value_str: str) -> float:
    s = value_str.strip().replace(" ", "").replace(",", ".")
    return float(s)


def _vector_m3h_desde_archivos_csv_guardados(paths: Sequence[Path], dia_chile: date) -> List[float]:
    """
    Reconstruye hora Chile 0..23 leyendo los CSV guardados (misma regla que ``get_hourly_measures_for_day``).

    Cada fila ``TIME,VALUE`` es un instante **inicio de hora en UTC** (``T00Z``…``T23Z``); el valor es el
    caudal medio (m³/h) de ese intervalo horario en UTC, reasignado al **inicio de la misma hora civil
    Chile** que contiene ese instante. Si hay más de una fila para la misma hora Chile, **gana la última**
    leída (no se suman duplicados).
    """
    acc: Dict[int, float] = {}
    for path in paths:
        text = path.read_text(encoding="utf-8", errors="replace")
        for line in text.strip().split("\n")[1:]:
            if not line.strip():
                continue
            parts = line.split(",", 1)
            if len(parts) < 2:
                continue
            try:
                time_str = parts[0].strip()
                value_str = parts[1].strip()
                ts_norm = time_str.strip().replace("Z", "+00:00")
                dt = datetime.fromisoformat(ts_norm)
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=timezone.utc)
                ch = _dt_to_chile(dt)
                if ch.date() != dia_chile:
                    continue
                hi = int(ch.hour)
                if 0 <= hi < 24:
                    acc[hi] = _parse_float_valor_csv(value_str)
            except (ValueError, TypeError, IndexError):
                continue
    return [float(acc.get(h, 0.0)) for h in range(24)]


def _vector_m3h_filas_orden_primer_csv_utc(
    paths: Sequence[Path], dia_chile: date
) -> Tuple[List[float], Optional[date]]:
    """
    Las 24 filas **en orden de archivo** del CSV del **primer** día UTC que cubre ``dia_chile``
    (mismo criterio que ``_utc_calendar_dates_for_chile_day``[0]).

    Muchas vistas móviles muestran primero este bloque ``T00Z…T23Z``; el índice 0 no es medianoche Chile.
    """
    uds = _utc_calendar_dates_for_chile_day(dia_chile)
    if not uds:
        return [0.0] * 24, None
    ud0 = uds[0]
    target_suffix = f"_utc{ud0:%Y%m%d}.csv"
    chosen: Optional[Path] = None
    for p in paths:
        if p.name.endswith(target_suffix):
            chosen = p
            break
    if chosen is None:
        return [0.0] * 24, ud0
    lines = chosen.read_text(encoding="utf-8", errors="replace").strip().split("\n")[1:]
    out: List[float] = []
    for line in lines:
        if not line.strip():
            continue
        parts = line.split(",", 1)
        if len(parts) < 2:
            continue
        try:
            out.append(_parse_float_valor_csv(parts[1]))
        except (ValueError, TypeError):
            out.append(0.0)
        if len(out) >= 24:
            break
    while len(out) < 24:
        out.append(0.0)
    return out[:24], ud0


def _texto_resumen_serie_dia(v: Sequence[float], d: date) -> str:
    yv = [float(x) for x in v]
    s = sum(yv)
    eps = 1e-6
    idx = next((i for i, x in enumerate(yv) if x > eps), None)
    if idx is None:
        return f"{d:%d-%m-%Y}: Σ≈{s:.2f} m³ · sin tramos con flujo>0 en esta API."
    return (
        f"{d:%d-%m-%Y}: Σ≈{s:.2f} m³ · 1.er tramo con flujo: {idx:02d}:00–{(idx+1)%24:02d}:00 "
        f"(≈{yv[idx]:.2f} m³/h)"
    )


def _vector_m3h_24_por_hora_utc_primer_archivo(node_id: str, dia_chile: date) -> List[float]:
    """
    Serie **24 puntos** hora civil Chile 0..23, **misma lógica que la app** y que el informe.

    Antes solo se leía el **primer** CSV de un día UTC y se agrupaba por hora UTC; un día civil Chile
    suele abarcar **dos** fechas UTC: faltaban datos y aparecían **ceros** erróneos (p. ej. Sin WES
    07-04-2026 con consumo real en la API).

    Ahora se reutiliza ``_vector_m3h_24_desde_api`` / ``get_hourly_measures_for_day`` (fusión de todos
    los CSV UTC que cubren el día Chile). El nombre del símbolo se mantiene por compatibilidad con
    ``WES_AUDITORIA_GRAFICO_EJE=utc_csv``.
    """
    return _vector_m3h_24_desde_api(node_id, dia_chile)


def _nice_ymax_consumo_m3h(need: float) -> float:
    """Tope Y legible (p. ej. 6 cuando el pico ~5), alineado con escalas tipo app WES."""
    if need <= 0:
        return 1.0
    candidatos = (
        1,
        2,
        3,
        4,
        5,
        6,
        8,
        10,
        12,
        15,
        18,
        20,
        25,
        30,
        40,
        50,
        60,
        80,
        100,
    )
    for c in candidatos:
        if float(c) >= need:
            return float(c)
    return float(np.ceil(need / 10.0) * 10.0)


def _ylim_superior(y_con: Sequence[float], y_sin: Sequence[float]) -> float:
    vals = [float(x) for x in y_con] + [float(x) for x in y_sin]
    mx = max(vals) if vals else 0.0
    if mx <= 0:
        return 1.0
    # Margen más ajustado que 1.18; el tope se redondea a escala “limpia” (0–6, 0–8, etc.).
    need = mx * 1.08
    return _nice_ymax_consumo_m3h(need)


def _ylim_superior_una_serie(y: Sequence[float]) -> float:
    vals = [float(x) for x in y]
    mx = max(vals) if vals else 0.0
    if mx <= 0:
        return 1.0
    return _nice_ymax_consumo_m3h(mx * 1.08)


def _aplicar_grilla_y_24h(ax, ymax: float) -> None:
    if ymax <= 15:
        ax.yaxis.set_major_locator(MultipleLocator(2))
        ax.yaxis.set_minor_locator(MultipleLocator(0.2))
    ax.grid(axis="y", which="major", linestyle="-", linewidth=0.85, alpha=0.5, color="0.78")
    ax.grid(axis="y", which="minor", linestyle="-", linewidth=0.45, alpha=0.28, color="0.85")
    ax.set_axisbelow(True)


def _fig_perfil_24h_un_dia(
    y: Sequence[float],
    *,
    titulo: str,
    etiqueta_serie: str,
    color_linea: str,
    color_fill: Tuple[float, float, float],
    eje_horario: str = "chile",
) -> plt.Figure:
    """Un solo día, una curva: escala Y ajustada solo a ese día."""
    x = np.arange(24, dtype=float)
    yv = np.round(np.asarray(y, dtype=float), 1)
    ymax = _ylim_superior_una_serie(yv)

    fig, ax = plt.subplots(figsize=(10.2, 4.6))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    ax.fill_between(x, 0, yv, color=color_fill + (0.42,), zorder=1)
    ax.plot(x, yv, color=color_linea, linewidth=2.0, zorder=3, label=etiqueta_serie)
    ax.set_xlim(-0.5, 23.5)
    ax.set_xticks(np.arange(24))
    ax.set_xticklabels([f"{h}:00" for h in range(24)], fontsize=7, rotation=45, ha="right")
    if eje_horario == "utc_csv":
        ax.set_xlabel(
            "Inicio de hora — Chile (America/Santiago); datos desde CSV/API fusionado",
            fontsize=8,
            color="#333333",
            labelpad=5,
        )
    else:
        ax.set_xlabel(
            "Inicio de hora — Chile (America/Santiago)",
            fontsize=9,
            color="#333333",
            labelpad=5,
        )
    ax.set_ylim(0, ymax)
    ax.set_ylabel("Consumo (m³/h)", fontsize=11)
    ax.set_title(titulo, fontsize=12, fontweight="bold", color="black", pad=8)
    _aplicar_grilla_y_24h(ax, ymax)
    ax.legend(loc="upper right", fontsize=9, frameon=True, framealpha=0.92)
    for spine in ax.spines.values():
        spine.set_edgecolor("#000000")
        spine.set_linewidth(0.9)
    fig.subplots_adjust(bottom=0.22, left=0.09, right=0.98, top=0.90)
    return fig


def _fig_prueba_dos_dias_apilados(
    y_con: Sequence[float],
    y_sin: Sequence[float],
    *,
    d_con: date,
    d_sin: date,
    eje_horario: str = "chile",
) -> plt.Figure:
    """Dos paneles (Con WES arriba, Sin WES abajo); cada uno con su propio tope en Y."""
    x = np.arange(24, dtype=float)
    yc = np.round(np.asarray(y_con, dtype=float), 1)
    ys = np.round(np.asarray(y_sin, dtype=float), 1)
    ymax_c = _ylim_superior_una_serie(yc)
    ymax_s = _ylim_superior_una_serie(ys)

    fig, axes = plt.subplots(2, 1, figsize=(10.2, 8.8), sharex=True)
    fig.patch.set_facecolor("white")
    fig.suptitle(
        f"{NOMBRE_PUNTO} — comparación homóloga (un día por panel)\n"
        f"Con WES (arriba): {d_con:%d-%m-%Y}  ·  Sin WES (abajo): {d_sin:%d-%m-%Y}",
        fontsize=12,
        fontweight="bold",
        y=0.98,
    )

    for ax, yv, ymax, lab, fc, lc, d_i, resumen in (
        (
            axes[0],
            yc,
            ymax_c,
            "Con WES",
            (0.62, 0.82, 0.98),
            "#2e7ac8",
            d_con,
            _texto_resumen_serie_dia(yc, d_con),
        ),
        (
            axes[1],
            ys,
            ymax_s,
            "Sin WES",
            (0.78, 0.22, 0.22),
            "#c41e1e",
            d_sin,
            _texto_resumen_serie_dia(ys, d_sin),
        ),
    ):
        ax.set_facecolor("white")
        ax.fill_between(x, 0, yv, color=fc + (0.42,), zorder=1)
        ax.plot(x, yv, color=lc, linewidth=2.0, zorder=3, label=lab)
        ax.set_ylim(0, ymax)
        ax.set_ylabel("m³/h", fontsize=10)
        _aplicar_grilla_y_24h(ax, ymax)
        ax.legend(loc="upper right", fontsize=9, frameon=True, framealpha=0.92)
        ax.text(
            0.02,
            0.04,
            resumen,
            transform=ax.transAxes,
            fontsize=7.5,
            color="#333333",
            verticalalignment="bottom",
            bbox=dict(boxstyle="round,pad=0.35", facecolor="white", edgecolor="#bbbbbb", alpha=0.92),
        )
        for spine in ax.spines.values():
            spine.set_edgecolor("#000000")
            spine.set_linewidth(0.9)

    axes[0].set_xticklabels([])
    axes[1].set_xticks(np.arange(24))
    axes[1].set_xticklabels([f"{h}:00" for h in range(24)], fontsize=7, rotation=45, ha="right")
    if eje_horario == "utc_csv":
        axes[1].set_xlabel(
            "Inicio de hora civil Chile (cada punto es el intervalo [h, h+1) en reloj local)",
            fontsize=8,
            color="#333333",
            labelpad=6,
        )
    else:
        axes[1].set_xlabel(
            "Inicio de hora civil Chile (cada punto es el intervalo [h, h+1) en reloj local)",
            fontsize=9,
            color="#333333",
            labelpad=6,
        )
    fig.text(
        0.5,
        0.012,
        "Datos: fusión de CSV por API; cada fila del CSV es TxxZ (inicio de hora UTC). "
        "Si la app muestra otra forma de noche, suele ser otro backend o la vista en orden T00Z…T23Z "
        "(PNG adicional con --prueba-referencia-filas-utc, activo por defecto).",
        ha="center",
        fontsize=7,
        color="#444444",
    )
    fig.subplots_adjust(bottom=0.16, left=0.09, right=0.98, top=0.90, hspace=0.18)
    return fig


def _fig_referencia_filas_utc_doble(
    y_con: Sequence[float],
    y_sin: Sequence[float],
    *,
    d_con: date,
    d_sin: date,
    ud_con: Optional[date],
    ud_sin: Optional[date],
) -> plt.Figure:
    """
    Misma forma que muchos CSV: 24 valores en orden ``T00Z``…``T23Z`` del **primer** día UTC del día Chile.
    El índice 0 **no** es medianoche Chile (suele ser la tarde/noche anterior en reloj local).
    """
    x = np.arange(24, dtype=float)
    yc = np.round(np.asarray(y_con, dtype=float), 1)
    ys = np.round(np.asarray(y_sin, dtype=float), 1)
    ymax_c = _ylim_superior_una_serie(yc)
    ymax_s = _ylim_superior_una_serie(ys)

    fig, axes = plt.subplots(2, 1, figsize=(10.2, 8.8), sharex=True)
    fig.patch.set_facecolor("white")
    t_ud = "fecha UTC del archivo"
    fig.suptitle(
        f"{NOMBRE_PUNTO} — referencia: orden de filas del CSV (T00Z→T23Z)\n"
        f"Con: {d_con:%d-%m-%Y} ({t_ud} {ud_con or '?'} · arriba)  ·  "
        f"Sin: {d_sin:%d-%m-%Y} ({t_ud} {ud_sin or '?'} · abajo)",
        fontsize=11,
        fontweight="bold",
        y=0.98,
    )

    for ax, yv, ymax, lab, fc, lc, ud in (
        (axes[0], yc, ymax_c, "Con WES", (0.62, 0.82, 0.98), "#2e7ac8", ud_con),
        (axes[1], ys, ymax_s, "Sin WES", (0.78, 0.22, 0.22), "#c41e1e", ud_sin),
    ):
        ax.set_facecolor("white")
        ax.fill_between(x, 0, yv, color=fc + (0.42,), zorder=1)
        ax.plot(x, yv, color=lc, linewidth=2.0, zorder=3, label=lab)
        ax.set_ylim(0, ymax)
        ax.set_ylabel("m³/h", fontsize=10)
        _aplicar_grilla_y_24h(ax, ymax)
        ax.legend(loc="upper right", fontsize=9, frameon=True, framealpha=0.92)
        ax.text(
            0.02,
            0.04,
            "Índices 0–7 = filas T00Z…T07Z (suelen ser ~0 en API). No confundir con 00:00–07:00 Chile.",
            transform=ax.transAxes,
            fontsize=7.5,
            color="#333333",
            verticalalignment="bottom",
            bbox=dict(boxstyle="round,pad=0.35", facecolor="#fff9e6", edgecolor="#ddaa33", alpha=0.95),
        )
        for spine in ax.spines.values():
            spine.set_edgecolor("#000000")
            spine.set_linewidth(0.9)

    axes[0].set_xticklabels([])
    axes[1].set_xticks(np.arange(24))
    axes[1].set_xticklabels([f"{k}" for k in range(24)], fontsize=7)
    axes[1].set_xlabel("Índice de fila en el CSV UTC (0 = T00Z … 23 = T23Z de ese archivo)", fontsize=8, color="#333333", labelpad=6)
    fig.text(
        0.5,
        0.012,
        "Este gráfico sirve para cotejar con pantallas que recorren el bloque horario UTC del día sin pasar antes por hora civil 0–23.",
        ha="center",
        fontsize=7,
        color="#444444",
    )
    fig.subplots_adjust(bottom=0.14, left=0.09, right=0.98, top=0.88, hspace=0.18)
    return fig


def _fig_wes_vs_linea_base_24h(
    y_con: Sequence[float],
    y_sin: Sequence[float],
    titulo: str,
    *,
    eje_horario: str = "chile",
) -> plt.Figure:
    """
    Estilo combinado: áreas semitransparentes + línea roja «Sin WES»; eje X 0–23.
    La API entrega hasta ~3 decimales; en el gráfico se usa **1 decimal** para lectura clara.

    ``eje_horario``:
    - ``chile``: hora civil Chile 0..23 (totales del informe).
    - ``utc_csv``: **mismos valores y eje** que ``chile``; el nombre alude al modo «como la app»
      (CSV fusionado por día civil). Se conserva para ``WES_AUDITORIA_GRAFICO_EJE=utc_csv``.
    """
    x = np.arange(24, dtype=float)
    y_con = np.round(np.asarray(y_con, dtype=float), 1)
    y_sin = np.round(np.asarray(y_sin, dtype=float), 1)

    fig, ax = plt.subplots(figsize=(10.2, 4.9))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    # Orden: primero línea base (rojo), encima Con WES (azul)
    ax.fill_between(
        x,
        0,
        y_sin,
        color=(0.78, 0.22, 0.22),
        alpha=0.38,
        label="Sin WES",
        zorder=1,
    )
    ax.fill_between(
        x,
        0,
        y_con,
        color=(0.62, 0.82, 0.98),
        alpha=0.45,
        label="Con WES",
        zorder=2,
    )
    ax.plot(x, y_sin, color="#c41e1e", linewidth=2.1, zorder=4)
    ax.plot(x, y_con, color="#2e7ac8", linewidth=1.35, alpha=0.85, zorder=3)

    ymax = _ylim_superior(y_con, y_sin)
    ax.set_xlim(-0.5, 23.5)
    ax.set_xticks(np.arange(24))
    ax.set_xticklabels([f"{h}:00" for h in range(24)], fontsize=7, rotation=45, ha="right")
    if eje_horario == "utc_csv":
        ax.set_xlabel(
            "Inicio de hora — Chile (America/Santiago); serie = fusión CSV/API (como la app WES)",
            fontsize=8,
            color="#333333",
            labelpad=6,
        )
    else:
        ax.set_xlabel(
            "Inicio de hora — Chile (America/Santiago)",
            fontsize=9,
            color="#333333",
            labelpad=6,
        )
    ax.set_ylim(0, ymax)
    ax.set_ylabel("Consumo (m³/h)", fontsize=11)
    ax.set_title(titulo, fontsize=12, fontweight="bold", color="black", pad=8)
    # Rejilla tipo informe: mayor cada 2 unidades y fina cada 0,2 (si el tope es moderado).
    if ymax <= 15:
        ax.yaxis.set_major_locator(MultipleLocator(2))
        ax.yaxis.set_minor_locator(MultipleLocator(0.2))
    ax.grid(axis="y", which="major", linestyle="-", linewidth=0.85, alpha=0.5, color="0.78")
    ax.grid(axis="y", which="minor", linestyle="-", linewidth=0.45, alpha=0.28, color="0.85")
    ax.set_axisbelow(True)
    ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.2),
        ncol=2,
        fontsize=9,
        frameon=False,
        columnspacing=1.4,
    )
    for spine in ax.spines.values():
        spine.set_edgecolor("#000000")
        spine.set_linewidth(0.9)
    fig.subplots_adjust(bottom=0.30, left=0.09, right=0.98, top=0.88)
    return fig


_SLUG_DIA = ("lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo")


def generar_png_wes_24h_lun_mie_promedio(
    node_id: str,
    ref: Periodo,
    aud: Periodo,
    out_dir: Path,
    ts: str,
    *,
    vecs_con: Optional[Sequence[Sequence[float]]] = None,
    vecs_sin: Optional[Sequence[Sequence[float]]] = None,
    series_eje: str = "chile",
) -> List[Path]:
    """
    Un PNG por cada día homólogo (ref vs aud) y uno de promedio horario sobre esos días.
    Si se pasan ``vecs_con`` / ``vecs_sin`` (mismos que en ``computar_auditoria_cpa``), no se vuelve a consultar la API.

    ``series_eje``: ``chile`` (por defecto) o ``utc_csv`` (misma serie hora Chile; nombre histórico).
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    paths: List[Path] = []
    nd = len(ref.dias)
    if nd != len(aud.dias) or nd < 1:
        return paths

    usar_precomputados = (
        series_eje == "chile"
        and vecs_con is not None
        and vecs_sin is not None
        and len(vecs_con) == nd
        and len(vecs_sin) == nd
    )
    if usar_precomputados:
        v_con_list = [list(map(float, v)) for v in vecs_con]
        v_sin_list = [list(map(float, v)) for v in vecs_sin]
    else:
        v_con_list = []
        v_sin_list = []
        for i in range(nd):
            if series_eje == "utc_csv":
                v_con_list.append(
                    _vector_m3h_24_por_hora_utc_primer_archivo(node_id, ref.dias[i])
                )
                v_sin_list.append(
                    _vector_m3h_24_por_hora_utc_primer_archivo(node_id, aud.dias[i])
                )
            else:
                v_con_list.append(_vector_m3h_24_desde_api(node_id, ref.dias[i]))
                v_sin_list.append(_vector_m3h_24_desde_api(node_id, aud.dias[i]))

    eje_fig = "utc_csv" if series_eje == "utc_csv" else "chile"

    for i in range(nd):
        d_con = ref.dias[i]
        d_sin = aud.dias[i]
        vc = v_con_list[i]
        vs = v_sin_list[i]
        titulo_dia = _nombre_dia_semana_es(d_con)
        titulo = (
            f"Consumos con WES vs línea base — {titulo_dia}\n"
            f"Con WES: {d_con:%d-%m-%Y}  ·  Sin WES (línea base): {d_sin:%d-%m-%Y}"
        )
        fig = _fig_wes_vs_linea_base_24h(vc, vs, titulo, eje_horario=eje_fig)
        slug = _SLUG_DIA[d_con.weekday()]
        path = out_dir / f"cpa_icco_renca_wes_24h_{slug}_{ts}.png"
        fig.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        paths.append(path)

    denom = float(nd)
    prom_con = [sum(z) / denom for z in zip(*v_con_list)]
    prom_sin = [sum(z) / denom for z in zip(*v_sin_list)]
    subt_prom = (
        "Promedio de los pares homólogos · fusión CSV / hora Chile (como app)"
        if eje_fig == "utc_csv"
        else "Promedio de los pares homólogos · Hora Chile"
    )
    fig_p = _fig_wes_vs_linea_base_24h(
        prom_con,
        prom_sin,
        f"Consumos con WES vs línea base — Promedio horario ({nd} días)\n" + subt_prom,
        eje_horario=eje_fig,
    )
    path_p = out_dir / f"cpa_icco_renca_wes_24h_promedio_{ts}.png"
    fig_p.savefig(path_p, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig_p)
    paths.append(path_p)
    return paths


def computar_auditoria_cpa(
    node_id: str,
    ref: Periodo = PERIODO_REFERENCIA,
    aud: Periodo = PERIODO_AUDITORIA,
) -> ResultadoAuditoriaCpa:
    vecs_ref = _vectores_m3h_por_dias(node_id, ref.dias)
    vecs_aud = _vectores_m3h_por_dias(node_id, aud.dias)
    x_ref, y_ref, tot_ref = _flatten_grilla_desde_vectores(ref.dias, vecs_ref)
    x_aud, y_aud, tot_aud = _flatten_grilla_desde_vectores(aud.dias, vecs_aud)
    if len(x_ref) != len(x_aud):
        raise ValueError("Series de distinta longitud; revisar grillas de periodos.")
    ahorro_m3 = tot_ref - tot_aud
    rendimiento_pct = (100.0 * ahorro_m3 / tot_ref) if tot_ref > 0 else 0.0
    return ResultadoAuditoriaCpa(
        node_id=node_id,
        etiquetas=x_ref,
        y_referencia=y_ref,
        y_auditoria=y_aud,
        total_ref_m3=tot_ref,
        total_aud_m3=tot_aud,
        ahorro_m3=ahorro_m3,
        rendimiento_pct=rendimiento_pct,
        vecs_ref=vecs_ref,
        vecs_aud=vecs_aud,
    )


def generar_png_barras_con_sin(
    tot_con: float,
    tot_sin: float,
    out_dir: Path,
    ts: str,
    *,
    ref: Periodo = PERIODO_REFERENCIA,
    aud: Periodo = PERIODO_AUDITORIA,
) -> Path:
    """Gráfico de barras comparando consumo acumulado con control vs sin control (informe auditoría)."""
    fig, ax = plt.subplots(figsize=(7.5, 4.2))
    fig.patch.set_facecolor("white")
    d0, d1 = ref.dias[0], ref.dias[-1]
    a0, a1 = aud.dias[0], aud.dias[-1]
    etiquetas = [
        f"Con control\n{d0:%d-%m} al {d1:%d-%m-%Y}",
        f"Sin control\n{a0:%d-%m} al {a1:%d-%m-%Y}",
    ]
    valores = [tot_con, tot_sin]
    colores = ["#4F81BD", "#C0504D"]
    x_pos = [0, 1]
    bars = ax.bar(
        x_pos,
        valores,
        color=colores,
        width=0.52,
        edgecolor="#444444",
        linewidth=0.8,
    )
    ax.set_xticks(x_pos)
    ax.set_xticklabels(etiquetas, fontsize=9)
    ax.set_ylabel("Consumo acumulado (m³)", fontsize=11)
    ax.set_title(f"Comparación consumo acumulado — {NOMBRE_PUNTO}", fontsize=12)
    ax.grid(axis="y", alpha=0.35, linestyle="--", linewidth=0.8)
    ax.set_axisbelow(True)
    mx = max(valores) if valores else 0.0
    top = mx * 1.25 if mx > 0 else 1.0
    ax.set_ylim(0, top)
    for bar, val in zip(bars, valores):
        ax.text(
            bar.get_x() + bar.get_width() / 2.0,
            bar.get_height() + top * 0.015,
            f"{round(val, 1):.1f} m³",
            ha="center",
            va="bottom",
            fontsize=10,
            weight="bold",
        )
    plt.tight_layout()
    path = out_dir / f"cpa_icco_renca_barras_{ts}.png"
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def _guardar_png_csv(
    res: ResultadoAuditoriaCpa,
    ref: Periodo,
    aud: Periodo,
    out_dir: Path,
    ts: str,
    *,
    graficos_24h: bool = True,
) -> ResultadoAuditoriaCpa:
    node_id = res.node_id
    n = len(res.etiquetas)
    tot_ref, tot_aud = res.total_ref_m3, res.total_aud_m3
    ahorro_m3, rendimiento_pct = res.ahorro_m3, res.rendimiento_pct
    y_ref, y_aud = res.y_referencia, res.y_auditoria
    x_ref = res.etiquetas

    # No se genera el PNG de líneas comparativas (serie larga); los perfiles 24 h sustituyen la lectura visual.

    barras_path = generar_png_barras_con_sin(tot_ref, tot_aud, out_dir, ts, ref=ref, aud=aud)

    if res.vecs_ref is not None and res.vecs_aud is not None:
        comps = _comparaciones_desde_vectores(ref, aud, res.vecs_ref, res.vecs_aud)
    else:
        comps = list_comparaciones_diarias_24h(node_id, ref, aud)
    res.comparaciones_diarias_24h = comps

    csv_path = out_dir / f"cpa_icco_renca_{ts}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["metrica", "valor"])
        w.writerow(["node_id", node_id])
        w.writerow(["punto", NOMBRE_PUNTO])
        w.writerow(
            [
                "ventana_horaria_chile",
                f"{HORA_INICIO:02d}:00 a {HORA_FIN_EXCL:02d}:00 (hora inicio inclusive, fin exclusivo)",
            ]
        )
        w.writerow(["total_m3_referencia", f"{tot_ref:.6f}"])
        w.writerow(["total_m3_auditoria", f"{tot_aud:.6f}"])
        w.writerow(["nota", "total_ref = periodo con control; total_aud = periodo sin control"])
        w.writerow(["ahorro_m3_referencia_menos_auditoria", f"{ahorro_m3:.6f}"])
        w.writerow(["rendimiento_ahorro_pct_sobre_referencia", f"{rendimiento_pct:.4f}"])
        w.writerow([])
        w.writerow(["indice", "etiqueta", "m3_h_referencia", "m3_h_auditoria"])
        for i in range(n):
            w.writerow([i, x_ref[i], y_ref[i], y_aud[i]])
        w.writerow([])
        w.writerow(
            [
                "nota_comparacion_24h",
                "Total diario = suma horas 0-23 (m³/h); aproxima consumo del día en m³.",
            ]
        )
        w.writerow(
            [
                "dia",
                "fecha_con",
                "fecha_sin",
                "total_m3_con",
                "total_m3_sin",
                "diferencia_con_menos_sin",
            ]
        )
        for c in comps:
            dif = c.total_con_m3 - c.total_sin_m3
            w.writerow(
                [
                    c.nombre_dia,
                    c.fecha_con.isoformat(),
                    c.fecha_sin.isoformat(),
                    f"{c.total_con_m3:.6f}",
                    f"{c.total_sin_m3:.6f}",
                    f"{dif:.6f}",
                ]
            )

    res.png_path = None
    res.csv_path = csv_path
    res.png_barras_path = barras_path
    if graficos_24h:
        se = os.environ.get("WES_AUDITORIA_GRAFICO_EJE", "chile").strip().lower()
        if se not in ("chile", "utc_csv"):
            se = "chile"
        res.png_paths_24h = generar_png_wes_24h_lun_mie_promedio(
            node_id,
            ref,
            aud,
            out_dir,
            ts,
            vecs_con=res.vecs_ref if se == "chile" else None,
            vecs_sin=res.vecs_aud if se == "chile" else None,
            series_eje=se,
        )
    return res


# Par homólogo martes para revisión manual del gráfico (no tiene que coincidir con PERIODO_*).
FECHA_PRUEBA_CON_WES_MARTES = date(2026, 3, 24)
FECHA_PRUEBA_SIN_WES_MARTES = date(2026, 4, 7)


def generar_png_prueba_martes_homologo(
    node_id: str = NODE_DEFAULT,
    out_dir: Optional[Path] = None,
    ts: Optional[str] = None,
    *,
    eje: str = "utc_csv",
    descargar_csv_api: bool = True,
    incluir_superpuesto_mismo_eje: bool = False,
    referencia_filas_utc: bool = True,
) -> Tuple[List[Path], Optional[Path]]:
    """
    Tres PNG 24 h para los martes homólogos (24-03 Con WES, 07-04 Sin WES):

    1. **Solo** día Con WES — escala Y solo de ese día.
    2. **Solo** día Sin WES — escala Y solo de ese día.
    3. **Combinado** — dos paneles apilados (arriba Con, abajo Sin), cada uno con su propio tope Y.

    Con ``descargar_csv_api=True``: descarga CSV por API en ``csv_api_descarga/prueba_martes_{ts}/``
    y las series se leen desde disco.

    Con ``incluir_superpuesto_mismo_eje=True``: además genera un cuarto PNG con ambas curvas en el
    mismo gráfico (comportamiento antiguo; puede deformar la lectura si los máximos difieren).

    Retorna ``(lista_png, carpeta_csv_o_None)``. La lista incluye 3 PNG base; opcionalmente el PNG
    de referencia ``T00Z…T23Z``, y otro más si ``incluir_superpuesto_mismo_eje=True``.
    """
    out_dir = out_dir or Path(__file__).resolve().parent / "reports" / "auditoria_cpa_icco"
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = ts or datetime.now().strftime("%Y%m%d_%H%M")
    d_con = FECHA_PRUEBA_CON_WES_MARTES
    d_sin = FECHA_PRUEBA_SIN_WES_MARTES
    csv_run_dir: Optional[Path] = None
    tag_csv = ""

    if descargar_csv_api:
        csv_run_dir = out_dir / "csv_api_descarga" / f"prueba_martes_{ts}"
        pc = descargar_csv_dia_chile_api(node_id, d_con, csv_run_dir / f"con_{d_con:%Y%m%d}")
        ps = descargar_csv_dia_chile_api(node_id, d_sin, csv_run_dir / f"sin_{d_sin:%Y%m%d}")
        vc = _vector_m3h_desde_archivos_csv_guardados(pc, d_con)
        vs = _vector_m3h_desde_archivos_csv_guardados(ps, d_sin)
        tag_csv = "_csvapi"
    elif eje == "utc_csv":
        vc = _vector_m3h_24_por_hora_utc_primer_archivo(node_id, d_con)
        vs = _vector_m3h_24_por_hora_utc_primer_archivo(node_id, d_sin)
    else:
        vc = _vector_m3h_24_desde_api(node_id, d_con)
        vs = _vector_m3h_24_desde_api(node_id, d_sin)

    if eje == "utc_csv":
        eje_fig = "utc_csv"
        suf = "utc_csv"
    else:
        eje_fig = "chile"
        suf = "chile"

    base = f"cpa_icco_prueba_{suf}{tag_csv}_{ts}"

    # 1) Un día por gráfico
    fig1 = _fig_perfil_24h_un_dia(
        vc,
        titulo=f"{NOMBRE_PUNTO} — Con WES · {d_con:%d-%m-%Y} (martes)",
        etiqueta_serie="Con WES",
        color_linea="#2e7ac8",
        color_fill=(0.62, 0.82, 0.98),
        eje_horario=eje_fig,
    )
    p1 = out_dir / f"{base}_solo_con_wes_{d_con:%Y%m%d}.png"
    fig1.savefig(p1, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig1)

    fig2 = _fig_perfil_24h_un_dia(
        vs,
        titulo=f"{NOMBRE_PUNTO} — Sin WES · {d_sin:%d-%m-%Y} (martes)",
        etiqueta_serie="Sin WES",
        color_linea="#c41e1e",
        color_fill=(0.78, 0.22, 0.22),
        eje_horario=eje_fig,
    )
    p2 = out_dir / f"{base}_solo_sin_wes_{d_sin:%Y%m%d}.png"
    fig2.savefig(p2, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig2)

    # 2) Dos paneles unidos (sin compartir escala Y forzada)
    fig3 = _fig_prueba_dos_dias_apilados(vc, vs, d_con=d_con, d_sin=d_sin, eje_horario=eje_fig)
    p3 = out_dir / f"{base}_doble_panel_con_arriba_sin_abajo.png"
    fig3.savefig(p3, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig3)

    out_paths: List[Path] = [p1, p2, p3]

    if descargar_csv_api and referencia_filas_utc:
        vcr, udc = _vector_m3h_filas_orden_primer_csv_utc(pc, d_con)
        vsr, uds = _vector_m3h_filas_orden_primer_csv_utc(ps, d_sin)
        fig_ref = _fig_referencia_filas_utc_doble(
            vcr,
            vsr,
            d_con=d_con,
            d_sin=d_sin,
            ud_con=udc,
            ud_sin=uds,
        )
        p_ref = out_dir / f"{base}_referencia_orden_filas_T00Z_T23Z_utc.png"
        fig_ref.savefig(p_ref, dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig_ref)
        out_paths.append(p_ref)

    if incluir_superpuesto_mismo_eje:
        titulo = (
            f"{NOMBRE_PUNTO} — superpuesto (mismo eje Y)\n"
            f"Con WES (azul): {d_con:%d-%m-%Y}  ·  Sin WES (rojo): {d_sin:%d-%m-%Y}"
        )
        fig4 = _fig_wes_vs_linea_base_24h(vc, vs, titulo, eje_horario=eje_fig)
        p4 = out_dir / f"{base}_superpuesto_mismo_eje.png"
        fig4.savefig(p4, dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig4)
        out_paths.append(p4)

    return out_paths, csv_run_dir


def ejecutar_auditoria_cpa_icco(
    node_id: str = NODE_DEFAULT,
    out_dir: Optional[Path] = None,
    ref: Periodo = PERIODO_REFERENCIA,
    aud: Periodo = PERIODO_AUDITORIA,
    timestamp: Optional[str] = None,
    *,
    graficos_24h: bool = True,
) -> ResultadoAuditoriaCpa:
    """
    Calcula métricas y, si se indica out_dir, escribe PNG + CSV con marca de tiempo.
    Opcionalmente cuatro PNG 24 h Con WES vs Sin WES (``graficos_24h``).
    """
    res = computar_auditoria_cpa(node_id, ref=ref, aud=aud)
    if out_dir is None:
        return res
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = timestamp or datetime.now().strftime("%Y%m%d_%H%M")
    return _guardar_png_csv(res, ref, aud, out_dir, ts, graficos_24h=graficos_24h)


def main() -> int:
    p = argparse.ArgumentParser(description="Gráfico auditoría CPA ICCO Renca")
    p.add_argument("--node-id", default=NODE_DEFAULT, help="ID nodo WES (default ICCO Renca)")
    p.add_argument(
        "--no-graficos-24h",
        action="store_true",
        help="No generar los 4 PNG 24 h (Lun/Mié/promedio Con WES vs Sin WES).",
    )
    p.add_argument(
        "--prueba-martes-homologo",
        action="store_true",
        help="PNG de prueba: un gráfico solo 24-03 Con WES, uno solo 07-04 Sin WES, y uno doble panel.",
    )
    p.add_argument(
        "--prueba-eje-chile",
        action="store_true",
        help="Con --prueba-martes-homologo: etiqueta de eje X 'solo Chile'. Por defecto: texto 'como app'.",
    )
    p.add_argument(
        "--prueba-sin-descarga-csv",
        action="store_true",
        help="Con --prueba-martes-homologo: no guardar CSV; agregar serie solo vía API en memoria.",
    )
    p.add_argument(
        "--prueba-superpuesto",
        action="store_true",
        help="Con --prueba-martes-homologo: además generar PNG con ambas curvas en el mismo eje Y.",
    )
    p.add_argument(
        "--no-prueba-referencia-filas-utc",
        action="store_true",
        help="Con --prueba-martes-homologo: no generar PNG de referencia (orden T00Z a T23Z en CSV UTC).",
    )
    args = p.parse_args()
    node_id = args.node_id

    root = Path(__file__).resolve().parent
    out_dir = root / "reports" / "auditoria_cpa_icco"
    if args.prueba_martes_homologo:
        paths, csv_dir = generar_png_prueba_martes_homologo(
            node_id=node_id,
            out_dir=out_dir,
            eje="chile" if args.prueba_eje_chile else "utc_csv",
            descargar_csv_api=not args.prueba_sin_descarga_csv,
            incluir_superpuesto_mismo_eje=args.prueba_superpuesto,
            referencia_filas_utc=not args.no_prueba_referencia_filas_utc,
        )
        print("PNG de prueba (martes homólogo):")
        for i, pth in enumerate(paths):
            nm = pth.name
            if i == 0:
                lab = "Solo Con WES (24-03)"
            elif i == 1:
                lab = "Solo Sin WES (07-04)"
            elif i == 2:
                lab = "Doble panel (hora civil Chile, fusión CSV)"
            elif "referencia_orden" in nm:
                lab = "Referencia orden filas CSV (T00Z a T23Z, no es eje 0-23 Chile)"
            elif "superpuesto" in nm:
                lab = "Superpuesto (mismo eje Y)"
            else:
                lab = nm
            print(f"  {i + 1}. {lab}: {pth.resolve()}")
        if csv_dir is not None:
            print("CSV descargados (API):")
            print(csv_dir.resolve())
        return 0

    ref, aud = PERIODO_REFERENCIA, PERIODO_AUDITORIA

    res = ejecutar_auditoria_cpa_icco(
        node_id=node_id,
        out_dir=out_dir,
        ref=ref,
        aud=aud,
        graficos_24h=not args.no_graficos_24h,
    )
    tot_ref, tot_aud = res.total_ref_m3, res.total_aud_m3
    ahorro_m3, rendimiento_pct = res.ahorro_m3, res.rendimiento_pct
    csv_path = res.csv_path

    print("=" * 72)
    print("AUDITORÍA CPA — ICCO RENCA")
    print("=" * 72)
    print(f"Nodo: {node_id}")
    print(f"Jornada (Chile): {HORA_INICIO:02d}:00 - {HORA_FIN_EXCL:02d}:00 (24 h/día)")
    print(
        f"Total m3 con control ({ref.dias[0]:%d-%m} a {ref.dias[-1]:%d-%m-%Y}): {tot_ref:.3f}"
    )
    print(
        f"Total m3 sin control ({aud.dias[0]:%d-%m} a {aud.dias[-1]:%d-%m-%Y}): {tot_aud:.3f}"
    )
    print(f"Diferencia (con - sin): {ahorro_m3:.3f} m3  |  Rendimiento sobre con: {rendimiento_pct:.2f}%")
    print(f"PNG (líneas comparativas): (no generado; ver perfiles 24 h y CSV)")
    print(f"PNG (barras): {res.png_barras_path.resolve() if res.png_barras_path else '(no)'}")
    if res.png_paths_24h:
        print("PNG (24 h Con WES vs Sin WES):")
        for pp in res.png_paths_24h:
            print(f"  {pp.resolve()}")
    print(f"CSV: {csv_path.resolve() if csv_path else '(no)'}")
    print()
    print(
        f"Nota: Con control = {ref.dias[0]:%d-%m} a {ref.dias[-1]:%d-%m-%Y}; "
        f"sin control = {aud.dias[0]:%d-%m} a {aud.dias[-1]:%d-%m-%Y}. "
        "Rendimiento % = 100 x (con - sin) / con."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
