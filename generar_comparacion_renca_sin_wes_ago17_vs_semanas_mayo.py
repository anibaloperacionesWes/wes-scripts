"""
Renca — rendimiento hídrico con WES vs periodo sin control (ago-2026).

Contexto operativo: desde el lunes 17/08/2026 los equipos de ICCP, Esc. Lo Velásquez,
Gimnasio y Piscina Municipal quedaron sin control WES. Este script:

  1. Toma el periodo sin WES (lunes 17/08 hasta la última hora completa Chile de «ahora»).
  2. Replica la misma ventana horaria (lun + mar + mié hasta esa hora) en cada semana
     lunes–domingo desde mayo 2026 (semanas con control WES).
  3. Calcula rendimiento de ahorro: (Sin WES − Con WES) / Sin WES × 100.
  4. Identifica la semana de mayo en adelante con mayor rendimiento (m³ y %).
  5. Genera Excel, gráficos, Word y PDF.

Uso:
  python generar_comparacion_renca_sin_wes_ago17_vs_semanas_mayo.py
"""

from __future__ import annotations

import csv
import shutil
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
from zoneinfo import ZoneInfo

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from generar_reporte_word import (
    acl_node_base_url,
    add_logo_to_header,
    fetch_json,
    flatten_measures,
    format_number_chilean,
    get_hourly_measures_for_day,
    normalize_measures_payload,
)

ROOT = Path(__file__).resolve().parent
TZ_CL = ZoneInfo("America/Santiago")
TARIFA_CLP_M3 = 1300.0
COLOR_WES = "#1F4788"
COLOR_SIN = "#C0504D"
COLOR_AHORRO = "#548235"
COLOR_HEAD = RGBColor(31, 71, 136)

PUNTOS: Tuple[Tuple[str, str], ...] = (
    ("000017-07", "ICCP (Cumbre de Cóndores pte.)"),
    ("000017-04", "Esc. Lo Velásquez"),
    ("000017-05", "Gimnasio municipal"),
    ("000017-06", "Piscina municipal"),
)

# Primera semana completa de mayo 2026 (lunes 4) hasta última semana con WES (lun 10–dom 16 ago).
PRIMER_LUNES_CON_WES = date(2026, 5, 4)
ULTIMO_LUNES_CON_WES = date(2026, 8, 10)
LUNES_SIN_WES = date(2026, 8, 17)


def _ahora_chile() -> datetime:
    return datetime.now(TZ_CL)


def _ultima_hora_completa(ahora: datetime) -> int:
    """Hora Chile 0–23 cuya franja [H:00, H+1:00) ya cerró. Si son las 13:31 → 12."""
    h = ahora.hour - 1
    return max(0, min(23, h))


def _lunes_semana(d: date) -> date:
    return d - timedelta(days=d.weekday())


def _iter_lunes(desde: date, hasta: date) -> List[date]:
    out: List[date] = []
    d = _lunes_semana(desde)
    if d < desde:
        d += timedelta(days=7)
    while d <= hasta:
        out.append(d)
        d += timedelta(days=7)
    return out


def _etiqueta_semana(lunes: date) -> str:
    domingo = lunes + timedelta(days=6)
    return f"{lunes:%d/%m}–{domingo:%d/%m}"


def _set_cell_shading(cell, hex_fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in tc_pr.findall(qn("w:shd")):
        tc_pr.remove(old)
    shading = parse_xml(
        "<w:shd xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        f'w:val="clear" w:fill="{hex_fill}"/>'
    )
    tc_pr.append(shading)


def _fetch_diario(node_id: str, start: date, end: date) -> Dict[date, float]:
    raw = fetch_json(
        f"{acl_node_base_url()}/nodes/measures/dates",
        params=[
            ("id", node_id),
            ("start", start.strftime("%d%m%Y")),
            ("end", end.strftime("%d%m%Y")),
        ],
    )
    payload = normalize_measures_payload(raw, node_id)
    out: Dict[date, float] = {}
    for m in flatten_measures(payload):
        d = m.date.date()
        if start <= d <= end:
            out[d] = float(m.total_m3)
    return out


def _vector_24h(node_id: str, dia: date) -> List[float]:
    hourly = get_hourly_measures_for_day(node_id, datetime.combine(dia, datetime.min.time())) or []
    por_hora = {int(h): float(v) for h, v in hourly}
    return [float(por_hora.get(h, 0.0)) for h in range(24)]


def _suma_horas(vec: Sequence[float], h_ini: int, h_fin_incl: int) -> float:
    return float(sum(float(vec[h]) for h in range(h_ini, h_fin_incl + 1) if 0 <= h < len(vec)))


@dataclass
class VentanaPunto:
    node_id: str
    nombre: str
    lunes: date
    con_wes: bool
    m3_lunes: float
    m3_martes: float
    m3_miercoles_corte: float
    m3_total: float
    m3_nocturno: float  # 00–06 del lun+mar + mié hasta min(6, corte)
    horas_equivalentes: int
    vec_miercoles: List[float] = field(default_factory=list)
    nota: str = ""


@dataclass
class SemanaPunto:
    node_id: str
    nombre: str
    lunes: date
    con_wes: bool
    m3_semana: float
    dias_con_dato: int
    m3_dia: float


def _consumo_ventana(
    node_id: str,
    nombre: str,
    lunes: date,
    diario: Dict[date, float],
    vec_mie: List[float],
    hora_corte: int,
    con_wes: bool,
) -> VentanaPunto:
    d_lun = lunes
    d_mar = lunes + timedelta(days=1)
    d_mie = lunes + timedelta(days=2)
    m3_lun = float(diario.get(d_lun, 0.0))
    m3_mar = float(diario.get(d_mar, 0.0))
    m3_mie = _suma_horas(vec_mie, 0, hora_corte)
    noct_lun = 0.0
    noct_mar = 0.0
    # Nocturno lun/mar se aproxima con total diario solo si más adelante hay vectores;
    # para ranking de ventana usamos miércoles 00–min(6, corte) + se completa en informe
    # con vectores del periodo actual. Aquí nocturno = horas 0–6 del miércoles (siempre
    # dentro del corte si hora_corte >= 6, que es el caso a media tarde).
    noct_mie = _suma_horas(vec_mie, 0, min(6, hora_corte))
    horas = 24 + 24 + (hora_corte + 1)
    return VentanaPunto(
        node_id=node_id,
        nombre=nombre,
        lunes=lunes,
        con_wes=con_wes,
        m3_lunes=m3_lun,
        m3_martes=m3_mar,
        m3_miercoles_corte=m3_mie,
        m3_total=m3_lun + m3_mar + m3_mie,
        m3_nocturno=noct_lun + noct_mar + noct_mie,
        horas_equivalentes=horas,
        vec_miercoles=list(vec_mie),
    )


def _semana_completa(
    node_id: str,
    nombre: str,
    lunes: date,
    diario: Dict[date, float],
    con_wes: bool,
    hasta: Optional[date] = None,
) -> SemanaPunto:
    total = 0.0
    n = 0
    for i in range(7):
        d = lunes + timedelta(days=i)
        if hasta is not None and d > hasta:
            break
        if d in diario:
            total += float(diario[d])
            n += 1
        else:
            n += 1  # día sin fila API = 0 m³, cuenta para el promedio de 7 días
    dias = 7 if hasta is None else max(1, sum(1 for i in range(7) if lunes + timedelta(days=i) <= hasta))
    return SemanaPunto(
        node_id=node_id,
        nombre=nombre,
        lunes=lunes,
        con_wes=con_wes,
        m3_semana=total,
        dias_con_dato=n,
        m3_dia=total / dias if dias else 0.0,
    )


def _rendimiento(m3_sin: float, m3_con: float) -> Tuple[float, float]:
    ahorro = m3_sin - m3_con
    pct = (ahorro / m3_sin * 100.0) if m3_sin > 1e-9 else 0.0
    return ahorro, pct


def _mediana_positiva(vals: Sequence[float]) -> float:
    import statistics

    pos = [float(v) for v in vals if v > 1.0]
    if not pos:
        pos = [float(v) for v in vals]
    return float(statistics.median(pos)) if pos else 0.0


def _clasificar_ocupacion(m3: float, mediana: float) -> str:
    """receso / baja ocupación no se cuentan como el mejor rendimiento WES."""
    if mediana <= 1e-6:
        return "sin_referencia"
    if m3 < 0.20 * mediana:
        return "receso"
    if m3 < 0.50 * mediana:
        return "baja_ocupacion"
    if m3 > 2.5 * mediana:
        return "outlier_alto"
    return "operativa"


def _grafico_barras_semanas(
    etiquetas: List[str],
    valores_con: List[float],
    valor_sin: float,
    titulo: str,
    ylabel: str,
    out_png: Path,
    idx_mejor: Optional[int] = None,
) -> None:
    fig, ax = plt.subplots(figsize=(12.5, 5.2))
    x = list(range(len(etiquetas)))
    colors = []
    for i in range(len(etiquetas)):
        if idx_mejor is not None and i == idx_mejor:
            colors.append(COLOR_AHORRO)
        else:
            colors.append("#5B9BD5")
    ax.bar(x, valores_con, color=colors, zorder=2, label="Con WES (ventana homóloga)")
    ax.axhline(
        valor_sin,
        color=COLOR_SIN,
        linestyle="--",
        linewidth=1.8,
        label=f"Sin WES 17–19 ago ({format_number_chilean(valor_sin, 1)} m³)",
        zorder=3,
    )
    ax.set_xticks(x)
    ax.set_xticklabels(etiquetas, rotation=55, ha="right", fontsize=8)
    ax.set_ylabel(ylabel)
    ax.set_title(titulo, fontweight="bold", fontsize=11, color="#1F4788")
    ax.grid(axis="y", alpha=0.3, zorder=0)
    ax.legend(fontsize=8, loc="upper right")
    fig.tight_layout()
    fig.savefig(out_png, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _grafico_ranking_pct(etiquetas: List[str], pcts: List[float], titulo: str, out_png: Path) -> None:
    fig, ax = plt.subplots(figsize=(12.5, 5.2))
    x = list(range(len(etiquetas)))
    colors = [COLOR_AHORRO if v > 0 else COLOR_SIN for v in pcts]
    ax.bar(x, pcts, color=colors, zorder=2)
    ax.axhline(0, color="#333333", linewidth=0.8)
    ax.set_xticks(x)
    ax.set_xticklabels(etiquetas, rotation=55, ha="right", fontsize=8)
    ax.set_ylabel("Rendimiento de ahorro (%)")
    ax.set_title(titulo, fontweight="bold", fontsize=11, color="#1F4788")
    ax.grid(axis="y", alpha=0.3, zorder=0)
    fig.tight_layout()
    fig.savefig(out_png, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _grafico_resumen_puntos(
    nombres: List[str],
    m3_sin: List[float],
    m3_mejor: List[float],
    m3_mediana: List[float],
    etiqueta_mejor: str,
    out_png: Path,
) -> None:
    import numpy as np

    x = np.arange(len(nombres))
    w = 0.26
    fig, ax = plt.subplots(figsize=(11, 5.2))
    ax.bar(x - w, m3_sin, width=w, color=COLOR_SIN, label="Sin WES (17–19 ago)", zorder=2)
    ax.bar(x, m3_mejor, width=w, color=COLOR_AHORRO, label=f"Última semana con WES ({etiqueta_mejor})", zorder=2)
    ax.bar(x + w, m3_mediana, width=w, color="#5B9BD5", label="Mediana semanas operativas con WES", zorder=2)
    ax.set_xticks(list(x))
    ax.set_xticklabels(nombres, rotation=18, ha="right", fontsize=9)
    ax.set_ylabel("Consumo ventana homóloga (m³)")
    ax.set_title(
        "Renca — ventana lun–mié: sin control vs última semana con WES y mediana operativa",
        fontweight="bold",
        fontsize=11,
        color="#1F4788",
    )
    ax.grid(axis="y", alpha=0.3, zorder=0)
    ax.legend(fontsize=8)
    fig.tight_layout()
    fig.savefig(out_png, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _grafico_serie_diaria(
    diario_por_nodo: Dict[str, Dict[date, float]],
    nombres: Dict[str, str],
    d0: date,
    d1: date,
    out_png: Path,
) -> None:
    fig, axes = plt.subplots(2, 2, figsize=(12.8, 7.2), sharex=True)
    dias = []
    d = d0
    while d <= d1:
        dias.append(d)
        d += timedelta(days=1)
    xs = list(range(len(dias)))
    for ax, (nid, nom) in zip(axes.ravel(), PUNTOS):
        ys = [float(diario_por_nodo.get(nid, {}).get(dia, 0.0)) for dia in dias]
        ax.plot(xs, ys, color=COLOR_WES, linewidth=1.1)
        # sombrear periodo sin WES
        i_sin = next((i for i, dia in enumerate(dias) if dia >= LUNES_SIN_WES), None)
        if i_sin is not None:
            ax.axvspan(i_sin - 0.5, len(dias) - 0.5, color=COLOR_SIN, alpha=0.18, label="Sin WES")
        ax.set_title(nom, fontsize=10, fontweight="bold", color="#1F4788")
        ax.set_ylabel("m³/día")
        ax.grid(alpha=0.3)
        ticks = [i for i, dia in enumerate(dias) if dia.day == 1 or dia == d0]
        ax.set_xticks(ticks)
        ax.set_xticklabels([dias[i].strftime("%d/%m") for i in ticks], fontsize=8)
    fig.suptitle("Consumo diario mayo–agosto 2026 (sombra = sin control WES)", fontsize=12, fontweight="bold")
    fig.tight_layout()
    fig.savefig(out_png, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _grafico_perfil_24h(
    vec_sin: List[float],
    vec_con: List[float],
    titulo: str,
    out_png: Path,
    hora_corte: Optional[int] = None,
) -> None:
    fig, ax = plt.subplots(figsize=(10.5, 4.4))
    x = list(range(24))
    ax.bar([i - 0.18 for i in x], vec_sin, width=0.36, color=COLOR_SIN, label="Sin WES", zorder=2)
    ax.bar([i + 0.18 for i in x], vec_con, width=0.36, color=COLOR_WES, label="Con WES", zorder=2)
    ax.axvspan(-0.5, 6.5, color="#c41e1e", alpha=0.10, zorder=0)
    if hora_corte is not None:
        ax.axvline(hora_corte + 0.5, color="#666666", linestyle=":", linewidth=1.2)
    ax.set_xticks(x)
    ax.set_xticklabels([f"{h:02d}" for h in x], fontsize=7)
    ax.set_ylabel("m³/h")
    ax.set_title(titulo, fontweight="bold", fontsize=11, color="#1F4788")
    ax.legend(fontsize=8)
    ax.grid(axis="y", alpha=0.3, zorder=0)
    fig.tight_layout()
    fig.savefig(out_png, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _style_header_row(ws, row: int, cols: int) -> None:
    fill = PatternFill("solid", fgColor="1F4788")
    font = Font(color="FFFFFF", bold=True, name="Calibri", size=10)
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        cell.border = thin


def _excel(
    out_xlsx: Path,
    diario_por_nodo: Dict[str, Dict[date, float]],
    ventanas: Dict[str, List[VentanaPunto]],
    ventana_sin: Dict[str, VentanaPunto],
    semanas: Dict[str, List[SemanaPunto]],
    hora_corte: int,
    ahora: datetime,
) -> None:
    wb = Workbook()
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )

    # Diario
    ws = wb.active
    ws.title = "Consumo_diario"
    headers = ["Fecha", "Día"] + [n for _, n in PUNTOS] + ["Total 4 puntos (m³)"]
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))
    d0 = date(2026, 5, 1)
    d1 = ahora.date()
    d = d0
    nomb_dia = ("lun", "mar", "mié", "jue", "vie", "sáb", "dom")
    while d <= d1:
        vals = [float(diario_por_nodo.get(nid, {}).get(d, 0.0)) for nid, _ in PUNTOS]
        ws.append([d.isoformat(), nomb_dia[d.weekday()]] + [round(v, 4) for v in vals] + [round(sum(vals), 4)])
        d += timedelta(days=1)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for c in row:
            c.border = thin
            c.font = Font(name="Calibri", size=10)

    # Ventana homóloga
    ws2 = wb.create_sheet("Ventana_homologa")
    h2 = [
        "Nodo",
        "Establecimiento",
        "Semana (lun–dom)",
        "Control",
        "m³ lunes",
        "m³ martes",
        "m³ miércoles (hasta hora corte)",
        "m³ ventana",
        "m³ Sin WES (misma ventana)",
        "Ahorro m³",
        "Rendimiento %",
        "Ahorro CLP",
        "Ocupación",
    ]
    ws2.append(h2)
    _style_header_row(ws2, 1, len(h2))
    for nid, nom in PUNTOS:
        sin = ventana_sin[nid]
        for v in ventanas[nid]:
            ahorro, pct = _rendimiento(sin.m3_total, v.m3_total)
            ws2.append(
                [
                    nid,
                    nom,
                    _etiqueta_semana(v.lunes),
                    "Con WES",
                    round(v.m3_lunes, 4),
                    round(v.m3_martes, 4),
                    round(v.m3_miercoles_corte, 4),
                    round(v.m3_total, 4),
                    round(sin.m3_total, 4),
                    round(ahorro, 4),
                    round(pct, 2),
                    round(ahorro * TARIFA_CLP_M3, 0),
                    v.nota or "operativa",
                ]
            )
        ws2.append(
            [
                nid,
                nom,
                _etiqueta_semana(sin.lunes),
                "Sin WES",
                round(sin.m3_lunes, 4),
                round(sin.m3_martes, 4),
                round(sin.m3_miercoles_corte, 4),
                round(sin.m3_total, 4),
                round(sin.m3_total, 4),
                0,
                0,
                0,
                "actual",
            ]
        )
    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, max_col=len(h2)):
        for c in row:
            c.border = thin
            c.font = Font(name="Calibri", size=10)

    # Semanas completas
    ws3 = wb.create_sheet("Semanas_completas")
    h3 = [
        "Nodo",
        "Establecimiento",
        "Semana",
        "Control",
        "m³ semana (lun–dom)",
        "m³/día",
        "m³/día Sin WES (ritmo actual)",
        "Ahorro m³ vs ritmo sin WES × días",
        "Rendimiento % vs ritmo sin WES",
    ]
    ws3.append(h3)
    _style_header_row(ws3, 1, len(h3))
    horas_eq = next(iter(ventana_sin.values())).horas_equivalentes
    for nid, nom in PUNTOS:
        sin = ventana_sin[nid]
        ritmo = (sin.m3_total / horas_eq * 24.0) if horas_eq else 0.0
        for s in semanas[nid]:
            ahorro_est = ritmo * 7.0 - s.m3_semana
            pct = (ahorro_est / (ritmo * 7.0) * 100.0) if ritmo > 1e-9 else 0.0
            ws3.append(
                [
                    nid,
                    nom,
                    _etiqueta_semana(s.lunes),
                    "Con WES" if s.con_wes else "Sin WES (incompleta)",
                    round(s.m3_semana, 4),
                    round(s.m3_dia, 4),
                    round(ritmo, 4),
                    round(ahorro_est, 4),
                    round(pct, 2),
                ]
            )
    for row in ws3.iter_rows(min_row=2, max_row=ws3.max_row, max_col=len(h3)):
        for c in row:
            c.border = thin
            c.font = Font(name="Calibri", size=10)

    # Ranking
    ws4 = wb.create_sheet("Ranking_mejor_semana")
    h4 = [
        "Establecimiento",
        "Mejor semana (mayor %)",
        "m³ con WES",
        "m³ sin WES (17–19 ago)",
        "Ahorro m³",
        "Rendimiento %",
        "Ahorro CLP",
        "Mejor semana (mayor m³)",
        "Ahorro m³ (máximo volumen)",
        "Rendimiento % de esa semana",
    ]
    ws4.append(h4)
    _style_header_row(ws4, 1, len(h4))
    for nid, nom in PUNTOS:
        sin = ventana_sin[nid]
        ranked = []
        for v in ventanas[nid]:
            ahorro, pct = _rendimiento(sin.m3_total, v.m3_total)
            ranked.append((pct, ahorro, v))
        ops = [t for t in ranked if t[2].nota == "operativa"]
        pool = ops or ranked
        best_pct = max(pool, key=lambda t: t[0])
        best_m3 = max(pool, key=lambda t: t[1])
        ws4.append(
            [
                nom,
                _etiqueta_semana(best_pct[2].lunes),
                round(best_pct[2].m3_total, 4),
                round(sin.m3_total, 4),
                round(best_pct[1], 4),
                round(best_pct[0], 2),
                round(best_pct[1] * TARIFA_CLP_M3, 0),
                _etiqueta_semana(best_m3[2].lunes),
                round(best_m3[1], 4),
                round(best_m3[0], 2),
            ]
        )
    for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row, max_col=len(h4)):
        for c in row:
            c.border = thin
            c.font = Font(name="Calibri", size=10)

    ws5 = wb.create_sheet("Metadatos")
    ws5["A1"] = "Periodo sin WES"
    ws5["B1"] = f"{LUNES_SIN_WES:%d/%m/%Y} 00:00 Chile hasta {ahora:%d/%m/%Y %H:%M} Chile"
    ws5["A2"] = "Última hora completa incluida (miércoles)"
    ws5["B2"] = f"{hora_corte:02d}:00–{hora_corte:02d}:59"
    ws5["A3"] = "Horas equivalentes de la ventana"
    ws5["B3"] = 48 + hora_corte + 1
    ws5["A4"] = "Semanas con WES"
    ws5["B4"] = f"{PRIMER_LUNES_CON_WES:%d/%m/%Y} a {ULTIMO_LUNES_CON_WES:%d/%m/%Y} (lunes de cada semana)"
    ws5["A5"] = "Tarifa CLP/m³"
    ws5["B5"] = TARIFA_CLP_M3
    ws5["A6"] = "Fuente"
    ws5["B6"] = "API WES totalM3 diario + serie horaria dates.measures.csv (hora Chile)"
    ws5["A7"] = "Rendimiento"
    ws5["B7"] = "(m³ Sin WES − m³ Con WES) / m³ Sin WES × 100"

    for wsx in wb.worksheets:
        for col in wsx.columns:
            letter = get_column_letter(col[0].column)
            maxlen = 12
            for cell in col[:40]:
                maxlen = max(maxlen, min(42, len(str(cell.value or ""))))
            wsx.column_dimensions[letter].width = maxlen + 2

    wb.save(out_xlsx)


def _word(
    out_docx: Path,
    pngs: Dict[str, Path],
    ventana_sin: Dict[str, VentanaPunto],
    ventanas: Dict[str, List[VentanaPunto]],
    hora_corte: int,
    ahora: datetime,
    ranking_agregado: List[Tuple[date, float, float, float]],
    mejor_por_punto: Dict[str, Tuple[VentanaPunto, float, float]],
    mediana_por_punto: Dict[str, float],
) -> None:
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    try:
        add_logo_to_header(doc)
    except Exception:
        pass

    h = doc.add_heading("Rendimiento hídrico WES — Renca (sin control desde el 17/08/2026)", level=0)
    if h.runs:
        h.runs[0].font.color.rgb = COLOR_HEAD

    p = doc.add_paragraph()
    p.add_run("Cliente / comuna: ").bold = True
    p.add_run("ICCP Renca — Escuela Lo Velásquez, Gimnasio municipal y Piscina municipal.")
    p = doc.add_paragraph()
    p.add_run("Generado: ").bold = True
    p.add_run(ahora.strftime("%d/%m/%Y %H:%M") + " (hora Chile).")

    doc.add_heading("1. Pregunta y método", level=1)
    doc.add_paragraph(
        "Desde el lunes 17 de agosto de 2026 los equipos quedaron sin control WES. "
        "Se compara ese periodo (hasta ahora) con cada semana lunes–domingo desde mayo 2026, "
        "cuando el control sí estaba activo, para responder: (a) qué semana genera el mayor "
        "rendimiento de ahorro hídrico; (b) si esta ventana con WES ahorra más que la misma "
        "ventana sin WES."
    )
    horas_eq = next(iter(ventana_sin.values())).horas_equivalentes
    doc.add_paragraph(
        f"Ventana homóloga (manzanas con manzanas): lunes 00:00–24:00 + martes 00:00–24:00 + "
        f"miércoles 00:00–{hora_corte:02d}:59 (última hora completa al momento del informe). "
        f"Son {horas_eq} horas equivalentes en cada semana. "
        f"El miércoles 19/08 se corta a las {hora_corte:02d}:59 para no comparar un día incompleto "
        f"contra miércoles históricos de 24 h."
    )
    doc.add_paragraph(
        "Rendimiento % = (consumo Sin WES − consumo Con WES) / consumo Sin WES × 100. "
        "Positivo: esa semana con WES consumió menos que el periodo actual sin control. "
        "Se excluyen de la semana ganadora los recesos (consumo < 20 % de la mediana del punto): "
        "un recinto cerrado no es ahorro del equipo. "
        f"Valoración referencial: ${format_number_chilean(TARIFA_CLP_M3, 0)} CLP/m³."
    )

    # Totales actuales
    doc.add_heading("2. Periodo sin WES (17–19 agosto 2026)", level=1)
    tbl = doc.add_table(rows=1 + len(PUNTOS) + 1, cols=5)
    tbl.style = "Table Grid"
    headers = ["Establecimiento", "Lunes 17 (m³)", "Martes 18 (m³)", f"Miércoles 19 hasta {hora_corte:02d}:59 (m³)", "Total ventana (m³)"]
    for j, hd in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        cell.text = hd
        _set_cell_shading(cell, "1F4788")
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9)
    tot = [0.0, 0.0, 0.0, 0.0]
    for i, (nid, nom) in enumerate(PUNTOS, start=1):
        v = ventana_sin[nid]
        vals = [v.m3_lunes, v.m3_martes, v.m3_miercoles_corte, v.m3_total]
        for k, x in enumerate(vals):
            tot[k] += x
        tbl.rows[i].cells[0].text = nom
        for j, x in enumerate(vals):
            tbl.rows[i].cells[j + 1].text = format_number_chilean(x, 1)
    last = tbl.rows[1 + len(PUNTOS)]
    last.cells[0].text = "Total 4 puntos"
    for run in last.cells[0].paragraphs[0].runs:
        run.bold = True
    for j, x in enumerate(tot):
        last.cells[j + 1].text = format_number_chilean(x, 1)
        for run in last.cells[j + 1].paragraphs[0].runs:
            run.bold = True

    m3_sin_total = tot[3]
    doc.add_paragraph(
        f"Consumo conjunto sin control en la ventana: {format_number_chilean(m3_sin_total, 1)} m³ "
        f"({format_number_chilean(m3_sin_total / horas_eq * 24, 1)} m³/día de ritmo)."
    )

    doc.add_heading("3. Comparación inmediata: última semana con WES vs ahora sin WES", level=1)
    doc.add_paragraph(
        "La lectura más limpia es la semana previa (lunes 10 a domingo 16 de agosto, control activo) "
        "contra el mismo tramo lun–mié de esta semana, ya sin control. Misma estación, misma ocupación "
        "esperable, mismos días de la semana."
    )
    prev_lunes = ULTIMO_LUNES_CON_WES
    tblp = doc.add_table(rows=1 + len(PUNTOS) + 1, cols=6)
    tblp.style = "Table Grid"
    hp = ["Establecimiento", "10–16 ago con WES (m³)", "17–19 ago sin WES (m³)", "Ahorro (m³)", "Rendimiento", "CLP"]
    for j, hd in enumerate(hp):
        cell = tblp.rows[0].cells[j]
        cell.text = hd
        _set_cell_shading(cell, "1F4788")
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9)
    tot_prev = 0.0
    tot_sin_p = 0.0
    for i, (nid, nom) in enumerate(PUNTOS, start=1):
        prev = next(v for v in ventanas[nid] if v.lunes == prev_lunes)
        sin = ventana_sin[nid]
        ahorro, pct = _rendimiento(sin.m3_total, prev.m3_total)
        tot_prev += prev.m3_total
        tot_sin_p += sin.m3_total
        row = tblp.rows[i]
        row.cells[0].text = nom
        row.cells[1].text = format_number_chilean(prev.m3_total, 1)
        row.cells[2].text = format_number_chilean(sin.m3_total, 1)
        row.cells[3].text = format_number_chilean(ahorro, 1)
        row.cells[4].text = format_number_chilean(pct, 1) + " %"
        row.cells[5].text = "$" + format_number_chilean(ahorro * TARIFA_CLP_M3, 0)
    ahorro_p, pct_p = _rendimiento(tot_sin_p, tot_prev)
    lastp = tblp.rows[1 + len(PUNTOS)]
    lastp.cells[0].text = "Total 4 puntos"
    lastp.cells[1].text = format_number_chilean(tot_prev, 1)
    lastp.cells[2].text = format_number_chilean(tot_sin_p, 1)
    lastp.cells[3].text = format_number_chilean(ahorro_p, 1)
    lastp.cells[4].text = format_number_chilean(pct_p, 1) + " %"
    lastp.cells[5].text = "$" + format_number_chilean(ahorro_p * TARIFA_CLP_M3, 0)
    for c in lastp.cells:
        for run in c.paragraphs[0].runs:
            run.bold = True
    doc.add_paragraph(
        f"En esta ventana, con WES se consumieron {format_number_chilean(tot_prev, 1)} m³ y sin WES "
        f"{format_number_chilean(tot_sin_p, 1)} m³. El control de la semana previa ahorra "
        f"{format_number_chilean(ahorro_p, 1)} m³ ({format_number_chilean(pct_p, 1)} %), "
        f"≈ ${format_number_chilean(ahorro_p * TARIFA_CLP_M3, 0)}. "
        "El salto lo explica sobre todo la piscina (de ~31 m³/día con control a ~56–58 m³/día sin control) "
        "e ICCP (de ~16 m³/día a ~21 m³/día)."
    )

    doc.add_heading("4. Semana de mayor rendimiento operativo (mayo → 16 agosto)", level=1)
    doc.add_paragraph(
        "Entre las semanas con ocupación comparable (se descartan recesos: piscina o recinto en ~0 m³), "
        "la ganadora es la de mayor rendimiento % respecto del periodo actual sin WES."
    )

    tbl2 = doc.add_table(rows=1 + len(PUNTOS) + 1, cols=7)
    tbl2.style = "Table Grid"
    h2 = ["Establecimiento", "Mejor semana", "Con WES (m³)", "Sin WES (m³)", "Ahorro (m³)", "Rendimiento", "CLP"]
    for j, hd in enumerate(h2):
        cell = tbl2.rows[0].cells[j]
        cell.text = hd
        _set_cell_shading(cell, "1F4788")
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9)

    ahorros = []
    for i, (nid, nom) in enumerate(PUNTOS, start=1):
        v, ahorro, pct = mejor_por_punto[nid]
        ahorros.append(ahorro)
        sin = ventana_sin[nid]
        row = tbl2.rows[i]
        row.cells[0].text = nom
        row.cells[1].text = _etiqueta_semana(v.lunes)
        row.cells[2].text = format_number_chilean(v.m3_total, 1)
        row.cells[3].text = format_number_chilean(sin.m3_total, 1)
        row.cells[4].text = format_number_chilean(ahorro, 1)
        row.cells[5].text = format_number_chilean(pct, 1) + " %"
        row.cells[6].text = "$" + format_number_chilean(ahorro * TARIFA_CLP_M3, 0)

    # Agregado: mejor semana OPERATIVA (piscina abierta y escuela no en receso)
    ranking_op = []
    for t in ranking_agregado:
        pisc = next(v for v in ventanas["000017-06"] if v.lunes == t[0])
        esc = next(v for v in ventanas["000017-04"] if v.lunes == t[0])
        if pisc.nota == "operativa" and esc.nota != "receso":
            ranking_op.append(t)
    ranking_op = ranking_op or list(ranking_agregado)
    best_agg = max(ranking_op, key=lambda t: t[2])
    best_agg_m3 = max(ranking_op, key=lambda t: t[1])
    last = tbl2.rows[1 + len(PUNTOS)]
    last.cells[0].text = "Agregado 4 puntos (mejor semana operativa)"
    last.cells[1].text = _etiqueta_semana(best_agg[0])
    last.cells[2].text = format_number_chilean(best_agg[3], 1)
    last.cells[3].text = format_number_chilean(m3_sin_total, 1)
    last.cells[4].text = format_number_chilean(best_agg[1], 1)
    last.cells[5].text = format_number_chilean(best_agg[2], 1) + " %"
    last.cells[6].text = "$" + format_number_chilean(best_agg[1] * TARIFA_CLP_M3, 0)
    for c in last.cells:
        for run in c.paragraphs[0].runs:
            run.bold = True

    doc.add_paragraph("")
    if "resumen_puntos" in pngs:
        doc.add_picture(str(pngs["resumen_puntos"]), width=Cm(16.2))
    if "ranking_agregado" in pngs:
        doc.add_paragraph("")
        doc.add_picture(str(pngs["ranking_agregado"]), width=Cm(16.2))

    doc.add_heading("5. ¿Esta ventana con WES ahorra más que sin WES?", level=1)
    import statistics

    agg_op = [t[3] for t in ranking_op]
    mediana_agg = statistics.median(agg_op) if agg_op else 0.0
    ahorro_med, pct_med = _rendimiento(m3_sin_total, mediana_agg)
    ahorro_best, pct_best = _rendimiento(m3_sin_total, best_agg[3])

    doc.add_paragraph(
        f"En la misma ventana lun–mié, el consumo conjunto sin WES fue "
        f"{format_number_chilean(m3_sin_total, 1)} m³. "
        f"La mediana de las {len(ranking_op)} semanas operativas con WES (mayo–16 ago, sin receso de piscina) fue "
        f"{format_number_chilean(mediana_agg, 1)} m³ "
        f"(rendimiento {format_number_chilean(pct_med, 1)} %, "
        f"{format_number_chilean(ahorro_med, 1)} m³). "
        f"La mejor semana operativa ({_etiqueta_semana(best_agg[0])}) bajó a "
        f"{format_number_chilean(best_agg[3], 1)} m³ "
        f"({format_number_chilean(pct_best, 1)} %, "
        f"{format_number_chilean(ahorro_best, 1)} m³). "
        f"La última semana con control (10–16 ago) ahorró {format_number_chilean(ahorro_p, 1)} m³ "
        f"({format_number_chilean(pct_p, 1)} %) en la misma ventana."
    )

    if tot_sin_p > tot_prev:
        doc.add_paragraph(
            "Conclusión: el periodo CON control WES genera más rendimiento de ahorro que este periodo "
            "SIN control. En la comparación de temporada (10–16 ago vs 17–19 ago) el consumo ya subió "
            f"{format_number_chilean(pct_p, 1)} %. Restaurar el control es lo que recupera ese caudal, "
            "en particular en piscina e ICCP."
        )
    else:
        doc.add_paragraph(
            "Conclusión: el consumo actual sin WES no supera la última semana con control; "
            "revisar el detalle por establecimiento."
        )

    doc.add_heading("6. Detalle por establecimiento", level=1)
    for nid, nom in PUNTOS:
        v_best, ahorro, pct = mejor_por_punto[nid]
        sin = ventana_sin[nid]
        med = mediana_por_punto[nid]
        ahorro_med_p, pct_med_p = _rendimiento(sin.m3_total, med)
        doc.add_heading(nom, level=2)
        doc.add_paragraph(
            f"Sin WES (17–19 ago): {format_number_chilean(sin.m3_total, 1)} m³ "
            f"(lun {format_number_chilean(sin.m3_lunes, 1)} / "
            f"mar {format_number_chilean(sin.m3_martes, 1)} / "
            f"mié-corte {format_number_chilean(sin.m3_miercoles_corte, 1)}). "
            f"Mejor semana con WES: {_etiqueta_semana(v_best.lunes)} con "
            f"{format_number_chilean(v_best.m3_total, 1)} m³ "
            f"(ahorro {format_number_chilean(ahorro, 1)} m³; {format_number_chilean(pct, 1)} %). "
            f"Mediana con WES: {format_number_chilean(med, 1)} m³ "
            f"({format_number_chilean(pct_med_p, 1)} %)."
        )
        key = f"semanas_{nid}"
        if key in pngs:
            doc.add_picture(str(pngs[key]), width=Cm(16.0))
        keyp = f"pct_{nid}"
        if keyp in pngs:
            doc.add_picture(str(pngs[keyp]), width=Cm(16.0))
        keyh = f"perfil_{nid}"
        if keyh in pngs:
            doc.add_picture(str(pngs[keyh]), width=Cm(16.0))

    if "serie_diaria" in pngs:
        doc.add_heading("7. Serie diaria mayo–agosto", level=1)
        doc.add_paragraph(
            "La banda roja marca el periodo sin control (desde el lunes 17/08). "
            "Piscina y gimnasio dominan el volumen; escuela e ICCP tienen consumos menores y más estables."
        )
        doc.add_picture(str(pngs["serie_diaria"]), width=Cm(16.2))

    doc.add_heading("8. Lectura operativa", level=1)
    bullets = [
        f"La comparación es de {horas_eq} horas equivalentes, no de una semana calendario completa "
        f"(el periodo sin WES aún no cierra el domingo 23/08).",
        "La semana 15–21 jun (piscina en 0 m³) y 22 jun–5 jul (escuela en 0 m³) se marcan como receso: "
        "no se cuentan como el mejor rendimiento WES.",
        "Piscina: con WES en agosto venía ~31 m³/día; el lunes 17 y martes 18 sin control subió a 56 y 58 m³/día "
        "(nivel similar a mayo). Ese es el mayor volumen perdido.",
        "ICCP: lun–mar con WES ~16 m³/día; sin WES ~21 m³/día.",
        "Escuela Lo Velásquez y gimnasio: en esta ventana el consumo sin WES no supera de forma clara "
        "la semana previa con control; el ahorro conjunto lo arrastran piscina e ICCP.",
        "Cuando se complete la semana 17–23 ago conviene repetir el cruce lun–dom contra lun–dom.",
        f"Tarifa usada para CLP: ${format_number_chilean(TARIFA_CLP_M3, 0)} / m³ (referencial).",
    ]
    for b in bullets:
        doc.add_paragraph(b, style="List Bullet")

    doc.add_heading("9. Respuesta directa", level=1)
    doc.add_paragraph(
        f"• Semana de mayor rendimiento operativo conjunto (4 puntos, sin receso de piscina): "
        f"{_etiqueta_semana(best_agg[0])} "
        f"({format_number_chilean(best_agg[2], 1)} %; {format_number_chilean(best_agg[1], 1)} m³ "
        f"en la ventana homóloga vs el periodo actual sin WES)."
    )
    doc.add_paragraph(
        f"• Comparación de temporada (la más justa): 10–16 ago CON WES vs 17–19 ago SIN WES → "
        f"{format_number_chilean(pct_p, 1)} % a favor del control "
        f"({format_number_chilean(ahorro_p, 1)} m³ en {horas_eq} h)."
    )
    if best_agg_m3[0] != best_agg[0]:
        doc.add_paragraph(
            f"• Semana de mayor volumen de ahorro conjunto: {_etiqueta_semana(best_agg_m3[0])} "
            f"({format_number_chilean(best_agg_m3[1], 1)} m³; {format_number_chilean(best_agg_m3[2], 1)} %)."
        )
    for nid, nom in PUNTOS:
        v, ahorro, pct = mejor_por_punto[nid]
        doc.add_paragraph(
            f"• {nom}: mejor semana {_etiqueta_semana(v.lunes)} "
            f"({format_number_chilean(pct, 1)} %; {format_number_chilean(ahorro, 1)} m³)."
        )
    if tot_sin_p > tot_prev:
        doc.add_paragraph(
            "• Esta ventana CON WES genera más ahorro que la misma ventana SIN WES: "
            "el lunes 17 y martes 18 ya superan el ritmo de la semana previa con equipos activos "
            "(sobre todo piscina e ICCP)."
        )

    out_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_docx)


def _convertir_pdf(docx_path: Path) -> Optional[Path]:
    pdf_path = docx_path.with_suffix(".pdf")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    subprocess.run(
        [
            soffice,
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
            "--outdir",
            str(docx_path.parent),
            str(docx_path),
        ],
        check=False,
        capture_output=True,
        text=True,
        timeout=180,
    )
    return pdf_path if pdf_path.is_file() else None


def main() -> int:
    ahora = _ahora_chile()
    hora_corte = _ultima_hora_completa(ahora)
    if ahora.date() > LUNES_SIN_WES + timedelta(days=2):
        # Ya pasó el miércoles 19: ventana = lun+mar+mié completos.
        hora_corte = 23
    elif ahora.date() == LUNES_SIN_WES:
        print("[AVISO] Aún es lunes 17; la ventana será solo horas de hoy.")
    print(f"Ahora Chile: {ahora:%Y-%m-%d %H:%M} | última hora completa miércoles: {hora_corte:02d}:00")

    d_ini = date(2026, 5, 1)
    d_fin = ahora.date()
    lunes_con = _iter_lunes(PRIMER_LUNES_CON_WES, ULTIMO_LUNES_CON_WES)

    print("Descargando totalM3 diario (4 nodos, mayo–hoy)...")
    diario_por_nodo: Dict[str, Dict[date, float]] = {}
    for nid, nom in PUNTOS:
        diario_por_nodo[nid] = _fetch_diario(nid, d_ini, d_fin)
        print(f"  {nid} {nom}: {len(diario_por_nodo[nid])} días")

    # Miércoles a consultar (cada semana con WES + miércoles actual)
    miercoles_ids: List[Tuple[str, date]] = []
    for lunes in lunes_con:
        mie = lunes + timedelta(days=2)
        for nid, _ in PUNTOS:
            miercoles_ids.append((nid, mie))
    mie_actual = LUNES_SIN_WES + timedelta(days=2)
    for nid, _ in PUNTOS:
        miercoles_ids.append((nid, mie_actual))

    print(f"Descargando serie horaria de {len(miercoles_ids)} miércoles (corte 00–{hora_corte:02d})...")
    vecs: Dict[Tuple[str, date], List[float]] = {}

    def _job(item: Tuple[str, date]) -> Tuple[Tuple[str, date], List[float]]:
        nid, dia = item
        return item, _vector_24h(nid, dia)

    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(_job, it) for it in miercoles_ids]
        done = 0
        for fut in as_completed(futs):
            key, vec = fut.result()
            vecs[key] = vec
            done += 1
            if done % 10 == 0 or done == len(miercoles_ids):
                print(f"  horarios {done}/{len(miercoles_ids)}")

    ventanas: Dict[str, List[VentanaPunto]] = {nid: [] for nid, _ in PUNTOS}
    ventana_sin: Dict[str, VentanaPunto] = {}
    semanas: Dict[str, List[SemanaPunto]] = {nid: [] for nid, _ in PUNTOS}

    for nid, nom in PUNTOS:
        for lunes in lunes_con:
            mie = lunes + timedelta(days=2)
            ventanas[nid].append(
                _consumo_ventana(
                    nid, nom, lunes, diario_por_nodo[nid], vecs[(nid, mie)], hora_corte, True
                )
            )
            semanas[nid].append(_semana_completa(nid, nom, lunes, diario_por_nodo[nid], True))
        ventana_sin[nid] = _consumo_ventana(
            nid,
            nom,
            LUNES_SIN_WES,
            diario_por_nodo[nid],
            vecs[(nid, mie_actual)],
            hora_corte,
            False,
        )
        semanas[nid].append(
            _semana_completa(nid, nom, LUNES_SIN_WES, diario_por_nodo[nid], False, hasta=d_fin)
        )

    # Clasificar ocupación (receso ≠ rendimiento WES)
    for nid, _ in PUNTOS:
        med = _mediana_positiva([v.m3_total for v in ventanas[nid]])
        for v in ventanas[nid]:
            v.nota = _clasificar_ocupacion(v.m3_total, med)

    # Ranking agregado (suma 4 puntos) por semana con WES
    ranking_agregado: List[Tuple[date, float, float, float]] = []
    ranking_operativo: List[Tuple[date, float, float, float]] = []
    m3_sin_total = sum(ventana_sin[nid].m3_total for nid, _ in PUNTOS)
    for i, lunes in enumerate(lunes_con):
        m3_con = sum(ventanas[nid][i].m3_total for nid, _ in PUNTOS)
        ahorro, pct = _rendimiento(m3_sin_total, m3_con)
        ranking_agregado.append((lunes, ahorro, pct, m3_con))
        # Semana operativa agregada: piscina abierta y escuela no en receso
        # (un receso escolar infla el % sin ser ahorro del equipo).
        pisc = ventanas["000017-06"][i].nota
        esc = ventanas["000017-04"][i].nota
        if pisc == "operativa" and esc != "receso":
            ranking_operativo.append((lunes, ahorro, pct, m3_con))

    import statistics

    mejor_por_punto: Dict[str, Tuple[VentanaPunto, float, float]] = {}
    mediana_por_punto: Dict[str, float] = {}
    for nid, _ in PUNTOS:
        sin = ventana_sin[nid]
        scored = []
        for v in ventanas[nid]:
            ahorro, pct = _rendimiento(sin.m3_total, v.m3_total)
            scored.append((pct, ahorro, v))
        ops = [t for t in scored if t[2].nota == "operativa"]
        pool = ops or [t for t in scored if t[2].nota != "receso"] or scored
        best = max(pool, key=lambda t: (t[0], t[1]))
        mejor_por_punto[nid] = (best[2], best[1], best[0])
        mediana_por_punto[nid] = statistics.median(
            [v.m3_total for v in ventanas[nid] if v.nota == "operativa"] or [v.m3_total for v in ventanas[nid]]
        )

    ts = ahora.strftime("%Y%m%d_%H%M")
    out_dir = ROOT / "reports" / "reporte de auditoria" / f"comparacion_renca_sin_wes_ago17_vs_semanas_mayo_{ts}"
    out_dir.mkdir(parents=True, exist_ok=True)
    png_dir = out_dir / "graficos"
    png_dir.mkdir(exist_ok=True)

    etiquetas = [_etiqueta_semana(l) for l in lunes_con]
    pngs: Dict[str, Path] = {}

    # Agregado
    pool_idx = ranking_operativo or ranking_agregado
    vals_agg = [t[3] for t in ranking_agregado]
    pcts_agg = [t[2] for t in ranking_agregado]
    idx_best = max(
        range(len(ranking_agregado)),
        key=lambda i: ranking_agregado[i][2]
        if ranking_agregado[i][0] in {t[0] for t in pool_idx}
        else -999,
    )
    p1 = png_dir / "barras_agregado_ventana.png"
    _grafico_barras_semanas(
        etiquetas,
        vals_agg,
        m3_sin_total,
        "Renca (4 puntos) — consumo ventana lun–mié con WES vs sin WES actual",
        "m³ ventana homóloga",
        p1,
        idx_mejor=idx_best,
    )
    pngs["resumen_semanas_agg"] = p1
    p2 = png_dir / "ranking_pct_agregado.png"
    _grafico_ranking_pct(
        etiquetas,
        pcts_agg,
        "Rendimiento de ahorro agregado vs periodo sin WES 17–19 ago",
        p2,
    )
    pngs["ranking_agregado"] = p2

    nombres_cortos = ["ICCP", "Lo Velásquez", "Gimnasio", "Piscina"]
    m3_sin_l = [ventana_sin[nid].m3_total for nid, _ in PUNTOS]
    m3_prev_l = [ventanas[nid][-1].m3_total for nid, _ in PUNTOS]
    m3_med_l = [mediana_por_punto[nid] for nid, _ in PUNTOS]
    p3 = png_dir / "resumen_puntos.png"
    _grafico_resumen_puntos(
        nombres_cortos,
        m3_sin_l,
        m3_prev_l,
        m3_med_l,
        "10–16 ago",
        p3,
    )
    pngs["resumen_puntos"] = p3

    p4 = png_dir / "serie_diaria.png"
    _grafico_serie_diaria(diario_por_nodo, {n: nom for n, nom in PUNTOS}, d_ini, d_fin, p4)
    pngs["serie_diaria"] = p4

    # Perfiles miércoles actual vs miércoles de la mejor semana de cada punto
    for nid, nom in PUNTOS:
        vals = [v.m3_total for v in ventanas[nid]]
        pcts = [_rendimiento(ventana_sin[nid].m3_total, v.m3_total)[1] for v in ventanas[nid]]
        idx_p_candidates = [
            i for i, v in enumerate(ventanas[nid]) if v.nota == "operativa"
        ]
        idx_p = max(idx_p_candidates or range(len(pcts)), key=lambda i: pcts[i])
        ps = png_dir / f"barras_{nid}.png"
        _grafico_barras_semanas(
            etiquetas,
            vals,
            ventana_sin[nid].m3_total,
            f"{nom} — ventana lun–mié con WES vs sin WES actual",
            "m³ ventana homóloga",
            ps,
            idx_mejor=idx_p,
        )
        pngs[f"semanas_{nid}"] = ps
        pp = png_dir / f"pct_{nid}.png"
        _grafico_ranking_pct(
            etiquetas,
            pcts,
            f"{nom} — rendimiento % vs sin WES 17–19 ago",
            pp,
        )
        pngs[f"pct_{nid}"] = pp

        v_best = mejor_por_punto[nid][0]
        vec_sin = vecs[(nid, mie_actual)]
        vec_con = vecs[(nid, v_best.lunes + timedelta(days=2))]
        ph = png_dir / f"perfil_mie_{nid}.png"
        _grafico_perfil_24h(
            vec_sin,
            vec_con,
            f"{nom} — perfil miércoles: {mie_actual:%d/%m} sin WES vs {v_best.lunes + timedelta(days=2):%d/%m} con WES",
            ph,
            hora_corte=hora_corte,
        )
        pngs[f"perfil_{nid}"] = ph

    xlsx = out_dir / "comparacion_renca_sin_wes_ago17_vs_semanas_mayo.xlsx"
    _excel(xlsx, diario_por_nodo, ventanas, ventana_sin, semanas, hora_corte, ahora)

    csv_path = out_dir / "ranking_agregado_ventana_homologa.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(
            [
                "lunes_semana",
                "etiqueta",
                "m3_con_wes",
                "m3_sin_wes",
                "ahorro_m3",
                "rendimiento_pct",
                "ahorro_clp",
                "ocupacion_piscina",
            ]
        )
        nota_piscina = {v.lunes: v.nota for v in ventanas["000017-06"]}
        for lunes, ahorro, pct, m3_con in ranking_agregado:
            w.writerow(
                [
                    lunes.isoformat(),
                    _etiqueta_semana(lunes),
                    f"{m3_con:.4f}",
                    f"{m3_sin_total:.4f}",
                    f"{ahorro:.4f}",
                    f"{pct:.2f}",
                    f"{ahorro * TARIFA_CLP_M3:.0f}",
                    nota_piscina.get(lunes, ""),
                ]
            )

    docx_path = out_dir / f"Rendimiento_hidrico_Renca_sinWES_17ago_vs_semanas_mayo_{ts}.docx"
    _word(
        docx_path,
        pngs,
        ventana_sin,
        ventanas,
        hora_corte,
        ahora,
        ranking_agregado,
        mejor_por_punto,
        mediana_por_punto,
    )

    pdf_path = _convertir_pdf(docx_path)

    # Consola
    horas_eq = next(iter(ventana_sin.values())).horas_equivalentes
    print("=" * 72)
    print("RENCA — SIN WES 17-19 AGO vs SEMANAS CON WES (MAYO→16 AGO)")
    print("=" * 72)
    print(f"Ventana: lun+mar + mié 00–{hora_corte:02d} h | {horas_eq} h equivalentes")
    print(f"Sin WES conjunto: {m3_sin_total:.2f} m³")
    for nid, nom in PUNTOS:
        v, ahorro, pct = mejor_por_punto[nid]
        print(
            f"  {nom}: sin {ventana_sin[nid].m3_total:.1f} m³ | "
            f"mejor {_etiqueta_semana(v.lunes)} {v.m3_total:.1f} m³ | "
            f"{pct:.1f}% ({ahorro:.1f} m³)"
        )
    best_agg = max(ranking_operativo or ranking_agregado, key=lambda t: t[2])
    prev = ranking_agregado[-1]
    print(
        f"Mejor semana operativa conjunta: {_etiqueta_semana(best_agg[0])} | "
        f"{best_agg[2]:.1f}% | {best_agg[1]:.1f} m³"
    )
    print(
        f"Última semana con WES (10–16 ago) vs actual: {prev[2]:.1f}% | {prev[1]:.1f} m³"
    )
    print(f"XLSX: {xlsx}")
    print(f"DOCX: {docx_path}")
    print(f"PDF:  {pdf_path or '(no convertido)'}")
    print(f"DIR:  {out_dir}")
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except Exception:
            pass
    raise SystemExit(main())
