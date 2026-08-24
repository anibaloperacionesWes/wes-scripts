"""
Informe de auditoría Renca — esta semana con WES (desde lun 24/08) vs la semana sin WES (lun 17/08).

Desde el lunes 24/08/2026 los 5 puntos vuelven con control.
Se compara día a día contra el mismo día de la semana del 17 (sin WES).
Hoy 24/08, 11:00–13:30: regulación de Lo Velásquez y gimnasio (mejoras de horario).

Uso:
  python generar_auditoria_renca_semana_pasada_vs_esta.py
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
from zoneinfo import ZoneInfo

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font as XlFont, PatternFill, Side
from openpyxl.utils import get_column_letter

import auditoria_cpa_icco_renca_grafico as graf
import generar_graficos_comparativos_desde_excel_consolidado as gxlsx
import generar_informe_auditoria_icco_renca_word as icco
from auditoria_cpa_icco_renca_grafico import Periodo, _vector_m3h_24_desde_api
from generar_excel_auditoria_consolidado_dos_periodos import generar_excel_consolidado
from generar_graficos_comparativos_desde_excel_consolidado import (
    _limpiar_pngs_carpeta_graficos,
    generar_pngs,
    leer_matriz_consolidado,
)
from generar_reporte_word import add_logo_to_header, format_number_chilean

ROOT = Path(__file__).resolve().parent
TZ_CL = ZoneInfo("America/Santiago")
TARIFA_CLP_M3 = 1300.0
COLOR_WES = "#2a6fad"
COLOR_SIN = "#C0504D"
COLOR_HEAD = RGBColor(31, 71, 136)

PUNTOS: Tuple[Tuple[str, str], ...] = (
    ("000017-07", "ICCP (Cumbre de Cóndores pte.)"),
    ("000017-04", "Esc. Lo Velásquez"),
    ("000017-05", "Gimnasio municipal"),
    ("000017-06", "Piscina municipal"),
    ("000017-08", "Colegio ICCO Renca"),
)
NODO_ICCO = "000017-08"
NODO_ESCUELA = "000017-04"
NODO_GIMNASIO = "000017-05"
LUNES_CON = date(2026, 8, 10)
LUNES_SIN = date(2026, 8, 17)
LUNES_PROX = date(2026, 8, 24)  # vuelve el control en los 5 puntos
HORAS_REGULADAS = (11, 12, 13)  # 11:00–13:59 cubre la regulación 11:00–13:30
HORAS_CONTROL_ICCO = frozenset(range(0, 6))
DOMINGO_SIN_ICCO = date(2026, 8, 16)
NOMBRE_ICCO = "Colegio ICCO Renca"
TEXTO_SALA_BOMBAS = ""  # texto viejo; el informe usa TEXTO_CONTROL
TEXTO_CONTROL = (
    "Desde el lunes 24/08/2026 los 5 puntos de Renca vuelven con control WES. "
    "Cada día de esta semana se compara con el mismo día de la semana del 17/08, "
    "que estuvo sin WES. Ahorro = (Sin WES − Con WES) / Sin WES × 100. "
    "Hoy 24/08, entre 11:00 y 13:30 Chile, se reguló Escuela Lo Velásquez y el "
    "gimnasio municipal con las mejoras de horario (el gimnasio ya no va con el "
    "tope parejo de 0,54 m³/h que se quedaba corto en un evento)."
)
WD_CORTO = ("Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom")
WD_LARGO = ("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo")


def _ventana_homologa(ahora: datetime) -> Tuple[Tuple[date, ...], Tuple[date, ...], int, bool]:
    """Días homólogos lun…hoy (máx. 7) y hora de corte del último día (23 = completo)."""
    hoy = ahora.date()
    if hoy >= LUNES_PROX:
        n = min(7, (hoy - LUNES_PROX).days + 1)
        ultimo_incompleto = hoy == (LUNES_PROX + timedelta(days=n - 1))
        hora_corte = max(0, min(23, ahora.hour - 1)) if ultimo_incompleto else 23
        if hora_corte < 0:
            hora_corte = 0
        dias_con = tuple(LUNES_PROX + timedelta(days=i) for i in range(n))
        dias_sin = tuple(LUNES_SIN + timedelta(days=i) for i in range(n))
        return dias_con, dias_sin, hora_corte, ultimo_incompleto
    if hoy < LUNES_SIN:
        n = 1
        hora_corte = 23
        ultimo_incompleto = False
    else:
        n = min(7, (hoy - LUNES_SIN).days + 1)
        ultimo_incompleto = hoy == (LUNES_SIN + timedelta(days=n - 1))
        hora_corte = max(0, min(23, ahora.hour - 1)) if ultimo_incompleto else 23
        if hora_corte < 0:
            hora_corte = 0
    dias_con = tuple(LUNES_CON + timedelta(days=i) for i in range(n))
    dias_sin = tuple(LUNES_SIN + timedelta(days=i) for i in range(n))
    return dias_con, dias_sin, hora_corte, ultimo_incompleto


def _fmt_rango_dias(dias: Sequence[date]) -> str:
    a, b = dias[0], dias[-1]
    return f"{a:%d/%m} al {b:%d/%m/%Y}"


def _etiqueta_horas(hora_corte: int) -> str:
    return "00:00–23:59" if hora_corte >= 23 else f"00:00–{hora_corte:02d}:59"


def _aplicar_corte_ultimo_dia(xlsx: Path, n_dias: int, hora_corte: int) -> None:
    """Horas posteriores al corte = 0 en el último día de cada periodo."""
    if hora_corte >= 23 or n_dias < 1:
        return
    wb = load_workbook(xlsx)
    ws = wb["Consolidado"]
    col_p1 = 1 + n_dias  # A=hora, B… = días periodo 1
    col_p2 = 1 + n_dias + n_dias
    for col in (col_p1, col_p2):
        for h in range(hora_corte + 1, 24):
            ws.cell(row=4 + h, column=col, value=0.0)
    wb.save(xlsx)
    wb.close()


def _set_cell_shading(cell, hex_fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in tc_pr.findall(qn("w:shd")):
        tc_pr.remove(old)
    tc_pr.append(
        parse_xml(
            "<w:shd xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
            f'w:val="clear" w:fill="{hex_fill}"/>'
        )
    )


def _rendimiento(m3_sin: float, m3_con: float) -> Tuple[float, float]:
    ahorro = m3_sin - m3_con
    pct = (ahorro / m3_sin * 100.0) if m3_sin > 1e-9 else 0.0
    return ahorro, pct


def _hallazgo(pct: float) -> str:
    if pct >= 15:
        return "WES baja el consumo"
    if pct > 0:
        return "Ahorro menor"
    return "Sin WES no gasta más"


def _grafico_4_puntos(
    nombres: List[str],
    m3_con: List[float],
    m3_sin: List[float],
    titulo: str,
    lab_con: str,
    lab_sin: str,
    out_png: Path,
) -> None:
    x = np.arange(len(nombres))
    w = 0.36
    fig, ax = plt.subplots(figsize=(11.4, 5.6))
    b1 = ax.bar(x - w / 2, m3_con, width=w, color=COLOR_WES, label=lab_con, zorder=2)
    b2 = ax.bar(x + w / 2, m3_sin, width=w, color=COLOR_SIN, label=lab_sin, zorder=2)
    ymax = max(list(m3_con) + list(m3_sin) + [1.0]) * 1.22
    ax.set_ylim(0, ymax)
    ax.set_xticks(list(x))
    ax.set_xticklabels(nombres, fontsize=10)
    ax.set_ylabel("Consumo (m³)")
    ax.set_title(titulo, fontweight="bold", fontsize=12, color="#1F4788")
    ax.grid(axis="y", alpha=0.3, zorder=0)
    ax.legend(fontsize=9, loc="upper left")
    for bars in (b1, b2):
        for bar in bars:
            h = bar.get_height()
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                h + ymax * 0.015,
                format_number_chilean(float(h), 1),
                ha="center",
                va="bottom",
                fontsize=8,
                color="#333333",
            )
    for i, (a, b) in enumerate(zip(m3_sin, m3_con)):
        _, pct = _rendimiento(float(a), float(b))
        ax.text(
            x[i],
            ymax * 0.96,
            f"{pct:+.0f} %",
            ha="center",
            va="top",
            fontsize=9,
            fontweight="bold",
            color="#548235" if pct > 0 else COLOR_SIN,
        )
    fig.tight_layout()
    fig.savefig(out_png, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _grafico_perfil_24h(
    vec_con: Sequence[float],
    vec_sin: Sequence[float],
    hora_corte: int,
    titulo: str,
    lab_con: str,
    lab_sin: str,
    out_png: Path,
    sombrear: Sequence[int] = (),
) -> None:
    h1 = 23 if hora_corte >= 23 else hora_corte
    horas = list(range(h1 + 1))
    y_con = [float(vec_con[h]) for h in horas]
    y_sin = [float(vec_sin[h]) for h in horas]
    fig, ax = plt.subplots(figsize=(11.4, 5.0))
    ax.plot(horas, y_con, color=COLOR_WES, linewidth=2.2, marker="o", markersize=4, label=lab_con, zorder=3)
    ax.plot(horas, y_sin, color=COLOR_SIN, linewidth=2.2, marker="o", markersize=4, label=lab_sin, zorder=3)
    for h in sombrear:
        if h <= h1:
            ax.axvspan(h - 0.5, h + 0.5, color="#FFF2CC", alpha=0.55, zorder=0)
    ax.set_xticks(list(range(0, h1 + 1, 1 if h1 <= 16 else 2)))
    ax.set_xlabel("Hora Chile (sombreado = regulación 11:00–13:30)")
    ax.set_ylabel("m³/h")
    ax.set_title(titulo, fontweight="bold", color="#1F4788")
    ax.grid(axis="y", alpha=0.3)
    ax.legend(fontsize=8, loc="upper left")
    ax.set_ylim(bottom=0)
    fig.tight_layout()
    fig.savefig(out_png, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _suma_hasta(vec: Sequence[float], hora_corte: int) -> float:
    h1 = 23 if hora_corte >= 23 else hora_corte
    return float(sum(float(vec[h]) for h in range(h1 + 1)))


def _horas_noche_icco(hora_corte: int) -> List[int]:
    """Horas 00–05 que ya cerraron (si el corte es antes de las 06:00, la noche va incompleta)."""
    h1 = 23 if hora_corte >= 23 else hora_corte
    return [h for h in sorted(HORAS_CONTROL_ICCO) if h <= h1]


def _suma_noche_icco(vec: Sequence[float], hora_corte: int = 23) -> float:
    return float(sum(float(vec[h]) for h in _horas_noche_icco(hora_corte)))


def _preparar_icco(
    dias_con: Sequence[date],
    dias_sin: Sequence[date],
    hora_corte: int,
    png_dir: Path,
    xlsx_path: Path,
) -> dict:
    """Homóloga lun–vie con WES vs sin WES. El KPI es la noche 00:01–06:00."""
    print(f"  ICCO: descargando {len(dias_con) + len(dias_sin) + 1} días (homóloga + dom 16)…")
    v16 = _vector_m3h_24_desde_api(NODO_ICCO, DOMINGO_SIN_ICCO)
    m3_16_full = float(sum(v16))
    testigo_noche_full = _suma_noche_icco(v16, 23)
    vecs: Dict[date, List[float]] = {DOMINGO_SIN_ICCO: v16}
    for d in list(dias_con) + list(dias_sin):
        vecs[d] = _vector_m3h_24_desde_api(NODO_ICCO, d)

    filas: List[dict] = []
    for i, (d_con, d_sin) in enumerate(zip(dias_con, dias_sin)):
        incompleto = i == len(dias_con) - 1 and hora_corte < 23
        hc = hora_corte if incompleto else 23
        vc, vs = vecs[d_con], vecs[d_sin]
        m3_con = _suma_hasta(vc, hc)
        m3_sin = _suma_hasta(vs, hc)
        noche_con = _suma_noche_icco(vc, 23 if hc >= 5 else hc)
        noche_sin = _suma_noche_icco(vs, 23 if hc >= 5 else hc)
        a_dia, p_dia = _rendimiento(m3_sin, m3_con)
        a_noc, p_noc = _rendimiento(noche_sin, noche_con)
        filas.append(
            {
                "d_con": d_con,
                "d_sin": d_sin,
                "wd": WD_LARGO[d_con.weekday()],
                "m3_con": m3_con,
                "m3_sin": m3_sin,
                "ahorro_dia": a_dia,
                "pct_dia": p_dia,
                "noche_con": noche_con,
                "noche_sin": noche_sin,
                "ahorro": a_noc,
                "pct": p_noc,
                "incompleto": incompleto,
                "hora_corte": hc,
                "vec_con": vc,
                "vec_sin": vs,
            }
        )

    tot_con = float(sum(f["m3_con"] for f in filas))
    tot_sin = float(sum(f["m3_sin"] for f in filas))
    ahorro_dia, pct_dia = _rendimiento(tot_sin, tot_con)
    tot_noche_con = float(sum(f["noche_con"] for f in filas))
    tot_noche_sin = float(sum(f["noche_sin"] for f in filas))
    ahorro_t, pct_t = _rendimiento(tot_noche_sin, tot_noche_con)
    n_cero = sum(1 for f in filas if f["noche_con"] < 1.0)
    noches_sin_parejas = all(f["noche_sin"] >= 7.0 for f in filas)
    inc_pct = (ahorro_t / tot_sin * 100.0) if tot_sin > 1e-9 else 0.0

    png_dir.mkdir(parents=True, exist_ok=True)
    labels = []
    for f in filas:
        suf = f"\nhasta {f['hora_corte']:02d}:59" if f["incompleto"] else ""
        labels.append(f"{WD_CORTO[f['d_con'].weekday()]}\n{f['d_con']:%d} vs {f['d_sin']:%d}{suf}")

    png_barras = png_dir / "icco_noche_homologa_00_06.png"
    fig, ax = plt.subplots(figsize=(11.4, 5.2))
    x = np.arange(len(labels))
    w = 0.36
    ax.bar(x - w / 2, [f["noche_con"] for f in filas], width=w, color=COLOR_WES, label="Noche con WES", zorder=2)
    ax.bar(x + w / 2, [f["noche_sin"] for f in filas], width=w, color=COLOR_SIN, label="Noche sin WES", zorder=2)
    ax.axhline(
        testigo_noche_full,
        color="#C0504D",
        linestyle="--",
        linewidth=1.3,
        label=f"Fuga dom 16 ({format_number_chilean(testigo_noche_full, 1)} m³)",
    )
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, fontsize=8)
    ax.set_ylabel("m³ en 00:01–06:00")
    ax.set_title("ICCO — noche homóloga (esta semana sin control todos los días)", fontweight="bold", color="#1F4788")
    ax.grid(axis="y", alpha=0.3, zorder=0)
    ax.legend(fontsize=8, loc="upper left")
    ax.set_ylim(bottom=0)
    fig.tight_layout()
    fig.savefig(png_barras, dpi=150, bbox_inches="tight")
    plt.close(fig)

    png_dias = png_dir / "icco_dia_homologo.png"
    fig, ax = plt.subplots(figsize=(11.4, 5.2))
    ax.bar(x - w / 2, [f["m3_con"] for f in filas], width=w, color=COLOR_WES, label="Día con WES", zorder=2)
    ax.bar(x + w / 2, [f["m3_sin"] for f in filas], width=w, color=COLOR_SIN, label="Día sin WES", zorder=2)
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, fontsize=8)
    ax.set_ylabel("m³ día completo")
    ax.set_title("ICCO — día homólogo (ocupación: esta semana gasta menos de día)", fontweight="bold", color="#1F4788")
    ax.grid(axis="y", alpha=0.3, zorder=0)
    ax.legend(fontsize=8, loc="upper left")
    ax.set_ylim(bottom=0)
    fig.tight_layout()
    fig.savefig(png_dias, dpi=150, bbox_inches="tight")
    plt.close(fig)

    png_perfil = png_dir / "icco_perfil_homologo.png"
    f_cero = next((f for f in filas if f["noche_con"] < 1.0), filas[-2] if len(filas) > 1 else filas[0])
    fig, ax = plt.subplots(figsize=(12.2, 5.0))
    horas = list(range(24))
    ax.plot(horas, v16, color="#C0504D", linewidth=1.6, linestyle=":", label="Dom 16 sin WES (fuga pareja)")
    ax.plot(horas, f_cero["vec_con"], color=COLOR_WES, linewidth=2.2, label=f"{WD_CORTO[f_cero['d_con'].weekday()]} {f_cero['d_con']:%d/%m} con WES")
    ax.plot(horas, f_cero["vec_sin"], color=COLOR_SIN, linewidth=2.2, label=f"{WD_CORTO[f_cero['d_sin'].weekday()]} {f_cero['d_sin']:%d/%m} sin WES")
    ax.axvspan(-0.5, 5.5, color="#FFF2CC", alpha=0.5, zorder=0)
    ax.set_xticks(list(range(0, 24, 2)))
    ax.set_xlabel("Hora Chile (sombreado = 00:01–06:00)")
    ax.set_ylabel("m³/h")
    ax.set_title("ICCO — de noche esta semana iguala la fuga; de día hay menos ocupación", fontweight="bold", color="#1F4788")
    ax.grid(axis="y", alpha=0.3)
    ax.legend(fontsize=8, loc="upper left")
    ax.set_ylim(bottom=0)
    fig.tight_layout()
    fig.savefig(png_perfil, dpi=150, bbox_inches="tight")
    plt.close(fig)

    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    wb = Workbook()
    ws = wb.active
    ws.title = "ICCO_homologa"
    headers = [
        "Día",
        "Con WES",
        "Sin WES",
        "Día con (m³)",
        "Día sin (m³)",
        "Ahorro día (m³)",
        "% día",
        "Noche 00–06 con (m³)",
        "Noche 00–06 sin (m³)",
        "Ahorro noche (m³)",
        "% noche",
        "Notas",
    ]
    head_font = XlFont(bold=True, color="FFFFFF", name="Calibri")
    fill_h = PatternFill("solid", fgColor="1F4788")
    fill_ok = PatternFill("solid", fgColor="E2EFDA")
    fill_no = PatternFill("solid", fgColor="F8CBAD")
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        c.font = head_font
        c.fill = fill_h
        c.alignment = Alignment(wrap_text=True, horizontal="center")
        c.border = thin
    for r, f in enumerate(filas, start=2):
        nota = "Noche sin WES ≈ fuga del domingo 16" if f["noche_sin"] >= 7 else ""
        if f["noche_con"] < 1:
            nota = (nota + "; " if nota else "") + "WES dejó la noche en 0"
        if f["incompleto"]:
            nota += f"; hasta {f['hora_corte']:02d}:59 Chile"
        vals = [
            f["wd"],
            f["d_con"].strftime("%d/%m/%Y"),
            f["d_sin"].strftime("%d/%m/%Y"),
            round(f["m3_con"], 2),
            round(f["m3_sin"], 2),
            round(f["ahorro_dia"], 2),
            round(f["pct_dia"], 1),
            round(f["noche_con"], 2),
            round(f["noche_sin"], 2),
            round(f["ahorro"], 2),
            round(f["pct"], 1),
            nota,
        ]
        fill = fill_ok if f["pct"] >= 15 else fill_no
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            c.border = thin
            c.fill = fill
    rr = 2 + len(filas)
    ws.cell(rr, 1, "TOTAL").font = XlFont(bold=True)
    ws.cell(rr, 4, round(tot_con, 2)).font = XlFont(bold=True)
    ws.cell(rr, 5, round(tot_sin, 2)).font = XlFont(bold=True)
    ws.cell(rr, 6, round(ahorro_dia, 2)).font = XlFont(bold=True)
    ws.cell(rr, 7, round(pct_dia, 1)).font = XlFont(bold=True)
    ws.cell(rr, 8, round(tot_noche_con, 2)).font = XlFont(bold=True)
    ws.cell(rr, 9, round(tot_noche_sin, 2)).font = XlFont(bold=True)
    ws.cell(rr, 10, round(ahorro_t, 2)).font = XlFont(bold=True)
    ws.cell(rr, 11, round(pct_t, 1)).font = XlFont(bold=True)
    ws2 = wb.create_sheet("Criterio")
    ws2["A1"] = "Qué cambió"
    ws2["A1"].font = XlFont(bold=True)
    ws2["B1"] = TEXTO_SALA_BOMBAS
    ws2["A2"] = "Dato a mostrar"
    ws2["A2"].font = XlFont(bold=True)
    ws2["B2"] = (
        f"Rendimiento ICCO ahora: noche homóloga {pct_t:.1f} % ({ahorro_t:.1f} m³). "
        f"Día completo {pct_dia:.1f} % no se usa (ocupación). "
        f"Fuga domingo 16: {m3_16_full:.1f} m³/día, {testigo_noche_full:.1f} m³ de 00:01 a 06:00. "
        "Semana del 24/08: repetir contra esta semana."
    )
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 110
    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions["L"].width = 48
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(xlsx_path)

    return {
        "filas": filas,
        "m3_16": m3_16_full,
        "m3_9": 0.0,
        "ahorro_dom": 0.0,
        "pct_dom": 0.0,
        "testigo_noche": testigo_noche_full,
        "tot_con": tot_con,
        "tot_sin": tot_sin,
        "ahorro_dia": ahorro_dia,
        "pct_dia": pct_dia,
        "tot_noche_con": tot_noche_con,
        "tot_noche_sin": tot_noche_sin,
        "ahorro_t": ahorro_t,
        "pct_t": pct_t,
        "n_noches": len(filas),
        "n_cero": n_cero,
        "vol_dias_wes": tot_con,
        "inc_pct": inc_pct,
        "sin_control_esta_semana": noches_sin_parejas,
        "png_barras": png_barras,
        "png_dias": png_dias,
        "png_perfil": png_perfil,
        "xlsx": xlsx_path,
    }


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


def _auditoria_un_punto(
    node_id: str,
    nombre: str,
    cliente_dir: Path,
    dias_con: Sequence[date],
    dias_sin: Sequence[date],
    hora_corte: int,
    per_con: Periodo,
    per_sin: Periodo,
    titulo_p1: str,
    titulo_p2: str,
) -> Tuple[Path, Optional[Path], float, float, List[float], List[float]]:
    cliente_dir.mkdir(parents=True, exist_ok=True)
    a, b = dias_con[0], dias_con[-1]
    c, d = dias_sin[0], dias_sin[-1]
    xlsx_out = (
        cliente_dir
        / f"consumo_consolidado_conWES_{a:%d%m}_sinWES_{c:%d%m}.xlsx"
    )
    print(f"  CSV + Excel {node_id} {nombre}…")
    generar_excel_consolidado(
        xlsx_out,
        node_id,
        dias_con,
        dias_sin,
        titulo_p1=titulo_p1,
        titulo_p2=titulo_p2,
        solo_desde_csv=False,
        csv_dir=cliente_dir / "csv_descarga_api",
    )
    _aplicar_corte_ultimo_dia(xlsx_out, len(dias_con), hora_corte)
    fechas, mats = leer_matriz_consolidado(xlsx_out)
    mid = len(fechas) // 2
    m3_con = float(sum(sum(col) for col in mats[:mid]))
    m3_sin = float(sum(sum(col) for col in mats[mid:]))
    dias_c = [float(sum(col)) for col in mats[:mid]]
    dias_s = [float(sum(col)) for col in mats[mid:]]
    return xlsx_out, None, m3_con, m3_sin, dias_c, dias_s


def _word_conjunto(
    out_docx: Path,
    hora_corte: int,
    ahora: datetime,
    dias_con: Sequence[date],
    dias_sin: Sequence[date],
    pngs: Dict[str, Path],
    m3_con: Dict[str, float],
    m3_sin: Dict[str, float],
    dias_m3_con: Dict[str, List[float]],
    dias_m3_sin: Dict[str, List[float]],
    dirs: Dict[str, Path],
    icco_info: Optional[dict] = None,
    horas_ultimo: Optional[Dict[str, Tuple[List[float], List[float]]]] = None,
) -> None:
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    try:
        add_logo_to_header(doc)
    except Exception:
        pass

    h = doc.add_heading("Informe de Auditoría — Renca (5 puntos, control activo)", level=0)
    if h.runs:
        h.runs[0].font.color.rgb = COLOR_HEAD
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r = p.add_run(
        f"Esta semana CON WES ({_fmt_rango_dias(dias_con)})  vs  "
        f"semana pasada SIN WES ({_fmt_rango_dias(dias_sin)})"
    )
    r.bold = True
    r.font.color.rgb = COLOR_HEAD

    horas = _etiqueta_horas(hora_corte)
    ultimo = WD_LARGO[dias_con[-1].weekday()]
    p = doc.add_paragraph()
    p.add_run("Comparación día a día. ").bold = True
    p.add_run(
        f"{WD_LARGO[dias_con[0].weekday()]} a {ultimo} "
        f"({horas} Chile el último día; corte {ahora:%d/%m/%Y %H:%M} Chile). "
        "Cada día con control se compara con el mismo día de la semana del 17/08 sin WES "
        "(lun 24 vs lun 17, mar 25 vs mar 18, …). "
        "Ahorro = (Sin WES − Con WES) / Sin WES × 100. Tarifa de referencia 1.300 CLP/m³. "
        + TEXTO_CONTROL
    )

    doc.add_heading("1. Los 5 puntos", level=1)
    if "todos" in pngs:
        doc.add_picture(str(pngs["todos"]), width=Cm(16.2))

    tot_con = sum(m3_con.values())
    tot_sin = sum(m3_sin.values())
    ahorro_t, pct_t = _rendimiento(tot_sin, tot_con)
    doc.add_paragraph(
        f"5 puntos: Con WES {format_number_chilean(tot_con, 1)} m³; "
        f"Sin WES {format_number_chilean(tot_sin, 1)} m³; "
        f"diferencia {format_number_chilean(ahorro_t, 1)} m³ "
        f"({format_number_chilean(pct_t, 1)} %; "
        f"~{format_number_chilean(max(0.0, ahorro_t) * TARIFA_CLP_M3, 0)} CLP)."
    )

    n_extra = 1
    tbl = doc.add_table(rows=1 + len(PUNTOS) + n_extra, cols=6)
    tbl.style = "Table Grid"
    headers = ["Punto", "Con WES (m³)", "Sin WES (m³)", "Ahorro (m³)", "%", "Hallazgo"]
    for j, hd in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        cell.text = hd
        _set_cell_shading(cell, "1F4788")
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9)

    for i, (nid, nom) in enumerate(PUNTOS, start=1):
        ahorro, pct = _rendimiento(m3_sin[nid], m3_con[nid])
        row = tbl.rows[i]
        vals = [
            nom,
            format_number_chilean(m3_con[nid], 1),
            format_number_chilean(m3_sin[nid], 1),
            format_number_chilean(ahorro, 1),
            format_number_chilean(pct, 1) + " %",
            _hallazgo(pct),
        ]
        for j, v in enumerate(vals):
            row.cells[j].text = v
            for run in row.cells[j].paragraphs[0].runs:
                run.font.size = Pt(9)
        if pct >= 15:
            for c in row.cells:
                _set_cell_shading(c, "E2EFDA")

    row = tbl.rows[1 + len(PUNTOS)]
    vals = [
        "Total 5 puntos",
        format_number_chilean(tot_con, 1),
        format_number_chilean(tot_sin, 1),
        format_number_chilean(ahorro_t, 1),
        format_number_chilean(pct_t, 1) + " %",
        "",
    ]
    for j, v in enumerate(vals):
        row.cells[j].text = v
        _set_cell_shading(row.cells[j], "D6DCE4")
        for run in row.cells[j].paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(9)

    nomb_d: List[str] = []
    for i, (d1, d2) in enumerate(zip(dias_con, dias_sin)):
        extra = f" ({horas})" if i == len(dias_con) - 1 and hora_corte < 23 else ""
        nomb_d.append(f"{WD_CORTO[d1.weekday()]} {d1:%d} vs {d2:%d}{extra}")

    for nid, nom in PUNTOS:
        doc.add_heading(nom, level=1)
        ahorro, pct = _rendimiento(m3_sin[nid], m3_con[nid])
        doc.add_paragraph(
            f"Con WES {format_number_chilean(m3_con[nid], 1)} m³ vs "
            f"Sin WES {format_number_chilean(m3_sin[nid], 1)} m³ → "
            f"{format_number_chilean(pct, 1)} % "
            f"({format_number_chilean(ahorro, 1)} m³)."
        )
        t2 = doc.add_table(rows=1 + len(nomb_d), cols=5)
        t2.style = "Table Grid"
        for j, hd in enumerate(["Día homólogo", "Con WES (m³)", "Sin WES (m³)", "Ahorro (m³)", "%"]):
            cell = t2.rows[0].cells[j]
            cell.text = hd
            _set_cell_shading(cell, "1F4788")
            for run in cell.paragraphs[0].runs:
                run.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(9)
        for i, lab in enumerate(nomb_d):
            a, pc = _rendimiento(dias_m3_sin[nid][i], dias_m3_con[nid][i])
            vals = [
                lab,
                format_number_chilean(dias_m3_con[nid][i], 1),
                format_number_chilean(dias_m3_sin[nid][i], 1),
                format_number_chilean(a, 1),
                format_number_chilean(pc, 1) + " %",
            ]
            for j, v in enumerate(vals):
                t2.rows[i + 1].cells[j].text = v
                for run in t2.rows[i + 1].cells[j].paragraphs[0].runs:
                    run.font.size = Pt(9)

    if horas_ultimo:
        doc.add_heading("2. Regulación de hoy 11:00–13:30 (Lo Velásquez y gimnasio)", level=1)
        p = doc.add_paragraph()
        p.add_run("Qué se hizo. ").bold = True
        p.add_run(
            "Hoy 24/08 entre 11:00 y 13:30 se cargaron las mejoras de horario. "
            "En el gimnasio el 100 % cubre el pico real (no el 0,54 m³/h parejo). "
            "En Lo Velásquez la noche sigue en 0 L/min; de día el tope deja pasar el uso del colegio. "
            "El sombreado amarillo en los perfiles es 11:00–13:59 (cubre ese tramo)."
        )
        for nid, nom, key in (
            (NODO_ESCUELA, "Esc. Lo Velásquez", "escuela"),
            (NODO_GIMNASIO, "Gimnasio municipal", "gimnasio"),
        ):
            if key in pngs and Path(pngs[key]).is_file():
                doc.add_paragraph(nom).runs[0].bold = True
                doc.add_picture(str(pngs[key]), width=Cm(16.2))
            vc, vs = horas_ultimo.get(nid, ([], []))
            if not vc:
                continue
            t3 = doc.add_table(rows=1 + len(HORAS_REGULADAS), cols=4)
            t3.style = "Table Grid"
            for j, hd in enumerate(["Hora", "Con WES hoy (m³/h)", "Sin WES lun 17 (m³/h)", "Δ"]):
                cell = t3.rows[0].cells[j]
                cell.text = hd
                _set_cell_shading(cell, "1F4788")
                for run in cell.paragraphs[0].runs:
                    run.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.size = Pt(9)
            for i, h in enumerate(HORAS_REGULADAS):
                a = float(vc[h]) if h < len(vc) else 0.0
                b = float(vs[h]) if h < len(vs) else 0.0
                vals = [
                    f"{h:02d}:00–{h:02d}:59",
                    format_number_chilean(a, 2),
                    format_number_chilean(b, 2),
                    format_number_chilean(b - a, 2),
                ]
                for j, v in enumerate(vals):
                    t3.rows[i + 1].cells[j].text = v
                    for run in t3.rows[i + 1].cells[j].paragraphs[0].runs:
                        run.font.size = Pt(9)
            doc.add_paragraph("")

    suben = [nom for nid, nom in PUNTOS if _rendimiento(m3_sin[nid], m3_con[nid])[1] >= 15]
    no_suben = [nom for nid, nom in PUNTOS if _rendimiento(m3_sin[nid], m3_con[nid])[1] <= 0]
    doc.add_heading("Conclusiones", level=1)
    doc.add_paragraph(
        "Los 5 puntos ya están con control desde el lunes 24. La comparación es día a día "
        "contra la semana del 17 sin WES. "
        + (
            f"Hoy el ahorro se concentra en {', '.join(suben)}. "
            if suben
            else "Hoy ningún punto muestra un ahorro claro todavía. "
        )
        + (
            f"En {', '.join(no_suben)} el día con WES no gasta menos que el lunes sin control. "
            if no_suben
            else ""
        )
        + "Lo Velásquez y el gimnasio se regularon a las 11:00–13:30: el gimnasio a las 12:00 "
        "tuvo más caudal que el lunes 17 (evento); el tope nuevo lo deja pasar. "
        "Se sigue el resto de la semana con el mismo corte horario."
    )
    out_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_docx)


def main() -> int:
    ahora = datetime.now(TZ_CL)
    dias_con, dias_sin, hora_corte, _inc = _ventana_homologa(ahora)
    ts = ahora.strftime("%Y%m%d_%H%M")
    lab_con = f"Con WES {_fmt_rango_dias(dias_con)}"
    lab_sin = f"Sin WES {_fmt_rango_dias(dias_sin)}"
    per_con = Periodo(f"Con control ({dias_con[0]:%d-%m} a {dias_con[-1]:%d-%m-%Y})", tuple(dias_con))
    per_sin = Periodo(f"Sin control ({dias_sin[0]:%d-%m} a {dias_sin[-1]:%d-%m-%Y})", tuple(dias_sin))
    titulo_p1 = f"Con WES: {dias_con[0]:%d-%m-%Y} al {dias_con[-1]:%d-%m-%Y}"
    titulo_p2 = f"Sin WES: {dias_sin[0]:%d-%m-%Y} al {dias_sin[-1]:%d-%m-%Y}"

    out_root = (
        ROOT
        / "reports"
        / "reporte de auditoria"
        / f"auditoria_renca_semana_pasada_vs_esta_{ts}"
    )
    out_root.mkdir(parents=True, exist_ok=True)
    print(
        f"Auditoría Renca | {lab_con} vs {lab_sin} | "
        f"último día hasta {hora_corte:02d}:59 Chile | {ahora:%Y-%m-%d %H:%M}"
    )
    print("5 puntos con control desde lun 24 vs semana del 17 sin WES")

    gxlsx.LABEL_P1 = lab_con
    gxlsx.LABEL_P2 = lab_sin

    m3_con: Dict[str, float] = {}
    m3_sin: Dict[str, float] = {}
    dias_m3_con: Dict[str, List[float]] = {}
    dias_m3_sin: Dict[str, List[float]] = {}
    dirs: Dict[str, Path] = {}
    horas_ultimo: Dict[str, Tuple[List[float], List[float]]] = {}

    for nid, nom in PUNTOS:
        print("=" * 64)
        print(f"{nid} {nom}")
        cdir = out_root / f"Auditoria_{nom.split('(')[0].strip().replace(' ', '_')}_{nid}"
        dirs[nid] = cdir
        xlsx, _pdf, c, s, dc, ds = _auditoria_un_punto(
            nid,
            nom,
            cdir,
            dias_con,
            dias_sin,
            hora_corte,
            per_con,
            per_sin,
            titulo_p1,
            titulo_p2,
        )
        m3_con[nid] = c
        m3_sin[nid] = s
        dias_m3_con[nid] = dc
        dias_m3_sin[nid] = ds
        _fechas, mats = leer_matriz_consolidado(xlsx)
        n = len(dias_con)
        horas_ultimo[nid] = (list(mats[n - 1]), list(mats[2 * n - 1]))
        print(f"  Con {c:.1f} m³ | Sin {s:.1f} m³")

    png_dir = out_root / "graficos"
    png_dir.mkdir(exist_ok=True)
    p_all = png_dir / "todos_casos_con_vs_sin.png"
    _grafico_4_puntos(
        ["ICCP", "Lo Velásquez", "Gimnasio", "Piscina", "ICCO"],
        [m3_con[n] for n, _ in PUNTOS],
        [m3_sin[n] for n, _ in PUNTOS],
        f"Renca — {lab_con} vs {lab_sin}",
        lab_con,
        lab_sin,
        p_all,
    )
    pngs: Dict[str, Path] = {"todos": p_all}
    d1, d2 = dias_con[-1], dias_sin[-1]
    for nid, key, nom in (
        (NODO_ESCUELA, "escuela", "Lo Velásquez"),
        (NODO_GIMNASIO, "gimnasio", "Gimnasio"),
        (NODO_ICCO, "icco", "ICCO"),
    ):
        png = png_dir / f"perfil_{key}_{d1:%d%m}_vs_{d2:%d%m}.png"
        vc, vs = horas_ultimo[nid]
        _grafico_perfil_24h(
            vc,
            vs,
            hora_corte,
            f"{nom} — {WD_CORTO[d1.weekday()]} {d1:%d/%m} con WES vs {d2:%d/%m} sin WES",
            lab_con,
            lab_sin,
            png,
            HORAS_REGULADAS if nid != NODO_ICCO else (),
        )
        pngs[key] = png

    docx_path = out_root / f"Informe_Auditoria_Renca_semana_pasada_vs_esta_{ts}.docx"
    _word_conjunto(
        docx_path,
        hora_corte,
        ahora,
        dias_con,
        dias_sin,
        pngs,
        m3_con,
        m3_sin,
        dias_m3_con,
        dias_m3_sin,
        dirs,
        None,
        horas_ultimo,
    )
    pdf_path = _convertir_pdf(docx_path)

    tot_c = sum(m3_con.values())
    tot_s = sum(m3_sin.values())
    ahorro, pct = _rendimiento(tot_s, tot_c)
    print("=" * 64)
    print(f"5 puntos Con {tot_c:.1f} | Sin {tot_s:.1f} | {pct:.1f}% ({ahorro:.1f} m³)")
    for nid, nom in PUNTOS:
        a, p = _rendimiento(m3_sin[nid], m3_con[nid])
        print(f"  {nom}: {p:.1f}% ({a:.1f} m³)")
    print(f"DOCX conjunto: {docx_path}")
    print(f"PDF  conjunto: {pdf_path or '(no convertido)'}")
    print(f"DIR: {out_root}")
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except Exception:
            pass
    raise SystemExit(main())
