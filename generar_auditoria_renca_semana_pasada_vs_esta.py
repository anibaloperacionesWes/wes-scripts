"""
Informe de auditoría Renca — semana pasada (con WES) vs esta semana (sin WES).

Semana pasada: lun 10/08/2026 (última semana con control).
Esta semana: lun 17/08/2026 (sin WES desde el 17 en ICCP, Lo Velásquez, gimnasio y piscina).
Colegio ICCO (000017-08): solo sin control el domingo 16/08 (estanque sala de bombas).

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
)
NODO_ICCO = "000017-08"
NOMBRE_ICCO = "Colegio ICCO Renca"
DOMINGO_SIN_ICCO = date(2026, 8, 16)
DOMINGO_CON_ICCO = date(2026, 8, 9)
LUNES_CON = date(2026, 8, 10)
LUNES_SIN = date(2026, 8, 17)
TEXTO_SALA_BOMBAS = (
    "En Colegio ICCO Renca (000017-08) no se puede dejar el punto sin control WES "
    "durante una semana completa. Si se elimina el control, el estanque de la sala "
    "de bombas se vacía y no es capaz de recuperarse; eso afectaría el funcionamiento "
    "normal del colegio. Por esa razón el único día sin control fue el domingo 16/08/2026, "
    "de 00:00 a 23:59. El resto de los días permaneció con WES. El rendimiento se mide "
    "comparando ese domingo sin control con cada día con control y, en particular, con el "
    "domingo homólogo anterior (09/08). No existe una semana homóloga lun–vie sin WES "
    "en este punto: no se puede construir sin vaciar el estanque."
)
WD_CORTO = ("Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom")
WD_LARGO = ("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo")


def _ventana_homologa(ahora: datetime) -> Tuple[Tuple[date, ...], Tuple[date, ...], int, bool]:
    """Días homólogos lun…hoy (máx. 7) y hora de corte del último día (23 = completo)."""
    hoy = ahora.date()
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


def _grafico_dias_homologos(
    nombres_dia: List[str],
    series_con: Dict[str, List[float]],
    series_sin: Dict[str, List[float]],
    titulo: str,
    out_png: Path,
) -> None:
    fig, axes = plt.subplots(2, 2, figsize=(13.2, 7.8))
    for ax, (nid, nom) in zip(axes.ravel(), PUNTOS):
        x = np.arange(len(nombres_dia))
        w = 0.36
        ax.bar(x - w / 2, series_con[nid], width=w, color=COLOR_WES, label="Con WES", zorder=2)
        ax.bar(x + w / 2, series_sin[nid], width=w, color=COLOR_SIN, label="Sin WES", zorder=2)
        ax.set_xticks(list(x))
        ax.set_xticklabels(nombres_dia, fontsize=7)
        ax.set_ylabel("m³")
        ax.set_title(nom, fontsize=10, fontweight="bold", color="#1F4788")
        ax.grid(axis="y", alpha=0.3, zorder=0)
        ax.legend(fontsize=7, loc="upper left")
    fig.suptitle(titulo, fontsize=13, fontweight="bold", color="#1F4788")
    fig.tight_layout()
    fig.savefig(out_png, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _suma_hasta(vec: Sequence[float], hora_corte: int) -> float:
    h1 = 23 if hora_corte >= 23 else hora_corte
    return float(sum(float(vec[h]) for h in range(h1 + 1)))


def _icco_dias_a_comparar(hoy: date) -> List[date]:
    """Domingo 09 (homólogo con WES) + 10/08 hasta hoy."""
    out = [DOMINGO_CON_ICCO]
    d = LUNES_CON
    while d <= hoy:
        out.append(d)
        d += timedelta(days=1)
    return out


def _preparar_icco(
    hoy: date,
    hora_corte: int,
    png_dir: Path,
    xlsx_path: Path,
) -> dict:
    dias = _icco_dias_a_comparar(hoy)
    vecs: Dict[date, List[float]] = {}
    print(f"  ICCO: descargando {len(dias)} días…")
    for d in dias:
        vecs[d] = _vector_m3h_24_desde_api(NODO_ICCO, d)
    v16 = vecs[DOMINGO_SIN_ICCO]
    m3_16_full = float(sum(v16))

    filas: List[dict] = []
    for d in dias:
        v = vecs[d]
        incompleto = d == hoy and hora_corte < 23
        hc = hora_corte if incompleto else 23
        m3 = _suma_hasta(v, hc)
        ref = _suma_hasta(v16, hc)
        ahorro, pct = _rendimiento(ref, m3)
        filas.append(
            {
                "fecha": d,
                "wd": WD_LARGO[d.weekday()],
                "control": "Sin WES" if d == DOMINGO_SIN_ICCO else "Con WES",
                "m3": m3,
                "ref": ref,
                "ahorro": ahorro,
                "pct": pct,
                "incompleto": incompleto,
                "hora_corte": hc,
            }
        )

    v9 = vecs[DOMINGO_CON_ICCO]
    m3_9 = float(sum(v9))
    ahorro_dom, pct_dom = _rendimiento(m3_16_full, m3_9)

    png_dir.mkdir(parents=True, exist_ok=True)
    png_barras = png_dir / "icco_dom16_vs_cada_dia.png"
    labels = []
    vals = []
    colors = []
    for f in filas:
        suf = f"\nhasta {f['hora_corte']:02d}:59" if f["incompleto"] else ""
        labels.append(f"{WD_CORTO[f['fecha'].weekday()]}\n{f['fecha']:%d/%m}{suf}")
        vals.append(f["m3"])
        colors.append(COLOR_SIN if f["fecha"] == DOMINGO_SIN_ICCO else COLOR_WES)
    fig, ax = plt.subplots(figsize=(12.8, 5.4))
    x = np.arange(len(vals))
    bars = ax.bar(x, vals, color=colors, zorder=2)
    ax.axhline(m3_16_full, color=COLOR_SIN, linestyle="--", linewidth=1.2, label="Dom 16 sin WES (día completo)")
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, fontsize=7)
    ax.set_ylabel("m³")
    ax.set_title(
        "ICCO — cada día vs domingo 16/08 sin control",
        fontweight="bold",
        fontsize=12,
        color="#1F4788",
    )
    ax.grid(axis="y", alpha=0.3, zorder=0)
    ax.legend(fontsize=8, loc="upper left")
    ymax = max(vals + [m3_16_full, 1.0]) * 1.18
    ax.set_ylim(0, ymax)
    for bar, v in zip(bars, vals):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + ymax * 0.012,
            format_number_chilean(float(v), 1),
            ha="center",
            va="bottom",
            fontsize=7,
        )
    fig.tight_layout()
    fig.savefig(png_barras, dpi=150, bbox_inches="tight")
    plt.close(fig)

    png_perfil = png_dir / "icco_perfil_dom9_dom16_mar11.png"
    fig, ax = plt.subplots(figsize=(12.2, 5.0))
    horas = list(range(24))
    ax.plot(horas, v9, color=COLOR_WES, linewidth=2.2, label=f"Dom 09/08 con WES ({format_number_chilean(m3_9, 1)} m³)")
    ax.plot(
        horas,
        v16,
        color=COLOR_SIN,
        linewidth=2.2,
        label=f"Dom 16/08 sin WES ({format_number_chilean(m3_16_full, 1)} m³)",
    )
    v11 = vecs.get(date(2026, 8, 11))
    if v11:
        ax.plot(
            horas,
            v11,
            color="#7B7B7B",
            linewidth=1.6,
            linestyle="--",
            label=f"Mar 11/08 con WES, día de clases ({format_number_chilean(sum(v11), 1)} m³)",
        )
    ax.set_xticks(list(range(0, 24, 2)))
    ax.set_xlabel("Hora Chile")
    ax.set_ylabel("m³/h")
    ax.set_title("ICCO — perfil 24 h: domingo sin control vs domingo y martes con WES", fontweight="bold", color="#1F4788")
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
    ws.title = "ICCO_dom16_vs_cada_dia"
    headers = [
        "Fecha",
        "Día",
        "Control",
        "m³ del día",
        "m³ dom 16 (misma ventana)",
        "Ahorro m³ (dom16 − día)",
        "Rendimiento %",
        "Notas",
    ]
    head_font = XlFont(bold=True, color="FFFFFF", name="Calibri")
    fill_h = PatternFill("solid", fgColor="1F4788")
    fill_sin = PatternFill("solid", fgColor="F8CBAD")
    fill_con = PatternFill("solid", fgColor="D6EAF8")
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        c.font = head_font
        c.fill = fill_h
        c.alignment = Alignment(wrap_text=True, horizontal="center")
        c.border = thin
    for r, f in enumerate(filas, start=2):
        nota = "Único día sin control (00:00–23:59)" if f["fecha"] == DOMINGO_SIN_ICCO else "Con WES"
        if f["incompleto"]:
            nota += f"; hasta {f['hora_corte']:02d}:59 Chile"
        if f["fecha"] == DOMINGO_CON_ICCO:
            nota += "; domingo homólogo"
        vals = [
            f["fecha"].strftime("%d/%m/%Y"),
            f["wd"],
            f["control"],
            round(f["m3"], 2),
            round(f["ref"], 2),
            round(f["ahorro"], 2),
            round(f["pct"], 1),
            nota,
        ]
        fill = fill_sin if f["fecha"] == DOMINGO_SIN_ICCO else fill_con
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            c.border = thin
            c.fill = fill
    ws2 = wb.create_sheet("Criterio")
    ws2["A1"] = "Sala de bombas"
    ws2["A1"].font = XlFont(bold=True)
    ws2["B1"] = TEXTO_SALA_BOMBAS
    ws2["A2"] = "Rendimiento homólogo"
    ws2["A2"].font = XlFont(bold=True)
    ws2["B2"] = (
        f"Domingo 09/08 con WES {m3_9:.2f} m³ vs domingo 16/08 sin WES {m3_16_full:.2f} m³ → "
        f"{pct_dom:.1f} % ({ahorro_dom:.2f} m³). Ahorro = (sin − con) / sin × 100."
    )
    ws2.column_dimensions["A"].width = 24
    ws2.column_dimensions["B"].width = 110
    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.column_dimensions["H"].width = 42
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(xlsx_path)

    return {
        "filas": filas,
        "m3_16": m3_16_full,
        "m3_9": m3_9,
        "ahorro_dom": ahorro_dom,
        "pct_dom": pct_dom,
        "png_barras": png_barras,
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

    gdir = cliente_dir / "graficos_comparativos"
    _limpiar_pngs_carpeta_graficos(gdir)
    fechas, mats = leer_matriz_consolidado(xlsx_out)
    generar_pngs(fechas, mats, gdir)

    mid = len(fechas) // 2
    m3_con = float(sum(sum(col) for col in mats[:mid]))
    m3_sin = float(sum(sum(col) for col in mats[mid:]))
    dias_c = [float(sum(col)) for col in mats[:mid]]
    dias_s = [float(sum(col)) for col in mats[mid:]]

    stem = f"Auditoria_{nombre.replace(' ', '_')}_{node_id}"
    docx_path = cliente_dir / f"{stem}.docx"
    graf.NOMBRE_PUNTO = nombre
    icco.NOMBRE_PUNTO = nombre
    icco.PORTADA_TITULO = "Informe de Auditoría"
    icco.PORTADA_REFERENCIA_BORRADOR = stem
    icco.PORTADA_ESTABLECIMIENTO_LINEA1 = nombre
    icco.PORTADA_ESTABLECIMIENTO_LINEA2 = node_id
    icco.CLP_POR_M3_REF = TARIFA_CLP_M3
    icco._ETIQUETA_COLEGIO_CUADRO_RESUMEN = nombre.upper()[:22]

    print(f"  Word {node_id}…")
    p = icco.generar_informe_word(
        node_id=node_id,
        out_dir=cliente_dir,
        output_docx=docx_path,
        mantener_borrador_manual=False,
        solo_consolidado=False,
        periodo_ref=per_con,
        periodo_aud=per_sin,
        figuras_desde_xlsx=xlsx_out,
    )
    pdf = _convertir_pdf(p)
    return p, pdf, m3_con, m3_sin, dias_c, dias_s


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
) -> None:
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    try:
        add_logo_to_header(doc)
    except Exception:
        pass

    h = doc.add_heading("Informe de Auditoría — Renca (5 puntos)", level=0)
    if h.runs:
        h.runs[0].font.color.rgb = COLOR_HEAD
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r = p.add_run(
        f"Semana pasada con WES ({_fmt_rango_dias(dias_con)})  vs  "
        f"esta semana sin WES ({_fmt_rango_dias(dias_sin)})"
    )
    r.bold = True
    r.font.color.rgb = COLOR_HEAD

    horas = _etiqueta_horas(hora_corte)
    ultimo = WD_LARGO[dias_con[-1].weekday()]
    p = doc.add_paragraph()
    p.add_run("Ventana homóloga: ").bold = True
    p.add_run(
        f"{WD_LARGO[dias_con[0].weekday()]} a {ultimo} "
        f"({horas} Chile el último día; corte {ahora:%d/%m/%Y %H:%M} Chile). "
        "Cada día de esta semana sin control se compara con el mismo día de la semana pasada con WES "
        "(lun 17 vs lun 10, mar 18 vs mar 11, mié 19 vs mié 12, jue 20 vs jue 13, vie 21 vs vie 14). "
        "Ahorro = (Sin WES − Con WES) / Sin WES × 100. Tarifa de referencia 1.300 CLP/m³. "
        "Sin WES desde el lunes 17/08/2026 en ICCP, Lo Velásquez, gimnasio y piscina. "
        "Colegio ICCO Renca se trata aparte: solo estuvo sin control el domingo 16/08 "
        "(estanque de la sala de bombas)."
    )

    doc.add_heading("1. ICCP, Lo Velásquez, gimnasio y piscina", level=1)
    if "todos" in pngs:
        doc.add_picture(str(pngs["todos"]), width=Cm(16.2))
    if "dias4" in pngs:
        doc.add_paragraph("")
        doc.add_picture(str(pngs["dias4"]), width=Cm(16.2))

    tot_con = sum(m3_con.values())
    tot_sin = sum(m3_sin.values())
    ahorro_t, pct_t = _rendimiento(tot_sin, tot_con)
    doc.add_paragraph(
        f"Conjunto: Con WES {format_number_chilean(tot_con, 1)} m³; "
        f"Sin WES {format_number_chilean(tot_sin, 1)} m³; "
        f"diferencia {format_number_chilean(ahorro_t, 1)} m³ "
        f"({format_number_chilean(pct_t, 1)} % sobre el periodo sin control; "
        f"~{format_number_chilean(max(0.0, ahorro_t) * TARIFA_CLP_M3, 0)} CLP)."
    )

    tbl = doc.add_table(rows=1 + len(PUNTOS) + 1, cols=6)
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
        "Total 4 puntos",
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

    slugs = ("Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo")
    for nid, nom in PUNTOS:
        doc.add_heading(nom, level=1)
        ahorro, pct = _rendimiento(m3_sin[nid], m3_con[nid])
        doc.add_paragraph(
            f"Con WES {format_number_chilean(m3_con[nid], 1)} m³ vs "
            f"Sin WES {format_number_chilean(m3_sin[nid], 1)} m³ → "
            f"{format_number_chilean(pct, 1)} % "
            f"({format_number_chilean(ahorro, 1)} m³)."
        )
        gdir = dirs[nid] / "graficos_comparativos"
        barras = gdir / "02_barras_total_rejilla_por_periodo.png"
        if barras.is_file():
            doc.add_picture(str(barras), width=Cm(12.5))
        for j, d1 in enumerate(dias_con):
            fname = f"{4 + j:02d}_area_{slugs[d1.weekday()]}.png"
            pth = gdir / fname
            if pth.is_file():
                doc.add_picture(str(pth), width=Cm(16.0))
        prom = gdir / "03_area_promedio_24h.png"
        if prom.is_file():
            doc.add_picture(str(prom), width=Cm(16.0))

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

    if icco_info:
        doc.add_heading("2. Quinto punto: Colegio ICCO Renca (000017-08)", level=1)
        p = doc.add_paragraph()
        p.add_run("Sala de bombas. ").bold = True
        p.add_run(TEXTO_SALA_BOMBAS)
        if icco_info.get("png_barras") and Path(icco_info["png_barras"]).is_file():
            doc.add_picture(str(icco_info["png_barras"]), width=Cm(16.2))
        if icco_info.get("png_perfil") and Path(icco_info["png_perfil"]).is_file():
            doc.add_paragraph("")
            doc.add_picture(str(icco_info["png_perfil"]), width=Cm(16.2))
        p = doc.add_paragraph()
        p.add_run("Rendimiento totalizado (único par homólogo justo). ").bold = True
        p.add_run(
            f"Domingo 09/08 con WES {format_number_chilean(icco_info['m3_9'], 1)} m³ vs "
            f"domingo 16/08 sin WES {format_number_chilean(icco_info['m3_16'], 1)} m³ → "
            f"{format_number_chilean(icco_info['pct_dom'], 1)} % "
            f"({format_number_chilean(icco_info['ahorro_dom'], 1)} m³; "
            f"~{format_number_chilean(max(0.0, icco_info['ahorro_dom']) * TARIFA_CLP_M3, 0)} CLP). "
            "No se puede extrapolar ese 40,1 m³/día a lun–vie: un día de clases con WES ya supera "
            "50–77 m³ por ocupación, y sin control el estanque de la sala de bombas se vaciaría."
        )
        filas = icco_info["filas"]
        t3 = doc.add_table(rows=1 + len(filas), cols=6)
        t3.style = "Table Grid"
        for j, hd in enumerate(["Día", "Control", "m³ del día", "m³ dom 16", "Ahorro (m³)", "%"]):
            cell = t3.rows[0].cells[j]
            cell.text = hd
            _set_cell_shading(cell, "1F4788")
            for run in cell.paragraphs[0].runs:
                run.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(9)
        for i, f in enumerate(filas, start=1):
            extra = f" hasta {f['hora_corte']:02d}:59" if f["incompleto"] else ""
            vals = [
                f"{f['wd']} {f['fecha']:%d/%m}{extra}",
                f["control"],
                format_number_chilean(f["m3"], 1),
                format_number_chilean(f["ref"], 1),
                format_number_chilean(f["ahorro"], 1),
                format_number_chilean(f["pct"], 1) + " %",
            ]
            for j, v in enumerate(vals):
                t3.rows[i].cells[j].text = v
                for run in t3.rows[i].cells[j].paragraphs[0].runs:
                    run.font.size = Pt(9)
            if f["fecha"] == DOMINGO_SIN_ICCO:
                for c in t3.rows[i].cells:
                    _set_cell_shading(c, "F8CBAD")
            elif f["fecha"] == DOMINGO_CON_ICCO:
                for c in t3.rows[i].cells:
                    _set_cell_shading(c, "E2EFDA")

    suben = [nom for nid, nom in PUNTOS if _rendimiento(m3_sin[nid], m3_con[nid])[1] >= 15]
    no_suben = [nom for nid, nom in PUNTOS if _rendimiento(m3_sin[nid], m3_con[nid])[1] <= 0]
    doc.add_heading("Conclusiones", level=1)
    doc.add_paragraph(
        "En ICCP, Lo Velásquez, gimnasio y piscina la comparación es homóloga: misma estación, "
        "mismos días de la semana y la última semana con control WES contra esta semana sin control. "
        + (
            f"El aumento de volumen se concentra en {', '.join(suben)}. "
            if suben
            else "Ningún punto de esos cuatro muestra un ahorro claro de WES en esta ventana. "
        )
        + (
            f"En {', '.join(no_suben)} el periodo sin WES no gasta más que la semana con control. "
            if no_suben
            else ""
        )
    )
    if icco_info:
        doc.add_paragraph(
            "En Colegio ICCO Renca el sin control no se puede repetir una semana completa: "
            "el estanque de la sala de bombas se vacía y no se recupera, y eso afectaría el "
            "funcionamiento normal del colegio. El único día sin WES fue el domingo 16/08 "
            f"({format_number_chilean(icco_info['m3_16'], 1)} m³). El rendimiento totalizado "
            f"frente al domingo homólogo 09/08 es {format_number_chilean(icco_info['pct_dom'], 1)} % "
            f"({format_number_chilean(icco_info['ahorro_dom'], 1)} m³). "
            "Los días de clases con WES gastan más que ese domingo por ocupación, no porque "
            "el control suba el consumo."
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
    print("Quinto punto: ICCO 000017-08 — único sin control: domingo 16/08 (sala de bombas)")

    gxlsx.LABEL_P1 = lab_con
    gxlsx.LABEL_P2 = lab_sin

    m3_con: Dict[str, float] = {}
    m3_sin: Dict[str, float] = {}
    dias_m3_con: Dict[str, List[float]] = {}
    dias_m3_sin: Dict[str, List[float]] = {}
    dirs: Dict[str, Path] = {}

    for nid, nom in PUNTOS:
        print("=" * 64)
        print(f"{nid} {nom}")
        cdir = out_root / f"Auditoria_{nom.split('(')[0].strip().replace(' ', '_')}_{nid}"
        dirs[nid] = cdir
        _docx, _pdf, c, s, dc, ds = _auditoria_un_punto(
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
        print(f"  Con {c:.1f} m³ | Sin {s:.1f} m³ | Word {_docx}")
        if _pdf:
            print(f"  PDF {_pdf}")

    png_dir = out_root / "graficos"
    png_dir.mkdir(exist_ok=True)
    p_all = png_dir / "todos_casos_con_vs_sin.png"
    _grafico_4_puntos(
        ["ICCP", "Lo Velásquez", "Gimnasio", "Piscina"],
        [m3_con[n] for n, _ in PUNTOS],
        [m3_sin[n] for n, _ in PUNTOS],
        f"Renca — auditoría: {lab_con} vs {lab_sin}",
        lab_con,
        lab_sin,
        p_all,
    )
    nombres_d = []
    for i, d1 in enumerate(dias_con):
        if i == len(dias_con) - 1 and hora_corte < 23:
            nombres_d.append(f"{WD_LARGO[d1.weekday()]}\nhasta {hora_corte:02d}:59")
        else:
            nombres_d.append(WD_LARGO[d1.weekday()])
    p_dias = png_dir / "dias_homologos_4_puntos.png"
    _grafico_dias_homologos(
        nombres_d,
        dias_m3_con,
        dias_m3_sin,
        f"Días homólogos: {lab_con} vs {lab_sin}",
        p_dias,
    )

    print("=" * 64)
    print(f"{NODO_ICCO} {NOMBRE_ICCO}")
    icco_dir = out_root / "Auditoria_Colegio_ICCO_Renca_000017-08"
    icco_dir.mkdir(parents=True, exist_ok=True)
    icco_info = _preparar_icco(
        ahora.date(),
        hora_corte,
        png_dir,
        icco_dir / "ICCO_domingo16_sinWES_vs_cada_dia.xlsx",
    )
    print(
        f"  Dom 16 sin WES {icco_info['m3_16']:.1f} m³ | Dom 09 con WES {icco_info['m3_9']:.1f} m³ | "
        f"{icco_info['pct_dom']:.1f}% ({icco_info['ahorro_dom']:.1f} m³)"
    )

    docx_path = out_root / f"Informe_Auditoria_Renca_semana_pasada_vs_esta_{ts}.docx"
    _word_conjunto(
        docx_path,
        hora_corte,
        ahora,
        dias_con,
        dias_sin,
        {"todos": p_all, "dias4": p_dias, "icco": icco_info["png_barras"]},
        m3_con,
        m3_sin,
        dias_m3_con,
        dias_m3_sin,
        dirs,
        icco_info,
    )
    pdf_path = _convertir_pdf(docx_path)

    tot_c = sum(m3_con.values())
    tot_s = sum(m3_sin.values())
    ahorro, pct = _rendimiento(tot_s, tot_c)
    print("=" * 64)
    print(f"TOTAL Con WES {tot_c:.1f} m³ | Sin WES {tot_s:.1f} m³ | {pct:.1f}% ({ahorro:.1f} m³)")
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
