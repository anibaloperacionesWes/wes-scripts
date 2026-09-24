"""
Total por colegio: m³ cuenta (Aguas Andinas) vs m³ WES.

El número que importa es el total de cada colegio (ene-2026 en adelante).
Los meses son el desglose: fila = establecimiento, columnas = mes de emisión
con m³ cuenta, m³ WES (medido + proyección de huecos) y % error de lectura.

Uso:
  python generar_totalizados_mensuales_cormup_penalolen.py --skip-download
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from generar_comparativo_facturaciones_cormup_penalolen import (
    COMPANY_ID,
    HEADING_RGB,
    MESES_CORTOS,
    OUT_DIR,
    PDF_CACHE,
    Sitio,
    _cargar_sitios,
    _set_cell,
    _tbl_full_width,
    convertir_a_pdf,
    cruzar_wes,
    descargar_facturaciones,
)
from generar_reporte_word import (
    add_formatted_heading,
    add_logo_to_header,
    estilizar_tabla_wes,
    format_number_chilean,
)

MES_DESDE = date(2026, 1, 1)
UMBRAL_FUERA_PCT = 15.0


def _mes_key(d: date) -> Tuple[int, int]:
    return (d.year, d.month)


def _mes_label(y: int, m: int) -> str:
    return f"{MESES_CORTOS[m]}-{y}"


def _meses_rango(sitios: List[Sitio]) -> List[Tuple[int, int]]:
    keys = set()
    for s in sitios:
        for f in s.filas:
            em = f.emision.date()
            if em >= MES_DESDE:
                keys.add(_mes_key(em))
    if not keys:
        return []
    y0, m0 = min(keys)
    y1, m1 = max(keys)
    out: List[Tuple[int, int]] = []
    y, m = y0, m0
    while (y, m) <= (y1, m1):
        out.append((y, m))
        m += 1
        if m == 13:
            y += 1
            m = 1
    return out


def _matriz(sitios: List[Sitio], meses: List[Tuple[int, int]]):
    """sitio.node_id -> mes -> {cuenta, wes, n, n_est}."""
    data: Dict[str, Dict[Tuple[int, int], dict]] = {}
    for s in sitios:
        acc: Dict[Tuple[int, int], dict] = {
            k: {"cuenta": 0.0, "wes": 0.0, "n": 0, "n_est": 0} for k in meses
        }
        for f in s.filas:
            em = f.emision.date()
            if em < MES_DESDE:
                continue
            k = _mes_key(em)
            if k not in acc:
                continue
            acc[k]["cuenta"] += float(f.m3_cuenta)
            acc[k]["wes"] += float(f.m3_wes)
            acc[k]["n"] += 1
            if f.estimado:
                acc[k]["n_est"] += 1
        data[s.node_id] = acc
    return data


def _pct_error_lectura(cuenta: float, wes: float) -> Optional[float]:
    """Error de lectura WES vs cuenta: (WES − cuenta) / cuenta × 100."""
    if not cuenta:
        return None
    return 100.0 * (wes - cuenta) / cuenta


def _mes_fuera(cel) -> Optional[bool]:
    """True = no cuadra, False = cuadra, None = mes no evaluado."""
    if not cel["n"] or cel["cuenta"] <= 0:
        return None
    pct = _pct_error_lectura(cel["cuenta"], cel["wes"])
    if pct is None:
        return None
    return abs(pct) >= UMBRAL_FUERA_PCT


def _fuera_colegio(
    s: Sitio,
    meses: List[Tuple[int, int]],
    data,
    *,
    excluir_promedio: bool = False,
):
    n_eval = n_fuera = 0
    labels: List[str] = []
    for k in meses:
        cel = data[s.node_id][k]
        if excluir_promedio and cel["n_est"]:
            continue
        st = _mes_fuera(cel)
        if st is None:
            continue
        n_eval += 1
        if st:
            n_fuera += 1
            labels.append(MESES_CORTOS[k[1]])
    return n_fuera, n_eval, labels


def _promedio_colegio(s: Sitio, meses: List[Tuple[int, int]], data):
    n_prom = 0
    labels: List[str] = []
    for k in meses:
        cel = data[s.node_id][k]
        if cel["n"] and cel["n_est"]:
            n_prom += 1
            labels.append(MESES_CORTOS[k[1]])
    return n_prom, labels


def _totales_colegio(sitios: List[Sitio], meses: List[Tuple[int, int]], data):
    filas = []
    for s in sorted(sitios, key=lambda x: x.node_id):
        sc = sw = 0.0
        for k in meses:
            cel = data[s.node_id][k]
            sc += cel["cuenta"]
            sw += cel["wes"]
        n_fuera, n_eval, labels = _fuera_colegio(s, meses, data)
        n_fuera_sp, n_eval_sp, labels_sp = _fuera_colegio(
            s, meses, data, excluir_promedio=True
        )
        n_prom, labels_prom = _promedio_colegio(s, meses, data)
        filas.append(
            (
                s,
                sc,
                sw,
                sc - sw,
                n_fuera,
                n_eval,
                labels,
                n_prom,
                labels_prom,
                n_fuera_sp,
                n_eval_sp,
                labels_sp,
            )
        )
    filas.sort(key=lambda t: (t[9], abs(t[3])), reverse=True)
    return filas


def _escala_desvio(filas):
    """Más desviado → menos, por |% error| = |(WES − cuenta) / cuenta|."""
    scored = []
    for t in filas:
        pct = _pct_error_lectura(t[1], t[2])
        scored.append((abs(pct) if pct is not None else -1.0, pct, t))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored


def _barra_escala(abs_pct: float, max_abs: float, width: int = 18) -> str:
    if max_abs <= 0 or abs_pct < 0:
        return ""
    n = int(round(width * min(abs_pct, max_abs) / max_abs))
    return "█" * n + "░" * (width - n)


def _fill_desvio(abs_pct: float) -> PatternFill:
    if abs_pct >= 100:
        return PatternFill("solid", fgColor="C00000")
    if abs_pct >= 50:
        return PatternFill("solid", fgColor="E67E22")
    if abs_pct >= 15:
        return PatternFill("solid", fgColor="F4B183")
    return PatternFill("solid", fgColor="C6EFCE")


def _sentido(pct: Optional[float]) -> str:
    if pct is None:
        return "—"
    if abs(pct) < UMBRAL_FUERA_PCT:
        return "Cercano"
    return "WES marca más" if pct > 0 else "WES marca menos"


def _grafico_escala(scored, out_png: Path) -> Path:
    names = [t[2][0].node_name for t in scored]
    vals = [max(t[0], 0.0) for t in scored]
    colors = []
    for a, _pct, _t in scored:
        if a >= 100:
            colors.append("#C00000")
        elif a >= 50:
            colors.append("#E67E22")
        elif a >= 15:
            colors.append("#F4B183")
        else:
            colors.append("#548235")
    fig, ax = plt.subplots(figsize=(11.2, 6.4))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    ax.barh(names, vals, color=colors, height=0.72)
    ax.invert_yaxis()
    ax.set_xlabel("|% desvío|   (WES − cuenta) / cuenta")
    ax.set_title("Escala de desvío porcentual — del más desviado al menos")
    ax.axvline(UMBRAL_FUERA_PCT, color="#888888", linestyle="--", linewidth=0.8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.xaxis.grid(True, linestyle="--", linewidth=0.6, color="#cbd5e1")
    ax.set_axisbelow(True)
    for i, (a, pct, _t) in enumerate(scored):
        if pct is None:
            continue
        ax.text(a + max(vals) * 0.012, i, f"{pct:+.1f}%", va="center", fontsize=8)
    fig.tight_layout()
    fig.savefig(out_png, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out_png


def _write_excel(
    sitios: List[Sitio],
    meses: List[Tuple[int, int]],
    data,
    out_xlsx: Path,
    generado: datetime,
) -> None:
    wb = Workbook()
    azul = PatternFill("solid", fgColor="003366")
    azul2 = PatternFill("solid", fgColor="1F4E79")
    blanco = Font(color="FFFFFF", bold=True, size=9)
    tot_fill = PatternFill("solid", fgColor="D9E1F2")
    resalte = PatternFill("solid", fgColor="FFF2CC")
    err_fill = PatternFill("solid", fgColor="F8CBAD")
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )

    filas_tot = _totales_colegio(sitios, meses, data)
    scored = _escala_desvio(filas_tot)
    max_abs = max((a for a, _p, _t in scored), default=0.0)

    we = wb.active
    we.title = "Escala_desvio"
    headers_e = [
        "#",
        "Colegio",
        "|% desvío|",
        "% desvío",
        "Escala",
        "Sentido",
        "m³ cuenta",
        "m³ WES",
        "Dif m³",
        "Fuera s/prom.",
    ]
    we.append(headers_e)
    for col in range(1, 11):
        c = we.cell(1, col)
        c.fill = azul
        c.font = blanco
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    blanco_txt = Font(color="FFFFFF", bold=True, size=9)
    for i, (abs_pct, pct, t) in enumerate(scored, start=1):
        s, sc, sw, dif = t[0], t[1], t[2], t[3]
        n_fuera_sp = t[9]
        we.append(
            [
                i,
                s.node_name,
                None if abs_pct < 0 else round(abs_pct, 1),
                None if pct is None else round(pct, 1),
                _barra_escala(abs_pct, max_abs) if abs_pct >= 0 else "",
                _sentido(pct),
                sc,
                round(sw, 1),
                round(dif, 1),
                n_fuera_sp,
            ]
        )
        fill = _fill_desvio(abs_pct if abs_pct >= 0 else 0)
        we.cell(we.max_row, 1).fill = fill
        we.cell(we.max_row, 3).fill = fill
        we.cell(we.max_row, 4).fill = fill
        if abs_pct >= 50:
            we.cell(we.max_row, 1).font = blanco_txt
            we.cell(we.max_row, 3).font = blanco_txt
            we.cell(we.max_row, 4).font = blanco_txt
    for row in we.iter_rows(min_row=2, max_row=we.max_row):
        row[2].number_format = "0.0"
        row[3].number_format = "+0.0;-0.0;0.0"
        row[6].number_format = "#,##0"
        row[7].number_format = "#,##0.0"
        row[8].number_format = "#,##0.0"
        for cell in row:
            cell.border = thin
            cell.alignment = Alignment(horizontal="center")
        row[1].alignment = Alignment(horizontal="left")
        row[4].alignment = Alignment(horizontal="left")
    if we.max_row >= 2:
        we.conditional_formatting.add(
            f"C2:C{we.max_row}",
            DataBarRule(
                start_type="num",
                start_value=0,
                end_type="num",
                end_value=max(max_abs, 1),
                color="C00000",
                showValue=True,
            ),
        )
    we.freeze_panes = "A2"
    we.row_dimensions[1].height = 22
    for i, w in enumerate([5, 28, 12, 12, 22, 16, 14, 14, 14, 14], start=1):
        we.column_dimensions[get_column_letter(i)].width = w

    # Listado por colegio + cuántos meses no cuadran
    wp = wb.create_sheet("Total_por_colegio")
    headers_p = [
        "Colegio",
        "Meses fuera (sin promedio)",
        "Meses fuera (todos)",
        "Meses evaluados",
        "Meses promedio",
        "Nodo",
        "m³ cuenta (total)",
        "m³ WES (total)",
        "Dif m³ (cuenta − WES)",
        "% vs cuenta",
        "Meses fuera sin promedio",
        "Meses que no cuadran (todos)",
        "Meses con cuenta promedio",
    ]
    wp.append(headers_p)
    for col in range(1, 14):
        c = wp.cell(1, col)
        c.fill = azul
        c.font = blanco
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    tot_c = tot_w = 0.0
    tot_fuera = tot_eval = tot_prom = tot_fuera_sp = 0
    for (
        s,
        sc,
        sw,
        dif,
        n_fuera,
        n_eval,
        labels,
        n_prom,
        labels_prom,
        n_fuera_sp,
        _n_eval_sp,
        labels_sp,
    ) in _totales_colegio(sitios, meses, data):
        pct = (100.0 * dif / sc) if sc else None
        wp.append(
            [
                s.node_name,
                n_fuera_sp,
                n_fuera,
                n_eval,
                n_prom,
                s.node_id,
                sc,
                round(sw, 1),
                round(dif, 1),
                None if pct is None else round(pct, 1),
                ", ".join(labels_sp) if labels_sp else "—",
                ", ".join(labels) if labels else "—",
                ", ".join(labels_prom) if labels_prom else "—",
            ]
        )
        tot_c += sc
        tot_w += sw
        tot_fuera += n_fuera
        tot_eval += n_eval
        tot_prom += n_prom
        tot_fuera_sp += n_fuera_sp
        if n_fuera_sp:
            wp.cell(wp.max_row, 2).fill = err_fill
        if n_prom:
            wp.cell(wp.max_row, 5).fill = resalte
    wp.append(
        [
            "TOTAL",
            tot_fuera_sp,
            tot_fuera,
            tot_eval,
            tot_prom,
            COMPANY_ID,
            tot_c,
            round(tot_w, 1),
            round(tot_c - tot_w, 1),
            round(100.0 * (tot_c - tot_w) / tot_c, 1) if tot_c else None,
            "",
            "",
            "",
        ]
    )
    for row in wp.iter_rows(min_row=2, max_row=wp.max_row):
        row[6].number_format = "#,##0"
        row[7].number_format = "#,##0.0"
        row[8].number_format = "#,##0.0"
        row[9].number_format = "0.0"
        for cell in row:
            cell.border = thin
    for cell in wp[wp.max_row]:
        cell.fill = tot_fill
        cell.font = Font(bold=True)
    wp.freeze_panes = "A2"
    for i, w in enumerate([28, 22, 16, 16, 15, 12, 18, 16, 22, 12, 36, 36, 36], start=1):
        wp.column_dimensions[get_column_letter(i)].width = w

    ws = wb.create_sheet("Mensual_cuenta_vs_WES")

    # Colegio | Nodo | TOTAL cta | WES | % error | luego meses (cta | promedio | WES | % error)
    n_meses = len(meses)
    cols_por_mes = 4
    first_mes_col = 6
    ws.cell(1, 1, "Colegio")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    ws.cell(1, 2, "Nodo")
    ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)
    ws.cell(1, 3, "TOTAL colegio")
    ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=5)
    ws.cell(2, 3, "m³ cuenta")
    ws.cell(2, 4, "m³ WES")
    ws.cell(2, 5, "% error")
    col = first_mes_col
    for y, m in meses:
        ws.cell(1, col, _mes_label(y, m).upper())
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 3)
        ws.cell(2, col, "m³ cuenta")
        ws.cell(2, col + 1, "Promedio")
        ws.cell(2, col + 2, "m³ WES")
        ws.cell(2, col + 3, "% error")
        col += cols_por_mes

    last_col = 5 + n_meses * cols_por_mes
    for r in (1, 2):
        for c in range(1, last_col + 1):
            cell = ws.cell(r, c)
            cell.fill = azul if r == 1 else azul2
            cell.font = blanco
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 18

    def _write_pct(cell, pct: Optional[float]) -> None:
        cell.value = None if pct is None else round(pct, 1)
        cell.number_format = "0.0"
        if pct is not None and abs(pct) >= 15:
            cell.fill = err_fill

    sitios_ord = sorted(sitios, key=lambda s: s.node_id)
    tot_mes = {k: {"cuenta": 0.0, "wes": 0.0, "n": 0, "n_est": 0} for k in meses}
    tot_c = tot_w = 0.0
    row_i = 3
    for s in sitios_ord:
        ws.cell(row_i, 1, s.node_name)
        ws.cell(row_i, 2, s.node_id)
        sc = sw = 0.0
        col = first_mes_col
        for k in meses:
            cel = data[s.node_id][k]
            cta, wes = cel["cuenta"], cel["wes"]
            c1 = ws.cell(row_i, col, cta if cel["n"] else None)
            c_prom = ws.cell(row_i, col + 1)
            c2 = ws.cell(row_i, col + 2, wes if cel["n"] else None)
            c1.number_format = "#,##0"
            c2.number_format = "#,##0.0"
            if cel["n"]:
                es_prom = bool(cel["n_est"])
                c_prom.value = "Sí" if es_prom else "No"
                if es_prom:
                    c1.fill = resalte
                    c_prom.fill = resalte
                _write_pct(ws.cell(row_i, col + 3), _pct_error_lectura(cta, wes))
            sc += cta
            sw += wes
            tot_mes[k]["cuenta"] += cta
            tot_mes[k]["wes"] += wes
            tot_mes[k]["n_est"] = tot_mes[k].get("n_est", 0) + (1 if cel["n_est"] else 0)
            tot_mes[k]["n"] = tot_mes[k].get("n", 0) + (1 if cel["n"] else 0)
            col += cols_por_mes
        ws.cell(row_i, 3, sc).number_format = "#,##0"
        ws.cell(row_i, 4, sw).number_format = "#,##0.0"
        ws.cell(row_i, 3).fill = resalte
        ws.cell(row_i, 4).fill = resalte
        ws.cell(row_i, 3).font = Font(bold=True)
        ws.cell(row_i, 4).font = Font(bold=True)
        _write_pct(ws.cell(row_i, 5), _pct_error_lectura(sc, sw))
        ws.cell(row_i, 5).font = Font(bold=True)
        tot_c += sc
        tot_w += sw
        for c in range(1, last_col + 1):
            ws.cell(row_i, c).border = thin
            ws.cell(row_i, c).alignment = Alignment(horizontal="center")
        ws.cell(row_i, 1).alignment = Alignment(horizontal="left")
        row_i += 1

    ws.cell(row_i, 1, "TOTAL colegios")
    ws.cell(row_i, 2, COMPANY_ID)
    ws.cell(row_i, 3, tot_c).number_format = "#,##0"
    ws.cell(row_i, 4, tot_w).number_format = "#,##0.0"
    _write_pct(ws.cell(row_i, 5), _pct_error_lectura(tot_c, tot_w))
    col = first_mes_col
    for k in meses:
        tc, tw = tot_mes[k]["cuenta"], tot_mes[k]["wes"]
        ws.cell(row_i, col, tc).number_format = "#,##0"
        n_est_m = tot_mes[k]["n_est"]
        ws.cell(row_i, col + 1, n_est_m if tot_mes[k]["n"] else None)
        ws.cell(row_i, col + 2, tw).number_format = "#,##0.0"
        _write_pct(ws.cell(row_i, col + 3), _pct_error_lectura(tc, tw))
        col += cols_por_mes
    for c in range(1, last_col + 1):
        cell = ws.cell(row_i, c)
        cell.fill = tot_fill
        cell.font = Font(bold=True)
        cell.border = thin

    ws.freeze_panes = "F3"
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 13
    ws.column_dimensions["E"].width = 10
    for c in range(first_mes_col, last_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 10

    wn = wb.create_sheet("Notas")
    wn["A1"] = (
        "Hoja Escala_desvio: colegios del MÁS desviado al MENOS, "
        "según |% desvío| = |(m³ WES − m³ cuenta) / m³ cuenta|. "
        "Positivo = WES marca más que la turbina. #1 = peor."
    )
    wn["A2"] = (
        "El total que importa es el de CADA COLEGIO (hoja Total_por_colegio): "
        "suma de todas las boletas ene-2026 en adelante, m³ cuenta vs m³ WES. "
        "Los meses son el desglose, no el consolidado principal."
    )
    wn["A3"] = (
        "Cada boleta se asigna al mes de FECHA DE EMISIÓN (ene-2026 en adelante). "
        "Si un colegio tiene más de una boleta en el mismo mes, se suman."
    )
    wn["A4"] = (
        "m³ cuenta = consumo total de la boleta Aguas Andinas. "
        "m³ WES = registro API del mismo período de lecturas + proyección de días sin dato."
    )
    wn["A5"] = (
        "Columna «Promedio» / «Meses promedio»: la boleta Aguas Andinas de ese mes "
        "es consumo promedio, estimado SISS, medidor detenido o casa cerrada. "
        "Celda amarilla. Igual entra al total del colegio porque es lo que marcó la cuenta, "
        "pero no es lectura real de turbina."
    )
    wn["A6"] = (
        "% error de lectura = (m³ WES − m³ cuenta) / m³ cuenta × 100. "
        "Positivo = WES marca más que la turbina de la cuenta. "
        f"Naranja si |error| ≥ {UMBRAL_FUERA_PCT:.0f}%. Vacío si la cuenta del mes es 0."
    )
    wn["A7"] = (
        "Meses fuera (todos) = cuántos meses evaluados no cuadran "
        f"(|% error| ≥ {UMBRAL_FUERA_PCT:.0f}%), incluyendo boletas a promedio. "
        "Mes evaluado = hay boleta ese mes con m³ cuenta > 0."
    )
    wn["A8"] = (
        "Meses fuera (sin promedio) = los mismos meses fuera, "
        "pero sin contar boletas a promedio/estimado. "
        "Es el número que sirve para alinear ultrasónico vs turbina."
    )
    wn["A9"] = f"Generado: {generado.strftime('%d-%m-%Y %H:%M')}"
    wn.column_dimensions["A"].width = 120
    wb.save(out_xlsx)


def _write_word(
    sitios: List[Sitio],
    meses: List[Tuple[int, int]],
    data,
    out_docx: Path,
    generado: datetime,
    out_png: Optional[Path] = None,
) -> None:
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.2)
    add_logo_to_header(doc)
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    title = doc.add_heading("Escala de desvío — cuenta vs WES", level=1)
    if title.runs:
        title.runs[0].font.color.rgb = HEADING_RGB
    p = doc.add_paragraph()
    p.add_run("CORMUP / Peñalolén. ").bold = True
    p.add_run(
        "Escala del más desviado al menos, por |% desvío| "
        "= |(WES − cuenta) / cuenta|. "
        f"Generado {generado.strftime('%d-%m-%Y %H:%M')}."
    )

    add_formatted_heading(doc, "1. Escala de desvío porcentual", level=1)
    scored = _escala_desvio(_totales_colegio(sitios, meses, data))
    headers_e = ["#", "Colegio", "|% desvío|", "% desvío", "Sentido", "m³ cuenta", "m³ WES"]
    rows_e: List[List[str]] = []
    hi_e: List[int] = []
    for i, (abs_pct, pct, t) in enumerate(scored, start=1):
        s, sc, sw = t[0], t[1], t[2]
        rows_e.append(
            [
                str(i),
                s.node_name,
                format_number_chilean(abs_pct, 1) if abs_pct >= 0 else "—",
                format_number_chilean(pct, 1) if pct is not None else "—",
                _sentido(pct),
                format_number_chilean(sc, 0),
                format_number_chilean(sw, 0),
            ]
        )
        if abs_pct >= UMBRAL_FUERA_PCT:
            hi_e.append(i)
    table_e = doc.add_table(rows=1 + len(rows_e), cols=len(headers_e))
    table_e.alignment = WD_TABLE_ALIGNMENT.CENTER
    _tbl_full_width(table_e)
    for j, h in enumerate(headers_e):
        _set_cell(table_e.rows[0].cells[j], h, bold=True, size=9)
    for i, row in enumerate(rows_e, start=1):
        for j, val in enumerate(row):
            _set_cell(table_e.rows[i].cells[j], val, bold=(j == 0), size=9)
    estilizar_tabla_wes(table_e, highlight_rows=hi_e, has_total_row=False)
    if out_png and out_png.is_file():
        doc.add_paragraph()
        doc.add_picture(str(out_png), width=Cm(22))

    add_formatted_heading(doc, "2. Total de cada colegio", level=1)
    headers_c = [
        "Colegio",
        "Fuera s/prom.",
        "Fuera todos",
        "Meses eval.",
        "Meses prom.",
        "Nodo",
        "m³ cuenta",
        "m³ WES",
        "Dif m³",
        "% vs cuenta",
    ]
    rows_c: List[List[str]] = []
    hi: List[int] = []
    gt_c = gt_w = 0.0
    gt_fuera = gt_eval = gt_prom = gt_fuera_sp = 0
    for i, (
        s,
        sc,
        sw,
        dif,
        n_fuera,
        n_eval,
        _labels,
        n_prom,
        _lp,
        n_fuera_sp,
        _nes,
        _lsp,
    ) in enumerate(_totales_colegio(sitios, meses, data), start=1):
        pct = (100.0 * dif / sc) if sc else 0.0
        rows_c.append(
            [
                s.node_name,
                str(n_fuera_sp),
                str(n_fuera),
                str(n_eval),
                str(n_prom),
                s.node_id,
                format_number_chilean(sc, 0),
                format_number_chilean(sw, 0),
                format_number_chilean(dif, 0),
                format_number_chilean(pct, 1),
            ]
        )
        if n_fuera_sp:
            hi.append(i)
        gt_c += sc
        gt_w += sw
        gt_fuera += n_fuera
        gt_eval += n_eval
        gt_prom += n_prom
        gt_fuera_sp += n_fuera_sp
    rows_c.append(
        [
            "TOTAL",
            str(gt_fuera_sp),
            str(gt_fuera),
            str(gt_eval),
            str(gt_prom),
            COMPANY_ID,
            format_number_chilean(gt_c, 0),
            format_number_chilean(gt_w, 0),
            format_number_chilean(gt_c - gt_w, 0),
            format_number_chilean(100.0 * (gt_c - gt_w) / gt_c, 1) if gt_c else "0",
        ]
    )
    table_c = doc.add_table(rows=1 + len(rows_c), cols=len(headers_c))
    table_c.alignment = WD_TABLE_ALIGNMENT.CENTER
    _tbl_full_width(table_c)
    for j, h in enumerate(headers_c):
        _set_cell(table_c.rows[0].cells[j], h, bold=True, size=9)
    for i, row in enumerate(rows_c, start=1):
        is_tot = i == len(rows_c)
        for j, val in enumerate(row):
            _set_cell(table_c.rows[i].cells[j], val, bold=is_tot, size=9)
    estilizar_tabla_wes(table_c, highlight_rows=hi, has_total_row=True)

    add_formatted_heading(doc, "3. Desglose mensual (detalle)", level=1)
    headers = ["Colegio", "Total cta", "Total WES"]
    for y, m in meses:
        lab = f"{MESES_CORTOS[m]}-{str(y)[2:]}"
        headers.append(f"{lab} cta")
        headers.append(f"{lab} WES")

    sitios_ord = sorted(sitios, key=lambda s: s.node_id)
    rows: List[List[str]] = []
    tot_mes = {k: {"cuenta": 0.0, "wes": 0.0} for k in meses}
    for s in sitios_ord:
        sc = sw = 0.0
        meses_txt: List[str] = []
        for k in meses:
            cel = data[s.node_id][k]
            if cel["n"]:
                meses_txt.append(format_number_chilean(cel["cuenta"], 0))
                meses_txt.append(format_number_chilean(cel["wes"], 0))
                sc += cel["cuenta"]
                sw += cel["wes"]
                tot_mes[k]["cuenta"] += cel["cuenta"]
                tot_mes[k]["wes"] += cel["wes"]
            else:
                meses_txt.append("—")
                meses_txt.append("—")
        rows.append([s.node_name, format_number_chilean(sc, 0), format_number_chilean(sw, 0), *meses_txt])
    tot_line = [
        "TOTAL",
        format_number_chilean(gt_c, 0),
        format_number_chilean(gt_w, 0),
    ]
    for k in meses:
        tot_line.append(format_number_chilean(tot_mes[k]["cuenta"], 0))
        tot_line.append(format_number_chilean(tot_mes[k]["wes"], 0))
    rows.append(tot_line)

    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _tbl_full_width(table)
    for j, h in enumerate(headers):
        _set_cell(table.rows[0].cells[j], h, bold=True, size=7)
    for i, row in enumerate(rows, start=1):
        is_tot = i == len(rows)
        for j, val in enumerate(row):
            _set_cell(table.rows[i].cells[j], val, bold=is_tot or j in (1, 2), size=7)
    estilizar_tabla_wes(table, has_total_row=True)

    doc.add_paragraph(
        "Mes = fecha de emisión de la boleta. WES = consumo API en el período de lecturas "
        "de esa misma boleta (más proyección si hubo huecos). "
        "El % de error de lectura por mes (WES − cuenta) / cuenta está en la hoja "
        "Excel Mensual_cuenta_vs_WES."
    )
    out_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_docx)


HOJAS_PDF = ("Escala_desvio", "Total_por_colegio", "Mensual_cuenta_vs_WES")


def _font_grande(cell, size: int) -> None:
    f = cell.font
    cell.font = Font(name="Calibri", size=size, bold=bool(f.bold), color=f.color)


def _agrandar_hoja(ws, size: int = 12) -> None:
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            _font_grande(cell, size)
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 22 if r <= 2 else 20
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        cur = ws.column_dimensions[letter].width or 10
        ws.column_dimensions[letter].width = max(cur * 1.08, 11)


def _pagina_horizontal(ws, *, encajar: bool, encabezado_filas: str = "1:1") -> None:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.horizontalCentered = True
    ws.page_setup.verticalCentered = False
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.45
    ws.page_margins.bottom = 0.4
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2
    ws.oddHeader.left.text = "&A — CORMUP Peñalolén"
    ws.oddFooter.right.text = "Pág. &P / &N"
    ws.print_title_rows = encabezado_filas
    if encajar:
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    else:
        ws.page_setup.fitToPage = False
        ws.page_setup.scale = 120
        ws.sheet_properties.pageSetUpPr.fitToPage = False


def _cortes_mensual(ws) -> Tuple[int, List[Tuple[int, int]]]:
    """Columna donde empiezan los meses y rangos (col_ini, col_fin) de cada mes."""
    bloques = []
    for rng in ws.merged_cells.ranges:
        if rng.min_row == 1 and (rng.max_col - rng.min_col + 1) >= 4:
            bloques.append((rng.min_col, rng.max_col))
    meses = sorted(set(bloques))
    if not meses:
        return 6, []
    return meses[0][0], meses


def _partir_mensual(wb, ws):
    """Dos hojas horizontales (ene–may / jun–sep) para que el texto no se achique."""
    first_mes, meses = _cortes_mensual(ws)
    if len(meses) < 3:
        _agrandar_hoja(ws, 11)
        _pagina_horizontal(ws, encajar=True, encabezado_filas="1:2")
        ws.print_title_cols = "A:E"
        return [ws]
    mid = (len(meses) + 1) // 2
    cortes = [meses[:mid], meses[mid:]]
    titulos = []
    for cols in cortes:
        a, b = cols[0][0], cols[-1][1]
        lab_a = str(ws.cell(1, cols[0][0]).value or "ini")
        lab_b = str(ws.cell(1, cols[-1][0]).value or "fin")
        titulos.append((f"Mensual_{lab_a}_{lab_b}".replace(" ", "")[:31], a, b))

    copias = []
    for titulo, c0, c1 in titulos:
        copia = wb.copy_worksheet(ws)
        copia.title = titulo
        for col in range(first_mes, ws.max_column + 1):
            if col < c0 or col > c1:
                copia.column_dimensions[get_column_letter(col)].hidden = True
        _agrandar_hoja(copia, 12)
        _pagina_horizontal(copia, encajar=True, encabezado_filas="1:2")
        copia.print_title_cols = "A:E"
        copias.append(copia)
    del wb[ws.title]
    return copias


def pdf_tres_hojas(xlsx_path: Path) -> Optional[Path]:
    """PDF horizontal de Escala, Total por colegio y Mensual (esta última en 2 páginas)."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("[WARN] No hay LibreOffice/soffice; se omite PDF de 3 hojas.", flush=True)
        return None
    wb = load_workbook(xlsx_path)
    for name in list(wb.sheetnames):
        if name not in HOJAS_PDF:
            del wb[name]
    faltan = [n for n in HOJAS_PDF if n not in wb.sheetnames]
    if faltan:
        raise SystemExit(f"Faltan hojas para el PDF: {faltan}")

    we = wb["Escala_desvio"]
    _agrandar_hoja(we, 12)
    _pagina_horizontal(we, encajar=True)

    wp = wb["Total_por_colegio"]
    _agrandar_hoja(wp, 11)
    _pagina_horizontal(wp, encajar=True)

    _partir_mensual(wb, wb["Mensual_cuenta_vs_WES"])

    orden = [n for n in wb.sheetnames]
    wb._sheets = [wb[n] for n in orden]
    out_pdf = xlsx_path.with_name(xlsx_path.stem + "_3hojas.pdf")
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td) / "tres_hojas.xlsx"
        wb.save(tmp)
        subprocess.run(
            [
                soffice,
                "--headless",
                "--nologo",
                "--nofirststartwizard",
                "--convert-to",
                "pdf:calc_pdf_Export",
                "--outdir",
                td,
                str(tmp),
            ],
            check=True,
            capture_output=True,
            text=True,
        )
        produced = Path(td) / "tres_hojas.pdf"
        if not produced.is_file():
            print("[WARN] LibreOffice no generó el PDF de 3 hojas.", flush=True)
            return None
        shutil.copy2(produced, out_pdf)
    return out_pdf


def generar(skip_download: bool = False) -> Tuple[Path, Optional[Path], Path]:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    generado = datetime.now(timezone.utc).astimezone()
    ts = generado.strftime("%Y%m%d_%H%M")

    if skip_download and any(PDF_CACHE.glob("*/*.pdf")):
        locales = {p.name: p for p in PDF_CACHE.iterdir() if p.is_dir()}
        print(f"[INFO] Reuso cache {PDF_CACHE}", flush=True)
    else:
        locales = descargar_facturaciones(PDF_CACHE)

    sitios = _cargar_sitios(locales)
    cruzar_wes(sitios)
    meses = _meses_rango(sitios)
    if not meses:
        raise SystemExit("No hay boletas con emisión desde enero 2026.")
    data = _matriz(sitios, meses)

    stem = f"Totalizados_mensuales_cuenta_vs_WES_CORMUP_{ts}"
    out_xlsx = OUT_DIR / f"{stem}.xlsx"
    out_docx = OUT_DIR / f"{stem}.docx"
    out_png = OUT_DIR / f"{stem}_escala.png"
    _grafico_escala(_escala_desvio(_totales_colegio(sitios, meses, data)), out_png)
    _write_excel(sitios, meses, data, out_xlsx, generado)
    _write_word(sitios, meses, data, out_docx, generado, out_png=out_png)
    out_pdf = convertir_a_pdf(out_docx)
    out_pdf_hojas = pdf_tres_hojas(out_xlsx)
    print(f"[OK] Excel: {out_xlsx}", flush=True)
    print(f"[OK] Word:  {out_docx}", flush=True)
    if out_pdf:
        print(f"[OK] PDF:   {out_pdf}", flush=True)
    if out_pdf_hojas:
        print(f"[OK] PDF 3 hojas: {out_pdf_hojas}", flush=True)
    return out_docx, out_pdf, out_xlsx


def main() -> int:
    if sys.platform == "win32":
        for s in (sys.stdout, sys.stderr):
            try:
                s.reconfigure(encoding="utf-8", line_buffering=True)
            except Exception:
                pass
    ap = argparse.ArgumentParser(description="Totalizados mensuales cuenta vs WES CORMUP")
    ap.add_argument("--skip-download", action="store_true")
    args = ap.parse_args()
    generar(skip_download=args.skip_download)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
