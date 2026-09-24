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
import sys
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
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


def _totales_colegio(sitios: List[Sitio], meses: List[Tuple[int, int]], data):
    filas = []
    for s in sorted(sitios, key=lambda x: x.node_id):
        sc = sw = 0.0
        for k in meses:
            cel = data[s.node_id][k]
            sc += cel["cuenta"]
            sw += cel["wes"]
        filas.append((s, sc, sw, sc - sw))
    filas.sort(key=lambda t: abs(t[3]), reverse=True)
    return filas


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

    # Hoja principal: un total por colegio (lo que importa)
    wp = wb.active
    wp.title = "Total_por_colegio"
    headers_p = [
        "Colegio",
        "Nodo",
        "m³ cuenta (total)",
        "m³ WES (total)",
        "Dif m³ (cuenta − WES)",
        "% vs cuenta",
    ]
    wp.append(headers_p)
    for col in range(1, 7):
        c = wp.cell(1, col)
        c.fill = azul
        c.font = blanco
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    tot_c = tot_w = 0.0
    for s, sc, sw, dif in _totales_colegio(sitios, meses, data):
        pct = (100.0 * dif / sc) if sc else None
        wp.append([s.node_name, s.node_id, sc, round(sw, 1), round(dif, 1), None if pct is None else round(pct, 1)])
        tot_c += sc
        tot_w += sw
    wp.append(
        [
            "TOTAL",
            COMPANY_ID,
            tot_c,
            round(tot_w, 1),
            round(tot_c - tot_w, 1),
            round(100.0 * (tot_c - tot_w) / tot_c, 1) if tot_c else None,
        ]
    )
    for row in wp.iter_rows(min_row=2, max_row=wp.max_row):
        row[2].number_format = "#,##0"
        row[3].number_format = "#,##0.0"
        row[4].number_format = "#,##0.0"
        row[5].number_format = "0.0"
        for cell in row:
            cell.border = thin
    for cell in wp[wp.max_row]:
        cell.fill = tot_fill
        cell.font = Font(bold=True)
    wp.freeze_panes = "A2"
    for i, w in enumerate([28, 12, 18, 16, 22, 12], start=1):
        wp.column_dimensions[get_column_letter(i)].width = w

    ws = wb.create_sheet("Mensual_cuenta_vs_WES")

    # Colegio | Nodo | TOTAL cta | WES | % error | luego meses (cta | WES | % error)
    n_meses = len(meses)
    cols_por_mes = 3
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
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
        ws.cell(2, col, "m³ cuenta")
        ws.cell(2, col + 1, "m³ WES")
        ws.cell(2, col + 2, "% error")
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
    tot_mes = {k: {"cuenta": 0.0, "wes": 0.0} for k in meses}
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
            c2 = ws.cell(row_i, col + 1, wes if cel["n"] else None)
            c1.number_format = "#,##0"
            c2.number_format = "#,##0.0"
            if cel["n"]:
                _write_pct(ws.cell(row_i, col + 2), _pct_error_lectura(cta, wes))
            if cel["n_est"] and cel["n"]:
                c1.fill = PatternFill("solid", fgColor="FFF2CC")
            sc += cta
            sw += wes
            tot_mes[k]["cuenta"] += cta
            tot_mes[k]["wes"] += wes
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
        ws.cell(row_i, col + 1, tw).number_format = "#,##0.0"
        _write_pct(ws.cell(row_i, col + 2), _pct_error_lectura(tc, tw))
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
    wn["A1"] = "Qué número importa"
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
        "Celda amarilla en «m³ cuenta» mensual: esa boleta es promedio/estimado. "
        "Igual entra al total del colegio porque es lo que marcó la cuenta."
    )
    wn["A6"] = (
        "% error de lectura = (m³ WES − m³ cuenta) / m³ cuenta × 100. "
        "Positivo = WES marca más que la turbina de la cuenta. "
        "Naranja si |error| ≥ 15%. Vacío si la cuenta del mes es 0."
    )
    wn["A7"] = f"Generado: {generado.strftime('%d-%m-%Y %H:%M')}"
    wn.column_dimensions["A"].width = 120
    wb.save(out_xlsx)


def _write_word(
    sitios: List[Sitio],
    meses: List[Tuple[int, int]],
    data,
    out_docx: Path,
    generado: datetime,
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

    title = doc.add_heading("Total por colegio — cuenta vs WES", level=1)
    if title.runs:
        title.runs[0].font.color.rgb = HEADING_RGB
    p = doc.add_paragraph()
    p.add_run("CORMUP / Peñalolén. ").bold = True
    p.add_run(
        "El número que importa es el total de cada colegio (ene-2026 en adelante). "
        "Los meses van después, como desglose. "
        f"Generado {generado.strftime('%d-%m-%Y %H:%M')}."
    )

    add_formatted_heading(doc, "1. Total de cada colegio", level=1)
    headers_c = ["Colegio", "Nodo", "m³ cuenta", "m³ WES", "Dif m³", "% vs cuenta"]
    rows_c: List[List[str]] = []
    hi: List[int] = []
    gt_c = gt_w = 0.0
    for i, (s, sc, sw, dif) in enumerate(_totales_colegio(sitios, meses, data), start=1):
        pct = (100.0 * dif / sc) if sc else 0.0
        rows_c.append(
            [
                s.node_name,
                s.node_id,
                format_number_chilean(sc, 0),
                format_number_chilean(sw, 0),
                format_number_chilean(dif, 0),
                format_number_chilean(pct, 1),
            ]
        )
        if abs(pct) >= 15:
            hi.append(i)
        gt_c += sc
        gt_w += sw
    rows_c.append(
        [
            "TOTAL",
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

    add_formatted_heading(doc, "2. Desglose mensual (detalle)", level=1)
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
    _write_excel(sitios, meses, data, out_xlsx, generado)
    _write_word(sitios, meses, data, out_docx, generado)
    out_pdf = convertir_a_pdf(out_docx)
    print(f"[OK] Excel: {out_xlsx}", flush=True)
    print(f"[OK] Word:  {out_docx}", flush=True)
    if out_pdf:
        print(f"[OK] PDF:   {out_pdf}", flush=True)
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
