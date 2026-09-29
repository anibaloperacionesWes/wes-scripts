"""
Juan Pablo II (000008-14): 1ª columna Hora, consumo por fecha a la derecha.

Fila 28 = Total (celeste). Fila 29 = Listado (verde). Rojo solo si Listado ≠ Total.

Uso:
  python generar_horario_jp2_hora_consumo.py
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from copy import copy

from wes_estilo_graficos_app import horas_api_chile
from wes_google_drive import credenciales_configuradas, subir_a_drive

NODE = "000008-14"
NOMBRE = "Juan Pablo II"
D0 = date(2026, 3, 1)
D1 = date(2026, 9, 28)
OUT_DIR = Path("reports/CORMUP/Facturaciones_vs_WES")
DRIVE_SUB = "CORMUP/Facturaciones_vs_WES"
# Mismo archivo de Drive en el que estamos trabajando.
DRIVE_NOMBRE = "Horario_JP2_hora_consumo_20260929_1649.xlsx"

# Listado diario pegado bajo Total, alineado por fecha (marzo–julio en 2026).
_LISTADO_MAR_JUL = """
2026-03-01 0.35
2026-03-02 3.81
2026-03-03 4.26
2026-03-04 4.11
2026-03-05 4.89
2026-03-06 4.61
2026-03-07 0.14
2026-03-08 0.04
2026-03-09 4.14
2026-03-10 4.44
2026-03-11 5.19
2026-03-12 1.04
2026-03-13 5.41
2026-03-14 0.01
2026-03-15 0.03
2026-03-16 4.01
2026-03-17 4.59
2026-03-18 4.21
2026-03-19 5.21
2026-03-20 0.45
2026-03-21 0.03
2026-03-22 4.4
2026-03-23 4.41
2026-03-24 1.21
2026-03-25 7.85
2026-03-26 4.83
2026-03-27 0.21
2026-03-28 0.01
2026-03-29 5.79
2026-03-30 4.16
2026-03-31 4.75
2026-04-01 5.06
2026-04-02 5.3
2026-04-03 5.53
2026-04-04 0.26
2026-04-05 0.03
2026-04-06 5.3
2026-04-07 5.45
2026-04-08 5.14
2026-04-09 4.77
2026-04-10 6.26
2026-04-11 0.26
2026-04-12 0.03
2026-04-13 5.96
2026-04-14 5.36
2026-04-15 5.28
2026-04-16 4.87
2026-04-17 4.38
2026-04-18 0.27
2026-04-19 0.02
2026-04-20 5.04
2026-04-21 5.08
2026-04-22 4.26
2026-04-23 5.77
2026-04-24 6.36
2026-04-25 0.47
2026-04-26 0.03
2026-04-27 6.04
2026-04-28 4.57
2026-04-29 4.38
2026-04-30 4.36
2026-05-01 4.87
2026-05-02 0.25
2026-05-03 0.03
2026-05-04 4.28
2026-05-05 4.81
2026-05-06 4.34
2026-05-07 4.91
2026-05-08 5.34
2026-05-09 0.26
2026-05-10 0.03
2026-05-11 4.88
2026-05-12 5.06
2026-05-13 5.16
2026-05-14 4.52
2026-05-15 4.44
2026-05-16 0.28
2026-05-17 0.03
2026-05-18 4.28
2026-05-19 4.24
2026-05-20 4.36
2026-05-21 4.18
2026-05-22 5.06
2026-05-23 0.27
2026-05-24 0.03
2026-05-25 4.26
2026-05-26 4.38
2026-05-27 4.42
2026-05-28 4.08
2026-05-29 4.14
2026-05-30 0.26
2026-05-31 0.03
2026-06-01 4.08
2026-06-02 4.06
2026-06-03 4.5
2026-06-04 4.08
2026-06-05 4.14
2026-06-06 0.26
2026-06-07 0.04
2026-06-08 3.56
2026-06-09 4.12
2026-06-10 4.14
2026-06-11 4.08
2026-06-12 3.13
2026-06-13 0.22
2026-06-14 0.02
2026-06-15 3.12
2026-06-16 4.08
2026-06-17 4.06
2026-06-18 4.04
2026-06-19 4.03
2026-06-20 0.22
2026-06-21 0.04
2026-06-22 4.06
2026-06-23 3.13
2026-06-24 3.08
2026-06-25 3.06
2026-06-26 4.03
2026-06-27 0.23
2026-06-28 0.03
2026-06-29 4.08
2026-06-30 4.04
2026-07-01 4.06
2026-07-02 4.12
2026-07-03 3.12
2026-07-04 0.22
2026-07-05 0.03
2026-07-06 4.02
2026-07-07 4.05
2026-07-08 4.04
2026-07-09 4.08
2026-07-10 4.05
2026-07-11 0.22
2026-07-12 0.04
2026-07-13 4.04
2026-07-14 4.06
2026-07-15 4.06
2026-07-16 4.18
2026-07-17 4.08
2026-07-18 0.23
2026-07-19 0.03
2026-07-20 4.02
2026-07-21 4.03
2026-07-22 4.12
2026-07-23 4.14
2026-07-24 4.08
2026-07-25 0.22
2026-07-26 0.03
2026-07-27 4.05
2026-07-28 4.7
"""


def _parse_listado_blob(blob: str) -> dict[date, float]:
    out: dict[date, float] = {}
    for line in blob.strip().splitlines():
        parts = line.split()
        if len(parts) != 2:
            continue
        y, m, d = (int(x) for x in parts[0].split("-"))
        out[date(2026, m, d)] = float(parts[1].replace(",", "."))
    return out


LISTADO = {
    **_parse_listado_blob(_LISTADO_MAR_JUL),
    date(2026, 7, 29): 4.17,
    date(2026, 7, 30): 5.16,
    date(2026, 7, 31): 4.87,
    date(2026, 8, 1): 5.2,
    date(2026, 8, 2): 0.02,
    date(2026, 8, 3): 7.29,
    date(2026, 8, 4): 6.22,
    date(2026, 8, 5): 5.49,
    date(2026, 8, 6): 4.81,
    date(2026, 8, 7): 5.02,
    date(2026, 8, 8): 0.41,
    date(2026, 8, 9): 0.03,
    date(2026, 8, 10): 4.94,
    date(2026, 8, 11): 4.14,
    date(2026, 8, 12): 3.3,
    date(2026, 8, 13): 3.15,
    date(2026, 8, 14): 4.01,
    date(2026, 8, 15): 0.57,
    date(2026, 8, 16): 0.04,
    date(2026, 8, 17): 2.98,
    date(2026, 8, 18): 3.02,
    date(2026, 8, 19): 3.67,
    date(2026, 8, 20): 5.14,
    date(2026, 8, 21): 3.43,
    date(2026, 8, 22): 0.4,
    date(2026, 8, 23): 0.03,
    date(2026, 8, 24): 4.62,
    date(2026, 8, 25): 4.2,
    date(2026, 8, 26): 5.26,
    date(2026, 8, 27): 3.7,
    date(2026, 8, 28): 3.99,
    date(2026, 8, 29): 5.17,
    date(2026, 8, 30): 0.4,
    date(2026, 8, 31): 3.43,
    date(2026, 9, 1): 3.26,
    date(2026, 9, 2): 5.68,
    date(2026, 9, 3): 4.99,
    date(2026, 9, 4): 4.69,
    date(2026, 9, 5): 4.85,
    date(2026, 9, 6): 0.36,
    date(2026, 9, 7): 5.42,
    date(2026, 9, 8): 7.21,
    date(2026, 9, 9): 5.02,
    date(2026, 9, 10): 5.43,
    date(2026, 9, 11): 4.01,
    date(2026, 9, 12): 0.0,
    date(2026, 9, 13): 0.0,
    date(2026, 9, 14): 0.0,
    date(2026, 9, 15): 0.0,
    date(2026, 9, 16): 0.0,
    date(2026, 9, 17): 0.0,
    date(2026, 9, 18): 0.0,
    date(2026, 9, 19): 0.0,
    date(2026, 9, 20): 0.0,
    date(2026, 9, 21): 10.62,
    date(2026, 9, 22): 11.28,
    date(2026, 9, 23): 11.36,
    date(2026, 9, 24): 11.02,
    date(2026, 9, 25): 7.58,
    date(2026, 9, 26): 0.0,
    date(2026, 9, 27): 0.0,
    date(2026, 9, 28): 12.44,
}

HDR = PatternFill("solid", fgColor="D9D9D9")
SUB = PatternFill("solid", fgColor="F2F2F2")
CELESTE = PatternFill("solid", fgColor="9DC3E6")
VERDE = PatternFill("solid", fgColor="C6EFCE")
ROJO = PatternFill("solid", fgColor="FFC7CE")
FONT_OK = Font(bold=True, size=9)
FONT_ROJO = Font(bold=True, size=9, color="9C0006")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


NUM_FMT = "0.00"
TOL_M3 = 0.01  # descuadre Listado vs Total WES (2 decimales)


def _num(v: float) -> float:
    """Número real con 2 decimales (punto) para que Excel/Sheets pueda sumar."""
    return round(float(v), 2)


def _no_cuadra(listado: float | None, wes_total: float) -> bool:
    if listado is None:
        return False
    return abs(_num(listado) - _num(wes_total)) > TOL_M3


def _dias_unicos(d0: date, d1: date) -> list[date]:
    dias: list[date] = []
    seen: set[date] = set()
    d = d0
    while d <= d1:
        if d not in seen:
            seen.add(d)
            dias.append(d)
        d += timedelta(days=1)
    return dias


def _cargar_horas_xlsx(path: Path) -> dict[date, dict[int, float]]:
    """Reusa horas ya exportadas para no volver a consultar días viejos."""
    out: dict[date, dict[int, float]] = {}
    if not path.is_file():
        return out
    wb = load_workbook(path, data_only=False)
    ws = wb.active
    for col in range(2, ws.max_column + 1):
        hdr = str(ws.cell(2, col).value or "")
        try:
            dia = datetime.strptime(hdr.replace("Fecha", "").strip(), "%d/%m/%Y").date()
        except ValueError:
            continue
        horas: dict[int, float] = {}
        for h in range(24):
            v = ws.cell(4 + h, col).value
            if isinstance(v, (int, float)):
                horas[h] = float(v)
            elif isinstance(v, str) and v and not v.startswith("="):
                horas[h] = float(v.replace(",", "."))
            else:
                horas[h] = 0.0
        out[dia] = horas
    return out


def _horas_rango(dias: list[date], cache: dict[date, dict[int, float]] | None = None) -> dict[date, dict[int, float]]:
    horas: dict[date, dict[int, float]] = dict(cache or {})
    for dia in dias:
        if dia in horas and len(horas[dia]) == 24:
            continue
        try:
            h = horas_api_chile(NODE, datetime.combine(dia, datetime.min.time()))
            horas[dia] = {i: float(h.get(i, 0.0)) for i in range(24)}
            print(f"  API {dia.isoformat()}", flush=True)
        except Exception as exc:
            print(f"  API FAIL {dia.isoformat()}: {exc}", flush=True)
            horas[dia] = {i: 0.0 for i in range(24)}
    return horas


def construir_horario(dias: list[date], horas: dict[date, dict[int, float]]) -> Workbook:

    wb = Workbook()
    ws = wb.active
    ws.title = "Horario"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + min(8, len(dias)))
    ws["A1"] = (
        f"{NOMBRE} ({NODE}) — 1ª columna Hora; consumo por fecha a la derecha "
        f"({D0.strftime('%d-%m')} a {D1.strftime('%d-%m')}). Día repetido no se suma. "
        "Fila Total = celeste. Fila Listado = verde. Rojo solo si Listado no coincide con Total."
    )
    ws["A1"].font = Font(bold=True, size=12, color="003366")
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 32

    ws.merge_cells(start_row=2, start_column=1, end_row=3, end_column=1)
    c = ws.cell(2, 1, "Hora")
    c.fill = HDR
    c.font = Font(bold=True, size=11)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = THIN
    ws.cell(3, 1).border = THIN
    ws.cell(3, 1).fill = HDR

    for i, dia in enumerate(dias):
        col = 2 + i
        top = ws.cell(2, col, f"Fecha  {dia.strftime('%d/%m/%Y')}")
        top.fill = HDR
        top.font = Font(bold=True, size=9)
        top.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        top.border = THIN
        sub = ws.cell(3, col, "Consumo")
        sub.fill = SUB
        sub.font = Font(bold=True, size=9)
        sub.alignment = CENTER
        sub.border = THIN
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions["A"].width = 12
    ws.row_dimensions[2].height = 22

    for h in range(24):
        r = 4 + h
        bg = PatternFill("solid", fgColor="FFFFFF" if h % 2 == 0 else "FCFCFC")
        a = ws.cell(r, 1, f"{h:02d}:00")
        a.alignment = CENTER
        a.border = THIN
        a.fill = bg
        a.font = Font(bold=True, size=9)
        for i, dia in enumerate(dias):
            b = ws.cell(r, 2 + i, _num(horas[dia][h]))
            b.number_format = NUM_FMT
            b.alignment = CENTER
            b.border = THIN
            b.fill = bg

    rt = 28
    t = ws.cell(rt, 1, "Total")
    t.font = Font(bold=True, size=9)
    t.alignment = CENTER
    t.border = THIN
    t.fill = CELESTE
    for i, dia in enumerate(dias):
        col = get_column_letter(2 + i)
        b = ws.cell(rt, 2 + i, f"=SUM({col}4:{col}27)")
        b.number_format = NUM_FMT
        b.font = FONT_OK
        b.alignment = CENTER
        b.border = THIN
        b.fill = CELESTE

    rl = 29
    lab = ws.cell(rl, 1, "Listado")
    lab.font = Font(bold=True, size=9)
    lab.alignment = CENTER
    lab.border = THIN
    lab.fill = VERDE
    n_rojo = 0
    for i, dia in enumerate(dias):
        val = LISTADO.get(dia)
        wes_tot = sum(horas[dia].values())
        descuadre = _no_cuadra(val, wes_tot)
        cell = ws.cell(rl, 2 + i, _num(val) if val is not None else None)
        if val is not None:
            cell.number_format = NUM_FMT
        cell.font = FONT_ROJO if descuadre else FONT_OK
        cell.alignment = CENTER
        cell.border = THIN
        cell.fill = ROJO if descuadre else VERDE
        if descuadre:
            n_rojo += 1
            print(
                f"  ROJO {dia.strftime('%d/%m')} listado={_num(val)} total={_num(wes_tot)} "
                f"delta={(_num(val) - _num(wes_tot)):+.2f}",
                flush=True,
            )

    ultima = get_column_letter(1 + len(dias))
    ws.conditional_formatting.add(
        f"B29:{ultima}29",
        FormulaRule(
            formula=["ABS(B29-B28)>0.01"],
            fill=ROJO,
            font=FONT_ROJO,
        ),
    )
    print(f"[INFO] {n_rojo} días no cuadran (Listado vs Total)", flush=True)

    ws.freeze_panes = "B4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.print_title_rows = "1:3"
    return wb


AZUL_HDR = PatternFill("solid", fgColor="1F4E79")
AZUL_APP = PatternFill("solid", fgColor="2E75B6")
ROJO_CTA = PatternFill("solid", fgColor="C00000")
VERDE_PROM = PatternFill("solid", fgColor="C6EFCE")
GRIS = PatternFill("solid", fgColor="D9E1F2")
FONT_BLANCO = Font(color="FFFFFF", bold=True, size=9)

# Períodos de boleta Juan Pablo II (Aguas Andinas, corte 12:00).
PERIODOS_FACTURA = [
    {
        "mes": "ene-2026",
        "d0": date(2025, 11, 29),
        "d1": date(2025, 12, 30),
        "ini": "29-11-2025 12:00",
        "fin": "30-12-2025 12:00",
        "dif": 90.0,
        "estimado": True,
        "hora_fin": 12,
    },
    {
        "mes": "feb-2026",
        "d0": date(2025, 12, 30),
        "d1": date(2026, 1, 29),
        "ini": "30-12-2025 12:00",
        "fin": "29-01-2026 12:00",
        "dif": 159.0,
        "estimado": True,
        "hora_fin": 12,
    },
    {
        "mes": "mar-2026",
        "d0": date(2026, 1, 29),
        "d1": date(2026, 2, 27),
        "ini": "29-01-2026 12:00",
        "fin": "27-02-2026 12:00",
        "dif": 159.0,
        "estimado": True,
        "hora_fin": 12,
    },
    {
        "mes": "abr-2026",
        "d0": date(2026, 2, 27),
        "d1": date(2026, 3, 30),
        "ini": "27-02-2026 12:00",
        "fin": "30-03-2026 12:00",
        "dif": 159.0,
        "estimado": True,
        "hora_fin": 12,
    },
    {
        "mes": "may-2026",
        "d0": date(2025, 10, 29),
        "d1": date(2026, 4, 29),
        "ini": "29-10-2025 12:00  (72.170 m³)",
        "fin": "29-04-2026 12:00  (72.937 m³)",
        "dif": 767.0,
        "estimado": False,
        "hora_fin": 12,
    },
    {
        "mes": "jun-2026",
        "d0": date(2026, 4, 29),
        "d1": date(2026, 5, 29),
        "ini": "29-04-2026 12:00  (72.937 m³)",
        "fin": "29-05-2026 12:00  (72.986 m³)",
        "dif": 49.0,
        "estimado": False,
        "hora_fin": 12,
    },
    {
        "mes": "jul-2026",
        "d0": date(2026, 5, 29),
        "d1": date(2026, 6, 27),
        "ini": "29-05-2026 12:00  (72.986 m³)",
        "fin": "27-06-2026 12:00",
        "dif": 125.0,
        "estimado": True,
        "hora_fin": 12,
    },
    {
        "mes": "ago-2026",
        "d0": date(2026, 5, 29),
        "d1": date(2026, 7, 29),
        "ini": "29-05-2026 12:00  (72.986 m³)",
        "fin": "29-07-2026 12:00  (73.206 m³)",
        "dif": 220.0,
        "estimado": False,
        "hora_fin": 12,
    },
    {
        "mes": "sep-2026",
        "d0": date(2026, 7, 29),
        "d1": date(2026, 8, 29),
        "ini": "29-07-2026 12:00  (73.206 m³)",
        "fin": "29-08-2026 12:00  (73.303 m³)",
        "dif": 97.0,
        "estimado": False,
        "hora_fin": 12,
    },
    {
        "mes": "terreno 25-sep",
        "d0": date(2026, 8, 29),
        "d1": date(2026, 9, 25),
        "ini": "29-08-2026 12:00  (73.303 m³)",
        "fin": "25-09-2026 11:00  (73.385 m³)",
        "dif": 82.0,
        "estimado": False,
        "hora_fin": 11,
    },
]


def _horas_dia(horas: dict[date, dict[int, float]], dia: date) -> dict[int, float]:
    return horas.get(dia) or {i: 0.0 for i in range(24)}


def _total_placa_periodo(
    horas: dict[date, dict[int, float]],
    d0: date,
    d1: date,
    hora_fin: int = 12,
) -> float:
    """Suma horaria WES con corte mediodía: inicio 12:00–23:59, final 00:00–hora_fin."""
    if d1 < d0:
        return 0.0
    h0 = _horas_dia(horas, d0)
    if d0 == d1:
        return sum(h0.get(i, 0.0) for i in range(12, hora_fin)) if hora_fin > 12 else 0.0
    tot = sum(h0.get(i, 0.0) for i in range(12, 24))
    d = d0 + timedelta(days=1)
    while d < d1:
        tot += sum(_horas_dia(horas, d).values())
        d += timedelta(days=1)
    h1 = _horas_dia(horas, d1)
    tot += sum(h1.get(i, 0.0) for i in range(0, hora_fin))
    return tot


def _fraccion_tarde(horas: dict[date, dict[int, float]], dia: date) -> float:
    h = _horas_dia(horas, dia)
    tot = sum(h.values())
    if tot <= 0:
        return 0.5
    return sum(h.get(i, 0.0) for i in range(12, 24)) / tot


def _fraccion_manana(horas: dict[date, dict[int, float]], dia: date, hora_fin: int) -> float:
    h = _horas_dia(horas, dia)
    tot = sum(h.values())
    if tot <= 0:
        return 0.5
    return sum(h.get(i, 0.0) for i in range(0, hora_fin)) / tot


def _listado_periodo(
    horas: dict[date, dict[int, float]],
    d0: date,
    d1: date,
    hora_fin: int = 12,
) -> float | None:
    """Suma Listado CSV del período, partiendo inicio/cierre con el mismo corte 12:00."""
    if d1 < d0:
        return None
    hay = False
    tot = 0.0
    if d0 == d1:
        if d0 in LISTADO:
            return _num(LISTADO[d0] * _fraccion_tarde(horas, d0))
        return None
    if d0 in LISTADO:
        hay = True
        tot += LISTADO[d0] * _fraccion_tarde(horas, d0)
    d = d0 + timedelta(days=1)
    while d < d1:
        if d in LISTADO:
            hay = True
            tot += LISTADO[d]
        d += timedelta(days=1)
    if d1 in LISTADO:
        hay = True
        tot += LISTADO[d1] * _fraccion_manana(horas, d1, hora_fin)
    return _num(tot) if hay else None


def _pintar_dif(cell, delta: float | None) -> None:
    if delta is None:
        return
    if delta > 0.05:
        cell.fill = AZUL_APP
        cell.font = Font(bold=True, color="FFFFFF", size=10)
    elif delta < -0.05:
        cell.fill = ROJO_CTA
        cell.font = Font(bold=True, color="FFFFFF", size=10)


def construir_facturaciones(wb: Workbook, horas: dict[date, dict[int, float]]) -> None:
    if "Facturaciones" in wb.sheetnames:
        del wb["Facturaciones"]
    ws = wb.create_sheet("Facturaciones", 1)
    headers = [
        "Mes",
        "Lectura inicial (12:00)",
        "Fecha lectura final y lectura (12:00)",
        "Diferencia entre lecturas (m³)",
        "Consumo app WES (Total consumo registro de la placa)",
        "Consumo app WES (Listado consumo sacado del csv de la placa)",
        "Diferencia lecturas con total",
        "Diferencia lecturas con Listado",
    ]
    ws.append(headers)
    for col in range(1, 9):
        c = ws.cell(1, col)
        c.fill = AZUL_HDR
        c.font = FONT_BLANCO
        c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        c.border = THIN
    ws.row_dimensions[1].height = 48
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
    ws["A2"] = f"{NOMBRE} ({NODE}) — un renglón por período de facturación. Corte Aguas Andinas 12:00."
    ws["A2"].font = Font(bold=True, size=11, color="003366")
    ws["A2"].alignment = Alignment(vertical="center")

    tot_cta = tot_placa = tot_lis = 0.0
    n_lis = 0
    for p in PERIODOS_FACTURA:
        placa = _total_placa_periodo(horas, p["d0"], p["d1"], p["hora_fin"])
        lista = _listado_periodo(horas, p["d0"], p["d1"], p["hora_fin"])
        dif_cta = float(p["dif"])
        d_tot = _num(placa) - dif_cta
        d_lis = (_num(lista) - dif_cta) if lista is not None else None
        ws.append(
            [
                p["mes"],
                p["ini"],
                p["fin"],
                _num(dif_cta),
                _num(placa),
                _num(lista) if lista is not None else None,
                _num(d_tot),
                _num(d_lis) if d_lis is not None else None,
            ]
        )
        row = ws.max_row
        if not str(p["mes"]).startswith("terreno"):
            tot_cta += dif_cta
            tot_placa += placa
            if lista is not None:
                tot_lis += lista
                n_lis += 1
        if p["estimado"]:
            for col in (1, 2, 3, 4):
                ws.cell(row, col).fill = VERDE_PROM
            ws.cell(row, 1).font = Font(bold=True, size=10, color="006100")
        ws.cell(row, 5).fill = CELESTE
        if lista is not None:
            ws.cell(row, 6).fill = VERDE
        _pintar_dif(ws.cell(row, 7), d_tot)
        _pintar_dif(ws.cell(row, 8), d_lis)
        for col in range(1, 9):
            ws.cell(row, col).border = THIN
            ws.cell(row, col).alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        for col in (4, 5, 6, 7, 8):
            ws.cell(row, col).number_format = NUM_FMT
        print(
            f"  FACT {p['mes']} cta={dif_cta:.1f} total={placa:.2f} listado="
            f"{lista if lista is not None else '-'} est={p['estimado']}",
            flush=True,
        )

    ws.append(
        [
            "TOTAL",
            "",
            "",
            _num(tot_cta),
            _num(tot_placa),
            _num(tot_lis) if n_lis else None,
            _num(tot_placa - tot_cta),
            _num(tot_lis - tot_cta) if n_lis else None,
        ]
    )
    last = ws.max_row
    for col in range(1, 9):
        cell = ws.cell(last, col)
        cell.font = Font(bold=True, size=10)
        cell.border = THIN
        cell.fill = GRIS
        cell.alignment = CENTER
    for col in (4, 5, 6, 7, 8):
        ws.cell(last, col).number_format = NUM_FMT
    ws.cell(last, 5).fill = CELESTE
    ws.cell(last, 5).font = Font(bold=True, size=10)
    if n_lis:
        ws.cell(last, 6).fill = VERDE
        ws.cell(last, 6).font = Font(bold=True, size=10)
    _pintar_dif(ws.cell(last, 7), tot_placa - tot_cta)
    if n_lis:
        _pintar_dif(ws.cell(last, 8), tot_lis - tot_cta)

    ws.freeze_panes = "A3"
    anchos = [16, 36, 40, 22, 28, 32, 24, 26]
    for i, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    nota = last + 2
    ws.merge_cells(start_row=nota, start_column=1, end_row=nota, end_column=8)
    ws.cell(
        nota,
        1,
        "Corte 12:00: día inicial 12:00–23:59 + días intermedios + día final 00:00–12:00 "
        "(terreno 25-09 corta a las 11:00). Total = suma horaria de la placa. "
        "Listado = CSV diario (01/03 a 28/09) partido con el mismo corte. "
        "Verde en Mes = cobro a promedio. Celeste = Total. Verde = Listado. "
        "Azul = app > cuenta. Rojo = cuenta > app.",
    )
    ws.cell(nota, 1).font = Font(size=9, italic=True, color="006100")
    ws.cell(nota, 1).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[nota].height = 42
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:2"
    ws.sheet_view.showGridLines = False


def _copiar_a_general(ws) -> Path | None:
    gen = OUT_DIR / "Lecturas_vs_WES_medio_dia_CORMUP_20260929_1418.xlsx"
    if not gen.is_file():
        return None
    gwb = load_workbook(gen)
    if "JP2 horario" in gwb.sheetnames:
        del gwb["JP2 horario"]
    gws = gwb.create_sheet("JP2 horario", 1)
    for row in ws.iter_rows():
        for c in row:
            dest = gws.cell(c.row, c.column, c.value)
            dest.font = copy(c.font)
            dest.fill = copy(c.fill)
            dest.alignment = copy(c.alignment)
            dest.border = copy(c.border)
            dest.number_format = c.number_format
    for m in ws.merged_cells.ranges:
        gws.merge_cells(str(m))
    gws.column_dimensions["A"].width = 12
    for i in range(2, ws.max_column + 1):
        gws.column_dimensions[get_column_letter(i)].width = 14
    gws.freeze_panes = "B4"
    gws.row_dimensions[1].height = 32
    gwb.save(gen)
    return gen


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    dias = _dias_unicos(D0, D1)
    prev = sorted(OUT_DIR.glob("Horario_JP2_hora_consumo_*.xlsx"))
    cache = _cargar_horas_xlsx(prev[-1]) if prev else {}
    print(f"[INFO] cache {len(cache)} días; faltan {sum(1 for d in dias if d not in cache)}", flush=True)
    horas = _horas_rango(dias, cache)
    extra = _dias_unicos(date(2025, 10, 29), D1)
    print(
        f"[INFO] facturaciones: extra {sum(1 for d in extra if d not in horas)} días",
        flush=True,
    )
    horas = _horas_rango(extra, horas)
    wb = construir_horario(dias, horas)
    construir_facturaciones(wb, horas)
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    out = OUT_DIR / f"Horario_JP2_hora_consumo_{ts}.xlsx"
    wb.save(out)
    estable = OUT_DIR / DRIVE_NOMBRE
    if out.resolve() != estable.resolve():
        estable.write_bytes(out.read_bytes())
    print("XLSX", out, "cols", wb.active.max_column)
    if credenciales_configuradas():
        info = subir_a_drive(estable, subcarpeta=DRIVE_SUB, nombre=DRIVE_NOMBRE)
        print("DRIVE", info["id"], info["web_view_link"])


if __name__ == "__main__":
    main()
