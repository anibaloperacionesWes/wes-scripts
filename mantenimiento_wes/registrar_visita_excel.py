# -*- coding: utf-8 -*-
"""Registra una visita completada en el Excel digital (hoja Datos)."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import openpyxl
from openpyxl.styles import Border, PatternFill, Side

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / "FORMULARIO_MANTENCION_WES_DIGITAL.xlsx"

FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

# columnas en hoja Datos (empezando en B=2)
COLS = {
    "cliente": 2,
    "maquina": 3,
    "tecnico": 4,
    "fecha": 5,
    "tipo_mtto": 6,
    "tipo_falla": 7,
    "falla_especifica": 8,
    "solucion": 9,
    "observaciones": 10,
    "firma": 11,
    "ot": 12,
    "estado": 13,
    "origen": 14,
}


def _parse_fecha(value: Any):
    if value is None or value == "":
        return datetime.now().date()
    if isinstance(value, datetime):
        return value.date()
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return s


def registrar_visita_en_excel(
    data: Dict[str, Any],
    xlsx_path: Optional[Path] = None,
) -> Tuple[Path, int]:
    path = Path(xlsx_path) if xlsx_path else XLSX
    if not path.is_file():
        raise FileNotFoundError(f"No existe Excel: {path}")

    wb = openpyxl.load_workbook(path)
    if "Datos" not in wb.sheetnames:
        raise RuntimeError("Falta hoja Datos en el Excel")
    if "Ingreso" in wb.sheetnames:
        # espejo en Ingreso para quien abre el archivo después
        ing = wb["Ingreso"]
        ing["C4"] = data.get("cliente")
        ing["C6"] = data.get("maquina")
        ing["C8"] = data.get("tecnico")
        ing["C10"] = _parse_fecha(data.get("fecha"))
        ing["C10"].number_format = "DD/MM/YYYY"
        ing["C12"] = data.get("tipo_mtto")
        ing["C14"] = data.get("tipo_falla")
        ing["C16"] = data.get("falla_especifica")
        ing["C18"] = data.get("solucion")
        ing["C20"] = data.get("observaciones")
        ing["C22"] = "Sí - acta firmada" if data.get("firma_png") else "No - solo digital"
        ing["C24"] = data.get("ot")
        ing["C26"] = data.get("estado_visita") or "cerrada"

    datos = wb["Datos"]
    row = 2
    while datos.cell(row, COLS["cliente"]).value not in (None, ""):
        row += 1
        if row > 8000:
            raise RuntimeError("Hoja Datos llena")

    valores = {
        "cliente": data.get("cliente"),
        "maquina": data.get("maquina"),
        "tecnico": data.get("tecnico"),
        "fecha": _parse_fecha(data.get("fecha")),
        "tipo_mtto": data.get("tipo_mtto"),
        "tipo_falla": data.get("tipo_falla"),
        "falla_especifica": data.get("falla_especifica"),
        "solucion": data.get("solucion"),
        "observaciones": data.get("observaciones"),
        "firma": (
            f"Sí - {data.get('recibido_por') or 'firmada'}"
            if data.get("firma_png") or data.get("recibido_por")
            else "No - solo digital"
        ),
        "ot": data.get("ot"),
        "estado": data.get("estado_visita") or "cerrada",
        "origen": f"Formulario web {datetime.now().strftime('%Y-%m-%d %H:%M')}",
    }
    for key, col in COLS.items():
        cell = datos.cell(row, col, valores.get(key))
        cell.fill = FILL_INPUT
        cell.border = THIN
        if key == "fecha":
            cell.number_format = "DD/MM/YYYY"

    wb.save(path)
    return path, row


if __name__ == "__main__":
    demo = {
        "cliente": "CORMUP",
        "maquina": "TOBALABA",
        "tecnico": "Anibal Aranda",
        "fecha": "2026-08-12",
        "tipo_mtto": "Mtto Preventivo",
        "tipo_falla": "Auditoría",
        "falla_especifica": "Validación Data",
        "solucion": "Demo registro",
        "observaciones": "",
        "ot": "OT-DEMO",
        "estado_visita": "cerrada",
        "recibido_por": "Demo",
        "firma_png": "x",
    }
    p, r = registrar_visita_en_excel(demo)
    print(f"OK {p} fila {r}")
