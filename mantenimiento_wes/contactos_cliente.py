# -*- coding: utf-8 -*-
"""
Contactos por cliente/sitio para actas WES (varios CC).

Hoja Excel «Contactos» + JSON catalogos/contactos_cliente.json

Columnas:
  Cliente | Sitio | Rol | Nombre | Email | Firmante habitual | Enviar TO | Enviar CC | Activo | Notas

Sitio vacío = aplica a todo el cliente.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / "FORMULARIO_MANTENCION_WES_DIGITAL.xlsx"
JSON_OUT = ROOT / "catalogos" / "contactos_cliente.json"

HEADERS = [
    "Cliente",
    "Sitio",
    "Rol",
    "Nombre",
    "Email",
    "Firmante habitual",
    "Enviar TO",
    "Enviar CC",
    "Activo",
    "Notas",
]

# Semilla inicial (MAE y malls PA). Completar emails reales en la planilla.
_PA_MALLS = ("MAE", "PAK", "AEB", "BOM", "CUR", "MAM", "MAQ")
_ROLES = (
    ("Jefe de operaciones", False, False, True, "Reportar acta de visita"),
    ("Líder de medio ambiente", False, False, True, "Copia ambiental"),
    ("Mantención Linkes", True, True, True, "Apoya la visita; suele firmar"),
)

SEED: List[Dict[str, str]] = []
for cli in _PA_MALLS:
    for rol, firmante, to, cc, notas in _ROLES:
        SEED.append({
            "Cliente": cli,
            "Sitio": "",
            "Rol": rol,
            "Nombre": "",
            "Email": "",
            "Firmante habitual": "Sí" if firmante else "No",
            "Enviar TO": "Sí" if to else "No",
            "Enviar CC": "Sí" if cc else "No",
            "Activo": "Sí",
            "Notas": notas,
        })

FILL_HEADER = PatternFill("solid", fgColor="1F4E79")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FONT_WHITE = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def _yes(v: Any) -> bool:
    s = str(v or "").strip().lower()
    return s in {"si", "sí", "yes", "y", "1", "true", "activo"}


def _email_ok(v: Any) -> bool:
    s = str(v or "").strip()
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", s))


def _read_existing(ws) -> List[Dict[str, str]]:
    rows = []
    headers = [ws.cell(1, c).value for c in range(1, len(HEADERS) + 1)]
    if not headers or headers[0] != "Cliente":
        return rows
    for r in range(2, (ws.max_row or 1) + 1):
        if not ws.cell(r, 1).value and not ws.cell(r, 5).value:
            continue
        item = {}
        for i, h in enumerate(HEADERS, start=1):
            val = ws.cell(r, i).value
            item[h] = "" if val is None else str(val).strip()
        rows.append(item)
    return rows


def _merge_seed(existing: List[Dict[str, str]]) -> List[Dict[str, str]]:
    out = list(existing)
    for s in SEED:
        soft = (s["Cliente"].upper(), s["Sitio"].upper(), s["Rol"].upper())
        already = any(
            (
                e.get("Cliente", "").upper(),
                e.get("Sitio", "").upper(),
                e.get("Rol", "").upper(),
            )
            == soft
            for e in existing
        )
        if already:
            continue
        out.append(dict(s))
    out.sort(
        key=lambda d: (
            d.get("Cliente", ""),
            d.get("Sitio", ""),
            d.get("Rol", ""),
            d.get("Nombre", ""),
        )
    )
    return out


def asegurar_hoja_contactos(xlsx: Path = XLSX) -> Path:
    wb = load_workbook(xlsx)
    if "Contactos" in wb.sheetnames:
        ws = wb["Contactos"]
        existing = _read_existing(ws)
    else:
        ws = wb.create_sheet("Contactos")
        existing = []

    rows = _merge_seed(existing)

    if ws.max_row and ws.max_row > 0:
        ws.delete_rows(1, ws.max_row)

    for i, h in enumerate(HEADERS, start=1):
        cell = ws.cell(1, i, h)
        cell.fill = FILL_HEADER
        cell.font = FONT_WHITE
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN
        ws.column_dimensions[cell.column_letter].width = 22 if h != "Notas" else 36

    for r, item in enumerate(rows, start=2):
        for c, h in enumerate(HEADERS, start=1):
            cell = ws.cell(r, c, item.get(h, ""))
            cell.fill = FILL_INPUT
            cell.border = THIN

    ws.data_validations.dataValidation = []
    for col, opts in (
        ("F", '"Sí,No"'),
        ("G", '"Sí,No"'),
        ("H", '"Sí,No"'),
        ("I", '"Sí,No"'),
    ):
        dv = DataValidation(type="list", formula1=opts, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}2000")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:J{max(2, len(rows) + 1)}"

    if "Instrucciones" in wb.sheetnames:
        inst = wb["Instrucciones"]
        note = (
            "Contactos multi-CC: hoja Contactos. "
            "Completá Email / Rol (JO, Medio ambiente, Mantención Linkes). "
            "Firmante habitual = quien suele firmar. Enviar TO/CC = destino del PDF."
        )
        found = False
        for r in range(1, min(50, (inst.max_row or 1) + 1)):
            val = inst.cell(r, 1).value
            if val and "Contactos multi-CC" in str(val):
                inst.cell(r, 1, note)
                found = True
                break
        if not found:
            inst.cell((inst.max_row or 1) + 2, 1, note)

    order = [
        "Instrucciones",
        "Ingreso",
        "Datos",
        "Contactos",
        "Resumen",
        "Formulario Visita",
        "Base1",
        "Base2",
        "Base3",
        "Base 4",
    ]
    for idx, name in enumerate(order):
        if name in wb.sheetnames:
            wb.move_sheet(name, offset=idx - wb.sheetnames.index(name))

    wb.save(xlsx)
    export_json(rows)
    return xlsx


def export_json(rows: Optional[List[Dict[str, str]]] = None) -> Path:
    if rows is None:
        wb = load_workbook(XLSX, data_only=True)
        rows = _read_existing(wb["Contactos"]) if "Contactos" in wb.sheetnames else []

    payload = []
    for r in rows:
        if r.get("Activo") and not _yes(r.get("Activo")):
            continue
        payload.append({
            "cliente": r.get("Cliente", ""),
            "sitio": r.get("Sitio", ""),
            "rol": r.get("Rol", ""),
            "nombre": r.get("Nombre", ""),
            "email": r.get("Email", ""),
            "firmante": _yes(r.get("Firmante habitual")),
            "enviar_to": _yes(r.get("Enviar TO")),
            "enviar_cc": _yes(r.get("Enviar CC")),
            "activo": True if r.get("Activo") in ("", None) else _yes(r.get("Activo")),
            "notas": r.get("Notas", ""),
            "email_ok": _email_ok(r.get("Email")),
        })
    JSON_OUT.parent.mkdir(parents=True, exist_ok=True)
    JSON_OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return JSON_OUT


def contactos_para(cliente: str, sitio: str = "") -> List[Dict[str, Any]]:
    if not JSON_OUT.is_file():
        return []
    data = json.loads(JSON_OUT.read_text(encoding="utf-8"))
    cli = (cliente or "").strip().upper()
    sit = (sitio or "").strip().upper()
    out = []
    for c in data:
        if not c.get("activo", True):
            continue
        if (c.get("cliente") or "").upper() != cli:
            continue
        c_sit = (c.get("sitio") or "").upper()
        if c_sit and sit and c_sit != sit:
            continue
        if c_sit and not sit:
            continue
        out.append(c)
    return out


def emails_to_cc(cliente: str, sitio: str = "") -> Dict[str, Any]:
    contacts = contactos_para(cliente, sitio)
    to_list: List[str] = []
    cc_list: List[str] = []
    firmantes = []
    for c in contacts:
        em = (c.get("email") or "").strip()
        if not _email_ok(em):
            continue
        if c.get("enviar_to"):
            to_list.append(em)
        if c.get("enviar_cc"):
            cc_list.append(em)
        if c.get("firmante"):
            firmantes.append({
                "nombre": c.get("nombre") or "",
                "email": em,
                "rol": c.get("rol") or "",
            })

    def uniq(xs: List[str]) -> List[str]:
        seen = set()
        out = []
        for x in xs:
            k = x.lower()
            if k in seen:
                continue
            seen.add(k)
            out.append(x)
        return out

    return {
        "to": uniq(to_list),
        "cc": uniq(cc_list),
        "firmantes": firmantes,
        "contactos": contacts,
    }


if __name__ == "__main__":
    path = asegurar_hoja_contactos()
    print(f"OK Contactos en {path}")
    print(f"JSON {JSON_OUT} ({JSON_OUT.stat().st_size} bytes)")
