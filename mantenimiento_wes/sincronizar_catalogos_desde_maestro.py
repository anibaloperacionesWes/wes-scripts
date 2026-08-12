# -*- coding: utf-8 -*-
"""
Sincroniza catálogos Cliente/Máquina y Tipo falla desde el maestro
(Registro de fallas WES / analisis de falla) hacia:
  - mantenimiento_wes/catalogos/*.json  (formulario web)
  - FORMULARIO_MANTENCION_WES_DIGITAL.xlsx (hojas Base*)

También completa Base1 con pares Cliente+Máquina que existan en Datos
pero falten en el catálogo (ej. MOLYMET, PAE).

Uso:
  python sincronizar_catalogos_desde_maestro.py
  python sincronizar_catalogos_desde_maestro.py --desde-drive
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parent
MAESTRO_DIR = ROOT / "maestro"
CAT_DIR = ROOT / "catalogos"
XLSX_FORM = ROOT / "FORMULARIO_MANTENCION_WES_DIGITAL.xlsx"

# IDs Drive
SHEET_REGISTRO_ID = "1GlRn7QXWEre7ziau29ojR5lTl-bZ8T3mCT3cD93HZgM"
FILE_ANALISIS_ID = "1mzIsNG9Kr8PLZkUz_JDDklJu5uv1HJgC"

FILL_HEADER = PatternFill("solid", fgColor="1F4E79")
FONT_WHITE = Font(name="Calibri", bold=True, color="FFFFFF", size=11)

Pair = Tuple[str, str]


def _download_drive(file_id: str, dest: Path, *, export_xlsx: bool = False) -> Path:
    sys.path.insert(0, str(ROOT.parent))
    from wes_google_drive import obtener_servicio_drive
    from googleapiclient.http import MediaIoBaseDownload
    import io

    svc = obtener_servicio_drive()
    meta = svc.files().get(fileId=file_id, fields="mimeType,name").execute()
    mime = meta["mimeType"]
    if export_xlsx or mime == "application/vnd.google-apps.spreadsheet":
        req = svc.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        req = svc.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    dl = MediaIoBaseDownload(fh, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(fh.getvalue())
    return dest


def _norm(v) -> str:
    return "" if v is None else str(v).strip()


def _pairs_from_sheet(ws, cli_col: int, maq_col: int, start_row: int = 2) -> Set[Pair]:
    out: Set[Pair] = set()
    for r in range(start_row, (ws.max_row or 1) + 1):
        c = _norm(ws.cell(r, cli_col).value)
        m = _norm(ws.cell(r, maq_col).value)
        if c and m:
            out.add((c, m))
    return out


def _fallas_from_sheet(ws, tipo_col: int = 2, esp_col: int = 3) -> Dict[str, List[str]]:
    d: Dict[str, List[str]] = defaultdict(list)
    for r in range(2, (ws.max_row or 1) + 1):
        t = _norm(ws.cell(r, tipo_col).value)
        e = _norm(ws.cell(r, esp_col).value)
        if t and e and e not in d[t]:
            d[t].append(e)
    return dict(d)


def _write_base1(ws, pairs: Iterable[Pair]) -> int:
    # clear old
    if ws.max_row and ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    ws["B1"] = "CLIENTE"
    ws["C1"] = "MAQUINA"
    ws["B1"].fill = FILL_HEADER
    ws["C1"].fill = FILL_HEADER
    ws["B1"].font = FONT_WHITE
    ws["C1"].font = FONT_WHITE
    n = 0
    for cli, maq in sorted(pairs, key=lambda x: (x[0].lower(), x[1].lower())):
        n += 1
        ws.cell(n + 1, 2, cli)
        ws.cell(n + 1, 3, maq)
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 42
    return n


def _write_base3(ws, fallas: Dict[str, List[str]]) -> int:
    if ws.max_row and ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    ws["B1"] = "Tipo de falla"
    ws["C1"] = "Falla Específica"
    ws["B1"].fill = FILL_HEADER
    ws["C1"].fill = FILL_HEADER
    ws["B1"].font = FONT_WHITE
    ws["C1"].font = FONT_WHITE
    n = 0
    for tipo in sorted(fallas.keys()):
        for esp in fallas[tipo]:
            n += 1
            ws.cell(n + 1, 2, tipo)
            ws.cell(n + 1, 3, esp)
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 36
    return n


def _update_filter_helpers(wb, n1: int, n3: int) -> None:
    if "Base2" not in wb.sheetnames:
        wb.create_sheet("Base2")
    ws2 = wb["Base2"]
    ws2["A1"] = f'=IFERROR(FILTER(Base1!B2:C{n1 + 1},Base1!B2:B{n1 + 1}=Ingreso!C4),"")'
    if "Base 4" not in wb.sheetnames:
        wb.create_sheet("Base 4")
    ws4 = wb["Base 4"]
    ws4["A1"] = f'=IFERROR(FILTER(Base3!B2:C{n3 + 1},Base3!B2:B{n3 + 1}=Ingreso!C14),"")'


def _clientes_dict(pairs: Set[Pair]) -> Dict[str, List[str]]:
    d: Dict[str, List[str]] = defaultdict(list)
    for cli, maq in sorted(pairs):
        if maq not in d[cli]:
            d[cli].append(maq)
    return dict(sorted(d.items()))


def sincronizar(
    *,
    desde_drive: bool = False,
    fuente: Optional[Path] = None,
) -> dict:
    MAESTRO_DIR.mkdir(parents=True, exist_ok=True)
    CAT_DIR.mkdir(parents=True, exist_ok=True)

    if desde_drive:
        fuente = _download_drive(
            SHEET_REGISTRO_ID,
            MAESTRO_DIR / "analisis_falla_google.xlsx",
            export_xlsx=True,
        )
        _download_drive(FILE_ANALISIS_ID, MAESTRO_DIR / "analisis_de_falla.xlsx")
    elif fuente is None:
        cand = [
            MAESTRO_DIR / "analisis_falla_google.xlsx",
            MAESTRO_DIR / "analisis_de_falla.xlsx",
            XLSX_FORM,
        ]
        fuente = next((p for p in cand if p.is_file()), None)
        if fuente is None:
            raise FileNotFoundError("No hay maestro local. Usá --desde-drive")

    wb_src = load_workbook(fuente, data_only=True)

    # Pares Base1
    pairs: Set[Pair] = set()
    if "Base1" in wb_src.sheetnames:
        pairs |= _pairs_from_sheet(wb_src["Base1"], 2, 3)
    # Enriquecer desde historial Datos / Data
    for sh_name, ccol, mcol in (("Datos", 2, 3), ("Data", 2, 3)):
        if sh_name in wb_src.sheetnames:
            extra = _pairs_from_sheet(wb_src[sh_name], ccol, mcol)
            pairs |= extra

    # Fallas
    if "Base3" in wb_src.sheetnames:
        fallas = _fallas_from_sheet(wb_src["Base3"])
    else:
        fallas = {}

    # JSON catálogos para el formulario web
    clientes = _clientes_dict(pairs)
    (CAT_DIR / "clientes_maquinas.json").write_text(
        json.dumps(clientes, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (CAT_DIR / "tipos_falla.json").write_text(
        json.dumps(fallas, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # Actualizar Excel digital si existe
    if XLSX_FORM.is_file():
        wb = load_workbook(XLSX_FORM)
        if "Base1" not in wb.sheetnames:
            wb.create_sheet("Base1")
        if "Base3" not in wb.sheetnames:
            wb.create_sheet("Base3")
        n1 = _write_base1(wb["Base1"], pairs)
        n3 = _write_base3(wb["Base3"], fallas)
        _update_filter_helpers(wb, n1, n3)
        # stamp en Instrucciones
        if "Instrucciones" in wb.sheetnames:
            wb["Instrucciones"]["A22"] = (
                f"Catálogos sincronizados: {datetime.now().strftime('%Y-%m-%d %H:%M')} "
                f"desde {fuente.name} · {len(clientes)} clientes · {n1} máquinas · {n3} fallas"
            )
        wb.save(XLSX_FORM)
    else:
        n1 = len(pairs)
        n3 = sum(len(v) for v in fallas.values())

    resumen = {
        "fuente": str(fuente),
        "clientes": len(clientes),
        "pares_maquina": n1,
        "tipos_falla": len(fallas),
        "fallas_especificas": n3,
        "clientes_lista": sorted(clientes.keys()),
        "sync_at": datetime.now().isoformat(timespec="seconds"),
    }
    (CAT_DIR / "ultima_sincronizacion.json").write_text(
        json.dumps(resumen, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return resumen


def main() -> int:
    parser = argparse.ArgumentParser(description="Sincroniza catálogos desde maestro de fallas")
    parser.add_argument("--desde-drive", action="store_true", help="Descarga Sheet/Excel desde Drive")
    parser.add_argument("--fuente", type=Path, default=None, help="Ruta local a xlsx maestro")
    args = parser.parse_args()
    info = sincronizar(desde_drive=args.desde_drive, fuente=args.fuente)
    print("OK sync catálogos")
    print(f"  fuente: {info['fuente']}")
    print(f"  clientes: {info['clientes']}")
    print(f"  máquinas: {info['pares_maquina']}")
    print(f"  fallas: {info['fallas_especificas']} en {info['tipos_falla']} tipos")
    print("  clientes:", ", ".join(info["clientes_lista"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
