"""
Horario por colegio CORMUP desde Drive:

  G:\\Mi unidad\\Colegios\\Peñalolén\\Facturaciones\\data de placas.xlsx

Una hoja H.* por establecimiento (F/H/M de la placa). Las últimas 3 hojas de
esa planilla (Erasmo Escala, Matilde Huici, CE Valle Hermoso) están vacías:
quedan como pendiente de descargar.

Uso:
  python generar_horario_placas_cormup.py
"""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from generar_comparativo_facturaciones_cormup_penalolen import (
    _buscar_carpeta_facturaciones,
    _descargar_archivo,
    _listar_hijos,
)
from generar_horario_jp2_hora_consumo import (
    AZUL_HDR,
    CELESTE,
    CENTER,
    FONT_OK,
    HDR,
    NARANJA,
    NUM_FMT,
    SUB,
    THIN,
    VERDE,
    construir_horario,
)
from wes_google_drive import credenciales_configuradas, obtener_servicio_drive, subir_a_drive

OUT_DIR = Path("reports/CORMUP/Facturaciones_vs_WES")
DRIVE_SUB = "CORMUP/Facturaciones_vs_WES"
DRIVE_NOMBRE = "Horario_placas_CORMUP.xlsx"
LOCAL_XLSX = OUT_DIR / "_pdfs" / "data_de_placas.xlsx"
PLACAS_NOMBRE = "data de placas.xlsx"

# Nombre en data de placas.xlsx → (nombre, nodo, hoja destino)
COLEGIOS = [
    ("171", "Liceo Antonio Hermida", "000008-01", "H. 171"),
    ("eduardo de la barra", "Eduardo de la Barra", "000008-02", "H. E. Barra"),
    ("Carlos Fernandez Peña", "Carlos Fernández Peña", "000008-03", "H. C.F. Peña"),
    ("tobalaba", "Tobalaba", "000008-04", "H. Tobalaba"),
    ("santa mari", "Santa María", "000008-05", "H. Santa María"),
    ("Luis Arrieta Cañas", "Luis Arrieta Cañas", "000008-06", "H. Arrieta"),
    ("erasmo escala", "Erasmo Escala", "000008-07", "H. Erasmo"),
    ("alicura", "Alicura", "000008-08", "H. Alicura"),
    ("juan bta pastenes", "Juan Bautista Pastenes", "000008-09", "H. J.B. Pastenes"),
    ("matilde ", "Matilde Huici Navas", "000008-10", "H. Matilde"),
    ("ce valle", "CE Valle Hermoso", "000008-11", "H. Valle Hermoso"),
    ("Union Nacional ", "Unión Nacional Árabe", "000008-12", "H. Unión Árabe"),
    ("Likankura", "Likancura", "000008-13", "H. Likancura"),
    ("JP II", "Juan Pablo II", "000008-14", "H. JPII"),
]
PENDIENTES = {"000008-07", "000008-10", "000008-11"}

PAT_F = re.compile(r"^F:\s*(\d{2}/\d{2}/\d{4})")
PAT_H = re.compile(r"^H:\s*(\d{1,2}):00")
PAT_M = re.compile(r"^M:\s*([\d.,]+)")
PAT_R = re.compile(r"^R:")


def descargar_data_de_placas() -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    LOCAL_XLSX.parent.mkdir(parents=True, exist_ok=True)
    if not credenciales_configuradas():
        if LOCAL_XLSX.is_file():
            print(f"[INFO] sin Drive; uso {LOCAL_XLSX}", flush=True)
            return LOCAL_XLSX
        raise FileNotFoundError(LOCAL_XLSX)
    service = obtener_servicio_drive()
    folder_id = _buscar_carpeta_facturaciones(service)
    hijos = _listar_hijos(service, folder_id)
    target = next(
        (
            h
            for h in hijos
            if h["name"].lower().replace(" ", "") == PLACAS_NOMBRE.lower().replace(" ", "")
            or "data de placas" in h["name"].lower()
        ),
        None,
    )
    if not target:
        raise FileNotFoundError("No está data de placas.xlsx en Colegios/Peñalolén/Facturaciones")
    print(f"[Drive] {target['name']} id={target['id']}", flush=True)
    _descargar_archivo(service, target["id"], LOCAL_XLSX)
    print(f"[Drive] bajado {LOCAL_XLSX} ({LOCAL_XLSX.stat().st_size} bytes)", flush=True)
    return LOCAL_XLSX


def _parse_fhm_sheet(ws) -> tuple[dict[date, dict[int, float]], dict]:
    """Lee F:/H:/M: por columna. Si un (fecha, hora) se repite, se queda la primera."""
    col_tokens: dict[int, list[str]] = {}
    for col in range(1, ws.max_column + 1):
        toks: list[str] = []
        for row in range(1, ws.max_row + 1):
            raw = ws.cell(row, col).value
            if raw is None:
                continue
            s = str(raw).strip()
            if s:
                toks.append(s)
        if toks:
            col_tokens[col] = toks

    horas: dict[date, dict[int, float]] = {}
    n_triples = 0
    n_dup = 0
    n_conf = 0
    for tokens in col_tokens.values():
        pending_f: date | None = None
        pending_h: int | None = None
        for tok in tokens:
            if PAT_R.match(tok):
                pending_f = pending_h = None
                continue
            mf, mh, mm = PAT_F.match(tok), PAT_H.match(tok), PAT_M.match(tok)
            if mf:
                pending_f = datetime.strptime(mf.group(1), "%d/%m/%Y").date()
                pending_h = None
                continue
            if mh and pending_f is not None:
                pending_h = int(mh.group(1))
                continue
            if mm and pending_f is not None and pending_h is not None:
                val = float(mm.group(1).replace(",", "."))
                n_triples += 1
                prev = (horas.get(pending_f) or {}).get(pending_h)
                if prev is not None:
                    n_dup += 1
                    if abs(prev - val) > 0.001:
                        n_conf += 1
                else:
                    horas.setdefault(pending_f, {})[pending_h] = val
                pending_f = pending_h = None
    dias = sorted(horas)
    stats = {
        "triples": n_triples,
        "unicos": sum(len(h) for h in horas.values()),
        "duplicados": n_dup,
        "conflictos": n_conf,
        "dias": len(dias),
        "d0": dias[0] if dias else None,
        "d1": dias[-1] if dias else None,
        "m3": round(sum(sum(h.values()) for h in horas.values()), 2),
    }
    return horas, stats


def _hoja_pendiente(wb: Workbook, sheet_name: str, nombre: str, node: str) -> None:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    ws["A1"] = (
        f"{nombre} ({node}) — data de placas pendiente de descargar. "
        "En la planilla de Drive esta hoja está vacía; cuando esté el pegado F/H/M "
        "se arma el horario igual que en los otros colegios."
    )
    ws["A1"].font = Font(bold=True, size=12, color="003366")
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 48
    ws["A3"] = "Estado"
    ws["B3"] = "Pendiente de descargar"
    ws["A3"].fill = NARANJA
    ws["B3"].fill = NARANJA
    ws["A3"].font = FONT_OK
    ws["B3"].font = FONT_OK
    ws["A3"].border = THIN
    ws["B3"].border = THIN
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 36


def _escribir_resumen(wb: Workbook, filas: list[dict]) -> None:
    if "Resumen" in wb.sheetnames:
        del wb["Resumen"]
    ws = wb.create_sheet("Resumen", 0)
    headers = [
        "Colegio",
        "Nodo",
        "Desde",
        "Hasta",
        "Días",
        "Horas únicas",
        "Duplicados",
        "m³ placa",
        "Estado",
    ]
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        c = ws.cell(1, col)
        c.fill = AZUL_HDR
        c.font = Font(color="FFFFFF", bold=True, size=9)
        c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        c.border = THIN
    ws.row_dimensions[1].height = 28
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    ws["A2"] = (
        "Fuente: Colegios / Peñalolén / Facturaciones / data de placas.xlsx. "
        "Una hoja horaria por colegio. Si un horario se repetía, queda una sola vez "
        "(el Total no suma dos veces). Erasmo Escala, Matilde Huici y CE Valle Hermoso "
        "están pendientes de descargar."
    )
    ws["A2"].font = Font(bold=True, size=11, color="003366")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 36

    for info in filas:
        pendiente = info["pendiente"]
        ws.append(
            [
                info["nombre"],
                info["node"],
                info["d0"].strftime("%d/%m/%Y") if info["d0"] else "—",
                info["d1"].strftime("%d/%m/%Y") if info["d1"] else "—",
                info["dias"] or None,
                info["unicos"] or None,
                info["duplicados"] if not pendiente else None,
                info["m3"] if not pendiente else None,
                "Pendiente de descargar" if pendiente else "OK",
            ]
        )
        row = ws.max_row
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row, col)
            cell.border = THIN
            cell.alignment = CENTER
        ws.cell(row, 1).alignment = Alignment(horizontal="left", vertical="center")
        if pendiente:
            for col in range(1, len(headers) + 1):
                ws.cell(row, col).fill = NARANJA
        else:
            ws.cell(row, 8).number_format = NUM_FMT
            ws.cell(row, 8).fill = CELESTE
            if info["duplicados"]:
                ws.cell(row, 7).fill = PatternFill("solid", fgColor="FFFF99")
            ws.cell(row, 9).fill = VERDE
            ws.cell(row, 9).font = FONT_OK

    anchos = [28, 14, 14, 14, 10, 14, 14, 14, 26]
    for i, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A3"
    ws.page_setup.orientation = "landscape"
    ws.sheet_view.showGridLines = False


def main() -> None:
    path = descargar_data_de_placas()
    src = load_workbook(path, data_only=False)
    print("[INFO] hojas fuente", src.sheetnames, flush=True)

    wb = Workbook()
    tmp = wb.active
    tmp.title = "_tmp"
    filas_resumen: list[dict] = []

    for src_name, nombre, node, dest in COLEGIOS:
        if src_name not in src.sheetnames:
            print(f"[WARN] no está la hoja {src_name!r} en data de placas", flush=True)
            _hoja_pendiente(wb, dest, nombre, node)
            filas_resumen.append(
                {
                    "nombre": nombre,
                    "node": node,
                    "d0": None,
                    "d1": None,
                    "dias": 0,
                    "unicos": 0,
                    "duplicados": 0,
                    "m3": 0.0,
                    "pendiente": True,
                }
            )
            continue
        horas, st = _parse_fhm_sheet(src[src_name])
        pendiente = node in PENDIENTES or not horas
        print(
            f"  {dest:18} {nombre:24} {st['triples']:5} triples "
            f"únicos={st['unicos']} dups={st['duplicados']} conf={st['conflictos']} "
            f"días={st['dias']} m³={st['m3']}"
            f"{' PENDIENTE' if pendiente else ''}",
            flush=True,
        )
        filas_resumen.append(
            {
                "nombre": nombre,
                "node": node,
                "d0": st["d0"],
                "d1": st["d1"],
                "dias": st["dias"],
                "unicos": st["unicos"],
                "duplicados": st["duplicados"],
                "m3": st["m3"],
                "pendiente": pendiente,
            }
        )
        if pendiente:
            _hoja_pendiente(wb, dest, nombre, node)
            continue
        dias = sorted(horas)
        nota = (
            "Fuente: data de placas.xlsx (pegado F/H/M de la placa). "
            "Si un horario se repetía, queda una sola vez; el Total no suma dos veces. "
            f"Duplicados={st['duplicados']} (conflictos de valor={st['conflictos']}). "
            "Fila Total = celeste. Solo días con registro."
        )
        construir_horario(
            dias,
            horas,
            wb=wb,
            sheet_name=dest,
            nombre=nombre,
            node=node,
            d0=dias[0],
            d1=dias[-1],
            listado=None,
            nota=nota,
        )

    if "_tmp" in wb.sheetnames:
        del wb["_tmp"]
    _escribir_resumen(wb, filas_resumen)

    ts = datetime.now().strftime("%Y%m%d_%H%M")
    out = OUT_DIR / f"Horario_placas_CORMUP_{ts}.xlsx"
    wb.save(out)
    estable = OUT_DIR / DRIVE_NOMBRE
    if out.resolve() != estable.resolve():
        estable.write_bytes(out.read_bytes())
    print("XLSX", out, "sheets", wb.sheetnames)
    if credenciales_configuradas():
        info = subir_a_drive(estable, subcarpeta=DRIVE_SUB, nombre=DRIVE_NOMBRE)
        print("DRIVE", info["id"], info["web_view_link"])


if __name__ == "__main__":
    main()
