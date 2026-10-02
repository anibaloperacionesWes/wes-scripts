"""
Horario por colegio CORMUP desde Drive:

  G:\\Mi unidad\\Colegios\\Peñalolén\\Facturaciones\\data de placas.xlsx

Una hoja H.* por establecimiento (F/H/M de la placa). Fila 28 = data placa
(suma de horas). Fila 29 = data app (totales diarios de la app). Las últimas
3 hojas de esa planilla (Erasmo Escala, Matilde Huici, CE Valle Hermoso)
están vacías: quedan como pendiente de descargar.

Hojas Placa vs cuenta / Detalle placa vs cuenta: m³ de la placa (sin
duplicar horas, corte 12:00) contra la boleta Aguas Andinas, por mes.

Uso:
  python generar_horario_placas_cormup.py
"""

from __future__ import annotations

import json
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from generar_horario_jp2_hora_consumo import (
    AZUL_APP,
    AZUL_HDR,
    CELESTE,
    CENTER,
    DORADO,
    FONT_BLANCO,
    FONT_OK,
    HDR,
    NARANJA,
    NUM_FMT,
    PCT_FMT,
    ROJO,
    ROJO_CTA,
    SUB,
    THIN,
    VERDE,
    VERDE_PROM,
    _error_pct,
    _fmt_lectura,
    _pintar_dif,
    construir_horario,
)
from wes_google_drive import credenciales_configuradas, obtener_servicio_drive, subir_a_drive

OUT_DIR = Path("reports/CORMUP/Facturaciones_vs_WES")
DRIVE_SUB = "CORMUP/Facturaciones_vs_WES"
DRIVE_NOMBRE = "Horario_placas_CORMUP.xlsx"
LOCAL_XLSX = OUT_DIR / "_pdfs" / "data_de_placas.xlsx"
PLACAS_NOMBRE = "data de placas.xlsx"
APP_DIR = OUT_DIR / "app_data"

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
NOMBRE_NODO = {node: nombre for _src, nombre, node, _dest in COLEGIOS}
MESES_CORTOS_CTA = {
    1: "ene",
    2: "feb",
    3: "mar",
    4: "abr",
    5: "may",
    6: "jun",
    7: "jul",
    8: "ago",
    9: "sep",
    10: "oct",
    11: "nov",
    12: "dic",
}

PAT_F = re.compile(r"^F:\s*(\d{2}/\d{2}/\d{4})")
PAT_H = re.compile(r"^H:\s*(\d{1,2}):00")
PAT_M = re.compile(r"^M:\s*([\d.,]+)")
PAT_R = re.compile(r"^R:")


def _cargar_data_app(node: str) -> dict[date, float]:
    """Totales diarios de la app (suma de horas del dump), por nodo."""
    path = APP_DIR / f"{node}_diario.json"
    if not path.is_file():
        return {}
    raw = json.loads(path.read_text(encoding="utf-8"))
    return {date.fromisoformat(str(k)): float(v) for k, v in raw.items()}


def _meta_data_app(node: str) -> dict[str, Any]:
    path = APP_DIR / f"{node}_meta.json"
    if not path.is_file():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def descargar_data_de_placas() -> Path:
    from generar_comparativo_facturaciones_cormup_penalolen import (
        _buscar_carpeta_facturaciones,
        _descargar_archivo,
        _listar_hijos,
    )

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


def _parse_fhm_sheet(
    ws,
) -> tuple[dict[date, dict[int, float]], dict, set[tuple[date, int]]]:
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
    dups_pares: set[tuple[date, int]] = set()
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
                    dups_pares.add((pending_f, pending_h))
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
        "dias_dup": len({d for d, _h in dups_pares}),
        "conflictos": n_conf,
        "dias": len(dias),
        "d0": dias[0] if dias else None,
        "d1": dias[-1] if dias else None,
        "m3": round(sum(sum(h.values()) for h in horas.values()), 2),
    }
    return horas, stats, dups_pares


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
        "Días app",
        "m³ app",
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
        "Fuente placa: Colegios / Peñalolén / Facturaciones / data de placas.xlsx. "
        "Una hoja horaria por colegio. Fila data placa = suma de horas de la placa "
        "(si un horario se repetía, queda una sola vez). Fila data app = totales "
        "diarios de la app. En cada hoja, encabezado y hora en rojo = ese día/"
        "horario se repetía en el pegado de la placa. Fila data app en rojo = no "
        "cuadra con data placa. Estado: Placa cargada = llegó el pegado F/H/M; "
        "Pendiente de placa = hay data app pero la hoja de la placa sigue vacía. "
        "Erasmo Escala, Matilde Huici y CE Valle Hermoso están pendientes de descargar. "
        "Hoja Placa vs cuenta: m³ placa (corte 12:00, sin duplicar) contra la boleta."
    )
    ws["A2"].font = Font(bold=True, size=11, color="003366")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 52

    for info in filas:
        pendiente = info["pendiente"]
        n_app = int(info.get("dias_app") or 0)
        m3_app = info.get("m3_app")
        if pendiente and n_app:
            estado = "Pendiente de placa"
        elif pendiente:
            estado = "Pendiente de descargar"
        else:
            estado = "Placa cargada"
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
                n_app or None,
                m3_app if n_app else None,
                estado,
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
            if n_app:
                ws.cell(row, 10).number_format = NUM_FMT
                ws.cell(row, 10).fill = VERDE
        else:
            ws.cell(row, 8).number_format = NUM_FMT
            ws.cell(row, 8).fill = CELESTE
            if info["duplicados"]:
                ws.cell(row, 7).fill = PatternFill("solid", fgColor="FFFF99")
            if n_app:
                ws.cell(row, 10).number_format = NUM_FMT
                ws.cell(row, 10).fill = VERDE
            ws.cell(row, 11).fill = VERDE
            ws.cell(row, 11).font = FONT_OK

    anchos = [28, 14, 14, 14, 10, 14, 14, 14, 12, 14, 26]
    for i, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A3"
    ws.page_setup.orientation = "landscape"
    ws.sheet_view.showGridLines = False


def _mes_label_cta(y: int, m: int) -> str:
    return f"{MESES_CORTOS_CTA[m]}-{y}"


def _placa_en_periodo(
    horas: dict[date, dict[int, float]],
    d0: date,
    d1: date,
    hora_fin: int = 12,
) -> tuple[float | None, int, int]:
    """Suma placa (primera ocurrencia) con corte 12:00. (m³, horas con dato, horas esperadas)."""
    if d1 < d0:
        return None, 0, 0
    slots: list[tuple[date, int]] = []
    if d0 == d1:
        slots = [(d0, h) for h in range(12, hora_fin)]
    else:
        slots.extend((d0, h) for h in range(12, 24))
        d = d0 + timedelta(days=1)
        while d < d1:
            slots.extend((d, h) for h in range(24))
            d += timedelta(days=1)
        slots.extend((d1, h) for h in range(0, hora_fin))
    esperadas = len(slots)
    tot = 0.0
    n = 0
    for dia, h in slots:
        val = (horas.get(dia) or {}).get(h)
        if val is not None:
            tot += float(val)
            n += 1
    if n == 0:
        return None, 0, esperadas
    return tot, n, esperadas


def _cargar_boletas_cuenta():
    from facturacion_aguas_andinas_pdf import extraer_texto_pdf
    from generar_comparativo_facturaciones_cormup_penalolen import (
        PDF_CACHE,
        _cargar_sitios,
        descargar_facturaciones,
    )
    from generar_hojas_lecturas_medio_dia_cormup import _lecturas_medidor

    locales: dict[str, Path] = {}
    if credenciales_configuradas():
        try:
            locales = descargar_facturaciones(PDF_CACHE)
        except Exception as exc:
            print(f"[WARN] descarga facturaciones: {exc}", flush=True)
    if not locales:
        locales = {p.name: p for p in PDF_CACHE.iterdir() if p.is_dir()}
    locales = {k: v for k, v in locales.items() if k not in {"_raiz", "_pdfs"}}
    if not locales:
        print("[WARN] no hay PDFs de facturaciones para cruzar con la placa", flush=True)
        return None
    sitios = _cargar_sitios(locales)
    return sitios, PDF_CACHE, extraer_texto_pdf, _lecturas_medidor


def _filas_placa_vs_cuenta(
    horas_por_nodo: dict[str, dict[date, dict[int, float]]],
) -> list[dict]:
    loaded = _cargar_boletas_cuenta()
    if not loaded:
        return []
    sitios, pdf_cache, extraer_texto_pdf, lecturas_medidor = loaded
    filas: list[dict] = []
    for sitio in sitios:
        node = sitio.node_id
        nombre = NOMBRE_NODO.get(node, sitio.node_name)
        horas = horas_por_nodo.get(node) or {}
        pendiente = node in PENDIENTES or not horas
        for f in sorted(sitio.filas, key=lambda x: (x.emision, x.lectura_actual)):
            d0, d1 = f.lectura_anterior.date(), f.lectura_actual.date()
            pdf = pdf_cache / f.carpeta / f.pdf_name
            ant = act = dif_l = None
            if pdf.is_file():
                try:
                    ant, act, dif_l = lecturas_medidor(extraer_texto_pdf(pdf))
                except Exception:
                    pass
            dif_cta = float(dif_l) if dif_l is not None else float(f.m3_cuenta)
            placa, n_h, esp = _placa_en_periodo(horas, d0, d1, 12)
            cob_pct = (100.0 * n_h / esp) if esp else 0.0
            err = _error_pct(placa, dif_cta) if placa is not None else None
            notas: list[str] = []
            if f.estimado:
                notas.append("Cobro a promedio: no cruzar con la cuenta.")
            if pendiente:
                notas.append("Placa pendiente de descargar.")
            elif placa is None:
                notas.append("Sin horas de placa en el período.")
            elif cob_pct < 80:
                notas.append(f"Placa incompleta ({cob_pct:.0f} % de horas).")
            mes = _mes_label_cta(f.emision.year, f.emision.month)
            filas.append(
                {
                    "nombre": nombre,
                    "node": node,
                    "mes": mes,
                    "ym": (f.emision.year, f.emision.month),
                    "d0": d0,
                    "d1": d1,
                    "ini": _fmt_lectura(d0, ant),
                    "fin": _fmt_lectura(d1, act),
                    "dif_cta": dif_cta,
                    "placa": placa,
                    "err": err,
                    "estimado": bool(f.estimado),
                    "pendiente": pendiente,
                    "n_h": n_h,
                    "esp": esp,
                    "cob_pct": cob_pct,
                    "nota": " ".join(notas),
                    "usable": (
                        not f.estimado
                        and placa is not None
                        and cob_pct >= 50
                    ),
                }
            )
            print(
                f"  CTA {nombre:24} {mes:10} cta={dif_cta:.1f} "
                f"placa={placa if placa is not None else '—'} "
                f"cob={cob_pct:.0f}% est={f.estimado}",
                flush=True,
            )
    filas.sort(key=lambda r: (r["node"], r["d1"], r["d0"]))
    return filas


def _escribir_detalle_placa_vs_cuenta(wb: Workbook, filas: list[dict]) -> None:
    name = "Detalle placa vs cuenta"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name, 2)
    headers = [
        "Colegio",
        "Nodo",
        "Mes",
        "Lectura inicial (12:00)",
        "Lectura final (12:00)",
        "m³ cuenta",
        "m³ placa",
        "Dif (placa − cuenta)",
        "Error placa (%)",
        "Promedio",
        "Cobertura placa",
        "Nota",
    ]
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        c = ws.cell(1, col)
        c.fill = AZUL_HDR
        c.font = FONT_BLANCO
        c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        c.border = THIN
    ws.row_dimensions[1].height = 36
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    ws["A2"] = (
        "Un renglón por boleta Aguas Andinas. Corte 12:00: el día de lectura inicial "
        "cuenta desde las 12:00; el día de lectura final, hasta las 12:00. "
        "m³ placa = horas únicas de data de placas (si un horario se repetía, queda una sola vez). "
        "Azul = la placa marca más que la cuenta. Rojo = la cuenta marca más que la placa. "
        "Verde = cobro a promedio (no cruzar). Naranja = placa ausente o incompleta."
    )
    ws["A2"].font = Font(bold=True, size=11, color="003366")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 48

    if not filas:
        ws["A3"] = "No se pudieron leer las boletas de Drive."
        return

    prev_node = None
    tot_cta = tot_placa = 0.0
    n_placa = 0

    def _fila_total(nombre: str, cta: float, pla: float, n: int) -> None:
        err = _error_pct(pla, cta) if n else None
        ws.append(
            [
                f"TOTAL {nombre}",
                None,
                None,
                None,
                None,
                cta,
                pla if n else None,
                (pla - cta) if n else None,
                err,
                None,
                None,
                "Suma de períodos con placa (≥50 % de horas), sin promedios.",
            ]
        )
        r = ws.max_row
        for col in range(1, len(headers) + 1):
            cell = ws.cell(r, col)
            cell.border = THIN
            cell.alignment = CENTER
            cell.fill = PatternFill("solid", fgColor="D9E1F2")
            cell.font = Font(bold=True, size=9)
        ws.cell(r, 1).alignment = Alignment(horizontal="left")
        ws.cell(r, 6).number_format = NUM_FMT
        ws.cell(r, 7).number_format = NUM_FMT
        ws.cell(r, 7).fill = CELESTE
        if n:
            ws.cell(r, 8).number_format = NUM_FMT
            _pintar_dif(ws.cell(r, 8), pla - cta)
            ws.cell(r, 9).number_format = PCT_FMT
            ws.cell(r, 9).value = None if err is None else round(err, 2)

    for info in filas:
        if prev_node is not None and info["node"] != prev_node:
            nombre_prev = NOMBRE_NODO.get(prev_node, prev_node)
            _fila_total(nombre_prev, tot_cta, tot_placa, n_placa)
            tot_cta = tot_placa = 0.0
            n_placa = 0
        prev_node = info["node"]
        placa = info["placa"]
        dif = (placa - info["dif_cta"]) if placa is not None else None
        cob = (
            f"{info['cob_pct']:.0f}% ({info['n_h']}/{info['esp']} h)"
            if info["esp"]
            else "—"
        )
        ws.append(
            [
                info["nombre"],
                info["node"],
                info["mes"],
                info["ini"],
                info["fin"],
                info["dif_cta"],
                placa,
                dif,
                info["err"],
                "Sí" if info["estimado"] else "No",
                cob,
                info["nota"],
            ]
        )
        r = ws.max_row
        for col in range(1, len(headers) + 1):
            cell = ws.cell(r, col)
            cell.border = THIN
            cell.alignment = CENTER
        ws.cell(r, 1).alignment = Alignment(horizontal="left")
        ws.cell(r, 6).number_format = NUM_FMT
        if placa is not None:
            ws.cell(r, 7).number_format = NUM_FMT
            ws.cell(r, 7).fill = CELESTE
            ws.cell(r, 8).number_format = NUM_FMT
        if info["err"] is not None:
            ws.cell(r, 9).number_format = PCT_FMT
            ws.cell(r, 9).value = round(info["err"], 2)
        if info["estimado"]:
            for col in range(1, len(headers) + 1):
                ws.cell(r, col).fill = VERDE_PROM
        elif info["pendiente"] or placa is None or info["cob_pct"] < 80:
            ws.cell(r, 7).fill = NARANJA
            ws.cell(r, 11).fill = NARANJA
        elif dif is not None:
            _pintar_dif(ws.cell(r, 8), dif)
        if info["usable"]:
            tot_cta += info["dif_cta"]
            tot_placa += placa
            n_placa += 1

    if prev_node is not None:
        _fila_total(NOMBRE_NODO.get(prev_node, prev_node), tot_cta, tot_placa, n_placa)

    anchos = [26, 12, 12, 36, 36, 14, 14, 18, 14, 12, 18, 52]
    for i, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.sheet_view.showGridLines = False


def _escribir_matriz_placa_vs_cuenta(wb: Workbook, filas: list[dict]) -> None:
    name = "Placa vs cuenta"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name, 1)
    meses = sorted({r["ym"] for r in filas}) if filas else []
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5 + max(1, len(meses) * 4))
    ws["A1"] = (
        "Placa vs cuenta por mes de cada colegio. Mes = emisión de la boleta. "
        "Corte Aguas Andinas 12:00. m³ placa = horas únicas (sin duplicar). "
        "Azul = placa > cuenta. Rojo = cuenta > placa. Verde = promedio. "
        "Naranja = placa pendiente o incompleta. Detalle de lecturas en la hoja "
        "«Detalle placa vs cuenta»."
    )
    ws["A1"].font = Font(bold=True, size=11, color="003366")
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 40

    ws.merge_cells(start_row=2, start_column=1, end_row=3, end_column=1)
    ws.cell(2, 1, "Colegio")
    ws.merge_cells(start_row=2, start_column=2, end_row=3, end_column=2)
    ws.cell(2, 2, "Nodo")
    ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=5)
    ws.cell(2, 3, "TOTAL (sin promedios)")
    ws.cell(3, 3, "m³ cuenta")
    ws.cell(3, 4, "m³ placa")
    ws.cell(3, 5, "% error")
    col = 6
    for y, m in meses:
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 3)
        ws.cell(2, col, _mes_label_cta(y, m).upper())
        ws.cell(3, col, "m³ cuenta")
        ws.cell(3, col + 1, "m³ placa")
        ws.cell(3, col + 2, "Dif")
        ws.cell(3, col + 3, "% error")
        col += 4
    last_col = 5 + len(meses) * 4
    if last_col < 5:
        last_col = 5
    for r in (2, 3):
        for c in range(1, last_col + 1):
            cell = ws.cell(r, c)
            cell.fill = AZUL_HDR
            cell.font = FONT_BLANCO
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 20

    por_nodo: dict[str, list[dict]] = {}
    for r in filas:
        por_nodo.setdefault(r["node"], []).append(r)

    row_i = 4
    tot_mes = {ym: {"cta": 0.0, "pla": 0.0, "n": 0} for ym in meses}
    tot_c = tot_p = 0.0
    for _src, nombre, node, _dest in COLEGIOS:
        rows = por_nodo.get(node) or []
        by_mes: dict[tuple[int, int], list[dict]] = {}
        for r in rows:
            by_mes.setdefault(r["ym"], []).append(r)
        ws.cell(row_i, 1, nombre)
        ws.cell(row_i, 2, node)
        sc = sp = 0.0
        n_ok = 0
        col = 6
        for ym in meses:
            grupo = by_mes.get(ym) or []
            cta = sum(g["dif_cta"] for g in grupo) if grupo else None
            piezas = [g for g in grupo if g["placa"] is not None]
            pla = sum(g["placa"] for g in piezas) if piezas else None
            est = any(g["estimado"] for g in grupo)
            pend = (not grupo) or all(g["pendiente"] or g["placa"] is None for g in grupo)
            incompleto = any(g["placa"] is not None and g["cob_pct"] < 80 for g in grupo)
            c1 = ws.cell(row_i, col, cta)
            c2 = ws.cell(row_i, col + 1, pla)
            c3 = ws.cell(row_i, col + 2)
            c4 = ws.cell(row_i, col + 3)
            if cta is not None:
                c1.number_format = NUM_FMT
            if pla is not None:
                c2.number_format = NUM_FMT
                c2.fill = CELESTE
                if cta is not None:
                    delta = pla - cta
                    c3.value = delta
                    c3.number_format = NUM_FMT
                    err = _error_pct(pla, cta)
                    c4.value = None if err is None else round(err, 1)
                    c4.number_format = "0.0"
            if est:
                for cc in (col, col + 1, col + 2, col + 3):
                    ws.cell(row_i, cc).fill = VERDE_PROM
            elif pend:
                c2.fill = NARANJA
            elif incompleto:
                c2.fill = NARANJA
            elif pla is not None and cta is not None:
                _pintar_dif(c3, pla - cta)
            usable = [g for g in grupo if g["usable"]]
            if usable:
                sc += sum(g["dif_cta"] for g in usable)
                sp += sum(g["placa"] for g in usable)
                n_ok += len(usable)
                tot_mes[ym]["cta"] += sum(g["dif_cta"] for g in usable)
                tot_mes[ym]["pla"] += sum(g["placa"] for g in usable)
                tot_mes[ym]["n"] += len(usable)
            col += 4
        ws.cell(row_i, 3, sc if n_ok else None)
        ws.cell(row_i, 4, sp if n_ok else None)
        if n_ok:
            ws.cell(row_i, 3).number_format = NUM_FMT
            ws.cell(row_i, 4).number_format = NUM_FMT
            ws.cell(row_i, 3).fill = DORADO
            ws.cell(row_i, 4).fill = CELESTE
            err_t = _error_pct(sp, sc)
            ws.cell(row_i, 5, None if err_t is None else round(err_t, 1))
            ws.cell(row_i, 5).number_format = "0.0"
            if err_t is not None:
                _pintar_dif(ws.cell(row_i, 5), sp - sc)
            tot_c += sc
            tot_p += sp
        elif node in PENDIENTES:
            ws.cell(row_i, 4).fill = NARANJA
        for c in range(1, last_col + 1):
            ws.cell(row_i, c).border = THIN
            ws.cell(row_i, c).alignment = CENTER
        ws.cell(row_i, 1).alignment = Alignment(horizontal="left")
        row_i += 1

    ws.cell(row_i, 1, "TOTAL colegios")
    ws.cell(row_i, 2, "000008")
    ws.cell(row_i, 3, tot_c if tot_c else None)
    ws.cell(row_i, 4, tot_p if tot_p else None)
    gris = PatternFill("solid", fgColor="D9E1F2")
    for c in range(1, last_col + 1):
        cell = ws.cell(row_i, c)
        cell.border = THIN
        cell.alignment = CENTER
        cell.font = Font(bold=True, size=9)
        cell.fill = gris
    ws.cell(row_i, 1).alignment = Alignment(horizontal="left")
    if tot_c:
        ws.cell(row_i, 3).number_format = NUM_FMT
        ws.cell(row_i, 4).number_format = NUM_FMT
        ws.cell(row_i, 4).fill = CELESTE
        err_all = _error_pct(tot_p, tot_c)
        ws.cell(row_i, 5, None if err_all is None else round(err_all, 1))
        ws.cell(row_i, 5).number_format = "0.0"
        if err_all is not None:
            _pintar_dif(ws.cell(row_i, 5), tot_p - tot_c)
    col = 6
    for ym in meses:
        tc, tp, n = tot_mes[ym]["cta"], tot_mes[ym]["pla"], tot_mes[ym]["n"]
        ws.cell(row_i, col, tc if n else None)
        ws.cell(row_i, col + 1, tp if n else None)
        if n:
            ws.cell(row_i, col).number_format = NUM_FMT
            ws.cell(row_i, col + 1).number_format = NUM_FMT
            ws.cell(row_i, col + 1).fill = CELESTE
            delta = tp - tc
            ws.cell(row_i, col + 2, delta).number_format = NUM_FMT
            _pintar_dif(ws.cell(row_i, col + 2), delta)
            err = _error_pct(tp, tc)
            ws.cell(row_i, col + 3, None if err is None else round(err, 1))
            ws.cell(row_i, col + 3).number_format = "0.0"
        col += 4

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 12
    for i in range(3, last_col + 1):
        ws.column_dimensions[get_column_letter(i)].width = 12
    ws.freeze_panes = "C4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.sheet_view.showGridLines = False


def _escribir_comparativo_placa_cuenta(
    wb: Workbook,
    horas_por_nodo: dict[str, dict[date, dict[int, float]]],
) -> None:
    filas = _filas_placa_vs_cuenta(horas_por_nodo)
    _escribir_matriz_placa_vs_cuenta(wb, filas)
    _escribir_detalle_placa_vs_cuenta(wb, filas)
    print(f"[INFO] Placa vs cuenta: {len(filas)} períodos", flush=True)


def main() -> None:
    path = descargar_data_de_placas()
    src = load_workbook(path, data_only=False)
    print("[INFO] hojas fuente", src.sheetnames, flush=True)

    wb = Workbook()
    tmp = wb.active
    tmp.title = "_tmp"
    filas_resumen: list[dict] = []
    horas_por_nodo: dict[str, dict[date, dict[int, float]]] = {}

    for src_name, nombre, node, dest in COLEGIOS:
        if src_name not in src.sheetnames:
            print(f"[WARN] no está la hoja {src_name!r} en data de placas", flush=True)
            horas_por_nodo[node] = {}
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
                    "dias_app": 0,
                    "m3_app": None,
                    "pendiente": True,
                }
            )
            continue
        horas, st, dups_pares = _parse_fhm_sheet(src[src_name])
        horas_por_nodo[node] = horas
        data_app = _cargar_data_app(node)
        meta_app = _meta_data_app(node)
        pendiente = node in PENDIENTES or not horas
        print(
            f"  {dest:18} {nombre:24} {st['triples']:5} triples "
            f"únicos={st['unicos']} dups={st['duplicados']} "
            f"días_dup={st['dias_dup']} conf={st['conflictos']} "
            f"días={st['dias']} m³={st['m3']} app={len(data_app)}"
            f"{' PENDIENTE' if pendiente else ''}",
            flush=True,
        )
        d0_res = st["d0"]
        d1_res = st["d1"]
        if data_app:
            d0_app, d1_app = min(data_app), max(data_app)
            d0_res = min(d0_res, d0_app) if d0_res else d0_app
            d1_res = max(d1_res, d1_app) if d1_res else d1_app
        filas_resumen.append(
            {
                "nombre": nombre,
                "node": node,
                "d0": d0_res,
                "d1": d1_res,
                "dias": st["dias"],
                "unicos": st["unicos"],
                "duplicados": st["duplicados"],
                "m3": st["m3"],
                "dias_app": len(data_app),
                "m3_app": round(sum(data_app.values()), 2) if data_app else None,
                "pendiente": pendiente,
            }
        )
        if pendiente and not data_app:
            _hoja_pendiente(wb, dest, nombre, node)
            continue
        dias = sorted(horas)
        if data_app:
            dias = sorted(set(dias) | set(data_app))
        nota = (
            "Fuente placa: data de placas.xlsx (pegado F/H/M). "
            "Fila data placa = celeste (suma de horas de la placa; "
            "si un horario se repetía, queda una sola vez). "
            "Fila data app = verde (totales diarios de la app). "
            "Encabezado y hora en rojo = ese día/horario se repetía en la placa. "
            "Fila data app en rojo = no coincide con data placa. "
            f"Duplicados placa={st['duplicados']} "
            f"(días={st['dias_dup']}, conflictos de valor={st['conflictos']}). "
        )
        if pendiente:
            nota += "Placa pendiente de descargar; esta hoja muestra solo data app. "
        nota += "Solo días con placa" + (" o con data app." if data_app else ".")
        if data_app:
            extra_app = meta_app.get("nota") or (
                f"Dump app {min(data_app).strftime('%d/%m/%Y')}–"
                f"{max(data_app).strftime('%d/%m/%Y')} ({len(data_app)} días)."
            )
            nota += " " + extra_app
        construir_horario(
            dias,
            horas,
            wb=wb,
            sheet_name=dest,
            nombre=nombre,
            node=node,
            d0=dias[0],
            d1=dias[-1],
            listado=data_app,
            nota=nota,
            resaltar=dups_pares or None,
            fill_resaltar=ROJO,
            etiqueta_total="data placa",
            etiqueta_listado="data app",
            mostrar_fila_listado=True,
        )

    if "_tmp" in wb.sheetnames:
        del wb["_tmp"]
    _escribir_resumen(wb, filas_resumen)
    _escribir_comparativo_placa_cuenta(wb, horas_por_nodo)

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
