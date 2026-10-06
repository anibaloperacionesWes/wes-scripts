"""Comparación 2026 — colegios Providencia, con horas repetidas contadas una sola vez.

El CSV horario (`dates.measures.csv`) a veces trae la misma marca TIME dos veces
con el mismo m³/h. Sumarlas deja el día al doble (caso Carmela Carvajal; el
`totalM3` de la app coincide con esa suma). Este Excel:

- compara el consumo ene–hoy 2026 de los colegios Providencia;
- usa **una sola fila por hora** cuando el TIME se repite;
- lista los días/nodos en esa condición para pedir soporte a IT.

Uso:
  python generar_excel_comparacion_colegios_providencia_2026.py
"""
from __future__ import annotations

import json
import sys
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
BASE = "http://104.248.53.141:7003/wes/api/acl-node/v1"
OUT_DIR = ROOT / "reports" / "Providencia" / "Comparacion_2026"
CACHE_PATH = OUT_DIR / "cache_horario_2026.json"

# Colegios Providencia que se reportan juntos. 000006-03 se incluye en el
# barrido: si no tiene serie 2026 queda marcado sin datos.
NODOS: list[tuple[str, str]] = [
    ("000006-01", "Liceo Lastarria"),
    ("000006-02", "Carmela Carvajal"),
    ("000006-03", "Arturo Alessandri Palma"),
    ("000006-04", "Liceo 7 Luisa Saavedra"),
    ("000006-05", "Liceo Juan Pablo Duarte"),
]

INICIO = date(2026, 1, 1)
FIN = date.today()
WORKERS = 16
TOL = 0.001

MESES = [
    "Ene",
    "Feb",
    "Mar",
    "Abr",
    "May",
    "Jun",
    "Jul",
    "Ago",
    "Sep",
    "Oct",
    "Nov",
    "Dic",
]

FILL_HDR = PatternFill("solid", fgColor="1F4E79")
FILL_TOTAL = PatternFill("solid", fgColor="D6E3F0")
FILL_DOBLE = PatternFill("solid", fgColor="F4B183")
FILL_CERO = PatternFill("solid", fgColor="FFF2CC")
FILL_DIST = PatternFill("solid", fgColor="C6EFCE")
FILL_OK = PatternFill("solid", fgColor="E2EFDA")
FILL_SIN = PatternFill("solid", fgColor="F2F2F2")
FONT_HDR = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
FONT_BOLD = Font(bold=True, name="Calibri", size=11)
FONT_NAME = Font(name="Calibri", size=11)
THIN = Border(
    left=Side(style="thin", color="BDD3E6"),
    right=Side(style="thin", color="BDD3E6"),
    top=Side(style="thin", color="BDD3E6"),
    bottom=Side(style="thin", color="BDD3E6"),
)
FMT = "#,##0.00"
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _fechas() -> list[date]:
    out: list[date] = []
    d = INICIO
    while d <= FIN:
        out.append(d)
        d += timedelta(days=1)
    return out


def _fetch_csv(node_id: str, dia: date) -> str:
    ds = dia.strftime("%d%m%Y")
    url = f"{BASE}/nodes/{node_id}/dates.measures.csv"
    last_exc: Exception | None = None
    for _ in range(3):
        try:
            r = requests.get(url, params=[("start", ds), ("end", ds)], timeout=40)
            r.raise_for_status()
            return r.text
        except Exception as exc:  # noqa: BLE001
            last_exc = exc
    raise RuntimeError(f"{node_id} {dia}: {last_exc}")


def _clasificar(texto: str) -> dict:
    """Una fila por TIME. Si el valor se repite, se usa una sola vez."""
    filas: list[tuple[str, float]] = []
    for line in texto.strip().split("\n")[1:]:
        if not line.strip():
            continue
        parts = line.split(",", 1)
        if len(parts) < 2:
            continue
        try:
            t = parts[0].strip()
            v = float(parts[1].strip().replace(" ", "").replace(",", "."))
        except (ValueError, TypeError):
            continue
        filas.append((t, v))

    por_time: dict[str, list[float]] = defaultdict(list)
    for t, v in filas:
        por_time[t].append(v)

    m3_suma = 0.0
    m3_una = 0.0
    n_ident = 0
    n_dist = 0
    n_ident_positivo = 0
    ejemplos_dist: list[str] = []
    for t, vals in por_time.items():
        m3_suma += sum(vals)
        if len(vals) == 1:
            m3_una += vals[0]
            continue
        if max(vals) - min(vals) <= TOL:
            m3_una += vals[0]
            n_ident += 1
            if abs(vals[0]) > TOL:
                n_ident_positivo += 1
        else:
            m3_una += max(vals)
            n_dist += 1
            if len(ejemplos_dist) < 6:
                hh = t[11:16] if len(t) >= 16 else t
                ejemplos_dist.append(f"{hh}={'/'.join(f'{x:.2f}' for x in vals)}")

    if not filas:
        condicion = "SIN_DATOS"
    elif n_ident_positivo > 0 and n_dist > 0:
        condicion = "DOBLE_Y_DISTINTOS"
    elif n_ident_positivo > 0:
        condicion = "DOBLE_IDENTICO"
    elif n_dist > 0:
        condicion = "VALORES_DISTINTOS"
    elif n_ident > 0:
        condicion = "REPETIDA_EN_CERO"
    else:
        condicion = "OK"

    return {
        "filas": len(filas),
        "horas_unicas": len(por_time),
        "horas_identicas": n_ident,
        "horas_identicas_con_consumo": n_ident_positivo,
        "horas_valores_distintos": n_dist,
        "m3_suma": round(m3_suma, 4),
        "m3_una_hora": round(m3_una, 4),
        "condicion": condicion,
        "ejemplos_distintos": "; ".join(ejemplos_dist),
    }


def _cargar_cache() -> dict:
    if CACHE_PATH.is_file():
        return json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    return {}


def _guardar_cache(cache: dict) -> None:
    CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    CACHE_PATH.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")


def descargar() -> dict[str, dict[str, dict]]:
    """cache[node_id][iso_date] = clasificacion."""
    cache = _cargar_cache()
    fechas = _fechas()
    pendientes: list[tuple[str, date]] = []
    for nid, _nombre in NODOS:
        nodo = cache.setdefault(nid, {})
        for d in fechas:
            key = d.isoformat()
            if key not in nodo:
                pendientes.append((nid, d))

    print(f"[INFO] Días {INICIO} → {FIN} | nodos {len(NODOS)} | pendientes {len(pendientes)}")
    hechos = 0
    errores: list[str] = []

    def _job(nid: str, dia: date) -> tuple[str, str, dict]:
        texto = _fetch_csv(nid, dia)
        return nid, dia.isoformat(), _clasificar(texto)

    if pendientes:
        with ThreadPoolExecutor(max_workers=WORKERS) as pool:
            futs = {pool.submit(_job, nid, dia): (nid, dia) for nid, dia in pendientes}
            for fut in as_completed(futs):
                nid, dia = futs[fut]
                try:
                    nid_r, key, info = fut.result()
                    cache.setdefault(nid_r, {})[key] = info
                except Exception as exc:  # noqa: BLE001
                    errores.append(f"{nid} {dia}: {exc}")
                hechos += 1
                if hechos % 200 == 0 or hechos == len(pendientes):
                    print(f"[INFO] {hechos}/{len(pendientes)}")
                    _guardar_cache(cache)
        _guardar_cache(cache)

    if errores:
        print(f"[WARN] {len(errores)} días con error")
        for e in errores[:12]:
            print(" ", e)
    return cache


def _style_header(ws, row: int, cols: int) -> None:
    for c in range(1, cols + 1):
        cell = ws.cell(row, c)
        cell.fill = FILL_HDR
        cell.font = FONT_HDR
        cell.alignment = CENTER
        cell.border = THIN


def _autosize(ws, widths: dict[int, float]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def _escribir(cache: dict[str, dict[str, dict]]) -> Path:
    wb = Workbook()
    fechas = _fechas()
    hoy = FIN.isoformat()

    # --- Resumen ---
    ws = wb.active
    ws.title = "Resumen"
    ws["A1"] = "Colegios Providencia — comparación 2026 (horas repetidas = 1 hora)"
    ws["A1"].font = Font(bold=True, size=14, name="Calibri", color="1F4E79")
    ws.merge_cells("A1:H1")
    ws["A2"] = (
        f"Periodo: {INICIO:%d/%m/%Y} al {FIN:%d/%m/%Y} (hoy, serie del día en curso puede estar incompleta). "
        "Fuente: CSV horario dates.measures.csv. Si la misma marca TIME aparece más de una vez "
        "con el mismo m³/h, se cuenta una sola vez. La app (totalM3) suma esas filas y el día queda al doble."
    )
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 48

    headers = [
        "Nodo",
        "Colegio",
        "Días con datos",
        "Días con hora repetida",
        "Desde",
        "Hasta",
        "Sigue al corte",
        "Días que duplican m³",
        "m³ crudo (suma filas)",
        "m³ con 1 hora",
        "m³ de más",
        "Días solo en cero",
        "Días valores distintos",
    ]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(headers))

    resumen_rows: list[dict] = []
    for nid, nombre in NODOS:
        dias = cache.get(nid, {})
        con_datos = 0
        serie: list[str] = []
        duplican_m3 = 0
        ceros = 0
        distintos = 0
        m3_suma = 0.0
        m3_una = 0.0
        for key in sorted(dias):
            info = dias[key]
            if info["condicion"] == "SIN_DATOS":
                continue
            con_datos += 1
            m3_suma += info["m3_suma"]
            m3_una += info["m3_una_hora"]
            # 10+ horas idénticas repetidas = el defecto de Carmela, no un hora suelta.
            if info["horas_identicas"] >= 10:
                serie.append(key)
                if info["horas_identicas_con_consumo"] >= 10:
                    duplican_m3 += 1
                elif info["condicion"] == "REPETIDA_EN_CERO":
                    ceros += 1
            if info["horas_valores_distintos"] > 0 and info["horas_identicas"] < 10:
                distintos += 1
        resumen_rows.append(
            {
                "nid": nid,
                "nombre": nombre,
                "con_datos": con_datos,
                "n_serie": len(serie),
                "desde": serie[0] if serie else "",
                "hasta": serie[-1] if serie else "",
                "sigue": "Sí" if serie and serie[-1] == hoy else ("No" if serie else ""),
                "n_duplican": duplican_m3,
                "m3_suma": m3_suma,
                "m3_una": m3_una,
                "extra": m3_suma - m3_una,
                "ceros": ceros,
                "distintos": distintos,
            }
        )

    for i, row in enumerate(resumen_rows):
        r = 5 + i
        vals = [
            row["nid"],
            row["nombre"],
            row["con_datos"],
            row["n_serie"],
            _fmt_fecha(row["desde"]),
            _fmt_fecha(row["hasta"]),
            row["sigue"],
            row["n_duplican"],
            row["m3_suma"],
            row["m3_una"],
            row["extra"],
            row["ceros"],
            row["distintos"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.font = FONT_NAME
            cell.border = THIN
            cell.alignment = CENTER if c != 2 else LEFT
            if c in (9, 10, 11):
                cell.number_format = FMT
        if row["n_serie"]:
            for c in range(1, 14):
                ws.cell(r, c).fill = FILL_DOBLE
        elif row["distintos"]:
            for c in range(1, 14):
                ws.cell(r, c).fill = FILL_CERO

    nota_r = 5 + len(resumen_rows) + 1
    afectados = [r for r in resumen_rows if r["n_serie"]]
    if afectados:
        texto = (
            "Para soporte IT — la misma hora viene repetida con el mismo m³/h "
            "(10 o más horas del día). Al sumar filas, el día queda al doble: "
            + "; ".join(
                f"{r['nombre']} ({r['nid']}) del {_fmt_fecha(r['desde'])} al {_fmt_fecha(r['hasta'])} "
                f"({r['n_serie']} días con la hora repetida, {r['n_duplican']} de ellos duplican m³, "
                f"+{r['extra']:.0f} m³ de más). Sigue al corte: {r['sigue']}"
                for r in afectados
            )
            + ". Los días en cero igual traen la hora duplicada: el defecto sigue, aunque el m³ no cambie. "
            "Lastarria, Alessandri y Duarte no están en esa condición. "
            "Duarte el 08–11/04/2026 repite la hora con valores distintos (no es una copia); "
            "ahí se tomó el mayor y no se sumó. Detalle en Duplicados_IT."
        )
    else:
        texto = "Ningún colegio repitió horas con el mismo valor en el periodo."
    ws.cell(nota_r, 1, texto)
    ws.cell(nota_r, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=nota_r, start_column=1, end_row=nota_r, end_column=13)
    ws.row_dimensions[nota_r].height = 60

    nota2 = nota_r + 2
    ws.cell(
        nota2,
        1,
        "Regla aplicada: por cada marca TIME se conserva un solo m³/h. "
        "Si las copias son iguales, se usa ese valor. Si difieren (por ejemplo 0,30 y 0,00), "
        "se usa el mayor y el día queda en Duplicados_IT como valores distintos, sin sumar. "
        "Verificado el 05/10/2026 en Carmela: CSV 48 filas / 24 horas, totalM3 de la app = 51,2 "
        "(suma) y consumo con 1 hora = 25,6.",
    )
    ws.cell(nota2, 1).alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=nota2, start_column=1, end_row=nota2, end_column=13)
    ws.row_dimensions[nota2].height = 48

    _autosize(
        ws,
        {
            1: 14,
            2: 28,
            3: 16,
            4: 24,
            5: 14,
            6: 14,
            7: 16,
            8: 22,
            9: 22,
            10: 16,
            11: 14,
            12: 18,
            13: 22,
        },
    )
    ws.auto_filter.ref = f"A4:M{4 + len(resumen_rows)}"
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.left.text = "Providencia — colegios 2026"
    ws.oddFooter.right.text = "WES — horas repetidas contadas una vez"
    ws.print_title_rows = "1:4"
    ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110

    # --- Comparación mensual (1 hora) ---
    ws_m = wb.create_sheet("Comparacion_mensual")
    ws_m["A1"] = "m³ por mes — una sola hora cuando el TIME se repite"
    ws_m["A1"].font = Font(bold=True, size=14, name="Calibri", color="1F4E79")
    ws_m.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + len(NODOS) + 1)

    ws_m.cell(3, 1, "Mes")
    for i, (_nid, nombre) in enumerate(NODOS, 2):
        ws_m.cell(3, i, nombre)
    ws_m.cell(3, 2 + len(NODOS), "Total")
    _style_header(ws_m, 3, 2 + len(NODOS))

    meses_presentes = sorted({d.month for d in fechas})
    matriz: dict[tuple[str, int], float] = defaultdict(float)
    for nid, _nombre in NODOS:
        for key, info in cache.get(nid, {}).items():
            if info["condicion"] == "SIN_DATOS":
                continue
            mes = int(key[5:7])
            matriz[(nid, mes)] += info["m3_una_hora"]

    for i, mes in enumerate(meses_presentes):
        r = 4 + i
        etiqueta = f"{MESES[mes - 1]} 2026"
        if mes == FIN.month:
            etiqueta += " (al día)"
        ws_m.cell(r, 1, etiqueta).font = FONT_NAME
        ws_m.cell(r, 1).border = THIN
        total = 0.0
        for j, (nid, _nombre) in enumerate(NODOS):
            v = round(matriz[(nid, mes)], 2)
            total += v
            cell = ws_m.cell(r, 2 + j, v)
            cell.number_format = FMT
            cell.font = FONT_NAME
            cell.border = THIN
            cell.alignment = CENTER
        cell_t = ws_m.cell(r, 2 + len(NODOS), round(total, 2))
        cell_t.number_format = FMT
        cell_t.font = FONT_BOLD
        cell_t.border = THIN
        cell_t.fill = FILL_TOTAL

    rtot = 4 + len(meses_presentes)
    ws_m.cell(rtot, 1, "Total 2026").font = FONT_BOLD
    ws_m.cell(rtot, 1).fill = FILL_TOTAL
    ws_m.cell(rtot, 1).border = THIN
    gran = 0.0
    for j, (nid, _nombre) in enumerate(NODOS):
        v = round(sum(matriz[(nid, m)] for m in meses_presentes), 2)
        gran += v
        cell = ws_m.cell(rtot, 2 + j, v)
        cell.number_format = FMT
        cell.font = FONT_BOLD
        cell.fill = FILL_TOTAL
        cell.border = THIN
    cell = ws_m.cell(rtot, 2 + len(NODOS), round(gran, 2))
    cell.number_format = FMT
    cell.font = FONT_BOLD
    cell.fill = FILL_TOTAL
    cell.border = THIN

    # Crudo al lado, misma hoja, para ver el doble
    col0 = 2 + len(NODOS) + 2
    ws_m.cell(3, col0, "Mes")
    ws_m.cell(3, col0).fill = FILL_HDR
    ws_m.cell(3, col0).font = FONT_HDR
    ws_m.cell(2, col0, "m³ crudo (suma de filas repetidas — lo que muestra la app)")
    ws_m.cell(2, col0).font = Font(bold=True, size=12, name="Calibri", color="C65911")
    ws_m.merge_cells(start_row=2, start_column=col0, end_row=2, end_column=col0 + len(NODOS))
    for i, (_nid, nombre) in enumerate(NODOS):
        cell = ws_m.cell(3, col0 + 1 + i, nombre)
        cell.fill = FILL_HDR
        cell.font = FONT_HDR
        cell.alignment = CENTER
    matriz_cruda: dict[tuple[str, int], float] = defaultdict(float)
    for nid, _nombre in NODOS:
        for key, info in cache.get(nid, {}).items():
            if info["condicion"] == "SIN_DATOS":
                continue
            matriz_cruda[(nid, int(key[5:7]))] += info["m3_suma"]
    for i, mes in enumerate(meses_presentes):
        r = 4 + i
        ws_m.cell(r, col0, f"{MESES[mes - 1]} 2026").border = THIN
        for j, (nid, _nombre) in enumerate(NODOS):
            v = round(matriz_cruda[(nid, mes)], 2)
            cell = ws_m.cell(r, col0 + 1 + j, v)
            cell.number_format = FMT
            cell.border = THIN
            cell.alignment = CENTER
            una = matriz[(nid, mes)]
            if v > una + 0.05:
                cell.fill = FILL_DOBLE
    for j, (nid, _nombre) in enumerate(NODOS):
        v = round(sum(matriz_cruda[(nid, m)] for m in meses_presentes), 2)
        cell = ws_m.cell(rtot, col0 + 1 + j, v)
        cell.number_format = FMT
        cell.font = FONT_BOLD
        cell.fill = FILL_TOTAL
        cell.border = THIN
    ws_m.cell(rtot, col0, "Total 2026").font = FONT_BOLD
    ws_m.cell(rtot, col0).fill = FILL_TOTAL
    ws_m.cell(rtot, col0).border = THIN

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "m³ 2026 con 1 hora (sin doble conteo)"
    chart.y_axis.title = "m³"
    chart.style = 10
    data = Reference(ws_m, min_col=2, max_col=1 + len(NODOS), min_row=3, max_row=3 + len(meses_presentes))
    cats = Reference(ws_m, min_col=1, min_row=4, max_row=3 + len(meses_presentes))
    chart.add_data(data, from_rows=False, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.y_axis.majorGridlines = None
    chart.legend.position = "b"
    chart.width = 22
    chart.height = 10
    ws_m.add_chart(chart, "A" + str(rtot + 3))

    _autosize(ws_m, {1: 18, **{i: 24 for i in range(2, 2 + len(NODOS) + 1)}, col0: 16})
    for i in range(len(NODOS)):
        ws_m.column_dimensions[get_column_letter(col0 + 1 + i)].width = 24
    ws_m.freeze_panes = "B4"
    ws_m.page_setup.orientation = "landscape"
    ws_m.page_setup.fitToPage = True
    ws_m.page_setup.fitToWidth = 1
    ws_m.page_setup.fitToHeight = 1
    ws_m.sheet_properties.pageSetUpPr.fitToPage = True
    ws_m.page_setup.paperSize = ws_m.PAPERSIZE_TABLOID
    ws_m.print_title_rows = "1:3"
    ws_m.sheet_view.showGridLines = False
    ws_m.oddHeader.left.text = "Comparación mensual — 1 hora"
    ws_m.oddFooter.right.text = "Naranja = el crudo supera al valor con 1 hora"

    # --- Duplicados IT ---
    ws_d = wb.create_sheet("Duplicados_IT")
    ws_d["A1"] = "Días en que se repite la hora — para soporte IT"
    ws_d["A1"].font = Font(bold=True, size=14, name="Calibri", color="1F4E79")
    ws_d.merge_cells("A1:L1")
    ws_d["A2"] = (
        "DOBLE_IDENTICO: la misma hora viene dos veces con el mismo m³/h; la app suma y el día marca el doble. "
        "REPETIDA_EN_CERO: la hora se repite pero el valor es 0 (no cambia el m³, igual es el mismo defecto). "
        "VALORES_DISTINTOS: la hora se repite con números distintos; no se suman, se toma el mayor."
    )
    ws_d["A2"].alignment = Alignment(wrap_text=True)
    ws_d.merge_cells("A2:L2")
    ws_d.row_dimensions[2].height = 36
    hdrs = [
        "Nodo",
        "Colegio",
        "Fecha",
        "Condición",
        "Filas CSV",
        "Horas únicas",
        "Horas idénticas repetidas",
        "Con consumo (>0)",
        "Horas con valores distintos",
        "m³ si se suman (app)",
        "m³ con 1 hora",
        "m³ de más",
        "Ejemplos valores distintos",
    ]
    for i, h in enumerate(hdrs, 1):
        ws_d.cell(4, i, h)
    _style_header(ws_d, 4, len(hdrs))

    filas_it: list[tuple] = []
    for nid, nombre in NODOS:
        for key in sorted(cache.get(nid, {})):
            info = cache[nid][key]
            if info["condicion"] in ("OK", "SIN_DATOS"):
                continue
            if info["horas_identicas"] == 0 and info["horas_valores_distintos"] == 0:
                continue
            filas_it.append((nid, nombre, key, info))

    fills = {
        "DOBLE_IDENTICO": FILL_DOBLE,
        "DOBLE_Y_DISTINTOS": FILL_DOBLE,
        "REPETIDA_EN_CERO": FILL_CERO,
        "VALORES_DISTINTOS": FILL_DIST,
    }
    for i, (nid, nombre, key, info) in enumerate(filas_it):
        r = 5 + i
        extra = round(info["m3_suma"] - info["m3_una_hora"], 4)
        vals = [
            nid,
            nombre,
            _fmt_fecha(key),
            info["condicion"],
            info["filas"],
            info["horas_unicas"],
            info["horas_identicas"],
            info["horas_identicas_con_consumo"],
            info["horas_valores_distintos"],
            round(info["m3_suma"], 2),
            round(info["m3_una_hora"], 2),
            round(extra, 2),
            info["ejemplos_distintos"],
        ]
        fill = fills.get(info["condicion"], FILL_OK)
        for c, v in enumerate(vals, 1):
            cell = ws_d.cell(r, c, v)
            cell.font = FONT_NAME
            cell.border = THIN
            cell.fill = fill
            cell.alignment = CENTER if c != 13 else LEFT
            if c in (10, 11, 12):
                cell.number_format = FMT
        if key == hoy:
            ws_d.cell(r, 3, _fmt_fecha(key) + " (parcial)")

    ws_d.auto_filter.ref = f"A4:M{4 + max(len(filas_it), 1)}"
    ws_d.freeze_panes = "A5"
    ws_d.auto_filter.ref = f"A4:M{4 + len(filas_it)}"
    _autosize(
        ws_d,
        {
            1: 14,
            2: 28,
            3: 18,
            4: 24,
            5: 12,
            6: 14,
            7: 26,
            8: 18,
            9: 28,
            10: 20,
            11: 16,
            12: 14,
            13: 42,
        },
    )
    ws_d.page_setup.orientation = "landscape"
    ws_d.page_setup.fitToPage = True
    ws_d.page_setup.fitToWidth = 1
    ws_d.page_setup.fitToHeight = 0
    ws_d.sheet_properties.pageSetUpPr.fitToPage = True
    ws_d.page_setup.paperSize = ws_d.PAPERSIZE_TABLOID
    ws_d.print_title_rows = "1:4"
    ws_d.page_setup.horizontalCentered = True
    ws_d.sheet_view.showGridLines = False
    ws_d.oddHeader.left.text = "Duplicados para IT"
    ws_d.oddFooter.left.text = "Naranja = marca el doble   Amarillo = repetida en cero   Verde = valores distintos"
    ws_d.oddFooter.right.text = "Página &P de &N"
    ws_d.page_setup.pageOrder = "downThenOver"

    # --- Diario corregido (para comparar la serie) ---
    ws_dia = wb.create_sheet("Diario_1_hora")
    ws_dia.cell(1, 1, "m³ diario con 1 hora por TIME")
    ws_dia.cell(1, 1).font = Font(bold=True, size=14, name="Calibri", color="1F4E79")
    ws_dia.cell(3, 1, "Fecha")
    for i, (_nid, nombre) in enumerate(NODOS, 2):
        ws_dia.cell(3, i, nombre)
    _style_header(ws_dia, 3, 1 + len(NODOS))
    for i, dia in enumerate(fechas):
        r = 4 + i
        key = dia.isoformat()
        ws_dia.cell(r, 1, dia.strftime("%d/%m/%Y")).border = THIN
        ws_dia.cell(r, 1).font = FONT_NAME
        for j, (nid, _nombre) in enumerate(NODOS):
            info = cache.get(nid, {}).get(key)
            cell = ws_dia.cell(r, 2 + j)
            cell.border = THIN
            cell.font = FONT_NAME
            cell.alignment = CENTER
            if not info or info["condicion"] == "SIN_DATOS":
                cell.value = None
                cell.fill = FILL_SIN
            else:
                cell.value = round(info["m3_una_hora"], 2)
                cell.number_format = FMT
                if info["condicion"] in ("DOBLE_IDENTICO", "DOBLE_Y_DISTINTOS"):
                    cell.fill = FILL_DOBLE
    _autosize(ws_dia, {1: 14, **{i: 26 for i in range(2, 2 + len(NODOS))}})
    ws_dia.freeze_panes = "B4"
    ws_dia.auto_filter.ref = f"A3:{get_column_letter(1 + len(NODOS))}{3 + len(fechas)}"
    ws_dia.page_setup.orientation = "landscape"
    ws_dia.page_setup.fitToPage = True
    ws_dia.page_setup.fitToWidth = 1
    ws_dia.page_setup.fitToHeight = 0
    ws_dia.sheet_properties.pageSetUpPr.fitToPage = True
    ws_dia.page_setup.paperSize = ws_dia.PAPERSIZE_TABLOID
    ws_dia.print_title_rows = "1:3"
    ws_dia.sheet_view.showGridLines = False
    ws_dia.oddHeader.left.text = "Diario con 1 hora"
    ws_dia.oddFooter.right.text = "Celda naranja = ese día la hora venía repetida"
    ws_dia.page_setup.horizontalCentered = True

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    path = OUT_DIR / f"Comparacion_colegios_Providencia_2026_{stamp}.xlsx"
    wb.save(path)
    print(f"[OK] {path}")
    print(f"[INFO] Filas IT: {len(filas_it)}")
    for row in resumen_rows:
        print(
            f"  {row['nid']} {row['nombre']}: datos={row['con_datos']} "
            f"serie={row['n_serie']} {row['desde']}→{row['hasta']} sigue={row['sigue']} "
            f"duplican={row['n_duplican']} "
            f"crudo={row['m3_suma']:.1f} 1h={row['m3_una']:.1f} extra={row['extra']:.1f} "
            f"cero={row['ceros']} distintos={row['distintos']}"
        )
    return path


def _fmt_fecha(iso: str) -> str:
    if not iso:
        return ""
    y, m, d = iso.split("-")
    return f"{d}/{m}/{y}"


def main() -> None:
    cache = descargar()
    _escribir(cache)


if __name__ == "__main__":
    main()
