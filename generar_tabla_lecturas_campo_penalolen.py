"""
Tabla de lecturas de contador (campo) vs app WES — colegios CORMUP / Peñalolén.

Consulta la API y arma Excel + Word con:
- Última medida recibida en la app (fecha/hora Chile + m³/h)
- Consumo acumulado del mes en la app
- Columnas vacías para anotar lectura de contador en terreno y fecha de visita

Cuando tengas lecturas de campo, podés actualizar el Excel o pasar un CSV:

  python generar_tabla_lecturas_campo_penalolen.py
  python generar_tabla_lecturas_campo_penalolen.py --lecturas-csv reports/CORMUP/Lecturas_Campo/lecturas_campo.csv

Formato CSV de lecturas:
  node_id,fecha_visita,lectura_contador,observaciones
  000008-04,25/09/2026 11:30,12345.6,Medidor patio

Con 2+ lecturas por nodo calcula Δ m³ terreno vs m³ app entre visitas.
"""

from __future__ import annotations

import argparse
import csv
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from generar_reporte_word import (
    NODE_NAMES,
    acl_node_base_url,
    add_formatted_heading,
    add_logo_to_header,
    estilizar_tabla_wes,
    format_number_chilean,
)

try:
    from zoneinfo import ZoneInfo

    CHILE_TZ = ZoneInfo("America/Santiago")
except Exception:
    CHILE_TZ = timezone(timedelta(hours=-4))

COMPANY_ID = "000008"
CORMUP_NODES = [f"000008-{i:02d}" for i in range(1, 15)]
ROOT = Path(__file__).resolve().parent
OUT_BASE = ROOT / "reports" / "CORMUP" / "Lecturas_Campo"


@dataclass
class NodoApp:
    node_id: str
    nombre: str
    ultima_utc: Optional[datetime]
    ultima_chile: str
    ultimo_m3h: Optional[float]
    m3_mes: float
    dias_con_dato: int
    ultimo_dia_total: Optional[str]
    ultimo_dia_m3: Optional[float]
    alerta: str = ""


@dataclass
class LecturaCampo:
    node_id: str
    fecha_visita: datetime
    lectura_contador: float
    observaciones: str = ""


@dataclass
class Comparacion:
    node_id: str
    nombre: str
    lectura_anterior: LecturaCampo
    lectura_actual: LecturaCampo
    delta_terreno: float
    m3_app: Optional[float]
    diferencia: Optional[float]
    pct: Optional[float]
    nota: str = ""


def _parse_time_utc(s: str) -> Optional[datetime]:
    try:
        return datetime.fromisoformat(s.strip().replace("Z", "+00:00"))
    except (ValueError, TypeError):
        return None


def _fmt_chile(dt: Optional[datetime]) -> str:
    if dt is None:
        return "—"
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(CHILE_TZ).strftime("%d/%m/%Y %H:%M")


def _parse_fecha_flexible(s: str) -> Optional[datetime]:
    s = (s or "").strip()
    if not s:
        return None
    for fmt in (
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d",
    ):
        try:
            dt = datetime.strptime(s, fmt)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=CHILE_TZ)
            return dt
        except ValueError:
            continue
    return None


def _ultima_medida_csv(node_id: str, dias: int = 45) -> Tuple[Optional[datetime], Optional[float]]:
    """Última marca TIME y VALUE (m³/h) del CSV horario."""
    hoy = datetime.now(timezone.utc).date()
    url = f"{acl_node_base_url()}/nodes/{node_id}/dates.measures.csv"
    ultima: Optional[datetime] = None
    ultimo_val: Optional[float] = None
    for d in range(dias):
        dia = hoy - timedelta(days=d)
        ds = dia.strftime("%d%m%Y")
        try:
            r = requests.get(url, params=[("start", ds), ("end", ds)], timeout=30)
            if r.status_code != 200:
                continue
            by_t: Dict[datetime, float] = {}
            for line in r.text.strip().split("\n")[1:]:
                if not line.strip() or "," not in line:
                    continue
                t, v = line.split(",", 1)
                dt = _parse_time_utc(t)
                if not dt:
                    continue
                try:
                    val = float(v.strip())
                except ValueError:
                    continue
                by_t[dt] = val
            if not by_t:
                continue
            max_dt = max(by_t)
            if ultima is None or max_dt > ultima:
                ultima = max_dt
                ultimo_val = by_t[max_dt]
            if ultima and (hoy - ultima.date()).days <= d:
                break
        except requests.RequestException:
            continue
    return ultima, ultimo_val


def _consumo_periodo_diario(
    node_id: str, start: date, end: date
) -> Tuple[float, int, Optional[str], Optional[float]]:
    """Suma totalM3 diario entre start y end (inclusive)."""
    url = f"{acl_node_base_url()}/nodes/measures/dates"
    r = requests.get(
        url,
        params={
            "id": node_id,
            "start": start.strftime("%d%m%Y"),
            "end": end.strftime("%d%m%Y"),
        },
        timeout=45,
    )
    r.raise_for_status()
    data = r.json()
    months: List[dict] = []
    if isinstance(data, list):
        for item in data:
            months.extend(item.get("month") or [])
    elif isinstance(data, dict):
        months = data.get("month") or []
    total = 0.0
    dias = 0
    last_day: Optional[str] = None
    last_m3: Optional[float] = None
    for m in months:
        tm = m.get("totalM3")
        if tm is None:
            continue
        try:
            tm_f = float(tm)
        except (TypeError, ValueError):
            continue
        d = m.get("date")
        total += tm_f
        dias += 1
        if d and (last_day is None or str(d) > last_day):
            last_day = str(d)
            last_m3 = tm_f
    return total, dias, last_day, last_m3


def _m3_app_entre(node_id: str, t0: datetime, t1: datetime) -> float:
    """Suma totalM3 de días civiles entre t0 y t1 (excluye día final si es mismo día parcial → usa diario completo)."""
    d0 = t0.astimezone(CHILE_TZ).date()
    d1 = t1.astimezone(CHILE_TZ).date()
    if d1 < d0:
        return 0.0
    # Si es el mismo día, intentar sumar CSV horario entre marcas
    if d1 == d0:
        return _m3_csv_entre(node_id, t0, t1)
    # Días intermedios completos + fracciones
    m_ini = _m3_csv_entre(node_id, t0, datetime.combine(d0 + timedelta(days=1), datetime.min.time(), tzinfo=CHILE_TZ))
    m_fin = _m3_csv_entre(node_id, datetime.combine(d1, datetime.min.time(), tzinfo=CHILE_TZ), t1)
    mid = 0.0
    if (d1 - d0).days > 1:
        mid, _, _, _ = _consumo_periodo_diario(node_id, d0 + timedelta(days=1), d1 - timedelta(days=1))
    return m_ini + mid + m_fin


def _m3_csv_entre(node_id: str, t0: datetime, t1: datetime) -> float:
    if t0.tzinfo is None:
        t0 = t0.replace(tzinfo=CHILE_TZ)
    if t1.tzinfo is None:
        t1 = t1.replace(tzinfo=CHILE_TZ)
    t0u, t1u = t0.astimezone(timezone.utc), t1.astimezone(timezone.utc)
    if t1u <= t0u:
        return 0.0
    url = f"{acl_node_base_url()}/nodes/{node_id}/dates.measures.csv"
    total = 0.0
    d = t0u.date()
    end = t1u.date()
    while d <= end:
        ds = d.strftime("%d%m%Y")
        try:
            r = requests.get(url, params=[("start", ds), ("end", ds)], timeout=30)
            if r.status_code != 200:
                d += timedelta(days=1)
                continue
            by_t: Dict[datetime, float] = {}
            for line in r.text.strip().split("\n")[1:]:
                if not line.strip() or "," not in line:
                    continue
                t, v = line.split(",", 1)
                dt = _parse_time_utc(t)
                if not dt:
                    continue
                try:
                    val = float(v.strip())
                except ValueError:
                    continue
                by_t[dt] = by_t.get(dt, 0.0) + val
            for dt, val in by_t.items():
                if t0u <= dt < t1u:
                    total += val
        except requests.RequestException:
            pass
        d += timedelta(days=1)
    return total


def consultar_nodos(mes: int, anio: int) -> List[NodoApp]:
    inicio = date(anio, mes, 1)
    hoy = datetime.now(CHILE_TZ).date()
    if mes == hoy.month and anio == hoy.year:
        fin = hoy
    else:
        if mes == 12:
            fin = date(anio + 1, 1, 1) - timedelta(days=1)
        else:
            fin = date(anio, mes + 1, 1) - timedelta(days=1)

    def work(nid: str) -> NodoApp:
        nombre = NODE_NAMES.get(nid, nid)
        ultima, val = _ultima_medida_csv(nid)
        m3, dias, ud, um = _consumo_periodo_diario(nid, inicio, fin)
        alerta = ""
        if ultima is None:
            alerta = "Sin datos"
        else:
            horas = (datetime.now(timezone.utc) - ultima.astimezone(timezone.utc)).total_seconds() / 3600
            if horas > 48:
                alerta = f"Sin datos hace {horas/24:.0f} d"
            elif dias < (fin - inicio).days + 1 - 2:
                alerta = f"Solo {dias} días con dato"
        return NodoApp(
            node_id=nid,
            nombre=nombre,
            ultima_utc=ultima,
            ultima_chile=_fmt_chile(ultima),
            ultimo_m3h=val,
            m3_mes=m3,
            dias_con_dato=dias,
            ultimo_dia_total=ud,
            ultimo_dia_m3=um,
            alerta=alerta,
        )

    out: List[NodoApp] = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = {ex.submit(work, nid): nid for nid in CORMUP_NODES}
        for f in as_completed(futs):
            out.append(f.result())
    out.sort(key=lambda x: x.node_id)
    return out


def cargar_lecturas_csv(path: Path) -> Dict[str, List[LecturaCampo]]:
    by_node: Dict[str, List[LecturaCampo]] = {}
    if not path.is_file():
        return by_node
    with path.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            nid = (row.get("node_id") or row.get("nodo") or "").strip()
            if not nid:
                continue
            fv = _parse_fecha_flexible(row.get("fecha_visita") or row.get("fecha") or "")
            if not fv:
                continue
            try:
                lec = float(str(row.get("lectura_contador") or row.get("lectura") or "").replace(",", "."))
            except ValueError:
                continue
            obs = (row.get("observaciones") or "").strip()
            by_node.setdefault(nid, []).append(LecturaCampo(nid, fv, lec, obs))
    for nid in by_node:
        by_node[nid].sort(key=lambda x: x.fecha_visita)
    return by_node


def calcular_comparaciones(
    nodos: List[NodoApp], lecturas: Dict[str, List[LecturaCampo]]
) -> List[Comparacion]:
    comps: List[Comparacion] = []
    nombres = {n.node_id: n.nombre for n in nodos}
    for nid, lista in lecturas.items():
        if len(lista) < 2:
            continue
        for i in range(1, len(lista)):
            a, b = lista[i - 1], lista[i]
            delta = b.lectura_contador - a.lectura_contador
            nota = ""
            try:
                m3_app = _m3_app_entre(nid, a.fecha_visita, b.fecha_visita)
                dif = delta - m3_app
                pct = (dif / delta * 100.0) if abs(delta) > 1e-6 else None
            except Exception as e:
                m3_app = None
                dif = None
                pct = None
                nota = str(e)
            comps.append(
                Comparacion(
                    node_id=nid,
                    nombre=nombres.get(nid, nid),
                    lectura_anterior=a,
                    lectura_actual=b,
                    delta_terreno=delta,
                    m3_app=m3_app,
                    diferencia=dif,
                    pct=pct,
                    nota=nota,
                )
            )
    return comps


def _border() -> Border:
    thin = Side(style="thin", color="B0B0B0")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def escribir_excel(
    path: Path,
    nodos: List[NodoApp],
    mes_label: str,
    lecturas: Dict[str, List[LecturaCampo]],
    comps: List[Comparacion],
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Lecturas campo"

    headers = [
        "Nodo",
        "Colegio",
        "Última medida app (Chile)",
        "Último m³/h app",
        f"m³ app {mes_label}",
        "Días con dato",
        "Alerta",
        "Fecha visita",
        "Lectura contador (m³)",
        "Observaciones",
    ]
    header_fill = PatternFill("solid", fgColor="003366")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    alert_fill = PatternFill("solid", fgColor="FFE6E6")
    alt_fill = PatternFill("solid", fgColor="F2F6FC")

    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    # Última lectura de campo por nodo (si hay)
    ultima_campo: Dict[str, LecturaCampo] = {}
    for nid, lista in lecturas.items():
        if lista:
            ultima_campo[nid] = lista[-1]

    for i, n in enumerate(nodos, 2):
        lc = ultima_campo.get(n.node_id)
        row = [
            n.node_id,
            n.nombre,
            n.ultima_chile,
            n.ultimo_m3h if n.ultimo_m3h is not None else "",
            round(n.m3_mes, 2),
            n.dias_con_dato,
            n.alerta,
            lc.fecha_visita.strftime("%d/%m/%Y %H:%M") if lc else "",
            lc.lectura_contador if lc else "",
            lc.observaciones if lc else "",
        ]
        for col, val in enumerate(row, 1):
            cell = ws.cell(i, col, val)
            cell.border = _border()
            cell.alignment = Alignment(vertical="center")
            if n.alerta:
                cell.fill = alert_fill
            elif i % 2 == 0:
                cell.fill = alt_fill

    widths = [12, 26, 20, 14, 14, 12, 18, 18, 18, 28]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:J{len(nodos) + 1}"

    # Hoja plantilla para ir cargando lecturas
    ws2 = wb.create_sheet("Plantilla carga")
    ws2.append(["node_id", "fecha_visita", "lectura_contador", "observaciones"])
    for cell in ws2[1]:
        cell.fill = header_fill
        cell.font = header_font
    for n in nodos:
        ws2.append([n.node_id, "", "", n.nombre])
    for i, w in enumerate([12, 18, 18, 28], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # Instrucciones
    ws3 = wb.create_sheet("Cómo usar")
    ws3["A1"] = "Cómo comparar lecturas de contador vs app WES"
    ws3["A1"].font = Font(bold=True, size=14, color="003366")
    instrucciones = [
        "",
        "1. En cada colegio anotá la fecha/hora de la visita y el número del contador (lectura acumulada en m³).",
        "2. Completá la hoja «Plantilla carga» (o un CSV con las mismas columnas) y guardalo.",
        "3. Volvé a correr el script con --lecturas-csv para calcular Δ terreno vs m³ de la app entre visitas.",
        "4. La app WES no guarda el número absoluto del medidor: guarda consumo horario (m³/h).",
        "   Por eso hace falta al menos DOS lecturas de campo del mismo colegio para comparar.",
        "5. Diferencia = (lectura_actual − lectura_anterior) − m³ app entre esas dos fechas/horas.",
        "6. Filas en rojo = puntos con datos atrasados o pocos días en el mes (revisar en terreno).",
    ]
    for i, line in enumerate(instrucciones, 2):
        ws3[f"A{i}"] = line
    ws3.column_dimensions["A"].width = 110

    if comps:
        ws4 = wb.create_sheet("Comparación vs app")
        h2 = [
            "Nodo",
            "Colegio",
            "Visita anterior",
            "Lectura ant.",
            "Visita actual",
            "Lectura act.",
            "Δ terreno m³",
            "m³ app",
            "Diferencia",
            "% vs terreno",
            "Nota",
        ]
        for col, h in enumerate(h2, 1):
            cell = ws4.cell(1, col, h)
            cell.fill = header_fill
            cell.font = header_font
        for i, c in enumerate(comps, 2):
            row = [
                c.node_id,
                c.nombre,
                c.lectura_anterior.fecha_visita.strftime("%d/%m/%Y %H:%M"),
                c.lectura_anterior.lectura_contador,
                c.lectura_actual.fecha_visita.strftime("%d/%m/%Y %H:%M"),
                c.lectura_actual.lectura_contador,
                round(c.delta_terreno, 3),
                round(c.m3_app, 3) if c.m3_app is not None else "",
                round(c.diferencia, 3) if c.diferencia is not None else "",
                round(c.pct, 1) if c.pct is not None else "",
                c.nota,
            ]
            for col, val in enumerate(row, 1):
                cell = ws4.cell(i, col, val)
                cell.border = _border()
        for i, w in enumerate([12, 24, 16, 12, 16, 12, 12, 10, 12, 12, 24], 1):
            ws4.column_dimensions[get_column_letter(i)].width = w

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def escribir_word(
    path: Path,
    nodos: List[NodoApp],
    mes_label: str,
    comps: List[Comparacion],
) -> None:
    doc = Document()
    add_logo_to_header(doc)
    add_formatted_heading(doc, f"Lecturas de contador — Peñalolén (CORMUP) — {mes_label}", level=1)

    p = doc.add_paragraph()
    p.add_run(
        "Tabla de apoyo para visitas a colegios: última medida en la app WES y espacio "
        "para anotar la lectura del contador en terreno. Con dos o más lecturas por colegio "
        "se calcula el consumo de campo contra el de la aplicación."
    ).font.size = Pt(10)

    gen = datetime.now(CHILE_TZ).strftime("%d/%m/%Y %H:%M")
    meta = doc.add_paragraph()
    meta.add_run(f"Generado: {gen} (Chile) · Empresa 000008 · 14 nodos").font.size = Pt(9)
    for run in meta.runs:
        run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    headers = [
        "Nodo",
        "Colegio",
        "Última medida app",
        "m³/h",
        f"m³ {mes_label}",
        "Días",
        "Alerta",
        "Fecha visita",
        "Lectura contador",
    ]
    rows: List[List[str]] = [headers]
    highlight: List[int] = []
    for i, n in enumerate(nodos, 1):
        rows.append(
            [
                n.node_id,
                n.nombre,
                n.ultima_chile,
                format_number_chilean(n.ultimo_m3h, 2) if n.ultimo_m3h is not None else "—",
                format_number_chilean(n.m3_mes, 1),
                str(n.dias_con_dato),
                n.alerta or "—",
                "",
                "",
            ]
        )
        if n.alerta:
            highlight.append(i)

    total_m3 = sum(n.m3_mes for n in nodos)
    rows.append(["", "TOTAL", "", "", format_number_chilean(total_m3, 1), "", "", "", ""])

    table = doc.add_table(rows=len(rows), cols=len(headers))
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            table.rows[i].cells[j].text = val
    estilizar_tabla_wes(table, highlight_rows=highlight, has_total_row=True)

    doc.add_paragraph("")
    add_formatted_heading(doc, "Cómo usarla en las visitas", level=2)
    for line in (
        "Anotá en cada colegio: fecha/hora de la visita y el número del contador (m³ acumulados).",
        "La app no guarda el número absoluto del medidor; guarda consumo horario. Hace falta "
        "una segunda visita (o una lectura previa) para comparar Δ terreno vs m³ app.",
        "Pasame las lecturas (colegio + fecha/hora + número de contador) y te calculo la diferencia contra la aplicación.",
        "Filas en rojo: dato atrasado o pocos días con medida en el mes (prioridad de revisión).",
    ):
        doc.add_paragraph(line, style="List Bullet")

    if comps:
        add_formatted_heading(doc, "Comparación campo vs app (períodos con 2+ lecturas)", level=2)
        h2 = ["Colegio", "Período", "Δ terreno", "m³ app", "Dif.", "%"]
        r2: List[List[str]] = [h2]
        for c in comps:
            per = (
                f"{c.lectura_anterior.fecha_visita.strftime('%d/%m %H:%M')} → "
                f"{c.lectura_actual.fecha_visita.strftime('%d/%m %H:%M')}"
            )
            r2.append(
                [
                    c.nombre,
                    per,
                    format_number_chilean(c.delta_terreno, 2),
                    format_number_chilean(c.m3_app, 2) if c.m3_app is not None else "—",
                    format_number_chilean(c.diferencia, 2) if c.diferencia is not None else "—",
                    f"{c.pct:.1f}%" if c.pct is not None else "—",
                ]
            )
        t2 = doc.add_table(rows=len(r2), cols=len(h2))
        for i, row in enumerate(r2):
            for j, val in enumerate(row):
                t2.rows[i].cells[j].text = val
        estilizar_tabla_wes(t2, has_total_row=False)

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Tabla lecturas campo Peñalolén / CORMUP")
    parser.add_argument("--mes", type=int, default=None, help="Mes 1-12 (default: mes actual Chile)")
    parser.add_argument("--anio", type=int, default=None, help="Año (default: año actual Chile)")
    parser.add_argument(
        "--lecturas-csv",
        type=Path,
        default=None,
        help="CSV con lecturas de campo ya tomadas",
    )
    parser.add_argument("--out-dir", type=Path, default=None)
    args = parser.parse_args()

    ahora = datetime.now(CHILE_TZ)
    mes = args.mes or ahora.month
    anio = args.anio or ahora.year
    meses = (
        "",
        "enero",
        "febrero",
        "marzo",
        "abril",
        "mayo",
        "junio",
        "julio",
        "agosto",
        "septiembre",
        "octubre",
        "noviembre",
        "diciembre",
    )
    mes_label = f"{meses[mes]} {anio}"

    out_dir = args.out_dir or (OUT_BASE / f"{anio}{mes:02d}")
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"[INFO] Consultando app WES — CORMUP / Peñalolén — {mes_label}…")
    nodos = consultar_nodos(mes, anio)
    for n in nodos:
        flag = f" ⚠ {n.alerta}" if n.alerta else ""
        print(
            f"  {n.node_id} {n.nombre:28} última={n.ultima_chile}  "
            f"m³ mes={n.m3_mes:8.1f}{flag}"
        )

    lecturas: Dict[str, List[LecturaCampo]] = {}
    csv_path = args.lecturas_csv or (out_dir / "lecturas_campo.csv")
    if csv_path.is_file():
        lecturas = cargar_lecturas_csv(csv_path)
        print(f"[INFO] Lecturas de campo cargadas: {sum(len(v) for v in lecturas.values())} de {csv_path}")
    else:
        # Plantilla vacía
        plantilla = out_dir / "lecturas_campo.csv"
        if not plantilla.exists():
            with plantilla.open("w", encoding="utf-8", newline="") as f:
                w = csv.writer(f)
                w.writerow(["node_id", "fecha_visita", "lectura_contador", "observaciones"])
                for n in nodos:
                    w.writerow([n.node_id, "", "", n.nombre])
            print(f"[INFO] Plantilla CSV creada: {plantilla}")

    comps = calcular_comparaciones(nodos, lecturas) if lecturas else []

    ts = ahora.strftime("%Y%m%d_%H%M")
    xlsx = out_dir / f"Lecturas_campo_Penalolen_{anio}{mes:02d}_{ts}.xlsx"
    docx = out_dir / f"Lecturas_campo_Penalolen_{anio}{mes:02d}_{ts}.docx"
    estado = out_dir / f"estado_app_{anio}{mes:02d}_{ts}.json"

    escribir_excel(xlsx, nodos, mes_label, lecturas, comps)
    escribir_word(docx, nodos, mes_label, comps)

    serial: List[Dict[str, Any]] = []
    for n in nodos:
        d = asdict(n)
        d["ultima_utc"] = n.ultima_utc.isoformat() if n.ultima_utc else None
        serial.append(d)
    estado.write_text(json.dumps(serial, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[OK] Excel: {xlsx}")
    print(f"[OK] Word:  {docx}")
    print(f"[OK] JSON:  {estado}")
    if comps:
        print(f"[OK] Comparaciones campo vs app: {len(comps)}")
    else:
        print("[INFO] Todavía no hay 2 lecturas por colegio → pasame las de terreno y recalculo.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
