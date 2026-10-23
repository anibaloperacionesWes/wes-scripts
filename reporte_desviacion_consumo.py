#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Alerta temprana de consumo anómalo (muy alto / muy bajo) vs baseline 6 meses.

Para cada punto operativo:
  - Promedio diario últimos 6 meses
  - Promedio lunes–viernes (L–V)
  - Promedio sábado–domingo (S–D)
  - Compara el día de evaluación (ayer completo + hoy parcial) vs el promedio
    del mismo tipo de día (hábil o fin de semana)

Salida: Excel + Word + PDF en reports/Desviacion_Consumo/ y opcional Drive.

Uso:
  python reporte_desviacion_consumo.py
  python reporte_desviacion_consumo.py --solo-alertas
  python reporte_desviacion_consumo.py --company-id 000021
"""

from __future__ import annotations

import argparse
import csv
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from exclusiones_reportes import is_node_excluded
from wes_api_bff import listar_companias, measures_dates

ROOT = Path(__file__).resolve().parent
OUT_ROOT = ROOT / "reports" / "Desviacion_Consumo"

WD_NAMES = ("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo")

# Umbrales relativos al promedio del tipo de día (L–V o S–D)
RATIO_ALTO = 1.50
RATIO_BAJO = 0.50
# No alertar BAJO si el baseline del tipo de día es muy chico (ruido)
MIN_BASELINE_M3 = 0.50
# Días mínimos con data en el baseline
MIN_DIAS_BASELINE = 14
# Lotes de nodos por llamada measures/dates
BATCH_SIZE = 8
WORKERS = 6

# Empresas fuera de este informe (además de exclusiones_reportes)
EXCLUDED_COMPANY_IDS_DESVIACION = {
    "000010",  # Corporación Puente Alto
}

# Baseline especial: fin inclusive del periodo “normal” (antes de la caída).
# Calle de Servicio (CDUC) se vino abajo en agosto-2026 → promedio hasta 31/07/2026.
BASELINE_END_OVERRIDE: Dict[str, date] = {
    "000021-05": date(2026, 7, 31),
}


@dataclass
class NodoInfo:
    company_id: str
    company_name: str
    node_id: str
    node_name: str


@dataclass
class ResultadoNodo:
    info: NodoInfo
    dias_baseline: int = 0
    baseline_start: Optional[date] = None
    baseline_end: Optional[date] = None
    avg_all: Optional[float] = None
    avg_lv: Optional[float] = None
    avg_sd: Optional[float] = None
    n_lv: int = 0
    n_sd: int = 0
    # Ayer (día completo)
    ayer: Optional[date] = None
    ayer_m3: Optional[float] = None
    ayer_tipo: str = ""
    ayer_ref: Optional[float] = None
    ayer_ratio: Optional[float] = None
    ayer_alerta: str = "OK"  # OK | ALTO | BAJO | SIN_DATA
    # Hoy (puede ser parcial)
    hoy: Optional[date] = None
    hoy_m3: Optional[float] = None
    hoy_tipo: str = ""
    hoy_ref: Optional[float] = None
    hoy_ratio: Optional[float] = None
    hoy_alerta: str = "OK"
    serie: List[Tuple[date, float]] = field(default_factory=list)
    error: str = ""


def _fmt(v: Optional[float], dec: int = 2) -> str:
    if v is None:
        return "—"
    s = f"{v:.{dec}f}"
    ent, _, frac = s.partition(".")
    neg = ent.startswith("-")
    if neg:
        ent = ent[1:]
    parts = []
    while ent:
        parts.append(ent[-3:])
        ent = ent[:-3]
    ent = ".".join(reversed(parts))
    out = f"{ent},{frac}" if frac else ent
    return f"-{out}" if neg else out


def _tipo_dia(d: date) -> str:
    return "L-V" if d.weekday() < 5 else "S-D"


def listar_nodos(company_id: Optional[str] = None) -> List[NodoInfo]:
    out: List[NodoInfo] = []
    for c in listar_companias():
        cid = str(c.get("companyId") or "").strip()
        cname = str(c.get("name") or cid).strip()
        if company_id and cid != company_id:
            continue
        if cid in EXCLUDED_COMPANY_IDS_DESVIACION:
            continue
        if "PUENTE ALTO" in cname.upper():
            continue
        for n in c.get("nodes") or []:
            nid = str(n.get("nodeId") or "").strip()
            nname = str(n.get("name") or nid).strip()
            if not nid:
                continue
            if is_node_excluded(nid, company_id=cid, company_name=cname):
                continue
            out.append(
                NodoInfo(
                    company_id=cid,
                    company_name=cname,
                    node_id=nid,
                    node_name=nname,
                )
            )
    out.sort(key=lambda x: (x.company_name.lower(), x.node_id))
    return out


def _baseline_window(
    node_id: str,
    *,
    hoy: date,
    meses: int,
) -> Tuple[date, date]:
    """Retorna (baseline_start, baseline_end) inclusive."""
    default_end = hoy - timedelta(days=1)
    end = BASELINE_END_OVERRIDE.get(node_id, default_end)
    if end > default_end:
        end = default_end
    start = end - timedelta(days=meses * 30 + 5)
    return start, end


def _promedio(vals: List[float]) -> Optional[float]:
    return (sum(vals) / len(vals)) if vals else None


def _clasificar(ratio: Optional[float], baseline: Optional[float]) -> str:
    if ratio is None or baseline is None:
        return "SIN_DATA"
    if baseline < MIN_BASELINE_M3:
        # Baseline chico: solo alertar ALTO extremo, no BAJO
        if ratio >= RATIO_ALTO:
            return "ALTO"
        return "OK"
    if ratio >= RATIO_ALTO:
        return "ALTO"
    if ratio <= RATIO_BAJO:
        return "BAJO"
    return "OK"


def analizar_serie(
    info: NodoInfo,
    daily: Dict[date, float],
    *,
    hoy: date,
    baseline_start: date,
    baseline_end: date,
) -> ResultadoNodo:
    r = ResultadoNodo(info=info, hoy=hoy)
    r.baseline_start = baseline_start
    r.baseline_end = baseline_end
    serie = sorted((d, v) for d, v in daily.items() if baseline_start <= d <= hoy)
    r.serie = serie

    base_vals = [
        (d, v) for d, v in daily.items() if baseline_start <= d <= baseline_end
    ]
    base_vals.sort(key=lambda x: x[0])
    r.dias_baseline = len(base_vals)
    if r.dias_baseline < MIN_DIAS_BASELINE:
        r.error = f"Pocos días baseline ({r.dias_baseline})"
        r.ayer_alerta = "SIN_DATA"
        r.hoy_alerta = "SIN_DATA"
        return r

    all_v = [v for _, v in base_vals]
    lv_v = [v for d, v in base_vals if d.weekday() < 5]
    sd_v = [v for d, v in base_vals if d.weekday() >= 5]
    r.avg_all = _promedio(all_v)
    r.avg_lv = _promedio(lv_v)
    r.avg_sd = _promedio(sd_v)
    r.n_lv = len(lv_v)
    r.n_sd = len(sd_v)

    ayer = hoy - timedelta(days=1)
    r.ayer = ayer
    r.ayer_m3 = daily.get(ayer)
    r.ayer_tipo = _tipo_dia(ayer)
    r.ayer_ref = r.avg_lv if r.ayer_tipo == "L-V" else r.avg_sd
    if r.ayer_m3 is not None and r.ayer_ref and r.ayer_ref > 0:
        r.ayer_ratio = r.ayer_m3 / r.ayer_ref
    r.ayer_alerta = _clasificar(r.ayer_ratio, r.ayer_ref)

    r.hoy_m3 = daily.get(hoy)
    r.hoy_tipo = _tipo_dia(hoy)
    r.hoy_ref = r.avg_lv if r.hoy_tipo == "L-V" else r.avg_sd
    if r.hoy_m3 is not None and r.hoy_ref and r.hoy_ref > 0:
        r.hoy_ratio = r.hoy_m3 / r.hoy_ref
    # Hoy es parcial durante el día: solo ALTO (ya superó el umbral del día
    # completo). BAJO se evalúa con el día cerrado (ayer).
    if r.hoy_m3 is None:
        r.hoy_alerta = "SIN_DATA"
    elif r.hoy_ratio is not None and r.hoy_ratio >= RATIO_ALTO:
        r.hoy_alerta = "ALTO"
    else:
        r.hoy_alerta = "OK"
    return r


def fetch_batch(
    nodos: List[NodoInfo],
    start_ddmmyyyy: str,
    end_ddmmyyyy: str,
) -> Dict[str, Dict[date, float]]:
    ids = [n.node_id for n in nodos]
    raw = measures_dates(ids, start_ddmmyyyy, end_ddmmyyyy, timeout=150.0)
    out: Dict[str, Dict[date, float]] = {nid: {} for nid in ids}
    for item in raw:
        nid = str(item.get("nodeId") or "").strip()
        if nid not in out:
            out[nid] = {}
        for m in item.get("month") or []:
            ds = str(m.get("date") or "")[:10]
            if not ds:
                continue
            try:
                d = date.fromisoformat(ds)
            except ValueError:
                continue
            try:
                out[nid][d] = float(m.get("totalM3") or 0.0)
            except (TypeError, ValueError):
                out[nid][d] = 0.0
    return out


def procesar_todos(
    nodos: List[NodoInfo],
    *,
    hoy: date,
    meses: int = 6,
) -> List[ResultadoNodo]:
    # Rango de descarga: cubre el baseline más antiguo (overrides) hasta hoy
    windows = {_baseline_window(n.node_id, hoy=hoy, meses=meses) for n in nodos}
    fetch_start = min(s for s, _ in windows)
    fetch_end = hoy
    start_s = fetch_start.strftime("%d%m%Y")
    end_s = fetch_end.strftime("%d%m%Y")

    batches = [nodos[i : i + BATCH_SIZE] for i in range(0, len(nodos), BATCH_SIZE)]
    daily_by_node: Dict[str, Dict[date, float]] = {}
    errors: Dict[str, str] = {}

    def _one(batch: List[NodoInfo]) -> Tuple[Dict[str, Dict[date, float]], Optional[str]]:
        try:
            return fetch_batch(batch, start_s, end_s), None
        except Exception as e:
            return {}, str(e)

    print(
        f"[INFO] {len(nodos)} nodos en {len(batches)} lotes "
        f"(data {fetch_start} → {hoy})",
        flush=True,
    )
    for nid, end in BASELINE_END_OVERRIDE.items():
        print(f"  [INFO] baseline especial {nid}: hasta {end.isoformat()} (pre-caída)", flush=True)

    done = 0
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(_one, b): b for b in batches}
        for fut in as_completed(futs):
            batch = futs[fut]
            data, err = fut.result()
            done += 1
            if err:
                print(f"  [WARN] lote {done}/{len(batches)}: {err}", flush=True)
                # Reintento individual
                for n in batch:
                    try:
                        one = fetch_batch([n], start_s, end_s)
                        daily_by_node.update(one)
                    except Exception as e2:
                        errors[n.node_id] = str(e2)
            else:
                daily_by_node.update(data)
            if done % 5 == 0 or done == len(batches):
                print(f"  … {done}/{len(batches)} lotes", flush=True)

    resultados: List[ResultadoNodo] = []
    for n in nodos:
        if n.node_id in errors and n.node_id not in daily_by_node:
            r = ResultadoNodo(info=n, error=errors[n.node_id])
            r.ayer_alerta = "SIN_DATA"
            r.hoy_alerta = "SIN_DATA"
            resultados.append(r)
            continue
        daily = daily_by_node.get(n.node_id) or {}
        b_start, b_end = _baseline_window(n.node_id, hoy=hoy, meses=meses)
        resultados.append(
            analizar_serie(
                n,
                daily,
                hoy=hoy,
                baseline_start=b_start,
                baseline_end=b_end,
            )
        )
    return resultados


def _alerta_prioridad(r: ResultadoNodo) -> int:
    # BAJO (fuga/corte silencioso) prioriza sobre ALTO; ayer sobre hoy
    score = 0
    for a in (r.ayer_alerta, r.hoy_alerta):
        if a == "BAJO":
            score = max(score, 3)
        elif a == "ALTO":
            score = max(score, 2)
        elif a == "SIN_DATA":
            score = max(score, 1)
    return score


def _flag_display(r: ResultadoNodo) -> str:
    flags = []
    if r.ayer_alerta in ("ALTO", "BAJO"):
        flags.append(f"ayer:{r.ayer_alerta}")
    if r.hoy_alerta in ("ALTO", "BAJO"):
        flags.append(f"hoy:{r.hoy_alerta}")
    return " | ".join(flags) if flags else "OK"


def graficar_serie(r: ResultadoNodo, path: Path) -> None:
    if not r.serie:
        return
    fechas = [d for d, _ in r.serie]
    vals = [v for _, v in r.serie]
    fig, ax = plt.subplots(figsize=(10, 3.6))
    ax.plot(fechas, vals, color="#2e7ac8", linewidth=1.2, label="Consumo diario")
    if r.avg_all is not None:
        ax.axhline(r.avg_all, color="#666666", linestyle="--", linewidth=1, label=f"Prom. baseline ({_fmt(r.avg_all)})")
    if r.avg_lv is not None:
        ax.axhline(r.avg_lv, color="#2e7d32", linestyle=":", linewidth=1, label=f"Prom. L–V ({_fmt(r.avg_lv)})")
    if r.avg_sd is not None:
        ax.axhline(r.avg_sd, color="#ef6c00", linestyle=":", linewidth=1, label=f"Prom. S–D ({_fmt(r.avg_sd)})")
    ax.set_title(f"{r.info.company_name} — {r.info.node_name}")
    ax.set_ylabel("m³/día")
    ax.grid(True, alpha=0.3)
    ax.legend(loc="upper right", fontsize=8)
    fig.autofmt_xdate()
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=120)
    plt.close(fig)


def _sheet_title(company_name: str, used: Dict[str, int]) -> str:
    """Nombre de hoja Excel válido (máx. 31) y único."""
    bad = '[]:*?/\\'
    base = "".join(ch if ch not in bad else " " for ch in company_name).strip() or "Cliente"
    base = " ".join(base.split())[:31]
    if base not in used:
        used[base] = 1
        return base
    used[base] += 1
    suffix = f"_{used[base]}"
    return (base[: 31 - len(suffix)] + suffix)


EXCEL_HEADERS = [
    "Alerta ayer",
    "Alerta hoy",
    "Punto",
    "Días baseline",
    "Baseline desde",
    "Baseline hasta",
    "Prom. baseline (m³)",
    "Prom. L–V (m³)",
    "N L–V",
    "Prom. S–D (m³)",
    "N S–D",
    "Ayer fecha",
    "Ayer tipo",
    "Ayer m³",
    "Ayer ref tipo (m³)",
    "Ayer ratio",
    "Hoy fecha",
    "Hoy tipo",
    "Hoy m³ (parcial)",
    "Hoy ref tipo (m³)",
    "Hoy ratio",
    "Error",
]


def _fila_excel(r: ResultadoNodo, hoy: date) -> list:
    return [
        r.ayer_alerta,
        r.hoy_alerta,
        r.info.node_name,
        r.dias_baseline,
        r.baseline_start.isoformat() if r.baseline_start else "",
        r.baseline_end.isoformat() if r.baseline_end else "",
        None if r.avg_all is None else round(r.avg_all, 3),
        None if r.avg_lv is None else round(r.avg_lv, 3),
        r.n_lv,
        None if r.avg_sd is None else round(r.avg_sd, 3),
        r.n_sd,
        r.ayer.isoformat() if r.ayer else "",
        r.ayer_tipo,
        None if r.ayer_m3 is None else round(r.ayer_m3, 3),
        None if r.ayer_ref is None else round(r.ayer_ref, 3),
        None if r.ayer_ratio is None else round(r.ayer_ratio, 3),
        hoy.isoformat(),
        r.hoy_tipo,
        None if r.hoy_m3 is None else round(r.hoy_m3, 3),
        None if r.hoy_ref is None else round(r.hoy_ref, 3),
        None if r.hoy_ratio is None else round(r.hoy_ratio, 3),
        r.error,
    ]


def _estilo_hoja_cliente(ws, hoy: date, filas: List[ResultadoNodo]) -> None:
    fill_header = PatternFill("solid", fgColor="1F4E79")
    fill_alto = PatternFill("solid", fgColor="F4B183")
    fill_bajo = PatternFill("solid", fgColor="F8CBAD")
    fill_ok = PatternFill("solid", fgColor="C6EFCE")

    ws.append(EXCEL_HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill_header

    ordered = sorted(filas, key=lambda r: (-_alerta_prioridad(r), r.info.node_name.lower()))
    for r in ordered:
        ws.append(_fila_excel(r, hoy))
        alerta = r.ayer_alerta if r.ayer_alerta in ("ALTO", "BAJO") else r.hoy_alerta
        fill = {"ALTO": fill_alto, "BAJO": fill_bajo, "OK": fill_ok}.get(alerta)
        if fill:
            for cell in ws[ws.max_row]:
                if cell.column <= 2:
                    cell.fill = fill

    for col in ws.columns:
        letter = col[0].column_letter
        width = min(36, max(10, max(len(str(c.value or "")) for c in col) + 2))
        ws.column_dimensions[letter].width = width


def escribir_excel(resultados: List[ResultadoNodo], path: Path, hoy: date) -> None:
    """Una hoja por cliente (todos los puntos del cliente) + hoja Resumen."""
    wb = Workbook()
    # Resumen primero
    ws_res = wb.active
    ws_res.title = "Resumen"
    n_alto = sum(1 for r in resultados if "ALTO" in (r.ayer_alerta, r.hoy_alerta))
    n_bajo = sum(1 for r in resultados if "BAJO" in (r.ayer_alerta, r.hoy_alerta))
    clientes = sorted({r.info.company_name for r in resultados}, key=str.lower)
    ws_res.append(["Fecha generación", datetime.now().strftime("%Y-%m-%d %H:%M")])
    ws_res.append(["Día evaluación (hoy)", hoy.isoformat()])
    ws_res.append(["Día de la semana", WD_NAMES[hoy.weekday()]])
    ws_res.append(["Puntos analizados", len(resultados)])
    ws_res.append(["Clientes (hojas)", len(clientes)])
    ws_res.append(["Alertas ALTO", n_alto])
    ws_res.append(["Alertas BAJO", n_bajo])
    ws_res.append(["Umbral ALTO", f">= {RATIO_ALTO}× promedio tipo día"])
    ws_res.append(["Umbral BAJO", f"<= {RATIO_BAJO}× promedio tipo día"])
    ws_res.append(["Baseline default", "Últimos ~6 meses hasta ayer"])
    ws_res.append(
        [
            "Baseline Calle de Servicio",
            "Hasta 2026-07-31 (antes de la caída de agosto)",
        ]
    )
    ws_res.append(["Excluidos", "Puente Alto (000010)"])
    ws_res.append([])
    ws_res.append(["Cliente", "Puntos", "Alertas"])
    for cname in clientes:
        filas = [r for r in resultados if r.info.company_name == cname]
        n_al = sum(
            1
            for r in filas
            if r.ayer_alerta in ("ALTO", "BAJO") or r.hoy_alerta in ("ALTO", "BAJO")
        )
        ws_res.append([cname, len(filas), n_al])
    for col in ws_res.columns:
        letter = col[0].column_letter
        width = min(48, max(12, max(len(str(c.value or "")) for c in col) + 2))
        ws_res.column_dimensions[letter].width = width

    used_titles: Dict[str, int] = {}
    by_company: Dict[str, List[ResultadoNodo]] = {}
    for r in resultados:
        by_company.setdefault(r.info.company_name, []).append(r)

    for cname in sorted(by_company.keys(), key=str.lower):
        title = _sheet_title(cname, used_titles)
        ws = wb.create_sheet(title)
        _estilo_hoja_cliente(ws, hoy, by_company[cname])

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def escribir_csv_alertas(resultados: List[ResultadoNodo], path: Path) -> None:
    alertas = [
        r
        for r in resultados
        if r.ayer_alerta in ("ALTO", "BAJO") or r.hoy_alerta in ("ALTO", "BAJO")
    ]
    alertas.sort(key=lambda r: (-_alerta_prioridad(r), r.info.company_name))
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(
            [
                "alerta_ayer",
                "alerta_hoy",
                "empresa",
                "punto",
                "baseline_hasta",
                "prom_baseline",
                "prom_lv",
                "prom_sd",
                "ayer_m3",
                "ayer_ratio",
                "hoy_m3",
                "hoy_ratio",
            ]
        )
        for r in alertas:
            w.writerow(
                [
                    r.ayer_alerta,
                    r.hoy_alerta,
                    r.info.company_name,
                    r.info.node_name,
                    r.baseline_end.isoformat() if r.baseline_end else "",
                    _fmt(r.avg_all),
                    _fmt(r.avg_lv),
                    _fmt(r.avg_sd),
                    _fmt(r.ayer_m3),
                    _fmt(r.ayer_ratio, 2),
                    _fmt(r.hoy_m3),
                    _fmt(r.hoy_ratio, 2),
                ]
            )


def escribir_word(
    resultados: List[ResultadoNodo],
    path: Path,
    *,
    hoy: date,
    charts_dir: Path,
    max_charts: int = 25,
) -> None:
    doc = Document()
    title = doc.add_heading("Alerta de desviación de consumo", level=0)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    p = doc.add_paragraph()
    p.add_run(
        f"Fecha: {hoy.isoformat()} ({WD_NAMES[hoy.weekday()]}). "
        f"Baseline: promedio diario de los últimos ~6 meses (hasta ayer), "
        f"separado en L–V y S–D. "
        f"ALTO ≥ {RATIO_ALTO:.2f}× el promedio del tipo de día; "
        f"BAJO ≤ {RATIO_BAJO:.2f}× (baseline ≥ {MIN_BASELINE_M3} m³). "
        f"Calle de Servicio: baseline hasta 2026-07-31 (antes de la caída de agosto). "
        f"Puente Alto excluido."
    )

    alertas = [
        r
        for r in resultados
        if r.ayer_alerta in ("ALTO", "BAJO") or r.hoy_alerta in ("ALTO", "BAJO")
    ]
    alertas.sort(key=lambda r: (-_alerta_prioridad(r), r.info.company_name, r.info.node_id))

    n_alto = sum(1 for r in alertas if "ALTO" in (r.ayer_alerta, r.hoy_alerta))
    n_bajo = sum(1 for r in alertas if "BAJO" in (r.ayer_alerta, r.hoy_alerta))
    doc.add_heading("Resumen ejecutivo", level=1)
    doc.add_paragraph(
        f"Puntos analizados: {len(resultados)}. "
        f"Con alerta: {len(alertas)} (ALTO: {n_alto}, BAJO: {n_bajo}). "
        f"La comparación principal es el día de ayer (completo); hoy se muestra "
        f"como lectura parcial."
    )

    if not alertas:
        doc.add_paragraph("No hay puntos fuera de umbral en esta corrida.")
    else:
        doc.add_heading("Puntos en alerta", level=1)
        table = doc.add_table(rows=1, cols=8)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        for i, h in enumerate(
            [
                "Alerta",
                "Empresa / Punto",
                "Prom baseline",
                "Prom L–V",
                "Prom S–D",
                "Ayer m³",
                "Ratio ayer",
                "Hoy m³",
            ]
        ):
            hdr[i].text = h
            for run in hdr[i].paragraphs[0].runs:
                run.bold = True
                run.font.size = Pt(9)

        for r in alertas:
            row = table.add_row().cells
            flag = _flag_display(r)
            row[0].text = flag
            row[1].text = f"{r.info.company_name}\n{r.info.node_name}"
            row[2].text = _fmt(r.avg_all)
            row[3].text = _fmt(r.avg_lv)
            row[4].text = _fmt(r.avg_sd)
            row[5].text = _fmt(r.ayer_m3)
            row[6].text = _fmt(r.ayer_ratio, 2) + "×"
            row[7].text = _fmt(r.hoy_m3)
            for c in row:
                for para in c.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(8)
                        if "BAJO" in flag:
                            run.font.color.rgb = RGBColor(0x9C, 0x57, 0x00)
                        elif "ALTO" in flag:
                            run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)

        doc.add_heading("Gráficos (top alertas)", level=1)
        for i, r in enumerate(alertas[:max_charts]):
            chart_path = charts_dir / f"chart_{r.info.node_id}.png"
            graficar_serie(r, chart_path)
            doc.add_heading(
                f"{r.info.company_name} — {r.info.node_name}",
                level=2,
            )
            flag = _flag_display(r)
            bl = ""
            if r.baseline_end:
                bl = f" Baseline hasta {r.baseline_end.isoformat()}."
            doc.add_paragraph(
                f"Alerta: {flag}.{bl} "
                f"Ayer ({r.ayer} {r.ayer_tipo}): {_fmt(r.ayer_m3)} m³ "
                f"vs ref {_fmt(r.ayer_ref)} m³ ({_fmt(r.ayer_ratio, 2)}×). "
                f"Hoy parcial ({hoy} {r.hoy_tipo}): {_fmt(r.hoy_m3)} m³ "
                f"vs ref {_fmt(r.hoy_ref)} m³ ({_fmt(r.hoy_ratio, 2)}×). "
                f"Promedios: total {_fmt(r.avg_all)}, L–V {_fmt(r.avg_lv)}, "
                f"S–D {_fmt(r.avg_sd)}."
            )
            if chart_path.is_file():
                doc.add_picture(str(chart_path), width=Inches(6.2))

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


def docx_a_pdf(docx_path: Path) -> Optional[Path]:
    try:
        from docx2pdf import convert  # type: ignore
    except ImportError:
        convert = None
    pdf_path = docx_path.with_suffix(".pdf")
    if convert is not None and sys.platform == "win32":
        try:
            convert(str(docx_path), str(pdf_path))
            if pdf_path.is_file():
                return pdf_path
        except Exception:
            pass
    # LibreOffice
    try:
        import subprocess

        subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(docx_path.parent),
                str(docx_path),
            ],
            check=True,
            capture_output=True,
            timeout=180,
        )
        if pdf_path.is_file():
            return pdf_path
    except Exception:
        pass
    return None


def main() -> int:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    ap = argparse.ArgumentParser(description="Alerta desviación consumo vs 6 meses")
    ap.add_argument("--company-id", default=None, help="Limitar a una empresa (ej. 000021)")
    ap.add_argument("--hoy", default=None, help="Fecha YYYY-MM-DD (default: hoy Chile/UTC)")
    ap.add_argument("--meses", type=int, default=6, help="Meses de baseline (default 6)")
    ap.add_argument("--solo-alertas", action="store_true", help="Solo listar alertas en consola")
    ap.add_argument("--no-pdf", action="store_true")
    args = ap.parse_args()

    if args.hoy:
        hoy = date.fromisoformat(args.hoy)
    else:
        # Preferir fecha Chile si zoneinfo está disponible
        try:
            from zoneinfo import ZoneInfo

            hoy = datetime.now(ZoneInfo("America/Santiago")).date()
        except Exception:
            hoy = date.today()

    t0 = time.time()
    nodos = listar_nodos(args.company_id)
    print(f"[INFO] Nodos a analizar: {len(nodos)}", flush=True)
    if not nodos:
        print("[ERROR] Sin nodos")
        return 1

    resultados = procesar_todos(nodos, hoy=hoy, meses=args.meses)
    alertas = [
        r
        for r in resultados
        if r.ayer_alerta in ("ALTO", "BAJO") or r.hoy_alerta in ("ALTO", "BAJO")
    ]
    alertas.sort(key=lambda r: (-_alerta_prioridad(r), r.info.company_name))

    print()
    print("=" * 72)
    print(
        f"HOY {hoy.isoformat()} ({WD_NAMES[hoy.weekday()]}) — "
        f"{len(alertas)} alertas / {len(resultados)} puntos"
    )
    print("=" * 72)
    for r in alertas[:40]:
        flag = _flag_display(r)
        print(
            f"[{flag}] {r.info.company_name} | {r.info.node_name} ({r.info.node_id}) "
            f"| 6m={_fmt(r.avg_all)} L-V={_fmt(r.avg_lv)} S-D={_fmt(r.avg_sd)} "
            f"| ayer={_fmt(r.ayer_m3)} ({_fmt(r.ayer_ratio, 2)}×) "
            f"| hoy={_fmt(r.hoy_m3)} ({_fmt(r.hoy_ratio, 2)}×)"
        )
    if len(alertas) > 40:
        print(f"... y {len(alertas) - 40} más")

    if args.solo_alertas:
        print(f"[OK] {time.time() - t0:.1f}s")
        return 0

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = OUT_ROOT / f"DESVIACION_{stamp}"
    charts_dir = out_dir / "charts"
    out_dir.mkdir(parents=True, exist_ok=True)

    xlsx = out_dir / f"desviacion_consumo_{hoy.isoformat()}.xlsx"
    docx = out_dir / f"desviacion_consumo_{hoy.isoformat()}.docx"
    csv_path = out_dir / f"alertas_{hoy.isoformat()}.csv"

    escribir_excel(resultados, xlsx, hoy)
    escribir_csv_alertas(resultados, csv_path)
    escribir_word(resultados, docx, hoy=hoy, charts_dir=charts_dir)
    print(f"[OK] Excel: {xlsx}")
    print(f"[OK] Word:  {docx}")
    print(f"[OK] CSV:   {csv_path}")

    if not args.no_pdf:
        pdf = docx_a_pdf(docx)
        if pdf:
            print(f"[OK] PDF:   {pdf}")
        else:
            print("[WARN] No se pudo generar PDF (instalar LibreOffice o docx2pdf)")

    print(f"[OK] Listo en {time.time() - t0:.1f}s → {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
