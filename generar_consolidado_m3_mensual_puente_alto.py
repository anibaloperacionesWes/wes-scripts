"""
Consolidado m³ por mes y por colegio (Corporación Puente Alto, empresa 000010), desde la API WES.

Estrategia pensada para notebooks / máquinas modestas:
  - Un **nodo (colegio) a la vez**, en orden.
  - Dentro de cada nodo, los meses se pueden pedir **en paralelo** (--workers) para ir más rápido.
  - Primero intenta **una sola petición** ``dates.measures.csv`` con start/end del mes (DDMMYYYY).
  - Si falla o el total sale vacío, hace **fallback día a día** (CSV por día + ``totalM3`` JSON como la app).
  - Checkpoint JSON: si se corta el proceso, ``--solo-exportar-excel`` genera el Excel **sin API** con lo ya descargado.
  - Tras **cada** colegio: se actualiza ``consolidado_*_{año}_PARCIAL.xlsx`` (y un respaldo con timestamp).

Salida por defecto:
  ``reports/proyeccion ahorre puente 2025/``

Por defecto, m3 "sin WES" = m3 medidos * (1 + %% informe/100) (ej. 24,2%% -> x1,242). ``--sin-wes-formula division_sin`` restaura m3/(1-%%/100).

Forma recomendada (notebook / estable): **un colegio completo tras otro**, y dentro de cada uno **mes por mes sin paralelo**:
  python generar_consolidado_m3_mensual_puente_alto.py

Opción más rápida (paralela solo entre meses del mismo colegio):
  python generar_consolidado_m3_mensual_puente_alto.py --workers 4

Solo regenerar Excel desde checkpoint (instantáneo):
  python generar_consolidado_m3_mensual_puente_alto.py --solo-exportar-excel

Revisar cada punto (CSV horario vs totalM3 JSON vs total tras reconciliar, mismo mes):
  python generar_consolidado_m3_mensual_puente_alto.py --verificar-csv --year 2025 --mes-verificar 1

Un colegio, un trimestre:
  python generar_consolidado_m3_mensual_puente_alto.py --solo-node-id 000010-01 --mes-desde 1 --mes-fin 3
"""
from __future__ import annotations

import argparse
import calendar
import json
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests

from generar_reporte_word import (
    _chile_hours_from_dates_measures_csv_text,
    _reconcile_chile_hours_with_total_m3,
    _total_m3_from_json_for_chile_day,
    acl_node_base_url,
)
from reporte_puente_alto_lxm import (
    obtener_nodos_puente_alto,
    mapear_establecimiento_a_nodo,
)

ROOT = Path(__file__).resolve().parent
OUT_DIR_DEFAULT = ROOT / "reports" / "proyeccion ahorre puente 2025"

# Informe Auditoría WES Modelos - Base Puente Alto: eficiencia agregada municipal (%)
EFICIENCIA_GLOBAL_MODELO_PA = 49.0

PROYECCION_PCT_DEFAULT = OUT_DIR_DEFAULT / "PROYECCION_AHORRO_PUENTE_ALTO_2025_v3_clp.xlsx"

# Misma base que ``reporte_puente_alto_lxm`` (configuration.kpi[].efficiency por nodo, escala 0–100).
ENTITY_BASE_PA = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
NOMBRE_CSV_PCT_INFORME = "pct_auditoria_informe_pa.csv"


def _consumo_dia_fallback(node_id: str, dia: date) -> Tuple[float, int]:
    """
    Un día por pedido CSV; misma conversión Chile + alineación con ``totalM3`` JSON que la app/backoffice.
    """
    try:
        sess = requests.Session()
        url = f"{acl_node_base_url()}/nodes/{node_id}/dates.measures.csv"
        ds = dia.strftime("%d%m%Y")
        r = sess.get(url, params=[("start", ds), ("end", ds)], timeout=180)
        r.raise_for_status()
        horas = _chile_hours_from_dates_measures_csv_text(r.text, dia)
        tjson = _total_m3_from_json_for_chile_day(node_id, dia)
        horas, _ = _reconcile_chile_hours_with_total_m3(horas, tjson)
        s = sum(float(horas.get(hi, 0.0)) for hi in range(24))
        nh = sum(1 for hi in range(24) if horas.get(hi, 0.0) > 1e-9)
        return round(float(s), 4), nh
    except Exception:
        return 0.0, 0


def _mes_clave(year: int, month: int) -> str:
    return f"{year}-{month:02d}"


def consumo_mes_un_nodo(
    session: requests.Session,
    node_id: str,
    year: int,
    month: int,
) -> Tuple[float, int, str]:
    """
    Retorna (m³ del mes civil Chile, días con suma > 0, método: rango_mensual | dia_a_dia).
    """
    last_d = calendar.monthrange(year, month)[1]
    first = date(year, month, 1)
    last = date(year, month, last_d)

    url = f"{acl_node_base_url()}/nodes/{node_id}/dates.measures.csv"
    params = [("start", first.strftime("%d%m%Y")), ("end", last.strftime("%d%m%Y"))]

    try:
        r = session.get(url, params=params, timeout=180)
        r.raise_for_status()
        text = r.text
        if not text.strip() or "TIME" not in text.upper():
            raise ValueError("csv vacío o sin encabezado")

        total = 0.0
        dias_con = 0
        for day in range(1, last_d + 1):
            d = date(year, month, day)
            horas = _chile_hours_from_dates_measures_csv_text(text, d)
            tjson = _total_m3_from_json_for_chile_day(node_id, d)
            horas, _ = _reconcile_chile_hours_with_total_m3(horas, tjson)
            s = sum(float(horas.get(h, 0.0)) for h in range(24))
            if s > 1e-9:
                dias_con += 1
            total += s

        if total <= 1e-9:
            raise ValueError("total mensual nulo tras parsear")

        return round(total, 4), dias_con, "rango_mensual"
    except Exception:
        pass

    total = 0.0
    dias_con = 0
    d = first
    while d <= last:
        s, nh = _consumo_dia_fallback(node_id, d)
        if s > 1e-9:
            dias_con += 1
        total += s
        d += timedelta(days=1)

    return round(total, 4), dias_con, "dia_a_dia"


def verificar_cada_punto_csv_vs_json(
    year: int,
    month: int,
    solo_node_id: Optional[str],
    out_dir: Path,
) -> int:
    """
    Por cada nodo Puente Alto: un GET ``dates.measures.csv`` mensual; por día civil Chile suma CSV crudo,
    ``totalM3`` JSON y total tras ``_reconcile_*`` (igual que el consolidado). Escribe tabla en CSV.
    """
    nodos = obtener_nodos_puente_alto()
    nodos.sort(key=lambda x: x["nodeName"])
    if solo_node_id:
        nid_f = solo_node_id.strip()
        nodos = [n for n in nodos if n["nodeId"] == nid_f]
        if not nodos:
            print(f"[ERROR] No existe el nodo {nid_f} en empresa 000010.")
            return 1

    last_d = calendar.monthrange(year, month)[1]
    first = date(year, month, 1)
    last = date(year, month, last_d)
    sess = requests.Session()

    print(
        f"[VERIFICAR] periodo={year}-{month:02d} | nodos={len(nodos)} "
        "| columnas: crudo_csv = suma horas CSV; json = suma totalM3/dia; "
        "coherente_app = tras reconciliar (usa el consolidado)",
        flush=True,
    )
    rows: List[Dict[str, object]] = []

    for n in nodos:
        nid = n["nodeId"]
        nombre = n["nodeName"]
        url = f"{acl_node_base_url()}/nodes/{nid}/dates.measures.csv"
        params = [("start", first.strftime("%d%m%Y")), ("end", last.strftime("%d%m%Y"))]
        row: Dict[str, object] = {
            "node_id": nid,
            "colegio": nombre,
            "http_ok": False,
            "m3_csv_crudo_mes": "",
            "m3_json_suma_dias": "",
            "m3_coherente_app_mes": "",
            "pct_diff_crudo_vs_json": "",
            "pct_diff_final_vs_json": "",
            "dias_sin_totalM3_json": "",
            "dias_con_medida_crudo": "",
            "error": "",
        }
        try:
            r = sess.get(url, params=params, timeout=180)
            r.raise_for_status()
            text = r.text
            if not text.strip() or "TIME" not in text.upper():
                row["error"] = "csv vacio o sin TIME"
                rows.append(row)
                print(f"  {nid} {nombre[:36]:<36} ERROR {row['error']}", flush=True)
                continue

            sum_crudo = 0.0
            sum_json = 0.0
            sum_fin = 0.0
            dias_json_null = 0
            dias_crudo = 0
            for day in range(1, last_d + 1):
                d = date(year, month, day)
                horas = _chile_hours_from_dates_measures_csv_text(text, d)
                raw = sum(float(horas.get(h, 0.0)) for h in range(24))
                tj = _total_m3_from_json_for_chile_day(nid, d)
                horas_rec, _ = _reconcile_chile_hours_with_total_m3(horas, tj)
                fin = sum(float(horas_rec.get(h, 0.0)) for h in range(24))
                sum_crudo += raw
                sum_fin += fin
                if tj is not None:
                    sum_json += float(tj)
                else:
                    dias_json_null += 1
                if raw > 1e-9:
                    dias_crudo += 1

            row["http_ok"] = True
            row["m3_csv_crudo_mes"] = round(sum_crudo, 4)
            row["m3_json_suma_dias"] = round(sum_json, 4)
            row["m3_coherente_app_mes"] = round(sum_fin, 4)
            row["dias_sin_totalM3_json"] = dias_json_null
            row["dias_con_medida_crudo"] = dias_crudo
            if sum_json > 1e-9:
                row["pct_diff_crudo_vs_json"] = round(
                    abs(sum_crudo - sum_json) / sum_json * 100.0, 4
                )
                row["pct_diff_final_vs_json"] = round(
                    abs(sum_fin - sum_json) / sum_json * 100.0, 4
                )
            rows.append(row)
            pj = row["pct_diff_final_vs_json"] if row["pct_diff_final_vs_json"] != "" else "-"
            print(
                f"  {nid} {nombre[:32]:<32} crudo={row['m3_csv_crudo_mes']} "
                f"json={row['m3_json_suma_dias']} app={row['m3_coherente_app_mes']} "
                f"dif%%_crudo/json={row.get('pct_diff_crudo_vs_json','')} "
                f"dif%%_final/json={pj} sin_json_dias={dias_json_null}",
                flush=True,
            )
        except Exception as ex:
            row["error"] = str(ex)[:200]
            rows.append(row)
            print(f"  {nid} {nombre[:36]:<36} ERROR {row['error']}", flush=True)

    out_dir.mkdir(parents=True, exist_ok=True)
    out_csv = out_dir / f"verificacion_csv_pa_{year}_{month:02d}.csv"
    pd.DataFrame(rows).to_csv(out_csv, index=False, encoding="utf-8-sig")
    print(f"[OK] Tabla: {out_csv}", flush=True)
    return 0


def _fecha_iso_a_date(s: str) -> Optional[date]:
    s = (s or "").strip()
    if not s:
        return None
    try:
        return date.fromisoformat(s[:10])
    except ValueError:
        return None


def _parse_pct_api_o_excel(raw: object) -> Optional[float]:
    """Interpreta %% en texto (API o CSV). Fracciones 0–1 se convierten a %% si aplica."""
    if raw is None:
        return None
    s = str(raw).strip().replace("%", "").replace(",", ".")
    if s == "" or s.lower() in ("nan", "none"):
        return None
    try:
        x = float(s)
    except ValueError:
        return None
    if x < 0:
        return None
    if x == 0:
        return 0.0
    if x <= 1.0:
        return round(x * 100.0, 6)
    return round(x, 6)


def _efficiency_kpi_vigente_nodo(node: dict, ref: Optional[date] = None) -> Optional[float]:
    """``configuration.kpi[]`` vigente (o bloque que expira más tarde). None si no hay KPI."""
    ref = ref or date.today()
    kpis = (node.get("configuration") or {}).get("kpi") or []
    bloques: List[Tuple[date, date, dict]] = []
    for k in kpis:
        d0 = _fecha_iso_a_date(str(k.get("creationDate") or ""))
        d1 = _fecha_iso_a_date(str(k.get("expirationDate") or ""))
        if d0 is None or d1 is None:
            continue
        bloques.append((d0, d1, k))
    if not bloques:
        return None
    for d0, d1, k in bloques:
        if d0 <= ref <= d1:
            return _parse_pct_api_o_excel(k.get("efficiency"))
    bloques.sort(key=lambda t: t[1], reverse=True)
    return _parse_pct_api_o_excel(bloques[0][2].get("efficiency"))


def _cargar_kpi_efficiency_desde_api() -> Tuple[Dict[str, float], Dict[str, Optional[float]]]:
    """
    - Primer dict: solo valores > 0 (para poder usarlos como %% activa si no hay informe CSV).
    - Segundo dict: valor KPI del período vigente por nodo (incluye 0), para columnas de comparación.
    """
    try:
        r = requests.get(f"{ENTITY_BASE_PA}/companies/000010", timeout=30)
        r.raise_for_status()
        nodes = r.json().get("nodes") or []
    except Exception:
        return {}, {}
    positivos: Dict[str, float] = {}
    display: Dict[str, Optional[float]] = {}
    for n in nodes:
        nid = (n.get("nodeId") or "").strip()
        if not nid:
            continue
        v = _efficiency_kpi_vigente_nodo(n)
        display[nid] = v
        if v is not None and v > 0:
            positivos[nid] = float(v)
    return positivos, display


def _cargar_pct_informe_csv(csv_path: Path) -> Dict[str, float]:
    """
    CSV autoritativo del informe por colegio (%% auditoría por nodo).
    Columnas típicas: node_id, pct_eficiencia_auditoria (o segunda columna numérica).
    """
    if not csv_path.is_file():
        return {}
    df = pd.read_csv(csv_path, dtype=str)
    df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]
    if df.empty:
        return {}
    col_id = None
    for c in df.columns:
        cl = str(c).strip().lower().replace(" ", "_")
        if cl in ("node_id", "nodeid", "nodo"):
            col_id = c
            break
    if col_id is None:
        col_id = df.columns[0]
    col_pct = None
    for c in df.columns:
        if c == col_id:
            continue
        cl = str(c).strip().lower()
        if "pct" in cl or "eficiencia" in cl or cl in ("%", "porcentaje", "valor"):
            col_pct = c
            break
    if col_pct is None and len(df.columns) >= 2:
        col_pct = df.columns[1]
    elif col_pct is None:
        return {}
    out: Dict[str, float] = {}
    for _, row in df.iterrows():
        nid = str(row[col_id]).strip()
        if not nid or nid.lower() == "nan":
            continue
        pv = _parse_pct_api_o_excel(row[col_pct])
        if pv is not None and pv > 0:
            out[nid] = float(pv)
    return out


def _fusionar_pct_eficiencia(
    node_ids: List[str],
    pct_informe: Dict[str, float],
    pct_proyeccion: Dict[str, float],
    fallback: float,
) -> Tuple[Dict[str, float], Dict[str, str]]:
    """
    Prioridad: informe CSV por colegio > Ahorro (%%) Excel proyeccion > %% municipal.

    El KPI de la app (escala 0-100) no sustituye al %% de ahorro del informe para estimar sin WES.
    """
    final: Dict[str, float] = {}
    fuente: Dict[str, str] = {}
    for nid in node_ids:
        if nid in pct_informe:
            final[nid] = pct_informe[nid]
            fuente[nid] = "informe_csv"
        elif nid in pct_proyeccion:
            final[nid] = pct_proyeccion[nid]
            fuente[nid] = "proyeccion_xlsx"
        else:
            final[nid] = fallback
            fuente[nid] = "fallback_global"
    return final, fuente


def _cargar_pct_eficiencia_por_node(proyeccion_xlsx: Path) -> Dict[str, float]:
    """
    Lee ``Ahorro (%)`` del Detalle (= (Sin-Con)/Sin * 100) y lo asocia a node_id.
    Usa ``mapear_establecimiento_a_nodo`` (MAPA Excel / API) como en otros informes PA.
    """
    if not proyeccion_xlsx.is_file():
        return {}
    df = pd.read_excel(proyeccion_xlsx, sheet_name="Detalle por Punto")
    df = df[df["ID"].astype(str) != "TOTAL"]
    nodos = obtener_nodos_puente_alto()
    out: Dict[str, float] = {}
    for _, r in df.iterrows():
        nombre = str(r["Punto"]).strip()
        nid = mapear_establecimiento_a_nodo(nombre, nodos)
        if nid:
            out[nid] = float(r["Ahorro (%)"])
    return out


def _fila_desde_checkpoint(
    nid: str,
    nombre: str,
    mes_keys: List[str],
    data: Dict[str, object],
) -> Dict[str, object]:
    fila: Dict[str, object] = {"node_id": nid, "colegio": nombre}
    ta = 0.0
    for mk in mes_keys:
        fila[mk] = float(data.get(mk, 0.0) or 0.0)
        ta += float(fila[mk])
        fila[f"{mk}_dias_con_dato"] = int(data.get(f"{mk}_dias_con_dato", 0) or 0)
        fila[f"{mk}_metodo"] = str(data.get(f"{mk}_metodo", "") or "")
    fila["total_anio_m3"] = round(float(data.get("total_anio_m3", ta)), 4)
    return fila


def _armar_dataframe_consolidado(
    filas: List[Dict[str, object]],
    mes_keys: List[str],
    *,
    pct_por_node: Optional[Dict[str, float]] = None,
    pct_fuente_por_node: Optional[Dict[str, str]] = None,
    pct_kpi_app_por_node: Optional[Dict[str, Optional[float]]] = None,
    pct_fallback_global: float = EFICIENCIA_GLOBAL_MODELO_PA,
    sin_wes_formula: str = "con_mas_pct",
) -> pd.DataFrame:
    """
    ``pct_eficiencia_auditoria``: %% de ahorro por colegio (informe CSV > proyeccion > fallback).

    ``sin_wes_formula``:
      - ``con_mas_pct`` (default): Sin_WES_est = m3_medidos * (1 + %%/100), es decir 100%% del consumo medido
        más ``pct_eficiencia_auditoria``.
      - ``division_sin``: Sin_WES_est = m3_medidos / (1 - %%/100), coherente con el %% de auditoría
        Ahorro = (Sin-Con)/Sin (misma base que ``pct_auditoria_informe_pa.csv``).

    ``pct_eficiencia_kpi_app``: ``configuration.kpi[].efficiency`` del nodo (referencia app, no mezclada al %%).
    """
    df = pd.DataFrame(filas)
    cols_base = ["node_id", "colegio"] + mes_keys
    for mk in mes_keys:
        cols_base.extend([f"{mk}_dias_con_dato", f"{mk}_metodo"])
    cols_base.append("total_anio_m3")
    df = df[[c for c in cols_base if c in df.columns]]

    pct_map = pct_por_node or {}
    df["pct_eficiencia_auditoria"] = df["node_id"].map(pct_map)
    df["pct_eficiencia_auditoria"] = df["pct_eficiencia_auditoria"].fillna(pct_fallback_global)

    fuente_map = pct_fuente_por_node or {}
    df["pct_eficiencia_fuente"] = df["node_id"].map(fuente_map).fillna("")

    kpi_map = pct_kpi_app_por_node or {}

    def _celda_kpi(nid: object) -> object:
        if str(nid) == "TOTAL":
            return ""
        v = kpi_map.get(str(nid))
        if v is None:
            return ""
        return round(float(v), 4)

    df["pct_eficiencia_kpi_app"] = df["node_id"].apply(_celda_kpi)

    pcol = df["pct_eficiencia_auditoria"].astype(float)
    if sin_wes_formula == "con_mas_pct":
        fac = 1.0 + pcol / 100.0
        df["factor_sin_WES_vs_medicion"] = fac.round(6)
        for mk in mes_keys:
            if mk not in df.columns:
                continue
            df[f"{mk}_sin_WES_est_m3"] = (df[mk].astype(float) * fac).round(4)
    else:
        denom = 1.0 - pcol / 100.0
        denom = denom.where(denom > 1e-6)
        df["factor_sin_WES_vs_medicion"] = (1.0 / denom).round(6)
        for mk in mes_keys:
            if mk not in df.columns:
                continue
            df[f"{mk}_sin_WES_est_m3"] = (df[mk].astype(float) / denom).round(4)

    total_sin_anual = []
    for _, row in df.iterrows():
        s = sum(float(row.get(f"{mk}_sin_WES_est_m3", 0) or 0) for mk in mes_keys)
        total_sin_anual.append(round(s, 4))
    df["total_anio_sin_WES_est_m3"] = total_sin_anual

    ordered: List[str] = [
        "node_id",
        "colegio",
        "pct_eficiencia_auditoria",
        "pct_eficiencia_fuente",
        "factor_sin_WES_vs_medicion",
        "pct_eficiencia_kpi_app",
    ]
    for mk in mes_keys:
        ordered.extend(
            [mk, f"{mk}_sin_WES_est_m3", f"{mk}_dias_con_dato", f"{mk}_metodo"]
        )
    ordered.append("total_anio_m3")
    ordered.append("total_anio_sin_WES_est_m3")
    df = df[[c for c in ordered if c in df.columns]]

    total_row: Dict[str, object] = {
        "node_id": "TOTAL",
        "colegio": "SUMATORIA",
        "pct_eficiencia_auditoria": "",
        "pct_eficiencia_fuente": "",
        "factor_sin_WES_vs_medicion": "",
        "pct_eficiencia_kpi_app": "",
    }
    for mk in mes_keys:
        total_row[mk] = round(float(df[mk].sum()), 4) if mk in df.columns else 0.0
        sk = f"{mk}_sin_WES_est_m3"
        total_row[sk] = round(float(df[sk].sum()), 4) if sk in df.columns else 0.0
        total_row[f"{mk}_dias_con_dato"] = ""
        total_row[f"{mk}_metodo"] = ""
    total_row["total_anio_m3"] = round(float(df["total_anio_m3"].sum()), 4)
    total_row["total_anio_sin_WES_est_m3"] = round(float(df["total_anio_sin_WES_est_m3"].sum()), 4)

    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def _mes_keys_desde_dataframe(df2: pd.DataFrame) -> List[str]:
    out: List[str] = []
    for c in df2.columns:
        s = str(c).strip()
        if re.match(r"^\d{4}-\d{2}$", s):
            out.append(s)
    out.sort()
    return out


def _proyectar_dataframe_a_30_dias(df2: pd.DataFrame, mes_keys: List[str]) -> pd.DataFrame:
    """
    Proyecta a 30 días usando días con data real del mes:
      valor_proyectado = (m3_mes / dias_con_data) * 30
    Solo aplica cuando ``dias_con_data`` está entre 1 y 29.
    """
    out = df2.copy()
    mask_total = out["node_id"].astype(str).str.upper() == "TOTAL"
    detalle = out.loc[~mask_total].copy()

    for mk in mes_keys:
        dias_col = f"{mk}_dias_con_dato"
        if dias_col not in detalle.columns:
            continue
        dias_data = pd.to_numeric(detalle[dias_col], errors="coerce").fillna(0.0)
        mask_proj = (dias_data > 0.0) & (dias_data < 30.0)
        factor = pd.Series(1.0, index=detalle.index)
        factor.loc[mask_proj] = 30.0 / dias_data.loc[mask_proj]
        for col in (mk, f"{mk}_sin_WES_est_m3"):
            if col in detalle.columns:
                base = pd.to_numeric(detalle[col], errors="coerce").fillna(0.0)
                detalle[col] = (base * factor).round(4)
        # En la hoja proyectada, el periodo queda normalizado a 30 días cuando aplica.
        detalle.loc[mask_proj, dias_col] = 30

    if "total_anio_m3" in detalle.columns:
        detalle["total_anio_m3"] = (
            detalle[mes_keys].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1).round(4)
        )
    skeys = [f"{mk}_sin_WES_est_m3" for mk in mes_keys if f"{mk}_sin_WES_est_m3" in detalle.columns]
    if "total_anio_sin_WES_est_m3" in detalle.columns and skeys:
        detalle["total_anio_sin_WES_est_m3"] = (
            detalle[skeys].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1).round(4)
        )

    total_row = out.loc[mask_total].copy()
    if total_row.empty:
        return detalle

    for mk in mes_keys:
        total_row[mk] = round(float(pd.to_numeric(detalle[mk], errors="coerce").fillna(0.0).sum()), 4)
        sk = f"{mk}_sin_WES_est_m3"
        if sk in detalle.columns:
            total_row[sk] = round(float(pd.to_numeric(detalle[sk], errors="coerce").fillna(0.0).sum()), 4)
    if "total_anio_m3" in total_row.columns:
        total_row["total_anio_m3"] = round(float(pd.to_numeric(detalle["total_anio_m3"], errors="coerce").fillna(0.0).sum()), 4)
    if "total_anio_sin_WES_est_m3" in total_row.columns and "total_anio_sin_WES_est_m3" in detalle.columns:
        total_row["total_anio_sin_WES_est_m3"] = round(
            float(pd.to_numeric(detalle["total_anio_sin_WES_est_m3"], errors="coerce").fillna(0.0).sum()), 4
        )

    return pd.concat([detalle, total_row], ignore_index=True)


def _sheet_totalizados(
    df_base: pd.DataFrame,
    df_proj_30: pd.DataFrame,
    mes_keys: List[str],
) -> pd.DataFrame:
    """Totales por periodo: comparación Sheet1 (base) vs Sheet2 (proyectada 30 días)."""
    mask_total_base = df_base["node_id"].astype(str).str.upper() == "TOTAL"
    mask_total_proj = df_proj_30["node_id"].astype(str).str.upper() == "TOTAL"
    cols = [
        "periodo",
        "con_WES_sheet1_m3",
        "sin_WES_sheet1_m3",
        "ahorro_sheet1_m3",
        "con_WES_sheet2_m3",
        "sin_WES_sheet2_m3",
        "ahorro_sheet2_m3",
    ]
    if not mask_total_base.any() or not mask_total_proj.any():
        return pd.DataFrame(columns=cols)
    tr1 = df_base.loc[mask_total_base].iloc[0]
    tr2 = df_proj_30.loc[mask_total_proj].iloc[0]
    rows: List[Dict[str, object]] = []
    for mk in mes_keys:
        con_1 = float(tr1.get(mk, 0.0) or 0.0)
        sin_1 = float(tr1.get(f"{mk}_sin_WES_est_m3", 0.0) or 0.0)
        con_2 = float(tr2.get(mk, 0.0) or 0.0)
        sin_2 = float(tr2.get(f"{mk}_sin_WES_est_m3", 0.0) or 0.0)
        rows.append(
            {
                "periodo": mk,
                "con_WES_sheet1_m3": round(con_1, 4),
                "sin_WES_sheet1_m3": round(sin_1, 4),
                "ahorro_sheet1_m3": round(sin_1 - con_1, 4),
                "con_WES_sheet2_m3": round(con_2, 4),
                "sin_WES_sheet2_m3": round(sin_2, 4),
                "ahorro_sheet2_m3": round(sin_2 - con_2, 4),
            }
        )
    con_1y = float(tr1.get("total_anio_m3", 0.0) or 0.0)
    sin_1y = float(tr1.get("total_anio_sin_WES_est_m3", 0.0) or 0.0)
    con_2y = float(tr2.get("total_anio_m3", 0.0) or 0.0)
    sin_2y = float(tr2.get("total_anio_sin_WES_est_m3", 0.0) or 0.0)
    rows.append(
        {
            "periodo": "TOTAL_ANUAL",
            "con_WES_sheet1_m3": round(con_1y, 4),
            "sin_WES_sheet1_m3": round(sin_1y, 4),
            "ahorro_sheet1_m3": round(sin_1y - con_1y, 4),
            "con_WES_sheet2_m3": round(con_2y, 4),
            "sin_WES_sheet2_m3": round(sin_2y, 4),
            "ahorro_sheet2_m3": round(sin_2y - con_2y, 4),
        }
    )
    return pd.DataFrame(rows)


def _guardar_consolidado(df2: pd.DataFrame, xlsx_path: Path, csv_path: Path) -> None:
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    mes_keys = _mes_keys_desde_dataframe(df2)
    df_proj_30 = _proyectar_dataframe_a_30_dias(df2, mes_keys)
    df_tot = _sheet_totalizados(df2, df_proj_30, mes_keys)
    with pd.ExcelWriter(xlsx_path) as writer:
        df2.to_excel(writer, sheet_name="Sheet1", index=False)
        df_proj_30.to_excel(writer, sheet_name="Sheet2", index=False)
        df_tot.to_excel(writer, sheet_name="Sheet3", index=False)
    df2.to_csv(csv_path, index=False, encoding="utf-8-sig")


def _exportar_consolidado_parcial(
    filas: List[Dict[str, object]],
    mes_keys: List[str],
    year: int,
    out_dir: Path,
    indice: int,
    nid: str,
    nombre_colegio: str,
    *,
    pct_por_node: Dict[str, float],
    pct_fuente_por_node: Dict[str, str],
    pct_kpi_app_por_node: Dict[str, Optional[float]],
    pct_fallback_global: float,
    sin_wes_formula: str,
) -> Tuple[Path, Path]:
    """
    Tras cada colegio: escribe consolidado actualizado + copia con timestamp por colegio.
    Siempre sobrescribe ``*_PARCIAL.xlsx`` para tener el último estado en un solo archivo.
    """
    df2 = _armar_dataframe_consolidado(
        filas,
        mes_keys,
        pct_por_node=pct_por_node,
        pct_fuente_por_node=pct_fuente_por_node,
        pct_kpi_app_por_node=pct_kpi_app_por_node,
        pct_fallback_global=pct_fallback_global,
        sin_wes_formula=sin_wes_formula,
    )
    parcial_xlsx = out_dir / f"consolidado_m3_mensual_colegios_puente_alto_{year}_PARCIAL.xlsx"
    parcial_csv = out_dir / f"consolidado_m3_mensual_colegios_puente_alto_{year}_PARCIAL.csv"
    _guardar_consolidado(df2, parcial_xlsx, parcial_csv)
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    safe = "".join(c if c.isalnum() or c in " -_" else "_" for c in nombre_colegio)[:40]
    snap = (
        out_dir
        / f"consolidado_m3_mensual_PA_{year}_paso{indice:02d}_{nid}_{safe}_{ts}.xlsx"
    )
    snap_csv = snap.with_suffix(".csv")
    _guardar_consolidado(df2, snap, snap_csv)
    return parcial_xlsx, snap


def _exportar_desde_checkpoint(
    year: int,
    mes_keys: List[str],
    nodos: List[Dict[str, str]],
    checkpoint: Dict[str, Dict[str, object]],
    out_dir: Path,
    *,
    pct_por_node: Dict[str, float],
    pct_fuente_por_node: Dict[str, str],
    pct_kpi_app_por_node: Dict[str, Optional[float]],
    pct_fallback_global: float,
    sin_wes_formula: str,
) -> Tuple[Path, Path]:
    """Genera Excel/CSV solo con datos ya guardados en checkpoint (sin API)."""
    filas: List[Dict[str, object]] = []
    for n in nodos:
        nid = n["nodeId"]
        ck = checkpoint.get(nid)
        if not ck:
            continue
        filas.append(_fila_desde_checkpoint(nid, n["nodeName"], mes_keys, ck))

    if not filas:
        raise FileNotFoundError(
            "No hay datos en checkpoint para ningun nodo (archivo ausente, vacio o sin claves de nodos). "
            "Ejecute sin --solo-exportar-excel para descargar desde la API, o restaure el JSON de checkpoint."
        )

    df2 = _armar_dataframe_consolidado(
        filas,
        mes_keys,
        pct_por_node=pct_por_node,
        pct_fuente_por_node=pct_fuente_por_node,
        pct_kpi_app_por_node=pct_kpi_app_por_node,
        pct_fallback_global=pct_fallback_global,
        sin_wes_formula=sin_wes_formula,
    )
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    xlsx = out_dir / f"consolidado_m3_mensual_colegios_puente_alto_{year}_{ts}_desde_checkpoint.xlsx"
    csv_path = out_dir / f"consolidado_m3_mensual_colegios_puente_alto_{year}_{ts}_desde_checkpoint.csv"
    _guardar_consolidado(df2, xlsx, csv_path)
    return xlsx, csv_path


def main() -> int:
    ap = argparse.ArgumentParser(description="Consolidado m³/mes por colegio Puente Alto (API WES)")
    ap.add_argument("--year", type=int, default=2025)
    ap.add_argument("--mes-desde", type=int, default=1, help="Mes inicial inclusive (1-12)")
    ap.add_argument("--mes-fin", type=int, default=12, help="Mes final inclusive (1-12)")
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=OUT_DIR_DEFAULT,
        help="Carpeta de salida Excel/CSV/checkpoint",
    )
    ap.add_argument(
        "--solo-node-id",
        default=None,
        help="Procesar solo este nodeId (ej. 000010-01)",
    )
    ap.add_argument(
        "--workers",
        type=int,
        default=1,
        help=(
            "Meses en paralelo **solo dentro del mismo colegio** (nunca se mezclan dos colegios). "
            "1 = recomendado: mes tras mes. 4–6 = más rápido si la API aguanta. Default: 1"
        ),
    )
    ap.add_argument(
        "--solo-exportar-excel",
        action="store_true",
        help="Solo generar Excel/CSV desde checkpoint_consolidado_m3_mensual_<año>.json (sin llamar API)",
    )
    ap.add_argument(
        "--proyeccion-pct",
        type=Path,
        default=PROYECCION_PCT_DEFAULT,
        help=(
            "Excel de proyeccion/auditoría con hoja Detalle por Punto "
            "(columnas Punto + Ahorro %%). Por defecto: PROYECCION_*_v3_clp.xlsx"
        ),
    )
    ap.add_argument(
        "--pct-fallback-global",
        type=float,
        default=EFICIENCIA_GLOBAL_MODELO_PA,
        help=(
            "%% eficiencia si no hay match por colegio (Informe modelo municipal Puente Alto: 49%%)"
        ),
    )
    ap.add_argument(
        "--pct-informe-csv",
        type=Path,
        default=None,
        help=(
            "CSV con %% del informe de auditoría por node_id (prioridad sobre Excel de proyección). "
            "Si no se indica y existe "
            f"<out-dir>/{NOMBRE_CSV_PCT_INFORME}, se usa automáticamente."
        ),
    )
    ap.add_argument(
        "--sin-wes-formula",
        choices=("con_mas_pct", "division_sin"),
        default="con_mas_pct",
        help=(
            "con_mas_pct (default): Sin_WES_est = m3_medidos*(1+p/100), "
            "equivale a 100%% + p%%. "
            "division_sin: Sin_WES_est = m3_medidos/(1-p/100), con p=%% ahorro (Sin-Con)/Sin."
        ),
    )
    ap.add_argument(
        "--verificar-csv",
        action="store_true",
        help=(
            "Solo revisar cada punto (nodo): GET mensual dates.measures.csv vs suma totalM3 JSON por día "
            "vs total tras reconciliar (coherente con la app). Escribe verificacion_csv_pa_<año>_<mes>.csv."
        ),
    )
    ap.add_argument(
        "--mes-verificar",
        type=int,
        default=None,
        metavar="N",
        help="Mes 1-12 para --verificar-csv (por defecto: --mes-desde).",
    )
    args = ap.parse_args()

    if args.verificar_csv:
        out_dir_v = Path(args.out_dir).expanduser().resolve()
        mv = args.mes_verificar if args.mes_verificar is not None else args.mes_desde
        mv = max(1, min(12, mv))
        return verificar_cada_punto_csv_vs_json(args.year, mv, args.solo_node_id, out_dir_v)

    year = args.year
    m0 = max(1, min(12, args.mes_desde))
    m1 = max(1, min(12, args.mes_fin))
    if m0 > m1:
        print("[ERROR] mes-desde no puede ser mayor que mes-fin.")
        return 1
    out_dir = Path(args.out_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    pct_path = Path(args.proyeccion_pct).expanduser().resolve()
    pct_fallback = float(args.pct_fallback_global)

    pct_proyeccion = _cargar_pct_eficiencia_por_node(pct_path)
    informe_csv_path = (
        Path(args.pct_informe_csv).expanduser().resolve()
        if args.pct_informe_csv
        else (out_dir / NOMBRE_CSV_PCT_INFORME)
    )
    pct_informe = _cargar_pct_informe_csv(informe_csv_path)
    pct_kpi_uso, pct_kpi_display = _cargar_kpi_efficiency_desde_api()

    nodos = obtener_nodos_puente_alto()
    nodos.sort(key=lambda x: x["nodeName"])
    if args.solo_node_id:
        nid_f = args.solo_node_id.strip()
        nodos = [n for n in nodos if n["nodeId"] == nid_f]
        if not nodos:
            print(f"[ERROR] No existe el nodo {nid_f} en empresa 000010.")
            return 1

    node_ids = [n["nodeId"] for n in nodos]
    pct_map, pct_fuente = _fusionar_pct_eficiencia(
        node_ids,
        pct_informe,
        pct_proyeccion,
        pct_fallback,
    )

    if pct_path.is_file():
        print(
            f"[INFO] %% proyeccion (Ahorro %%): {len(pct_proyeccion)} nodos desde {pct_path.name}",
            flush=True,
        )
    else:
        print(f"[INFO] No existe {pct_path}; sin columna intermedia de proyeccion.", flush=True)
    if pct_informe:
        print(
            f"[INFO] %% informe por colegio (CSV): {len(pct_informe)} nodos desde {informe_csv_path.name}",
            flush=True,
        )
    elif informe_csv_path.is_file():
        print(f"[WARN] {informe_csv_path.name} existe pero no hay filas %% validas.", flush=True)
    if pct_kpi_uso:
        print(f"[INFO] KPI eficiencia en nodo (solo columna pct_eficiencia_kpi_app): {len(pct_kpi_uso)} nodos.", flush=True)

    fuentes = list(pct_fuente.values())
    print(
        "[INFO] %% ahorro usado sin_WES por fuente: "
        f"informe_csv={fuentes.count('informe_csv')} "
        f"proyeccion_xlsx={fuentes.count('proyeccion_xlsx')} "
        f"fallback_global={fuentes.count('fallback_global')}",
        flush=True,
    )
    print(f"[INFO] Formula sin_WES_est: {args.sin_wes_formula}", flush=True)

    meses = list(range(m0, m1 + 1))
    mes_keys = [_mes_clave(year, m) for m in meses]

    ck_path = out_dir / f"checkpoint_consolidado_m3_mensual_{year}.json"
    checkpoint: Dict[str, Dict[str, object]] = {}
    if ck_path.is_file():
        try:
            checkpoint = json.loads(ck_path.read_text(encoding="utf-8"))
            print(f"[INFO] Checkpoint cargado: {ck_path}")
        except Exception:
            checkpoint = {}

    if args.solo_exportar_excel:
        try:
            xlsx, csv_p = _exportar_desde_checkpoint(
                year,
                mes_keys,
                nodos,
                checkpoint,
                out_dir,
                pct_por_node=pct_map,
                pct_fuente_por_node=pct_fuente,
                pct_kpi_app_por_node=pct_kpi_display,
                pct_fallback_global=pct_fallback,
                sin_wes_formula=args.sin_wes_formula,
            )
        except FileNotFoundError as e:
            print(f"[ERROR] {e}", flush=True)
            return 1
        print(f"[OK] Excel (solo checkpoint): {xlsx}")
        print(f"[OK] CSV:   {csv_p}")
        return 0

    workers = max(1, int(args.workers))
    print(
        "[INFO] Modo: un colegio a la vez; "
        + ("mes por mes secuencial" if workers == 1 else f"hasta {workers} meses en paralelo por colegio"),
        flush=True,
    )
    filas: List[Dict[str, object]] = []

    for idx, n in enumerate(nodos, start=1):
        nid = n["nodeId"]
        nombre = n["nodeName"]
        ck_key = nid
        print("=" * 72)
        print(f"[{idx}/{len(nodos)}] {nid} — {nombre}")

        ck = checkpoint.get(ck_key, {})
        meta_ok = (
            ck.get("_meta_year") == year
            and ck.get("_meta_mes_desde") == m0
            and ck.get("_meta_mes_fin") == m1
        )
        if ck.get("completo") and meta_ok:
            print("  [SKIP] Ya descargado para este mismo rango de meses (checkpoint).")
            filas.append(_fila_desde_checkpoint(nid, nombre, mes_keys, ck))
            px, snap = _exportar_consolidado_parcial(
                filas,
                mes_keys,
                year,
                out_dir,
                len(filas),
                nid,
                nombre,
                pct_por_node=pct_map,
                pct_fuente_por_node=pct_fuente,
                pct_kpi_app_por_node=pct_kpi_display,
                pct_fallback_global=pct_fallback,
                sin_wes_formula=args.sin_wes_formula,
            )
            print(
                f"  [OK] Consolidado parcial: {px.name} + copia: {snap.name}",
                flush=True,
            )
            continue
        if ck.get("completo") and not meta_ok:
            print(
                "  [INFO] Checkpoint de otro rango de meses: se vuelve a consultar este período.",
                flush=True,
            )

        fila: Dict[str, object] = {"node_id": nid, "colegio": nombre}
        total_anio = 0.0

        def _fetch_month(m: int) -> Tuple[int, float, int, str]:
            # Sesión propia por hilo (evita condiciones de carrera).
            sess = requests.Session()
            m3, dias_con, metodo = consumo_mes_un_nodo(sess, nid, year, m)
            return m, m3, dias_con, metodo

        if workers == 1:
            print("  Descargando mes a mes ...", flush=True)
            resultados = {}
            for m in meses:
                _, m3, dias_con, metodo = _fetch_month(m)
                resultados[m] = (m3, dias_con, metodo)
                print(f"    Mes {m:02d}: {m3} m3 | dias: {dias_con} | {metodo}", flush=True)
        else:
            print(f"  Meses en paralelo solo este colegio (workers={workers}) ...", flush=True)
            resultados = {}
            with ThreadPoolExecutor(max_workers=workers) as ex:
                fut_map = {ex.submit(_fetch_month, m): m for m in meses}
                for fut in as_completed(fut_map):
                    m, m3, dias_con, metodo = fut.result()
                    resultados[m] = (m3, dias_con, metodo)
                    print(f"    Mes {m:02d}: {m3} m3 | dias: {dias_con} | {metodo}", flush=True)

        for m in meses:
            mk = _mes_clave(year, m)
            m3, dias_con, metodo = resultados[m]
            fila[mk] = m3
            fila[f"{mk}_dias_con_dato"] = dias_con
            fila[f"{mk}_metodo"] = metodo
            total_anio += m3

            checkpoint[ck_key] = checkpoint.get(ck_key, {})
            checkpoint[ck_key].update(
                {k: v for k, v in fila.items() if isinstance(v, (str, int, float))}
            )
            checkpoint[ck_key]["total_anio_m3"] = round(total_anio, 4)
            ck_path.write_text(json.dumps(checkpoint, ensure_ascii=False, indent=2), encoding="utf-8")

        fila["total_anio_m3"] = round(total_anio, 4)
        checkpoint[ck_key]["completo"] = True
        checkpoint[ck_key]["_meta_year"] = year
        checkpoint[ck_key]["_meta_mes_desde"] = m0
        checkpoint[ck_key]["_meta_mes_fin"] = m1
        ck_path.write_text(json.dumps(checkpoint, ensure_ascii=False, indent=2), encoding="utf-8")
        filas.append(fila)

        px, snap = _exportar_consolidado_parcial(
            filas,
            mes_keys,
            year,
            out_dir,
            len(filas),
            nid,
            nombre,
            pct_por_node=pct_map,
            pct_fuente_por_node=pct_fuente,
            pct_kpi_app_por_node=pct_kpi_display,
            pct_fallback_global=pct_fallback,
            sin_wes_formula=args.sin_wes_formula,
        )
        print(
            f"  [OK] Consolidado parcial: {px.name} + copia: {snap.name}",
            flush=True,
        )

    df2 = _armar_dataframe_consolidado(
        filas,
        mes_keys,
        pct_por_node=pct_map,
        pct_fuente_por_node=pct_fuente,
        pct_kpi_app_por_node=pct_kpi_display,
        pct_fallback_global=pct_fallback,
        sin_wes_formula=args.sin_wes_formula,
    )
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    xlsx = out_dir / f"consolidado_m3_mensual_colegios_puente_alto_{year}_{ts}.xlsx"
    csv_path = out_dir / f"consolidado_m3_mensual_colegios_puente_alto_{year}_{ts}.csv"
    _guardar_consolidado(df2, xlsx, csv_path)

    total_row_val = float(df2.iloc[-1]["total_anio_m3"])

    print("=" * 72)
    print(f"[OK] Excel final: {xlsx}")
    print(f"[OK] CSV final:   {csv_path}")
    print(f"[OK] Ultimo PARCIAL (mismo contenido que final si termino todo): consolidado_*_{year}_PARCIAL.xlsx")
    print(f"[OK] Total ano (suma nodos): {total_row_val} m3")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
