"""
Informes de gestión hídrica (formato Zapallar) — lote comercial de 8.

Periodo: 1 al 31 de agosto de 2026.
Clientes: Fundo Zapallar, Inchcape, Nido de Águilas, Lo Valledor,
UDD, Club Providencia, AGUNSA Lampa, AGUNSA Intermodal.

Uso:
    python generar_informes_gestion_hidrica_lote_agosto2026.py
"""

from __future__ import annotations

import json
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import requests

from agregado_extendido_extra import _serie_mensual_nodo
from exclusiones_reportes import FUNDO_ZAPALLAR_NODE_IDS, filter_node_ids
from generar_reporte_word import (
    acl_node_base_url,
    calculate_nocturnal_metrics,
    fetch_json,
    flatten_measures,
    get_node_name,
    get_water_price_per_m3,
    normalize_measures_payload,
    parse_date,
    summarize_consumption,
)
from informe_gestion_hidrica_pdf import (
    Accion,
    Hallazgo,
    InformeSpec,
    PuntoIndicador,
    SerieDiaria,
    VisitaTecnicaSpec,
    _fecha_es,
    _fmt,
    _fmt_clp,
    build_chart_6_meses,
    build_chart_nocturno,
    build_chart_puntos,
    render_mensual,
    render_one_pager,
    resolve_logo,
)
from perfiles_clientes_hidricos import aplicar_perfil_valledor
from visitas_tecnicas_formulario import (
    VisitaTecnica,
    cargar_visitas_periodo,
    visitas_de_cliente,
)

ENTITY = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
START = "01/08/2026"
END = "31/08/2026"
START_DT = datetime(2026, 8, 1)
END_DT = datetime(2026, 8, 31)
PERIODO_DIAS = 31
CACHE_DIR = Path("/tmp/gh_agosto31")

CLIENTES: List[Dict[str, Any]] = [
    {
        "key": "zapallar",
        "company_id": "000027",
        "folder": "Fundo_Zapallar",
        "cliente": "Fundo Zapallar",
        "sitio": "Fundo Zapallar",
        "sujeto": "el fundo",
        "node_ids": list(FUNDO_ZAPALLAR_NODE_IDS),
        "start": "05/08/2026",
        "end": "31/08/2026",
        "periodo_corto": "Fundo Zapallar · 5 al 31 de agosto de 2026",
        "usar_kpi_ultimo_mes_6m": True,
        "nota_agosto": (
            "Agosto se informa del 5 al 31. Los días 1 al 4 se excluyen y no se extrapola el mes."
        ),
        "panorama_nota": (
            "El consumo se considera del 5 al 31 de agosto. Los días 1 al 4 se excluyen "
            "y no se extrapola el resto del mes."
        ),
        "apply_exclusions": False,
        "matriz_id": "000027-01",
        "matriz_name": "Matriz ESVAL",
        "additive": False,
        "nocturnal_explain": "bombas_estanques",
        "kpi_label": "Consumo de entrada",
        "short_names": {
            "000027-01": "Matriz ESVAL",
            "000027-02": "Estanque inferior",
            "000027-03": "Etapa N.º 5",
            "000027-04": "Etapa N.º 1 al 4",
            "000027-06": "Etapa N.º 1",
            "000027-07": "Etapa N.º 2",
            "000027-08": "Etapa N.º 3",
            "000027-09": "Riego estanque ESVAL",
        },
        "leyenda": ("Matriz ESVAL (entrada real)", "Estanques y etapas (aguas abajo)"),
        "chart_nota": (
            "la Matriz ESVAL representa la entrada real al fundo. Los demás medidores "
            "están aguas abajo y miden el mismo flujo en distintos tramos; por lo tanto, "
            "sus consumos no se suman."
        ),
        "nocturno_nota": (
            "Los volúmenes de los medidores interiores se expresan respecto de la entrada "
            "principal, pero no son aditivos porque corresponden a tramos encadenados."
        ),
    },
    {
        "key": "inchcape",
        "company_id": "000012",
        "folder": "Inchcape",
        "cliente": "Inchcape",
        "sitio": "Inchcape Quilicura",
        "sujeto": "la sucursal",
        "node_ids": [
            "000012-06",
            "000012-07",
            "000012-08",
            "000012-09",
            "000012-10",
            "000012-11",
            "000012-12",
        ],
        "apply_exclusions": False,
        "matriz_id": "000012-06",
        "matriz_name": "Matriz Principal",
        "additive": False,
        "nocturnal_explain": "wes",
        "kpi_label": "Consumo de entrada",
        "short_names": {
            "000012-06": "Matriz Principal",
            "000012-07": "Dercomaq",
            "000012-08": "Lav. Máquinas",
            "000012-09": "Casino",
            "000012-10": "Proderco",
            "000012-11": "Camarines",
            "000012-12": "Edificio JCB",
        },
        "leyenda": ("Matriz Principal (total sucursal)", "Puntos internos (parte de la matriz)"),
        "chart_nota": (
            "la Matriz Principal representa el consumo real de la sucursal. Los demás "
            "medidores están aguas abajo y miden derivaciones del mismo caudal; por lo tanto, "
            "sus consumos no se suman."
        ),
        "nocturno_nota": (
            "Los volúmenes de los medidores interiores se expresan respecto de la Matriz "
            "Principal, pero no son aditivos porque corresponden a derivaciones internas."
        ),
    },
    {
        "key": "nido",
        "company_id": "000007",
        "folder": "Nido_de_Aguilas",
        "cliente": "Nido de Águilas",
        "sitio": "Nido de Águilas",
        "sujeto": "el colegio",
        "node_ids": None,
        "apply_exclusions": True,
        "matriz_id": None,
        "matriz_name": "",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "short_names": {
            "000007-01": "Estanque B",
            "000007-02": "Teatro",
            "000007-03": "High School",
            "000007-04": "Elementary",
            "000007-05": "Piscina",
            "000007-06": "Pozo profundo",
            "000007-07": "Estanque C",
        },
        "leyenda": None,
        "chart_nota": (
            "cada barra es el consumo del punto en el periodo. Los volúmenes de estanques, "
            "pozo y recintos se presentan por separado y se suman al total del colegio."
        ),
        "nocturno_nota": (
            "El consumo nocturno se suma entre puntos porque corresponden a recintos y "
            "sistemas distintos del colegio."
        ),
    },
    aplicar_perfil_valledor(
        {
            "key": "valledor",
            "company_id": "000002",
            "folder": "Lo_Valledor",
            "cliente": "Lo Valledor",
            "sitio": "Lo Valledor",
            "sujeto": "el recinto",
            "node_ids": ["000002-01", "000002-03"],
            "apply_exclusions": False,
            "matriz_id": None,
            "matriz_name": "",
            "additive": True,
            "kpi_label": "Consumo total",
            "short_names": {
                "000002-01": "P1",
                "000002-03": "Barrio Norte",
            },
            "leyenda": None,
            "chart_nota": (
                "P1 y Barrio Norte son recintos distintos y se suman al total. "
                "El anillo muestra la participación; las barras, el volumen en m³. "
                "P1 concentra ~94 % del recinto."
            ),
        }
    ),
    {
        "key": "udd",
        "company_id": "000026",
        "folder": "UDD",
        "cliente": "UDD",
        "sitio": "Universidad del Desarrollo",
        "sujeto": "el campus",
        "node_ids": ["000026-01", "000026-02"],
        "apply_exclusions": False,
        "matriz_id": None,
        "matriz_name": "",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "short_names": {
            "000026-01": "Impulsión Honduras",
            "000026-02": "Aula Magna",
        },
        "leyenda": None,
        "chart_nota": (
            "Impulsión Honduras y Aula Magna son puntos distintos: sus consumos se suman al total."
        ),
        "nocturno_nota": (
            "El consumo nocturno se suma entre ambos puntos porque no corresponden al mismo caudal."
        ),
    },
    {
        "key": "club",
        "company_id": "000031",
        "folder": "Club_Providencia",
        "cliente": "Club Providencia",
        "sitio": "Club Providencia",
        "sujeto": "el club",
        "node_ids": ["000031-01", "000031-02"],
        "apply_exclusions": False,
        "matriz_id": None,
        "matriz_name": "",
        "additive": True,
        "nocturnal_explain": "piscina",
        "kpi_label": "Consumo total",
        "short_names": {
            "000031-01": "Matriz Fitness",
            "000031-02": "Matriz Piscina",
        },
        "leyenda": None,
        "chart_nota": (
            "Fitness y Piscina son matrices distintas: sus consumos se suman al total del club."
        ),
        "nocturno_nota": (
            "El consumo nocturno de piscina puede incluir llenado o reposición; se suma al de fitness."
        ),
    },
    {
        "key": "lampa",
        "company_id": "000020",
        "folder": "AGUNSA_Lampa",
        "cliente": "AGUNSA Lampa",
        "sitio": "AGUNSA Lampa",
        "sujeto": "el depósito",
        "node_ids": ["000020-01", "000020-02", "000020-03", "000020-04"],
        "apply_exclusions": False,
        "matriz_id": None,
        "matriz_name": "",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "short_names": {
            "000020-01": "Depósito",
            "000020-02": "Módulo D",
            "000020-03": "Módulo ABC",
            "000020-04": "Módulo E",
        },
        "leyenda": None,
        "chart_nota": (
            "Depósito y módulos se presentan por separado y se suman al total del recinto Lampa."
        ),
        "nocturno_nota": (
            "El consumo nocturno se suma entre depósito y módulos porque son puntos distintos."
        ),
    },
    {
        "key": "intermodal",
        "company_id": "000020",
        "folder": "Agunsa_Intermodal",
        "cliente": "AGUNSA Intermodal",
        "sitio": "AGUNSA Intermodal San Antonio",
        "sujeto": "el recinto",
        "node_ids": ["000020-05"],
        "apply_exclusions": False,
        "matriz_id": "000020-05",
        "matriz_name": "Intermodal",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "short_names": {"000020-05": "Intermodal"},
        "leyenda": None,
        "chart_nota": "un único punto de monitoreo representa el consumo del recinto Intermodal.",
        "nocturno_nota": (
            "El consumo nocturno corresponde al único medidor del recinto Intermodal San Antonio."
        ),
    },
]


def _nodos_api(company_id: str, company_name: str) -> List[str]:
    r = requests.get(f"{ENTITY}/companies/{company_id}", timeout=30)
    r.raise_for_status()
    ids = sorted(n["nodeId"] for n in r.json().get("nodes", []) if n.get("nodeId"))
    return filter_node_ids(ids, company_id=company_id, company_name=company_name)


def _daily_series(measures) -> List[Dict[str, Any]]:
    by_day: Dict[str, float] = {}
    for m in measures:
        key = m.date.strftime("%Y-%m-%d")
        by_day[key] = by_day.get(key, 0.0) + float(m.total_m3)
    return [{"date": d, "m3": by_day[d]} for d in sorted(by_day)]


def _fetch_node(cfg: dict, node_id: str, start_dt: datetime, end_dt: datetime) -> Dict[str, Any]:
    payload_raw = fetch_json(
        f"{acl_node_base_url()}/nodes/measures/dates",
        params=[
            ("id", node_id),
            ("start", start_dt.strftime("%d%m%Y")),
            ("end", end_dt.strftime("%d%m%Y")),
        ],
    )
    payload = normalize_measures_payload(payload_raw, node_id)
    measures = flatten_measures(payload)
    summary = summarize_consumption(measures)
    noct = calculate_nocturnal_metrics(node_id, start_dt, end_dt, company_id=cfg["company_id"])
    mx = summary.get("max")
    short = (cfg.get("short_names") or {}).get(node_id) or get_node_name(node_id)
    return {
        "node_id": node_id,
        "node_name": get_node_name(node_id),
        "short_name": short,
        "total": float(summary.get("total") or 0),
        "promedio_diario": float(summary.get("promedio_diario") or 0),
        "dias": int(summary.get("dias") or 0),
        "max_m3": float(mx.total_m3) if mx else 0.0,
        "max_fecha": mx.date.strftime("%Y-%m-%d") if mx else None,
        "nocturno_m3": float(noct.get("consumo_nocturno_total") or 0),
        "nocturno_dias": int(noct.get("dias_con_consumo_nocturno") or 0),
        "nocturno_cobertura": int(noct.get("dias_con_datos_horarios") or 0),
        "daily": _daily_series(measures),
    }


def _periodo_bounds(cfg: dict) -> Tuple[datetime, datetime, int]:
    start_dt = parse_date(cfg.get("start") or START)
    end_dt = parse_date(cfg.get("end") or END, end_of_day=True)
    dias = (end_dt.date() - start_dt.date()).days + 1
    return start_dt, end_dt, dias


_MES_ABREV_NUM = {
    "ene": 1,
    "feb": 2,
    "mar": 3,
    "abr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "ago": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dic": 12,
}


def _parse_lab_mes(lab: str) -> Optional[Tuple[int, int]]:
    parts = str(lab).replace("*", "").strip().split()
    if len(parts) != 2:
        return None
    mes = _MES_ABREV_NUM.get(parts[0].lower()[:3])
    if mes is None:
        return None
    try:
        yy = int(parts[1])
    except ValueError:
        return None
    year = 2000 + yy if yy < 100 else yy
    return (year, mes)


def _aplicar_serie_6m(
    serie: List[Tuple[str, float]],
    excluir_meses: Optional[Sequence[Tuple[int, int]]] = None,
    ultimo_m3: Optional[float] = None,
    end_dt: Optional[datetime] = None,
) -> List[Tuple[str, float]]:
    excl = {(int(y), int(m)) for y, m in (excluir_meses or [])}
    out: List[Tuple[str, float]] = []
    for lab, val in serie:
        parsed = _parse_lab_mes(lab)
        if parsed and parsed in excl:
            continue
        out.append((lab, float(val)))
    if ultimo_m3 is not None and out:
        parsed = _parse_lab_mes(out[-1][0])
        if end_dt is None or (parsed and parsed == (end_dt.year, end_dt.month)):
            out[-1] = (out[-1][0], float(ultimo_m3))
    return out


def _iter_days(start: datetime, end: datetime):
    cur = datetime(start.year, start.month, start.day)
    last = datetime(end.year, end.month, end.day)
    while cur <= last:
        yield cur
        cur += timedelta(days=1)


def fetch_cliente(cfg: dict) -> dict:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache = CACHE_DIR / f"{cfg['key']}.json"
    if cache.is_file():
        print(f"[CACHE] {cfg['cliente']}", flush=True)
        return json.loads(cache.read_text(encoding="utf-8"))

    node_ids = cfg["node_ids"]
    if not node_ids:
        node_ids = _nodos_api(cfg["company_id"], cfg["cliente"])
        cfg["node_ids"] = node_ids
    start_dt, end_dt, dias = _periodo_bounds(cfg)
    results: Dict[str, dict] = {}
    workers = int(cfg.get("workers") or max(2, min(6, len(node_ids))))
    print(f"[INFO] {cfg['cliente']}: {len(node_ids)} nodos", flush=True)
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(_fetch_node, cfg, nid, start_dt, end_dt): nid for nid in node_ids}
        for fut in as_completed(futs):
            nid = futs[fut]
            results[nid] = fut.result()
            r = results[nid]
            print(
                f"  {nid} {r['short_name']}: {_fmt(r['total'], 1)} m³  "
                f"noct {_fmt(r['nocturno_m3'], 1)} cob={r['nocturno_cobertura']}",
                flush=True,
            )
    ordered = [results[nid] for nid in node_ids if nid in results]
    excluir_meses = [tuple(x) for x in (cfg.get("excluir_meses_6m") or [])]
    n_meses = 6 + len(excluir_meses)
    matriz_id = cfg.get("matriz_id")
    if matriz_id and not cfg["additive"]:
        ref = next(n for n in ordered if n["node_id"] == matriz_id)
        entrada = float(ref["total"])
        nocturno = float(ref["nocturno_m3"])
        max_m3 = float(ref["max_m3"])
        max_fecha = ref["max_fecha"]
        serie6 = _serie_mensual_nodo(matriz_id, end_dt, n_meses)
    else:
        entrada = sum(float(n["total"]) for n in ordered)
        nocturno = sum(float(n["nocturno_m3"]) for n in ordered)
        top = max(ordered, key=lambda n: float(n["max_m3"] or 0))
        max_m3 = float(top["max_m3"])
        max_fecha = top["max_fecha"]
        by_label: Dict[str, float] = {}
        labels_order: List[str] = []
        for n in ordered:
            try:
                serie = _serie_mensual_nodo(n["node_id"], end_dt, n_meses)
            except Exception as e:
                print(f"[ADVERTENCIA] 6 meses {n['short_name']}: {e}", flush=True)
                continue
            for lab, val in serie:
                lab2 = lab.replace("*", "")
                if lab2 not in by_label:
                    labels_order.append(lab2)
                    by_label[lab2] = 0.0
                by_label[lab2] += float(val)
        serie6 = [(lab, by_label[lab]) for lab in labels_order]
    ultimo_6m = float(entrada) if cfg.get("usar_kpi_ultimo_mes_6m") else None
    serie6 = _aplicar_serie_6m(serie6, excluir_meses, ultimo_6m, end_dt)
    price = get_water_price_per_m3(cfg["company_id"], node_ids[0], {})
    payload = {
        "cfg_key": cfg["key"],
        "company_id": cfg["company_id"],
        "nodos": ordered,
        "serie_6_meses": [{"label": a, "m3": b} for a, b in serie6],
        "price_per_m3": price,
        "periodo_dias": dias,
        "start_iso": start_dt.strftime("%Y-%m-%d"),
        "end_iso": end_dt.strftime("%Y-%m-%d"),
        "kpi": {
            "entrada": entrada,
            "promedio": entrada / dias if dias else 0.0,
            "nocturno": nocturno,
            "pct_nocturno": (nocturno / entrada * 100.0) if entrada else 0.0,
            "max_m3": max_m3,
            "max_fecha": max_fecha,
            "costo_nocturno": nocturno * price,
        },
    }
    cache.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return payload


def _daily_full(nodo: dict, start: datetime, end: datetime) -> SerieDiaria:
    by = {d["date"]: float(d["m3"]) for d in nodo.get("daily") or []}
    fechas = []
    valores: List[Optional[float]] = []
    for cur in _iter_days(start, end):
        key = cur.strftime("%Y-%m-%d")
        fechas.append(cur)
        valores.append(by[key] if key in by else None)
    return SerieDiaria(nombre=nodo["short_name"], fechas=fechas, valores=valores, lectura="")


def _cobertura_huecos(nodo: dict, start: datetime, end: datetime) -> List[str]:
    have = {d["date"] for d in nodo.get("daily") or []}
    missing = []
    for cur in _iter_days(start, end):
        key = cur.strftime("%Y-%m-%d")
        if key not in have:
            missing.append(key)
    return missing


def _bounds_from_data(data: dict) -> Tuple[datetime, datetime, int]:
    if data.get("start_iso") and data.get("end_iso"):
        start = datetime.strptime(data["start_iso"], "%Y-%m-%d")
        end = datetime.strptime(data["end_iso"], "%Y-%m-%d")
        dias = int(data.get("periodo_dias") or ((end.date() - start.date()).days + 1))
        return start, end, dias
    return START_DT, END_DT, PERIODO_DIAS


def _salto_mensual(data: dict) -> Optional[Tuple[float, float, float]]:
    """(actual, mediana previa, ratio) si el mes ≥ 2,5× la mediana de los meses anteriores."""
    serie = data.get("serie_6_meses") or []
    if len(serie) < 3:
        return None
    vals = [float(x.get("m3") or 0) for x in serie]
    actual = vals[-1]
    prev = [v for v in vals[:-1] if v > 0]
    if not prev:
        return None
    med = sorted(prev)[len(prev) // 2]
    if med < 1:
        return None
    ratio = actual / med
    if ratio < 2.5:
        return None
    return actual, med, ratio


def _clasificar(
    cfg: dict,
    pct: float,
    evento_fuerte: bool,
    salto: Optional[Tuple[float, float, float]] = None,
) -> Tuple[str, str]:
    explain = cfg.get("nocturnal_explain")
    if cfg.get("cpa_estado") == "instalado_pendiente":
        return (
            "REQUIERE ATENCIÓN",
            "el equipo CPA está instalado y falta activarlo y programarlo; sin ese control "
            "el recinto no tiene regulación automática de caudal.",
        )
    if salto:
        actual, med, ratio = salto
        return (
            "REQUIERE ATENCIÓN",
            f"el consumo del mes alcanzó {_fmt(actual, 0)} m³, "
            f"{_fmt(ratio, 1)} veces la mediana de los meses previos ({_fmt(med, 0)} m³).",
        )
    if explain == "wes" and pct <= 16:
        return (
            "BAJO CONTROL",
            "el caudal nocturno de la Matriz Principal corresponde al funcionamiento del "
            "sistema WES y no se identifica pérdida en la sucursal.",
        )
    if explain == "mercado":
        return (
            "BAJO CONTROL",
            "el horario full del mercado es 22:00–03:00 y el caudal de madrugada "
            "corresponde a esa operación, no a una pérdida.",
        )
    if explain == "bombas_estanques" and pct >= 18:
        return (
            "EN OBSERVACIÓN",
            "consumo nocturno que requiere seguimiento y validación frente a la operación "
            "habitual del condominio (bombas, estanques y distintos horarios).",
        )
    if explain == "piscina" and pct >= 18:
        return (
            "EN OBSERVACIÓN",
            "consumo nocturno que puede corresponder a reposición o llenado de piscina y "
            "requiere validación frente a la operación del club.",
        )
    if pct >= 35:
        return (
            "REQUIERE ATENCIÓN",
            f"la participación nocturna alcanzó {_fmt(round(pct), 0)} % y debe revisarse "
            "para separar uso programado de posibles pérdidas.",
        )
    if pct >= 18 or evento_fuerte:
        return (
            "EN OBSERVACIÓN",
            "se observa un patrón o evento que requiere seguimiento antes de cerrar el periodo.",
        )
    return (
        "BAJO CONTROL",
        "el consumo se mantuvo estable y la participación nocturna no indica pérdida.",
    )


def _hallazgos(cfg: dict, data: dict) -> Tuple[List[Hallazgo], bool]:
    kpi = data["kpi"]
    nodos = data["nodos"]
    start, end, dias = _bounds_from_data(data)
    pct = float(kpi["pct_nocturno"])
    nocturno = float(kpi["nocturno"])
    entrada = float(kpi["entrada"])
    explain = cfg.get("nocturnal_explain")
    evento_fuerte = False
    hall: List[Hallazgo] = []

    salto = _salto_mensual(data)
    if cfg.get("cpa_estado") == "instalado_pendiente":
        hall.append(
            Hallazgo(
                "ATENCIÓN",
                "Equipo CPA pendiente de activar y programar",
                "El control WES está instalado y aún no opera.",
                "Activarlo y programarlo con el peak 22:00–03:00. Sin CPA no hay "
                "control automático de caudal: por eso el estado del periodo es Requiere atención.",
            )
        )

    if salto:
        actual, med, ratio = salto
        mx = float(kpi.get("max_m3") or 0)
        mx_f = kpi.get("max_fecha")
        extra = (
            f" Máximo diario {_fmt(mx, 1)} m³ el {_fecha_es(mx_f)}."
            if mx_f
            else ""
        )
        hall.append(
            Hallazgo(
                "ATENCIÓN",
                f"Agosto superó {_fmt(ratio, 1)} veces la línea base",
                f"{_fmt(actual, 0)} m³ frente a una mediana de {_fmt(med, 0)} m³ en los meses previos.",
                "Revisar fuga, riego continuo o error de medición. No es el patrón de mar–jul."
                + extra,
            )
        )

    hd = cfg.get("hallazgo_dato") or {}
    if hd.get("titulo"):
        hall.append(
            Hallazgo(
                hd.get("prioridad") or "INFORMATIVA",
                hd["titulo"],
                hd.get("dato") or "",
                hd.get("lectura") or "",
            )
        )

    if explain == "wes":
        hall.append(
            Hallazgo(
                "INFORMATIVA",
                f"{_fmt(round(pct), 0)} % del consumo nocturno es operación WES",
                f"{_fmt(nocturno, 1)} m³ entre 00:00 y 06:59.",
                "El caudal de madrugada de la Matriz Principal se explica por los ciclos de "
                "control y regulación del sistema WES, no por una pérdida. Se mantiene como "
                "referencia operativa.",
            )
        )
    elif explain == "mercado":
        hall.append(
            Hallazgo(
                "INFORMATIVA",
                "El caudal 24 h es operación del recinto",
                f"{_fmt(round(pct), 0)} % en 00:00–06:59 ({_fmt(nocturno, 1)} m³); peak 22:00–03:00.",
                "No se interpreta como pérdida. Queda como línea base del periodo.",
            )
        )
    elif explain == "bombas_estanques":
        hall.append(
            Hallazgo(
                "SEGUIMIENTO" if pct >= 18 else "INFORMATIVA",
                "Consumo nocturno en seguimiento",
                f"{_fmt(nocturno, 1)} m³ entre 00:00 y 06:59: {_fmt(round(pct), 0)} % de la entrada.",
                "Por la operación mediante bombas y estanques, puede corresponder a "
                "funcionamiento habitual. Definir una línea base nocturna propia y observar "
                "su evolución.",
            )
        )
    else:
        prio = "ATENCIÓN" if pct >= 35 else ("SEGUIMIENTO" if pct >= 18 else "INFORMATIVA")
        hall.append(
            Hallazgo(
                prio,
                f"{_fmt(round(pct), 0)} % del consumo fue nocturno",
                f"{_fmt(nocturno, 1)} m³ entre 00:00 y 06:59.",
                "Corresponde confirmar horarios de uso, riego o llenado. Si el patrón no tiene "
                "explicación operacional, elevar el estado a Requiere atención.",
            )
        )

    # Evento: peak relativo al promedio del punto
    candidatos = []
    for n in nodos:
        avg = float(n["total"]) / max(int(n["dias"]) or 1, 1)
        mx = float(n["max_m3"] or 0)
        if avg > 0 and mx >= max(3.0 * avg, 8.0):
            candidatos.append((mx / avg, n))
    if candidatos:
        candidatos.sort(key=lambda t: t[0], reverse=True)
        n = candidatos[0][1]
        evento_fuerte = candidatos[0][0] >= 4.0 and float(n["max_m3"]) >= 12
        hall.append(
            Hallazgo(
                "SEGUIMIENTO",
                f"Evento en {n['short_name']}",
                f"{_fmt(n['max_m3'], 1)} m³ el {_fecha_es(n['max_fecha'])}.",
                "Confirmar si correspondió a una maniobra operacional. Si se repite sin "
                "explicación, elevar el estado a Requiere atención.",
            )
        )

    # Cobertura: el punto con más días faltantes (mínimo 3).
    ref = None
    if cfg.get("matriz_id"):
        ref = next((n for n in nodos if n["node_id"] == cfg["matriz_id"]), None)
    if ref is None:
        ref = max(nodos, key=lambda n: float(n["total"]))
    gaps = sorted(
        ((len(_cobertura_huecos(n, start, end)), n) for n in nodos),
        key=lambda t: t[0],
        reverse=True,
    )
    cover_n = gaps[0][1] if gaps and gaps[0][0] >= 3 else None
    if cover_n is None and _cobertura_huecos(ref, start, end):
        cover_n = ref
    if cover_n is not None and _cobertura_huecos(cover_n, start, end):
        hall.append(
            Hallazgo(
                "INFORMATIVA",
                f"Cobertura parcial en {cover_n['short_name']}",
                f"La serie tiene {cover_n['dias']} días con datos de {dias} del periodo.",
                "Los días sin registro no se interpolan ni se extrapolan. La menor cobertura "
                "no representa por sí sola ausencia de consumo.",
            )
        )
    else:
        # tercer hallazgo: punto de menor/mayor participación o internos no suman
        if not cfg["additive"] and len(nodos) > 1:
            hall.append(
                Hallazgo(
                    "INFORMATIVA",
                    "Los puntos internos no se suman al total",
                    f"La {cfg['matriz_name']} es el consumo real de {cfg['sujeto']}.",
                    "Los demás medidores están aguas abajo o en derivaciones del mismo caudal. "
                    "Se leen como control interno, no como volumen adicional.",
                )
            )
        elif len(nodos) >= 2:
            top = max(nodos, key=lambda n: float(n["total"]))
            pct_top = float(top["total"]) / entrada * 100.0 if entrada else 0.0
            hall.append(
                Hallazgo(
                    "INFORMATIVA",
                    f"{top['short_name']} concentra el {_fmt(round(pct_top), 0)} % del total",
                    f"{_fmt(top['total'], 1)} m³ en el periodo.",
                    "Es el punto de mayor demanda. Sirve como referencia de seguimiento mes a mes.",
                )
            )
        else:
            if cfg.get("hallazgo_dato") or cfg.get("excluir_meses_6m") or start.day != 1:
                hall.append(
                    Hallazgo(
                        "INFORMATIVA",
                        f"Periodo informado: {dias} días válidos",
                        f"{ref['short_name']} registra los {dias} días del recorte operativo.",
                        cfg.get("panorama_nota")
                        or "No se incluyen días con dato anómalo ni se extrapola el mes.",
                    )
                )
            else:
                hall.append(
                    Hallazgo(
                        "INFORMATIVA",
                        "Serie continua en el mes completo",
                        f"{ref['short_name']} registra los {dias} días del periodo.",
                        "Agosto se evalúa del 1 al 31 y no se extrapola. La serie queda como línea base.",
                    )
                )

    return hall[:3], evento_fuerte


def _acciones(hallazgos: Sequence[Hallazgo], cfg: dict) -> List[Accion]:
    acts: List[Accion] = []
    for h in hallazgos:
        if h.prioridad == "SEGUIMIENTO" and "nocturno" in h.titulo.lower():
            acts.append(
                Accion(
                    "Confirmar horarios de riego, llenado o uso nocturno.",
                    "7 días",
                    "Separar consumo programado de posibles pérdidas.",
                    "Administración / operación",
                )
            )
        elif h.prioridad == "SEGUIMIENTO" and h.titulo.startswith("Evento"):
            acts.append(
                Accion(
                    f"Revisar {h.titulo.lower()}.",
                    "7 días",
                    "Validar causa y descartar fuga o dato anómalo.",
                    "Operación + WES",
                )
            )
        elif "CPA" in h.titulo:
            acts.append(
                Accion(
                    "Activar y programar el equipo CPA.",
                    "7 días",
                    "Poner el control en servicio alineado al horario del mercado.",
                    "WES + operación",
                )
            )
        elif h.prioridad == "ATENCIÓN" and "línea base" in h.titulo.lower():
            acts.append(
                Accion(
                    "Diagnosticar el salto de consumo de agosto.",
                    "7 días",
                    "Identificar fuga, riego continuo o falla de medición.",
                    "Operación + WES",
                )
            )
        elif "sensor" in h.titulo.lower() or "pulso" in h.titulo.lower():
            acts.append(
                Accion(
                    "Mantener el sensor de pulso operativo y validar lecturas diarias.",
                    "Próximo informe",
                    "Volver a incluir todos los meses con dato válido en el comparativo.",
                    "WES",
                )
            )
        elif "WES" in h.titulo:
            acts.append(
                Accion(
                    "Mantener el nocturno WES como línea base.",
                    "Próximo informe",
                    "Detectar desviaciones sobre la operación habitual.",
                    "WES",
                )
            )
        elif "mercado" in h.titulo.lower():
            acts.append(
                Accion(
                    "Mantener el horario full 22:00–03:00 como línea base.",
                    "Próximo informe",
                    "Detectar desviaciones sobre el patrón continuo del mercado.",
                    "WES",
                )
            )
        elif "Cobertura" in h.titulo:
            acts.append(
                Accion(
                    "Completar cobertura de medición en el punto afectado.",
                    "7 días",
                    "Asegurar serie continua para el próximo periodo.",
                    "WES",
                )
            )
    if len(acts) < 3:
        if cfg.get("cpa_estado") == "instalado_pendiente":
            acts.append(
                Accion(
                    "Programar el CPA con el peak 22:00–03:00.",
                    "7 días",
                    "Que el control no corte agua en el horario de operación.",
                    "WES + operación",
                )
            )
            acts.append(
                Accion(
                    "Mantener observación del caudal 24 h hasta que el CPA esté en servicio.",
                    "Próximo informe",
                    "Línea base previa al control automático.",
                    "WES + cliente",
                )
            )
        elif cfg.get("nocturnal_explain") == "mercado":
            acts.append(
                Accion(
                    "Alertar solo si el caudal 24 h se dispara sin operación de mercado.",
                    "Próximo informe",
                    "No usar 00:00–06:59 como indicador de pérdida en este recinto.",
                    "WES + cliente",
                )
            )
        else:
            acts.append(
                Accion(
                    "Definir línea base nocturna y umbral de alerta.",
                    "Próximo informe",
                    "Detección temprana de desviaciones.",
                    "WES + cliente",
                )
            )
    # unique by accion, max 3
    seen = set()
    out = []
    for a in acts:
        if a.accion in seen:
            continue
        seen.add(a.accion)
        out.append(a)
        if len(out) == 3:
            break
    return out


def _series_diarias(
    cfg: dict, nodos: List[dict], hallazgos: Sequence[Hallazgo], data: dict
) -> List[SerieDiaria]:
    start, end, _dias = _bounds_from_data(data)
    picked: List[dict] = []
    if cfg.get("matriz_id"):
        m = next((n for n in nodos if n["node_id"] == cfg["matriz_id"]), None)
        if m:
            picked.append(m)
    if not picked:
        picked.append(max(nodos, key=lambda n: float(n["total"])))
    for h in hallazgos:
        for n in nodos:
            if n["short_name"] in h.titulo and n not in picked:
                picked.append(n)
                break
    for n in sorted(nodos, key=lambda x: -float(x["total"])):
        if n not in picked:
            picked.append(n)
        if len(picked) >= 3:
            break
    out = []
    for n in picked[:3]:
        serie = _daily_full(n, start, end)
        missing = _cobertura_huecos(n, start, end)
        if n.get("max_fecha"):
            serie.lectura = (
                f"El máximo fue de {_fmt(n['max_m3'], 1)} m³ el {_fecha_es(n['max_fecha'])}. "
            )
        else:
            serie.lectura = "Serie del periodo. "
        if missing:
            serie.lectura += (
                f"Hay {len(missing)} día(s) sin dato; no se interpolan. "
            )
        elif start.day != 1 or cfg.get("excluir_meses_6m"):
            serie.lectura += "La serie cubre los días del periodo informado."
        else:
            serie.lectura += "La serie cubre el mes completo."
        out.append(serie)
    return out


def _visitas_spec(visitas: Sequence[VisitaTecnica]) -> List[VisitaTecnicaSpec]:
    return [
        VisitaTecnicaSpec(
            fecha=v.fecha,
            tecnico=v.tecnico,
            punto=_punto_lectura(v.punto),
            motivo=v.motivo,
            diagnostico=v.diagnostico,
        )
        for v in visitas
    ]


def _punto_lectura(nombre: str) -> str:
    text = (nombre or "—").strip() or "—"
    if text.isupper() and len(text) > 3:
        return text.title()
    return text


def _parrafo_visitas(visitas: Sequence[VisitaTecnica]) -> List[Tuple[str, bool]]:
    if not visitas:
        return []
    if len(visitas) == 1:
        v = visitas[0]
        return [
            ("En el periodo se registró una visita técnica el ", False),
            (_fecha_es(v.fecha_iso), True),
            (f" en {_punto_lectura(v.punto)} (", False),
            (v.motivo, True),
            ("). El detalle está en la sección de visitas técnicas.", False),
        ]
    fechas_u: List[str] = []
    for v in visitas:
        lab = _fecha_es(v.fecha_iso)
        if lab not in fechas_u:
            fechas_u.append(lab)
    if len(fechas_u) == 1:
        fechas = fechas_u[0]
    elif len(fechas_u) == 2:
        fechas = f"{fechas_u[0]} y el {fechas_u[1]}"
    else:
        fechas = ", ".join(fechas_u[:-1]) + f" y el {fechas_u[-1]}"
    puntos: List[str] = []
    for v in visitas:
        p = _punto_lectura(v.punto)
        if p not in puntos:
            puntos.append(p)
    punto_txt = puntos[0] if len(puntos) == 1 else ", ".join(puntos)
    return [
        (f"En el periodo se registraron {len(visitas)} visitas técnicas el ", False),
        (fechas, True),
        (f" en {punto_txt}. El detalle está en la sección de visitas técnicas.", False),
    ]


def build_spec(
    cfg: dict,
    data: dict,
    visitas: Optional[Sequence[VisitaTecnica]] = None,
) -> InformeSpec:
    visitas = list(visitas or [])
    kpi = data["kpi"]
    nodos = data["nodos"]
    entrada = float(kpi["entrada"])
    nocturno = float(kpi["nocturno"])
    pct = float(kpi["pct_nocturno"])
    promedio = float(kpi["promedio"])
    max_m3 = float(kpi["max_m3"])
    max_fecha = kpi["max_fecha"]
    costo = float(kpi["costo_nocturno"])
    hallazgos, evento_fuerte = _hallazgos(cfg, data)
    clasificacion, motivo = _clasificar(cfg, pct, evento_fuerte, salto=_salto_mensual(data))
    acciones = _acciones(hallazgos, cfg)
    entrada_txt = f"{_fmt(entrada, 1)} m³"
    noct_txt = f"{_fmt(nocturno, 1)} m³"
    prom_txt = f"{_fmt(promedio, 1)} m³/día"
    pct_txt = f"{_fmt(round(pct), 0)} %"
    max_txt = f"{_fmt(max_m3, 1)} m³" if max_fecha else "—"
    fecha_max = _fecha_es(max_fecha) if max_fecha else "—"
    _start, _end, dias = _bounds_from_data(data)
    verbo = cfg.get("verbo_registro") or "registró"

    panorama = [
        (f"En {cfg['sitio']} se registraron ", False),
        (entrada_txt, True),
        (", con un promedio de ", False),
        (f"{_fmt(promedio, 1)} m³ diarios", True),
        (f" sobre los {dias} días del periodo. El mayor consumo diario ocurrió el ", False),
        (fecha_max, True),
        (", con ", False),
        (max_txt, True),
        (".", False),
    ]
    lectura = [
        [
            (f"Durante el periodo analizado {cfg['sujeto']} {verbo} ", False),
            (entrada_txt, True),
            (cfg.get("lectura_nocturno") or ". El consumo entre las 00:00 y las 06:59 alcanzó ", False),
            (noct_txt, True),
            (", equivalente al ", False),
            (pct_txt, True),
            (" del volumen de referencia y a un costo referencial de ", False),
            (_fmt_clp(costo), True),
            (".", False),
        ],
        [
            ("El estado se clasifica como ", False),
            (clasificacion.title().replace("ó", "ó"), True),
            (f". {motivo[0].upper() + motivo[1:]}" if motivo else "", False),
        ],
    ]
    # Fix clasificacion display in lectura: "Bajo control" etc.
    lectura[1] = [
        ("El estado se clasifica como ", False),
        (
            {
                "BAJO CONTROL": "Bajo control",
                "EN OBSERVACIÓN": "En observación",
                "REQUIERE ATENCIÓN": "Requiere atención",
                "CRÍTICO": "Crítico",
            }[clasificacion],
            True,
        ),
        (". ", False),
        (motivo[0].upper() + motivo[1:] if motivo else "", False),
    ]
    extra_visitas = _parrafo_visitas(visitas)
    if extra_visitas:
        lectura.append(extra_visitas)

    conclusion = [
        [
            ("El estado se clasifica como ", False),
            (
                f"“{_pretty_cls(clasificacion)}”",
                True,
            ),
            (". ", False),
            (motivo[0].upper() + motivo[1:] if motivo else "", False),
        ],
        [
            (
                (
                    "Cuando el CPA esté activo y programado se reevaluará a En observación o Bajo control. "
                    "Si el caudal 24 h se dispara sin operación de mercado, la clasificación avanzará a "
                    if cfg.get("cpa_estado") == "instalado_pendiente"
                    else "Si el caudal 24 h se dispara sin operación de mercado, la clasificación avanzará a "
                    if cfg.get("nocturnal_explain") == "mercado"
                    else "Si el consumo nocturno se desvía de esta referencia, aumenta o presenta "
                    "eventos sin explicación operacional, la clasificación avanzará a "
                ),
                False,
            ),
            (
                "“Crítico”"
                if clasificacion == "REQUIERE ATENCIÓN"
                else "“Requiere atención”",
                True,
            ),
            (".", False),
        ],
    ]

    indicadores = []
    orden = sorted(nodos, key=lambda n: -float(n["total"]))
    if cfg.get("matriz_id") and not cfg["additive"]:
        # matriz first then rest by total
        matriz = next(n for n in nodos if n["node_id"] == cfg["matriz_id"])
        resto = [n for n in orden if n["node_id"] != cfg["matriz_id"]]
        orden = [matriz] + resto
    for n in orden:
        max_dt = datetime.strptime(n["max_fecha"], "%Y-%m-%d") if n.get("max_fecha") else None
        indicadores.append(
            PuntoIndicador(
                nombre=n["short_name"],
                total=float(n["total"]),
                promedio=float(n["total"]) / dias,
                max_m3=float(n["max_m3"]),
                max_fecha=max_dt.strftime("%d/%m") if max_dt else "—",
                nocturno=float(n["nocturno_m3"]),
                cobertura=int(n["nocturno_cobertura"]),
                es_matriz=n["node_id"] == cfg.get("matriz_id"),
            )
        )

    labels_6, vals_6 = [], []
    for item in data["serie_6_meses"]:
        lab = item["label"].replace("*", "").split()[0].capitalize()
        labels_6.append(lab)
        vals_6.append(float(item["m3"]))

    out_dir = Path("reports") / cfg["folder"] / "GESTION_HIDRICA"
    charts = out_dir / "_charts"
    charts.mkdir(parents=True, exist_ok=True)
    highlight = cfg.get("matriz_name") or ""
    names = [n["short_name"] for n in (orden if not cfg["additive"] else sorted(nodos, key=lambda x: -x["total"]))]
    # keep original order for matriz sites: as in config node_ids
    if not cfg["additive"]:
        names = [n["short_name"] for n in nodos]
        totals = [float(n["total"]) for n in nodos]
        nocts = [float(n["nocturno_m3"]) for n in nodos]
    else:
        by_tot = sorted(nodos, key=lambda x: -float(x["total"]))
        names = [n["short_name"] for n in by_tot]
        totals = [float(n["total"]) for n in by_tot]
        nocts = [float(n["nocturno_m3"]) for n in by_tot]

    chart_6m = build_chart_6_meses(charts / f"{cfg['key']}_6m.png", labels_6, vals_6)
    chart_pts = build_chart_puntos(
        charts / f"{cfg['key']}_puntos.png",
        names,
        totals,
        highlight,
        additive=bool(cfg.get("additive")),
    )
    chart_noc = build_chart_nocturno(
        charts / f"{cfg['key']}_nocturno.png",
        names,
        nocts,
        highlight,
        cfg.get("leyenda"),
    )
    series = _series_diarias(cfg, nodos, hallazgos, data)

    return InformeSpec(
        cliente=cfg["cliente"],
        sitio=cfg["sitio"],
        periodo_corto=cfg.get("periodo_corto")
        or f"{cfg['sitio']} · 1 al 31 de agosto de 2026",
        footer=f"Informe mensual - {cfg['cliente']} | Agosto 2026",
        titulo_onepager="Resumen ejecutivo de gestión hídrica",
        titulo_mensual="Informe mensual de gestión hídrica",
        clasificacion=clasificacion,
        motivo=motivo,
        kpi_entrada=entrada_txt,
        kpi_promedio=prom_txt,
        kpi_nocturno=noct_txt,
        kpi_pct=pct_txt,
        panorama=panorama,
        panorama_nota=cfg.get("panorama_nota")
        or "Agosto se evalúa del 1 al 31 y no se extrapola.",
        hallazgos=hallazgos,
        acciones=acciones,
        conclusion=conclusion,
        lectura_ejecutiva=lectura,
        nota_agosto=cfg.get("nota_agosto")
        or "Agosto comprende 31 días. No se extrapola el consumo.",
        kpi_consumo_label=cfg.get("kpi_label") or "Consumo de entrada",
        chart_6m=chart_6m,
        chart_puntos=chart_pts,
        chart_puntos_nota=cfg["chart_nota"],
        max_entrada_txt=(
            f"El mayor consumo diario de la referencia ocurrió el {fecha_max}, con {max_txt}."
        ),
        chart_nocturno=chart_noc,
        chart_nocturno_nota=cfg["nocturno_nota"],
        indicadores=indicadores,
        criterio_nocturno=[
            [
                (
                    (
                        cfg.get("ventana_nocturna")
                        or (
                            "Se considera nocturno el volumen medido entre las 00:00 y las 06:59, "
                            "hora de Chile. "
                        )
                    )
                    + "Los valores corresponden únicamente a días con datos y no "
                    "se proyectan. El costo nocturno de la referencia se estima en ",
                    False,
                ),
                (_fmt_clp(costo), True),
                (", con tarifa referencial de ", False),
                (f"{_fmt_clp(float(data['price_per_m3']))}/m³", True),
                (".", False),
            ]
        ],
        nota_cobertura=(
            "La cobertura nocturna indica cuántos días cuentan con registros en esa franja. "
            "Los días sin datos no se interpolan."
        ),
        series_diarias=series,
        logo_path=resolve_logo(),
        visitas=_visitas_spec(visitas),
    )


def _pretty_cls(c: str) -> str:
    return {
        "BAJO CONTROL": "Bajo control",
        "EN OBSERVACIÓN": "En observación",
        "REQUIERE ATENCIÓN": "Requiere atención",
        "CRÍTICO": "Crítico",
    }[c]


def generar_cliente(
    cfg: dict,
    visitas: Optional[Sequence[VisitaTecnica]] = None,
) -> Tuple[Path, Path]:
    data = fetch_cliente(cfg)
    spec = build_spec(cfg, data, visitas)
    out_dir = Path("reports") / cfg["folder"] / "GESTION_HIDRICA"
    slug = cfg["cliente"].replace(" ", "_").replace("Á", "A").replace("á", "a")
    one = out_dir / f"One_Pager_Gestion_Hidrica_{slug}_Agosto_2026.pdf"
    monthly = out_dir / f"Informe_Mensual_{slug}_Agosto_2026.pdf"
    render_one_pager(spec, one)
    render_mensual(spec, monthly, out_dir / "_charts")
    if spec.visitas:
        print(f"  visitas: {len(spec.visitas)}", flush=True)
    return one, monthly


def run_lote(clientes: Sequence[dict], titulo: str) -> None:
    print(f"{titulo}\n", flush=True)
    try:
        todas = cargar_visitas_periodo(START_DT, END_DT)
        print(f"[INFO] Visitas del formulario en el periodo: {len(todas)}", flush=True)
    except Exception as e:
        todas = []
        print(f"[ADVERTENCIA] No se pudieron leer visitas técnicas: {e}", flush=True)
    ok = []
    errors = []
    for cfg in clientes:
        try:
            visitas = visitas_de_cliente(todas, cfg)
            one, monthly = generar_cliente(cfg, visitas)
            print(f"[OK] {cfg['cliente']}: {one.name} | {monthly.name}", flush=True)
            ok.append((cfg, one, monthly))
        except Exception as e:
            errors.append(f"{cfg['cliente']}: {e}")
            print(f"[ERROR] {cfg['cliente']}: {e}", flush=True)
            import traceback

            traceback.print_exc()
    print(f"\n[INFO] Completados: {len(ok)}/{len(clientes)}", flush=True)
    if errors:
        print("[INFO] Fallidos:")
        for e in errors:
            print("  -", e)
        raise SystemExit(1)


def main() -> None:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
            sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
        except Exception:
            pass
    run_lote(
        CLIENTES,
        "GESTIÓN HÍDRICA — lote 8 — 01/08/2026 a 31/08/2026",
    )


if __name__ == "__main__":
    main()
