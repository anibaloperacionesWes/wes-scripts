"""Descarga métricas Inchcape Quilicura (01–28/08/2026) para informes de gestión hídrica."""

from __future__ import annotations

import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

from agregado_extendido_extra import INCHCAPE_NODE_MATRIZ_PRINCIPAL, _serie_mensual_nodo
from generar_reporte_word import (
    acl_node_base_url,
    calculate_nocturnal_metrics,
    fetch_json,
    flatten_measures,
    format_number_chilean,
    get_node_name,
    get_water_price_per_m3,
    normalize_measures_payload,
    parse_date,
    summarize_consumption,
)

COMPANY_ID = "000012"
START = "01/08/2026"
END = "28/08/2026"
NODE_IDS = [
    "000012-06",
    "000012-07",
    "000012-08",
    "000012-09",
    "000012-10",
    "000012-11",
    "000012-12",
]
SHORT_NAMES = {
    "000012-06": "Matriz Principal",
    "000012-07": "Dercomaq",
    "000012-08": "Lav. Máquinas",
    "000012-09": "Casino",
    "000012-10": "Proderco",
    "000012-11": "Camarines",
    "000012-12": "Edificio JCB",
}
OUT = Path("/tmp/inchcape_gh/datos.json")


def _daily_series(measures) -> List[Dict[str, Any]]:
    by_day: Dict[str, float] = {}
    for m in measures:
        key = m.date.strftime("%Y-%m-%d")
        by_day[key] = by_day.get(key, 0.0) + float(m.total_m3)
    return [{"date": d, "m3": by_day[d]} for d in sorted(by_day)]


def _fetch_node(node_id: str, start_dt: datetime, end_dt: datetime) -> Dict[str, Any]:
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
    noct = calculate_nocturnal_metrics(node_id, start_dt, end_dt, company_id=COMPANY_ID)
    mx = summary.get("max")
    return {
        "node_id": node_id,
        "node_name": get_node_name(node_id),
        "short_name": SHORT_NAMES.get(node_id, get_node_name(node_id)),
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


def main() -> None:
    start_dt = parse_date(START)
    end_dt = parse_date(END, end_of_day=True)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    results: Dict[str, Dict[str, Any]] = {}
    print(f"[INFO] Fetch Inchcape {START}–{END}", flush=True)
    with ThreadPoolExecutor(max_workers=6) as ex:
        futs = {ex.submit(_fetch_node, nid, start_dt, end_dt): nid for nid in NODE_IDS}
        for fut in as_completed(futs):
            nid = futs[fut]
            try:
                results[nid] = fut.result()
                r = results[nid]
                print(
                    f"  {nid} {r['short_name']}: {format_number_chilean(r['total'], 1)} m³ "
                    f"noct {format_number_chilean(r['nocturno_m3'], 1)} cob={r['nocturno_cobertura']}",
                    flush=True,
                )
            except Exception as e:
                print(f"[ERROR] {nid}: {e}", flush=True)
                raise
    ordered = [results[nid] for nid in NODE_IDS if nid in results]
    matriz = results[INCHCAPE_NODE_MATRIZ_PRINCIPAL]
    price = get_water_price_per_m3(COMPANY_ID, INCHCAPE_NODE_MATRIZ_PRINCIPAL, {})
    serie6 = _serie_mensual_nodo(INCHCAPE_NODE_MATRIZ_PRINCIPAL, end_dt, 6)
    payload = {
        "company_id": COMPANY_ID,
        "cliente": "Inchcape",
        "sitio": "Inchcape Quilicura",
        "periodo": {"start": START, "end": END, "dias": 28, "mes_label": "Agosto 2026"},
        "price_per_m3": price,
        "matriz_id": INCHCAPE_NODE_MATRIZ_PRINCIPAL,
        "nodos": ordered,
        "serie_6_meses": [{"label": a, "m3": b} for a, b in serie6],
        "kpi": {
            "entrada": matriz["total"],
            "promedio": matriz["promedio_diario"],
            "nocturno": matriz["nocturno_m3"],
            "pct_nocturno": (matriz["nocturno_m3"] / matriz["total"] * 100.0)
            if matriz["total"]
            else 0.0,
            "max_m3": matriz["max_m3"],
            "max_fecha": matriz["max_fecha"],
            "costo_nocturno": matriz["nocturno_m3"] * price,
        },
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OK] {OUT}", flush=True)
    print(json.dumps(payload["kpi"], ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
