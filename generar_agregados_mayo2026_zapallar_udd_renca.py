"""Agregados mayo 2026: Fundo Zapallar, UDD y Renca."""

from __future__ import annotations

import time

import requests

from exclusiones_reportes import FUNDO_ZAPALLAR_NODE_IDS, filter_node_ids
from generar_reporte_word import generate_aggregated_report

ENTITY = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
PERIODO = ("01/05/2026", "31/05/2026")


def _nodos_renca() -> list[str]:
    r = requests.get(f"{ENTITY}/companies/000017", timeout=30)
    r.raise_for_status()
    ids = [n["nodeId"] for n in r.json().get("nodes", []) if n.get("nodeId")]
    return filter_node_ids(ids, company_id="000017", company_name="Renca")


def main() -> None:
    jobs = [
        ("000027", list(FUNDO_ZAPALLAR_NODE_IDS), "Fundo Zapallar", False, "000027-01"),
        ("000026", ["000026-01", "000026-02"], "UDD", False, None),
        ("000017", _nodos_renca(), "Renca", True, None),
    ]
    for cid, nids, label, apply_ex, fuente in jobs:
        print("=" * 60)
        print(f"[INFO] {label} — {len(nids)} nodo(s): {', '.join(nids)}")
        t0 = time.perf_counter()
        out = generate_aggregated_report(
            company_id=cid,
            node_ids=nids,
            start_date=PERIODO[0],
            end_date=PERIODO[1],
            output_dir="reports",
            apply_exclusions=apply_ex,
            generate_ppt=False,
            parallel_node_fetch=True,
            max_parallel_workers=8,
            fuente_agua_id=fuente,
        )
        print(f"[OK] {label}: {out}")
        print(f"[INFO] Tiempo: {time.perf_counter() - t0:.1f} s")


if __name__ == "__main__":
    main()
