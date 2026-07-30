"""
Reportes agregados mayo 2026:
  Fundo Zapallar, DERCO, Nido de Aguilas, Lo Valledor (P1), Barrio Norte, UDD, Club Providencia.

Uso:
  python generar_agregados_mayo2026_clientes.py
"""

from __future__ import annotations

import sys
import time
from typing import List, Optional

import requests

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from exclusiones_reportes import FUNDO_ZAPALLAR_NODE_IDS, filter_node_ids
from generar_reporte_word import generate_aggregated_report

ENTITY = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
PERIODO = ("01/05/2026", "31/05/2026")


def _nodos_empresa(company_id: str, company_name: str = "") -> List[str]:
    r = requests.get(f"{ENTITY}/companies/{company_id}", timeout=30)
    r.raise_for_status()
    ids = [n["nodeId"] for n in r.json().get("nodes", []) if n.get("nodeId")]
    return filter_node_ids(ids, company_id=company_id, company_name=company_name)


def _run(
    label: str,
    company_id: str,
    node_ids: List[str],
    *,
    output_dir: str = "reports",
    apply_exclusions: bool = True,
    fuente_agua_id: Optional[str] = None,
) -> None:
    print("=" * 70)
    print(f"[INFO] {label}")
    print(f"       Empresa {company_id} | {len(node_ids)} nodo(s): {', '.join(node_ids)}")
    print(f"       Periodo: {PERIODO[0]} – {PERIODO[1]}")
    t0 = time.perf_counter()
    out = generate_aggregated_report(
        company_id=company_id,
        node_ids=node_ids,
        start_date=PERIODO[0],
        end_date=PERIODO[1],
        output_dir=output_dir,
        apply_exclusions=apply_exclusions,
        generate_ppt=False,
        parallel_node_fetch=True,
        max_parallel_workers=8,
        fuente_agua_id=fuente_agua_id,
    )
    print(f"[OK] {label}: {out}")
    print(f"[INFO] Tiempo: {time.perf_counter() - t0:.1f} s\n")


def main() -> None:
    jobs = [
        (
            "Fundo Zapallar",
            "000027",
            list(FUNDO_ZAPALLAR_NODE_IDS),
            dict(output_dir="reports", apply_exclusions=False, fuente_agua_id="000027-01"),
        ),
        (
            "DERCO",
            "000012",
            _nodos_empresa("000012", "DERCO"),
            dict(apply_exclusions=True),
        ),
        (
            "Nido de Aguilas",
            "000007",
            _nodos_empresa("000007", "Nido de Aguilas"),
            dict(apply_exclusions=True),
        ),
        (
            "Lo Valledor — P1",
            "000002",
            ["000002-01"],
            dict(output_dir="reports/Lo_Valledor", apply_exclusions=True),
        ),
        (
            "Lo Valledor — Barrio Norte",
            "000002",
            ["000002-03"],
            dict(output_dir="reports/Barrio_Norte", apply_exclusions=True),
        ),
        (
            "UDD",
            "000026",
            ["000026-01", "000026-02"],
            dict(apply_exclusions=False),
        ),
        (
            "Club Providencia",
            "000031",
            ["000031-01", "000031-02"],
            dict(apply_exclusions=False),
        ),
    ]

    print("GENERACIÓN AGREGADOS MAYO 2026\n")
    for label, cid, nids, kwargs in jobs:
        if not nids:
            print(f"[ERROR] {label}: sin nodos tras exclusiones.\n")
            continue
        try:
            _run(label, cid, nids, **kwargs)
        except Exception as e:
            print(f"[ERROR] {label}: {e}\n")
            import traceback

            traceback.print_exc()


if __name__ == "__main__":
    main()
