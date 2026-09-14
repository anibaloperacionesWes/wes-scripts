"""
Agregados agosto 2026 — mismo lote de fin de mes que julio:

  Fundo Zapallar, Inchcape (ex DERCO), Nido de Águilas, Lo Valledor,
  UDD, Club Providencia, AGUNSA (Lampa + Intermodal).

Formato: generate_aggregated_report, sin PPT, fetch paralelo.

Uso:
  python generar_agregados_agosto2026_lote.py
"""

from __future__ import annotations

import sys
import time
from typing import List, Optional

import requests

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
    except Exception:
        pass

from exclusiones_reportes import FUNDO_ZAPALLAR_NODE_IDS, filter_node_ids
from generar_reporte_word import generate_aggregated_report

ENTITY = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
# Agosto 2026 hasta el día de corrida (mes en curso)
PERIODO = ("01/08/2026", "28/08/2026")

NODOS_LAMPA = ["000020-01", "000020-02", "000020-03", "000020-04"]
NODO_INTERMODAL = ["000020-05"]
NODOS_CLUB_PROVIDENCIA = ["000031-01", "000031-02"]
NODOS_UDD = ["000026-01", "000026-02"]
# Lo Valledor: P1 + Barrio Norte (000002-02 excluido por configuración)
NODOS_LO_VALLEDOR = ["000002-01", "000002-03"]


def _nodos_api(company_id: str, company_name: str = "", *, apply_exclusions: bool) -> List[str]:
    r = requests.get(f"{ENTITY}/companies/{company_id}", timeout=30)
    r.raise_for_status()
    ids = sorted(n["nodeId"] for n in r.json().get("nodes", []) if n.get("nodeId"))
    if apply_exclusions:
        return filter_node_ids(ids, company_id=company_id, company_name=company_name)
    return ids


def _run(
    label: str,
    company_id: str,
    node_ids: List[str],
    *,
    apply_exclusions: bool = True,
    fuente_agua_id: Optional[str] = None,
    company_folder_override: Optional[str] = None,
) -> None:
    if not node_ids:
        print(f"[ERROR] {label}: sin nodos.\n")
        return
    w = max(4, min(8, len(node_ids)))
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
        output_dir="reports",
        apply_exclusions=apply_exclusions,
        generate_ppt=False,
        parallel_node_fetch=True,
        max_parallel_workers=w,
        fuente_agua_id=fuente_agua_id,
        company_folder_override=company_folder_override,
    )
    print(f"[OK] {label}: {out}")
    print(f"[INFO] Tiempo: {time.perf_counter() - t0:.1f} s\n")


def main() -> None:
    jobs = [
        (
            "Fundo Zapallar",
            "000027",
            list(FUNDO_ZAPALLAR_NODE_IDS),
            dict(apply_exclusions=False, fuente_agua_id="000027-01"),
        ),
        (
            "Inchcape (ex DERCO)",
            "000012",
            _nodos_api("000012", "Inchcape", apply_exclusions=False),
            dict(apply_exclusions=True, company_folder_override="Inchcape"),
        ),
        (
            "Nido de Aguilas",
            "000007",
            _nodos_api("000007", "Nido de Aguilas", apply_exclusions=False),
            dict(apply_exclusions=True),
        ),
        (
            "Lo Valledor (P1 + Barrio Norte)",
            "000002",
            list(NODOS_LO_VALLEDOR),
            dict(apply_exclusions=False, company_folder_override="Lo_Valledor"),
        ),
        (
            "UDD",
            "000026",
            list(NODOS_UDD),
            dict(apply_exclusions=False),
        ),
        (
            "Club Providencia",
            "000031",
            list(NODOS_CLUB_PROVIDENCIA),
            dict(apply_exclusions=False),
        ),
        (
            "AGUNSA Lampa (Depósito + módulos)",
            "000020",
            list(NODOS_LAMPA),
            dict(apply_exclusions=False, company_folder_override="AGUNSA_Lampa"),
        ),
        (
            "AGUNSA Intermodal San Antonio",
            "000020",
            list(NODO_INTERMODAL),
            dict(apply_exclusions=False, company_folder_override="Agunsa_Intermodal"),
        ),
    ]

    print("GENERACIÓN AGREGADOS AGOSTO 2026 — LOTE FIN DE MES\n")
    print(f"Periodo: {PERIODO[0]} – {PERIODO[1]}\n")
    ok = 0
    errors: list[str] = []
    for label, cid, nids, kwargs in jobs:
        try:
            _run(label, cid, nids, **kwargs)
            ok += 1
        except Exception as e:
            errors.append(f"{label}: {e}")
            print(f"[ERROR] {label}: {e}\n")
            import traceback

            traceback.print_exc()
    print(f"[INFO] Completados: {ok}/{len(jobs)}")
    if errors:
        print("[INFO] Fallidos:")
        for e in errors:
            print(f"  - {e}")


if __name__ == "__main__":
    main()
