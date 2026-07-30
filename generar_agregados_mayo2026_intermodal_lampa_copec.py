"""Agregados mayo 2026: AGUNSA Intermodal, AGUNSA Lampa y COPEC."""

from __future__ import annotations

import time

import requests

from generar_reporte_word import generate_aggregated_report

ENTITY = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
PERIODO = ("01/05/2026", "31/05/2026")

# Lampa: depósito + módulos en planta Lampa (sin Intermodal San Antonio).
NODOS_LAMPA = ["000020-01", "000020-02", "000020-03", "000020-04"]
NODO_INTERMODAL = ["000020-05"]


def _nodos_copec() -> list[str]:
    r = requests.get(f"{ENTITY}/companies/000009", timeout=30)
    r.raise_for_status()
    ids = [n["nodeId"] for n in r.json().get("nodes", []) if n.get("nodeId")]
    return sorted(ids)


def _generar(label: str, cid: str, nids: list[str]) -> None:
    print("=" * 60)
    print(f"[INFO] {label} — {len(nids)} nodo(s): {', '.join(nids)}")
    t0 = time.perf_counter()
    out = generate_aggregated_report(
        company_id=cid,
        node_ids=nids,
        start_date=PERIODO[0],
        end_date=PERIODO[1],
        output_dir="reports",
        apply_exclusions=False,
        generate_ppt=False,
        parallel_node_fetch=True,
        max_parallel_workers=8,
    )
    print(f"[OK] {label}: {out}")
    print(f"[INFO] Tiempo: {time.perf_counter() - t0:.1f} s\n")


def main() -> None:
    _generar("AGUNSA Intermodal San Antonio", "000020", NODO_INTERMODAL)
    _generar("AGUNSA Lampa (Depósito + módulos)", "000020", NODOS_LAMPA)
    copec = _nodos_copec()
    _generar("COPEC", "000009", copec)


if __name__ == "__main__":
    main()
