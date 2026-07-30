"""
Reportes agregados junio 2026:
  Club Providencia, Colegios Providencia, UDD, CDUC,
  Renca (Gimnasio + Piscina), Fundo Zapallar.

Uso:
  python generar_agregados_junio2026_clientes_batch.py
"""

from __future__ import annotations

import sys
import time
from typing import List, Optional

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from exclusiones_reportes import FUNDO_ZAPALLAR_NODE_IDS
from generar_reporte_word import generate_aggregated_report

PERIODO = ("01/06/2026", "30/06/2026")

NODOS_CDUC = [
    "000021-01",
    "000021-02",
    "000021-03",
    "000021-04",
    "000021-05",
    "000021-07",
    "000021-08",
]


def _run(
    label: str,
    company_id: str,
    node_ids: List[str],
    *,
    output_dir: str = "reports",
    apply_exclusions: bool = True,
    fuente_agua_id: Optional[str] = None,
    workers: Optional[int] = None,
) -> None:
    w = workers if workers is not None else max(4, min(8, len(node_ids)))
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
        max_parallel_workers=w,
        fuente_agua_id=fuente_agua_id,
    )
    print(f"[OK] {label}: {out}")
    print(f"[INFO] Tiempo: {time.perf_counter() - t0:.1f} s\n")


def main() -> None:
    jobs = [
        (
            "Club Providencia",
            "000031",
            ["000031-01", "000031-02"],
            dict(apply_exclusions=False, workers=4),
        ),
        (
            "Colegios Providencia",
            "000006",
            ["000006-01", "000006-02", "000006-04", "000006-05"],
            dict(apply_exclusions=False, workers=4),
        ),
        (
            "UDD",
            "000026",
            ["000026-01", "000026-02"],
            dict(apply_exclusions=False, workers=4),
        ),
        (
            "CDUC",
            "000021",
            list(NODOS_CDUC),
            dict(apply_exclusions=True, workers=8),
        ),
        (
            "Renca — Gimnasio y Piscina",
            "000017",
            ["000017-05", "000017-06"],
            dict(apply_exclusions=False, workers=4),
        ),
        (
            "Fundo Zapallar",
            "000027",
            list(FUNDO_ZAPALLAR_NODE_IDS),
            dict(apply_exclusions=False, fuente_agua_id="000027-01", workers=8),
        ),
    ]

    print("GENERACIÓN AGREGADOS JUNIO 2026\n")
    for label, cid, nids, kwargs in jobs:
        if not nids:
            print(f"[ERROR] {label}: sin nodos.\n")
            continue
        try:
            _run(label, cid, nids, **kwargs)
        except Exception as e:
            print(f"[ERROR] {label}: {e}\n")
            import traceback

            traceback.print_exc()


if __name__ == "__main__":
    main()
