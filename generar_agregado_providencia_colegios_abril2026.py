"""Agregado Providencia — colegios (sin 000006-03 Arturo Alessandri Palma), abril 2026.

Nodos: Liceo Lastarria, Carmela Carvajal, Liceo 7, Liceo Juan Pablo Duarte.
"""
from __future__ import annotations

import argparse
import os
import sys
import time

from generar_reporte_word import generate_aggregated_report

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

COMPANY_ID = "000006"
# Colegios Providencia reportados; 000006-03 excluido por pedido operativo.
NODE_IDS = ["000006-01", "000006-02", "000006-04", "000006-05"]
DEFAULT_START = "01/04/2026"
DEFAULT_END = "30/04/2026"


def _default_workers() -> int:
    # Para este agregado (4 nodos) el máximo útil es 4.
    return max(2, min(len(NODE_IDS), os.cpu_count() or 4, 4))


def main() -> None:
    ap = argparse.ArgumentParser(description="Reporte agregado Providencia (colegios).")
    ap.add_argument("--start-date", default=DEFAULT_START)
    ap.add_argument("--end-date", default=DEFAULT_END)
    ap.add_argument("--workers", type=int, default=None)
    ap.add_argument("--output-dir", default="reports")
    args = ap.parse_args()
    w = args.workers if args.workers is not None else _default_workers()
    w = max(1, min(int(w), len(NODE_IDS)))

    print("[INFO] Agregado Providencia (colegios)")
    print(f"[INFO] Periodo: {args.start_date} → {args.end_date}")
    print(f"[INFO] Nodos: {len(NODE_IDS)} | workers={w}")
    print("[INFO] Modo rápido: sin PPT, descarga paralela por nodo")
    t0 = time.perf_counter()
    out = generate_aggregated_report(
        company_id=COMPANY_ID,
        node_ids=list(NODE_IDS),
        start_date=args.start_date,
        end_date=args.end_date,
        output_dir=args.output_dir,
        apply_exclusions=False,
        generate_ppt=False,
        parallel_node_fetch=True,
        max_parallel_workers=w,
    )
    print(f"[OK] {out}")
    print(f"[INFO] Tiempo: {time.perf_counter() - t0:.1f} s")


if __name__ == "__main__":
    main()
