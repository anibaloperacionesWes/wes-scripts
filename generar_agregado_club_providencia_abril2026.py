"""Agregado Club Providencia — 01/04/2026 a 30/04/2026 (solo Word, rápido).

Por defecto:
- Sin PPT (más rápido).
- Descarga en paralelo por nodo (ajustable con --workers).

Uso:
  python generar_agregado_club_providencia_abril2026.py
  python generar_agregado_club_providencia_abril2026.py --start-date 01/04/2026 --end-date 30/04/2026 --workers 4
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


COMPANY_ID = "000031"  # Club Providencia
NODE_IDS = ["000031-01", "000031-02"]  # Matriz Fitness, Matriz Piscina

DEFAULT_START = "01/04/2026"
DEFAULT_END = "30/04/2026"


def _default_workers() -> int:
    # Es un agregado pequeño (2 nodos). 2 workers es suficiente.
    cpu = os.cpu_count() or 4
    return max(2, min(cpu, len(NODE_IDS), 4))


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Reporte agregado Club Providencia (rápido, sin PPT)."
    )
    ap.add_argument("--start-date", default=DEFAULT_START, help="Inicio (dd/mm/aaaa o ISO).")
    ap.add_argument("--end-date", default=DEFAULT_END, help="Fin (dd/mm/aaaa o ISO).")
    ap.add_argument("--workers", type=int, default=None, help="Workers descarga paralela.")
    ap.add_argument("--output-dir", default="reports", help="Directorio base de salida.")
    args = ap.parse_args()

    workers = args.workers if args.workers is not None else _default_workers()
    workers = max(1, min(int(workers), len(NODE_IDS)))

    print("[INFO] Club Providencia — reporte agregado")
    print(f"[INFO] Periodo: {args.start_date} → {args.end_date}")
    print(f"[INFO] Nodos: {len(NODE_IDS)} | workers={workers}")
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
        max_parallel_workers=workers,
    )
    elapsed = time.perf_counter() - t0
    print(f"[OK] {out}")
    print(f"[INFO] Tiempo total agregado: {elapsed:.1f} s")


if __name__ == "__main__":
    main()

