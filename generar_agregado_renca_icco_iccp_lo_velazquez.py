"""Agregado Renca — ICCO, ICCP y Lo Velázquez."""

from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

if sys.platform == "win32":
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

from generar_reporte_word import generate_aggregated_report

COMPANY_ID = "000017"
NODE_IDS = ["000017-04", "000017-07", "000017-08"]  # Lo Velázquez, ICCP, ICCO
DEFAULT_START = "01/06/2026"
DEFAULT_END = "30/06/2026"


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser(description="Agregado Renca — ICCO, ICCP, Lo Velázquez")
    ap.add_argument("--start-date", default=DEFAULT_START)
    ap.add_argument("--end-date", default=DEFAULT_END)
    args = ap.parse_args()

    print("=" * 72)
    print("[INFO] Renca — agregado ICCO + ICCP + Lo Velázquez")
    print("[INFO] Nodos:", ", ".join(NODE_IDS))
    print(f"[INFO] Periodo: {args.start_date} -> {args.end_date}")
    t0 = time.perf_counter()
    out = generate_aggregated_report(
        company_id=COMPANY_ID,
        node_ids=NODE_IDS,
        start_date=args.start_date,
        end_date=args.end_date,
        output_dir="reports",
        apply_exclusions=False,
        generate_ppt=False,
        parallel_node_fetch=True,
        max_parallel_workers=6,
    )
    print(f"[OK] {out}")
    print(f"[INFO] Tiempo: {time.perf_counter() - t0:.1f} s")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
