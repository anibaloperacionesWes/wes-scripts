"""Agregado CORMUP (Peñalolén) — solo Word; orientado a velocidad.

Por defecto:
- Sin PPT (generate_ppt=False).
- Descarga medidas/alertas en paralelo; más workers = menos tiempo si la API lo aguanta.
  Si ves errores o timeouts, baja --workers (p. ej. 4).

Uso:
  python generar_agregado_cormup_abril2026.py
  python generar_agregado_cormup_abril2026.py --start-date 01/04/2026 --end-date 30/04/2026
  python generar_agregado_cormup_abril2026.py --workers 14
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

COMPANY_ID = "000008"
# 14 establecimientos CORMUP mapeados en generar_reporte_word
NODE_IDS = [f"000008-{i:02d}" for i in range(1, 15)]

DEFAULT_START = "01/04/2026"
DEFAULT_END = "30/04/2026"


def _default_workers() -> int:
    """Hasta todos los nodos en paralelo, acotado por CPUs (evita spawns absurdos)."""
    n_nodos = len(NODE_IDS)
    cpu = os.cpu_count() or 8
    return max(6, min(n_nodos, cpu, 14))


def main() -> None:
    ap = argparse.ArgumentParser(description="Reporte agregado CORMUP / Peñalolén (rápido, sin PPT).")
    ap.add_argument("--start-date", default=DEFAULT_START, help="Inicio dd/mm/aaaa o ISO.")
    ap.add_argument("--end-date", default=DEFAULT_END, help="Fin dd/mm/aaaa o ISO.")
    ap.add_argument(
        "--workers",
        type=int,
        default=None,
        help=f"Workers descarga paralela (default: {_default_workers()}, acotado al nº de nodos).",
    )
    ap.add_argument("--output-dir", default="reports", help="Directorio base de salida.")
    args = ap.parse_args()

    w = args.workers if args.workers is not None else _default_workers()
    w = max(1, min(int(w), len(NODE_IDS)))

    print("[INFO] Modo rápido: sin PPT, paralelo nodos.")
    print(f"[INFO] Periodo: {args.start_date} → {args.end_date}")
    print(f"[INFO] Nodos: {len(NODE_IDS)} | workers={w}")
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
    elapsed = time.perf_counter() - t0
    print(f"[OK] {out}")
    print(f"[INFO] Tiempo total agregado: {elapsed:.1f} s")


if __name__ == "__main__":
    main()
