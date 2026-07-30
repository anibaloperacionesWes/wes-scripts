"""
Genera reportes Word por cada punto AGUNSA y, solo después, el reporte agregado.
El agregado no se ejecuta antes de terminar los individuales (orden garantizado).

Uso:
  python generar_reportes_agunsa_periodo.py
  python generar_reportes_agunsa_periodo.py --start-date 01/01/2026 --end-date 06/04/2026
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from generar_reporte_word import (
    generate_aggregated_report,
    generate_report,
    get_company_name,
)

COMPANY_ID = "000020"
NODOS_AGUNSA = [
    "000020-01",
    "000020-02",
    "000020-03",
    "000020-04",
    "000020-05",
]


def main() -> int:
    p = argparse.ArgumentParser(description="Reportes AGUNSA: individuales y luego agregado")
    p.add_argument("--start-date", default="01/01/2026", help="Inicio (dd/mm/aaaa)")
    p.add_argument("--end-date", default="06/04/2026", help="Término (dd/mm/aaaa)")
    p.add_argument(
        "--parallel-agregado",
        action="store_true",
        help="Descarga en paralelo al armar el agregado (solo la fase final).",
    )
    p.add_argument(
        "--workers-agregado",
        type=int,
        default=4,
        metavar="N",
        help="Hilos máx. para --parallel-agregado (default: 4).",
    )
    args = p.parse_args()

    root = Path(__file__).resolve().parent
    reports_dir = root / "reports"
    company_name = get_company_name(COMPANY_ID)

    print("=" * 70)
    print("AGUNSA — individuales primero, agregado al final")
    print(f"Empresa: {company_name} ({COMPANY_ID})")
    print(f"Periodo: {args.start_date} -> {args.end_date}")
    print(f"Nodos: {', '.join(NODOS_AGUNSA)}")
    print("=" * 70)

    exitosos: list[str] = []
    for i, nid in enumerate(NODOS_AGUNSA, 1):
        print(f"[{i}/{len(NODOS_AGUNSA)}] Generando {nid}...")
        gen_args = argparse.Namespace(
            company_id=COMPANY_ID,
            node_id=nid,
            start_date=args.start_date,
            end_date=args.end_date,
            output_dir=str(reports_dir),
            enviar_correo=False,
        )
        try:
            out = generate_report(gen_args)
            print(f"[OK] {out}")
            exitosos.append(nid)
        except Exception as e:
            print(f"[ERROR] {nid}: {e}", file=sys.stderr)

    if not exitosos:
        print("[ERROR] No hay nodos exitosos; no se genera el agregado.", file=sys.stderr)
        return 1

    print()
    print("[INFO] Generando reporte agregado (después de los individuales)...")
    out_agg = generate_aggregated_report(
        COMPANY_ID,
        exitosos,
        args.start_date,
        args.end_date,
        output_dir=str(reports_dir),
        generate_ppt=False,
        parallel_node_fetch=args.parallel_agregado,
        max_parallel_workers=max(1, args.workers_agregado),
    )
    print(f"[OK] Agregado: {out_agg.resolve()}")
    if len(exitosos) < len(NODOS_AGUNSA):
        print(
            f"[INFO] Agregado con {len(exitosos)} nodo(s) exitoso(s) "
            f"(se excluyen fallidos de la lista anterior)."
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
