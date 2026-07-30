"""
Genera reportes Word individuales por nodo CDUC + reporte agregado para un periodo.

Uso:
  python generar_reportes_cduc_periodo.py
  python generar_reportes_cduc_periodo.py --start-date 01/01/2026 --end-date 06/04/2026
  python generar_reportes_cduc_periodo.py --solo-agregado
  python generar_reportes_cduc_periodo.py --parallel-agregado

--solo-agregado: solo el Word agregado (omite individuales; mucho más rápido si no necesitas un .docx por nodo).
--parallel-agregado: descarga medidas/alertas de los nodos en paralelo al armar el agregado.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from generar_reporte_word import (
    generate_aggregated_report,
    generate_report,
    get_company_name,
    is_node_excluded,
)

COMPANY_ID = "000021"
# Todos los nodos CDUC conocidos (no existe 000021-06 en catálogo WES).
NODOS_CDUC_TODOS = [
    "000021-01",
    "000021-02",
    "000021-03",
    "000021-04",
    "000021-05",
    "000021-07",
    "000021-08",
]


def main() -> int:
    p = argparse.ArgumentParser(description="Reportes CDUC individuales + agregado")
    p.add_argument("--start-date", default="01/01/2026", help="Inicio (dd/mm/aaaa)")
    p.add_argument("--end-date", default="06/04/2026", help="Término (dd/mm/aaaa)")
    p.add_argument(
        "--sin-exclusiones-en-agregado",
        action="store_true",
        help="Incluye en el agregado nodos que normalmente están excluidos (000021-02, 000021-08).",
    )
    p.add_argument(
        "--solo-agregado",
        action="store_true",
        help="No genera reportes por nodo; solo el Word agregado (más rápido).",
    )
    p.add_argument(
        "--parallel-agregado",
        action="store_true",
        help="Descarga datos de nodos en paralelo al generar el agregado (acelera esa fase).",
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
    # Ruta absoluta al proyecto: no depende del cwd (evita escrituras en carpetas equivocadas).
    reports_dir = root / "reports"

    company_name = get_company_name(COMPANY_ID)
    individuales = [
        nid
        for nid in NODOS_CDUC_TODOS
        if not is_node_excluded(nid, company_id=COMPANY_ID, company_name=company_name)
    ]
    excluidos = [n for n in NODOS_CDUC_TODOS if n not in individuales]

    print("=" * 70)
    print("CDUC — reportes individuales + agregado")
    print(f"Periodo: {args.start_date} -> {args.end_date}")
    print(f"Individuales ({len(individuales)}): {', '.join(individuales)}")
    if excluidos:
        print(f"[INFO] Sin informe individual (exclusión global): {', '.join(excluidos)}")
    print("=" * 70)

    if not args.solo_agregado:
        for i, nid in enumerate(individuales, 1):
            print(f"[{i}/{len(individuales)}] Generando {nid}...")
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
            except Exception as e:
                print(f"[ERROR] Falló {nid}: {e}", file=sys.stderr)
                return 1
    else:
        print("[INFO] --solo-agregado: se omiten reportes individuales.\n")

    print()
    print("[INFO] Generando reporte agregado...")
    apply_ex = not args.sin_exclusiones_en_agregado
    out_agg = generate_aggregated_report(
        COMPANY_ID,
        list(NODOS_CDUC_TODOS),
        args.start_date,
        args.end_date,
        output_dir=str(reports_dir),
        apply_exclusions=apply_ex,
        generate_ppt=False,
        parallel_node_fetch=args.parallel_agregado,
        max_parallel_workers=max(1, args.workers_agregado),
    )
    print(f"[OK] Agregado: {out_agg.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
