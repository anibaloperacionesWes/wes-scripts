"""
Reporte agregado — Clínica Bupa Antofagasta (nodos 000029-07..10).

Formato fin de mes (agregado extendido, como UDD / Club Providencia / Zapallar):
  - Portada: resumen ejecutivo + comparación por punto
  - Cuerpo: evolución diaria, día de mayor consumo, resumen con costo nocturno,
    análisis de consumos nocturnos (suma real del periodo, sin proyección a 30 días)

Uso:
  python generar_agregado_bupa_antofagasta.py
  python generar_agregado_bupa_antofagasta.py --start-date 23/07/2026 --end-date 11/09/2026
"""

from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from generar_reporte_word import generate_aggregated_report

COMPANY_ID = "000029"
NODOS = [
    "000029-07",
    "000029-08",
    "000029-09",
    "000029-10",
]
# Inicio operativo de Antofagasta
DEFAULT_START = "23/07/2026"
FOLDER = "Bupa_Antofagasta"


def main() -> None:
    ap = argparse.ArgumentParser(description="Agregado Bupa Antofagasta (formato fin de mes)")
    ap.add_argument("--start-date", default=DEFAULT_START, help="dd/mm/aaaa")
    ap.add_argument(
        "--end-date",
        default=datetime.now().strftime("%d/%m/%Y"),
        help="dd/mm/aaaa (default: hoy)",
    )
    args = ap.parse_args()

    print("=" * 70)
    print("[INFO] Clínica Bupa Antofagasta — reporte agregado (formato fin de mes)")
    print(f"       Empresa {COMPANY_ID} | {len(NODOS)} nodo(s): {', '.join(NODOS)}")
    print(f"       Periodo: {args.start_date} – {args.end_date}")
    print(f"       Carpeta: reports/{FOLDER}/ABREGADO/")
    print("=" * 70)
    t0 = time.perf_counter()
    out = generate_aggregated_report(
        company_id=COMPANY_ID,
        node_ids=list(NODOS),
        start_date=args.start_date,
        end_date=args.end_date,
        output_dir="reports",
        apply_exclusions=False,
        generate_ppt=False,
        parallel_node_fetch=True,
        max_parallel_workers=4,
        fuente_agua_id=None,
        company_folder_override=FOLDER,
        # 1ª página = resumen/comparación estándar; resto = formato fin de mes.
        conservar_portada_estandar=True,
    )
    print(f"[OK] {out}")
    print(f"[INFO] Tiempo: {time.perf_counter() - t0:.1f} s")


if __name__ == "__main__":
    main()
