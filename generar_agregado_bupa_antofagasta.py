"""
Reporte agregado — Clínica Bupa Antofagasta (nodos 000029-07..10).

Uso:
  python generar_agregado_bupa_antofagasta.py
"""

from __future__ import annotations

import sys
import time

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
# Días civiles completos (excluye el día parcial de hoy)
START_DATE = "23/07/2026"
END_DATE = "27/07/2026"
FOLDER = "Bupa_Antofagasta"


def main() -> None:
    print("=" * 70)
    print("[INFO] Clínica Bupa Antofagasta — reporte agregado")
    print(f"       Empresa {COMPANY_ID} | {len(NODOS)} nodo(s): {', '.join(NODOS)}")
    print(f"       Periodo: {START_DATE} – {END_DATE}")
    print(f"       Carpeta: reports/{FOLDER}/ABREGADO/")
    print("=" * 70)
    t0 = time.perf_counter()
    out = generate_aggregated_report(
        company_id=COMPANY_ID,
        node_ids=list(NODOS),
        start_date=START_DATE,
        end_date=END_DATE,
        output_dir="reports",
        apply_exclusions=False,
        generate_ppt=False,
        parallel_node_fetch=True,
        max_parallel_workers=4,
        fuente_agua_id=None,
        company_folder_override=FOLDER,
    )
    print(f"[OK] {out}")
    print(f"[INFO] Tiempo: {time.perf_counter() - t0:.1f} s")


if __name__ == "__main__":
    main()
