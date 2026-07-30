"""
Reporte agregado Las Condes — mayo (mes completo).

Uso:
  python generar_agregado_las_condes_mayo.py
  python generar_agregado_las_condes_mayo.py --year 2025
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))

from generar_reporte_word import (
    ENTITY_BASE_URL,
    generate_aggregated_report,
    get_company_name,
)

COMPANY_ID = "000022"


def main() -> int:
    ap = argparse.ArgumentParser(description="Reporte agregado Las Condes — mayo")
    ap.add_argument("--year", type=int, default=2026, help="Año del mes de mayo (default: 2026)")
    args = ap.parse_args()
    yy = args.year
    start = f"01/05/{yy}"
    end = f"31/05/{yy}"

    print("=" * 70)
    print("  REPORTE AGREGADO - LAS CONDES - MAYO")
    print(f"  Periodo: {start} - {end}")
    print("=" * 70)

    company_name = get_company_name(COMPANY_ID)
    print(f"[INFO] Empresa: {company_name} ({COMPANY_ID})")

    r = requests.get(f"{ENTITY_BASE_URL}/companies/{COMPANY_ID}", timeout=30)
    r.raise_for_status()
    nodes = r.json().get("nodes", [])
    node_ids = [n["nodeId"] for n in nodes if n.get("nodeId")]
    if not node_ids:
        print("[ERROR] No hay nodos para Las Condes.", file=sys.stderr)
        return 1

    for n in nodes:
        print(f"  - {n.get('nodeId')}: {n.get('name')}")

    path = generate_aggregated_report(
        company_id=COMPANY_ID,
        node_ids=node_ids,
        start_date=start,
        end_date=end,
        output_dir="reports",
    )
    print(f"[OK] Reporte agregado: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
