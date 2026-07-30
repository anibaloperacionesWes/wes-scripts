"""
Script para completar el reporte faltante de Estadio Israelita y luego generar agregado + PPT
"""

import sys
from pathlib import Path
import argparse

sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import generate_report

COMPANY_ID = "000030"
NODE_ID_FALTANTE = "000030-02"  # Matriz Chesterton
START_DATE = "28/12/2025"
END_DATE = "28/12/2026"

print("=" * 70)
print("  COMPLETANDO REPORTE FALTANTE - ESTADIO ISRAELITA")
print(f"  Nodo: {NODE_ID_FALTANTE}")
print(f"  Período: {START_DATE} - {END_DATE}")
print("=" * 70)
print()

print(f"[1/1] Generando reporte para {NODE_ID_FALTANTE}...")
try:
    args = argparse.Namespace(
        company_id=COMPANY_ID,
        node_id=NODE_ID_FALTANTE,
        start_date=START_DATE,
        end_date=END_DATE,
        output_dir="reports",
        enviar_correo=False
    )
    report_path = generate_report(args)
    if report_path:
        print(f"[OK] Reporte generado: {report_path}")
    else:
        print(f"[ADVERTENCIA] generate_report retornó None")
except Exception as e:
    print(f"[ERROR] No se pudo generar el reporte: {e}")
    import traceback
    traceback.print_exc()
