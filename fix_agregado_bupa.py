# -*- coding: utf-8 -*-
"""Regenerar reporte agregado de BUPA con todos los nodos."""

import sys
import shutil
from pathlib import Path
from datetime import datetime

# Configurar codificación
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

from generar_reporte_word import generate_aggregated_report, get_company_name, get_node_name

COMPANY_ID = "000029"
NODOS_BUPA = [
    "000029-01",  # Llenado de Estanques
    "000029-02",  # Torre A
    "000029-03",  # Torre B1
    "000029-04",  # Torre B2
    "000029-05",  # Torre C
    "000029-06",  # Central Térmica
]
START_DATE = "03/10/2025"
END_DATE = datetime.now().strftime("%d/%m/%Y")

print("=" * 70)
print("REGENERANDO REPORTE AGREGADO DE BUPA")
print("=" * 70)
print(f"Empresa: {get_company_name(COMPANY_ID)} ({COMPANY_ID})")
print(f"Periodo: {START_DATE} - {END_DATE}")
print(f"Total de nodos: {len(NODOS_BUPA)}")
print()
print("Nodos a incluir:")
for node_id in NODOS_BUPA:
    print(f"  - {get_node_name(node_id)} ({node_id})")
print()

# Eliminar reporte anterior
start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
start_str = start_dt.strftime("%Y%m%d")
end_str = end_dt.strftime("%Y%m%d")
pattern = f"Reporte_Agregado_{COMPANY_ID}_{start_str}_{end_str}.docx"

agregado_dir = Path("reports") / "BUPA" / "ABREGADO"

if agregado_dir.exists():
    print("Eliminando reportes anteriores...")
    for carpeta in agregado_dir.iterdir():
        if carpeta.is_dir():
            reporte_file = carpeta / pattern
            if reporte_file.exists():
                print(f"  Eliminando: {carpeta.name}")
                try:
                    shutil.rmtree(carpeta)
                    print(f"    [OK] Eliminado")
                except Exception as e:
                    print(f"    [ERROR] {e}")

print()
print("Generando nuevo reporte agregado con TODOS los nodos...")
print()

try:
    reporte = generate_aggregated_report(
        COMPANY_ID,
        NODOS_BUPA,
        START_DATE,
        END_DATE,
        fuente_agua_id=None
    )
    print()
    print("=" * 70)
    print("REPORTE AGREGADO GENERADO EXITOSAMENTE")
    print("=" * 70)
    print(f"Ubicación: {reporte}")
    print(f"Nodos incluidos: {len(NODOS_BUPA)}")
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()














