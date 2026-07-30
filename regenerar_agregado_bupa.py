"""Script para regenerar el reporte agregado de BUPA con TODOS los nodos."""

import sys
from datetime import datetime
from pathlib import Path
import shutil
from generar_reporte_word import generate_aggregated_report, get_company_name, get_node_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Todos los nodos de BUPA (empresa 000029)
NODOS_BUPA = [
    "000029-01",  # Llenado de Estanques
    "000029-02",  # Torre A
    "000029-03",  # Torre B1
    "000029-04",  # Torre B2
    "000029-05",  # Torre C
    "000029-06",  # Central Térmica
]

COMPANY_ID = "000029"
START_DATE = "03/10/2025"
END_DATE = datetime.now().strftime("%d/%m/%Y")
FUENTE_AGUA_ID = None  # BUPA no tiene fuente de agua

def main():
    print("=" * 70)
    print("REGENERANDO REPORTE AGREGADO DE BUPA CON TODOS LOS NODOS")
    print("=" * 70)
    print(f"Empresa: BUPA ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Total de nodos: {len(NODOS_BUPA)}")
    print("Nodos incluidos:")
    for node_id in NODOS_BUPA:
        print(f"  - {get_node_name(node_id)} ({node_id})")
    print("=" * 70)
    print()
    
    # Eliminar reporte agregado anterior si existe
    start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
    end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
    start_str = start_dt.strftime("%Y%m%d")
    end_str = end_dt.strftime("%Y%m%d")
    pattern = f"Reporte_Agregado_{COMPANY_ID}_{start_str}_{end_str}.docx"
    
    company_name = get_company_name(COMPANY_ID)
    safe_company_name = company_name.replace(" ", "_")
    agregado_dir = Path("reports") / safe_company_name / "ABREGADO"
    
    if agregado_dir.exists():
        for carpeta in agregado_dir.iterdir():
            if carpeta.is_dir():
                reporte_file = carpeta / pattern
                if reporte_file.exists():
                    print(f"Eliminando reporte agregado anterior: {carpeta}")
                    try:
                        shutil.rmtree(carpeta)
                        print("  [OK] Reporte anterior eliminado")
                    except Exception as e:
                        print(f"  [ADVERTENCIA] No se pudo eliminar: {e}")
    
    print()
    print("Generando reporte agregado con TODOS los nodos...")
    print()
    
    try:
        reporte_agregado = generate_aggregated_report(
            COMPANY_ID,
            NODOS_BUPA,  # Pasar TODOS los nodos
            START_DATE,
            END_DATE,
            fuente_agua_id=FUENTE_AGUA_ID
        )
        print()
        print("=" * 70)
        print("REPORTE AGREGADO GENERADO EXITOSAMENTE")
        print("=" * 70)
        print(f"Ubicación: {reporte_agregado}")
        print(f"Nodos incluidos: {len(NODOS_BUPA)}")
    except Exception as e:
        print(f"[ERROR] Error al generar reporte agregado: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

