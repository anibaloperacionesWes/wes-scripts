"""Script para generar reportes de todos los nodos de PARQUE_ARAUCO para noviembre 2025."""

import subprocess
import sys
from pathlib import Path

# Configurar codificación UTF-8 para la consola
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Todos los nodos de PARQUE_ARAUCO (empresa 000025)
NODOS_PARQUE_ARAUCO = [
    "000025-08",  # PLACA BANCARIA
    "000025-09",  # SALA DE BOMBAS FALABELLA
    "000025-10",  # SALA DE BOMBAS RIPLEY
    "000025-11",  # MATRIZ PRIMER PISO
    "000025-12",  # MATRIZ RIEGO PLAZA
    "000025-14",  # MATRIZ RED DE INCENDIO
    "000025-15",  # MATRIZ PRINCIPAL
    "000025-16",  # BAÑOS
    "000025-17",  # SAN IGANCION 300
    "000025-18",  # SAN IGNACIO 500
    "000025-19",  # ESTANQUE SUR
    "000025-20",  # ANDEN 3-4 MATRIZ PRINCIPAL
    "000025-21",  # SANDIA BAÑOS 2-3-6-7 FREDO
    "000025-23",  # LLENADO DE PILETA
    "000025-24",  # PILETA CASCADA
    "000025-35",  # PAK BAZAR GOURMET (reemplazo 000025-25)
    "000025-36",  # PAK DL KENNEDY (reemplazo 000025-26)
    "000025-27",  # DISTRITO DE LUJO
    "000025-28",  # SANDIA MALL 1 PISO -4
    "000025-29",  # ANDEN 3-4 RESTAURANTE
]

COMPANY_ID = "000025"
START_DATE = "01112025"  # 01 de noviembre 2025
END_DATE = "30112025"    # 30 de noviembre 2025

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def main():
    print("=" * 60)
    print("GENERANDO REPORTES INDIVIDUALES PARA NOVIEMBRE 2025")
    print("=" * 60)
    print(f"Empresa: PARQUE_ARAUCO ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Total de nodos: {len(NODOS_PARQUE_ARAUCO)}")
    print("=" * 60)
    print()
    
    nodos_exitosos = []
    nodos_fallidos = []
    
    for i, node_id in enumerate(NODOS_PARQUE_ARAUCO, 1):
        print(f"[{i}/{len(NODOS_PARQUE_ARAUCO)}] Generando reporte para nodo {node_id}...")
        
        cmd = [
            PYTHON_EXE,
            SCRIPT_PATH,
            "--company-id", COMPANY_ID,
            "--node-id", node_id,
            "--start-date", START_DATE,
            "--end-date", END_DATE,
        ]
        
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
            if result.returncode == 0:
                print(f"  [OK] Reporte generado exitosamente")
                nodos_exitosos.append(node_id)
            else:
                print(f"  [ERROR] {result.stderr[:200]}")
                nodos_fallidos.append(node_id)
        except subprocess.TimeoutExpired:
            print(f"  [ERROR] Timeout al generar reporte")
            nodos_fallidos.append(node_id)
        except Exception as e:
            print(f"  [ERROR] {e}")
            nodos_fallidos.append(node_id)
        
        print()
    
    print("=" * 60)
    print("RESUMEN DE REPORTES INDIVIDUALES")
    print("=" * 60)
    print(f"Exitosos: {len(nodos_exitosos)}")
    print(f"Fallidos: {len(nodos_fallidos)}")
    
    if nodos_fallidos:
        print(f"\nNodos con errores: {', '.join(nodos_fallidos)}")
    
    print()
    print("=" * 60)
    print("GENERANDO REPORTE AGREGADO")
    print("=" * 60)
    
    # Generar reporte agregado solo con los nodos exitosos
    if nodos_exitosos:
        print(f"Generando reporte agregado con {len(nodos_exitosos)} nodos...")
        
        # Importar la función de reporte agregado
        sys.path.insert(0, str(Path.cwd()))
        from generar_reporte_word import generate_aggregated_report
        
        try:
            output_path = generate_aggregated_report(
                COMPANY_ID,
                nodos_exitosos,
                START_DATE,
                END_DATE
            )
            print(f"[OK] Reporte agregado generado exitosamente:")
            print(f"  {output_path}")
        except Exception as e:
            print(f"[ERROR] Error al generar reporte agregado: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("No se pueden generar reportes agregados porque no hay nodos exitosos.")
    
    print()
    print("=" * 60)
    print("PROCESO COMPLETADO")
    print("=" * 60)

if __name__ == "__main__":
    main()

