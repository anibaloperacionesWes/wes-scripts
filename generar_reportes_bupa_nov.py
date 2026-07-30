"""Script para generar reportes de todos los nodos de BUPA para noviembre 2025."""

import subprocess
import sys
from pathlib import Path

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
START_DATE = "01112025"  # 01 de noviembre 2025
END_DATE = "30112025"    # 30 de noviembre 2025

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def main():
    # Configurar codificación UTF-8
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    
    print("=" * 60)
    print("GENERANDO REPORTES INDIVIDUALES PARA BUPA - NOVIEMBRE 2025")
    print("=" * 60)
    print(f"Empresa: BUPA ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Total de nodos: {len(NODOS_BUPA)}")
    print("=" * 60)
    print()
    
    nodos_exitosos = []
    nodos_fallidos = []
    
    for i, node_id in enumerate(NODOS_BUPA, 1):
        print(f"[{i}/{len(NODOS_BUPA)}] Generando reporte para nodo {node_id}...")
        
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
                error_msg = result.stderr[:200] if result.stderr else result.stdout[:200]
                print(f"  [ERROR] {error_msg}")
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


