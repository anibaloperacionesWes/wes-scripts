"""Script para completar los reportes faltantes de Fundo Zapallar y generar el agregado."""

import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path
from generar_reporte_word import generate_aggregated_report, get_company_name, get_node_name
from exclusiones_reportes import FUNDO_ZAPALLAR_NODE_IDS

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

NODOS_FUNDO_ZAPALLAR = list(FUNDO_ZAPALLAR_NODE_IDS)

COMPANY_ID = "000027"
NODO_FUENTE = "000027-01"  # Matriz ESVAL

# Periodo: 01 de diciembre hasta ayer
START_DATE = "01/12/2025"
END_DATE = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def main():
    print("=" * 60)
    print("COMPLETANDO REPORTES FALTANTES DE FUNDO ZAPALLAR")
    print("=" * 60)
    print(f"Empresa: Fundo Zapallar ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print("=" * 60)
    print()
    
    # Nodos que ya tienen reporte (según lo que vimos)
    nodos_con_reporte = ["000027-01", "000027-02", "000027-03"]
    nodos_faltantes = [n for n in NODOS_FUNDO_ZAPALLAR if n not in nodos_con_reporte]
    
    print(f"Reportes existentes: {len(nodos_con_reporte)}")
    print(f"Reportes faltantes: {len(nodos_faltantes)}")
    print(f"Nodos faltantes: {', '.join([get_node_name(n) for n in nodos_faltantes])}")
    print()
    
    nodos_exitosos = list(nodos_con_reporte)  # Incluir los que ya existen
    nodos_fallidos = []
    
    # Generar solo los reportes faltantes
    if nodos_faltantes:
        print("=" * 60)
        print("GENERANDO REPORTES FALTANTES")
        print("=" * 60)
        for i, node_id in enumerate(nodos_faltantes, 1):
            node_name = get_node_name(node_id)
            print(f"[{i}/{len(nodos_faltantes)}] Generando reporte para {node_name} ({node_id})...")
            
            cmd = [
                PYTHON_EXE,
                SCRIPT_PATH,
                "--company-id", COMPANY_ID,
                "--node-id", node_id,
                "--start-date", START_DATE,
                "--end-date", END_DATE,
            ]
            
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
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
    else:
        print("No hay reportes faltantes.")
        print()
    
    print("=" * 60)
    print("GENERANDO REPORTE AGREGADO")
    print("=" * 60)
    
    reporte_agregado = None
    
    if nodos_exitosos:
        print(f"Generando reporte agregado con {len(nodos_exitosos)} nodos...")
        print(f"Nodos incluidos: {', '.join([get_node_name(n) for n in nodos_exitosos])}")
        print(f"NOTA: El reporte excluirá ESVAL del cálculo de consumo efectivo")
        print(f"y agregará narrativa especial sobre ESVAL como fuente de agua.")
        print()
        
        try:
            reporte_agregado = generate_aggregated_report(
                COMPANY_ID,
                nodos_exitosos,
                START_DATE,
                END_DATE
            )
            print(f"[OK] Reporte agregado generado exitosamente:")
            print(f"  {reporte_agregado}")
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
    print()
    print("Resumen:")
    print(f"  - Reportes individuales totales: {len(nodos_exitosos)}/{len(NODOS_FUNDO_ZAPALLAR)}")
    if nodos_fallidos:
        print(f"  - Reportes fallidos: {len(nodos_fallidos)} ({', '.join([get_node_name(n) for n in nodos_fallidos])})")
    print(f"  - Reporte agregado: {'Sí' if reporte_agregado else 'No'}")
    if reporte_agregado and reporte_agregado.exists():
        print(f"  - Ubicación: {reporte_agregado}")
        print()
        print("NOTA: El reporte agregado:")
        print(f"  - Excluye {get_node_name(NODO_FUENTE)} del cálculo de consumo efectivo")
        print(f"  - Incluye narrativa explicando que ESVAL es la fuente que alimenta los demás puntos")

if __name__ == "__main__":
    main()














