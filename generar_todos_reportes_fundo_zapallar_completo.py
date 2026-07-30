"""Script para generar todos los reportes de Fundo Zapallar desde cero (individuales + agregado)."""

import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path
from generar_reporte_word import generate_aggregated_report, get_node_name
from exclusiones_reportes import FUNDO_ZAPALLAR_NODE_IDS

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Ocho puntos de monitoreo (000027-05 Riego excluido en exclusiones_reportes)
NODOS_FUNDO_ZAPALLAR = list(FUNDO_ZAPALLAR_NODE_IDS)

COMPANY_ID = "000027"
NODO_FUENTE = "000027-01"  # Matriz ESVAL

# Periodo: 01 de diciembre hasta ayer
START_DATE = "01/12/2025"
END_DATE = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def main():
    print("=" * 70)
    print("GENERANDO TODOS LOS REPORTES DE FUNDO ZAPALLAR DESDE CERO")
    print("=" * 70)
    print(f"Empresa: Fundo Zapallar ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Total de nodos: {len(NODOS_FUNDO_ZAPALLAR)}")
    print(f"Fuente de agua: {get_node_name(NODO_FUENTE)} ({NODO_FUENTE})")
    print("=" * 70)
    print()
    print("NOTA: Este script generará:")
    print(f"  - {len(NODOS_FUNDO_ZAPALLAR)} reportes individuales (uno por cada nodo)")
    print("  - 1 reporte agregado con lógica especial para ESVAL")
    print("  - Incluirá gráfica del día con mayor alerta en todos los reportes")
    print("  - Manejará correctamente cuando la filtración proyectada es cero")
    print()
    print("=" * 70)
    print("GENERANDO REPORTES INDIVIDUALES")
    print("=" * 70)
    print()
    
    nodos_exitosos = []
    nodos_fallidos = []
    
    # Generar reportes individuales
    for i, node_id in enumerate(NODOS_FUNDO_ZAPALLAR, 1):
        node_name = get_node_name(node_id)
        tipo = "FUENTE" if node_id == NODO_FUENTE else "CONSUMIDOR"
        print(f"[{i}/{len(NODOS_FUNDO_ZAPALLAR)}] Generando reporte para {node_name} ({node_id}) [{tipo}]...")
        
        cmd = [
            PYTHON_EXE,
            SCRIPT_PATH,
            "--company-id", COMPANY_ID,
            "--node-id", node_id,
            "--start-date", START_DATE,
            "--end-date", END_DATE,
        ]
        
        try:
            result = subprocess.run(
                cmd, 
                capture_output=True, 
                text=True, 
                timeout=600,
                encoding='utf-8',
                errors='ignore'  # Ignorar errores de codificación
            )
            if result.returncode == 0:
                print(f"  [OK] Reporte generado exitosamente")
                nodos_exitosos.append(node_id)
            else:
                error_msg = result.stderr[:300] if result.stderr else result.stdout[:300]
                print(f"  [ERROR] {error_msg}")
                nodos_fallidos.append(node_id)
        except subprocess.TimeoutExpired:
            print(f"  [ERROR] Timeout al generar reporte")
            nodos_fallidos.append(node_id)
        except Exception as e:
            print(f"  [ERROR] {e}")
            nodos_fallidos.append(node_id)
        
        print()
    
    print("=" * 70)
    print("RESUMEN DE REPORTES INDIVIDUALES")
    print("=" * 70)
    print(f"Exitosos: {len(nodos_exitosos)}/{len(NODOS_FUNDO_ZAPALLAR)}")
    if nodos_exitosos:
        print(f"  Nodos exitosos: {', '.join([f'{get_node_name(n)} ({n})' for n in nodos_exitosos])}")
    if nodos_fallidos:
        print(f"Fallidos: {len(nodos_fallidos)}")
        print(f"  Nodos fallidos: {', '.join([f'{get_node_name(n)} ({n})' for n in nodos_fallidos])}")
    print()
    
    print("=" * 70)
    print("GENERANDO REPORTE AGREGADO")
    print("=" * 70)
    
    reporte_agregado = None
    
    if nodos_exitosos:
        print(f"Generando reporte agregado con {len(nodos_exitosos)} nodos...")
        print(f"Nodos incluidos:")
        for n in nodos_exitosos:
            print(f"  - {get_node_name(n)} ({n})")
        print()
        print("Características del reporte agregado:")
        print(f"  - Excluirá {get_node_name(NODO_FUENTE)} del cálculo de consumo efectivo")
        print(f"  - Incluirá narrativa explicando que ESVAL es la fuente que alimenta los demás puntos")
        print(f"  - Incluirá gráfica del día con mayor alerta")
        print(f"  - Incluirá gráfica de balance hídrico (ESVAL vs consumidores)")
        print()
        
        try:
            reporte_agregado = generate_aggregated_report(
                COMPANY_ID,
                nodos_exitosos,
                START_DATE,
                END_DATE
            )
            print("[OK] Reporte agregado generado exitosamente:")
            print(f"  {reporte_agregado}")
        except Exception as e:
            print(f"[ERROR] Error al generar reporte agregado: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("No se pueden generar reportes agregados porque no hay nodos exitosos.")
    
    print()
    print("=" * 70)
    print("PROCESO COMPLETADO")
    print("=" * 70)
    print()
    print("RESUMEN FINAL:")
    print(f"  ✓ Reportes individuales generados: {len(nodos_exitosos)}/{len(NODOS_FUNDO_ZAPALLAR)}")
    if nodos_fallidos:
        print(f"  ✗ Reportes fallidos: {len(nodos_fallidos)}")
    print(f"  {'✓' if reporte_agregado else '✗'} Reporte agregado: {'Sí' if reporte_agregado else 'No'}")
    
    if reporte_agregado and reporte_agregado.exists():
        print()
        print("UBICACIÓN DE LOS REPORTES:")
        print(f"  - Individuales: reports/Fundo_Zapallar/REPORTE/")
        print(f"  - Agregado: {reporte_agregado}")
        print()
        print("CARACTERÍSTICAS IMPLEMENTADAS:")
        print("  ✓ Gráfica del día con mayor alerta en todos los reportes")
        print("  ✓ Manejo de filtración proyectada cero (sin gráfica comparativa)")
        print("  ✓ Lógica especial para ESVAL como fuente de agua")
        print("  ✓ Gráfica de balance hídrico en reporte agregado")

if __name__ == "__main__":
    main()














