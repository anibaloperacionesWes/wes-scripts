"""Script para regenerar todos los reportes de Fundo Zapallar con los cambios implementados."""

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

NODOS_FUNDO_ZAPALLAR = list(FUNDO_ZAPALLAR_NODE_IDS)

COMPANY_ID = "000027"
NODO_FUENTE = "000027-01"  # Matriz ESVAL

# Periodo: 01 de diciembre hasta ayer
START_DATE = "01/12/2025"
END_DATE = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def buscar_reporte_existente(node_id: str) -> bool:
    """Verifica si ya existe un reporte para el nodo en el periodo."""
    start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
    end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
    start_str = start_dt.strftime("%Y%m%d")
    end_str = end_dt.strftime("%Y%m%d")
    
    pattern = f"Reporte_{COMPANY_ID}_{node_id}_{start_str}_{end_str}.docx"
    reportes_dir = Path("reports/Fundo_Zapallar/REPORTE")
    
    if not reportes_dir.exists():
        return False
    
    # Buscar en todas las subcarpetas
    for carpeta in reportes_dir.iterdir():
        if carpeta.is_dir():
            reporte = carpeta / pattern
            if reporte.exists():
                return True
    return False

def main():
    # Primero limpiar duplicados existentes
    print("Limpiando duplicados existentes...")
    try:
        from limpiar_duplicados_fundo_zapallar import limpiar_duplicados
        limpiar_duplicados()
        print()
    except Exception as e:
        print(f"Advertencia: No se pudieron limpiar duplicados: {e}")
        print()
    
    print("=" * 70)
    print("REGENERANDO TODOS LOS REPORTES DE FUNDO ZAPALLAR")
    print("=" * 70)
    print(f"Empresa: Fundo Zapallar ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Total de nodos: {len(NODOS_FUNDO_ZAPALLAR)}")
    print("=" * 70)
    print()
    print("CAMBIOS IMPLEMENTADOS:")
    print("  ✓ Gráfica de pie NO se muestra cuando proyección de fuga es cero")
    print("  ✓ Gráfica mensual NO se muestra cuando proyección de fuga es cero")
    print("  ✓ Gráfica del día con mayor alerta marca el valor exacto de la alerta")
    print("  ✓ Hora de la alerta se muestra con 1 decimal (ej: 9.5 h)")
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
        
        # Verificar si ya existe un reporte
        if buscar_reporte_existente(node_id):
            print(f"[{i}/{len(NODOS_FUNDO_ZAPALLAR)}] {node_name} ({node_id}) [{tipo}]...")
            print(f"  [SALTANDO] Ya existe un reporte para este nodo. Eliminando y regenerando...")
            # Eliminar reporte existente antes de regenerar
            start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
            end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
            start_str = start_dt.strftime("%Y%m%d")
            end_str = end_dt.strftime("%Y%m%d")
            pattern = f"Reporte_{COMPANY_ID}_{node_id}_{start_str}_{end_str}.docx"
            reportes_dir = Path("reports/Fundo_Zapallar/REPORTE")
            for carpeta in reportes_dir.iterdir():
                if carpeta.is_dir():
                    reporte = carpeta / pattern
                    if reporte.exists():
                        try:
                            import shutil
                            shutil.rmtree(carpeta)
                            print(f"  [ELIMINADO] Carpeta anterior eliminada")
                        except Exception as e:
                            print(f"  [ADVERTENCIA] No se pudo eliminar carpeta anterior: {e}")
        
        print(f"[{i}/{len(NODOS_FUNDO_ZAPALLAR)}] Generando reporte para {node_name} ({node_id}) [{tipo}]...")
        sys.stdout.flush()
        
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
                errors='ignore'
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
        sys.stdout.flush()
    
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
        print(f"  - Incluirá gráfica del día con mayor alerta con hora en decimales")
        print(f"  - Incluirá gráfica de balance hídrico (ESVAL vs consumidores)")
        print()
        sys.stdout.flush()
        
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
        print("VERIFICAR EN LOS REPORTES:")
        print("  ✓ Gráfica de pie NO aparece cuando proyección de fuga es cero")
        print("  ✓ Gráfica mensual NO aparece cuando proyección de fuga es cero")
        print("  ✓ Gráfica del día con mayor alerta muestra hora con 1 decimal (ej: 9.5 h)")
        print("  ✓ Punto destacado en gráfica corresponde al valor exacto de la alerta")

if __name__ == "__main__":
    main()

