"""Script interactivo para generar reportes individuales y agregado con preguntas sobre fuente de agua."""

import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path
from generar_reporte_word import generate_aggregated_report, get_company_name, get_node_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def preguntar_empresa():
    """Pregunta al usuario por el ID de la empresa."""
    print("=" * 70)
    print("GENERACIÓN DE REPORTES INTERACTIVA")
    print("=" * 70)
    print()
    company_id = input("Ingrese el ID de la empresa (ej: 000027): ").strip()
    if not company_id:
        print("Error: Debe ingresar un ID de empresa.")
        sys.exit(1)
    return company_id

def preguntar_nodos(company_id):
    """Pregunta al usuario por los nodos a analizar."""
    print()
    print("Ingrese los IDs de los nodos a analizar (uno por línea).")
    print("Presione Enter en una línea vacía para terminar.")
    print()
    node_ids = []
    while True:
        node_id = input(f"Nodo {len(node_ids) + 1} (o Enter para terminar): ").strip()
        if not node_id:
            break
        node_ids.append(node_id)
    
    if not node_ids:
        print("Error: Debe ingresar al menos un nodo.")
        sys.exit(1)
    
    return node_ids

def preguntar_fuente_agua(node_ids):
    """Pregunta si existe una fuente de agua y cuál es."""
    print()
    print("=" * 70)
    print("CONFIGURACIÓN DE FUENTE DE AGUA")
    print("=" * 70)
    print()
    print("¿Existe un punto que actúa como fuente de agua que alimenta a los demás?")
    print("(Por ejemplo: Matriz ESVAL en Fundo Zapallar)")
    print()
    
    tiene_fuente = input("¿Tiene fuente de agua? (s/n): ").strip().lower()
    
    fuente_id = None
    if tiene_fuente in ['s', 'si', 'sí', 'y', 'yes']:
        print()
        print("Nodos disponibles:")
        for i, node_id in enumerate(node_ids, 1):
            node_name = get_node_name(node_id)
            print(f"  {i}. {node_id} - {node_name}")
        print()
        fuente_id = input("Ingrese el ID del nodo que es la fuente de agua: ").strip()
        if fuente_id not in node_ids:
            print(f"Advertencia: {fuente_id} no está en la lista de nodos. Continuando sin fuente.")
            fuente_id = None
    
    return fuente_id

def preguntar_periodo():
    """Pregunta por el periodo de análisis."""
    print()
    print("=" * 70)
    print("PERIODO DE ANÁLISIS")
    print("=" * 70)
    print()
    start_date = input("Fecha inicio (DD/MM/YYYY, ej: 01/12/2025): ").strip()
    end_date = input("Fecha fin (DD/MM/YYYY, ej: 08/12/2025, o Enter para ayer): ").strip()
    
    if not end_date:
        end_date = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
    
    return start_date, end_date

def generar_reporte_individual(company_id, node_id, start_date, end_date):
    """Genera un reporte individual."""
    cmd = [
        PYTHON_EXE,
        SCRIPT_PATH,
        "--company-id", company_id,
        "--node-id", node_id,
        "--start-date", start_date,
        "--end-date", end_date,
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
        return result.returncode == 0
    except Exception as e:
        print(f"  [ERROR] {e}")
        return False

def main():
    # Obtener información del usuario
    company_id = preguntar_empresa()
    node_ids = preguntar_nodos(company_id)
    fuente_id = preguntar_fuente_agua(node_ids)
    start_date, end_date = preguntar_periodo()
    
    company_name = get_company_name(company_id)
    
    # Resumen
    print()
    print("=" * 70)
    print("RESUMEN DE CONFIGURACIÓN")
    print("=" * 70)
    print(f"Empresa: {company_name} ({company_id})")
    print(f"Periodo: {start_date} - {end_date}")
    print(f"Total de nodos: {len(node_ids)}")
    if fuente_id:
        print(f"Fuente de agua: {get_node_name(fuente_id)} ({fuente_id})")
    else:
        print("Fuente de agua: No especificada (todos son consumidores)")
    print("=" * 70)
    print()
    
    confirmar = input("¿Continuar con la generación? (s/n): ").strip().lower()
    if confirmar not in ['s', 'si', 'sí', 'y', 'yes']:
        print("Operación cancelada.")
        sys.exit(0)
    
    print()
    print("=" * 70)
    print("GENERANDO REPORTES INDIVIDUALES")
    print("=" * 70)
    print()
    
    nodos_exitosos = []
    nodos_fallidos = []
    
    # Generar reportes individuales
    for i, node_id in enumerate(node_ids, 1):
        node_name = get_node_name(node_id)
        tipo = "FUENTE" if node_id == fuente_id else "CONSUMIDOR"
        print(f"[{i}/{len(node_ids)}] Generando reporte para {node_name} ({node_id}) [{tipo}]...")
        sys.stdout.flush()
        
        if generar_reporte_individual(company_id, node_id, start_date, end_date):
            print(f"  [OK] Reporte generado exitosamente")
            nodos_exitosos.append(node_id)
        else:
            print(f"  [ERROR] Fallo al generar reporte")
            nodos_fallidos.append(node_id)
        
        print()
        sys.stdout.flush()
    
    # Generar reporte agregado al final con TODOS los nodos exitosos
    if len(nodos_exitosos) >= 2:
        print()
        print("=" * 70)
        print("GENERANDO REPORTE AGREGADO CON TODOS LOS NODOS")
        print("=" * 70)
        print(f"Total de nodos exitosos: {len(nodos_exitosos)}")
        print("Nodos incluidos:")
        for node_id in nodos_exitosos:
            print(f"  - {get_node_name(node_id)} ({node_id})")
        if fuente_id:
            print(f"Fuente de agua configurada: {get_node_name(fuente_id)} ({fuente_id})")
        print()
        sys.stdout.flush()
        
        # Eliminar reporte agregado anterior si existe (para regenerarlo con todos los nodos)
        start_dt = datetime.strptime(start_date, "%d/%m/%Y")
        end_dt = datetime.strptime(end_date, "%d/%m/%Y")
        start_str = start_dt.strftime("%Y%m%d")
        end_str = end_dt.strftime("%Y%m%d")
        pattern = f"Reporte_Agregado_{company_id}_{start_str}_{end_str}.docx"
        
        agregado_dir = Path("reports") / company_name.replace(" ", "_") / "ABREGADO"
        
        if agregado_dir.exists():
            for carpeta in agregado_dir.iterdir():
                if carpeta.is_dir():
                    reporte_file = carpeta / pattern
                    if reporte_file.exists():
                        print(f"Eliminando reporte agregado anterior: {carpeta}")
                        import shutil
                        try:
                            shutil.rmtree(carpeta)
                            print("  [OK] Reporte anterior eliminado")
                        except Exception as e:
                            print(f"  [ADVERTENCIA] No se pudo eliminar: {e}")
        
        try:
            reporte_agregado = generate_aggregated_report(
                company_id,
                nodos_exitosos,  # Pasar TODOS los nodos exitosos
                start_date,
                end_date,
                fuente_agua_id=fuente_id if fuente_id else None
            )
            print(f"[OK] Reporte agregado generado exitosamente con {len(nodos_exitosos)} nodos:")
            print(f"  {reporte_agregado}")
        except Exception as e:
            print(f"[ERROR] Error al generar reporte agregado: {e}")
            import traceback
            traceback.print_exc()
    
    print()
    print("=" * 70)
    print("PROCESO COMPLETADO")
    print("=" * 70)
    print()
    print("RESUMEN:")
    print(f"  - Reportes individuales generados: {len(nodos_exitosos)}/{len(node_ids)}")
    if nodos_fallidos:
        print(f"  - Reportes fallidos: {len(nodos_fallidos)}")
    print(f"  - Reporte agregado: {'Sí' if len(nodos_exitosos) >= 2 else 'No (se requieren al menos 2 reportes individuales)'}")

if __name__ == "__main__":
    main()

