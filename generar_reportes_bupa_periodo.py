"""Script para generar todos los reportes de BUPA desde el 03 de octubre hasta ayer (16-12-25)."""

import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path
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

# Periodo: 03 de octubre hasta ayer (16-12-25)
START_DATE = "03/10/2025"
# Calcular fecha de ayer
ayer = datetime.now() - timedelta(days=1)
END_DATE = ayer.strftime("%d/%m/%Y")

# Configuración: Para BUPA, normalmente todos son consumidores (sin fuente de agua)
FUENTE_AGUA_ID = None  # Todos los puntos son consumidores

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def main():
    print("=" * 70)
    print("GENERANDO REPORTES DE BUPA")
    print("=" * 70)
    print(f"Empresa: BUPA ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Total de nodos: {len(NODOS_BUPA)}")
    print("Configuración: Todos los puntos son consumidores (sin fuente de agua)")
    print("=" * 70)
    print()
    
    print("=" * 70)
    print("GENERANDO REPORTES INDIVIDUALES")
    print("=" * 70)
    print()
    
    nodos_exitosos = []
    nodos_fallidos = []
    
    # Generar reportes individuales
    for i, node_id in enumerate(NODOS_BUPA, 1):
        node_name = get_node_name(node_id)
        print(f"[{i}/{len(NODOS_BUPA)}] Generando reporte para {node_name} ({node_id})...")
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
        print()
        sys.stdout.flush()
        
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
                        import shutil
                        try:
                            shutil.rmtree(carpeta)
                            print("  [OK] Reporte anterior eliminado")
                        except Exception as e:
                            print(f"  [ADVERTENCIA] No se pudo eliminar: {e}")
        
        try:
            reporte_agregado = generate_aggregated_report(
                COMPANY_ID,
                nodos_exitosos,
                START_DATE,
                END_DATE,
                fuente_agua_id=FUENTE_AGUA_ID
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
    print(f"  - Reportes individuales generados: {len(nodos_exitosos)}/{len(NODOS_BUPA)}")
    if nodos_fallidos:
        print(f"  - Reportes fallidos: {len(nodos_fallidos)}")
        for node_id in nodos_fallidos:
            print(f"    * {get_node_name(node_id)} ({node_id})")
    print(f"  - Reporte agregado: {'Sí' if len(nodos_exitosos) >= 2 else 'No (se requieren al menos 2 reportes individuales)'}")
    print(f"  - Configuración: Todos los puntos son consumidores")
    print()
    print("Los reportes se han guardado en disco en la carpeta reports/BUPA/")

if __name__ == "__main__":
    main()










