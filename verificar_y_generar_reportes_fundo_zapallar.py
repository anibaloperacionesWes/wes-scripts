"""Script para verificar y generar reportes de Fundo Zapallar solo si no existen."""

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

# Formatear fechas para buscar archivos
start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
start_str = start_dt.strftime("%Y%m%d")
end_str = end_dt.strftime("%Y%m%d")

def buscar_reporte_individual(node_id: str) -> bool:
    """Busca si existe un reporte individual para el nodo."""
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

def buscar_reporte_agregado() -> bool:
    """Busca si existe un reporte agregado."""
    pattern = f"Reporte_Agregado_{COMPANY_ID}_{start_str}_{end_str}.docx"
    agregado_dir = Path("reports/Fundo_Zapallar/ABREGADO")
    if not agregado_dir.exists():
        return False
    
    # Buscar en todas las subcarpetas
    for carpeta in agregado_dir.iterdir():
        if carpeta.is_dir():
            reporte = carpeta / pattern
            if reporte.exists():
                return True
    return False

def generar_reporte_individual(node_id: str) -> bool:
    """Genera un reporte individual para el nodo."""
    node_name = get_node_name(node_id)
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
        return result.returncode == 0
    except Exception:
        return False

def main():
    # Forzar flush de salida
    sys.stdout.flush()
    sys.stderr.flush()
    
    print("=" * 70)
    print("VERIFICACIÓN Y GENERACIÓN DE REPORTES DE FUNDO ZAPALLAR")
    print("=" * 70)
    print(f"Empresa: Fundo Zapallar ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Total de nodos: {len(NODOS_FUNDO_ZAPALLAR)}")
    print("=" * 70)
    sys.stdout.flush()
    print()
    
    # Verificar reportes individuales
    print("VERIFICANDO REPORTES INDIVIDUALES...")
    print("-" * 70)
    
    nodos_existentes = []
    nodos_faltantes = []
    nodos_exitosos = []
    nodos_fallidos = []
    
    for node_id in NODOS_FUNDO_ZAPALLAR:
        node_name = get_node_name(node_id)
        if buscar_reporte_individual(node_id):
            print(f"[EXISTE] {node_name} ({node_id}) - No se regenerará")
            nodos_existentes.append(node_id)
            nodos_exitosos.append(node_id)
        else:
            print(f"[FALTA] {node_name} ({node_id}) - Se generará")
            nodos_faltantes.append(node_id)
    
    print()
    
    # Generar reportes faltantes
    if nodos_faltantes:
        print("=" * 70)
        print(f"GENERANDO {len(nodos_faltantes)} REPORTE(S) FALTANTE(S)")
        print("=" * 70)
        print()
        
        for i, node_id in enumerate(nodos_faltantes, 1):
            node_name = get_node_name(node_id)
            tipo = "FUENTE" if node_id == NODO_FUENTE else "CONSUMIDOR"
            print(f"[{i}/{len(nodos_faltantes)}] Generando {node_name} ({node_id}) [{tipo}]...")
            
            if generar_reporte_individual(node_id):
                print(f"  [OK] Reporte generado exitosamente")
                nodos_exitosos.append(node_id)
            else:
                print(f"  [ERROR] Fallo al generar reporte")
                nodos_fallidos.append(node_id)
            print()
    else:
        print("Todos los reportes individuales ya existen. No se generarán nuevos.")
        print()
    
    # Verificar reporte agregado
    print("=" * 70)
    print("VERIFICANDO REPORTE AGREGADO...")
    print("-" * 70)
    
    if buscar_reporte_agregado():
        print("[EXISTE] Reporte agregado - No se regenerará")
        reporte_agregado_existe = True
    else:
        print("[FALTA] Reporte agregado - Se generará")
        reporte_agregado_existe = False
    
    print()
    
    # Generar reporte agregado si falta
    reporte_agregado = None
    if not reporte_agregado_existe:
        print("=" * 70)
        print("GENERANDO REPORTE AGREGADO")
        print("=" * 70)
        
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
    else:
        print("El reporte agregado ya existe. No se regenerará.")
    
    print()
    print("=" * 70)
    print("RESUMEN FINAL")
    print("=" * 70)
    print()
    print("REPORTES INDIVIDUALES:")
    print(f"  - Existentes (no regenerados): {len(nodos_existentes)}")
    print(f"  - Generados ahora: {len(nodos_faltantes) - len(nodos_fallidos)}")
    if nodos_fallidos:
        print(f"  - Fallidos: {len(nodos_fallidos)}")
        for n in nodos_fallidos:
            print(f"    ✗ {get_node_name(n)} ({n})")
    print(f"  - Total disponibles: {len(nodos_exitosos)}/{len(NODOS_FUNDO_ZAPALLAR)}")
    print()
    print("REPORTE AGREGADO:")
    if reporte_agregado_existe:
        print("  ✓ Ya existía (no se regeneró)")
    elif reporte_agregado:
        print(f"  ✓ Generado exitosamente: {reporte_agregado}")
    else:
        print("  ✗ No se pudo generar")
    print()
    print("CARACTERÍSTICAS IMPLEMENTADAS EN LOS REPORTES:")
    print("  ✓ Gráfica del día con mayor alerta en todos los reportes")
    print("  ✓ Manejo de filtración proyectada cero (sin gráfica comparativa)")
    print("  ✓ Lógica especial para ESVAL como fuente de agua (en agregado)")
    print("  ✓ Gráfica de balance hídrico en reporte agregado")

if __name__ == "__main__":
    main()

