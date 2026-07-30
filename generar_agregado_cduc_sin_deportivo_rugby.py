"""
Script para generar solo el reporte agregado de CDUC
usando el período detectado desde los archivos guardados.
Excluye los nodos "Edificio Deportivo" (000021-02) y "Rugby CDUC" (000021-08).
"""

import sys
from pathlib import Path
from datetime import datetime
import re

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import (
    get_company_name,
    get_node_name,
    generate_aggregated_report,
)

# Configuración
COMPANY_ID = "000021"  # CDUC
NODOS_EXCLUIDOS = ["000021-02", "000021-08"]  # Edificio Deportivo y Rugby CDUC

def detectar_periodo_desde_archivos() -> tuple:
    """
    Detecta el período de los reportes Word existentes de CDUC.
    Retorna (start_date, end_date) en formato "YYYY-MM-DD".
    """
    reports_dir = Path("reports/CDUC/REPORTE")
    
    if not reports_dir.exists():
        raise ValueError("No se encontró la carpeta de reportes de CDUC")
    
    # Buscar todos los archivos Word de reportes
    report_files = list(reports_dir.rglob("Reporte_*.docx"))
    
    if not report_files:
        raise ValueError("No se encontraron reportes Word de CDUC")
    
    # Extraer fechas de los nombres de archivo
    # Formato: Reporte_CDUC_Nombre_20251201_20260108.docx
    fechas_encontradas = []
    
    for report_file in report_files:
        # Buscar patrón de fechas en el nombre del archivo
        match = re.search(r'(\d{8})_(\d{8})', report_file.name)
        if match:
            start_str = match.group(1)
            end_str = match.group(2)
            
            try:
                start_date = datetime.strptime(start_str, "%Y%m%d").date()
                end_date = datetime.strptime(end_str, "%Y%m%d").date()
                fechas_encontradas.append((start_date, end_date))
            except ValueError:
                continue
    
    if not fechas_encontradas:
        raise ValueError("No se pudieron extraer fechas de los reportes existentes")
    
    # Encontrar el rango más amplio (fecha mínima de inicio y fecha máxima de fin)
    min_start = min(fecha[0] for fecha in fechas_encontradas)
    max_end = max(fecha[1] for fecha in fechas_encontradas)
    
    return (
        min_start.strftime("%Y-%m-%d"),
        max_end.strftime("%Y-%m-%d")
    )

def obtener_nodos_cduc() -> list:
    """
    Obtiene todos los nodos de CDUC excepto Edificio Deportivo y Rugby CDUC.
    Retorna lista de node_ids.
    """
    # Nodos conocidos de CDUC según generar_reporte_word.py
    nodos_cduc = {
        "000021-01": "Club House CDUC",
        "000021-02": "Edificio Deportivo",  # EXCLUIDO
        "000021-03": "Raimundo Tupper",
        "000021-04": "Equitación",
        "000021-05": "Calle de Servicio",
        "000021-07": "Canchas de Tenis",
        "000021-08": "Rugby CDUC",  # EXCLUIDO
    }
    
    # Filtrar excluyendo los nodos especificados
    nodos_incluidos = [
        node_id for node_id, node_name in nodos_cduc.items()
        if node_id not in NODOS_EXCLUIDOS
    ]
    
    return nodos_incluidos

def main():
    print("=" * 60)
    print("  GENERACIÓN DE REPORTE AGREGADO CDUC")
    print("  (Desde archivos guardados, excluyendo Edificio Deportivo y Rugby CDUC)")
    print("=" * 60)
    print()
    
    # Detectar período desde archivos
    print("[1/3] Detectando período desde archivos guardados...")
    try:
        start_date, end_date = detectar_periodo_desde_archivos()
        print(f"[OK] Período detectado: {start_date} a {end_date}")
    except Exception as e:
        print(f"[ERROR] No se pudo detectar el período: {e}")
        return
    print()
    
    # Obtener nodos (excluyendo Edificio Deportivo y Rugby CDUC)
    print("[2/3] Obteniendo nodos de CDUC (excluyendo Edificio Deportivo y Rugby CDUC)...")
    node_ids = obtener_nodos_cduc()
    
    print(f"[OK] Se procesarán {len(node_ids)} nodo(s):")
    for node_id in node_ids:
        node_name = get_node_name(node_id)
        print(f"  - {node_id}: {node_name}")
    print()
    
    print(f"[INFO] Nodos excluidos:")
    for node_id in NODOS_EXCLUIDOS:
        node_name = get_node_name(node_id)
        print(f"  - {node_id}: {node_name}")
    print()
    
    # Generar reporte agregado
    print("[3/3] Generando reporte agregado...")
    print("NOTA: Este proceso requiere acceso a la API WES para obtener los datos.")
    print("      Los archivos guardados solo se usan para detectar el período.")
    print()
    
    try:
        aggregated_report_path = generate_aggregated_report(
            company_id=COMPANY_ID,
            node_ids=node_ids,
            start_date=start_date,
            end_date=end_date,
            output_dir="reports",
            fuente_agua_id=None
        )
        print(f"[OK] Reporte agregado generado: {aggregated_report_path}")
    except Exception as e:
        print(f"[ERROR] No se pudo generar el reporte agregado: {e}")
        import traceback
        traceback.print_exc()
        return
    
    print()
    print("=" * 60)
    print("  PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 60)
    print()
    print(f"[INFO] Reporte agregado generado: 1")
    print(f"[INFO] Periodo: {start_date} a {end_date}")
    print(f"[INFO] Nodos incluidos: {len(node_ids)}")
    print(f"[INFO] Nodos excluidos: {len(NODOS_EXCLUIDOS)} (Edificio Deportivo y Rugby CDUC)")

if __name__ == "__main__":
    main()



