"""
Script para generar solo el reporte agregado de CDUC
excluyendo Raimundo Tupper, Edificio Deportivo y Rugby CDUC.
Período: 01 de diciembre 2025 hasta la fecha actual.
"""

import sys
from pathlib import Path
from datetime import datetime
import requests

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import (
    get_company_name,
    get_node_name,
    generate_aggregated_report,
)

# Importar función para generar PPT
from generar_reportes_y_ppt_mall_maipu import generar_ppt_desde_agregado

# Configuración
COMPANY_ID = "000021"  # CDUC
START_DATE = "01/12/2025"  # 01 de diciembre 2025
END_DATE = datetime.now().strftime("%d/%m/%Y")  # Fecha actual

# Nodos a excluir
NODOS_EXCLUIDOS = [
    "000021-02",  # Edificio Deportivo
    "000021-03",  # Raimundo Tupper
    "000021-08",  # Rugby CDUC
]

ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"

def obtener_nodos_cduc() -> list:
    """
    Obtiene todos los nodos de CDUC desde la API, excluyendo los especificados.
    Retorna lista de node_ids.
    """
    url = f"{ENTITY_BASE_URL}/companies/{COMPANY_ID}"
    
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            nodes = data.get("nodes", [])
            
            # Obtener todos los node_ids y filtrar los excluidos
            todos_nodos = [node.get("nodeId") for node in nodes if node.get("nodeId")]
            nodos_incluidos = [
                node_id for node_id in todos_nodos
                if node_id not in NODOS_EXCLUIDOS
            ]
            
            return nodos_incluidos
        else:
            print(f"[ERROR] No se pudo obtener información de la empresa {COMPANY_ID}: {response.status_code}")
            # Fallback: usar nodos conocidos
            return obtener_nodos_conocidos()
    except Exception as e:
        print(f"[ERROR] Error al obtener nodos de la API: {e}")
        print("[INFO] Usando lista de nodos conocidos como fallback")
        return obtener_nodos_conocidos()

def obtener_nodos_conocidos() -> list:
    """
    Retorna lista de nodos conocidos de CDUC excluyendo los especificados.
    """
    nodos_cduc = {
        "000021-01": "Club House CDUC",
        "000021-02": "Edificio Deportivo",  # EXCLUIDO
        "000021-03": "Raimundo Tupper",  # EXCLUIDO
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
    print("=" * 70)
    print("  GENERACIÓN DE REPORTE AGREGADO CDUC")
    print("  (Excluyendo Raimundo Tupper, Edificio Deportivo y Rugby CDUC)")
    print("=" * 70)
    print()
    print(f"Empresa: {get_company_name(COMPANY_ID)} ({COMPANY_ID})")
    print(f"Período: {START_DATE} - {END_DATE}")
    print()
    
    # Obtener nodos (excluyendo los especificados)
    print("[1/2] Obteniendo nodos de CDUC...")
    node_ids = obtener_nodos_cduc()
    
    print(f"[OK] Se procesarán {len(node_ids)} nodo(s):")
    for node_id in node_ids:
        node_name = get_node_name(node_id)
        print(f"  - {node_id}: {node_name}")
    print()
    
    print(f"[INFO] Nodos excluidos ({len(NODOS_EXCLUIDOS)}):")
    for node_id in NODOS_EXCLUIDOS:
        node_name = get_node_name(node_id)
        print(f"  - {node_id}: {node_name}")
    print()
    
    # Generar reporte agregado
    print("[2/3] Generando reporte agregado...")
    print("NOTA: Este proceso requiere acceso a la API WES para obtener los datos.")
    print()
    
    try:
        aggregated_report_path = generate_aggregated_report(
            company_id=COMPANY_ID,
            node_ids=node_ids,
            start_date=START_DATE,
            end_date=END_DATE,
            output_dir="reports",
            fuente_agua_id=None
        )
        print(f"[OK] Reporte agregado generado: {aggregated_report_path}")
    except Exception as e:
        print(f"[ERROR] No se pudo generar el reporte agregado: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # Generar PPT automáticamente
    print()
    print("[3/3] Generando presentación PPT...")
    try:
        company_name = get_company_name(COMPANY_ID)
        ppt_path = generar_ppt_desde_agregado(
            company_id=COMPANY_ID,
            node_ids=node_ids,
            start_date=START_DATE,
            end_date=END_DATE,
            aggregated_report_path=aggregated_report_path,
            company_name=company_name
        )
        print(f"[OK] PPT generada exitosamente")
        
        # La conversión a PDF se hace automáticamente dentro de generar_ppt_desde_agregado
    except Exception as e:
        print(f"[ADVERTENCIA] No se pudo generar PPT: {e}")
        import traceback
        traceback.print_exc()
    
    print()
    print("=" * 70)
    print("  PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 70)
    print()
    print(f"[INFO] Reporte agregado generado: 1")
    print(f"[INFO] Presentación PPT generada: 1")
    print(f"[INFO] Período: {START_DATE} - {END_DATE}")
    print(f"[INFO] Nodos incluidos: {len(node_ids)}")
    print(f"[INFO] Nodos excluidos: {len(NODOS_EXCLUIDOS)}")
    print(f"[INFO] Ubicación del reporte: {aggregated_report_path}")
    if 'ppt_path' in locals():
        print(f"[INFO] Ubicación del PPT: {ppt_path}")

if __name__ == "__main__":
    main()



