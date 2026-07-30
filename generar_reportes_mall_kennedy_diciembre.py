"""
Script para generar todos los reportes individuales y el reporte agregado
de Mall Kennedy (Parque Arauco) para el mes de diciembre 2025.
No envía correos electrónicos.
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
    get_mall_name_for_parque_arauco,
    generate_report,
    generate_aggregated_report,
)

# Configuración
COMPANY_ID = "000025"  # Parque Arauco
START_DATE = "2025-12-01"
END_DATE = "2025-12-31"
ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"

# Nodos de Kennedy (basado en el mapeo del código)
NODOS_KENNEDY = [
    "000025-20",  # PAK Impulsión Ander3-4 Matriz Principal
    "000025-21",  # PAK Impulsión Ander3-4 Locales Gast.
    "000025-22",  # PAK Impulsión Sandia Baños 2-3-6-7 Fredo
    "000025-23",  # PAK Llenado Pileta
    "000025-24",  # PAK Llenado Pileta Cascada
    "000025-35",  # PAK BAZAR GOURMET (reemplazo 000025-25)
    "000025-36",  # PAK DL KENNEDY (reemplazo 000025-26)
    "000025-27",  # PAK Distrito de lujo DL
    "000025-28",  # PAK Impulsión Mall 1 Piso-4
    "000025-29",  # PAK Impulsión Anden 3-4 Restaurante
]

def get_company_nodes(company_id: str) -> list:
    """
    Obtiene todos los nodos de una empresa.
    Retorna lista de diccionarios con nodeId y name.
    """
    url = f"{ENTITY_BASE_URL}/companies/{company_id}"
    
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            nodes = data.get("nodes", [])
            return [
                {
                    "nodeId": node.get("nodeId", ""),
                    "name": node.get("name", "").strip()
                }
                for node in nodes
                if node.get("nodeId") and node.get("name")
            ]
        else:
            print(f"[ERROR] No se pudo obtener información de la empresa {company_id}: {response.status_code}")
            return []
    except Exception as e:
        print(f"[ERROR] Error al obtener nodos de la empresa {company_id}: {e}")
        return []

def filter_kennedy_nodes(all_nodes: list) -> list:
    """
    Filtra solo los nodos que pertenecen a Kennedy.
    Usa el mapeo de nodos conocidos y también verifica el nombre del mall.
    """
    kennedy_nodes = []
    
    for node in all_nodes:
        node_id = node["nodeId"]
        node_name = node["name"]
        
        # Verificar si está en la lista conocida de Kennedy
        if node_id in NODOS_KENNEDY:
            kennedy_nodes.append(node)
        else:
            # Si no está en la lista, verificar por el nombre del mall
            mall = get_mall_name_for_parque_arauco(node_id, node_name)
            if mall == "Kennedy":
                kennedy_nodes.append(node)
    
    return kennedy_nodes

def main():
    print("=" * 60)
    print("  GENERACIÓN DE REPORTES MALL KENNEDY - DICIEMBRE 2025")
    print("=" * 60)
    print()
    
    # Obtener información de la empresa
    company_name = get_company_name(COMPANY_ID)
    print(f"[INFO] Empresa: {company_name} (ID: {COMPANY_ID})")
    print()
    
    # Obtener todos los nodos
    print(f"[1/3] Obteniendo nodos de {company_name}...")
    all_nodes = get_company_nodes(COMPANY_ID)
    
    if not all_nodes:
        print("[ERROR] No se encontraron nodos para la empresa.")
        return
    
    print(f"[OK] Se encontraron {len(all_nodes)} nodo(s) en total")
    print()
    
    # Filtrar solo nodos de Kennedy
    print(f"[INFO] Filtrando nodos de Mall Kennedy...")
    nodes = filter_kennedy_nodes(all_nodes)
    
    if not nodes:
        print("[ERROR] No se encontraron nodos de Kennedy.")
        print("[INFO] Nodos conocidos de Kennedy esperados:")
        for node_id in NODOS_KENNEDY:
            print(f"  - {node_id}")
        return
    
    print(f"[OK] Se encontraron {len(nodes)} nodo(s) de Kennedy")
    for node in nodes:
        print(f"  - {node['nodeId']}: {node['name']}")
    print()
    
    # Generar reportes individuales
    print(f"[2/3] Generando reportes individuales...")
    generated_reports = []
    
    for i, node in enumerate(nodes, 1):
        node_id = node["nodeId"]
        node_name = node["name"]
        
        print(f"  [{i}/{len(nodes)}] Generando reporte para {node_id} ({node_name})...", flush=True)
        
        try:
            # Crear argumentos para generate_report usando argparse.Namespace
            import argparse
            args = argparse.Namespace(
                company_id=COMPANY_ID,
                node_id=node_id,
                start_date=START_DATE,
                end_date=END_DATE,
                output_dir=None,  # Se generará automáticamente
                enviar_correo=False  # No enviar correos
            )
            report_path = generate_report(args)
            if report_path:
                generated_reports.append(report_path)
                print(f"    [OK] Reporte generado: {report_path}", flush=True)
            else:
                print(f"    [ADVERTENCIA] generate_report retornó None para {node_id}", flush=True)
        except Exception as e:
            print(f"    [ERROR] No se pudo generar el reporte para {node_id}: {e}", flush=True)
            import traceback
            traceback.print_exc()
            continue
    
    print()
    print(f"[OK] Se generaron {len(generated_reports)} reporte(s) individual(es)")
    print()
    
    # Generar reporte agregado
    print(f"[3/3] Generando reporte agregado...")
    
    try:
        node_ids = [node["nodeId"] for node in nodes]
        aggregated_report_path = generate_aggregated_report(
            company_id=COMPANY_ID,
            node_ids=node_ids,
            start_date=START_DATE,
            end_date=END_DATE,
            output_dir="reports",
            fuente_agua_id=None  # Todos los puntos son consumidores (sin fuente de agua)
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
    print(f"[INFO] Reportes individuales generados: {len(generated_reports)}")
    print(f"[INFO] Reporte agregado generado: 1")
    print(f"[INFO] Periodo: {START_DATE} a {END_DATE}")
    print(f"[INFO] Mall: Kennedy")
    print(f"[INFO] No se enviaron correos electrónicos (según solicitud)")

if __name__ == "__main__":
    main()

