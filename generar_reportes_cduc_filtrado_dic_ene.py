"""
Script para generar todos los reportes individuales y el reporte agregado
de CDUC para el período 1 dic 2025 - 8 ene 2026.
Excluye los nodos "Edificio Deportivo" y "Rugby CDUC".
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
    generate_report,
    generate_aggregated_report,
)

# Configuración
COMPANY_ID = "000021"  # CDUC
START_DATE = "2025-12-01"
END_DATE = "2026-01-08"
ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"

# Nodos a excluir
NODOS_EXCLUIDOS = [
    "Edificio Deportivo",
    "Rugby CDUC"
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

def filter_nodes(all_nodes: list) -> list:
    """
    Filtra los nodos excluyendo los especificados en NODOS_EXCLUIDOS.
    """
    filtered_nodes = []
    excluded_nodes = []
    
    for node in all_nodes:
        node_name = node["name"]
        if node_name in NODOS_EXCLUIDOS:
            excluded_nodes.append(node)
        else:
            filtered_nodes.append(node)
    
    if excluded_nodes:
        print(f"[INFO] Nodos excluidos ({len(excluded_nodes)}):")
        for node in excluded_nodes:
            print(f"  - {node['nodeId']}: {node['name']}")
        print()
    
    return filtered_nodes

def main():
    print("=" * 60)
    print("  GENERACIÓN DE REPORTES CDUC - DIC 2025 - ENE 2026")
    print("  (Excluyendo Edificio Deportivo y Rugby CDUC)")
    print("=" * 60)
    print()
    
    # Obtener información de la empresa
    company_name = get_company_name(COMPANY_ID)
    print(f"[INFO] Empresa: {company_name} (ID: {COMPANY_ID})")
    print()
    
    # Obtener todos los nodos
    print(f"[1/4] Obteniendo nodos de {company_name}...")
    all_nodes = get_company_nodes(COMPANY_ID)
    
    if not all_nodes:
        print("[ERROR] No se encontraron nodos para la empresa.")
        return
    
    print(f"[OK] Se encontraron {len(all_nodes)} nodo(s) en total")
    print()
    
    # Filtrar nodos excluidos
    print(f"[2/4] Filtrando nodos (excluyendo: {', '.join(NODOS_EXCLUIDOS)})...")
    nodes = filter_nodes(all_nodes)
    
    if not nodes:
        print("[ERROR] No quedaron nodos después del filtrado.")
        return
    
    print(f"[OK] Se procesarán {len(nodes)} nodo(s) después del filtrado")
    for node in nodes:
        print(f"  - {node['nodeId']}: {node['name']}")
    print()
    
    # Generar reportes individuales
    print(f"[3/4] Generando reportes individuales...")
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
                output_dir="reports",  # Directorio base para los reportes
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
    print(f"[4/4] Generando reporte agregado...")
    
    try:
        node_ids = [node["nodeId"] for node in nodes]
        aggregated_report_path = generate_aggregated_report(
            company_id=COMPANY_ID,
            node_ids=node_ids,
            start_date=START_DATE,
            end_date=END_DATE,
            output_dir="reports",
            fuente_agua_id=None  # Verificar si CDUC tiene fuente de agua específica
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
    print(f"[INFO] Nodos incluidos: {len(nodes)}")
    print(f"[INFO] Nodos excluidos: {len(NODOS_EXCLUIDOS)} ({', '.join(NODOS_EXCLUIDOS)})")
    print(f"[INFO] No se enviaron correos electrónicos (según solicitud)")

if __name__ == "__main__":
    main()




