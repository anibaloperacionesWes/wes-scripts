"""
Script para generar todos los reportes individuales y el reporte agregado
de AGUNSA desde el 01 de diciembre 2025 hasta la fecha actual.
No envía correos electrónicos.
"""

import sys
from pathlib import Path
from datetime import datetime
import requests
import argparse

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import (
    get_company_name,
    get_node_name,
    generate_report,
    generate_aggregated_report,
)

# Configuración
COMPANY_ID = "000020"  # AGUNSA
START_DATE = "01/12/2025"  # 01 de diciembre 2025
END_DATE = datetime.now().strftime("%d/%m/%Y")  # Fecha actual
ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"

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

def main():
    print("=" * 70)
    print("  GENERACIÓN DE REPORTES AGUNSA")
    print("  Período: 01 de diciembre 2025 - Fecha actual")
    print("=" * 70)
    print()
    
    # Obtener información de la empresa
    company_name = get_company_name(COMPANY_ID)
    print(f"[INFO] Empresa: {company_name} (ID: {COMPANY_ID})")
    print(f"[INFO] Período: {START_DATE} - {END_DATE}")
    print()
    
    # Obtener todos los nodos
    print(f"[1/3] Obteniendo nodos de {company_name}...")
    nodes = get_company_nodes(COMPANY_ID)
    
    if not nodes:
        print("[ERROR] No se encontraron nodos para la empresa.")
        return
    
    print(f"[OK] Se encontraron {len(nodes)} nodo(s)")
    for node in nodes:
        print(f"  - {node['nodeId']}: {node['name']}")
    print()
    
    # Generar reportes individuales (el agregado va después y solo con nodos exitosos)
    print(f"[2/3] Generando reportes individuales...")
    generated_reports = []
    node_ids_exitosos = []
    
    for i, node in enumerate(nodes, 1):
        node_id = node["nodeId"]
        node_name = node["name"]
        
        print(f"  [{i}/{len(nodes)}] Generando reporte para {node_id} ({node_name})...", flush=True)
        
        try:
            # Crear argumentos para generate_report usando argparse.Namespace
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
                node_ids_exitosos.append(node_id)
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
    
    if not node_ids_exitosos:
        print("[ERROR] No hay nodos con reporte individual exitoso; no se genera el agregado.")
        return
    
    # Generar reporte agregado solo después de los individuales y solo con esos nodos
    print(f"[3/3] Generando reporte agregado...")
    
    try:
        aggregated_report_path = generate_aggregated_report(
            company_id=COMPANY_ID,
            node_ids=node_ids_exitosos,
            start_date=START_DATE,
            end_date=END_DATE,
            output_dir="reports",
            fuente_agua_id=None  # Sin fuente de agua específica
        )
        print(f"[OK] Reporte agregado generado: {aggregated_report_path}")
    except Exception as e:
        print(f"[ERROR] No se pudo generar el reporte agregado: {e}")
        import traceback
        traceback.print_exc()
        return
    
    print()
    print("=" * 70)
    print("  PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 70)
    print()
    print(f"[INFO] Reportes individuales generados: {len(generated_reports)}")
    print(f"[INFO] Reporte agregado generado: 1")
    print(f"[INFO] Período: {START_DATE} - {END_DATE}")
    print(f"[INFO] No se enviaron correos electrónicos (según solicitud)")

if __name__ == "__main__":
    main()



