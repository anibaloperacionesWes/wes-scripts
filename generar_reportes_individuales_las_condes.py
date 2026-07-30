"""
Script para generar solo reportes individuales de Las Condes (sin presentación ni agregado)
Período: 01-01-2026 al 31-01-2026
"""

import sys
from pathlib import Path
import argparse
import requests

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import (
    get_company_name,
    get_node_name,
    generate_report,
    ENTITY_BASE_URL,
)

# Configuración
COMPANY_ID = "000022"  # Las Condes
START_DATE = "01/01/2026"
END_DATE = "31/01/2026"

def get_las_condes_nodes(company_id: str) -> list:
    """Obtiene todos los nodos de Las Condes."""
    url = f"{ENTITY_BASE_URL}/companies/{company_id}"
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            all_nodes = data.get("nodes", [])
            nodes_encontrados = []
            
            for node in all_nodes:
                node_id = node.get("nodeId", "")
                node_name = node.get("name", "").strip()
                nodes_encontrados.append({
                    "nodeId": node_id,
                    "name": node_name
                })
                print(f"  [OK] Nodo encontrado: {node_id} - {node_name}")
            
            return nodes_encontrados
        return []
    except Exception as e:
        print(f"[ERROR] Error al obtener nodos: {e}")
        return []

def main():
    print("=" * 70)
    print("  GENERACIÓN DE REPORTES INDIVIDUALES - LAS CONDES")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print("=" * 70)
    print()
    
    # Obtener información de la empresa
    try:
        company_name = get_company_name(COMPANY_ID)
        print(f"[INFO] Empresa: {company_name} (ID: {COMPANY_ID})")
    except Exception as e:
        print(f"[ERROR] No se pudo obtener el nombre de la empresa: {e}")
        print(f"[INFO] Continuando con ID: {COMPANY_ID}")
        company_name = "Las Condes"
    print()
    
    # Obtener todos los nodos
    print("[1/2] Obteniendo nodos de Las Condes...")
    nodes = get_las_condes_nodes(COMPANY_ID)
    
    if not nodes:
        print("[ERROR] No se encontraron nodos de Las Condes.")
        print(f"[INFO] Verificando si el COMPANY_ID {COMPANY_ID} es correcto...")
        return
    
    print(f"[OK] Se encontraron {len(nodes)} nodo(s)")
    for node in nodes:
        print(f"  - {node['nodeId']}: {node['name']}")
    print()
    
    # Generar reportes individuales
    print("[2/2] Generando reportes individuales...")
    generated_reports = []
    
    for i, node in enumerate(nodes, 1):
        node_id = node["nodeId"]
        node_name = node["name"]
        
        print(f"  [{i}/{len(nodes)}] Generando reporte para {node_id} ({node_name})...", flush=True)
        
        try:
            args = argparse.Namespace(
                company_id=COMPANY_ID,
                node_id=node_id,
                start_date=START_DATE,
                end_date=END_DATE,
                output_dir="reports",
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
    print("=" * 70)
    print("  PROCESO COMPLETADO")
    print("=" * 70)
    print(f"[OK] Reportes individuales generados: {len(generated_reports)}/{len(nodes)}")
    for report_path in generated_reports:
        print(f"  - {report_path}")

if __name__ == "__main__":
    main()
