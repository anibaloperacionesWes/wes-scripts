"""
Script para generar solo reportes individuales de Providencia (sin presentación ni agregado)
Período: 01-01-2026 al 31-01-2026
Nodos: 000006-01, 000006-02, 000006-04, 000006-05
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
COMPANY_ID = "000006"  # Providencia
START_DATE = "01/01/2026"
END_DATE = "31/01/2026"

# IDs de nodos solicitados explícitamente
NODE_IDS_SOLICITADOS = {
    "000006-01",
    "000006-02",
    "000006-04",
    "000006-05",
}

# Nombres de nodos a buscar (parciales) solo como ayuda para mapear IDs -> nombres
NODOS_SOLICITADOS = [
    "liceo lastarria",
    "carmela carvajal",
    "liceo 7",
    "liceo juan pablo duarte",
]

def get_providencia_nodes(company_id: str, nombres_buscar: list) -> list:
    """Obtiene los nodos de Providencia que coincidan con los nombres solicitados."""
    url = f"{ENTITY_BASE_URL}/companies/{company_id}"
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            all_nodes = data.get("nodes", [])
            nodes_encontrados = []
            
            for node in all_nodes:
                node_id = node.get("nodeId", "")
                node_name = node.get("name", "").strip().lower()
                
                # Buscar coincidencias parciales
                for nombre_buscar in nombres_buscar:
                    nombre_buscar_lower = nombre_buscar.lower()
                    # Verificar si el nombre del nodo contiene el nombre buscado o viceversa
                    if (nombre_buscar_lower in node_name or 
                        node_name in nombre_buscar_lower or
                        any(palabra in node_name for palabra in nombre_buscar_lower.split() if len(palabra) > 3)):
                        nodes_encontrados.append({
                            "nodeId": node_id,
                            "name": node.get("name", "").strip()
                        })
                        print(f"  [OK] Nodo encontrado: {node_id} - {node.get('name', '').strip()}")
                        break
            
            return nodes_encontrados
        return []
    except Exception as e:
        print(f"[ERROR] Error al obtener nodos: {e}")
        return []

def main():
    print("=" * 70)
    print("  GENERACIÓN DE REPORTES INDIVIDUALES - PROVIDENCIA")
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
        company_name = "Providencia"
    print()
    
    # Obtener nodos solicitados (por nombre) y luego filtrar por ID
    print("[1/2] Obteniendo nodos de Providencia...")
    nodes = get_providencia_nodes(COMPANY_ID, NODOS_SOLICITADOS)
    # Filtrar solo los IDs requeridos explícitamente
    nodes = [n for n in nodes if n["nodeId"] in NODE_IDS_SOLICITADOS]
    
    if not nodes:
        print("[ERROR] No se encontraron nodos para los nombres solicitados.")
        print("[INFO] IDs de nodos requeridos:")
        for nid in sorted(NODE_IDS_SOLICITADOS):
            print(f"  - {nid}")
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
