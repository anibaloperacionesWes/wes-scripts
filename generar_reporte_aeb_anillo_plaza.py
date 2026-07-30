"""
Script para generar reporte individual de Mall AEB - Anillo plaza
Período: 01 de diciembre 2025 al 16 de enero 2026
"""

import sys
from pathlib import Path
import requests
import argparse

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import (
    get_company_name,
    get_node_name,
    generate_report,
)

ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
START_DATE = "01/12/2025"
END_DATE = "16/01/2026"

def buscar_empresa_por_nombre(nombre_buscado: str) -> tuple[str, str] | None:
    """
    Busca una empresa por nombre en el rango 000001-000100.
    Retorna (company_id, company_name) si la encuentra, None si no.
    """
    print(f"[INFO] Buscando empresa '{nombre_buscado}'...")
    
    for i in range(1, 101):
        company_id = f"{i:06d}"
        try:
            url = f"{ENTITY_BASE_URL}/companies/{company_id}"
            response = requests.get(url, timeout=5)
            if response.status_code == 200:
                data = response.json()
                company_name = data.get("name", "").strip()
                if nombre_buscado.lower() in company_name.lower() or company_name.lower() in nombre_buscado.lower():
                    print(f"[OK] Empresa encontrada: ID {company_id}, Nombre: {company_name}")
                    return (company_id, company_name)
        except Exception as e:
            continue
    
    print(f"[ERROR] No se encontró la empresa '{nombre_buscado}'")
    return None

def buscar_nodo_por_nombre(company_id: str, nombre_nodo: str) -> tuple[str, str] | None:
    """
    Busca un nodo por nombre dentro de una empresa.
    Retorna (node_id, node_name) si lo encuentra, None si no.
    """
    print(f"[INFO] Buscando nodo '{nombre_nodo}' en empresa {company_id}...")
    
    try:
        url = f"{ENTITY_BASE_URL}/companies/{company_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            nodes = data.get("nodes", [])
            
            for node in nodes:
                node_id = node.get("nodeId", "")
                node_name = node.get("name", "").strip()
                if nombre_nodo.lower() in node_name.lower() or node_name.lower() in nombre_nodo.lower():
                    print(f"[OK] Nodo encontrado: ID {node_id}, Nombre: {node_name}")
                    return (node_id, node_name)
            
            print(f"[ERROR] No se encontró el nodo '{nombre_nodo}' en la empresa {company_id}")
            print(f"[INFO] Nodos disponibles:")
            for node in nodes:
                print(f"  - {node.get('nodeId', '')}: {node.get('name', '')}")
            return None
        else:
            print(f"[ERROR] No se pudo obtener información de la empresa {company_id}: {response.status_code}")
            return None
    except Exception as e:
        print(f"[ERROR] Error al buscar nodo: {e}")
        return None

def main():
    print("=" * 70)
    print("  GENERACIÓN DE REPORTE - MALL AEB - ANILLO PLAZA")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print("=" * 70)
    print()
    
    # Buscar empresa "Mall AEB"
    empresa_info = buscar_empresa_por_nombre("Mall AEB")
    if not empresa_info:
        # Intentar buscar solo "AEB"
        empresa_info = buscar_empresa_por_nombre("AEB")
    
    if not empresa_info:
        print("[ERROR] No se pudo encontrar la empresa. Por favor, verifica el nombre.")
        return
    
    company_id, company_name = empresa_info
    print()
    
    # Buscar nodo "Anillo plaza"
    nodo_info = buscar_nodo_por_nombre(company_id, "Anillo plaza")
    if not nodo_info:
        # Intentar buscar solo "Anillo"
        nodo_info = buscar_nodo_por_nombre(company_id, "Anillo")
    
    if not nodo_info:
        print("[ERROR] No se pudo encontrar el nodo. Por favor, verifica el nombre.")
        return
    
    node_id, node_name = nodo_info
    print()
    
    # Generar reporte
    print(f"[INFO] Generando reporte individual...")
    print(f"  Empresa: {company_name} (ID: {company_id})")
    print(f"  Nodo: {node_name} (ID: {node_id})")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print()
    
    try:
        args = argparse.Namespace(
            company_id=company_id,
            node_id=node_id,
            start_date=START_DATE,
            end_date=END_DATE,
            output_dir="reports",
            enviar_correo=False
        )
        
        report_path = generate_report(args)
        
        if report_path:
            print()
            print("=" * 70)
            print("  REPORTE GENERADO EXITOSAMENTE")
            print("=" * 70)
            print(f"[OK] Reporte generado: {report_path}")
        else:
            print("[ERROR] No se pudo generar el reporte")
    except Exception as e:
        print(f"[ERROR] Error al generar el reporte: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
