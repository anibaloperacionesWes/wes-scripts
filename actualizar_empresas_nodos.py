"""Script para obtener todas las empresas y nodos desde la API y actualizar los diccionarios."""

import requests
import json
import re
from pathlib import Path

ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"

def obtener_empresas_y_nodos():
    """Obtiene todas las empresas y sus nodos desde la API."""
    companies = {}
    all_nodes = {}
    
    print("Obteniendo empresas y nodos desde la API...")
    print("=" * 60)
    
    for i in range(51):  # 000000 a 000050
        company_id = f"{i:06d}"
        url = f"{ENTITY_BASE_URL}/companies/{company_id}"
        
        try:
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                data = response.json()
                company_name = data.get("name", "").strip()
                
                if company_name:
                    companies[company_id] = company_name
                    print(f"[OK] {company_id}: {company_name}")
                    
                    # Obtener nodos de la empresa
                    nodes = data.get("nodes", [])
                    for node in nodes:
                        node_id = node.get("nodeId", "")
                        node_name = node.get("name", "").strip()
                        if node_id and node_name:
                            all_nodes[node_id] = node_name
                            print(f"      - {node_id}: {node_name}")
                else:
                    print(f"[SKIP] {company_id}: Sin nombre")
            elif response.status_code == 404:
                # Empresa no existe, continuar
                pass
            else:
                print(f"[ERROR] {company_id}: Status {response.status_code}")
        except requests.RequestException as e:
            print(f"[ERROR] {company_id}: {e}")
        except Exception as e:
            print(f"[ERROR] {company_id}: {e}")
    
    print("=" * 60)
    print(f"Total empresas encontradas: {len(companies)}")
    print(f"Total nodos encontrados: {len(all_nodes)}")
    
    return companies, all_nodes

def actualizar_diccionarios(companies, nodes):
    """Actualiza los diccionarios en generar_reporte_word.py"""
    script_path = Path("generar_reporte_word.py")
    
    if not script_path.exists():
        print(f"Error: No se encuentra {script_path}")
        return False
    
    content = script_path.read_text(encoding='utf-8')
    
    # Generar nuevo diccionario COMPANY_NAMES
    company_lines = ["COMPANY_NAMES = {"]
    for company_id in sorted(companies.keys()):
        company_name = companies[company_id]
        # Escapar comillas en el nombre
        company_name_escaped = company_name.replace('"', '\\"')
        company_lines.append(f'    "{company_id}": "{company_name_escaped}",')
    company_lines.append("}")
    
    # Generar nuevo diccionario NODE_NAMES
    node_lines = ["NODE_NAMES = {"]
    for node_id in sorted(nodes.keys()):
        node_name = nodes[node_id]
        # Escapar comillas en el nombre
        node_name_escaped = node_name.replace('"', '\\"')
        node_lines.append(f'    "{node_id}": "{node_name_escaped}",')
    node_lines.append("}")
    
    # Buscar y reemplazar COMPANY_NAMES
    company_pattern = r'COMPANY_NAMES\s*=\s*\{[^}]*\}'
    company_replacement = '\n'.join(company_lines)
    content = re.sub(company_pattern, company_replacement, content, flags=re.DOTALL)
    
    # Buscar y reemplazar NODE_NAMES
    node_pattern = r'NODE_NAMES\s*=\s*\{[^}]*\}'
    node_replacement = '\n'.join(node_lines)
    content = re.sub(node_pattern, node_replacement, content, flags=re.DOTALL)
    
    # Guardar archivo actualizado
    script_path.write_text(content, encoding='utf-8')
    print(f"\n[OK] Diccionarios actualizados en {script_path}")
    return True

def main():
    print("=" * 60)
    print("ACTUALIZACIÓN DE EMPRESAS Y NODOS DESDE LA API")
    print("=" * 60)
    print()
    
    # Obtener datos de la API
    companies, nodes = obtener_empresas_y_nodos()
    
    if not companies:
        print("\n[ERROR] No se encontraron empresas. Verifica la conexión a la API.")
        return
    
    # Actualizar diccionarios
    print("\nActualizando diccionarios en generar_reporte_word.py...")
    if actualizar_diccionarios(companies, nodes):
        print("\n[OK] Proceso completado exitosamente")
        
        # Mostrar resumen
        print("\n" + "=" * 60)
        print("RESUMEN")
        print("=" * 60)
        print(f"Empresas actualizadas: {len(companies)}")
        print(f"Nodos actualizados: {len(nodes)}")
        
        # Buscar BUPA
        bupa_companies = {k: v for k, v in companies.items() if 'bupa' in v.lower()}
        if bupa_companies:
            print(f"\nEmpresas BUPA encontradas:")
            for company_id, company_name in bupa_companies.items():
                print(f"  {company_id}: {company_name}")
                # Mostrar nodos de BUPA
                bupa_nodes = {k: v for k, v in nodes.items() if k.startswith(company_id)}
                print(f"    Nodos: {len(bupa_nodes)}")
                for node_id, node_name in sorted(bupa_nodes.items()):
                    print(f"      - {node_id}: {node_name}")
    else:
        print("\n[ERROR] No se pudieron actualizar los diccionarios")

if __name__ == "__main__":
    main()


