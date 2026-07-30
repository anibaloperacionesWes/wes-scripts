"""
Script para generar y enviar reportes individuales y agregado de La Florida
Período: 25-12-2025 al 25-01-2026
Destinatarios: juan, diego, jose
"""

import sys
from pathlib import Path
from datetime import datetime
import argparse
import requests

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import (
    get_company_name,
    get_node_name,
    generate_report,
    generate_aggregated_report,
    enviar_reporte_por_correo,
    ENTITY_BASE_URL,
)

from config_correos_equipo import obtener_correos_por_rol

# Configuración
COMPANY_ID = "000028"  # La Florida
START_DATE = "25/12/2025"
END_DATE = "25/01/2026"

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

# Destinatarios: juan, diego, jose
DESTINATARIOS = [
    "juanlopez@wes.cl",      # Juan
    "diegocarrasco@wes.cl",  # Diego
    "joseotarola@wes.cl",    # Jose
]

def get_florida_nodes(company_id: str) -> list:
    """Obtiene todos los nodos de La Florida."""
    url = f"{ENTITY_BASE_URL}/companies/{company_id}"
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            all_nodes = data.get("nodes", [])
            nodes = []
            for node in all_nodes:
                node_id = node.get("nodeId", "")
                node_name = node.get("name", "").strip()
                if node_id and node_name:
                    nodes.append({
                        "nodeId": node_id,
                        "name": node_name
                    })
            return nodes
        return []
    except Exception as e:
        print(f"[ERROR] Error al obtener nodos: {e}")
        return []

def main():
    print("=" * 70)
    print("  GENERACIÓN Y ENVÍO DE REPORTES - LA FLORIDA")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print("=" * 70)
    print()
    
    # Obtener información de la empresa
    company_name = get_company_name(COMPANY_ID)
    print(f"[INFO] Empresa: {company_name} (ID: {COMPANY_ID})")
    print()
    
    # Obtener nodos (intentar desde API, si falla usar nodo conocido)
    print("[1/4] Obteniendo nodos de La Florida...")
    nodes = get_florida_nodes(COMPANY_ID)
    
    # Si no se obtuvieron nodos de la API, usar el nodo conocido
    if not nodes:
        print("[INFO] No se pudieron obtener nodos de la API, usando nodo conocido...")
        nodes = [{
            "nodeId": "000028-01",
            "name": "Liceo Alto Cordillera"
        }]
    
    print(f"[OK] Se encontraron {len(nodes)} nodo(s)")
    node_ids = [node["nodeId"] for node in nodes]
    for node in nodes:
        print(f"  - {node['nodeId']}: {node['name']}")
    print()
    
    # Generar reportes individuales
    print("[2/3] Generando reporte individual...")
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
                enviar_correo=False  # No enviar correos automáticamente
            )
            report_path = generate_report(args)
            if report_path:
                generated_reports.append({
                    "path": report_path,
                    "node_id": node_id,
                    "node_name": node_name
                })
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
    
    # Formatear fechas para el correo
    start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
    end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
    start_date_str = start_dt.strftime("%d-%m-%y")
    end_date_str = end_dt.strftime("%d-%m-%y")
    
    # Enviar reportes por correo
    print("[3/3] Enviando reporte individual por correo...")
    print(f"  Destinatarios: {', '.join(DESTINATARIOS)}")
    print()
    
    # Enviar reportes individuales
    for report_info in generated_reports:
        report_path = report_info["path"]
        node_name = report_info["node_name"]
        
        print(f"  Enviando reporte individual: {node_name}...")
        for destinatario in DESTINATARIOS:
            try:
                exito = enviar_reporte_por_correo(
                    reporte_path=report_path,
                    destinatario=destinatario,
                    smtp_servidor=SMTP_SERVIDOR,
                    smtp_puerto=SMTP_PUERTO,
                    smtp_usuario=SMTP_USUARIO,
                    smtp_password=SMTP_PASSWORD,
                    company_name=company_name,
                    node_name=node_name,
                    start_date=start_date_str,
                    end_date=end_date_str,
                )
                if exito:
                    print(f"    [OK] Enviado a {destinatario}")
                else:
                    print(f"    [ERROR] No se pudo enviar a {destinatario}")
            except Exception as e:
                print(f"    [ERROR] Error al enviar a {destinatario}: {e}")
        print()
    
    print()
    print("=" * 70)
    print("  PROCESO COMPLETADO")
    print("=" * 70)
    print(f"[OK] Reporte individual generado: {len(generated_reports)}")
    print(f"[OK] Reporte enviado a: {len(DESTINATARIOS)} destinatario(s)")
    print(f"[INFO] Periodo: {START_DATE} a {END_DATE}")
    print(f"[INFO] Empresa: {company_name}")

if __name__ == "__main__":
    main()
