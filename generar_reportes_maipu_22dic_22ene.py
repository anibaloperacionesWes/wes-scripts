"""
Script para generar todos los reportes individuales, agregado y PPT de Mall Maipú
Período: 22 de diciembre 2025 al 22 de enero 2026
Incluye también el nodo 000025-32
"""

import sys
from pathlib import Path
from datetime import datetime
from typing import Optional
import requests
import argparse
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from io import BytesIO

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import (
    get_company_name,
    get_node_name,
    generate_report,
    generate_aggregated_report,
    get_mall_name_for_parque_arauco,
    BASE_URL,
    ENTITY_BASE_URL,
    fetch_json,
    normalize_measures_payload,
    flatten_measures,
    summarize_consumption,
    parse_date,
    format_number_chilean,
    format_currency_chilean,
    get_water_price_per_m3,
)

# Importar funciones del script de Maipú
from generar_reportes_y_ppt_mall_maipu import (
    get_maipu_nodes,
    obtener_datos_agregados,
    crear_ppt_analisis,
    convertir_ppt_a_pdf,
    guardar_datos_json,
    cargar_datos_json,
)

COMPANY_ID = "000025"  # Parque Arauco
START_DATE = "22/12/2025"
END_DATE = "22/01/2026"
MALL_KEY = "maipu"

def main():
    print("=" * 70)
    print("  GENERACIÓN DE REPORTES Y PPT - MALL MAIPÚ")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print("=" * 70)
    print()
    
    # Obtener nodos de Maipú
    print("[1/5] Obteniendo nodos del Mall Maipú...")
    nodes = get_maipu_nodes(COMPANY_ID)
    
    if not nodes:
        print("[ERROR] No se encontraron nodos del Mall Maipú.")
        return
    
    # Agregar el nodo 000025-32 si no está ya incluido
    node_ids_list = [node["nodeId"] for node in nodes]
    if "000025-32" not in node_ids_list:
        print("[INFO] Agregando nodo 000025-32 a la lista...")
        # Obtener el nombre del nodo desde la API
        try:
            url = f"{ENTITY_BASE_URL}/companies/{COMPANY_ID}"
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                data = response.json()
                all_nodes = data.get("nodes", [])
                for node in all_nodes:
                    if node.get("nodeId") == "000025-32":
                        nodes.append({
                            "nodeId": "000025-32",
                            "name": node.get("name", "000025-32")
                        })
                        print(f"  [OK] Nodo agregado: 000025-32 ({node.get('name', '000025-32')})")
                        break
        except Exception as e:
            print(f"  [ADVERTENCIA] No se pudo obtener el nombre del nodo 000025-32: {e}")
            # Agregar de todas formas
            nodes.append({
                "nodeId": "000025-32",
                "name": "000025-32"
            })
            print(f"  [OK] Nodo agregado: 000025-32")
    
    print(f"[OK] Total de nodos a procesar: {len(nodes)}")
    node_ids = [node["nodeId"] for node in nodes]
    for node in nodes:
        print(f"  - {node['nodeId']}: {node['name']}")
    print()
    
    # Generar reportes individuales
    print("[2/5] Generando reportes individuales...")
    generated_reports = []
    
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
                output_dir="reports",
                enviar_correo=False
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
    
    # Obtener datos para reporte agregado y PPT
    print("[3/5] Obteniendo datos para análisis agregado...")
    try:
        datos = obtener_datos_agregados(node_ids, START_DATE, END_DATE)
        print(f"[OK] Datos obtenidos: {len(datos['all_measures'])} medidas")
    except Exception as e:
        print(f"[ERROR] Error al obtener datos: {e}")
        import traceback
        traceback.print_exc()
        return
    
    print()
    
    # Generar reporte agregado
    print("[4/5] Generando reporte agregado...")
    aggregated_report_path = None
    try:
        aggregated_report_path = generate_aggregated_report(
            company_id=COMPANY_ID,
            node_ids=node_ids,
            start_date=START_DATE,
            end_date=END_DATE,
            output_dir="reports",
            fuente_agua_id=None,
            mall_name="Maipú"
        )
        print(f"[OK] Reporte agregado: {aggregated_report_path}")
        
        # Guardar datos en JSON para futuras generaciones de PPT sin API
        if aggregated_report_path:
            if aggregated_report_path.is_dir():
                json_path = aggregated_report_path / "datos_agregados.json"
            else:
                json_path = aggregated_report_path.parent / "datos_agregados.json"
            guardar_datos_json(datos, json_path)
    except Exception as e:
        print(f"[ADVERTENCIA] Error al generar reporte agregado: {e}")
        print("[INFO] Continuando con la generación de PPT...")
        import traceback
        traceback.print_exc()
        # Crear directorio para PPT si no existe el reporte agregado
        from datetime import datetime as dt
        company_name = get_company_name(COMPANY_ID)
        safe_company_name = "".join(c for c in company_name if c.isalnum() or c in (" ", "-", "_")).strip().replace(" ", "_")
        timestamp = dt.now().strftime("%Y%m%d_%H%M")
        output_dir_base = Path("reports") / safe_company_name / "Maipú" / "ABREGADO"
        output_dir_base.mkdir(parents=True, exist_ok=True)
        ppt_dir = output_dir_base / f"AGREGADO_{timestamp}"
        ppt_dir.mkdir(exist_ok=True)
        aggregated_report_path = ppt_dir  # Usar este directorio para la PPT
    
    print()
    
    # Generar PPT
    print("[5/5] Generando presentación PPT...")
    try:
        # Guardar PPT en la misma carpeta que el reporte agregado
        if aggregated_report_path:
            if aggregated_report_path.is_dir():
                ppt_dir = aggregated_report_path
            else:
                ppt_dir = aggregated_report_path.parent
        else:
            # Fallback: crear directorio por defecto
            from datetime import datetime as dt
            company_name = get_company_name(COMPANY_ID)
            safe_company_name = "".join(c for c in company_name if c.isalnum() or c in (" ", "-", "_")).strip().replace(" ", "_")
            timestamp = dt.now().strftime("%Y%m%d_%H%M")
            output_dir_base = Path("reports") / safe_company_name / "Maipú" / "ABREGADO"
            output_dir_base.mkdir(parents=True, exist_ok=True)
            ppt_dir = output_dir_base / f"AGREGADO_{timestamp}"
            ppt_dir.mkdir(exist_ok=True)
        
        ppt_path = ppt_dir / "Agregado PPT.pptx"
        crear_ppt_analisis(datos, ppt_path, mall_name="Maipú")
        
        # Convertir PPT a PDF
        pdf_path = convertir_ppt_a_pdf(ppt_path)
        
        print()
        print("=" * 70)
        print("  GENERACIÓN COMPLETADA EXITOSAMENTE")
        print("=" * 70)
        print(f"[OK] Reportes individuales generados: {len(generated_reports)}")
        if aggregated_report_path:
            print(f"[OK] Reporte agregado: {aggregated_report_path}")
        print(f"[OK] PPT generada en: {ppt_path}")
        if pdf_path:
            print(f"[OK] PDF generado en: {pdf_path}")
        print()
    except Exception as e:
        print(f"[ERROR] Error al generar PPT: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
