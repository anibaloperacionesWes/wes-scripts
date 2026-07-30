"""
Script para generar reporte agregado de Parque Arauco Estación y PPT
Período: 01 de enero 2026 hasta hoy
"""

import sys
from pathlib import Path
from datetime import datetime
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
    obtener_datos_agregados,
    guardar_datos_json,
    cargar_datos_json,
    crear_ppt_analisis,
    get_estacion_nodes,
)

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.dml import MSO_LINE
except ImportError:
    print("[ERROR] Se requiere python-pptx. Instálalo con: pip install python-pptx")
    sys.exit(1)

COMPANY_ID = "000025"  # Parque Arauco
# Fechas: desde 01 de diciembre 2025 al 18 de enero 2026
START_DATE = "01/12/2025"
END_DATE = "18/01/2026"

def buscar_ultimo_reporte_agregado_estacion() -> Path:
    """Busca el último reporte agregado de Parque Arauco Estación."""
    reports_dir = Path("reports") / "Parque_Arauco" / "ABREGADO"
    if not reports_dir.exists():
        return None
    
    # Buscar directorios que contengan "ESTACION" o "Estación" en el nombre
    dirs = [d for d in reports_dir.iterdir() if d.is_dir() and ("ESTACION" in d.name.upper() or "ESTACIÓN" in d.name.upper())]
    if not dirs:
        return None
    
    # Ordenar por fecha de modificación (más reciente primero)
    dirs.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return dirs[0]

def main():
    print("=" * 70)
    print("  GENERACIÓN DE REPORTES Y PPT - PARQUE ARAUCO ESTACIÓN")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print("=" * 70)
    print()
    
    # Obtener nodos de Estación
    print("[1/5] Obteniendo nodos de Parque Arauco Estación...")
    nodes = get_estacion_nodes(COMPANY_ID)
    
    if not nodes:
        print("[ERROR] No se encontraron nodos de Parque Arauco Estación.")
        return
    
    print(f"[OK] Se encontraron {len(nodes)} nodo(s)")
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
            fuente_agua_id=None
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
        company_name = get_company_name(COMPANY_ID)
        safe_company_name = "".join(c for c in company_name if c.isalnum() or c in (" ", "-", "_")).strip().replace(" ", "_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_dir_base = Path("reports") / safe_company_name / "ABREGADO"
        output_dir_base.mkdir(parents=True, exist_ok=True)
        ppt_dir = output_dir_base / f"AGREGADO_ESTACION_{timestamp}"
        ppt_dir.mkdir(exist_ok=True)
        aggregated_report_path = ppt_dir
    
    print()
    
    # Generar PPT (siempre se genera)
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
            company_name = get_company_name(COMPANY_ID)
            safe_company_name = "".join(c for c in company_name if c.isalnum() or c in (" ", "-", "_")).strip().replace(" ", "_")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            output_dir_base = Path("reports") / safe_company_name / "ABREGADO"
            output_dir_base.mkdir(parents=True, exist_ok=True)
            ppt_dir = output_dir_base / f"AGREGADO_ESTACION_{timestamp}"
            ppt_dir.mkdir(exist_ok=True)
        
        ppt_path = ppt_dir / "Agregado PPT.pptx"
        crear_ppt_analisis(datos, ppt_path, mall_name="Estación")
        
        # Convertir PPT a PDF
        from generar_reportes_y_ppt_mall_maipu import convertir_ppt_a_pdf
        pdf_path = convertir_ppt_a_pdf(ppt_path)
        
        print()
        print("=" * 70)
        print("  PROCESO COMPLETADO")
        print("=" * 70)
        print(f"[OK] Reportes individuales generados: {len(generated_reports)}")
        print(f"[OK] Reporte agregado: {aggregated_report_path}")
        print(f"[OK] Presentación PPT: {ppt_path}")
        if pdf_path:
            print(f"[OK] Presentación PDF: {pdf_path}")
    except Exception as e:
        print(f"[ERROR] Error al generar PPT: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
