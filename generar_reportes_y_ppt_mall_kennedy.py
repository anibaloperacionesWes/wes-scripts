"""
Script para generar reporte agregado de Mall Kennedy y PPT en PDF de análisis similar
Período: 25 de diciembre 2025 al 25 de enero 2026
"""

import sys
from pathlib import Path
from datetime import datetime, timedelta
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
    get_hourly_measures_for_day,
    calculate_nocturnal_metrics,
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
START_DATE = "26/12/2025"
END_DATE = "26/01/2026"

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

def get_kennedy_nodes(company_id: str) -> list:
    """Obtiene todos los nodos del Mall Kennedy."""
    url = f"{ENTITY_BASE_URL}/companies/{company_id}"
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            all_nodes = data.get("nodes", [])
            kennedy_nodes = []
            for node in all_nodes:
                node_id = node.get("nodeId", "")
                node_name = node.get("name", "").strip()
                # Verificar si está en la lista conocida de Kennedy
                if node_id in NODOS_KENNEDY:
                    kennedy_nodes.append({
                        "nodeId": node_id,
                        "name": node_name
                    })
                else:
                    # Si no está en la lista, verificar por el nombre del mall
                    mall_name = get_mall_name_for_parque_arauco(node_id, node_name)
                    if mall_name == "Kennedy":
                        kennedy_nodes.append({
                            "nodeId": node_id,
                            "name": node_name
                        })
            return kennedy_nodes
        return []
    except Exception as e:
        print(f"[ERROR] Error al obtener nodos: {e}")
        return []

# Importar funciones necesarias desde el script de Maipú
# Copiamos las funciones necesarias aquí para evitar dependencias circulares
def obtener_datos_agregados(node_ids: list, start_date: str, end_date: str) -> dict:
    """Obtiene todos los datos agregados de los nodos para el período especificado."""
    from generar_reporte_word import MeasurePoint
    
    start_dt = parse_date(start_date)
    end_dt = parse_date(end_date)
    
    all_measures = []
    nodes_summary = []
    total_consumption = 0.0
    
    for node_id in node_ids:
        node_name = get_node_name(node_id)
        print(f"  Obteniendo datos de {node_id} ({node_name})...", flush=True)
        
        try:
            # Obtener medidas del nodo
            url = f"{BASE_URL}/nodes/{node_id}/measures"
            params = {
                "startDate": start_dt.strftime("%Y-%m-%d"),
                "endDate": end_dt.strftime("%Y-%m-%d")
            }
            
            response = requests.get(url, params=params, timeout=30)
            if response.status_code == 200:
                data = response.json()
                payload = normalize_measures_payload(data)
                measures = flatten_measures(payload)
                
                # Convertir a MeasurePoint
                measure_points = []
                for m in measures:
                    dt_obj = datetime.fromisoformat(m["date"].replace("Z", "+00:00"))
                    measure_points.append(MeasurePoint(
                        date=dt_obj,
                        total_m3=float(m["total_m3"]),
                        details=m.get("details", {})
                    ))
                
                # Calcular summary para este nodo
                summary = summarize_consumption(measure_points)
                node_total = sum(m.total_m3 for m in measure_points)
                total_consumption += node_total
                
                nodes_summary.append({
                    "node_id": node_id,
                    "node_name": node_name,
                    "summary": summary,
                    "measures": measure_points
                })
                
                all_measures.extend(measure_points)
            else:
                print(f"    [ADVERTENCIA] No se pudieron obtener datos de {node_id}: {response.status_code}")
        except Exception as e:
            print(f"    [ERROR] Error al obtener datos de {node_id}: {e}")
            continue
    
    # Calcular summary agregado
    aggregate_summary = summarize_consumption(all_measures)
    
    return {
        "total_consumption": total_consumption,
        "aggregate_summary": aggregate_summary,
        "nodes_summary": nodes_summary,
        "all_measures": all_measures,
        "start_date": start_dt,
        "end_date": end_dt
    }

def guardar_datos_json(datos: dict, output_path: Path):
    """Guarda los datos agregados en un archivo JSON para uso futuro."""
    def serializar_measurepoint(m):
        """Convierte un MeasurePoint a diccionario serializable."""
        return {
            "date": m.date.isoformat(),
            "total_m3": m.total_m3,
            "details": m.details
        }
    
    def serializar_summary(summ):
        """Convierte un summary a formato serializable."""
        if summ is None:
            return None
        result = {}
        for key, value in summ.items():
            if hasattr(value, 'date') and hasattr(value, 'total_m3'):
                # Es un MeasurePoint
                result[key] = serializar_measurepoint(value)
            elif isinstance(value, datetime):
                result[key] = value.isoformat()
            else:
                result[key] = value
        return result
    
    datos_serializados = {
        "total_consumption": datos["total_consumption"],
        "aggregate_summary": serializar_summary(datos["aggregate_summary"]),
        "nodes_summary": [
            {
                "node_id": node["node_id"],
                "node_name": node["node_name"],
                "summary": serializar_summary(node["summary"]),
                "measures": [serializar_measurepoint(m) for m in node["measures"]]
            }
            for node in datos["nodes_summary"]
        ],
        "all_measures": [serializar_measurepoint(m) for m in datos["all_measures"]],
        "start_date": datos["start_date"].isoformat(),
        "end_date": datos["end_date"].isoformat()
    }
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(datos_serializados, f, indent=2, ensure_ascii=False)
    
    print(f"[OK] Datos guardados en JSON: {output_path}")

def cargar_datos_json(json_path: Path) -> dict:
    """Carga los datos agregados desde un archivo JSON."""
    from generar_reporte_word import MeasurePoint
    
    with open(json_path, 'r', encoding='utf-8') as f:
        datos_json = json.load(f)
    
    def deserializar_measurepoint(d):
        """Convierte un diccionario a MeasurePoint."""
        date_obj = datetime.fromisoformat(d["date"])
        return MeasurePoint(
            date=date_obj,
            total_m3=d["total_m3"],
            details=d.get("details", {})
        )
    
    def deserializar_summary(summ):
        """Convierte un diccionario de summary a formato con MeasurePoint."""
        if summ is None:
            return None
        result = {}
        for key, value in summ.items():
            if isinstance(value, dict) and "date" in value and "total_m3" in value:
                result[key] = deserializar_measurepoint(value)
            elif isinstance(value, str) and 'T' in value:
                result[key] = datetime.fromisoformat(value)
            else:
                result[key] = value
        return result
    
    datos = {
        "total_consumption": datos_json["total_consumption"],
        "aggregate_summary": deserializar_summary(datos_json["aggregate_summary"]),
        "nodes_summary": [
            {
                "node_id": node["node_id"],
                "node_name": node["node_name"],
                "summary": deserializar_summary(node["summary"]),
                "measures": [deserializar_measurepoint(m) for m in node["measures"]]
            }
            for node in datos_json["nodes_summary"]
        ],
        "all_measures": [deserializar_measurepoint(m) for m in datos_json["all_measures"]],
        "start_date": datetime.fromisoformat(datos_json["start_date"]),
        "end_date": datetime.fromisoformat(datos_json["end_date"])
    }
    
    print(f"[OK] Datos cargados desde JSON: {json_path}")
    return datos

def buscar_ultimo_reporte_agregado_kennedy() -> Optional[Path]:
    """Busca el último reporte agregado de Kennedy para reutilizar datos."""
    company_name = get_company_name(COMPANY_ID)
    safe_company_name = "".join(c for c in company_name if c.isalnum() or c in (" ", "-", "_")).strip().replace(" ", "_")
    
    # Buscar SOLO en reports/Parque_Arauco/Kennedy/ABREGADO
    base_dir = Path("reports") / safe_company_name / "Kennedy"
    if not base_dir.exists():
        return None
    
    # Buscar directorios que contengan "AGREGADO" o "ABREGADO" SOLO dentro de Kennedy
    agregado_dirs = []
    for subdir in base_dir.iterdir():
        if subdir.is_dir() and ("AGREGADO" in subdir.name.upper() or "ABREGADO" in subdir.name.upper()):
            agregado_dirs.append(subdir)
        elif subdir.is_dir():
            # Buscar dentro de subdirectorios
            for subsubdir in subdir.iterdir():
                if subsubdir.is_dir() and ("AGREGADO" in subsubdir.name.upper() or "ABREGADO" in subsubdir.name.upper()):
                    agregado_dirs.append(subsubdir)
    
    if not agregado_dirs:
        return None
    
    # Ordenar por fecha de modificación (más reciente primero)
    agregado_dirs.sort(key=lambda p: p.stat().st_mtime if p.is_dir() else 0, reverse=True)
    
    # Buscar archivo JSON de datos
    for dir_path in agregado_dirs:
        json_path = dir_path / "datos_agregados.json"
        if json_path.exists():
            return dir_path
    
    return None

def convertir_ppt_a_pdf(ppt_path: Path) -> Optional[Path]:
    """Convierte un archivo PPT a PDF usando comtypes (Windows) o fallback."""
    pdf_path = ppt_path.with_suffix('.pdf')
    
    try:
        import comtypes.client
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1
        
        try:
            presentation = powerpoint.Presentations.Open(str(ppt_path.absolute()))
            presentation.SaveAs(str(pdf_path.absolute()), 32)  # 32 = ppSaveAsPDF
            presentation.Close()
            powerpoint.Quit()
            print(f"[OK] PPT convertido a PDF: {pdf_path}")
            return pdf_path
        except Exception as e:
            print(f"[ADVERTENCIA] Error al convertir PPT a PDF: {e}")
            powerpoint.Quit()
            return None
    except ImportError:
        print("[INFO] comtypes no disponible. PPT no se convertirá a PDF automáticamente.")
        return None
    except Exception as e:
        print(f"[ADVERTENCIA] Error al convertir PPT a PDF: {e}")
        return None

# Importar funciones de creación de PPT desde el script de Maipú
# Necesitamos importar crear_ppt_analisis y funciones relacionadas
# Por simplicidad, vamos a importar todo el módulo de Maipú y usar sus funciones
def importar_funciones_ppt():
    """Importa las funciones necesarias para crear PPT desde el script de Maipú."""
    import importlib.util
    maipu_script = Path(__file__).parent / "generar_reportes_y_ppt_mall_maipu.py"
    spec = importlib.util.spec_from_file_location("maipu_ppt", maipu_script)
    maipu_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(maipu_module)
    return maipu_module

def main():
    print("=" * 70)
    print("  GENERACIÓN DE REPORTES Y PPT - MALL KENNEDY")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print("=" * 70)
    print()
    
    # Importar funciones de PPT
    maipu_module = importar_funciones_ppt()
    crear_ppt_analisis = maipu_module.crear_ppt_analisis
    
    # Obtener nodos de Kennedy
    print("[1/5] Obteniendo nodos del Mall Kennedy...")
    nodes = get_kennedy_nodes(COMPANY_ID)
    
    if not nodes:
        print("[ERROR] No se encontraron nodos del Mall Kennedy.")
        return
    
    print(f"[OK] Se encontraron {len(nodes)} nodo(s)")
    node_ids = [node["nodeId"] for node in nodes]
    for node in nodes:
        print(f"  - {node['nodeId']}: {node['name']}")
    print()
    
    # Buscar último reporte agregado existente
    print("[INFO] Buscando último reporte agregado existente...")
    ultimo_reporte_dir = buscar_ultimo_reporte_agregado_kennedy()
    
    if ultimo_reporte_dir:
        print(f"[OK] Reporte agregado encontrado: {ultimo_reporte_dir}")
        print("[INFO] Generando solo PPT sin generar reportes nuevos")
        print()
        
        # Cargar datos desde JSON
        json_path = ultimo_reporte_dir / "datos_agregados.json"
        if json_path.exists():
            datos = cargar_datos_json(json_path)
            
            # Generar PPT
            print("[2/2] Generando presentación PPT...")
            try:
                ppt_path = ultimo_reporte_dir / "Agregado PPT.pptx"
                crear_ppt_analisis(datos, ppt_path, mall_name="Kennedy")
                
                # Convertir PPT a PDF
                pdf_path = convertir_ppt_a_pdf(ppt_path)
                
                print()
                print("=" * 70)
                print("  PROCESO COMPLETADO")
                print("=" * 70)
                print(f"[OK] Presentación PPT: {ppt_path}")
                if pdf_path:
                    print(f"[OK] Presentación PDF: {pdf_path}")
            except Exception as e:
                print(f"[ERROR] Error al generar PPT: {e}")
                import traceback
                traceback.print_exc()
        else:
            print("[ERROR] No se encontró archivo JSON de datos en el directorio.")
    else:
        print("[INFO] No se encontró reporte agregado previo. Generando reportes nuevos...")
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
                mall_name="Kennedy"
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
            output_dir_base = Path("reports") / safe_company_name / "Kennedy" / "ABREGADO"
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
                output_dir_base = Path("reports") / safe_company_name / "Kennedy" / "ABREGADO"
                output_dir_base.mkdir(parents=True, exist_ok=True)
                ppt_dir = output_dir_base / f"AGREGADO_{timestamp}"
                ppt_dir.mkdir(exist_ok=True)
            
            ppt_path = ppt_dir / "Agregado PPT.pptx"
            crear_ppt_analisis(datos, ppt_path, mall_name="Kennedy")
            
            # Convertir PPT a PDF
            pdf_path = convertir_ppt_a_pdf(ppt_path)
            
            print()
            print("=" * 70)
            print("  PROCESO COMPLETADO")
            print("=" * 70)
            if 'generated_reports' in locals():
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
