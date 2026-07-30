"""
Script para generar reportes individuales, agregado y PPT de Estadio Israelita Maccabi
Período: 28-12-2025 al 28-12-2026
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
    ENTITY_BASE_URL,
)

# Importar funciones de PPT desde el script de Maipú
import importlib.util
maipu_script = Path(__file__).parent / "generar_reportes_y_ppt_mall_maipu.py"
spec = importlib.util.spec_from_file_location("maipu_ppt", maipu_script)
maipu_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(maipu_module)

# Configuración
COMPANY_ID = "000030"  # Estadio Israelita Maccabi
START_DATE = "28/12/2025"
END_DATE = "28/12/2026"

def get_estadio_israelita_nodes(company_id: str) -> list:
    """Obtiene todos los nodos de Estadio Israelita Maccabi."""
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
    print("  GENERACIÓN DE REPORTES Y PPT - ESTADIO ISRAELITA MACCABI")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print("=" * 70)
    print()
    
    # Importar funciones de PPT
    crear_ppt_analisis = maipu_module.crear_ppt_analisis
    obtener_datos_agregados = maipu_module.obtener_datos_agregados
    guardar_datos_json = maipu_module.guardar_datos_json
    convertir_ppt_a_pdf = maipu_module.convertir_ppt_a_pdf
    
    # Obtener información de la empresa
    company_name = get_company_name(COMPANY_ID)
    print(f"[INFO] Empresa: {company_name} (ID: {COMPANY_ID})")
    print()
    
    # Obtener todos los nodos
    print("[1/5] Obteniendo nodos de Estadio Israelita Maccabi...")
    nodes = get_estadio_israelita_nodes(COMPANY_ID)
    
    if not nodes:
        print("[ERROR] No se encontraron nodos de Estadio Israelita Maccabi.")
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
        from datetime import datetime as dt
        company_name = get_company_name(COMPANY_ID)
        safe_company_name = "".join(c for c in company_name if c.isalnum() or c in (" ", "-", "_")).strip().replace(" ", "_")
        timestamp = dt.now().strftime("%Y%m%d_%H%M")
        output_dir_base = Path("reports") / safe_company_name / "ABREGADO"
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
            output_dir_base = Path("reports") / safe_company_name / "ABREGADO"
            output_dir_base.mkdir(parents=True, exist_ok=True)
            ppt_dir = output_dir_base / f"AGREGADO_{timestamp}"
            ppt_dir.mkdir(exist_ok=True)
        
        ppt_path = ppt_dir / "Agregado PPT.pptx"
        crear_ppt_analisis(datos, ppt_path, mall_name="Estadio Israelita Maccabi")
        
        # Convertir PPT a PDF
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
