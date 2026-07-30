"""
Script para generar reportes individuales, agregado y PPT de Providencia
Nodos: Liceo Lastarria, Liceo 7, Liceo Juan Pablo Duarte, Carmela Carvajal, Arturo Alessandri Palma
Período: 26-12-2025 al 27-01-2026
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
COMPANY_ID = "000006"  # Providencia
# Periodo solicitado: 27-12-2025 al 27-01-2026
START_DATE = "27/12/2025"
END_DATE = "27/01/2026"

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
    "liceo 7",
    "liceo juan pablo duarte",
    "carmela carvajal",
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
    print("  GENERACIÓN DE REPORTES Y PPT - PROVIDENCIA")
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
    
    # Obtener nodos solicitados (por nombre) y luego filtrar por ID
    print("[1/5] Obteniendo nodos de Providencia...")
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
        crear_ppt_analisis(datos, ppt_path, mall_name="Providencia")
        
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
