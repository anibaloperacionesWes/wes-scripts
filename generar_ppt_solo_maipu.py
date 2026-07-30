"""
Script para generar solo la PPT desde el último reporte agregado de Maipú
USANDO LA API para obtener todos los datos
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from generar_reportes_y_ppt_mall_maipu import (
    buscar_ultimo_reporte_agregado_maipu,
    obtener_datos_agregados,
    crear_ppt_analisis,
    convertir_ppt_a_pdf,
    get_maipu_nodes,
    COMPANY_ID,
    START_DATE,
    END_DATE,
)

def main():
    print("=" * 70)
    print("  GENERACIÓN DE PPT - MALL MAIPÚ (USANDO API)")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print("=" * 70)
    print()
    
    # Obtener nodos de Maipú
    print("[1/3] Obteniendo nodos del Mall Maipú...")
    nodes = get_maipu_nodes(COMPANY_ID)
    
    if not nodes:
        print("[ERROR] No se encontraron nodos del Mall Maipú.")
        return
    
    print(f"[OK] Se encontraron {len(nodes)} nodo(s)")
    node_ids = [node["nodeId"] for node in nodes]
    for node in nodes:
        print(f"  - {node['nodeId']}: {node['name']}")
    print()
    
    # Buscar último reporte agregado para guardar la PPT ahí
    print("[2/3] Buscando último reporte agregado para guardar PPT...")
    reports_base = Path("reports")
    posibles_rutas = [
        reports_base / "Parque_Arauco" / "Maipú" / "ABREGADO",
        reports_base / "Parque_Arauco" / "Maipu" / "ABREGADO",
    ]
    
    ultimo_reporte_dir = None
    for ruta_base in posibles_rutas:
        if ruta_base.exists():
            agregado_dirs = [d for d in ruta_base.iterdir() if d.is_dir() and d.name.startswith("AGREGADO_")]
            if agregado_dirs:
                agregado_dirs.sort(key=lambda x: x.stat().st_mtime, reverse=True)
                ultimo_reporte_dir = agregado_dirs[0]
                break
    
    if not ultimo_reporte_dir:
        print("[ADVERTENCIA] No se encontró directorio de reporte agregado.")
        print("[INFO] Se creará un nuevo directorio para la PPT")
        from datetime import datetime as dt
        timestamp = dt.now().strftime("%Y%m%d_%H%M")
        ultimo_reporte_dir = posibles_rutas[0] / f"AGREGADO_{timestamp}"
        ultimo_reporte_dir.mkdir(parents=True, exist_ok=True)
    else:
        print(f"[OK] Directorio encontrado: {ultimo_reporte_dir}")
    
    print()
    
    # Obtener datos desde la API
    print("[3/3] Obteniendo datos desde la API...")
    try:
        datos = obtener_datos_agregados(node_ids, START_DATE, END_DATE)
        # Marcar que SÍ se debe usar la API
        datos['_usar_api'] = True
        print(f"[OK] Datos obtenidos desde API: {len(datos['all_measures'])} medidas")
    except Exception as e:
        print(f"[ERROR] Error al obtener datos desde API: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # Generar PPT
    print()
    print("[3/3] Generando presentación PPT (con todos los datos desde API)...")
    try:
        ppt_path = ultimo_reporte_dir / "Agregado PPT.pptx"
        crear_ppt_analisis(datos, ppt_path, mall_name="Maipú")
        
        # Convertir PPT a PDF
        pdf_path = convertir_ppt_a_pdf(ppt_path)
        
        print()
        print("=" * 70)
        print("  PPT GENERADA EXITOSAMENTE (CON API)")
        print("=" * 70)
        print(f"[OK] PPT generada en: {ppt_path}")
        if pdf_path:
            print(f"[OK] PDF generado en: {pdf_path}")
        print(f"[OK] Directorio del reporte: {ultimo_reporte_dir}")
    except Exception as e:
        print(f"[ERROR] Error al generar PPT: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
