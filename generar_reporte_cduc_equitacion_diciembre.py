"""
Script para generar el reporte del nodo Equitación de CDUC
para el período 1 de diciembre al 31 de diciembre 2025.
No envía correos electrónicos.
"""

import sys
from pathlib import Path
import argparse

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import (
    get_company_name,
    get_node_name,
    generate_report,
)

# Configuración
COMPANY_ID = "000021"  # CDUC
NODE_ID = "000021-04"  # Equitación
START_DATE = "2025-12-01"
END_DATE = "2025-12-31"

def main():
    print("=" * 60)
    print("  REPORTE CDUC - NODO EQUITACIÓN")
    print("  Diciembre 2025")
    print("=" * 60)
    print()
    
    # Obtener información de la empresa y nodo
    company_name = get_company_name(COMPANY_ID)
    node_name = get_node_name(NODE_ID)
    
    print(f"[INFO] Empresa: {company_name} (ID: {COMPANY_ID})")
    print(f"[INFO] Nodo: {node_name} (ID: {NODE_ID})")
    print(f"[INFO] Periodo: {START_DATE} a {END_DATE}")
    print()
    
    # Generar reporte
    print("[1/1] Generando reporte...")
    
    try:
        # Crear argumentos para generate_report usando argparse.Namespace
        args = argparse.Namespace(
            company_id=COMPANY_ID,
            node_id=NODE_ID,
            start_date=START_DATE,
            end_date=END_DATE,
            output_dir="reports",  # Directorio base para los reportes
            enviar_correo=False  # No enviar correos
        )
        report_path = generate_report(args)
        
        if report_path:
            print(f"[OK] Reporte generado exitosamente: {report_path}")
        else:
            print(f"[ERROR] No se pudo generar el reporte")
            return
    except Exception as e:
        print(f"[ERROR] No se pudo generar el reporte: {e}")
        import traceback
        traceback.print_exc()
        return
    
    print()
    print("=" * 60)
    print("  PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 60)
    print()
    print(f"[INFO] Reporte generado: 1")
    print(f"[INFO] Periodo: {START_DATE} a {END_DATE}")
    print(f"[INFO] Nodo: {node_name} ({NODE_ID})")
    print(f"[INFO] Ruta del reporte: {report_path}")
    print(f"[INFO] No se enviaron correos electrónicos (según solicitud)")

if __name__ == "__main__":
    main()



