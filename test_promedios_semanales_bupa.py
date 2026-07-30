"""
Script de prueba para generar reporte individual de BUPA nodo "Llenado de Estanques"
con la nueva gráfica de promedios por día de la semana.
Período: diciembre 2025 - 10 enero 2026
"""

import sys
from pathlib import Path
import argparse

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import generate_report

# Configuración
COMPANY_ID = "000029"  # BUPA
NODE_ID = "000029-01"  # Llenado de Estanques
START_DATE = "01/12/2025"  # 01 de diciembre 2025
END_DATE = "10/01/2026"  # 10 de enero 2026

def main():
    print("=" * 70)
    print("  PRUEBA: GRÁFICA DE PROMEDIOS POR DÍA DE LA SEMANA")
    print("  BUPA - Llenado de Estanques")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print("=" * 70)
    print()
    
    # Crear argumentos para generate_report
    args = argparse.Namespace(
        company_id=COMPANY_ID,
        node_id=NODE_ID,
        start_date=START_DATE,
        end_date=END_DATE,
        output_dir="reports",  # Directorio base para los reportes
        enviar_correo=False  # No enviar correos
    )
    
    try:
        print(f"[INFO] Generando reporte para nodo {NODE_ID}...")
        report_path = generate_report(args)
        
        if report_path:
            print()
            print("=" * 70)
            print("  PRUEBA COMPLETADA EXITOSAMENTE")
            print("=" * 70)
            print()
            print(f"[OK] Reporte generado: {report_path}")
            print(f"[INFO] El reporte incluye la nueva gráfica de promedios por día de la semana")
            print(f"[INFO] (solo si hay 2 o más semanas completas en el período)")
        else:
            print("[ERROR] No se pudo generar el reporte")
            return
    except Exception as e:
        print(f"[ERROR] Error al generar el reporte: {e}")
        import traceback
        traceback.print_exc()
        return

if __name__ == "__main__":
    main()

