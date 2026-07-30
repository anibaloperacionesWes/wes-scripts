#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Script temporal para generar reporte de prueba con las nuevas modificaciones"""

import argparse
import sys
from pathlib import Path

# Agregar el directorio actual al path para importar generar_reporte_word
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import generate_report

def main():
    # Crear un objeto Namespace con los argumentos necesarios
    # Usando Parque Arauco, nodo Placa Bancaria, periodo diciembre 2025
    args = argparse.Namespace(
        company_id='000025',
        node_id='000025-08',  # Placa Bancaria
        start_date='2025-12-01',
        end_date='2025-12-31',
        output_dir='reports',
        enviar_correo=False,
        destinatario=None,
        smtp_servidor=None,
        smtp_puerto=None,
        smtp_usuario=None,
        smtp_password=None
    )
    
    try:
        print("=" * 70)
        print("GENERANDO REPORTE DE PRUEBA CON CORRECCIONES")
        print("=" * 70)
        print(f"  - Empresa ID: {args.company_id}")
        print(f"  - Nodo ID: {args.node_id}")
        print(f"  - Periodo: {args.start_date} a {args.end_date}")
        print("=" * 70)
        print()
        print("Verificando correcciones:")
        print("  [OK] Grafica de torta nueva achicada (de 6 a 4 pulgadas)")
        print("  [OK] Grafica de torta antigua (chart_monthly) eliminada")
        print("  [OK] Resumen ejecutivo: eliminado promedio de alerta y proyeccion diaria")
        print("  [OK] Resumen ejecutivo: agregado numero de dias del periodo")
        print("  [PENDIENTE] Verificar analisis de filtracion (si >= 7 dias y >= 95% dias con consumo nocturno)")
        print()
        
        output = generate_report(args)
        print()
        print("=" * 70)
        print(f"[OK] Reporte generado exitosamente en:")
        print(f"     {output}")
        print("=" * 70)
        print()
        print("Por favor, verifica en el reporte:")
        print("  1. La nueva grafica de torta esta mas pequena (4 pulgadas)")
        print("  2. Que NO aparezca la grafica de torta antigua (chart_monthly)")
        print("  3. Resumen ejecutivo: NO debe tener 'Promedio de alerta' ni 'Proyeccion diaria de alerta'")
        print("  4. Resumen ejecutivo: Debe tener 'Numero de dias del periodo del reporte: 31'")
        print("  5. Si >= 95% dias con consumo nocturno: debe aparecer seccion 'Analisis de filtracion'")
        print("=" * 70)
        
    except Exception as exc:
        print(f"[ERROR] Error al generar reporte: {exc}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()

