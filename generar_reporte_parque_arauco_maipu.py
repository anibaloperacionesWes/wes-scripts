#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Script para generar reporte de Parque Arauco Maipú - Placa Bancaria - Diciembre 2025"""

import argparse
import sys
from pathlib import Path

# Agregar el directorio actual al path para importar generar_reporte_word
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import generate_report

def main():
    # Crear un objeto Namespace con los argumentos necesarios
    # Parque Arauco Maipú, nodo Placa Bancaria, periodo diciembre 2025
    args = argparse.Namespace(
        company_id='000025',  # Parque Arauco
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
        print("GENERANDO REPORTE: PARQUE ARAUCO MAIPÚ - PLACA BANCARIA")
        print("=" * 70)
        print(f"  - Empresa: Parque Arauco (Maipú)")
        print(f"  - Empresa ID: {args.company_id}")
        print(f"  - Nodo: Placa Bancaria")
        print(f"  - Nodo ID: {args.node_id}")
        print(f"  - Periodo: {args.start_date} a {args.end_date} (31 días)")
        print("=" * 70)
        print()
        print("Verificando funcionalidades:")
        print("  [OK] Resumen ejecutivo: número de días del periodo")
        print("  [OK] Resumen ejecutivo: sin promedio de alerta ni proyección diaria")
        print("  [OK] Gráfica de torta nueva (consumo nocturno vs diurno efectivo)")
        print("  [OK] Gráfica de torta antigua eliminada")
        print("  [PENDIENTE] Análisis de filtración (si >= 95% días con consumo nocturno)")
        print()
        print("Generando reporte...")
        print()
        
        output = generate_report(args)
        
        print()
        print("=" * 70)
        print(f"[OK] Reporte generado exitosamente!")
        print(f"     Ubicación: {output}")
        print("=" * 70)
        print()
        print("Verificaciones en el reporte:")
        print("  1. Resumen ejecutivo debe tener:")
        print("     - Consumo total")
        print("     - Promedio diario")
        print("     - Número de días del periodo del reporte: 31")
        print("     - Día pico")
        print("     - Número de alertas de consumo nocturno")
        print("     - NO debe tener 'Promedio de alerta'")
        print("     - NO debe tener 'Proyección diaria de alerta'")
        print()
        print("  2. Sección de alertas debe tener:")
        print("     - Tabla 'Métricas de consumo nocturno'")
        print("     - Gráfica de torta nueva (consumo nocturno vs diurno efectivo)")
        print("     - NO debe tener gráfica de torta antigua (chart_monthly)")
        print()
        print("  3. Si >= 95% días con consumo nocturno:")
        print("     - Debe aparecer sección 'Análisis de filtración'")
        print("     - Tabla 'Proyección de filtración'")
        print("     - Gráfica de anillo (proyección filtración vs consumo efectivo)")
        print()
        print("  4. Sección 'Conclusiones' al final")
        print("=" * 70)
        
    except Exception as exc:
        print()
        print("=" * 70)
        print(f"[ERROR] Error al generar reporte: {exc}")
        print("=" * 70)
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()






