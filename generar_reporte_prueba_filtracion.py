#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Script temporal para generar reporte de prueba con análisis de filtración"""

import argparse
import sys
from pathlib import Path

# Agregar el directorio actual al path para importar generar_reporte_word
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import generate_report

def main():
    # Crear un objeto Namespace con los argumentos necesarios
    # Usando Parque Arauco, nodo Placa Bancaria, periodo diciembre 2025 (31 días)
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
        print("GENERANDO REPORTE DE PRUEBA CON ANÁLISIS DE FILTRACIÓN")
        print("=" * 70)
        print(f"  - Empresa ID: {args.company_id}")
        print(f"  - Nodo ID: {args.node_id}")
        print(f"  - Periodo: {args.start_date} a {args.end_date} (31 días)")
        print("=" * 70)
        print()
        print("Verificando condiciones para análisis de filtración:")
        print("  [OK] Periodo >= 7 días: 31 días")
        print("  [PENDIENTE] Verificar porcentaje días con consumo nocturno >= 95%")
        print()
        print("Si se cumplen las condiciones, el reporte incluirá:")
        print("  1. Tabla 'Proyección de filtración' con:")
        print("     - Promedio hora consumo nocturno")
        print("     - Proyección día fuga")
        print("     - Proyección filtración del periodo")
        print("     - Consumo efectivo del periodo")
        print("  2. Gráfica de anillo: Proyección filtración vs Consumo efectivo")
        print()
        
        output = generate_report(args)
        print()
        print("=" * 70)
        print(f"[OK] Reporte generado exitosamente en:")
        print(f"     {output}")
        print("=" * 70)
        print()
        print("Por favor, verifica en el reporte:")
        print("  1. Si el porcentaje de días con consumo nocturno >= 95%:")
        print("     - Debe aparecer la sección 'Análisis de filtración'")
        print("     - Debe aparecer la tabla 'Proyección de filtración'")
        print("     - Debe aparecer la gráfica de anillo")
        print("  2. Si NO se cumple la condición:")
        print("     - NO debe aparecer la sección 'Análisis de filtración'")
        print("=" * 70)
        
    except Exception as exc:
        print(f"[ERROR] Error al generar reporte: {exc}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()






