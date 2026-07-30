#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Script simple para probar la generación de reporte"""

import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

try:
    print("Importando módulo...")
    from generar_reporte_word import generate_report
    print("Módulo importado correctamente")
    
    print("Creando argumentos...")
    args = argparse.Namespace(
        company_id='000025',
        node_id='000025-08',
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
    
    print("Iniciando generación de reporte...")
    print(f"  - Empresa ID: {args.company_id}")
    print(f"  - Nodo ID: {args.node_id}")
    print(f"  - Periodo: {args.start_date} a {args.end_date}")
    print()
    
    result = generate_report(args)
    
    print()
    print("=" * 70)
    print(f"Reporte generado exitosamente!")
    print(f"Ubicación: {result}")
    print("=" * 70)
    
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)






