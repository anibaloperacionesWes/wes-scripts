"""
Script de prueba para el bot de WhatsApp sin necesidad de Twilio.

Este script simula el procesamiento de mensajes para probar la lógica
del bot sin necesidad de configurar Twilio.
"""

import sys
from pathlib import Path
from bot_whatsapp_reportes import parse_report_request, get_company_name, get_node_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')


def test_parse_request():
    """Prueba el parser de solicitudes."""
    print("=" * 70)
    print("PRUEBA DEL PARSER DE SOLICITUDES")
    print("=" * 70)
    print()
    
    test_cases = [
        "reporte empresa 000025 nodo 000025-12 desde 01/12/2025 hasta 15/12/2025",
        "reporte 000025 000025-12 01/12/2025 15/12/2025",
        "reporte empresa 000025 nodo 000025-12 ultimos 7 dias",
        "reporte agregado empresa 000025 desde 01/12/2025 hasta 15/12/2025",
        "reporte 000029 000029-01 ultimos 14 dias",
    ]
    
    for i, test_message in enumerate(test_cases, 1):
        print(f"Test {i}: {test_message}")
        result = parse_report_request(test_message)
        if result:
            print(f"  ✅ Parseado correctamente:")
            for key, value in result.items():
                print(f"     {key}: {value}")
        else:
            print(f"  ❌ No se pudo parsear")
        print()


if __name__ == "__main__":
    test_parse_request()
    
    print("=" * 70)
    print("PRUEBA DE OBTENCIÓN DE NOMBRES")
    print("=" * 70)
    print()
    
    # Probar obtener nombres
    company_id = "000025"
    node_id = "000025-12"
    
    print(f"Obteniendo nombre de empresa {company_id}...")
    company_name = get_company_name(company_id)
    print(f"  Resultado: {company_name}")
    print()
    
    print(f"Obteniendo nombre de nodo {node_id}...")
    node_name = get_node_name(node_id)
    print(f"  Resultado: {node_name}")
    print()











