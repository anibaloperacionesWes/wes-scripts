"""Script de prueba para verificar el comportamiento con 3 o más alertas."""

import sys
from generar_reporte_word import generate_report
from generar_reporte_word import get_node_name

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

COMPANY_ID = "000029"
START_DATE = "03/10/2025"
END_DATE = "16/12/2025"

# Probar con Torre B1 que probablemente tenga alertas
NODE_ID = "000029-03"

print(f"Probando generación de reporte para {get_node_name(NODE_ID)} ({NODE_ID})")
print(f"Periodo: {START_DATE} - {END_DATE}")
print()

try:
    result = generate_report(
        company_id=COMPANY_ID,
        node_id=NODE_ID,
        start_date=START_DATE,
        end_date=END_DATE
    )
    print(f"[OK] Reporte generado: {result}")
except Exception as e:
    print(f"[ERROR] {e}")
    import traceback
    traceback.print_exc()










