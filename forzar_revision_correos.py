"""
Script para forzar una revisión inmediata de correos sin esperar el intervalo.
Útil cuando se necesita procesar un correo recién recibido.
"""
import sys
from pathlib import Path

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Importar la función de monitoreo
from monitorear_correos_y_generar_reportes import monitorear_y_procesar_correos

if __name__ == "__main__":
    print("=" * 70)
    print("FORZANDO REVISIÓN INMEDIATA DE CORREOS")
    print("=" * 70)
    print()
    print("[INFO] Ejecutando revisión de correos ahora...")
    print()
    
    try:
        monitorear_y_procesar_correos()
        print()
        print("[OK] Revisión completada")
    except Exception as e:
        print()
        print(f"[ERROR] Error durante la revisión: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)







