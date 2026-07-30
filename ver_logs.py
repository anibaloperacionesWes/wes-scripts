"""
Script para ver los logs del monitoreo de correos.
Muestra las últimas líneas del archivo de log.
"""
import sys
from pathlib import Path

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

LOG_DIR = Path(__file__).parent / "logs"
LOG_FILE = LOG_DIR / "monitoreo_correos.log"

def ver_logs(ultimas_lineas=50):
    """Muestra las últimas líneas del archivo de log"""
    if not LOG_FILE.exists():
        print(f"[ERROR] No se encontró el archivo de log: {LOG_FILE}")
        print("[INFO] El archivo se creará cuando el script de monitoreo se ejecute por primera vez.")
        return
    
    try:
        with open(LOG_FILE, 'r', encoding='utf-8') as f:
            lineas = f.readlines()
        
        total_lineas = len(lineas)
        inicio = max(0, total_lineas - ultimas_lineas)
        
        print("=" * 70)
        print(f"LOGS DEL MONITOREO DE CORREOS")
        print("=" * 70)
        print(f"Archivo: {LOG_FILE}")
        print(f"Total de líneas: {total_lineas}")
        print(f"Mostrando últimas {min(ultimas_lineas, total_lineas)} líneas:")
        print("=" * 70)
        print()
        
        for linea in lineas[inicio:]:
            print(linea.rstrip())
        
        print()
        print("=" * 70)
        print(f"[INFO] Para ver más líneas, ejecuta: python ver_logs.py <número>")
        print(f"[INFO] Ejemplo: python ver_logs.py 100")
        
    except Exception as e:
        print(f"[ERROR] No se pudo leer el archivo de log: {e}")

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Ver logs del monitoreo de correos")
    parser.add_argument(
        "lineas",
        type=int,
        nargs="?",
        default=50,
        help="Número de líneas a mostrar (default: 50)"
    )
    
    args = parser.parse_args()
    ver_logs(args.lineas)







