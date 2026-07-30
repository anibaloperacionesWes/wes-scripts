"""Script para generar el reporte de Etapa N°5 que falta."""

import subprocess
import sys
from pathlib import Path

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Configuración del nodo
COMPANY_ID = "000027"  # Fundo Zapallar
NODE_ID = "000027-03"  # Etapa N°5
NODE_NAME = "Etapa N°5"

# Fechas: del 1 al 8 de diciembre 2025
START_DATE = "01/12/2025"
END_DATE = "08/12/2025"

# Ruta del script de generación de reportes
PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def main():
    print("=" * 60)
    print("GENERANDO REPORTE DE ETAPA N°5")
    print("=" * 60)
    print(f"Empresa: Fundo Zapallar ({COMPANY_ID})")
    print(f"Nodo: {NODE_NAME} ({NODE_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print("=" * 60)
    print()
    
    # Construir comando
    cmd = [
        PYTHON_EXE,
        SCRIPT_PATH,
        "--company-id", COMPANY_ID,
        "--node-id", NODE_ID,
        "--start-date", START_DATE,
        "--end-date", END_DATE,
    ]
    
    print(f"Ejecutando: {' '.join(cmd)}")
    print()
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=600, encoding='utf-8')
        
        # Mostrar salida
        if result.stdout:
            print("STDOUT:")
            print(result.stdout)
        if result.stderr:
            print("STDERR:")
            print(result.stderr)
        
        if result.returncode == 0:
            print()
            print("=" * 60)
            print("[OK] REPORTE GENERADO EXITOSAMENTE")
            print("=" * 60)
            
            # Verificar si existe el archivo
            reportes = list(Path("reports/Fundo_Zapallar/REPORTE").rglob(f"Reporte_000027_000027-03_*.docx"))
            if reportes:
                print(f"Reporte encontrado en: {reportes[0]}")
            else:
                print("ADVERTENCIA: No se encontró el archivo del reporte")
        else:
            print()
            print("=" * 60)
            print("[ERROR] FALLO AL GENERAR EL REPORTE")
            print("=" * 60)
            print(f"Código de salida: {result.returncode}")
            sys.exit(1)
    except subprocess.TimeoutExpired:
        print()
        print("=" * 60)
        print("[ERROR] TIMEOUT AL GENERAR EL REPORTE")
        print("=" * 60)
        sys.exit(1)
    except Exception as e:
        print()
        print("=" * 60)
        print(f"[ERROR] {e}")
        print("=" * 60)
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()














