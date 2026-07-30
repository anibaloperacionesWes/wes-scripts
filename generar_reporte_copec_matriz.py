"""Script para generar reporte de Matriz Principal COPEC y enviar por correo."""

import subprocess
import sys
from datetime import datetime
from pathlib import Path
from generar_reporte_word import enviar_reporte_por_correo, get_company_name, get_node_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Configuración del nodo
COMPANY_ID = "000009"  # COPEC
NODE_ID = "000009-06"  # Copec Matriz Principal
NODE_NAME = "Copec Matriz Principal"

# Fechas: del 5 al 8 de diciembre 2025
START_DATE = "05/12/2025"
END_DATE = "08/12/2025"

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"  # Contraseña de aplicación
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
DESTINATARIO = "anibal.aoperaciones@wes.cl"

# Ruta del script de generación de reportes
PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def main():
    print("=" * 60)
    print("GENERANDO REPORTE DE MATRIZ PRINCIPAL COPEC")
    print("=" * 60)
    print(f"Empresa: COPEC ({COMPANY_ID})")
    print(f"Nodo: {NODE_NAME} ({NODE_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Destinatario: {DESTINATARIO}")
    print("=" * 60)
    print()
    
    # Generar reporte
    print("Generando reporte...")
    cmd = [
        PYTHON_EXE,
        SCRIPT_PATH,
        "--company-id", COMPANY_ID,
        "--node-id", NODE_ID,
        "--start-date", START_DATE,
        "--end-date", END_DATE,
    ]
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
        if result.returncode == 0:
            print("[OK] Reporte generado exitosamente")
        else:
            error_msg = result.stderr[:200] if result.stderr else result.stdout[:200]
            print(f"[ERROR] {error_msg}")
            sys.exit(1)
    except subprocess.TimeoutExpired:
        print("[ERROR] Timeout al generar reporte")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] {e}")
        sys.exit(1)
    
    # Buscar el archivo generado
    print()
    print("Buscando reporte generado...")
    company_name = get_company_name(COMPANY_ID).replace(" ", "_")
    node_name_clean = NODE_NAME.replace(" ", "_")
    report_dir = Path("reports") / company_name / "REPORTE"
    
    reporte_path = None
    
    if report_dir.exists():
        # Buscar la carpeta más reciente que contenga el nombre del nodo
        matching_dirs = [d for d in report_dir.iterdir() if d.is_dir() and node_name_clean in d.name]
        if matching_dirs:
            latest_dir = max(matching_dirs, key=lambda d: d.stat().st_mtime)
            report_files = list(latest_dir.glob("Reporte_*.docx"))
            if report_files:
                reporte_path = report_files[0]
                print(f"[OK] Reporte encontrado: {reporte_path}")
        else:
            # Intentar buscar por patrón de fecha en el nombre del archivo
            report_files = list(report_dir.rglob(f"Reporte_{COMPANY_ID}_{NODE_ID}_*.docx"))
            if report_files:
                # Ordenar por fecha de modificación y tomar el más reciente
                report_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
                reporte_path = report_files[0]
                print(f"[OK] Reporte encontrado: {reporte_path}")
    
    if not reporte_path or not reporte_path.exists():
        print("[ERROR] No se encontró el reporte generado")
        sys.exit(1)
    
    # Enviar por correo
    print()
    print("=" * 60)
    print(f"ENVIANDO REPORTE A {DESTINATARIO}")
    print("=" * 60)
    
    company_name_display = get_company_name(COMPANY_ID)
    node_name_display = get_node_name(NODE_ID)
    
    # Formatear fechas para el correo
    start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
    end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
    start_date_str = start_dt.strftime("%d-%m-%y")
    end_date_str = end_dt.strftime("%d-%m-%y")
    
    print(f"Enviando reporte a {DESTINATARIO}...")
    
    exito = enviar_reporte_por_correo(
        reporte_path=reporte_path,
        destinatario=DESTINATARIO,
        smtp_servidor=SMTP_SERVIDOR,
        smtp_puerto=SMTP_PUERTO,
        smtp_usuario=SMTP_USUARIO,
        smtp_password=SMTP_PASSWORD,
        company_name=company_name_display,
        node_name=node_name_display,
        start_date=start_date_str,
        end_date=end_date_str,
    )
    
    if exito:
        print()
        print("=" * 60)
        print("[OK] CORREO ENVIADO EXITOSAMENTE")
        print("=" * 60)
        print(f"Destinatario: {DESTINATARIO}")
        print(f"Reporte: {reporte_path.name}")
        print(f"Periodo: {START_DATE} - {END_DATE}")
    else:
        print()
        print("=" * 60)
        print("[ERROR] FALLO EL ENVÍO DEL CORREO")
        print("=" * 60)
        sys.exit(1)

if __name__ == "__main__":
    main()















