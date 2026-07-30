"""Script para generar reportes de los 2 puntos BOM de Parque Arauco y enviar por correo."""

import argparse
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from generar_reporte_word import enviar_reporte_por_correo, get_company_name, get_node_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Configuración de los nodos BOM de Parque Arauco
COMPANY_ID = "000025"  # Parque Arauco
NODOS_BOM = [
    "000025-17",  # BOM San Ignacio 300
    "000025-18",  # BOM San Ignacio 500
]

# Fechas: del 3 al 8 de diciembre 2025
START_DATE = "03/12/2025"
END_DATE = "08/12/2025"

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"  # Contraseña de aplicación
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
DESTINATARIO = "benjamingumucio@wes.cl"

# Ruta del script de generación de reportes
PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def main():
    print("=" * 60)
    print("GENERANDO REPORTES DE BOM PARQUE ARAUCO")
    print("=" * 60)
    print(f"Empresa: Parque Arauco ({COMPANY_ID})")
    print(f"Nodos: {', '.join(NODOS_BOM)}")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Destinatario: {DESTINATARIO}")
    print("=" * 60)
    print()
    
    nodos_exitosos = []
    nodos_fallidos = []
    reportes_generados = []
    
    # Generar reportes individuales
    for i, node_id in enumerate(NODOS_BOM, 1):
        node_name = get_node_name(node_id)
        print(f"[{i}/{len(NODOS_BOM)}] Generando reporte para {node_name} ({node_id})...")
        
        cmd = [
            PYTHON_EXE,
            SCRIPT_PATH,
            "--company-id", COMPANY_ID,
            "--node-id", node_id,
            "--start-date", START_DATE,
            "--end-date", END_DATE,
        ]
        
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
            if result.returncode == 0:
                print(f"  [OK] Reporte generado exitosamente")
                nodos_exitosos.append(node_id)
                
                # Buscar el archivo generado
                company_name = get_company_name(COMPANY_ID).replace(" ", "_")
                node_name_clean = node_name.replace(" ", "_")
                report_dir = Path("reports") / company_name / "REPORTE"
                if report_dir.exists():
                    # Buscar la carpeta más reciente que contenga el nombre del nodo
                    matching_dirs = [d for d in report_dir.iterdir() if d.is_dir() and node_name_clean in d.name]
                    if matching_dirs:
                        latest_dir = max(matching_dirs, key=lambda d: d.stat().st_mtime)
                        report_files = list(latest_dir.glob("Reporte_*.docx"))
                        if report_files:
                            reportes_generados.append((node_id, node_name, report_files[0]))
            else:
                error_msg = result.stderr[:200] if result.stderr else result.stdout[:200]
                print(f"  [ERROR] {error_msg}")
                nodos_fallidos.append(node_id)
        except subprocess.TimeoutExpired:
            print(f"  [ERROR] Timeout al generar reporte")
            nodos_fallidos.append(node_id)
        except Exception as e:
            print(f"  [ERROR] {e}")
            nodos_fallidos.append(node_id)
        
        print()
    
    print("=" * 60)
    print("RESUMEN DE GENERACIÓN")
    print("=" * 60)
    print(f"Exitosos: {len(nodos_exitosos)}")
    print(f"Fallidos: {len(nodos_fallidos)}")
    
    if nodos_fallidos:
        print(f"\nNodos con errores: {', '.join(nodos_fallidos)}")
    
    # Enviar reportes por correo
    if reportes_generados:
        print()
        print("=" * 60)
        print(f"ENVIANDO REPORTES A {DESTINATARIO}")
        print("=" * 60)
        
        company_name = get_company_name(COMPANY_ID)
        
        # Formatear fechas para el correo
        start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
        end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
        start_date_str = start_dt.strftime("%d-%m-%y")
        end_date_str = end_dt.strftime("%d-%m-%y")
        
        for i, (node_id, node_name, reporte_path) in enumerate(reportes_generados, 1):
            if reporte_path.exists():
                print(f"[{i}/{len(reportes_generados)}] Enviando reporte de {node_name}...")
                
                exito = enviar_reporte_por_correo(
                    reporte_path=reporte_path,
                    destinatario=DESTINATARIO,
                    smtp_servidor=SMTP_SERVIDOR,
                    smtp_puerto=SMTP_PUERTO,
                    smtp_usuario=SMTP_USUARIO,
                    smtp_password=SMTP_PASSWORD,
                    company_name=company_name,
                    node_name=node_name,
                    start_date=start_date_str,
                    end_date=end_date_str,
                )
                
                if exito:
                    print(f"  [OK] Reporte de {node_name} enviado exitosamente")
                else:
                    print(f"  [ERROR] Fallo al enviar reporte de {node_name}")
            else:
                print(f"  [ERROR] Archivo no encontrado: {reporte_path}")
        
        print()
        print("=" * 60)
        print("PROCESO COMPLETADO")
        print("=" * 60)
        print(f"Reportes generados: {len(reportes_generados)}")
        print(f"Reportes enviados: {len(reportes_generados)}")
    else:
        print()
        print("=" * 60)
        print("[ERROR] NO SE PUDIERON GENERAR REPORTES")
        print("=" * 60)

if __name__ == "__main__":
    main()

