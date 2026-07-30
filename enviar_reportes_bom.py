"""Script para enviar reportes BOM de Parque Arauco por correo."""

import sys
from datetime import datetime
from pathlib import Path
from generar_reporte_word import enviar_reporte_por_correo, get_company_name, get_node_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Configuración
COMPANY_ID = "000025"
NODOS_BOM = [
    ("000025-17", "San Ignacio 300"),
    ("000025-18", "San Ignacio 500"),
]

START_DATE = "03/12/2025"
END_DATE = "08/12/2025"

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
DESTINATARIO = "benjamingumucio@wes.cl"

def main():
    print("=" * 60)
    print("ENVIANDO REPORTES BOM PARQUE ARAUCO")
    print("=" * 60)
    print(f"Destinatario: {DESTINATARIO}")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print("=" * 60)
    print()
    
    company_name = get_company_name(COMPANY_ID)
    company_name_clean = company_name.replace(" ", "_")
    report_dir = Path("reports") / company_name_clean / "REPORTE"
    
    # Formatear fechas para el correo
    start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
    end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
    start_date_str = start_dt.strftime("%d-%m-%y")
    end_date_str = end_dt.strftime("%d-%m-%y")
    
    reportes_enviados = 0
    
    for node_id, node_name in NODOS_BOM:
        print(f"Buscando reporte para {node_name} ({node_id})...")
        
        if not report_dir.exists():
            print(f"  [ERROR] Directorio no existe: {report_dir}")
            continue
        
        # Buscar la carpeta más reciente que contenga el nombre del nodo
        node_name_clean = node_name.replace(" ", "_")
        matching_dirs = [d for d in report_dir.iterdir() if d.is_dir() and node_name_clean in d.name]
        
        if not matching_dirs:
            print(f"  [ERROR] No se encontró carpeta para {node_name}")
            # Intentar buscar por patrón de fecha en el nombre del archivo
            report_files = list(report_dir.rglob(f"Reporte_{COMPANY_ID}_{node_id}_*.docx"))
            if report_files:
                # Ordenar por fecha de modificación y tomar el más reciente
                report_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
                reporte_path = report_files[0]
                print(f"  [OK] Encontrado: {reporte_path}")
            else:
                print(f"  [ERROR] No se encontró reporte para {node_id}")
                continue
        else:
            latest_dir = max(matching_dirs, key=lambda d: d.stat().st_mtime)
            report_files = list(latest_dir.glob("Reporte_*.docx"))
            if not report_files:
                print(f"  [ERROR] No se encontró archivo en {latest_dir}")
                continue
            reporte_path = report_files[0]
            print(f"  [OK] Encontrado: {reporte_path}")
        
        # Enviar reporte
        print(f"  Enviando a {DESTINATARIO}...")
        try:
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
                reportes_enviados += 1
            else:
                print(f"  [ERROR] Fallo al enviar reporte de {node_name}")
        except Exception as e:
            print(f"  [ERROR] {e}")
        
        print()
    
    print("=" * 60)
    print("PROCESO COMPLETADO")
    print("=" * 60)
    print(f"Reportes enviados: {reportes_enviados}/{len(NODOS_BOM)}")

if __name__ == "__main__":
    main()
















