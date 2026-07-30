"""Script directo para enviar reportes BOM por correo."""

import sys
from datetime import datetime
from pathlib import Path
from generar_reporte_word import enviar_reporte_por_correo, get_company_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Rutas exactas de los reportes
REPORTES = [
    ("reports/Parque_Arauco/REPORTE/San_Ignacio_300_20251209_1224/Reporte_000025_000025-17_20251203_20251208.docx", "San Ignacio 300"),
    ("reports/Parque_Arauco/REPORTE/San_Ignacio_500_20251209_1224/Reporte_000025_000025-18_20251203_20251208.docx", "San Ignacio 500"),
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
    # Forzar salida inmediata
    sys.stdout.flush()
    sys.stderr.flush()
    
    print("=" * 60, flush=True)
    print("ENVIANDO REPORTES BOM PARQUE ARAUCO", flush=True)
    print("=" * 60, flush=True)
    print(f"Destinatario: {DESTINATARIO}", flush=True)
    print(f"Periodo: {START_DATE} - {END_DATE}", flush=True)
    print("=" * 60, flush=True)
    print(flush=True)
    
    company_name = get_company_name("000025")
    
    # Formatear fechas para el correo
    start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
    end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
    start_date_str = start_dt.strftime("%d-%m-%y")
    end_date_str = end_dt.strftime("%d-%m-%y")
    
    reportes_enviados = 0
    
    for reporte_path_str, node_name in REPORTES:
        reporte_path = Path(reporte_path_str)
        print(f"Enviando reporte: {node_name}", flush=True)
        print(f"  Archivo: {reporte_path.name}", flush=True)
        
        if not reporte_path.exists():
            print(f"  [ERROR] Archivo no encontrado: {reporte_path}")
            continue
        
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
                print(f"  [OK] Reporte enviado exitosamente")
                reportes_enviados += 1
            else:
                print(f"  [ERROR] Fallo al enviar reporte")
        except Exception as e:
            print(f"  [ERROR] {e}")
            import traceback
            traceback.print_exc()
        
        print()
    
    print("=" * 60)
    print("PROCESO COMPLETADO")
    print("=" * 60)
    print(f"Reportes enviados: {reportes_enviados}/{len(REPORTES)}")
    print(f"Destinatario: {DESTINATARIO}")

if __name__ == "__main__":
    main()

