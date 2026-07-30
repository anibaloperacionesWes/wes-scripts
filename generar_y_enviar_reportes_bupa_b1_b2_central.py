"""
Script para generar reportes individuales de BUPA para los nodos B1, B2 y Central Térmica
del período 15 de octubre 2025 hasta 29 de diciembre 2025, y enviarlos por correo a Diego y José.
"""

import sys
import argparse
import smtplib
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import (
    generate_report,
    get_company_name,
    get_node_name,
)

# Configuración
COMPANY_ID = "000029"  # BUPA
START_DATE = "15/10/2025"
END_DATE = "29/12/2025"

# Nodos a generar (B1, B2, Central Térmica)
NODOS_SOLICITADOS = {
    "000029-03": "Torre B1",
    "000029-04": "Torre B2",
    "000029-06": "Central Térmica",
}

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

# Destinatarios
DESTINATARIOS = [
    "diegocarrasco@wes.cl",  # Diego
    "joseotarola@wes.cl",  # José
]

def generar_reportes_individuales():
    """
    Genera los reportes individuales para los nodos solicitados.
    Retorna lista de paths de los reportes generados.
    """
    reportes_generados = []
    
    print("=" * 70)
    print("GENERANDO REPORTES INDIVIDUALES DE BUPA")
    print("=" * 70)
    print(f"Empresa: {get_company_name(COMPANY_ID)} ({COMPANY_ID})")
    print(f"Período: {START_DATE} - {END_DATE}")
    print(f"Nodos a generar: {len(NODOS_SOLICITADOS)}")
    print()
    
    for i, (node_id, node_name) in enumerate(NODOS_SOLICITADOS.items(), 1):
        print(f"[{i}/{len(NODOS_SOLICITADOS)}] Generando reporte para {node_name} ({node_id})...")
        
        try:
            # Crear args para generate_report
            args = argparse.Namespace(
                company_id=COMPANY_ID,
                node_id=node_id,
                start_date=START_DATE,
                end_date=END_DATE,
                output_dir="reports",
                enviar_correo=False,  # No enviar correo individual, lo haremos al final
            )
            
            # Generar reporte
            reporte_path = generate_report(args)
            
            if reporte_path and Path(reporte_path).exists():
                reportes_generados.append(Path(reporte_path))
                print(f"  [OK] Reporte generado: {reporte_path}")
            else:
                print(f"  [ERROR] No se pudo generar el reporte para {node_name}")
        
        except Exception as e:
            print(f"  [ERROR] Error al generar reporte para {node_name}: {e}")
            import traceback
            traceback.print_exc()
        
        print()
    
    return reportes_generados

def enviar_reportes_por_correo(reportes_paths: list[Path]) -> bool:
    """
    Envía los reportes generados por correo a los destinatarios.
    """
    if not reportes_paths:
        print("[ERROR] No hay reportes para enviar")
        return False
    
    print("=" * 70)
    print("ENVIANDO REPORTES POR CORREO")
    print("=" * 70)
    print(f"Destinatarios: {', '.join(DESTINATARIOS)}")
    print(f"Total de reportes: {len(reportes_paths)}")
    print()
    
    # Formatear fechas para el asunto
    start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
    end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
    start_date_str = start_dt.strftime("%d-%m-%Y")
    end_date_str = end_dt.strftime("%d-%m-%Y")
    
    # Crear mensaje
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(DESTINATARIOS)
    msg["Subject"] = f"Reportes Individuales BUPA - {start_date_str} a {end_date_str}"
    
    # Crear cuerpo del mensaje
    cuerpo = f"""
Estimados Diego y José,

Se adjuntan los reportes individuales de consumo y análisis de agua para BUPA correspondientes al período del {start_date_str} al {end_date_str}.

REPORTES INCLUIDOS:
"""
    
    for node_id, node_name in NODOS_SOLICITADOS.items():
        cuerpo += f"- {node_name} ({node_id})\n"
    
    cuerpo += f"""
Período analizado: {start_date_str} a {end_date_str}

Cada reporte contiene información detallada sobre:
- Consumo diario y total del período
- Patrones de consumo horario
- Detección de alertas de fuga
- Métricas estadísticas (promedios, máximos, mínimos)
- Gráficos y visualizaciones

Los reportes están en formato Word (.docx) y se adjuntan en este correo.

Saludos cordiales,
Sistema WES - Water Efficiency System
"""
    
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))
    
    # Adjuntar todos los reportes
    for reporte_path in reportes_paths:
        if not reporte_path.exists():
            print(f"  [ADVERTENCIA] Reporte no encontrado: {reporte_path}")
            continue
        
        try:
            with open(reporte_path, "rb") as f:
                adjunto = MIMEApplication(f.read(), _subtype="docx")
                adjunto.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=reporte_path.name
                )
                msg.attach(adjunto)
            print(f"  [OK] Adjuntado: {reporte_path.name}")
        except Exception as e:
            print(f"  [ERROR] No se pudo adjuntar {reporte_path.name}: {e}")
    
    # Enviar correo
    try:
        print()
        print("Conectando al servidor SMTP...")
        server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO)
        server.starttls()
        server.login(SMTP_USUARIO, SMTP_PASSWORD)
        
        print("Enviando correo...")
        server.sendmail(SMTP_USUARIO, DESTINATARIOS, msg.as_string())
        server.quit()
        
        print()
        print("[OK] Correo enviado exitosamente")
        return True
    
    except Exception as e:
        print()
        print(f"[ERROR] Error al enviar correo: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("=" * 70)
    print("  GENERACIÓN Y ENVÍO DE REPORTES BUPA")
    print("  Nodos: B1, B2, Central Térmica")
    print("=" * 70)
    print()
    
    # Generar reportes
    reportes = generar_reportes_individuales()
    
    if not reportes:
        print("[ERROR] No se generaron reportes. Abortando envío de correo.")
        return
    
    print()
    print("=" * 70)
    print("RESUMEN DE GENERACIÓN")
    print("=" * 70)
    print(f"Reportes generados exitosamente: {len(reportes)}/{len(NODOS_SOLICITADOS)}")
    for reporte in reportes:
        print(f"  - {reporte.name}")
    print()
    
    # Enviar correos
    exito = enviar_reportes_por_correo(reportes)
    
    print()
    print("=" * 70)
    if exito:
        print("  PROCESO COMPLETADO EXITOSAMENTE")
        print("=" * 70)
        print(f"Reportes generados: {len(reportes)}")
        print(f"Correos enviados a: {', '.join(DESTINATARIOS)}")
    else:
        print("  PROCESO COMPLETADO CON ERRORES")
        print("=" * 70)
        print("Los reportes se generaron pero hubo un error al enviar el correo.")
        print(f"Reportes disponibles en: {reportes[0].parent}")

if __name__ == "__main__":
    main()

