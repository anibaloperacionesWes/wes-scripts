"""
Script para generar reporte del Mall Maipú (Parque Arauco) para diciembre 2025
y enviarlo por correo a Diego y Juan.
"""

import sys
import argparse
from pathlib import Path
from datetime import datetime
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import (
    generate_aggregated_report,
    convertir_word_a_pdf,
    get_company_name,
)

# Variable global para controlar envío de correo
ENVIAR_CORREO = False

# Configuración SMTP
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

# Configuración del reporte
COMPANY_ID = "000025"  # Parque Arauco
START_DATE = "2025-12-01"
END_DATE = "2025-12-31"

# Destinatarios
DESTINATARIOS = [
    "diegocarrasco@wes.cl",
    "juanlopez@wes.cl"
]

# URL base de la API
ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"


def obtener_nodos_maipu():
    """
    Obtiene los nodos del Mall Maipú de Parque Arauco.
    """
    print(f"[1/4] Obteniendo nodos del Mall Maipú...")
    
    # Obtener información de la empresa
    url = f"{ENTITY_BASE_URL}/companies/{COMPANY_ID}"
    response = requests.get(url)
    response.raise_for_status()
    company_data = response.json()
    
    print(f"[OK] Empresa: {company_data.get('name', 'N/A')}")
    
    # Obtener todos los nodos
    all_nodes = company_data.get('nodes', [])
    print(f"[INFO] Total de nodos encontrados: {len(all_nodes)}")
    
    # Filtrar nodos de Maipú según el mapeo conocido
    # Basado en generar_reporte_word.py, los nodos de Maipú son:
    maipu_node_ids = {
        "000025-08": "Maipú",
        "000025-12": "Maipú",
        # Agregar más nodos según sea necesario
    }
    
    # También buscar por nombre
    maipu_nodes = []
    for node in all_nodes:
        node_id = node.get('nodeId', '')
        node_name = node.get('name', '').lower()
        
        # Verificar si está en el mapeo o si el nombre contiene "maipú" o "maipu"
        if (node_id in maipu_node_ids or 
            "maipú" in node_name or 
            "maipu" in node_name):
            maipu_nodes.append(node_id)
            print(f"  [OK] Nodo de Maipú encontrado: {node_id} - {node.get('name', 'N/A')}")
    
    if not maipu_nodes:
        print(f"[ADVERTENCIA] No se encontraron nodos específicos de Maipú, usando todos los nodos de Parque Arauco")
        maipu_nodes = [node.get('nodeId') for node in all_nodes]
    
    print(f"[OK] Total de nodos de Maipú a procesar: {len(maipu_nodes)}")
    return maipu_nodes


def generar_reporte():
    """
    Genera el reporte agregado del Mall Maipú.
    """
    print(f"[2/4] Generando reporte agregado...")
    print(f"  Periodo: {START_DATE} a {END_DATE}")
    
    # Obtener nodos de Maipú
    node_ids = obtener_nodos_maipu()
    
    if not node_ids:
        print("[ERROR] No se encontraron nodos para procesar")
        return None
    
    # Generar reporte agregado
    try:
        reporte_path = generate_aggregated_report(
            company_id=COMPANY_ID,
            node_ids=node_ids,
            start_date=START_DATE,
            end_date=END_DATE,
            output_dir="reports"
        )
        
        print(f"[OK] Reporte generado: {reporte_path}")
        return reporte_path
    except Exception as e:
        print(f"[ERROR] Error al generar reporte: {e}")
        import traceback
        traceback.print_exc()
        return None


def enviar_correo(reporte_path: Path, destinatarios: list):
    """
    Envía el reporte por correo a los destinatarios.
    """
    print(f"[3/4] Enviando correo a {len(destinatarios)} destinatario(s)...")
    
    if not reporte_path or not reporte_path.exists():
        print(f"[ERROR] El archivo del reporte no existe: {reporte_path}")
        return False
    
    company_name = get_company_name(COMPANY_ID)
    
    # Crear mensaje
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = f"Reporte Agregado - Parque Arauco MAll Maipú - Diciembre 2025"
    
    # Crear cuerpo del mensaje
    cuerpo = f"""
Estimados Diego y Juan,

Se adjunta el reporte agregado del Mall Maipú (Parque Arauco) correspondiente al mes de diciembre 2025.

Detalles del reporte:
- Empresa: {company_name}
- Mall: Maipú
- Periodo: {START_DATE} a {END_DATE}

Este reporte contiene análisis detallado de consumo, alertas de consumo nocturno y métricas consolidadas de todos los puntos de monitoreo del mall.

Saludos cordiales,
Sistema WES
"""
    
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))
    
    # Convertir Word a PDF si es posible
    pdf_path = None
    try:
        pdf_path = convertir_word_a_pdf(reporte_path)
        if pdf_path and pdf_path.exists():
            with open(pdf_path, "rb") as f:
                adjunto = MIMEApplication(f.read(), _subtype="pdf")
                adjunto.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=f"Reporte_Agregado_Mall_Maipu_Diciembre_2025.pdf"
                )
                msg.attach(adjunto)
            print(f"[OK] Reporte convertido a PDF para envío")
        else:
            # Si falla la conversión, adjuntar Word original
            print(f"[ADVERTENCIA] No se pudo convertir a PDF, adjuntando Word original")
            with open(reporte_path, "rb") as f:
                adjunto = MIMEApplication(f.read(), _subtype="docx")
                adjunto.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=f"Reporte_Agregado_Mall_Maipu_Diciembre_2025.docx"
                )
                msg.attach(adjunto)
    except Exception as e:
        print(f"[ADVERTENCIA] Error al convertir a PDF: {e}")
        # Adjuntar Word original
        with open(reporte_path, "rb") as f:
            adjunto = MIMEApplication(f.read(), _subtype="docx")
            adjunto.add_header(
                "Content-Disposition",
                "attachment",
                filename=f"Reporte_Agregado_Mall_Maipu_Diciembre_2025.docx"
            )
            msg.attach(adjunto)
    
    # Enviar correo
    try:
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg)
        
        print(f"[OK] Correo enviado exitosamente a:")
        for destinatario in destinatarios:
            print(f"  - {destinatario}")
        
        # Limpiar PDF temporal si existe
        if pdf_path and pdf_path.exists():
            try:
                pdf_path.unlink()
                print(f"[OK] Archivo PDF temporal eliminado")
            except:
                pass
        
        return True
    except Exception as e:
        print(f"[ERROR] Error al enviar correo: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """
    Función principal.
    """
    global ENVIAR_CORREO
    
    # Parsear argumentos de línea de comandos
    parser = argparse.ArgumentParser(description="Genera reporte agregado del Mall Maipú")
    parser.add_argument("--enviar-correo", action="store_true", 
                       help="Enviar el reporte por correo electrónico")
    args = parser.parse_args()
    ENVIAR_CORREO = args.enviar_correo
    
    print("=" * 70)
    print("  GENERACION DE REPORTE - MALL MAIPU - DICIEMBRE 2025")
    print("=" * 70)
    print("")
    
    # Generar reporte
    reporte_path = generar_reporte()
    
    if not reporte_path:
        print("[ERROR] No se pudo generar el reporte")
        return 1
    
    # Enviar correo solo si se solicita
    print("")
    if ENVIAR_CORREO:
        exito = enviar_correo(reporte_path, DESTINATARIOS)
        if not exito:
            print("[ADVERTENCIA] No se pudo enviar el correo, pero el reporte se generó correctamente.")
    else:
        print("[INFO] Envío de correo no solicitado. Solo se generó el reporte.")
        print(f"[INFO] Para enviar correo, ejecuta: python {Path(__file__).name} --enviar-correo")
    
    print("")
    print("=" * 70)
    print("  PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 70)
    
    return 0


if __name__ == "__main__":
    sys.exit(main())

