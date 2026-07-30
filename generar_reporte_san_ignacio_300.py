"""
Script para generar reporte del nodo San Ignacio 300 de Parque Arauco
desde 03 de diciembre a 31 de diciembre 2025 y enviarlo por correo.
"""

import sys
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
    convertir_word_a_pdf,
    get_company_name,
    get_node_name,
    generate_report,
)
import argparse

# Variable global para controlar envío de correo
ENVIAR_CORREO = False

# Configuración SMTP
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

# Configuración del reporte
COMPANY_ID = "000025"  # Parque Arauco
START_DATE = "2025-12-03"
END_DATE = "2025-12-31"

# Destinatarios
DESTINATARIOS = [
    "diegocarrasco@wes.cl",
    "juanlopez@wes.cl"
]

# URL base de la API
ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"


def obtener_node_id_san_ignacio_300():
    """
    Obtiene el ID del nodo San Ignacio 300 de Parque Arauco.
    """
    print(f"[1/4] Buscando nodo San Ignacio 300...")
    
    # Obtener información de la empresa
    url = f"{ENTITY_BASE_URL}/companies/{COMPANY_ID}"
    response = requests.get(url)
    response.raise_for_status()
    company_data = response.json()
    
    print(f"[OK] Empresa: {company_data.get('name', 'N/A')}")
    
    # Buscar el nodo San Ignacio 300
    all_nodes = company_data.get('nodes', [])
    node_id = None
    node_name = None
    
    for node in all_nodes:
        node_name_check = node.get('name', '')
        if 'San Ignacio 300' in node_name_check or 'san ignacio 300' in node_name_check.lower():
            node_id = node.get('nodeId', '')
            node_name = node_name_check
            break
    
    if not node_id:
        print(f"[ERROR] No se encontró el nodo 'San Ignacio 300' en Parque Arauco")
        print(f"[INFO] Nodos disponibles:")
        for node in all_nodes:
            print(f"  - {node.get('nodeId', '')}: {node.get('name', '')}")
        return None, None
    
    print(f"[OK] Nodo encontrado: {node_id} - {node_name}")
    return node_id, node_name


def generar_reporte(node_id: str):
    """
    Genera el reporte individual para San Ignacio 300.
    """
    print(f"[2/4] Generando reporte...")
    print(f"  Periodo: {START_DATE} a {END_DATE}")
    
    try:
        # Crear argumentos para generate_report
        args = argparse.Namespace(
            company_id=COMPANY_ID,
            node_id=node_id,
            start_date=START_DATE,
            end_date=END_DATE,
            output_dir="reports"
        )
        
        # Generar el reporte
        reporte_path = generate_report(args)
        
        if reporte_path and reporte_path.exists():
            print(f"[OK] Reporte generado: {reporte_path}")
            return reporte_path
        else:
            print(f"[ERROR] No se pudo generar el reporte")
            return None
            
    except Exception as e:
        print(f"[ERROR] Error al generar reporte: {e}")
        import traceback
        traceback.print_exc()
        return None


def enviar_reporte_por_correo(reporte_path: Path, node_id: str, node_name: str):
    """
    Envía el reporte generado por correo electrónico.
    """
    if not reporte_path or not reporte_path.exists():
        print(f"[ERROR] El archivo del reporte no existe: {reporte_path}")
        return False
    
    print(f"[3/4] Preparando envío de correo...")
    
    company_name = get_company_name(COMPANY_ID)
    
    try:
        # Crear mensaje
        msg = MIMEMultipart()
        msg["From"] = SMTP_USUARIO
        msg["To"] = ", ".join(DESTINATARIOS)
        
        # Crear asunto
        asunto = f"Reporte - Parque Arauco {node_name} - Diciembre 2025"
        msg["Subject"] = asunto
        
        # Crear cuerpo del mensaje
        cuerpo = f"""
Estimados Diego y Juan,

Se adjunta el reporte de consumo y fugas generado para el nodo {node_name} de {company_name}.

Detalles:
- Nodo: {node_name}
- Periodo: {START_DATE} a {END_DATE}

Este reporte contiene análisis detallado de consumo, alertas de consumo nocturno y métricas del punto de monitoreo.

Saludos cordiales,
Sistema WES
"""
        
        msg.attach(MIMEText(cuerpo, "plain", "utf-8"))
        
        # Convertir Word a PDF si es posible
        pdf_path = convertir_word_a_pdf(reporte_path)
        
        if pdf_path and pdf_path.exists():
            # Adjuntar PDF
            with open(pdf_path, "rb") as f:
                adjunto = MIMEApplication(f.read(), _subtype="pdf")
                nombre_archivo = node_name.replace('/', '_').replace('\\', '_')
                adjunto.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=f"Reporte_Parque_Arauco_{nombre_archivo}_Diciembre_2025.pdf"
                )
                msg.attach(adjunto)
            print(f"[OK] Reporte convertido a PDF para envío")
        else:
            # Si falla la conversión, adjuntar Word original
            print(f"[ADVERTENCIA] No se pudo convertir a PDF, adjuntando Word original")
            with open(reporte_path, "rb") as f:
                adjunto = MIMEApplication(f.read(), _subtype="docx")
                nombre_archivo = node_name.replace('/', '_').replace('\\', '_')
                adjunto.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=f"Reporte_Parque_Arauco_{nombre_archivo}_Diciembre_2025.docx"
                )
                msg.attach(adjunto)
        
        # Enviar correo
        print(f"[4/4] Enviando correo...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg)
        
        print(f"[OK] Correo enviado exitosamente a:")
        for dest in DESTINATARIOS:
            print(f"  - {dest}")
        
        # Limpiar PDF temporal si existe
        if pdf_path and pdf_path.exists():
            try:
                pdf_path.unlink()
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
    parser = argparse.ArgumentParser(description="Genera reporte de San Ignacio 300")
    parser.add_argument("--enviar-correo", action="store_true", 
                       help="Enviar el reporte por correo electrónico")
    args = parser.parse_args()
    ENVIAR_CORREO = args.enviar_correo
    
    print("=" * 70)
    print("  GENERACION DE REPORTE: PARQUE ARAUCO - SAN IGNACIO 300")
    print("=" * 70)
    print("")
    
    # Obtener node ID primero
    node_id, node_name = obtener_node_id_san_ignacio_300()
    if not node_id:
        print("[ERROR] No se pudo obtener el ID del nodo")
        return 1
    
    # Generar reporte
    reporte_path = generar_reporte(node_id)
    
    if not reporte_path:
        print("[ERROR] No se pudo generar el reporte")
        return 1
    
    # Enviar correo solo si se solicita
    print("")
    if ENVIAR_CORREO:
        exito = enviar_reporte_por_correo(reporte_path, node_id, node_name)
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

