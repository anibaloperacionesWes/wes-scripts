"""
Script para generar reportes individuales y agregado de todos los nodos de Parque Arauco
para diciembre 2025 y enviarlos por correo a Diego y Juan.
"""

import sys
import os
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
    generate_aggregated_report,
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
START_DATE = "2025-12-01"
END_DATE = "2025-12-31"

# Destinatarios
DESTINATARIOS = [
    "diegocarrasco@wes.cl",
    "juanlopez@wes.cl"
]

# URL base de la API
ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"


def obtener_nodos_parque_arauco():
    """
    Obtiene todos los nodos de Parque Arauco.
    """
    print(f"[1/6] Obteniendo nodos de Parque Arauco...")
    
    # Obtener información de la empresa
    url = f"{ENTITY_BASE_URL}/companies/{COMPANY_ID}"
    response = requests.get(url)
    response.raise_for_status()
    company_data = response.json()
    
    print(f"[OK] Empresa: {company_data.get('name', 'N/A')}")
    
    # Obtener todos los nodos
    all_nodes = company_data.get('nodes', [])
    print(f"[INFO] Total de nodos encontrados: {len(all_nodes)}")
    
    # Listar todos los nodos
    nodes_info = []
    for node in all_nodes:
        node_id = node.get('nodeId', '')
        node_name = node.get('name', '')
        nodes_info.append({
            'nodeId': node_id,
            'name': node_name
        })
        print(f"  [OK] Nodo encontrado: {node_id} - {node_name}")
    
    if not nodes_info:
        print(f"[ERROR] No se encontraron nodos para Parque Arauco")
        return []
    
    print(f"[OK] Total de nodos a procesar: {len(nodes_info)}")
    return nodes_info


def generar_reporte_individual(node_id: str, node_name: str):
    """
    Genera un reporte individual para un nodo específico.
    """
    print(f"  Generando reporte para {node_id} ({node_name})...")
    
    try:
        # Crear argumentos para generate_report
        args = argparse.Namespace(
            company_id=COMPANY_ID,
            node_id=node_id,
            start_date=START_DATE,
            end_date=END_DATE,
            output_dir="reports"
        )
        
        # Generar el reporte directamente
        reporte_path = generate_report(args)
        
        if reporte_path and reporte_path.exists():
            print(f"    [OK] Reporte generado: {reporte_path}")
            return reporte_path
        else:
            print(f"    [ERROR] No se pudo generar el reporte")
            return None
            
    except Exception as e:
        print(f"    [ERROR] Error al generar reporte: {e}")
        import traceback
        traceback.print_exc()
        return None


def generar_todos_los_reportes_individuales():
    """
    Genera reportes individuales para todos los nodos de Parque Arauco.
    """
    print(f"[2/6] Generando reportes individuales...")
    print(f"  Periodo: {START_DATE} a {END_DATE}")
    print("")
    
    # Obtener nodos de Parque Arauco
    nodes_info = obtener_nodos_parque_arauco()
    
    if not nodes_info:
        print("[ERROR] No se encontraron nodos para procesar")
        return []
    
    # Generar reporte para cada nodo
    reportes_generados = []
    for i, node_info in enumerate(nodes_info, 1):
        node_id = node_info['nodeId']
        node_name = node_info['name']
        print(f"[{i}/{len(nodes_info)}] Procesando nodo {node_id}...")
        
        reporte_path = generar_reporte_individual(node_id, node_name)
        if reporte_path:
            reportes_generados.append({
                'nodeId': node_id,
                'nodeName': node_name,
                'path': reporte_path
            })
        print("")
    
    print(f"[OK] Total de reportes individuales generados: {len(reportes_generados)}/{len(nodes_info)}")
    return reportes_generados


def generar_reporte_agregado():
    """
    Genera el reporte agregado de todos los nodos de Parque Arauco.
    """
    print(f"[3/6] Generando reporte agregado...")
    
    try:
        # Obtener todos los nodos desde la API
        url = f"{ENTITY_BASE_URL}/companies/{COMPANY_ID}"
        response = requests.get(url)
        response.raise_for_status()
        company_data = response.json()
        
        # Obtener todos los nodos
        all_nodes = company_data.get('nodes', [])
        nodos_ids = [node.get('nodeId', '') for node in all_nodes if node.get('nodeId')]
        
        if not nodos_ids:
            print("[ERROR] No se encontraron nodos para el reporte agregado")
            return None
        
        print(f"  [INFO] Generando reporte agregado para {len(nodos_ids)} nodos...")
        
        # Generar reporte agregado
        reporte_path = generate_aggregated_report(
            company_id=COMPANY_ID,
            node_ids=nodos_ids,
            start_date=START_DATE,
            end_date=END_DATE,
            output_dir="reports",
        )
        
        if reporte_path and reporte_path.exists():
            print(f"  [OK] Reporte agregado generado: {reporte_path}")
            return reporte_path
        else:
            print(f"  [ERROR] No se pudo generar el reporte agregado")
            return None
            
    except Exception as e:
        print(f"  [ERROR] Error al generar reporte agregado: {e}")
        import traceback
        traceback.print_exc()
        return None


def enviar_correo(reportes_individuales: list, reporte_agregado: Path = None):
    """
    Envía todos los reportes por correo a los destinatarios.
    """
    print(f"[4/6] Preparando envío de correo...")
    print(f"  Reportes individuales a enviar: {len(reportes_individuales)}")
    if reporte_agregado:
        print(f"  Reporte agregado: Sí")
    else:
        print(f"  Reporte agregado: No")
    print(f"  Destinatarios: {len(DESTINATARIOS)}")
    print("")
    
    if not reportes_individuales and not reporte_agregado:
        print("[ERROR] No hay reportes para enviar")
        return False
    
    company_name = get_company_name(COMPANY_ID)
    
    # Crear mensaje
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(DESTINATARIOS)
    msg["Subject"] = f"Reportes Parque Arauco - Diciembre 2025"
    
    # Crear cuerpo del mensaje
    cuerpo = f"""
Estimados Diego y Juan,

Se adjuntan los reportes de Parque Arauco correspondientes al mes de diciembre 2025.

Detalles:
- Empresa: {company_name}
- Periodo: {START_DATE} a {END_DATE}
- Total de reportes individuales: {len(reportes_individuales)}
"""
    
    if reporte_agregado:
        cuerpo += "- Reporte agregado: Sí\n"
    else:
        cuerpo += "- Reporte agregado: No disponible\n"
    
    if reportes_individuales:
        cuerpo += "\nReportes individuales incluidos:\n"
        for i, reporte in enumerate(reportes_individuales, 1):
            node_name = reporte['nodeName']
            cuerpo += f"{i}. {node_name}\n"
    
    cuerpo += """
Cada reporte contiene análisis detallado de consumo, alertas de consumo nocturno y métricas del punto de monitoreo correspondiente.

Saludos cordiales,
Sistema WES
"""
    
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))
    
    # Adjuntar reporte agregado primero (si existe)
    pdf_paths = []
    if reporte_agregado and reporte_agregado.exists():
        try:
            pdf_path = convertir_word_a_pdf(reporte_agregado)
            if pdf_path and pdf_path.exists():
                with open(pdf_path, "rb") as f:
                    adjunto = MIMEApplication(f.read(), _subtype="pdf")
                    adjunto.add_header(
                        "Content-Disposition",
                        "attachment",
                        filename=f"Reporte_Agregado_Parque_Arauco_Diciembre_2025.pdf"
                    )
                    msg.attach(adjunto)
                pdf_paths.append(pdf_path)
                print(f"  [OK] Adjuntado: Reporte Agregado Parque Arauco.pdf")
            else:
                # Si falla la conversión, adjuntar Word original
                print(f"  [ADVERTENCIA] No se pudo convertir a PDF, adjuntando Word original")
                with open(reporte_agregado, "rb") as f:
                    adjunto = MIMEApplication(f.read(), _subtype="docx")
                    adjunto.add_header(
                        "Content-Disposition",
                        "attachment",
                        filename=f"Reporte_Agregado_Parque_Arauco_Diciembre_2025.docx"
                    )
                    msg.attach(adjunto)
                print(f"  [OK] Adjuntado: Reporte Agregado Parque Arauco.docx")
        except Exception as e:
            print(f"  [ADVERTENCIA] Error al procesar reporte agregado: {e}")
    
    # Adjuntar todos los reportes individuales
    for reporte in reportes_individuales:
        reporte_path = reporte['path']
        node_name = reporte['nodeName']
        
        if not reporte_path.exists():
            print(f"  [ADVERTENCIA] El archivo no existe: {reporte_path}")
            continue
        
        # Convertir Word a PDF si es posible
        try:
            pdf_path = convertir_word_a_pdf(reporte_path)
            if pdf_path and pdf_path.exists():
                with open(pdf_path, "rb") as f:
                    adjunto = MIMEApplication(f.read(), _subtype="pdf")
                    # Limpiar nombre del nodo para el nombre del archivo
                    nombre_archivo = node_name.replace('/', '_').replace('\\', '_')
                    adjunto.add_header(
                        "Content-Disposition",
                        "attachment",
                        filename=f"Reporte_Parque_Arauco_{nombre_archivo}_Diciembre_2025.pdf"
                    )
                    msg.attach(adjunto)
                pdf_paths.append(pdf_path)
                print(f"  [OK] Adjuntado: Reporte - Parque Arauco {node_name}.pdf")
            else:
                # Si falla la conversión, adjuntar Word original
                print(f"  [ADVERTENCIA] No se pudo convertir a PDF, adjuntando Word original")
                with open(reporte_path, "rb") as f:
                    adjunto = MIMEApplication(f.read(), _subtype="docx")
                    nombre_archivo = node_name.replace('/', '_').replace('\\', '_')
                    adjunto.add_header(
                        "Content-Disposition",
                        "attachment",
                        filename=f"Reporte_Parque_Arauco_{nombre_archivo}_Diciembre_2025.docx"
                    )
                    msg.attach(adjunto)
                print(f"  [OK] Adjuntado: Reporte - Parque Arauco {node_name}.docx")
        except Exception as e:
            print(f"  [ADVERTENCIA] Error al procesar {node_name}: {e}")
            # Intentar adjuntar Word original
            try:
                with open(reporte_path, "rb") as f:
                    adjunto = MIMEApplication(f.read(), _subtype="docx")
                    nombre_archivo = node_name.replace('/', '_').replace('\\', '_')
                    adjunto.add_header(
                        "Content-Disposition",
                        "attachment",
                        filename=f"Reporte_Parque_Arauco_{nombre_archivo}_Diciembre_2025.docx"
                    )
                    msg.attach(adjunto)
                print(f"  [OK] Adjuntado: Reporte - Parque Arauco {node_name}.docx (Word)")
            except:
                print(f"  [ERROR] No se pudo adjuntar el reporte de {node_name}")
    
    print("")
    print(f"[5/6] Enviando correo...")
    
    # Enviar correo
    try:
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg)
        
        print(f"[OK] Correo enviado exitosamente a:")
        for destinatario in DESTINATARIOS:
            print(f"  - {destinatario}")
        
        # Limpiar PDFs temporales si existen
        for pdf_path in pdf_paths:
            try:
                if pdf_path.exists():
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
    parser = argparse.ArgumentParser(description="Genera reportes de Parque Arauco")
    parser.add_argument("--enviar-correo", action="store_true", 
                       help="Enviar los reportes por correo electrónico")
    args = parser.parse_args()
    ENVIAR_CORREO = args.enviar_correo
    
    print("=" * 70)
    print("  GENERACION DE REPORTES - PARQUE ARAUCO - DICIEMBRE 2025")
    print("=" * 70)
    print("")
    
    # Generar todos los reportes individuales
    reportes_individuales = generar_todos_los_reportes_individuales()
    
    # Generar reporte agregado
    print("")
    reporte_agregado = generar_reporte_agregado()
    
    # Enviar correo solo si se solicita
    print("")
    if ENVIAR_CORREO:
        exito = enviar_correo(reportes_individuales, reporte_agregado)
        if not exito:
            print("[ADVERTENCIA] No se pudo enviar el correo, pero los reportes se generaron correctamente.")
    else:
        print("[INFO] Envío de correo no solicitado. Solo se generaron los reportes.")
        print(f"[INFO] Para enviar correo, ejecuta: python {Path(__file__).name} --enviar-correo")
    
    print("")
    print("=" * 70)
    print("  PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 70)
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
