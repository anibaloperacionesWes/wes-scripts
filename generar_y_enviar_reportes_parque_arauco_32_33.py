"""
Script para generar reportes individuales de los nodos 000025-32 y 000025-33
y enviarlos por correo a anibal.
Período: 01-01-2026 al 05-02-2026
"""

import sys
from pathlib import Path
import argparse
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import (
    get_company_name,
    get_node_name,
    generate_report,
    ENTITY_BASE_URL,
)

# Configuración
COMPANY_ID = "000025"  # Parque Arauco
START_DATE = "01/01/2026"
END_DATE = "05/02/2026"

# IDs de nodos solicitados
NODE_IDS = ["000025-32", "000025-33"]

# Configuración SMTP
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

# Destinatario
DESTINATARIO = "anibal.aoperaciones@wes.cl"


def obtener_nodos_info(company_id: str, node_ids: list) -> list:
    """Obtiene información de los nodos especificados."""
    url = f"{ENTITY_BASE_URL}/companies/{company_id}"
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            all_nodes = data.get("nodes", [])
            nodes_encontrados = []
            
            for node in all_nodes:
                node_id = node.get("nodeId", "")
                if node_id in node_ids:
                    nodes_encontrados.append({
                        "nodeId": node_id,
                        "name": node.get("name", "").strip()
                    })
            
            return nodes_encontrados
        return []
    except Exception as e:
        print(f"[ERROR] Error al obtener nodos: {e}")
        return []


def enviar_correo_con_reportes(reportes_paths: list, destinatario: str):
    """Envía un correo con los reportes adjuntos."""
    if not reportes_paths:
        print("[ADVERTENCIA] No hay reportes para enviar")
        return False
    
    try:
        # Crear mensaje
        msg = MIMEMultipart()
        msg['From'] = SMTP_USUARIO
        msg['To'] = destinatario
        msg['Subject'] = f"Reportes Individuales - Parque Arauco (Nodos 000025-32, 000025-33) - {START_DATE} al {END_DATE}"
        
        # Cuerpo del correo
        cuerpo = f"""
Estimado Aníbal,

Se adjuntan los reportes individuales solicitados para los siguientes puntos de monitoreo:

- Nodo 000025-32
- Nodo 000025-33

Período: {START_DATE} al {END_DATE}

Los reportes fueron generados automáticamente desde el sistema WES.

Saludos cordiales,
Sistema WES
"""
        msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))
        
        # Adjuntar reportes
        for report_path in reportes_paths:
            if Path(report_path).exists():
                with open(report_path, 'rb') as f:
                    adjunto = MIMEApplication(f.read())
                    adjunto.add_header(
                        'Content-Disposition',
                        'attachment',
                        filename=Path(report_path).name
                    )
                    msg.attach(adjunto)
                print(f"  [OK] Adjuntado: {Path(report_path).name}")
            else:
                print(f"  [ADVERTENCIA] No se encontró el archivo: {report_path}")
        
        # Enviar correo
        print(f"[INFO] Enviando correo a {destinatario}...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg)
        
        print(f"[OK] Correo enviado exitosamente a {destinatario}")
        return True
        
    except Exception as e:
        print(f"[ERROR] Error al enviar correo: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    print("=" * 70)
    print("  GENERACIÓN Y ENVÍO DE REPORTES INDIVIDUALES")
    print(f"  Nodos: {', '.join(NODE_IDS)}")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print("=" * 70)
    print()
    
    # Obtener información de la empresa
    try:
        company_name = get_company_name(COMPANY_ID)
        print(f"[INFO] Empresa: {company_name} (ID: {COMPANY_ID})")
    except Exception as e:
        print(f"[ERROR] No se pudo obtener el nombre de la empresa: {e}")
        print(f"[INFO] Continuando con ID: {COMPANY_ID}")
        company_name = "Parque Arauco"
    print()
    
    # Obtener información de los nodos
    print("[1/3] Obteniendo información de los nodos...")
    nodes = obtener_nodos_info(COMPANY_ID, NODE_IDS)
    
    if not nodes:
        print(f"[ERROR] No se encontraron los nodos solicitados: {NODE_IDS}")
        return
    
    print(f"[OK] Se encontraron {len(nodes)} nodo(s)")
    for node in nodes:
        print(f"  - {node['nodeId']}: {node['name']}")
    print()
    
    # Generar reportes individuales
    print("[2/3] Generando reportes individuales...")
    generated_reports = []
    
    for i, node in enumerate(nodes, 1):
        node_id = node["nodeId"]
        node_name = node["name"]
        
        print(f"  [{i}/{len(nodes)}] Generando reporte para {node_id} ({node_name})...", flush=True)
        
        try:
            args = argparse.Namespace(
                company_id=COMPANY_ID,
                node_id=node_id,
                start_date=START_DATE,
                end_date=END_DATE,
                output_dir="reports",
                enviar_correo=False  # No enviar correos automáticamente, lo haremos manualmente
            )
            report_path = generate_report(args)
            if report_path:
                generated_reports.append(report_path)
                print(f"    [OK] Reporte generado: {report_path}", flush=True)
            else:
                print(f"    [ADVERTENCIA] generate_report retornó None para {node_id}", flush=True)
        except Exception as e:
            print(f"    [ERROR] No se pudo generar el reporte para {node_id}: {e}", flush=True)
            import traceback
            traceback.print_exc()
            continue
    
    print()
    
    # Enviar correo con los reportes
    if generated_reports:
        print("[3/3] Enviando correo con los reportes...")
        enviar_correo_con_reportes(generated_reports, DESTINATARIO)
    else:
        print("[ADVERTENCIA] No se generaron reportes para enviar")
    
    print()
    print("=" * 70)
    print("  PROCESO COMPLETADO")
    print("=" * 70)
    print(f"[OK] Reportes generados: {len(generated_reports)}/{len(nodes)}")
    for report_path in generated_reports:
        print(f"  - {report_path}")


if __name__ == "__main__":
    main()
