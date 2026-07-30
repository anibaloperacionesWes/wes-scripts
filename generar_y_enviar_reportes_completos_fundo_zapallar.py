"""
Script para generar reportes individuales, agregado y presentación de Fundo Zapallar
y enviarlos por correo a José, Diego y Aníbal.
Período: 01-01-2026 al 04-02-2026
"""

import sys
from pathlib import Path
from datetime import datetime
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
    generate_aggregated_report,
    ENTITY_BASE_URL,
)

# Importar funciones de PPT desde mall_maipu
from generar_reportes_y_ppt_mall_maipu import (
    obtener_datos_agregados,
    guardar_datos_json,
    crear_ppt_analisis,
    convertir_ppt_a_pdf,
)

# Configuración
COMPANY_ID = "000027"  # Fundo Zapallar
START_DATE = "01/01/2026"
END_DATE = "04/02/2026"

# Configuración SMTP
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

# Destinatarios
DESTINATARIOS = [
    "joseotarola@wes.cl",      # José
    "diegocarrasco@wes.cl",    # Diego
    "anibal.aoperaciones@wes.cl"  # Aníbal
]


def obtener_todos_nodos_fundo_zapallar(company_id: str) -> list:
    """Obtiene todos los nodos de Fundo Zapallar desde la API."""
    url = f"{ENTITY_BASE_URL}/companies/{company_id}"
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            all_nodes = data.get("nodes", [])
            nodes = []
            for node in all_nodes:
                nodes.append({
                    "nodeId": node.get("nodeId", ""),
                    "name": node.get("name", "").strip()
                })
            return nodes
        return []
    except Exception as e:
        print(f"[ERROR] Error al obtener nodos: {e}")
        return []


def enviar_correo_con_reportes(
    reportes_individuales: list,
    reporte_agregado_path: Path,
    ppt_path: Path,
    pdf_path: Path,
    destinatarios: list
):
    """Envía un correo con todos los reportes adjuntos."""
    try:
        # Crear mensaje
        msg = MIMEMultipart()
        msg['From'] = SMTP_USUARIO
        msg['To'] = ", ".join(destinatarios)
        msg['Subject'] = f"Reportes Fundo Zapallar - {START_DATE} al {END_DATE}"
        
        # Cuerpo del correo
        cuerpo = f"""
Estimados José, Diego y Aníbal,

Se adjuntan los reportes completos de Fundo Zapallar para el período {START_DATE} al {END_DATE}:

- Reportes individuales de todos los puntos de monitoreo
- Reporte agregado consolidado
- Presentación en PowerPoint (PPT)
- Presentación en PDF

Los reportes fueron generados automáticamente desde el sistema WES.

Saludos cordiales,
Sistema WES
"""
        msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))
        
        # Adjuntar reportes individuales (solo PDFs si existen)
        print(f"[INFO] Adjuntando reportes individuales...")
        reportes_adjuntados = 0
        for report_path in reportes_individuales:
            if isinstance(report_path, Path) and report_path.exists():
                # Buscar PDF correspondiente
                pdf_path_individual = report_path.parent / report_path.name.replace('.docx', '.pdf')
                if pdf_path_individual.exists():
                    try:
                        with open(pdf_path_individual, 'rb') as f:
                            adjunto = MIMEApplication(f.read())
                            adjunto.add_header(
                                'Content-Disposition',
                                'attachment',
                                filename=pdf_path_individual.name
                            )
                            msg.attach(adjunto)
                        reportes_adjuntados += 1
                    except Exception as e:
                        print(f"  [ADVERTENCIA] No se pudo adjuntar {pdf_path_individual.name}: {e}")
        
        print(f"  [OK] {reportes_adjuntados} reporte(s) individual(es) adjuntado(s)")
        
        # Adjuntar reporte agregado (PDF si existe)
        if reporte_agregado_path and reporte_agregado_path.exists():
            pdf_agregado = reporte_agregado_path.parent / reporte_agregado_path.name.replace('.docx', '.pdf')
            if pdf_agregado.exists():
                print(f"[INFO] Adjuntando reporte agregado: {pdf_agregado.name}")
                with open(pdf_agregado, 'rb') as f:
                    adjunto = MIMEApplication(f.read())
                    adjunto.add_header(
                        'Content-Disposition',
                        'attachment',
                        filename=pdf_agregado.name
                    )
                    msg.attach(adjunto)
        
        # Adjuntar presentación PPT
        if ppt_path and ppt_path.exists():
            print(f"[INFO] Adjuntando presentación PPT: {ppt_path.name}")
            with open(ppt_path, 'rb') as f:
                adjunto = MIMEApplication(f.read())
                adjunto.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=ppt_path.name
                )
                msg.attach(adjunto)
        
        # Adjuntar presentación PDF
        if pdf_path and pdf_path.exists():
            print(f"[INFO] Adjuntando presentación PDF: {pdf_path.name}")
            with open(pdf_path, 'rb') as f:
                adjunto = MIMEApplication(f.read())
                adjunto.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=pdf_path.name
                )
                msg.attach(adjunto)
        
        # Enviar correo
        print(f"[INFO] Enviando correo a {len(destinatarios)} destinatario(s)...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg)
        
        print(f"[OK] Correo enviado exitosamente a: {', '.join(destinatarios)}")
        return True
        
    except Exception as e:
        print(f"[ERROR] Error al enviar correo: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    print("=" * 70)
    print("  GENERACIÓN DE REPORTES COMPLETOS - FUNDO ZAPALLAR")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print("=" * 70)
    print()
    
    # Obtener información de la empresa
    try:
        company_name = get_company_name(COMPANY_ID)
        print(f"[INFO] Empresa: {company_name} (ID: {COMPANY_ID})")
    except Exception as e:
        print(f"[ERROR] No se pudo obtener el nombre de la empresa: {e}")
        company_name = "Fundo Zapallar"
    print()
    
    # Obtener todos los nodos desde la API
    print("[1/5] Obteniendo todos los nodos desde la API...")
    nodes = obtener_todos_nodos_fundo_zapallar(COMPANY_ID)
    
    if not nodes:
        print("[ERROR] No se encontraron nodos para Fundo Zapallar.")
        return
    
    print(f"[OK] Se encontraron {len(nodes)} nodo(s)")
    for node in nodes:
        print(f"  - {node['nodeId']}: {node['name']}")
    print()
    
    # Generar reportes individuales
    print("[2/5] Generando reportes individuales...")
    generated_reports = []
    node_ids = []
    
    for i, node in enumerate(nodes, 1):
        node_id = node["nodeId"]
        node_name = node["name"]
        node_ids.append(node_id)
        
        print(f"  [{i}/{len(nodes)}] Generando reporte para {node_id} ({node_name})...", flush=True)
        
        try:
            args = argparse.Namespace(
                company_id=COMPANY_ID,
                node_id=node_id,
                start_date=START_DATE,
                end_date=END_DATE,
                output_dir="reports",
                enviar_correo=False
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
    print(f"[OK] Reportes individuales generados: {len(generated_reports)}/{len(nodes)}")
    print()
    
    # Generar reporte agregado
    print("[3/5] Generando reporte agregado...")
    aggregated_report_path = None
    try:
        aggregated_report_path = generate_aggregated_report(
            COMPANY_ID,
            node_ids,
            START_DATE,
            END_DATE
        )
        if aggregated_report_path:
            print(f"[OK] Reporte agregado generado: {aggregated_report_path}")
        else:
            print("[ADVERTENCIA] No se pudo generar el reporte agregado")
    except Exception as e:
        print(f"[ERROR] Error al generar reporte agregado: {e}")
        import traceback
        traceback.print_exc()
    
    print()
    
    # Obtener datos agregados para la presentación
    print("[4/5] Obteniendo datos agregados para la presentación...")
    datos = None
    try:
        datos = obtener_datos_agregados(node_ids, START_DATE, END_DATE)
        if datos:
            print("[OK] Datos agregados obtenidos")
            # Guardar datos en JSON para referencia
            if aggregated_report_path:
                json_path = aggregated_report_path.parent / "datos_agregados.json"
                guardar_datos_json(datos, json_path)
        else:
            print("[ADVERTENCIA] No se pudieron obtener datos agregados")
    except Exception as e:
        print(f"[ERROR] Error al obtener datos agregados: {e}")
        import traceback
        traceback.print_exc()
    
    print()
    
    # Generar presentación PPT
    print("[5/5] Generando presentación PPT...")
    ppt_path = None
    pdf_path = None
    
    if datos and aggregated_report_path:
        try:
            # Crear directorio para PPT
            if aggregated_report_path.is_dir():
                ppt_dir = aggregated_report_path
            else:
                ppt_dir = aggregated_report_path.parent
            
            ppt_path = ppt_dir / "Presentacion Fundo Zapallar.pptx"
            crear_ppt_analisis(datos, ppt_path, mall_name="Fundo Zapallar")
            print(f"[OK] Presentación PPT generada: {ppt_path}")
            
            # Convertir PPT a PDF
            pdf_path = convertir_ppt_a_pdf(ppt_path)
            if pdf_path:
                print(f"[OK] Presentación PDF generada: {pdf_path}")
        except Exception as e:
            print(f"[ERROR] Error al generar presentación: {e}")
            import traceback
            traceback.print_exc()
    
    print()
    
    # Enviar correo con todos los reportes
    print("[6/6] Enviando correo con todos los reportes...")
    if generated_reports or aggregated_report_path or ppt_path:
        enviar_correo_con_reportes(
            generated_reports,
            aggregated_report_path,
            ppt_path,
            pdf_path,
            DESTINATARIOS
        )
    else:
        print("[ADVERTENCIA] No hay reportes para enviar")
    
    print()
    print("=" * 70)
    print("  PROCESO COMPLETADO")
    print("=" * 70)
    print(f"[OK] Reportes individuales: {len(generated_reports)}/{len(nodes)}")
    if aggregated_report_path:
        print(f"[OK] Reporte agregado: {aggregated_report_path}")
    if ppt_path:
        print(f"[OK] Presentación PPT: {ppt_path}")
    if pdf_path:
        print(f"[OK] Presentación PDF: {pdf_path}")


if __name__ == "__main__":
    main()
