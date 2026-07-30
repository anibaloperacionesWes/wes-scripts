"""
Bot de WhatsApp para solicitar y generar reportes WES automáticamente.

Este bot escucha mensajes de WhatsApp, procesa solicitudes de reportes,
y envía los reportes generados de vuelta por WhatsApp.

Requisitos:
- pip install twilio flask python-dotenv
- Cuenta de Twilio con WhatsApp Business API habilitada
- Variables de entorno configuradas (ver .env.example)
"""

import os
import re
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import requests
from dotenv import load_dotenv
from flask import Flask, request
from twilio.rest import Client
from twilio.twiml.messaging_response import MessagingResponse

# Cargar variables de entorno
load_dotenv()

# Configuración Flask
app = Flask(__name__)

# Configuración Twilio
TWILIO_ACCOUNT_SID = os.getenv("TWILIO_ACCOUNT_SID")
TWILIO_AUTH_TOKEN = os.getenv("TWILIO_AUTH_TOKEN")
TWILIO_WHATSAPP_FROM = os.getenv("TWILIO_WHATSAPP_FROM")  # Ej: "whatsapp:+14155238886"
TWILIO_WHATSAPP_TO = os.getenv("TWILIO_WHATSAPP_TO")  # Número autorizado para recibir mensajes

# Configuración WES
ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

# Inicializar cliente Twilio
if TWILIO_ACCOUNT_SID and TWILIO_AUTH_TOKEN:
    twilio_client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
else:
    twilio_client = None
    print("[ADVERTENCIA] Twilio no configurado. El bot no funcionará sin credenciales.")


def get_company_name(company_id: str) -> Optional[str]:
    """Obtiene el nombre de la empresa desde la API."""
    try:
        url = f"{ENTITY_BASE_URL}/companies/{company_id}"
        response = requests.get(url, headers={"Accept": "application/json"})
        if response.status_code == 200:
            data = response.json()
            return data.get("name")
    except Exception as e:
        print(f"[ERROR] Error al obtener nombre de empresa: {e}")
    return None


def get_node_name(node_id: str) -> Optional[str]:
    """Obtiene el nombre del nodo desde la API."""
    try:
        # Extraer company_id del node_id (formato: COMPANY_ID-NODE_NUMBER)
        company_id = node_id.split("-")[0]
        url = f"{ENTITY_BASE_URL}/companies/{company_id}"
        response = requests.get(url, headers={"Accept": "application/json"})
        if response.status_code == 200:
            data = response.json()
            nodes = data.get("nodes", [])
            for node in nodes:
                if node.get("nodeId") == node_id:
                    return node.get("name")
    except Exception as e:
        print(f"[ERROR] Error al obtener nombre de nodo: {e}")
    return None


def parse_report_request(message: str) -> Optional[Dict]:
    """
    Parsea un mensaje de texto para extraer parámetros del reporte.
    
    Formatos soportados:
    - "reporte empresa 000025 nodo 000025-12 desde 01/12/2025 hasta 15/12/2025"
    - "reporte 000025 000025-12 01/12/2025 15/12/2025"
    - "reporte empresa 000025 nodo 000025-12 ultimos 7 dias"
    - "reporte agregado empresa 000025 desde 01/12/2025 hasta 15/12/2025"
    """
    message = message.lower().strip()
    
    # Patrones de búsqueda
    patterns = [
        # Formato: "reporte empresa X nodo Y desde DD/MM/YYYY hasta DD/MM/YYYY"
        r"reporte\s+(?:empresa\s+)?(\d+)\s+(?:nodo\s+)?([\d-]+)\s+desde\s+(\d{1,2}/\d{1,2}/\d{4})\s+hasta\s+(\d{1,2}/\d{1,2}/\d{4})",
        # Formato: "reporte X Y DD/MM/YYYY DD/MM/YYYY"
        r"reporte\s+(\d+)\s+([\d-]+)\s+(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}/\d{1,2}/\d{4})",
        # Formato: "reporte empresa X nodo Y ultimos N dias"
        r"reporte\s+(?:empresa\s+)?(\d+)\s+(?:nodo\s+)?([\d-]+)\s+ultimos\s+(\d+)\s+dias?",
        # Formato: "reporte agregado empresa X desde DD/MM/YYYY hasta DD/MM/YYYY"
        r"reporte\s+agregado\s+(?:empresa\s+)?(\d+)\s+desde\s+(\d{1,2}/\d{1,2}/\d{4})\s+hasta\s+(\d{1,2}/\d{1,2}/\d{4})",
    ]
    
    for pattern in patterns:
        match = re.search(pattern, message)
        if match:
            groups = match.groups()
            
            # Formato con fechas específicas (individual)
            if len(groups) == 4 and "/" in groups[2]:
                company_id = groups[0].zfill(6)  # Asegurar formato 000025
                node_id = groups[1]
                start_date = groups[2]
                end_date = groups[3]
                return {
                    "type": "individual",
                    "company_id": company_id,
                    "node_id": node_id,
                    "start_date": start_date,
                    "end_date": end_date,
                }
            
            # Formato con "últimos N días" (individual)
            elif len(groups) == 3 and groups[2].isdigit():
                company_id = groups[0].zfill(6)
                node_id = groups[1]
                days = int(groups[2])
                end_date = datetime.now() - timedelta(days=1)
                start_date = end_date - timedelta(days=days - 1)
                return {
                    "type": "individual",
                    "company_id": company_id,
                    "node_id": node_id,
                    "start_date": start_date.strftime("%d/%m/%Y"),
                    "end_date": end_date.strftime("%d/%m/%Y"),
                }
            
            # Formato agregado
            elif len(groups) == 3 and "agregado" in message:
                company_id = groups[0].zfill(6)
                start_date = groups[1]
                end_date = groups[2]
                return {
                    "type": "agregado",
                    "company_id": company_id,
                    "start_date": start_date,
                    "end_date": end_date,
                }
    
    return None


def generate_individual_report(company_id: str, node_id: str, start_date: str, end_date: str) -> Optional[Path]:
    """Genera un reporte individual."""
    try:
        cmd = [
            PYTHON_EXE,
            SCRIPT_PATH,
            "--company-id", company_id,
            "--node-id", node_id,
            "--start-date", start_date,
            "--end-date", end_date,
        ]
        
        print(f"[INFO] Generando reporte individual: {company_id} - {node_id}")
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore'
        )
        
        if result.returncode == 0:
            # Buscar el archivo generado en la salida
            output_lines = result.stdout.split("\n")
            for line in output_lines:
                if "Reporte generado en:" in line or "reports" in line.lower():
                    path_str = line.split(":")[-1].strip() if ":" in line else line.strip()
                    if path_str:
                        report_path = Path(path_str)
                        if report_path.exists():
                            return report_path
            
            # Buscar en el directorio de reportes
            reports_dir = Path("reports")
            if reports_dir.exists():
                # Buscar el archivo más reciente para este nodo
                pattern = f"*{node_id}*"
                matching_files = list(reports_dir.rglob(f"*{node_id}*.docx"))
                if matching_files:
                    return max(matching_files, key=lambda p: p.stat().st_mtime)
        
        print(f"[ERROR] Error al generar reporte: {result.stderr}")
        return None
        
    except Exception as e:
        print(f"[ERROR] Excepción al generar reporte: {e}")
        return None


def generate_aggregated_report(company_id: str, start_date: str, end_date: str, fuente_agua_id: Optional[str] = None) -> Optional[Path]:
    """Genera un reporte agregado."""
    try:
        from generar_reporte_word import generate_aggregated_report as gen_agg
        
        # Obtener todos los nodos de la empresa
        url = f"{ENTITY_BASE_URL}/companies/{company_id}"
        response = requests.get(url, headers={"Accept": "application/json"})
        if response.status_code != 200:
            print(f"[ERROR] No se pudo obtener información de la empresa {company_id}")
            return None
        
        data = response.json()
        nodes = data.get("nodes", [])
        node_ids = [node.get("nodeId") for node in nodes if node.get("nodeId")]
        
        if not node_ids:
            print(f"[ERROR] No se encontraron nodos para la empresa {company_id}")
            return None
        
        print(f"[INFO] Generando reporte agregado para {len(node_ids)} nodos")
        
        # Generar reporte agregado
        report_path = gen_agg(
            company_id=company_id,
            node_ids=node_ids,
            start_date=start_date,
            end_date=end_date,
            fuente_agua_id=fuente_agua_id,
        )
        
        return report_path
        
    except Exception as e:
        print(f"[ERROR] Excepción al generar reporte agregado: {e}")
        import traceback
        traceback.print_exc()
        return None


def convert_docx_to_pdf(docx_path: Path) -> Optional[Path]:
    """Convierte un archivo DOCX a PDF."""
    try:
        from generar_reporte_word import convertir_word_a_pdf
        pdf_path = convertir_word_a_pdf(docx_path)
        if pdf_path and pdf_path.exists():
            return pdf_path
    except Exception as e:
        print(f"[ERROR] Error al convertir a PDF: {e}")
    return None


def send_whatsapp_message(to: str, message: str, media_url: Optional[str] = None) -> bool:
    """Envía un mensaje de WhatsApp usando Twilio."""
    if not twilio_client:
        print("[ERROR] Twilio no configurado")
        return False
    
    try:
        message_params = {
            "from": TWILIO_WHATSAPP_FROM,
            "to": to,
            "body": message,
        }
        
        if media_url:
            message_params["media_url"] = [media_url]
        
        message = twilio_client.messages.create(**message_params)
        print(f"[OK] Mensaje enviado: {message.sid}")
        return True
        
    except Exception as e:
        print(f"[ERROR] Error al enviar mensaje WhatsApp: {e}")
        return False


def send_file_via_whatsapp(to: str, file_path: Path, message: str = "") -> bool:
    """
    Envía un archivo por WhatsApp.
    
    Nota: Twilio requiere que el archivo esté en una URL pública.
    Para producción, necesitarás subir el archivo a un servidor web o usar un servicio de almacenamiento.
    """
    # Por ahora, solo enviamos un mensaje con la ubicación del archivo
    # En producción, necesitarías subir el archivo a un servidor web
    file_info = f"""
📄 Reporte generado exitosamente

📁 Archivo: {file_path.name}
📍 Ubicación: {file_path}

⚠️ Nota: Para recibir el archivo directamente por WhatsApp, el bot necesita acceso a un servidor web para alojar el archivo temporalmente.

Por ahora, puedes encontrar el reporte en la carpeta de reportes.
"""
    
    return send_whatsapp_message(to, file_info)


@app.route("/whatsapp", methods=["POST"])
def whatsapp_webhook():
    """Webhook para recibir mensajes de WhatsApp."""
    incoming_message = request.values.get("Body", "").strip()
    from_number = request.values.get("From", "")
    
    print(f"[INFO] Mensaje recibido de {from_number}: {incoming_message}")
    
    # Verificar que el número esté autorizado
    if TWILIO_WHATSAPP_TO and from_number not in TWILIO_WHATSAPP_TO:
        print(f"[ADVERTENCIA] Número no autorizado: {from_number}")
        resp = MessagingResponse()
        resp.message("Lo siento, no estás autorizado para usar este bot.")
        return str(resp)
    
    resp = MessagingResponse()
    
    # Comando de ayuda
    if incoming_message.lower() in ["ayuda", "help", "comandos"]:
        help_text = """
🤖 *Bot de Reportes WES*

*Comandos disponibles:*

1. *Reporte Individual:*
   `reporte empresa 000025 nodo 000025-12 desde 01/12/2025 hasta 15/12/2025`
   `reporte 000025 000025-12 01/12/2025 15/12/2025`
   `reporte empresa 000025 nodo 000025-12 ultimos 7 dias`

2. *Reporte Agregado:*
   `reporte agregado empresa 000025 desde 01/12/2025 hasta 15/12/2025`

*Ejemplos:*
- `reporte 000025 000025-12 ultimos 7 dias`
- `reporte agregado 000025 desde 01/12/2025 hasta 15/12/2025`

Escribe *ayuda* para ver este mensaje nuevamente.
"""
        resp.message(help_text)
        return str(resp)
    
    # Procesar solicitud de reporte
    report_params = parse_report_request(incoming_message)
    
    if not report_params:
        resp.message(
            "❌ No pude entender tu solicitud.\n\n"
            "Formato esperado:\n"
            "`reporte empresa 000025 nodo 000025-12 desde 01/12/2025 hasta 15/12/2025`\n\n"
            "Escribe *ayuda* para ver todos los comandos disponibles."
        )
        return str(resp)
    
    # Enviar mensaje de confirmación
    if report_params["type"] == "individual":
        company_name = get_company_name(report_params["company_id"])
        node_name = get_node_name(report_params["node_id"])
        confirm_msg = f"""
✅ Solicitud recibida:

📊 *Reporte Individual*
🏢 Empresa: {report_params["company_id"]} {f"({company_name})" if company_name else ""}
📍 Nodo: {report_params["node_id"]} {f"({node_name})" if node_name else ""}
📅 Periodo: {report_params["start_date"]} a {report_params["end_date"]}

⏳ Generando reporte... Esto puede tomar unos minutos.
"""
    else:
        company_name = get_company_name(report_params["company_id"])
        confirm_msg = f"""
✅ Solicitud recibida:

📊 *Reporte Agregado*
🏢 Empresa: {report_params["company_id"]} {f"({company_name})" if company_name else ""}
📅 Periodo: {report_params["start_date"]} a {report_params["end_date"]}

⏳ Generando reporte... Esto puede tomar unos minutos.
"""
    
    resp.message(confirm_msg)
    
    # Generar reporte en segundo plano (en producción, usar una cola de tareas)
    try:
        if report_params["type"] == "individual":
            report_path = generate_individual_report(
                report_params["company_id"],
                report_params["node_id"],
                report_params["start_date"],
                report_params["end_date"],
            )
        else:
            # Para reportes agregados, necesitaríamos preguntar por fuente de agua
            # Por ahora, asumimos que no hay fuente
            report_path = generate_aggregated_report(
                report_params["company_id"],
                report_params["start_date"],
                report_params["end_date"],
                fuente_agua_id=None,
            )
        
        if report_path and report_path.exists():
            # Intentar convertir a PDF
            pdf_path = convert_docx_to_pdf(report_path)
            
            # Enviar mensaje de éxito
            success_msg = f"""
✅ *Reporte generado exitosamente*

📄 Archivo: {report_path.name}
📁 Ubicación: {report_path.parent}

⚠️ *Nota:* Para recibir el archivo directamente por WhatsApp, el bot necesita estar configurado con un servidor web para alojar archivos temporalmente.

Por ahora, puedes encontrar el reporte en la carpeta de reportes de tu servidor.
"""
            send_whatsapp_message(from_number, success_msg)
        else:
            error_msg = "❌ Error al generar el reporte. Por favor, verifica los parámetros e intenta nuevamente."
            send_whatsapp_message(from_number, error_msg)
            
    except Exception as e:
        error_msg = f"❌ Error inesperado: {str(e)}"
        send_whatsapp_message(from_number, error_msg)
        print(f"[ERROR] Error al procesar solicitud: {e}")
        import traceback
        traceback.print_exc()
    
    return str(resp)


@app.route("/health", methods=["GET"])
def health_check():
    """Endpoint de salud para verificar que el bot está funcionando."""
    return {"status": "ok", "service": "whatsapp-bot-wes"}


if __name__ == "__main__":
    # Verificar que las variables de entorno estén configuradas
    if not TWILIO_ACCOUNT_SID or not TWILIO_AUTH_TOKEN:
        print("[ERROR] Variables de entorno de Twilio no configuradas.")
        print("Por favor, crea un archivo .env con:")
        print("TWILIO_ACCOUNT_SID=tu_account_sid")
        print("TWILIO_AUTH_TOKEN=tu_auth_token")
        print("TWILIO_WHATSAPP_FROM=whatsapp:+14155238886")
        print("TWILIO_WHATSAPP_TO=whatsapp:+56912345678")
        sys.exit(1)
    
    print("=" * 70)
    print("🤖 BOT DE WHATSAPP PARA REPORTES WES")
    print("=" * 70)
    print(f"📱 Twilio configurado: {TWILIO_ACCOUNT_SID[:10]}...")
    print(f"📞 Número autorizado: {TWILIO_WHATSAPP_TO}")
    print()
    print("🌐 Iniciando servidor Flask...")
    print("📡 El bot escuchará en: http://localhost:5000/whatsapp")
    print()
    print("⚠️  IMPORTANTE: Configura el webhook de Twilio para apuntar a:")
    print("   https://tu-dominio.com/whatsapp")
    print("   (Usa ngrok o similar para desarrollo local)")
    print("=" * 70)
    
    # Ejecutar servidor Flask
    app.run(host="0.0.0.0", port=5000, debug=True)

