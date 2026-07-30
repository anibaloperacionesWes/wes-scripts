"""
Sistema automatizado para monitorear correos y generar reportes personalizados.
Lee correos, detecta solicitudes de reportes y genera/envía reportes personalizados
basados en la información de la lista de contactos.
"""

import os
import sys
import re
import base64
from pathlib import Path
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email import message_from_string
import smtplib
import imaplib
from email import message_from_bytes
from email.header import decode_header
from email.utils import formatdate, make_msgid, parseaddr
import pickle
import json
from typing import List, Optional
import logging
from logging.handlers import RotatingFileHandler
import subprocess

# Google API
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Importar lista de contactos
from lista_contactos_reportes import (
    obtener_contacto_por_email,
    obtener_contacto,
    obtener_contactos_por_empresa,
    esta_autorizado,
    obtener_configuracion_autorizado
)

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Configuración de logging
LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / "monitoreo_correos.log"

# Archivo para guardar contexto de solicitudes pendientes
CONTEXTO_PENDIENTES_FILE = LOG_DIR / "solicitudes_pendientes.json"

# Configurar logger que escribe tanto en consola como en archivo
logger = logging.getLogger("monitoreo_correos")
logger.setLevel(logging.INFO)

# Evitar duplicar handlers si ya existen
if not logger.handlers:
    # Handler para archivo con rotación (máx 10MB, mantener 5 archivos)
    file_handler = RotatingFileHandler(
        LOG_FILE,
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.INFO)
    file_formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)
    
    # Handler para consola
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(message)s')
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

# Clase para interceptar prints y escribirlos también en el log
class TeeOutput:
    """Clase que escribe tanto en consola como en archivo de log"""
    def __init__(self, *files):
        self.files = files
    
    def write(self, obj):
        for f in self.files:
            f.write(obj)
            f.flush()
    
    def flush(self):
        for f in self.files:
            f.flush()

# Interceptar stdout para escribir también en el log
# Asegurar encoding UTF-8 correcto
log_file_handle = open(LOG_FILE, 'a', encoding='utf-8', errors='replace')
original_stdout = sys.stdout
sys.stdout = TeeOutput(original_stdout, log_file_handle)

# Raíz del proyecto (credenciales locales, IMAP/SMTP)
_repo_dir = Path(__file__).resolve().parent

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
# Contraseña de aplicación Google: 1) variable WES_GMAIL_APP_PASSWORD, 2) archivo local
# gmail_oauth/app_password.txt (una línea; ignorado por git), 3) valor por defecto.
SMTP_PASSWORD = os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
if not SMTP_PASSWORD:
    _app_pw_file = _repo_dir / "gmail_oauth" / "app_password.txt"
    if _app_pw_file.is_file():
        try:
            for _line in _app_pw_file.read_text(encoding="utf-8", errors="replace").splitlines():
                _line = _line.strip()
                if _line and not _line.startswith("#"):
                    SMTP_PASSWORD = _line
                    break
        except OSError:
            pass
if not SMTP_PASSWORD:
    SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
IMAP_SERVIDOR = "imap.gmail.com"
IMAP_PUERTO = 993

# Pausa automática si Excel está abierto (reduce conflictos al guardar Excel y evita picos de RAM).
# Puedes desactivarlo con WES_PAUSE_ON_EXCEL=0
PAUSA_SI_EXCEL_ABIERTO = os.environ.get("WES_PAUSE_ON_EXCEL", "1").strip() != "0"
TIEMPO_POLL_EXCEL_SEG = int(os.environ.get("WES_EXCEL_PAUSE_SECONDS", "15"))

def excel_abierto() -> bool:
    """
    Detecta si Microsoft Excel está en ejecución.
    Se usa para pausar el monitor mientras el usuario edita el Excel.
    """
    if not PAUSA_SI_EXCEL_ABIERTO:
        return False
    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", "IMAGENAME eq excel.exe"],
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        return "excel.exe" in out.lower()
    except Exception:
        # Si falla la detección, no pausamos por seguridad operativa.
        return False

# Configuración de Google API (Gmail)
# Nota: antes se usaba una ruta fija (C:\Users\joseo\...), que en esta máquina puede no existir.
# Priorizamos credenciales locales dentro del repo para que el monitor sea portable.
_local_credentials_dir = _repo_dir / "gmail_oauth"
_legacy_credentials_dir = Path(r"C:\Users\joseo\Desktop\WES\2026\Agente Derco")

_candidate_dirs = [
    _local_credentials_dir,
    _legacy_credentials_dir,
]

CREDENTIALS_DIR = None
for _d in _candidate_dirs:
    if (_d / "credentials_drive.json").exists():
        CREDENTIALS_DIR = _d
        break

# Si ninguna existe todavía, elegimos por defecto la local (para que el usuario pueda copiar ahí las credenciales).
if CREDENTIALS_DIR is None:
    CREDENTIALS_DIR = _local_credentials_dir

CREDENTIALS_FILE = CREDENTIALS_DIR / "credentials_drive.json"
TOKEN_GMAIL_FILE = CREDENTIALS_DIR / "token_gmail.pickle"

# Scopes necesarios (Drive + Gmail)
SCOPES = [
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.modify'  # Para marcar como leído
]

# Palabras clave para detectar solicitudes de reportes (más específicas)
PALABRAS_CLAVE_REPORTE = [
    "generar reporte",
    "necesito reporte",
    "necesito un reporte",
    "necesito el reporte",
    "envía reporte",
    "envíame reporte",
    "quiero reporte",
    "reporte de bupa",
    "reporte de parque arauco",
    "reporte de consumo",
    "reporte de monitoreo",
    "reporte detallado",
    "reporte de derco",
    "análisis de consumo",
    "datos de consumo",
    "reporte wes",
    "reporte ia wes",
    # Variaciones con "informe" (equivalente a "reporte")
    "generar informe",
    "necesito informe",
    "necesito un informe",
    "necesito el informe",
    "envía informe",
    "envíame informe",
    "quiero informe",
    "informe de bupa",
    "informe de parque arauco",
    "informe de consumo",
    "informe de monitoreo",
    "informe detallado",
    "informe de derco",
    "informe wes",
    "informe ia wes",
    # Control nocturno (Renca / Excel de horarios)
    "reporte de control nocturno",
    "reporte control nocturno",
    "informe de control nocturno",
    "informe control nocturno",
    "reporte control nocturno renca",
    "informe control nocturno renca",
]

# Palabras individuales que indican solicitud (búsqueda más flexible)
PALABRAS_SOLICITUD = ["necesito", "quiero", "solicito", "requiero", "envía", "envíame", "genera", "generar"]
PALABRAS_REPORTE = ["reporte", "informe", "análisis", "datos", "estadísticas"]

# Palabras que indican que NO es una solicitud (filtros)
PALABRAS_EXCLUIR = [
    "secret santa",
    "descuento",
    "oferta",
    "promoción",
    "pago",
    "cuenta",
    "factura",
    "newsletter",
    "suscripción"
]

# Empresas conocidas
EMPRESAS_CONOCIDAS = ["BUPA", "Parque Arauco", "Fundo Zapallar", "COPEC", "AGUNSA", "DERCO", "Derco"]

# Estructura de malls para Parque Arauco
# Mapea cada mall a sus nodos correspondientes
PARQUE_ARAUCO_MALLS = {
    "estacion": {
        "nombres": ["estacion", "estación", "parque arauco estacion", "parque arauco estación"],
        # Solo incluir los nodos que aparecen en la imagen compartida
        "nodos": [
            "000025-01",  # PAE Estanque Norte Locales
            "000025-19",  # MAE Sala de Bomba Estanque Sur
            "000025-04",  # PAE Baños Públicos
            "000025-07"   # PIZZA HUT
        ]
    },
    "maipu": {
        "nombres": ["maipu", "maipú", "parque arauco maipu", "parque arauco maipú", "mall maipu", "mall maipú"],
        "nodos": ["000025-08", "000025-09", "000025-10"]
    },
    "el bosque": {
        "nombres": ["el bosque", "parque arauco el bosque", "mall el bosque"],
        "nodos": ["000025-11", "000025-12"]
    },
    "quilicura": {
        "nombres": ["quilicura", "parque arauco quilicura", "mall quilicura"],
        "nodos": ["000025-13", "000025-14"]
    },
    "curauma": {
        "nombres": ["curauma", "parque arauco curauma", "mall curauma"],
        "nodos": ["000025-15", "000025-16"]
    },
    "buenaventura": {
        "nombres": ["buenaventura", "parque arauco buenaventura", "mall buenaventura"],
        "nodos": ["000025-17", "000025-18"]
    },
    "kennedy": {
        "nombres": ["kennedy", "parque arauco kennedy", "mall kennedy"],
        "nodos": ["000025-20", "000025-21", "000025-22", "000025-23", "000025-24", "000025-35", "000025-36", "000025-27", "000025-28", "000025-29"]
    }
}


def verificar_conectividad():
    """Verifica si hay conectividad a internet."""
    import socket
    try:
        # Intentar resolver DNS de Google
        socket.gethostbyname('oauth2.googleapis.com')
        return True
    except socket.gaierror:
        return False

def obtener_servicio_gmail():
    """Obtiene el servicio de Gmail autenticado."""
    creds = None
    
    # Cargar token si existe
    if os.path.exists(TOKEN_GMAIL_FILE):
        try:
            with open(TOKEN_GMAIL_FILE, 'rb') as token:
                creds = pickle.load(token)
        except Exception as e:
            print(f"  [ADVERTENCIA] Error al cargar token: {e}")
            creds = None
    
    # Si no hay credenciales válidas, solicitar autorización
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            # Verificar conectividad antes de intentar refrescar
            if not verificar_conectividad():
                print("  [ERROR] No hay conexión a internet")
                print("  [ERROR] No se puede refrescar el token sin conexión")
                print("  [INFO] Verifique su conexión a internet e intente nuevamente")
                print("  [INFO] El script reintentará en el próximo ciclo")
                # Retornar None para que el ciclo actual se salte pero continúe el siguiente
                return None
            
            try:
                print("  [INFO] Refrescando token de autenticación...")
                # Intentar refrescar el token con reintentos
                import time
                max_reintentos = 3
                reintento = 0
                token_refrescado = False
                
                while reintento < max_reintentos and not token_refrescado:
                    try:
                        creds.refresh(Request())
                        print("  [OK] Token refrescado exitosamente")
                        token_refrescado = True
                    except Exception as e_refresh:
                        reintento += 1
                        error_msg = str(e_refresh)
                        if reintento < max_reintentos:
                            print(f"  [ADVERTENCIA] Intento {reintento}/{max_reintentos} falló al refrescar token")
                            if "SSLError" in error_msg or "SSL" in error_msg:
                                print(f"  [INFO] Error SSL detectado, reintentando en 5 segundos...")
                            else:
                                print(f"  [INFO] Reintentando en 3 segundos...")
                            time.sleep(5 if "SSL" in error_msg else 3)
                        else:
                            # Último intento falló, lanzar excepción para manejo externo
                            raise e_refresh
                            
            except Exception as e:
                error_msg = str(e)
                if "getaddrinfo failed" in error_msg or "NameResolutionError" in error_msg:
                    print("  [ERROR] Error de resolución DNS - No se puede conectar a los servidores de Google")
                    print("  [ERROR] Verifique su conexión a internet y configuración de DNS")
                    print(f"  [DETALLE] {error_msg}")
                    print("  [INFO] El script reintentará en el próximo ciclo")
                    # Retornar None en lugar de salir para permitir que continúe
                    return None
                elif "Max retries exceeded" in error_msg or "SSLError" in error_msg or "SSL" in error_msg:
                    print("  [ERROR] No se puede conectar a los servidores de Google")
                    print("  [ERROR] Verifique su conexión a internet")
                    if "SSLError" in error_msg or "SSL" in error_msg:
                        print("  [INFO] Error SSL detectado - puede ser un problema temporal de red o SSL")
                        print("  [INFO] El script reintentará en el próximo ciclo")
                    else:
                        print("  [INFO] El script reintentará en el próximo ciclo")
                    # Guardar las credenciales existentes aunque estén expiradas para intentar refrescar después
                    # No salir del script, permitir que continúe con el ciclo siguiente
                    # Retornar None para que el ciclo actual se salte pero continúe el siguiente
                    if creds:
                        try:
                            # Guardar credenciales para intentar refrescar en el siguiente ciclo
                            with open(TOKEN_GMAIL_FILE, 'wb') as token:
                                pickle.dump(creds, token)
                        except:
                            pass
                    return None
                else:
                    print(f"  [ADVERTENCIA] Error al refrescar token: {e}")
                    print("  [INFO] El script reintentará en el próximo ciclo")
                    # Retornar None en lugar de establecer creds = None y continuar
                    return None
        
        if not creds or not creds.valid:
            # Verificar conectividad antes de iniciar OAuth2
            if not verificar_conectividad():
                print("  [ERROR] No hay conexión a internet")
                print("  [ERROR] No se puede iniciar autenticación OAuth2 sin conexión")
                print("  [INFO] Verifique su conexión a internet e intente nuevamente")
                print("  [INFO] El script reintentará en el próximo ciclo")
                # Retornar None en lugar de salir para permitir que continúe
                return None
            
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"[ERROR] No se encontró el archivo de credenciales:")
                print(f"  {CREDENTIALS_FILE}")
                print("  [INFO] El script no puede continuar sin credenciales")
                # Este es un error crítico, pero aún así retornamos None para que el ciclo continúe
                # y se pueda intentar de nuevo en el siguiente ciclo
                return None
            
            print("  [INFO] Iniciando flujo de autenticación OAuth2...")
            print("  [INFO] Se solicitarán permisos para leer y enviar correos")
            print("  [INFO] Se abrirá una ventana del navegador o se mostrará un código para autorizar")
            try:
                flow = InstalledAppFlow.from_client_secrets_file(
                    str(CREDENTIALS_FILE), SCOPES)
                # Intentar usar run_local_server primero, si falla usar run_console
                try:
                    creds = flow.run_local_server(port=0)
                except AttributeError:
                    # Fallback para versiones más antiguas
                    creds = flow.run_console()
            except Exception as e:
                error_msg = str(e)
                if "getaddrinfo failed" in error_msg or "NameResolutionError" in error_msg:
                    print("  [ERROR] Error de resolución DNS - No se puede conectar a los servidores de Google")
                    print("  [ERROR] Verifique su conexión a internet y configuración de DNS")
                    print(f"  [DETALLE] {error_msg}")
                    print("  [INFO] El script reintentará en el próximo ciclo")
                    return None
                else:
                    print(f"  [ERROR] Error al iniciar autenticación OAuth2: {e}")
                    print("  [INFO] El script reintentará en el próximo ciclo")
                    return None
        
        # Guardar credenciales para la próxima vez
        try:
            with open(TOKEN_GMAIL_FILE, 'wb') as token:
                pickle.dump(creds, token)
            print("  [OK] Token guardado exitosamente")
        except Exception as e:
            print(f"  [ADVERTENCIA] No se pudo guardar el token: {e}")
    
    return build('gmail', 'v1', credentials=creds)


def obtener_correos_no_leidos(service, max_results=10):
    """
    Obtiene correos no leídos de la bandeja de entrada, ordenados por fecha (más recientes primero).
    
    Args:
        service: Servicio de Gmail autenticado
        max_results: Número máximo de correos a obtener
    
    Returns:
        Lista de mensajes ordenados por fecha (más recientes primero)
    """
    try:
        # Buscar correos no leídos, ordenados por fecha interna (más recientes primero)
        results = service.users().messages().list(
            userId='me',
            q='is:unread',
            maxResults=max_results
        ).execute()
        
        messages = results.get('messages', [])
        
        if not messages:
            return []
        
        # Obtener información de fecha interna para cada mensaje y ordenar
        messages_with_dates = []
        for msg in messages:
            try:
                # Obtener solo metadata para la fecha interna (más eficiente)
                msg_data = service.users().messages().get(
                    userId='me',
                    id=msg['id'],
                    format='metadata',
                    metadataHeaders=['Date']
                ).execute()
                
                # Usar internalDate que es más confiable y rápido
                internal_date = msg_data.get('internalDate')
                if internal_date:
                    from datetime import datetime, timedelta
                    date_obj = datetime.fromtimestamp(int(internal_date) / 1000)
                    messages_with_dates.append((date_obj, msg))
                else:
                    # Si no hay internalDate, usar fecha actual como fallback
                    from datetime import datetime, timedelta
                    messages_with_dates.append((datetime.now(), msg))
            except Exception as e:
                # Si hay error obteniendo la fecha, usar fecha actual
                from datetime import datetime, timedelta
                messages_with_dates.append((datetime.now(), msg))
        
        # Ordenar por fecha (más recientes primero)
        messages_with_dates.sort(key=lambda x: x[0], reverse=True)
        
        # Retornar solo los mensajes (sin las fechas)
        sorted_messages = [msg for _, msg in messages_with_dates]
        print(f"[INFO] Correos ordenados por fecha (más recientes primero): {len(sorted_messages)} correo(s)")
        return sorted_messages
        
    except Exception as e:
        print(f"[ERROR] Error al obtener correos: {e}")
        return []


def obtener_cuerpo_correo(service, msg_id):
    """
    Obtiene el cuerpo completo de un correo.
    
    Args:
        service: Servicio de Gmail autenticado
        msg_id: ID del mensaje
    
    Returns:
        Tupla (asunto, cuerpo, remitente_email, remitente_nombre, message_id, in_reply_to, references)
    """
    try:
        message = service.users().messages().get(
            userId='me',
            id=msg_id,
            format='full'
        ).execute()
        
        # Extraer headers
        headers = message['payload'].get('headers', [])
        asunto = next((h['value'] for h in headers if h['name'] == 'Subject'), 'Sin asunto')
        remitente = next((h['value'] for h in headers if h['name'] == 'From'), '')
        message_id_header = next((h['value'] for h in headers if h['name'] == 'Message-ID'), None)
        in_reply_to = next((h['value'] for h in headers if h['name'] == 'In-Reply-To'), None)
        references = next((h['value'] for h in headers if h['name'] == 'References'), None)
        
        # Parsear remitente
        remitente_email = remitente
        remitente_nombre = ''
        if '<' in remitente:
            match = re.match(r'(.+?)\s*<(.+?)>', remitente)
            if match:
                remitente_nombre = match.group(1).strip().strip('"')
                remitente_email = match.group(2).strip()
        
        # Extraer cuerpo del mensaje
        cuerpo = ''
        payload = message['payload']
        
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/plain':
                    data = part['body'].get('data')
                    if data:
                        cuerpo = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
                        break
                elif part['mimeType'] == 'text/html' and not cuerpo:
                    data = part['body'].get('data')
                    if data:
                        # Extraer texto del HTML (simple)
                        html = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
                        # Remover tags HTML básicos
                        cuerpo = re.sub(r'<[^>]+>', '', html)
                        break
        else:
            # Mensaje simple sin partes
            if payload['mimeType'] == 'text/plain':
                data = payload['body'].get('data')
                if data:
                    cuerpo = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
        
        return asunto, cuerpo, remitente_email, remitente_nombre, message_id_header, in_reply_to, references
    
    except Exception as e:
        print(f"[ERROR] Error al obtener cuerpo del correo {msg_id}: {e}")
        return None, None, None, None, None, None, None


def marcar_como_leido(service, msg_id):
    """Marca un correo como leído."""
    try:
        service.users().messages().modify(
            userId='me',
            id=msg_id,
            body={'removeLabelIds': ['UNREAD']}
        ).execute()
        return True
    except Exception as e:
        print(f"[ERROR] Error al marcar correo como leído: {e}")
        return False


def _smtp_password_normalizado() -> str:
    """Quita espacios típicos de la contraseña de aplicación de Google."""
    return (SMTP_PASSWORD or "").replace(" ", "").strip()


def _decode_mime_header(value: Optional[str]) -> str:
    if not value:
        return ""
    parts = decode_header(value)
    out: List[str] = []
    for text, enc in parts:
        if isinstance(text, bytes):
            out.append(text.decode(enc or "utf-8", errors="replace"))
        else:
            out.append(str(text))
    return "".join(out)


def _extraer_cuerpo_desde_email_message(msg) -> str:
    cuerpo = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            if ctype == "text/plain":
                payload = part.get_payload(decode=True)
                if payload:
                    cuerpo = payload.decode("utf-8", errors="replace")
                    break
            if ctype == "text/html" and not cuerpo:
                payload = part.get_payload(decode=True)
                if payload:
                    html = payload.decode("utf-8", errors="replace")
                    cuerpo = re.sub(r"<[^>]+>", "", html)
    else:
        if msg.get_content_type() == "text/plain":
            payload = msg.get_payload(decode=True)
            if payload:
                cuerpo = payload.decode("utf-8", errors="replace")
    return cuerpo


def _parsear_correo_rfc822(raw_bytes: bytes):
    """
    Devuelve la misma tupla que obtener_cuerpo_correo (Gmail API).
    """
    msg = message_from_bytes(raw_bytes)
    asunto = _decode_mime_header(msg.get("Subject"))
    remitente_raw = msg.get("From", "")
    remitente_nombre, remitente_email = parseaddr(remitente_raw)
    remitente_nombre = remitente_nombre.strip().strip('"')
    remitente_email = (remitente_email or remitente_raw).strip()
    message_id_header = msg.get("Message-ID")
    in_reply_to = msg.get("In-Reply-To")
    references = msg.get("References")
    cuerpo = _extraer_cuerpo_desde_email_message(msg)
    return asunto, cuerpo, remitente_email, remitente_nombre, message_id_header, in_reply_to, references


def monitorear_y_procesar_correos_imap():
    """
    Lee correos no leídos (UNSEEN) por IMAP con usuario + contraseña de aplicación.
    Misma lógica de negocio que la ruta Gmail API (analizar_correo / procesar_solicitud_reporte).
    """
    pwd = _smtp_password_normalizado()
    if not pwd or not SMTP_USUARIO:
        print("[ERROR] IMAP: falta SMTP_USUARIO o contraseña de aplicación (SMTP_PASSWORD / WES_GMAIL_APP_PASSWORD).")
        return None

    print("=" * 70)
    print("MONITOREO DE CORREOS (IMAP) Y GENERACIÓN DE REPORTES")
    print("=" * 70)
    print()

    mail = None
    try:
        print("Conectando por IMAP...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVIDOR, IMAP_PUERTO)
        mail.login(SMTP_USUARIO, pwd)
        mail.select("INBOX")

        status, data = mail.uid("search", None, "UNSEEN")
        if status != "OK":
            print(f"  [ERROR] IMAP SEARCH falló: {status}")
            try:
                mail.logout()
            except Exception:
                pass
            return None

        uids = data[0].split() if data and data[0] else []
        if not uids:
            print("Buscando correos UNSEEN...")
            print("  [INFO] No hay correos no leídos (IMAP).")
            try:
                mail.logout()
            except Exception:
                pass
            print()
            return None

        uids = sorted(uids, key=lambda u: int(u), reverse=True)[:10]
        print("Buscando correos UNSEEN...")
        print(f"  [INFO] Se encontraron {len(uids)} correo(s) no leído(s) (IMAP).")
        print()

        print("Limpiando contextos caducados...")
        contextos_eliminados = limpiar_contextos_caducados(None, dias_limite=1)
        if contextos_eliminados > 0:
            print(f"  [OK] {contextos_eliminados} contexto(s) caducado(s) eliminado(s)")
        else:
            print("  [INFO] No hay contextos caducados")
        print()

        solicitudes_procesadas = 0

        def imap_marcar_leido(uid_bytes):
            try:
                mail.uid("store", uid_bytes, "+FLAGS", r"(\Seen)")
                return True
            except Exception as e:
                print(f"  [ERROR] IMAP no pudo marcar como leído: {e}")
                return False

        for i, uid in enumerate(uids, 1):
            uid_s = uid.decode() if isinstance(uid, bytes) else str(uid)
            print(f"[{i}/{len(uids)}] Procesando correo IMAP UID/ID {uid_s}...")

            typ, msg_data = mail.uid("fetch", uid, "(RFC822)")
            if typ != "OK" or not msg_data or not msg_data[0]:
                print("  [ERROR] No se pudo obtener RFC822")
                continue

            raw = msg_data[0]
            if isinstance(raw, tuple):
                raw = raw[1]
            if not isinstance(raw, (bytes, bytearray)):
                print("  [ERROR] Formato de mensaje IMAP inesperado")
                continue

            asunto, cuerpo, remitente_email, remitente_nombre, message_id_header, in_reply_to, references = _parsear_correo_rfc822(
                bytes(raw)
            )

            if not asunto and not cuerpo:
                print("  [ERROR] No se pudo obtener el contenido del correo")
                continue

            print(f"  Asunto: {asunto}")
            print(f"  Remitente: {remitente_email}")

            config_autorizado = esta_autorizado(remitente_email)
            if not config_autorizado:
                print("  [INFO] Remitente no autorizado (IMAP), se marca como leído.")
                imap_marcar_leido(uid)
                print()
                continue

            info_detectada = analizar_correo(asunto, cuerpo, remitente_email)

            if info_detectada and info_detectada.get("es_solicitud"):
                print("  [OK] Solicitud de reporte detectada (IMAP).")
                info_detectada["message_id_original"] = message_id_header
                info_detectada["gmail_msg_id"] = uid_s
                info_detectada["in_reply_to"] = in_reply_to
                info_detectada["references"] = references

                if procesar_solicitud_reporte(info_detectada):
                    print("  [OK] Reporte procesado y correo enviado (IMAP)")
                    solicitudes_procesadas += 1
                    imap_marcar_leido(uid)
                    print("  [OK] Correo marcado como leído (IMAP)")
                else:
                    print("  [ERROR] No se pudo procesar la solicitud (IMAP)")
            else:
                print("  [INFO] No es una solicitud de reporte (IMAP). Se marca como leído.")
                imap_marcar_leido(uid)

            print()

        timestamp_fin = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print("=" * 70)
        print("PROCESO COMPLETADO (IMAP)")
        print("=" * 70)
        print(f"  - Correos procesados: {len(uids)}")
        print(f"  - Solicitudes de reporte procesadas: {solicitudes_procesadas}")
        print(f"  - Última revisión: {timestamp_fin}")
        print()

        try:
            mail.logout()
        except Exception:
            pass
        return None

    except Exception as e:
        print(f"[ERROR] Fallback IMAP falló: {e}")
        import traceback
        traceback.print_exc()
        if mail:
            try:
                mail.logout()
            except Exception:
                pass
        return None


def guardar_contexto_solicitud(message_id_respuesta, info_detectada, parametros_faltantes):
    """
    Guarda el contexto de una solicitud pendiente para poder combinarlo con respuestas futuras.
    
    Args:
        message_id_respuesta: Message-ID del correo de respuesta que se enviará
        info_detectada: Información detectada del correo original
        parametros_faltantes: Lista de parámetros que faltan
    """
    try:
        # Cargar contexto existente
        if CONTEXTO_PENDIENTES_FILE.exists():
            with open(CONTEXTO_PENDIENTES_FILE, 'r', encoding='utf-8') as f:
                contexto_existente = json.load(f)
        else:
            contexto_existente = {}
        
        # Guardar contexto usando el Message-ID de la respuesta como clave
        contexto_existente[message_id_respuesta] = {
            'info_detectada': info_detectada,
            'parametros_faltantes': parametros_faltantes,
            'fecha': datetime.now().isoformat(),
            'remitente_email': info_detectada.get('remitente_email'),
            'gmail_msg_id': info_detectada.get('gmail_msg_id')  # ID del mensaje en Gmail API para marcarlo como leído
        }
        
        # Guardar archivo
        with open(CONTEXTO_PENDIENTES_FILE, 'w', encoding='utf-8') as f:
            json.dump(contexto_existente, f, indent=2, ensure_ascii=False)
        
        print(f"[OK] Contexto de solicitud guardado para Message-ID: {message_id_respuesta}")
        return True
    except Exception as e:
        print(f"[ERROR] Error al guardar contexto de solicitud: {e}")
        return False


def obtener_contexto_solicitud(in_reply_to, references):
    """
    Obtiene el contexto de una solicitud pendiente basado en In-Reply-To o References.
    
    Args:
        in_reply_to: Header In-Reply-To del correo
        references: Header References del correo
    
    Returns:
        Diccionario con el contexto o None si no se encuentra
    """
    try:
        if not CONTEXTO_PENDIENTES_FILE.exists():
            return None
        
        with open(CONTEXTO_PENDIENTES_FILE, 'r', encoding='utf-8') as f:
            contexto_existente = json.load(f)
        
        # Buscar por In-Reply-To primero
        if in_reply_to:
            # Limpiar el Message-ID (puede venir con < >)
            message_id_limpio = in_reply_to.strip('<>')
            if message_id_limpio in contexto_existente:
                contexto = contexto_existente[message_id_limpio]
                print(f"[OK] Contexto encontrado por In-Reply-To: {message_id_limpio}")
                return contexto
        
        # Buscar por References si In-Reply-To no funcionó
        if references:
            # References puede contener múltiples Message-IDs separados por espacios
            message_ids = references.split()
            for msg_id in message_ids:
                msg_id_limpio = msg_id.strip('<>')
                if msg_id_limpio in contexto_existente:
                    contexto = contexto_existente[msg_id_limpio]
                    print(f"[OK] Contexto encontrado por References: {msg_id_limpio}")
                    return contexto
        
        return None
    except Exception as e:
        print(f"[ERROR] Error al obtener contexto de solicitud: {e}")
        return None


def eliminar_contexto_solicitud(message_id):
    """
    Elimina el contexto de una solicitud pendiente después de procesarla.
    
    Args:
        message_id: Message-ID del correo de respuesta
    """
    try:
        if not CONTEXTO_PENDIENTES_FILE.exists():
            return
        
        with open(CONTEXTO_PENDIENTES_FILE, 'r', encoding='utf-8') as f:
            contexto_existente = json.load(f)
        
        if message_id in contexto_existente:
            del contexto_existente[message_id]
            
            with open(CONTEXTO_PENDIENTES_FILE, 'w', encoding='utf-8') as f:
                json.dump(contexto_existente, f, indent=2, ensure_ascii=False)
            
            print(f"[OK] Contexto eliminado para Message-ID: {message_id}")
    except Exception as e:
        print(f"[ERROR] Error al eliminar contexto de solicitud: {e}")


def limpiar_contextos_caducados(service, dias_limite=1):
    """
    Limpia contextos de solicitudes pendientes que han caducado (más de días_limite).
    Marca los correos originales como leídos si es posible.
    
    Args:
        service: Servicio de Gmail API
        dias_limite: Número de días límite para considerar un contexto como caducado (default: 1)
    
    Returns:
        Número de contextos eliminados
    """
    try:
        if not CONTEXTO_PENDIENTES_FILE.exists():
            return 0
        
        with open(CONTEXTO_PENDIENTES_FILE, 'r', encoding='utf-8') as f:
            contexto_existente = json.load(f)
        
        if not contexto_existente:
            return 0
        
        contextos_caducados = []
        fecha_limite = datetime.now() - timedelta(days=dias_limite)
        
        for message_id_respuesta, contexto in contexto_existente.items():
            fecha_str = contexto.get('fecha')
            if not fecha_str:
                # Si no tiene fecha, considerar caducado
                contextos_caducados.append(message_id_respuesta)
                continue
            
            try:
                fecha_contexto = datetime.fromisoformat(fecha_str)
                if fecha_contexto < fecha_limite:
                    contextos_caducados.append(message_id_respuesta)
            except (ValueError, TypeError):
                # Si hay error al parsear la fecha, considerar caducado
                contextos_caducados.append(message_id_respuesta)
        
        if not contextos_caducados:
            return 0
        
        print(f"[INFO] Encontrados {len(contextos_caducados)} contexto(s) caducado(s) (más de {dias_limite} día(s))")
        
        # Intentar marcar correos originales como leídos
        for message_id_respuesta in contextos_caducados:
            contexto = contexto_existente[message_id_respuesta]
            gmail_msg_id = contexto.get('gmail_msg_id')
            
            if gmail_msg_id and service:
                try:
                    # Marcar el correo original como leído usando el ID de Gmail guardado
                    if marcar_como_leido(service, gmail_msg_id):
                        print(f"[OK] Correo original marcado como leído (Gmail ID: {gmail_msg_id})")
                    else:
                        print(f"[ADVERTENCIA] No se pudo marcar el correo como leído (Gmail ID: {gmail_msg_id})")
                except Exception as e:
                    print(f"[ADVERTENCIA] Error al marcar el correo original como leído: {e}")
            
            # Eliminar el contexto caducado
            del contexto_existente[message_id_respuesta]
            print(f"[OK] Contexto caducado eliminado (Message-ID respuesta: {message_id_respuesta})")
        
        # Guardar el archivo actualizado
        with open(CONTEXTO_PENDIENTES_FILE, 'w', encoding='utf-8') as f:
            json.dump(contexto_existente, f, indent=2, ensure_ascii=False)
        
        return len(contextos_caducados)
        
    except Exception as e:
        print(f"[ERROR] Error al limpiar contextos caducados: {e}")
        return 0


def analizar_correo(asunto, cuerpo, remitente_email):
    """
    Analiza un correo para detectar si es una solicitud de reporte.
    Detecta los parámetros (empresa, nodo, periodo) de forma independiente del orden.
    
    Args:
        asunto: Asunto del correo
        cuerpo: Cuerpo del correo
        remitente_email: Email del remitente
    
    Returns:
        Diccionario con información detectada o None si no es una solicitud
    """
    texto_completo = f"{asunto} {cuerpo}".lower()
    asunto_lower = asunto.lower()
    
    # REQUISITO OBLIGATORIO: El asunto debe contener "reporte", "reportes", "informe" o "informes"
    # Si no contiene estas palabras en el asunto, no procesar (incluso si el remitente está autorizado)
    palabras_obligatorias_asunto = ["reporte", "reportes", "informe", "informes"]
    tiene_palabra_obligatoria_asunto = any(palabra in asunto_lower for palabra in palabras_obligatorias_asunto)
    
    if not tiene_palabra_obligatoria_asunto:
        print(f"[INFO] El asunto no contiene 'reporte', 'reportes', 'informe' o 'informes'. No se procesará el correo.")
        return None
    
    # Excluir correos promocionales/spam
    if any(palabra in texto_completo for palabra in PALABRAS_EXCLUIR):
        return None
    
    # Verificar si contiene palabras clave de reporte (búsqueda exacta)
    es_solicitud = any(palabra in texto_completo for palabra in PALABRAS_CLAVE_REPORTE)
    
    # Si no se encontró con búsqueda exacta, intentar búsqueda flexible
    # (buscar palabras de solicitud + palabras de reporte)
    if not es_solicitud:
        tiene_solicitud = any(palabra in texto_completo for palabra in PALABRAS_SOLICITUD)
        tiene_reporte = any(palabra in texto_completo for palabra in PALABRAS_REPORTE)
        es_solicitud = tiene_solicitud and tiene_reporte
    
    # Pedido explícito de reporte de control nocturno: el asunto ya exige "reporte"/"informe"
    if not es_solicitud and (
        "control nocturno" in texto_completo
        or ("renca" in texto_completo and "nocturn" in texto_completo)
    ):
        es_solicitud = True
    
    if not es_solicitud:
        return None
    
    # Obtener información del remitente
    contacto = obtener_contacto_por_email(remitente_email)
    
    # Verificar si el remitente está autorizado para todas las empresas
    config_autorizado = esta_autorizado(remitente_email)
    puede_solicitar_cualquier_empresa = False
    if config_autorizado:
        puntos_monitoreo = config_autorizado.get("puntos_monitoreo", [])
        puede_solicitar_cualquier_empresa = "Todas" in puntos_monitoreo or (isinstance(puntos_monitoreo, list) and "Todas" in [p.upper() for p in puntos_monitoreo])
    
    # ===================================================================
    # DETECCIÓN INDEPENDIENTE DEL ORDEN DE PARÁMETROS
    # Los tres parámetros (empresa, nodo, periodo) se detectan de forma
    # independiente analizando TODO el texto completo, sin importar el orden
    # ===================================================================
    
    # 0. DETECTAR MALL DE PARQUE ARAUCO PRIMERO (antes de detectar empresa)
    # Si se detecta un mall, automáticamente se asigna "Parque Arauco" como empresa
    mall_detectado = detectar_mall_parque_arauco(texto_completo)
    empresa_detectada = None
    
    if mall_detectado:
        # Si se detecta un mall de Parque Arauco, automáticamente asignar la empresa
        empresa_detectada = "Parque Arauco"
        print(f"[INFO] Mall de Parque Arauco detectado: {mall_detectado.get('nombre_completo')}")
        print(f"[INFO] Empresa asignada automáticamente: Parque Arauco (por detección de mall)")
    
    # 1. DETECTAR EMPRESA (analiza todo el texto completo, independiente del orden)
    # Solo buscar empresa si no se detectó un mall (que ya asigna Parque Arauco)
    if not empresa_detectada:
        # Obtener todas las empresas disponibles desde la API (siempre, para búsqueda flexible)
        empresas_dict = obtener_todas_las_empresas()
        
        # Primero buscar en la lista conocida (búsqueda exacta)
        for empresa in EMPRESAS_CONOCIDAS:
            if empresa.lower() in texto_completo:
                empresa_detectada = empresa
                print(f"[INFO] Empresa detectada desde lista conocida: {empresa}")
                break
    
    # Si no se detectó, buscar en todas las empresas disponibles desde la API
    # Buscar coincidencias parciales y flexibles (sin necesidad de mencionar "empresa")
    if not empresa_detectada:
        texto_upper = texto_completo.upper()
        texto_normalizado = texto_completo.replace("empresa", "").replace("empresas", "").strip()
        
        # Buscar coincidencias exactas primero
        for nombre_empresa_api, empresa_id in empresas_dict.items():
            nombre_empresa_lower = nombre_empresa_api.lower()
            
            # Coincidencia exacta del nombre completo
            if nombre_empresa_lower in texto_completo or nombre_empresa_lower in texto_normalizado:
                empresa_detectada = nombre_empresa_api
                print(f"[INFO] Empresa detectada desde API (coincidencia exacta): {nombre_empresa_api}")
                break
            
            # Coincidencia parcial: si el nombre de la empresa contiene palabras del texto o viceversa
            palabras_empresa = set(nombre_empresa_lower.split())
            palabras_texto = set(texto_normalizado.split())
            
            # Si hay al menos 2 palabras en común (y no son palabras comunes)
            palabras_comunes = palabras_empresa.intersection(palabras_texto)
            palabras_comunes = {p for p in palabras_comunes if len(p) > 3 and p not in ['del', 'de', 'la', 'el', 'los', 'las', 'y', 'en', 'con', 'para', 'por', 'los', 'las']}
            
            if len(palabras_comunes) >= 2:
                empresa_detectada = nombre_empresa_api
                print(f"[INFO] Empresa detectada desde API (coincidencia parcial, {len(palabras_comunes)} palabras): {nombre_empresa_api}")
                break
            
            # Búsqueda por palabras clave significativas (ej: "cormup" en "colegios cormup")
            # Extraer palabras significativas del nombre de la empresa (más de 4 caracteres)
            palabras_significativas_empresa = [p for p in nombre_empresa_lower.split() if len(p) > 4]
            for palabra_empresa in palabras_significativas_empresa:
                if palabra_empresa in texto_normalizado:
                    empresa_detectada = nombre_empresa_api
                    print(f"[INFO] Empresa detectada desde API (palabra clave '{palabra_empresa}'): {nombre_empresa_api}")
                    break
            if empresa_detectada:
                break
    
    # Si aún no se detectó y el remitente puede solicitar cualquier empresa,
    # hacer una búsqueda más agresiva
    if not empresa_detectada and puede_solicitar_cualquier_empresa:
        # Buscar cualquier palabra significativa que pueda ser parte de un nombre de empresa
        palabras_texto_significativas = [p for p in texto_normalizado.split() if len(p) > 4 and p not in ['reporte', 'informe', 'necesito', 'quiero', 'solicito', 'desde', 'hasta', 'periodo']]
        for palabra in palabras_texto_significativas:
            for nombre_empresa_api, empresa_id in empresas_dict.items():
                if palabra in nombre_empresa_api.lower():
                    empresa_detectada = nombre_empresa_api
                    print(f"[INFO] Empresa detectada desde API (búsqueda agresiva, palabra '{palabra}'): {nombre_empresa_api}")
                    break
            if empresa_detectada:
                break
    
    # 2. DETECTAR PERIODO (analiza todo el texto completo, independiente del orden)
    # La función detectar_periodo() ya analiza todo el texto, por lo que funciona
    # independientemente de dónde esté el periodo en el correo
    # IMPORTANTE: Usar el texto ORIGINAL (asunto + cuerpo) sin normalizar para mantener formato de fechas
    texto_para_periodo = f"{asunto} {cuerpo}" if asunto and cuerpo else texto_completo
    print(f"[DEBUG] Detectando período en texto original (primeros 200 caracteres): {texto_para_periodo[:200]}...")
    periodo = detectar_periodo(texto_para_periodo)
    if periodo:
        print(f"[INFO] Período detectado correctamente: {periodo}")
    else:
        print(f"[ADVERTENCIA] No se detectó período en el correo")
    
    # 3. DETECTAR TIPO DE REPORTE (analiza todo el texto completo)
    tipo_reporte = detectar_tipo_reporte(texto_completo)
    
    # NOTA: Los nodos específicos se detectan más tarde en procesar_solicitud_reporte()
    # cuando ya se tiene la empresa, pero la detección también analiza todo el texto
    # del correo original, por lo que funciona independientemente del orden
    
    # Detectar formato solicitado (Word o PDF)
    # Buscar en el cuerpo del correo si se menciona Word, docx, formato word, etc.
    cuerpo_lower = cuerpo.lower()
    palabras_word = ["word", "docx", "formato word", "en word", "archivo word", "documento word", ".docx", ".doc"]
    solicita_word = any(palabra in cuerpo_lower for palabra in palabras_word)
    
    formato_solicitado = "word" if solicita_word else "pdf"
    if solicita_word:
        print(f"[INFO] Se detectó solicitud de formato Word en el correo")
    
    # Detectar solicitud de presentación (PPT/PDF)
    # Buscar palabras clave como "presentacion", "ppt", "powerpoint", "slides", etc.
    asunto_lower = asunto.lower() if asunto else ""
    palabras_presentacion = [
        "presentacion", "presentación", "ppt", "powerpoint", "slides", 
        "diapositivas", "presentacion ppt", "presentación ppt",
        "presentacion powerpoint", "presentación powerpoint",
        "con presentacion", "con presentación", "incluir presentacion", "incluir presentación"
    ]
    solicita_presentacion = any(palabra in cuerpo_lower or palabra in asunto_lower for palabra in palabras_presentacion)
    
    # Detectar formato de presentación solicitado (PPT o PDF)
    # Por defecto, intentar PDF si se puede convertir, sino PPT
    formato_presentacion = None  # None = no se solicita, "ppt" o "pdf"
    if solicita_presentacion:
        # Buscar preferencia específica por PPT o PDF
        palabras_ppt = ["ppt", "powerpoint", ".pptx", ".ppt", "en ppt", "formato ppt"]
        palabras_pdf_presentacion = ["presentacion pdf", "presentación pdf", "pdf de la presentacion", "pdf de la presentación"]
        
        prefiere_ppt = any(palabra in cuerpo_lower or palabra in asunto_lower for palabra in palabras_ppt)
        prefiere_pdf = any(palabra in cuerpo_lower or palabra in asunto_lower for palabra in palabras_pdf_presentacion)
        
        if prefiere_ppt:
            formato_presentacion = "ppt"
            print(f"[INFO] Se detectó solicitud de presentación en formato PPT")
        elif prefiere_pdf:
            formato_presentacion = "pdf"
            print(f"[INFO] Se detectó solicitud de presentación en formato PDF")
        else:
            # Por defecto, intentar PDF (más compatible), si no se puede, usar PPT
            formato_presentacion = "pdf"
            print(f"[INFO] Se detectó solicitud de presentación (formato por defecto: PDF)")
    
    # Reporte operativo control nocturno (Renca): no usa empresa/nodos del flujo estándar
    es_control_nocturno = (
        "control nocturno" in texto_completo
        or ("renca" in texto_completo and "nocturn" in texto_completo)
    )
    
    # Preparar información detectada
    info_detectada = {
        "es_solicitud": True,
        "remitente": contacto,
        "remitente_email": remitente_email,
        "empresa": empresa_detectada,
        "periodo": periodo,
        "tipo_reporte": tipo_reporte,  # "individual", "agregado", "ambos"
        "formato_solicitado": formato_solicitado,  # "word" o "pdf"
        "solicita_presentacion": solicita_presentacion,  # True o False
        "formato_presentacion": formato_presentacion,  # "ppt", "pdf" o None
        "asunto_original": asunto,
        "cuerpo_original": cuerpo
    }
    
    if es_control_nocturno:
        info_detectada["reporte_especial"] = "control_nocturno"
    
    # Si se detectó un mall, guardarlo en info_detectada
    if mall_detectado:
        info_detectada['mall_detectado'] = mall_detectado
        print(f"[INFO] Mall guardado en información detectada: {mall_detectado.get('nombre_completo')}")
    
    return info_detectada


def detectar_periodo(texto):
    """
    Detecta fechas o periodos mencionados en el texto.
    Retorna un diccionario con 'inicio' y 'fin' si se detectan, o None si no se detecta periodo.
    Soporta múltiples formatos: DD/MM/YYYY, DD-MM-YYYY, "DD de mes YYYY", "hasta hoy", etc.
    """
    import re
    from datetime import datetime, timedelta
    
    # Normalizar texto
    texto_lower = texto.lower()
    texto_original = texto
    
    # Mapeo de meses en español
    meses_espanol = {
        'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
        'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
    }
    
    fechas_encontradas = []
    
    # Patrón 1: Fechas en formato DD/MM/YYYY, DD-MM-YYYY, DD/MM/YY o DD-MM-YY
    # También captura formatos con "al" o "hasta" entre fechas
    # IMPORTANTE: Buscar en el texto original para mantener formato exacto
    patron_fecha_numerica = r"(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})"
    matches_numericas = list(re.finditer(patron_fecha_numerica, texto_original))
    
    print(f"[DEBUG] Deteccion de periodo: Se encontraron {len(matches_numericas)} fecha(s) numerica(s)")
    
    # Si hay múltiples fechas y palabras como "al", "hasta", "a" entre ellas, procesarlas como rango
    if len(matches_numericas) >= 2:
        # Buscar palabras que indiquen rango entre las fechas (usar texto en minúsculas para búsqueda)
        texto_entre_fechas = texto_lower[matches_numericas[0].end():matches_numericas[1].start()]
        print(f"[DEBUG] Texto entre fechas: '{texto_entre_fechas}'")
        palabras_rango = ['al', 'hasta', ' a ', 'del ', 'al ', 'hasta el', 'hasta la']
        # Limpiar espacios al inicio/fin para la búsqueda
        texto_entre_limpio = texto_entre_fechas.strip()
        tiene_rango = any(palabra in texto_entre_limpio or palabra.strip() in texto_entre_limpio for palabra in palabras_rango)
        print(f"[DEBUG] Tiene palabra de rango: {tiene_rango}")
        if tiene_rango:
            # Procesar ambas fechas como inicio y fin
            for match in matches_numericas[:2]:
                dia, mes, anio = match.groups()
                try:
                    if len(anio) == 2:
                        # Convertir año de 2 dígitos: 25 = 2025, 26 = 2026, etc.
                        anio = f"20{anio}" if int(anio) < 50 else f"19{anio}"
                    fecha_obj = datetime(int(anio), int(mes), int(dia))
                    fechas_encontradas.append((fecha_obj, match.start(), match.end()))
                except ValueError:
                    continue
            # Si se encontraron ambas fechas con "al" entre ellas, retornar el rango inmediatamente
            if len(fechas_encontradas) >= 2:
                # Ordenar fechas
                fechas_encontradas.sort(key=lambda x: x[0])
                inicio = fechas_encontradas[0][0].strftime('%d/%m/%Y')
                fin = fechas_encontradas[-1][0].strftime('%d/%m/%Y')
                print(f"[INFO] Periodo detectado (rango con 'al'): {inicio} - {fin}")
                return {'inicio': inicio, 'fin': fin}
    else:
        # Procesar fechas individuales normalmente
        for match in matches_numericas:
            dia, mes, anio = match.groups()
            try:
                if len(anio) == 2:
                    # Convertir año de 2 dígitos: 25 = 2025, 26 = 2026, etc.
                    anio = f"20{anio}" if int(anio) < 50 else f"19{anio}"
                fecha_obj = datetime(int(anio), int(mes), int(dia))
                fechas_encontradas.append((fecha_obj, match.start(), match.end()))
            except ValueError:
                continue
    
    # Patrón 2: Fechas en formato "DD de mes YYYY" o "DD de mes de YYYY" (con año)
    patron_fecha_texto_con_anio = r"(\d{1,2})\s+de\s+([a-z]+)(?:\s+de)?\s+(\d{4})"
    matches_texto_con_anio = re.finditer(patron_fecha_texto_con_anio, texto_lower)
    for match in matches_texto_con_anio:
        dia, mes_texto, anio = match.groups()
        mes_numero = meses_espanol.get(mes_texto.lower())
        if mes_numero:
            try:
                fecha_obj = datetime(int(anio), mes_numero, int(dia))
                fechas_encontradas.append((fecha_obj, match.start(), match.end()))
            except ValueError:
                continue
    
    # Patrón 3: Fechas en formato "DD de mes" (sin año, asumir año actual)
    patron_fecha_texto_sin_anio = r"(\d{1,2})\s+de\s+([a-z]+)(?!\s+de\s+\d{4})"
    matches_texto_sin_anio = re.finditer(patron_fecha_texto_sin_anio, texto_lower)
    for match in matches_texto_sin_anio:
        dia, mes_texto = match.groups()
        mes_numero = meses_espanol.get(mes_texto.lower())
        if mes_numero:
            try:
                # Asumir año actual
                anio_actual = datetime.now().year
                fecha_obj = datetime(anio_actual, mes_numero, int(dia))
                fechas_encontradas.append((fecha_obj, match.start(), match.end()))
            except ValueError:
                continue
    
    # Patrón 4: Detectar "mes YYYY" o "mes de YYYY" o "periodo mes YYYY" (ej: "diciembre 2025", "periodo diciembre 2025")
    # Si se detecta este patrón, procesarlo directamente y retornar el periodo completo del mes
    patron_mes_anio = r"(?:periodo\s+)?(?:mes\s+de\s+)?([a-z]+)\s+(?:del\s+)?(\d{4})"
    matches_mes_anio = list(re.finditer(patron_mes_anio, texto_lower))
    if matches_mes_anio:
        # Si se encontró "mes YYYY", procesarlo directamente sin mezclar con otras fechas
        match = matches_mes_anio[0]  # Tomar el primer match
        mes_texto, anio = match.groups()
        mes_numero = meses_espanol.get(mes_texto.lower())
        if mes_numero:
            try:
                # Primero del mes como inicio
                fecha_inicio = datetime(int(anio), mes_numero, 1)
                # Último día del mes como fin
                if mes_numero == 12:
                    fecha_fin = datetime(int(anio), 12, 31)
                else:
                    # Obtener el último día del mes
                    siguiente_mes = datetime(int(anio), mes_numero + 1, 1)
                    from datetime import timedelta
                    fecha_fin = siguiente_mes - timedelta(days=1)
                
                # Retornar directamente el periodo completo del mes
                print(f"[INFO] Periodo detectado: mes completo {mes_texto} {anio} (01/{mes_numero:02d}/{anio} - {fecha_fin.strftime('%d/%m/%Y')})")
                return {'inicio': fecha_inicio.strftime('%d/%m/%Y'), 'fin': fecha_fin.strftime('%d/%m/%Y')}
            except ValueError:
                pass
    
    # Patrón 5: Detectar solo el mes sin año (ej: "enero", "diciembre", "periodo enero")
    # Si solo se menciona el mes, usar año actual: del día 1 del mes hasta hoy
    # Buscar palabras de meses que no estén precedidas por un día ni seguidas por un año
    palabras_meses = list(meses_espanol.keys())
    mes_solo_detectado = None
    for mes_texto in palabras_meses:
        # Patrón para encontrar el mes solo (no precedido por número de día ni seguido por año)
        patron = r'(?:^|\s)(?:periodo\s+)?(?:mes\s+de\s+)?\b' + re.escape(mes_texto) + r'\b(?!\s+(?:del\s+)?\d{4})(?!\s+de\s+\d{4})(?!\s+\d{1,2})'
        matches = re.finditer(patron, texto_lower)
        for match in matches:
            # Verificar que no sea parte de una fecha ya detectada
            contexto_antes = texto_lower[max(0, match.start()-15):match.start()]
            contexto_despues = texto_lower[match.end():min(len(texto_lower), match.end()+15)]
            
            # Si hay un número antes (día), no es solo el mes
            tiene_dia_antes = re.search(r'\d{1,2}\s+(?:de\s+)?$', contexto_antes)
            # Si hay un año después, ya fue detectado por otro patrón
            tiene_anio_despues = re.search(r'^\s+(?:del\s+)?\d{4}', contexto_despues)
            
            if not tiene_dia_antes and not tiene_anio_despues:
                mes_numero = meses_espanol[mes_texto]
                mes_solo_detectado = mes_numero
                try:
                    anio_actual = datetime.now().year
                    # Día 1 del mes del año actual como inicio
                    fecha_inicio = datetime(anio_actual, mes_numero, 1)
                    # Fecha de hoy como fin (siempre hasta hoy cuando solo se menciona el mes)
                    fecha_fin = datetime.now()
                    
                    # Agregar ambas fechas
                    fechas_encontradas.append((fecha_inicio, match.start(), match.end()))
                    fechas_encontradas.append((fecha_fin, match.start(), match.end()))
                    break  # Solo procesar el primer mes encontrado
                except ValueError:
                    continue
        if mes_solo_detectado:
            break  # Solo procesar un mes solo
    
    # Si se detectó un mes solo (sin año), procesarlo directamente sin mezclar con otras fechas
    if mes_solo_detectado:
        anio_actual = datetime.now().year
        mes_actual = datetime.now().month
        
        # Si estamos en los primeros meses del año (enero, febrero, marzo) y se menciona
        # un mes del año anterior (noviembre, diciembre), usar el año anterior
        if mes_actual <= 3 and mes_solo_detectado >= 11:
            anio_a_usar = anio_actual - 1
        else:
            anio_a_usar = anio_actual
        
        fecha_inicio = datetime(anio_a_usar, mes_solo_detectado, 1)
        fecha_fin = datetime.now()
        return {'inicio': fecha_inicio.strftime('%d/%m/%Y'), 'fin': fecha_fin.strftime('%d/%m/%Y')}
    
    # Detectar "hasta hoy", "a la fecha de hoy", "hasta la fecha"
    tiene_hasta_hoy = any(expresion in texto_lower for expresion in [
        'hasta hoy', 'a la fecha de hoy', 'hasta la fecha', 'hasta el día de hoy',
        'a hoy', 'hasta ahora', 'hasta el momento'
    ])
    
    if tiene_hasta_hoy:
        fecha_hoy = datetime.now()
        fechas_encontradas.append((fecha_hoy, len(texto), len(texto)))
    
    if not fechas_encontradas:
        return None
    
    # Ordenar fechas por posición en el texto
    fechas_encontradas.sort(key=lambda x: x[1] if len(x) > 1 else 0)
    # Extraer fechas (manejar tuplas de 3 o 4 elementos)
    fechas = [fecha for fecha, *resto in fechas_encontradas]
    
    if len(fechas) == 0:
        return None
    
    # Si hay 2 o más fechas, usar la primera como inicio y la última como fin
    if len(fechas) >= 2:
        fechas.sort()
        inicio = fechas[0].strftime('%d/%m/%Y')
        fin = fechas[-1].strftime('%d/%m/%Y')
        return {'inicio': inicio, 'fin': fin}
    elif len(fechas) == 1:
        # Si solo hay una fecha, buscar contexto para determinar si es inicio o fin
        fecha_str = fechas[0].strftime('%d/%m/%Y')
        
        # Buscar palabras clave alrededor de la fecha
        fecha_pos = fechas_encontradas[0][1]
        contexto_antes = texto_lower[max(0, fecha_pos-50):fecha_pos]
        contexto_despues = texto_lower[fecha_pos:min(len(texto_lower), fecha_pos+50)]
        
        # Verificar si es fecha de inicio
        es_inicio = any(palabra in contexto_antes for palabra in [
            'desde', 'inicio', 'del periodo', 'del', 'a partir', 'comenzando'
        ]) or any(palabra in contexto_despues for palabra in [
            'hasta hoy', 'a la fecha de hoy', 'hasta la fecha', 'a hoy'
        ])
        
        # Verificar si es fecha de fin
        es_fin = any(palabra in contexto_antes for palabra in [
            'hasta', 'al', 'fin', 'termino', 'final'
        ]) or tiene_hasta_hoy
        
        if es_inicio or tiene_hasta_hoy:
            # Es fecha de inicio, usar fecha actual como fin
            fin = datetime.now().strftime('%d/%m/%Y')
            return {'inicio': fecha_str, 'fin': fin}
        elif es_fin:
            # Es fecha de fin, usar fecha anterior como inicio (ej: último mes)
            from datetime import timedelta
            inicio_obj = fechas[0] - timedelta(days=30)
            inicio = inicio_obj.strftime('%d/%m/%Y')
            return {'inicio': inicio, 'fin': fecha_str}
        else:
            # Si no hay contexto claro pero hay "hasta hoy", asumir que es inicio
            if tiene_hasta_hoy:
                fin = datetime.now().strftime('%d/%m/%Y')
                return {'inicio': fecha_str, 'fin': fin}
            # Si no, asumir que es fecha de fin y usar 30 días antes como inicio
            from datetime import timedelta
            inicio_obj = fechas[0] - timedelta(days=30)
            inicio = inicio_obj.strftime('%d/%m/%Y')
            return {'inicio': inicio, 'fin': fecha_str}
    
    return None


def detectar_tipo_reporte(texto):
    """
    Detecta qué tipo de reporte se solicita.
    """
    texto_lower = texto.lower()
    
    if "agregado" in texto_lower or "consolidado" in texto_lower or "todos" in texto_lower:
        return "agregado"
    elif "individual" in texto_lower or "punto" in texto_lower or "nodo" in texto_lower:
        return "individual"
    else:
        return "ambos"  # Por defecto, ambos


def obtener_nombre_descriptivo_archivo(nombre_archivo):
    """
    Extrae los IDs del nombre del archivo y obtiene los nombres descriptivos.
    Maneja tanto formatos antiguos (con IDs) como nuevos (con nombres descriptivos).
    
    Args:
        nombre_archivo: Nombre del archivo PDF (ej: "Reporte_000028_000028-01_20251201_20251224.pdf" 
                      o "Reporte_La_Florida_Liceo_Alto_Cordillera_20251201_20251224.pdf")
    
    Returns:
        Nombre descriptivo legible (ej: "Reporte - La Florida - Liceo Alto Cordillera") 
        o el nombre original si no se puede procesar
    """
    import requests
    import re
    
    try:
        # Remover extensión
        nombre_sin_ext = nombre_archivo.replace('.pdf', '').replace('.docx', '')
        
        # Verificar si ya tiene formato descriptivo (contiene nombres en lugar de solo números)
        # Patrón para reportes agregados descriptivos: Reporte_Agregado_{nombreEmpresa}_{fechas}
        patron_agregado_descriptivo = r"Reporte_Agregado_([A-Za-z_]+)_(\d{8})_(\d{8})"
        match_agregado_desc = re.match(patron_agregado_descriptivo, nombre_sin_ext)
        
        if match_agregado_desc:
            nombre_empresa_limpio = match_agregado_desc.group(1).replace('_', ' ')
            return f"Reporte Agregado - {nombre_empresa_limpio}"
        
        # Patrón para reportes individuales descriptivos: Reporte_{nombreEmpresa}_{nombreNodo}_{fechas}
        patron_descriptivo = r"Reporte_([A-Za-z_]+)_([A-Za-z_]+)_(\d{8})_(\d{8})"
        match_desc = re.match(patron_descriptivo, nombre_sin_ext)
        
        if match_desc:
            nombre_empresa_limpio = match_desc.group(1).replace('_', ' ')
            nombre_nodo_limpio = match_desc.group(2).replace('_', ' ')
            return f"Reporte - {nombre_empresa_limpio} - {nombre_nodo_limpio}"
        
        # Si no es descriptivo, intentar convertir desde formato antiguo con IDs
        # Patrón para reportes agregados antiguos: Reporte_Agregado_{companyId}_{startDate}_{endDate}
        patron_agregado = r"Reporte_Agregado_(\d+)_(\d{8})_(\d{8})"
        match_agregado = re.match(patron_agregado, nombre_sin_ext)
        
        if match_agregado:
            company_id = match_agregado.group(1)
            
            # Obtener nombre de la empresa
            try:
                url_empresa = f"http://104.248.53.141:7001/wes/api/acl-entities/v1/companies/{company_id}"
                response = requests.get(url_empresa, timeout=5)
                if response.status_code == 200:
                    empresa_data = response.json()
                    nombre_empresa = empresa_data.get('name', company_id)
                    return f"Reporte Agregado - {nombre_empresa}"
            except Exception:
                pass
            
            return f"Reporte Agregado - {company_id}"
        
        # Patrón para reportes individuales antiguos: Reporte_{companyId}_{nodeId}_{startDate}_{endDate}
        patron = r"Reporte_(\d+)_(\d+-\d+)_(\d+)_(\d+)"
        match = re.match(patron, nombre_sin_ext)
        
        if not match:
            # Intentar otro formato posible con fechas de 8 dígitos
            patron2 = r"Reporte_(\d+)_(\d+-\d+)_(\d{8})_(\d{8})"
            match = re.match(patron2, nombre_sin_ext)
        
        if match:
            company_id = match.group(1)
            node_id = match.group(2)
            
            # Obtener nombre de la empresa
            try:
                url_empresa = f"http://104.248.53.141:7001/wes/api/acl-entities/v1/companies/{company_id}"
                response = requests.get(url_empresa, timeout=5)
                if response.status_code == 200:
                    empresa_data = response.json()
                    nombre_empresa = empresa_data.get('name', company_id)
                    
                    # Obtener nombre del nodo
                    nodes = empresa_data.get('nodes', [])
                    nombre_nodo = node_id
                    for node in nodes:
                        if node.get('nodeId') == node_id:
                            nombre_nodo = node.get('name', node_id)
                            break
                    
                    return f"Reporte - {nombre_empresa} - {nombre_nodo}"
            except Exception:
                pass
        
        # Si no se puede procesar, retornar nombre original sin extensión pero formateado
        return nombre_sin_ext.replace('_', ' ')
    except Exception:
        return nombre_archivo.replace('.pdf', '').replace('.docx', '').replace('_', ' ')


def generar_correo_personalizado(contacto, info_reporte, archivos_adjuntos=None, link_carpeta=None, formato="PDF"):
    """
    Genera un correo personalizado para un contacto basado en su información.
    
    Args:
        contacto: Diccionario con información del contacto
        info_reporte: Información sobre el reporte generado
        archivos_adjuntos: Lista de rutas a archivos PDF o Word a adjuntar
        link_carpeta: Link a la carpeta en Google Drive (opcional)
        formato: Formato de los archivos adjuntos ("PDF" o "Word")
    
    Returns:
        Tupla (cuerpo_texto, cuerpo_html)
    """
    if not contacto:
        contacto = {
            "tratamiento": "Estimado/a",
            "despedida": "Quedo atento a tus comentarios.",
            "firma": "Agente IA de WES"
        }
    
    tratamiento = contacto.get("tratamiento", "Estimado/a")
    despedida = contacto.get("despedida", "Quedo atento a tus comentarios.")
    firma = "Agente IA de WES"  # Siempre presentarse como agente IA
    
    empresa = info_reporte.get("empresa", "la empresa solicitada")
    tipo_reporte = info_reporte.get("tipo_reporte", "reporte")
    puntos_monitoreo = info_reporte.get("puntos_monitoreo", [])
    periodo = info_reporte.get("periodo")
    
    # Reseña de cómo se elaboró el reporte
    reseña_elaboracion = f"""
RESEÑA DE ELABORACIÓN DEL REPORTE:

Este reporte fue elaborado mediante un proceso automatizado que incluye:

1. RECOPILACIÓN DE DATOS:
   - Acceso automatizado a los sistemas instalados en terreno y sistemas de monitoreo
   - Consulta a la API de WES para obtener información de los puntos de monitoreo
   - Extracción de datos de consumo horario y diario del periodo analizado desde los sistemas de monitoreo
   - Recopilación de alertas y eventos registrados en el sistema
   - Obtención de métricas de consumo efectivo y consumos nocturnos

2. PROCESAMIENTO Y ANÁLISIS:
   - Cálculo de consumos totales, promedios diarios y mensuales
   - Identificación de consumos nocturnos (22:00-07:00) en los últimos 2 días
   - Proyección de consumos nocturnos basada en las últimas 2 alertas nocturnas registradas
   - Cálculo de consumo efectivo (consumo total menos proyección de consumos nocturnos)
   - Análisis comparativo entre puntos de monitoreo

3. GENERACIÓN DE VISUALIZACIONES:
   - Creación de gráficos de consumo diario y horario
   - Generación de gráficos comparativos entre puntos
   - Elaboración de tablas resumen con métricas clave
   - Análisis de ranking de consumo por punto

4. CONSOLIDACIÓN:
   - Generación de reportes individuales por punto de monitoreo (cuando corresponde)
   - Generación de reporte agregado con análisis comparativo (cuando corresponde)
   - Inclusión de narrativas explicativas y recomendaciones
   - Formateo y estructuración del documento final

Los datos fueron procesados de manera automatizada, garantizando consistencia y precisión en los cálculos y visualizaciones presentadas.
"""
    
    # Información de archivos adjuntos con nombres descriptivos
    info_adjuntos = ""
    if archivos_adjuntos:
        tipo_archivo = formato.upper()
        info_adjuntos = f"\nREPORTES ADJUNTOS ({tipo_archivo}):\n"
        for i, archivo_path in enumerate(archivos_adjuntos, 1):
            nombre_archivo = Path(archivo_path).name
            nombre_descriptivo = obtener_nombre_descriptivo_archivo(nombre_archivo)
            info_adjuntos += f"- {nombre_descriptivo}\n"
        info_adjuntos += "\n"
    
    # Presentación (al inicio)
    presentacion = """PRESENTACIÓN:

Soy un agente de inteligencia artificial al servicio de WES, diseñado específicamente para apoyar la generación y distribución automatizada de reportes de monitoreo de consumo de agua. Obtengo los datos directamente de los sistemas instalados en terreno y de los sistemas de monitoreo, procesándolos de manera eficiente y precisa para generar análisis detallados que facilitan la toma de decisiones.

"""
    
    # Cuerpo en texto plano
    cuerpo_texto = f"""{presentacion}{tratamiento},

Te informo que he generado el reporte solicitado para {empresa}."""
    
    if link_carpeta:
        cuerpo_texto += f"""

ENLACE A LA CARPETA COMPARTIDA:
{link_carpeta}"""
    
    cuerpo_texto += f"""

El reporte incluye:
- Análisis de consumo de agua
- Identificación de consumos nocturnos
- Gráficos y visualizaciones
- Métricas clave y recomendaciones

{reseña_elaboracion}

{despedida}

Saludos,
{firma}"""
    
    # Agregar reportes adjuntos al final
    if info_adjuntos:
        cuerpo_texto += info_adjuntos
        cuerpo_texto += "\nTe invito a revisar los archivos adjuntos que contienen el análisis completo del reporte generado.\n"
    else:
        cuerpo_texto += "\nTe invito a revisar los archivos adjuntos que contienen el análisis completo del reporte generado.\n"
    
    # Cuerpo en HTML
    cuerpo_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
        .header {{ background-color: #f4f4f4; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
        .content {{ margin: 20px 0; }}
        .section {{ margin: 15px 0; padding: 15px; background-color: #f9f9f9; border-left: 4px solid #4472C4; }}
        .link-box {{ background-color: #e8f4f8; padding: 15px; border-radius: 5px; margin: 15px 0; }}
        .link-box a {{ color: #1a73e8; text-decoration: none; font-weight: bold; font-size: 1.1em; }}
        .link-box a:hover {{ text-decoration: underline; }}
        .footer {{ margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 0.9em; color: #666; }}
        ul {{ margin: 10px 0; padding-left: 20px; }}
        li {{ margin: 5px 0; }}
        h3 {{ color: #4472C4; margin-top: 0; }}
    </style>
</head>
<body>
    <div class="header">
        <h2>Reporte Generado - {empresa}</h2>
        <p><strong>Fecha de generación:</strong> {datetime.now().strftime("%d de %B de %Y")}</p>
    </div>
    
    <div class="content">
        <div class="section">
            <h3>Presentación:</h3>
            <p>Soy un <strong>agente de inteligencia artificial al servicio de WES</strong>, diseñado específicamente para apoyar la generación y distribución automatizada de reportes de monitoreo de consumo de agua. Obtengo los datos directamente de los <strong>sistemas instalados en terreno</strong> y de los <strong>sistemas de monitoreo</strong>, procesándolos de manera eficiente y precisa para generar análisis detallados que facilitan la toma de decisiones.</p>
        </div>
        
        <p>{tratamiento},</p>
        
        <p>Te informo que he generado el reporte solicitado para <strong>{empresa}</strong>.</p>
"""
    
    if link_carpeta:
        cuerpo_html += f"""
        <div class="link-box">
            <h3>🔗 Enlace a la carpeta compartida:</h3>
            <p><a href="{link_carpeta}" target="_blank">{link_carpeta}</a></p>
        </div>
"""
    
    cuerpo_html += f"""
        <div class="section">
            <h3>Contenido del reporte:</h3>
            <ul>
                <li>Análisis de consumo de agua</li>
                <li>Identificación de consumos nocturnos</li>
                <li>Gráficos y visualizaciones</li>
                <li>Métricas clave y recomendaciones</li>
            </ul>
        </div>
        
        <div class="section">
            <h3>Metodología de obtención de datos:</h3>
            <p>Los datos fueron obtenidos directamente de los <strong>sistemas instalados en terreno</strong> y de los <strong>sistemas de monitoreo</strong> mediante consultas automatizadas a la API de WES, accediendo a:</p>
            <ul>
                <li>Información de todos los puntos de monitoreo instalados en terreno</li>
                <li>Datos de consumo horario y diario del periodo analizado desde los sistemas de monitoreo</li>
                <li>Alertas y eventos registrados en el sistema</li>
                <li>Métricas de consumo efectivo y consumos nocturnos</li>
            </ul>
        </div>
        
        <div class="section">
            <h3>Metodología de elaboración del informe:</h3>
            <ol>
                <li><strong>Recopilación de datos</strong> desde la API de WES</li>
                <li><strong>Procesamiento y análisis</strong> de consumos</li>
                <li><strong>Generación de visualizaciones</strong> (gráficos, tablas)</li>
                <li><strong>Consolidación</strong> en reportes individuales y/o agregados</li>
            </ol>
        </div>
        
        <div class="section">
            <h3>Reseña de elaboración del reporte:</h3>
            <p>Este reporte fue elaborado mediante un proceso automatizado que incluye:</p>
            <ol>
                <li><strong>Recopilación de datos:</strong> Acceso automatizado a los <strong>sistemas instalados en terreno</strong> y <strong>sistemas de monitoreo</strong> mediante consultas a la API de WES para obtener información de los puntos de monitoreo, datos de consumo horario y diario, alertas y eventos registrados, y métricas de consumo efectivo y consumos nocturnos.</li>
                <li><strong>Procesamiento y análisis:</strong> Cálculo de consumos totales, promedios diarios y mensuales, identificación de consumos nocturnos (22:00-07:00) en los últimos 2 días, proyección basada en las últimas 2 alertas nocturnas, cálculo de consumo efectivo, y análisis comparativo entre puntos.</li>
                <li><strong>Generación de visualizaciones:</strong> Creación de gráficos de consumo diario y horario, gráficos comparativos entre puntos, tablas resumen con métricas clave, y análisis de ranking de consumo por punto.</li>
                <li><strong>Consolidación:</strong> Generación de reportes individuales y/o agregados, inclusión de narrativas explicativas y recomendaciones, y formateo del documento final.</li>
            </ol>
            <p>Los datos fueron procesados de manera automatizada, garantizando consistencia y precisión en los cálculos y visualizaciones presentadas.</p>
        </div>
        
        <p>{despedida}</p>
"""
    
    # Buscar si hay un reporte agregado y su gráfica de comparación
    grafica_comparacion_path = None
    if archivos_adjuntos:
        for archivo_path in archivos_adjuntos:
            archivo_path_obj = Path(archivo_path)
            nombre_archivo = archivo_path_obj.name
            
            # Verificar si es un reporte agregado
            if "Agregado" in nombre_archivo or "agregado" in nombre_archivo.lower():
                # El directorio del reporte (PDF y Word están en el mismo directorio)
                directorio_reporte = archivo_path_obj.parent
                
                # Buscar la gráfica de comparación en el directorio del reporte
                grafica_path = directorio_reporte / "chart_comparacion_nodos.png"
                
                if grafica_path.exists():
                    grafica_comparacion_path = grafica_path
                    print(f"[INFO] Gráfica de comparación encontrada: {grafica_path}")
                    break
                else:
                    # Si no está en el mismo directorio, buscar en el directorio padre (por si el PDF está en otro lugar)
                    directorio_padre = directorio_reporte.parent
                    grafica_path_padre = directorio_padre / "chart_comparacion_nodos.png"
                    if grafica_path_padre.exists():
                        grafica_comparacion_path = grafica_path_padre
                        print(f"[INFO] Gráfica de comparación encontrada en directorio padre: {grafica_path_padre}")
                        break
    
    # Agregar gráfica de comparación si existe (antes de los adjuntos)
    if grafica_comparacion_path and grafica_comparacion_path.exists():
        # Leer la imagen y convertirla a base64 para embebida
        try:
            import base64
            with open(grafica_comparacion_path, 'rb') as img_file:
                img_data = base64.b64encode(img_file.read()).decode('utf-8')
                img_ext = grafica_comparacion_path.suffix[1:]  # 'png' sin el punto
                img_data_uri = f"data:image/{img_ext};base64,{img_data}"
                
                cuerpo_html += f"""
        <div class="section">
            <h3>📊 Comparación de Consumo por Punto</h3>
            <p>Consumo total registrado en cada punto de monitoreo durante el periodo analizado:</p>
            <div style="text-align: center; margin: 20px 0;">
                <img src="{img_data_uri}" alt="Comparación de Consumo por Punto" style="max-width: 100%; height: auto; border: 1px solid #ddd; border-radius: 5px;" />
            </div>
        </div>
"""
        except Exception as e:
            print(f"[ADVERTENCIA] No se pudo incluir la gráfica de comparación: {e}")
    
    # Agregar reportes adjuntos al final (antes del footer)
    if archivos_adjuntos:
        tipo_archivo = formato.upper()
        cuerpo_html += f"""
        <div class="section">
            <h3>📎 Reportes adjuntos ({tipo_archivo}):</h3>
            <p>Te invito a revisar los archivos adjuntos que contienen el análisis completo del reporte generado.</p>
            <ul>
"""
        for archivo_path in archivos_adjuntos:
            nombre_archivo = Path(archivo_path).name
            nombre_descriptivo = obtener_nombre_descriptivo_archivo(nombre_archivo)
            cuerpo_html += f"                <li><strong>{nombre_descriptivo}</strong></li>\n"
        cuerpo_html += """            </ul>
        </div>
"""
    else:
        # Si no hay adjuntos, igual invitar a revisar (por si acaso)
        cuerpo_html += """
        <div class="section">
            <p><strong>Te invito a revisar los archivos adjuntos</strong> que contienen el análisis completo del reporte generado.</p>
        </div>
"""
    
    cuerpo_html += f"""
    </div>
    
    <div class="footer">
        <p>Saludos,<br>
        <strong>{firma}</strong></p>
    </div>
</body>
</html>
"""
    
    return cuerpo_texto, cuerpo_html


def enviar_correo_personalizado(contacto, info_reporte, archivos_pdf=None, archivos_word=None, archivos_ppt=None, link_carpeta=None):
    """
    Envía un correo personalizado a un contacto con reportes PDF o Word adjuntos.
    
    Args:
        contacto: Diccionario con información del contacto
        info_reporte: Información sobre el reporte generado
        archivos_pdf: Lista de rutas a archivos PDF a adjuntar
        archivos_word: Lista de rutas a archivos Word (.docx) a adjuntar
        link_carpeta: Link a la carpeta en Google Drive (opcional)
    """
    if not contacto:
        print("[ERROR] No se pudo obtener información del contacto")
        return False
    
    # Determinar qué archivos adjuntar (PDF o Word)
    archivos_adjuntos = archivos_pdf if archivos_pdf else archivos_word
    formato_adjunto = "PDF" if archivos_pdf else "Word"
    
    cuerpo_texto, cuerpo_html = generar_correo_personalizado(contacto, info_reporte, archivos_adjuntos, link_carpeta, formato_adjunto)
    
    # Construir correo
    msg = MIMEMultipart('alternative')
    
    # Headers
    msg["From"] = f"Agente IA de WES <{SMTP_USUARIO}>"
    msg["To"] = contacto["email"]
    msg["Reply-To"] = SMTP_USUARIO
    empresa = info_reporte.get('empresa', '')
    msg["Subject"] = f"Reporte {empresa} - Generado por Agente IA de WES"
    
    # Headers importantes para evitar spam
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="wes.cl")
    msg["X-Mailer"] = "WES AI Agent - Report Generator"
    msg["X-Priority"] = "3"
    msg["MIME-Version"] = "1.0"
    
    # Agregar ambas versiones (texto y HTML)
    part1 = MIMEText(cuerpo_texto, "plain", "utf-8")
    part2 = MIMEText(cuerpo_html, "html", "utf-8")
    
    msg.attach(part1)
    msg.attach(part2)
    
    # Adjuntar archivos PDF
    if archivos_pdf:
        print(f"[INFO] Adjuntando {len(archivos_pdf)} archivo(s) PDF...")
        for pdf_path in archivos_pdf:
            if os.path.exists(pdf_path):
                try:
                    with open(pdf_path, "rb") as f:
                        pdf_attachment = MIMEApplication(f.read(), _subtype="pdf")
                        nombre_archivo_original = Path(pdf_path).name
                        # Obtener nombre descriptivo para el adjunto
                        nombre_descriptivo = obtener_nombre_descriptivo_archivo(nombre_archivo_original)
                        # Agregar extensión .pdf al nombre descriptivo
                        nombre_adjunto = f"{nombre_descriptivo}.pdf"
                        pdf_attachment.add_header(
                            "Content-Disposition",
                            "attachment",
                            filename=nombre_adjunto
                        )
                        msg.attach(pdf_attachment)
                        print(f"  [OK] Adjuntado: {nombre_adjunto}")
                except Exception as e:
                    print(f"  [ERROR] No se pudo adjuntar {pdf_path}: {e}")
            else:
                print(f"  [ERROR] Archivo no encontrado: {pdf_path}")
    
    # Adjuntar archivos Word
    if archivos_word:
        print(f"[INFO] Adjuntando {len(archivos_word)} archivo(s) Word...")
        for word_path in archivos_word:
            if os.path.exists(word_path):
                try:
                    with open(word_path, "rb") as f:
                        word_attachment = MIMEApplication(f.read(), _subtype="vnd.openxmlformats-officedocument.wordprocessingml.document")
                        nombre_archivo_original = Path(word_path).name
                        # Obtener nombre descriptivo para el adjunto
                        nombre_descriptivo = obtener_nombre_descriptivo_archivo(nombre_archivo_original)
                        # Si el nombre descriptivo no tiene extensión, agregar .docx
                        if not nombre_descriptivo.endswith('.docx'):
                            nombre_adjunto = f"{nombre_descriptivo}.docx"
                        else:
                            nombre_adjunto = nombre_descriptivo
                        word_attachment.add_header(
                            "Content-Disposition",
                            "attachment",
                            filename=nombre_adjunto
                        )
                        msg.attach(word_attachment)
                        print(f"  [OK] Adjuntado: {nombre_adjunto}")
                except Exception as e:
                    print(f"  [ERROR] No se pudo adjuntar {word_path}: {e}")
            else:
                print(f"  [ERROR] Archivo no encontrado: {word_path}")
    
    # Adjuntar archivos PPT (solo si se solicita presentación)
    if archivos_ppt:
        print(f"[INFO] Adjuntando {len(archivos_ppt)} archivo(s) PPT...")
        for ppt_path in archivos_ppt:
            if os.path.exists(ppt_path):
                try:
                    with open(ppt_path, "rb") as f:
                        ppt_attachment = MIMEApplication(f.read(), _subtype="vnd.openxmlformats-officedocument.presentationml.presentation")
                        nombre_archivo_original = Path(ppt_path).name
                        # Obtener nombre descriptivo para el adjunto
                        nombre_descriptivo = obtener_nombre_descriptivo_archivo(nombre_archivo_original)
                        # Si el nombre descriptivo no tiene extensión, agregar .pptx
                        if not nombre_descriptivo.endswith('.pptx'):
                            nombre_adjunto = f"{nombre_descriptivo}.pptx"
                        else:
                            nombre_adjunto = nombre_descriptivo
                        ppt_attachment.add_header(
                            "Content-Disposition",
                            "attachment",
                            filename=nombre_adjunto
                        )
                        msg.attach(ppt_attachment)
                        print(f"  [OK] Adjuntado: {nombre_adjunto}")
                except Exception as e:
                    print(f"  [ERROR] No se pudo adjuntar {ppt_path}: {e}")
            else:
                print(f"  [ERROR] Archivo no encontrado: {ppt_path}")
    
    # Enviar correo
    try:
        print(f"[INFO] Enviando correo personalizado a {contacto['email']}...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.sendmail(SMTP_USUARIO, [contacto["email"]], msg.as_string())
        print(f"[OK] Correo enviado correctamente a {contacto.get('nombre_completo', contacto['email'])}")
        return True
    except Exception as e:
        print(f"[ERROR] Falló el envío del correo: {e}")
        import traceback
        traceback.print_exc()
        return False


def obtener_todas_las_empresas():
    """
    Obtiene todas las empresas disponibles desde la API WES.
    
    Returns:
        Diccionario con mapeo de nombres de empresas (en mayúsculas) a sus IDs.
        Formato: {"NOMBRE_EMPRESA": "company_id", ...}
    """
    import requests
    
    empresas_dict = {}
    
    try:
        url = "http://104.248.53.141:7001/wes/api/acl-entities/v1/configuration/companies"
        print(f"[DEBUG] Conectando a API para obtener empresas: {url}")
        response = requests.get(url, timeout=15)
        print(f"[DEBUG] Respuesta recibida: Status {response.status_code}")
        
        if response.status_code == 200:
            empresas = response.json()
            
            if isinstance(empresas, list):
                for empresa in empresas:
                    company_id = empresa.get('companyId')
                    company_name = empresa.get('name', '')
                    
                    if company_id and company_name:
                        # Agregar variaciones del nombre para facilitar la búsqueda
                        nombre_upper = company_name.upper().strip()
                        empresas_dict[nombre_upper] = company_id
                        
                        # También agregar sin espacios extra
                        nombre_sin_espacios = nombre_upper.replace('  ', ' ').strip()
                        if nombre_sin_espacios != nombre_upper:
                            empresas_dict[nombre_sin_espacios] = company_id
                        
                        # Agregar variaciones comunes
                        if 'ESVAL' in nombre_upper:
                            nombre_sin_esval = nombre_upper.replace(' ESVAL', '').replace('ESVAL ', '').strip()
                            if nombre_sin_esval:
                                empresas_dict[nombre_sin_esval] = company_id
                
                print(f"[OK] Se obtuvieron {len(empresas)} empresa(s) desde la API")
            else:
                print(f"[ADVERTENCIA] La respuesta de la API no es una lista: {type(empresas)}")
        else:
            print(f"[ADVERTENCIA] Error al obtener empresas desde la API: código {response.status_code}")
    
    except Exception as e:
        print(f"[ADVERTENCIA] Error al consultar API de empresas: {e}")
        print(f"[INFO] Se usará el diccionario de empresas conocido como respaldo")
    
    return empresas_dict


def obtener_empresa_id_por_nombre(nombre_empresa, empresas_dict=None):
    """
    Busca el ID de una empresa por su nombre.
    
    Args:
        nombre_empresa: Nombre de la empresa a buscar
        empresas_dict: Diccionario de empresas (si es None, se obtiene desde la API)
    
    Returns:
        ID de la empresa si se encuentra, None en caso contrario
    """
    if not nombre_empresa:
        return None
    
    # Si no se proporciona el diccionario, obtenerlo desde la API
    if empresas_dict is None:
        empresas_dict = obtener_todas_las_empresas()
    
    nombre_upper = nombre_empresa.upper().strip()
    
    # Buscar coincidencia exacta primero
    empresa_id = empresas_dict.get(nombre_upper)
    if empresa_id:
        return empresa_id
    
    # Buscar coincidencia parcial
    for nombre_api, id_api in empresas_dict.items():
        if nombre_upper in nombre_api or nombre_api in nombre_upper:
            return id_api
    
    # Si no se encuentra, intentar con el diccionario conocido como respaldo
    EMPRESA_IDS_CONOCIDOS = {
        "DERCO": "000012",
        "BUPA": "000029",
        "PARQUE ARAUCO": "000025",
        "Parque Arauco": "000025",
        "FUNDO ZAPALLAR": "000027",
        "Fundo Zapallar": "000027",
        "FUNDO ZAPALLAR ESVAL": "000027",
        "Fundo Zapallar ESVAL": "000027",
    }
    
    empresa_id = EMPRESA_IDS_CONOCIDOS.get(nombre_upper)
    if empresa_id:
        return empresa_id
    
    # Buscar coincidencia parcial en el diccionario conocido
    for nombre_conocido, id_conocido in EMPRESA_IDS_CONOCIDOS.items():
        if nombre_upper in nombre_conocido or nombre_conocido in nombre_upper:
            return id_conocido
    
    return None


def detectar_mall_parque_arauco(texto):
    """
    Detecta si se menciona un mall específico de Parque Arauco en el texto.
    Detecta variaciones como "mall maipu", "parque arauco maipu", "maipu", etc.
    
    Args:
        texto: Texto del correo (asunto + cuerpo) en minúsculas
    
    Returns:
        Diccionario con información del mall detectado o None si no se detecta
        Formato: {'mall': 'maipu', 'nodos': ['000025-08', ...], 'nombre_completo': 'Parque Arauco Maipu'}
    """
    texto_lower = texto.lower()
    
    # Buscar cada mall en el texto
    for mall_key, mall_info in PARQUE_ARAUCO_MALLS.items():
        for nombre_mall in mall_info["nombres"]:
            # Buscar el nombre del mall en el texto
            if nombre_mall in texto_lower:
                # Verificar que también se mencione "parque arauco", "mall", "centro comercial" o que el nombre del mall ya lo incluya
                tiene_contexto = any(palabra in texto_lower for palabra in ["parque arauco", "mall", "centro comercial"])
                
                # Si el nombre del mall ya incluye "parque arauco", no necesitamos verificar contexto adicional
                # También aceptar si solo se menciona el nombre del mall (ej: "maipu") sin contexto, ya que es específico
                if "parque arauco" in nombre_mall or tiene_contexto:
                    return {
                        'mall': mall_key,
                        'nodos': mall_info["nodos"],
                        'nombre_completo': f"Parque Arauco {mall_key.title()}"
                    }
                # Si el nombre del mall es específico (no es una palabra común), aceptarlo sin contexto adicional
                # Esto permite detectar "mall maipu" o solo "maipu" cuando está claro que es un mall
                elif mall_key not in ['norte', 'sur', 'este', 'oeste']:  # Evitar palabras comunes
                    # Verificar que no sea parte de otra palabra (usar límites de palabra)
                    import re
                    patron = r'\b' + re.escape(nombre_mall) + r'\b'
                    if re.search(patron, texto_lower):
                        return {
                            'mall': mall_key,
                            'nodos': mall_info["nodos"],
                            'nombre_completo': f"Parque Arauco {mall_key.title()}"
                        }
    
    return None


def detectar_nodos_especificos_en_texto(texto_correo, todos_los_nodos):
    """
    Detecta nodos específicos mencionados en el texto del correo.
    
    Args:
        texto_correo: Texto del correo (asunto + cuerpo) en minúsculas
        todos_los_nodos: Lista de todos los nodos disponibles de la empresa
    
    Returns:
        Lista de diccionarios con información de nodos específicos detectados, o lista vacía si no se detectan
    """
    import unicodedata
    
    nodos_especificos_mencionados = []
    
    # Palabras clave comunes que indican un nodo específico
    palabras_filtro = ["solo", "solamente", "únicamente", "solo el", "solo la", "solo los", "solo las", "solo de", "punto", "puntos", "nodo", "nodos"]
    
    # Normalizar nombres: quitar acentos y caracteres especiales para comparación
    def normalizar(texto):
        texto = unicodedata.normalize('NFD', texto)
        texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
        return texto.lower()
    
    texto_normalizado = normalizar(texto_correo)
    
    # Buscar si hay palabras de filtro que indiquen un nodo específico
    tiene_filtro = any(palabra in texto_correo for palabra in palabras_filtro)
    
    # SIEMPRE buscar nodos específicos mencionados
    # Primero buscar con palabras de filtro si existen
    if tiene_filtro:
        # Buscar coincidencias entre nombres de nodos y el texto del correo
        for node in todos_los_nodos:
            node_name = node.get('name', '').strip()
            node_name_lower = node_name.lower()
            
            # Buscar coincidencia exacta del nombre completo del nodo
            if node_name_lower in texto_correo:
                nodos_especificos_mencionados.append({
                    'nodeId': node.get('nodeId'),
                    'name': node_name,
                    'match': 'nombre completo'
                })
                continue
            
            # Si no hay coincidencia exacta, buscar por palabras significativas
            palabras_nodo = [p for p in node_name_lower.split() if len(p) > 3 and p not in ['del', 'de', 'la', 'el', 'los', 'las', 'y', 'en', 'con', 'para', 'por']]
            
            # Si todas las palabras significativas del nodo aparecen en el correo
            if palabras_nodo and all(palabra in texto_correo for palabra in palabras_nodo):
                nodos_especificos_mencionados.append({
                    'nodeId': node.get('nodeId'),
                    'name': node_name,
                    'match': 'palabras significativas'
                })
                continue
            
            # Buscar si alguna palabra clave del nodo aparece cerca de palabras de filtro
            for palabra_filtro in palabras_filtro:
                if palabra_filtro in texto_correo:
                    # Buscar el contexto después de la palabra de filtro
                    indice_filtro = texto_correo.find(palabra_filtro)
                    contexto = texto_correo[indice_filtro:indice_filtro+100]  # 100 caracteres después del filtro
                    
                    # Buscar palabras del nodo en el contexto
                    palabras_encontradas = [p for p in palabras_nodo if p in contexto]
                    if len(palabras_encontradas) >= 2 or (len(palabras_nodo) == 1 and palabras_encontradas):
                        nodos_especificos_mencionados.append({
                            'nodeId': node.get('nodeId'),
                            'name': node_name,
                            'match': f'cerca de "{palabra_filtro}"'
                        })
                        break
    
    # Si no se detectaron nodos específicos pero el texto es muy corto, intentar buscar de todas formas
    if not nodos_especificos_mencionados and len(texto_correo.split()) < 50:
        for node in todos_los_nodos:
            node_name = node.get('name', '').strip()
            node_name_lower = node_name.lower()
            
            # Buscar coincidencia exacta del nombre completo
            if node_name_lower in texto_correo:
                nodos_especificos_mencionados.append({
                    'nodeId': node.get('nodeId'),
                    'name': node_name,
                    'match': 'coincidencia exacta en texto corto'
                })
                break
    
    # Búsqueda adicional: buscar nombres de nodos que aparezcan en el correo sin palabras de filtro
    if not nodos_especificos_mencionados:
        for node in todos_los_nodos:
            node_name = node.get('name', '').strip()
            node_name_lower = node_name.lower()
            nombre_normalizado = normalizar(node_name_lower)
            
            # Buscar coincidencia exacta del nombre normalizado
            if nombre_normalizado in texto_normalizado:
                nodos_especificos_mencionados.append({
                    'nodeId': node.get('nodeId'),
                    'name': node_name,
                    'match': 'coincidencia directa (normalizada)'
                })
                break  # Si encontramos uno, probablemente es el que se busca
            
            # También buscar por palabras clave significativas (al menos 2 palabras del nodo)
            palabras_nodo = [p for p in nombre_normalizado.split() if len(p) > 3 and p not in ['del', 'de', 'la', 'el', 'los', 'las', 'y', 'en', 'con', 'para', 'por', 'impulsion', 'impulsi', 'matriz', 'principal']]
            if len(palabras_nodo) >= 2:
                palabras_encontradas = sum(1 for p in palabras_nodo if p in texto_normalizado)
                # Si al menos 2 palabras significativas aparecen, considerar que es el nodo
                if palabras_encontradas >= 2:
                    nodos_especificos_mencionados.append({
                        'nodeId': node.get('nodeId'),
                        'name': node_name,
                        'match': f'coincidencia por {palabras_encontradas} palabras clave'
                    })
                    break
            
            # Búsqueda especial para nombres cortos o palabras clave únicas
            if len(palabras_nodo) == 1 and palabras_nodo[0] in texto_normalizado:
                # Verificar que la palabra no sea muy común
                palabra_unica = palabras_nodo[0]
                if palabra_unica not in ['norte', 'sur', 'este', 'oeste', 'principal', 'matriz']:
                    nodos_especificos_mencionados.append({
                        'nodeId': node.get('nodeId'),
                        'name': node_name,
                        'match': f'coincidencia por palabra clave única: {palabra_unica}'
                    })
                    break
    
    return nodos_especificos_mencionados


def validar_parametros_solicitud(info_detectada):
    """
    Valida que todos los parámetros necesarios estén presentes en la solicitud.
    
    Args:
        info_detectada: Diccionario con información detectada del correo
    
    Returns:
        Tupla (es_valida, parametros_faltantes)
        - es_valida: True si todos los parámetros están presentes
        - parametros_faltantes: Lista de parámetros que faltan
    """
    if info_detectada.get("reporte_especial") == "control_nocturno":
        return True, []
    
    parametros_faltantes = []
    
    # Validar empresa
    empresa_detectada = info_detectada.get("empresa")
    # Si hay un mall detectado de Parque Arauco, la empresa está implícita (Parque Arauco)
    mall_detectado = info_detectada.get("mall_detectado")
    if not empresa_detectada and not mall_detectado:
        parametros_faltantes.append("empresa")
    elif mall_detectado and not empresa_detectada:
        # Si hay mall pero no empresa, asignar Parque Arauco automáticamente
        info_detectada["empresa"] = "Parque Arauco"
        print(f"[INFO] Empresa asignada automáticamente en validación: Parque Arauco (por mall detectado)")
    
    # Validar periodo (debe estar explícitamente mencionado)
    periodo = info_detectada.get("periodo")
    if not periodo:
        parametros_faltantes.append("periodo")
    
    # El punto/nodo es opcional (si no se especifica, se procesan todos)
    # Pero si se menciona un punto específico, debe estar claro
    
    return len(parametros_faltantes) == 0, parametros_faltantes


def enviar_correo_solicitando_informacion(contacto, parametros_faltantes, asunto_original=None, message_id_original=None, info_detectada=None):
    """
    Envía un correo al remitente solicitando la información faltante.
    Responde al correo original si se proporciona message_id_original.
    
    Args:
        contacto: Diccionario con información del contacto
        parametros_faltantes: Lista de parámetros que faltan
        asunto_original: Asunto del correo original (opcional)
        message_id_original: Message-ID del correo original para responder (opcional)
        info_detectada: Información detectada del correo original para guardar contexto (opcional)
    
    Returns:
        Message-ID del correo enviado si se envió correctamente, None en caso contrario
    """
    if not contacto:
        print("[ERROR] No se pudo obtener información del contacto para enviar correo de solicitud")
        return False
    
    tratamiento = contacto.get("tratamiento", "Estimado/a")
    email = contacto.get("email")
    
    if not email:
        print("[ERROR] No se pudo obtener el email del contacto para enviar correo de solicitud")
        return False
    
    # Mapear nombres de parámetros a descripciones
    descripciones = {
        "empresa": "el nombre de la empresa",
        "periodo": "el periodo o rango de fechas (fecha inicio y fecha fin)",
        "punto": "el punto o nodo específico (opcional, si no se especifica se procesan todos)"
    }
    
    # Construir lista de parámetros faltantes
    lista_faltantes = []
    for param in parametros_faltantes:
        if param in descripciones:
            lista_faltantes.append(f"- {descripciones[param]}")
    
    lista_faltantes_texto = "\n".join(lista_faltantes)
    
    # Cuerpo del correo en texto plano
    cuerpo_texto = f"""{tratamiento},

He recibido tu solicitud de reporte, pero para poder procesarla necesito que me proporciones la siguiente información:

{lista_faltantes_texto}

PARÁMETROS NECESARIOS PARA GENERAR REPORTES:

Para generar un reporte, necesito que me proporciones los siguientes 3 parámetros:

1. EMPRESA: El nombre de la empresa para la cual deseas el reporte.
   Ejemplo: "Parque Arauco", "La Florida", "BUPA", etc.

2. PERIODO: El rango de fechas para el reporte.
   Ejemplo: "del 01/12/2025 al 24/12/2025" o "desde 01/12/2025 hasta 24/12/2025"

3. PUNTO/NODO (opcional): Si deseas un reporte de un punto específico, menciónalo.
   Si no lo especificas, se generarán reportes de todos los puntos de la empresa.
   Ejemplo: "Placa Bancaria", "Matriz Principal", etc.

Por favor, responde a este correo con la información faltante para poder procesar tu solicitud.

Saludos,
Agente IA de WES
"""
    
    # Cuerpo del correo en HTML
    cuerpo_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
        .header {{ background-color: #f4f4f4; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
        .content {{ margin: 20px 0; }}
        .section {{ margin: 15px 0; padding: 15px; background-color: #fff3cd; border-left: 4px solid #ffc107; }}
        .info-box {{ background-color: #e8f4f8; padding: 15px; border-radius: 5px; margin: 15px 0; border-left: 4px solid #4472C4; }}
        .footer {{ margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 0.9em; color: #666; }}
        ul {{ margin: 10px 0; padding-left: 20px; }}
        li {{ margin: 5px 0; }}
        h3 {{ color: #4472C4; margin-top: 0; }}
        .parametro {{ margin: 10px 0; padding: 10px; background-color: #f9f9f9; border-radius: 3px; }}
    </style>
</head>
<body>
    <div class="header">
        <h2>Información Faltante para Generar Reporte</h2>
    </div>
    
    <div class="content">
        <p>{tratamiento},</p>
        
        <p>He recibido tu solicitud de reporte, pero para poder procesarla necesito que me proporciones la siguiente información:</p>
        
        <div class="section">
            <h3>⚠️ Información Faltante:</h3>
            <ul>
"""
    
    for param in parametros_faltantes:
        if param in descripciones:
            cuerpo_html += f"                <li><strong>{descripciones[param]}</strong></li>\n"
    
    cuerpo_html += """            </ul>
        </div>
        
        <div class="info-box">
            <h3>📋 Parámetros Necesarios para Generar Reportes:</h3>
            
            <div class="parametro">
                <strong>1. EMPRESA:</strong> El nombre de la empresa para la cual deseas el reporte.<br>
                <em>Ejemplo:</em> "Parque Arauco", "La Florida", "BUPA", etc.
            </div>
            
            <div class="parametro">
                <strong>2. PERIODO:</strong> El rango de fechas para el reporte.<br>
                <em>Ejemplo:</em> "del 01/12/2025 al 24/12/2025" o "desde 01/12/2025 hasta 24/12/2025"
            </div>
            
            <div class="parametro">
                <strong>3. PUNTO/NODO (opcional):</strong> Si deseas un reporte de un punto específico, menciónalo.<br>
                Si no lo especificas, se generarán reportes de todos los puntos de la empresa.<br>
                <em>Ejemplo:</em> "Placa Bancaria", "Matriz Principal", etc.
            </div>
        </div>
        
        <p><strong>Por favor, responde a este correo con la información faltante para poder procesar tu solicitud.</strong></p>
    </div>
    
    <div class="footer">
        <p>Saludos,<br>
        <strong>Agente IA de WES</strong></p>
    </div>
</body>
</html>
"""
    
    # Construir correo
    msg = MIMEMultipart('alternative')
    
    # Headers
    msg["From"] = f"Agente IA de WES <{SMTP_USUARIO}>"
    msg["To"] = email
    msg["Reply-To"] = SMTP_USUARIO
    
    # Generar Message-ID para este correo
    message_id_respuesta = make_msgid(domain="wes.cl")
    
    # Si hay message_id_original, responder al correo original
    if message_id_original:
        msg["In-Reply-To"] = message_id_original
        msg["References"] = message_id_original
        if asunto_original:
            # Remover "Re: " si ya existe para evitar "Re: Re: ..."
            asunto_limpio = re.sub(r'^Re:\s*', '', asunto_original, flags=re.IGNORECASE)
            msg["Subject"] = f"Re: {asunto_limpio}"
        else:
            msg["Subject"] = "Re: Información Faltante para Generar Reporte"
    else:
        if asunto_original:
            msg["Subject"] = f"Re: {asunto_original} - Información Faltante para Reporte"
        else:
            msg["Subject"] = "Información Faltante para Generar Reporte"
    
    # Headers importantes
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = message_id_respuesta
    msg["X-Mailer"] = "WES AI Agent - Report Generator"
    msg["X-Priority"] = "3"
    msg["MIME-Version"] = "1.0"
    
    # Agregar ambas versiones (texto y HTML)
    part1 = MIMEText(cuerpo_texto, "plain", "utf-8")
    part2 = MIMEText(cuerpo_html, "html", "utf-8")
    
    msg.attach(part1)
    msg.attach(part2)
    
    # Enviar correo
    try:
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg)
        
        message_id_str = str(message_id_respuesta).strip('<>')
        print(f"[OK] Correo de solicitud de información enviado a {email}")
        print(f"[OK] Message-ID de respuesta: {message_id_str}")
        
        # Guardar contexto de la solicitud pendiente
        if info_detectada:
            guardar_contexto_solicitud(message_id_str, info_detectada, parametros_faltantes)
        
        return message_id_str
    except Exception as e:
        print(f"[ERROR] No se pudo enviar correo de solicitud de información: {e}")
        return None


def generar_y_enviar_reporte_puntos_cero(remitente_email: str, contacto: dict, info_detectada: dict) -> bool:
    """
    Genera el reporte de puntos en cero y lo envía por correo.
    
    Args:
        remitente_email: Email del remitente
        contacto: Información del contacto
        info_detectada: Información detectada del correo
    
    Returns:
        True si se generó y envió exitosamente, False en caso contrario
    """
    try:
        print("[INFO] Generando reporte de puntos en cero...")
        
        # Importar funciones del script de puntos en cero
        from reporte_puntos_en_cero import (
            obtener_todos_los_nodos,
            verificar_consumo_cero,
            crear_reporte_word,
            construir_resumen_alertas
        )
        from pathlib import Path
        from datetime import datetime, timezone, timedelta
        
        # Crear carpeta de salida
        output_dir = Path("reporte en cero")
        output_dir.mkdir(exist_ok=True)
        
        # Obtener todos los nodos
        print("[INFO] Obteniendo todos los nodos del sistema...")
        todos_nodos = obtener_todos_los_nodos()
        
        if not todos_nodos:
            print("[ERROR] No se encontraron nodos en el sistema.")
            return False
        
        print(f"[OK] Se encontraron {len(todos_nodos)} nodos para verificar")
        
        # Verificar cada nodo
        print("[INFO] Verificando consumo en cero...")
        puntos_en_cero = []
        puntos_sin_datos = []
        
        for i, nodo in enumerate(todos_nodos, 1):
            node_id = nodo["nodeId"]
            node_name = nodo["nodeName"]
            
            if i % 10 == 0:
                print(f"  Progreso: {i}/{len(todos_nodos)} puntos verificados...")
            
            esta_en_cero, error = verificar_consumo_cero(node_id)
            
            if esta_en_cero:
                puntos_en_cero.append(nodo)
            elif error and "Sin datos" in error:
                puntos_sin_datos.append(nodo)
        
        print(f"[OK] Verificación completada:")
        print(f"  - Puntos en cero: {len(puntos_en_cero)}")
        print(f"  - Puntos sin datos: {len(puntos_sin_datos)}")
        
        # Construir análisis de alertas del día anterior
        fecha_alertas = datetime.now(timezone.utc) - timedelta(days=1)
        alertas_resumen = construir_resumen_alertas(todos_nodos, fecha_alertas)

        # Generar reporte Word
        print("[INFO] Generando documento Word...")
        reporte_path = crear_reporte_word(
            puntos_en_cero,
            puntos_sin_datos,
            len(todos_nodos),
            output_dir,
            alertas_resumen=alertas_resumen,
            alertas_fecha=fecha_alertas.strftime("%d-%m-%Y")
        )
        
        if not reporte_path or not Path(reporte_path).exists():
            print("[ERROR] No se pudo generar el reporte")
            return False
        
        print(f"[OK] Reporte generado: {reporte_path}")
        
        # Enviar correo con el reporte
        print(f"[INFO] Enviando correo a {remitente_email}...")
        
        # Usar configuración SMTP global
        SMTP_PUERTO = 587
        
        # Crear mensaje
        msg = MIMEMultipart()
        msg['From'] = SMTP_USUARIO
        msg['To'] = remitente_email
        msg['Subject'] = "Reporte de Puntos en Cero - Sistema WES"
        
        # Cuerpo del correo
        nombre_contacto = contacto.get('nombre_completo', 'Estimado/a')
        tratamiento = contacto.get('tratamiento', 'Estimado/a')
        despedida = contacto.get('despedida', 'Quedo atento a tus comentarios.')
        firma = contacto.get('firma', 'Sistema WES')
        
        cuerpo = f"""
{tratamiento} {nombre_contacto},

Se adjunta el reporte de puntos de monitoreo que están marcando cero consumo o sin datos disponibles.

El reporte incluye:
- Resumen ejecutivo con estadísticas generales
- Tabla detallada de puntos marcando cero
- Tabla detallada de puntos sin datos disponibles

{despedida}

{firma}
"""
        
        msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))
        
        # Adjuntar reporte
        with open(reporte_path, 'rb') as f:
            adjunto = MIMEApplication(f.read())
            adjunto.add_header(
                'Content-Disposition',
                'attachment',
                filename=Path(reporte_path).name
            )
            msg.attach(adjunto)
        
        # Enviar correo
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg)
        
        print(f"[OK] Correo enviado exitosamente a {remitente_email}")
        return True
        
    except Exception as e:
        print(f"[ERROR] Error al generar/enviar reporte de puntos en cero: {e}")
        import traceback
        traceback.print_exc()
        return False


def _procesar_reporte_control_nocturno_email(info_detectada, contacto):
    """
    Genera el reporte de control nocturno (Renca / Excel de horarios) y lo envía por SMTP.
    No usa el flujo de empresas/nodos de la API de reportes estándar.
    """
    from datetime import datetime

    from control_nocturno import generar_reporte_control_nocturno

    remitente_email = info_detectada.get("remitente_email")
    if not contacto:
        contacto = {
            "nombre_completo": remitente_email or "Usuario",
            "email": remitente_email,
            "tratamiento": "Estimado/a",
            "despedida": "Quedo atento a tus comentarios.",
            "firma": "Agente IA de WES",
        }

    ahora = datetime.now()
    desde = datetime.combine(ahora.date(), datetime.min.time())
    hasta = datetime.combine(ahora.date(), datetime.min.time())

    periodo = info_detectada.get("periodo")
    if periodo and isinstance(periodo, dict):
        ini_s = periodo.get("inicio")
        fin_s = periodo.get("fin")
        if ini_s and fin_s:
            try:
                desde = datetime.strptime(ini_s, "%d/%m/%Y")
                hasta = datetime.strptime(fin_s, "%d/%m/%Y")
            except (TypeError, ValueError):
                print("[ADVERTENCIA] Período detectado inválido para control nocturno; se usa la fecha de hoy")

    print(f"[INFO] Control nocturno Renca: ventana {desde:%Y-%m-%d} → {hasta:%Y-%m-%d}")

    try:
        _rows, _csv, out_docx, out_pdf = generar_reporte_control_nocturno(
            desde, hasta, umbral=0.0, excel_path=None
        )
    except Exception as e:
        print(f"[ERROR] No se pudo generar el reporte de control nocturno: {e}")
        import traceback
        traceback.print_exc()
        return False

    if not out_pdf.exists() and not out_docx.exists():
        print("[ERROR] No hay PDF ni DOCX generados para control nocturno")
        return False

    periodo_txt = f"{desde:%d/%m/%Y} - {hasta:%d/%m/%Y}"
    info_reporte = {
        "empresa": "Control nocturno Renca",
        "tipo_reporte": "control nocturno",
        "puntos_monitoreo": [],
        "periodo": periodo_txt,
    }

    # Mantener estilo de correo simple (similar a "puntos en cero")
    formato = info_detectada.get("formato_solicitado", "pdf")
    archivo_path = None
    if formato == "word" and out_docx.exists():
        archivo_path = out_docx
    elif out_pdf.exists():
        archivo_path = out_pdf
    elif out_docx.exists():
        archivo_path = out_docx

    if not archivo_path:
        print("[ERROR] No se encontró archivo para adjuntar (control nocturno)")
        return False

    nombre_contacto = contacto.get("nombre_completo", remitente_email or "Estimado/a")
    tratamiento = contacto.get("tratamiento", "Estimado/a")
    despedida = contacto.get("despedida", "Quedo atento a tus comentarios.")
    firma = contacto.get("firma", "Sistema WES")

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = remitente_email
    msg["Subject"] = "Reporte de Control Nocturno - Sistema WES"

    cuerpo = f"""
:{tratamiento} {nombre_contacto},

Se adjunta el reporte de control nocturno (Renca) del período {desde:%d-%m-%Y} a {hasta:%d-%m-%Y}.

El reporte incluye:
- Resumen ejecutivo con estadísticas generales
- Tabla de puntos fuera de control
- Nota metodológica sobre el criterio de "fuera de control"

{despedida}

{firma}
"""
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    with open(archivo_path, "rb") as f:
        adjunto = MIMEApplication(f.read())
        adjunto.add_header(
            "Content-Disposition",
            "attachment",
            filename=Path(archivo_path).name,
        )
        msg.attach(adjunto)

    try:
        print(f"[INFO] Enviando correo a {remitente_email} (control nocturno)...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg)
        print(f"[OK] Correo enviado exitosamente a {remitente_email}")
        return True
    except Exception as e:
        print(f"[ERROR] Error al generar/enviar correo de control nocturno: {e}")
        import traceback
        traceback.print_exc()
        return False


def procesar_solicitud_reporte(info_detectada):
    """
    Procesa una solicitud de reporte detectada en un correo.
    Esta función se llamará cuando se detecte una solicitud.
    SOLO procesa si el remitente está autorizado.
    Usa la configuración del correo autorizado para determinar qué reportes generar.
    Si el correo es una respuesta, combina parámetros con la solicitud original.
    """
    remitente_email = info_detectada.get("remitente_email")
    contacto = info_detectada.get("remitente")
    empresa_detectada = info_detectada.get("empresa")
    tipo_reporte_detectado = info_detectada.get("tipo_reporte", "ambos")
    asunto_original = info_detectada.get("asunto_original")
    in_reply_to = info_detectada.get("in_reply_to")
    references = info_detectada.get("references")
    message_id_original = info_detectada.get("message_id_original")
    
    # VERIFICAR AUTORIZACIÓN Y OBTENER CONFIGURACIÓN
    config_autorizado = esta_autorizado(remitente_email)
    
    if not config_autorizado:
        print(f"[ADVERTENCIA] El correo {remitente_email} NO está autorizado para recibir reportes")
        print(f"[INFO] La solicitud fue registrada pero NO se procesará")
        print(f"[INFO] Para autorizar este correo, agrégalo a CORREOS_AUTORIZADOS en lista_contactos_reportes.py")
        return False
    
    print(f"[OK] Remitente autorizado detectado: {remitente_email}")
    print(f"[INFO] Se generarán y ENVIARÁN reportes por correo electrónico")
    
    if info_detectada.get("reporte_especial") == "control_nocturno":
        print("[INFO] Ruta: reporte especial — control nocturno (Renca)")
        return _procesar_reporte_control_nocturno_email(info_detectada, contacto)
    
    # DETECTAR SI ES UNA RESPUESTA A UNA SOLICITUD PENDIENTE
    contexto_anterior = None
    message_id_para_eliminar = None
    if in_reply_to or references:
        print(f"[INFO] Detectado correo de respuesta (In-Reply-To o References presente)")
        contexto_anterior = obtener_contexto_solicitud(in_reply_to, references)
        
        if contexto_anterior:
            print(f"[OK] Contexto de solicitud anterior encontrado")
            info_anterior = contexto_anterior.get('info_detectada', {})
            parametros_faltantes_anterior = contexto_anterior.get('parametros_faltantes', [])
            
            print(f"[INFO] Parámetros faltantes en solicitud anterior: {', '.join(parametros_faltantes_anterior)}")
            
            # IMPORTANTE: Si hay contexto anterior, priorizar los parámetros del contexto anterior
            # para los parámetros que faltaban, incluso si el correo actual detecta otros valores.
            # Esto evita que se mezclen empresas cuando el usuario solo responde con el periodo faltante.
            
            # EMPRESA: Lógica mejorada para evitar que se detecte una empresa incorrecta en la respuesta
            empresa_anterior = info_anterior.get('empresa')
            empresa_detectada_en_respuesta = empresa_detectada  # Guardar la empresa detectada en la respuesta antes de modificarla
            
            # Obtener mall del contexto anterior PRIMERO (antes de usarlo)
            mall_anterior = info_anterior.get('mall_detectado')
            
            # PRIORIDAD 1: Si hay mall detectado (en respuesta o contexto anterior), asignar Parque Arauco
            mall_en_respuesta = info_detectada.get('mall_detectado')
            if mall_anterior or mall_en_respuesta:
                empresa_detectada = "Parque Arauco"
                info_detectada['empresa'] = empresa_detectada
                if mall_anterior:
                    print(f"[INFO] Empresa asignada: Parque Arauco (por mall del contexto anterior: {mall_anterior.get('nombre_completo')})")
                elif mall_en_respuesta:
                    print(f"[INFO] Empresa asignada: Parque Arauco (por mall detectado en respuesta: {mall_en_respuesta.get('nombre_completo')})")
            elif empresa_anterior:
                if "empresa" in parametros_faltantes_anterior:
                    # Si "empresa" estaba en los faltantes, el usuario podría estar corrigiendo la empresa
                    # Pero si hay una empresa en el contexto anterior, usarla (probablemente se detectó pero faltaba confirmación)
                    empresa_detectada = empresa_anterior
                    info_detectada['empresa'] = empresa_detectada
                    print(f"[INFO] Empresa obtenida del contexto anterior (era parámetro faltante): {empresa_detectada}")
                else:
                    # Si "empresa" NO estaba en los faltantes, significa que se detectó correctamente en el primer correo
                    # SIEMPRE priorizar la empresa del contexto anterior, ignorando cualquier empresa detectada en la respuesta
                    # Esto evita que se detecte incorrectamente otra empresa (ej: BUPA) cuando el usuario solo responde con el periodo
                    if empresa_detectada_en_respuesta and empresa_detectada_en_respuesta != empresa_anterior:
                        print(f"[ADVERTENCIA] Empresa detectada en respuesta ({empresa_detectada_en_respuesta}) ignorada. Usando empresa del contexto anterior: {empresa_anterior}")
                    empresa_detectada = empresa_anterior
                    info_detectada['empresa'] = empresa_detectada
                    print(f"[INFO] Empresa obtenida del contexto anterior: {empresa_anterior}")
            elif not empresa_detectada and empresa_anterior:
                # Si no se detectó empresa en la respuesta pero hay una en el contexto anterior, usarla
                empresa_detectada = empresa_anterior
                info_detectada['empresa'] = empresa_detectada
                print(f"[INFO] Empresa obtenida del contexto anterior (no detectada en respuesta): {empresa_anterior}")
            
            # PERIODO: Si "periodo" estaba en los parámetros faltantes, usar el del correo actual (respuesta)
            # Si no estaba en los faltantes pero tampoco se detectó en el correo actual, usar el anterior
            if "periodo" in parametros_faltantes_anterior:
                # El periodo debe venir del correo actual (la respuesta)
                if info_detectada.get('periodo'):
                    print(f"[INFO] Periodo obtenido del correo de respuesta: {info_detectada['periodo']}")
                else:
                    # Si no se detectó en la respuesta pero estaba faltante, intentar detectarlo del cuerpo actual
                    # Usar el cuerpo completo del correo actual para detectar el período
                    asunto_actual = info_detectada.get('asunto_original', '')
                    cuerpo_actual = info_detectada.get('cuerpo_original', '')
                    texto_respuesta_completo = f"{asunto_actual} {cuerpo_actual}".lower()
                    periodo_respuesta = detectar_periodo(texto_respuesta_completo)
                    if periodo_respuesta:
                        info_detectada['periodo'] = periodo_respuesta
                        print(f"[INFO] Periodo detectado del correo de respuesta: {periodo_respuesta}")
                    else:
                        # Si aún no se detecta, mantener el anterior como fallback
                        if info_anterior.get('periodo'):
                            info_detectada['periodo'] = info_anterior.get('periodo')
                            print(f"[ADVERTENCIA] Periodo no detectado en respuesta, usando del contexto anterior: {info_detectada['periodo']}")
            elif not info_detectada.get('periodo') and info_anterior.get('periodo'):
                info_detectada['periodo'] = info_anterior.get('periodo')
                print(f"[INFO] Periodo obtenido del contexto anterior: {info_detectada['periodo']}")
            
            # PUNTOS: Si "punto" estaba en los parámetros faltantes, usar los del correo actual (respuesta)
            # Si no estaba en los faltantes pero tampoco se detectó en el correo actual, usar los anteriores
            if "punto" in parametros_faltantes_anterior:
                # Los puntos deben venir del correo actual (la respuesta)
                if info_detectada.get('puntos'):
                    print(f"[INFO] Puntos obtenidos del correo de respuesta: {info_detectada['puntos']}")
            elif not info_detectada.get('puntos') and info_anterior.get('puntos'):
                info_detectada['puntos'] = info_anterior.get('puntos')
                print(f"[INFO] Puntos obtenidos del contexto anterior: {info_detectada['puntos']}")
            
            # Actualizar asunto y cuerpo original del contexto anterior (importante para detectar nodos específicos)
            if info_anterior.get('asunto_original'):
                asunto_original = info_anterior.get('asunto_original')
                info_detectada['asunto_original'] = asunto_original
                print(f"[INFO] Asunto original obtenido del contexto anterior: {asunto_original}")
            
            if info_anterior.get('cuerpo_original'):
                info_detectada['cuerpo_original'] = info_anterior.get('cuerpo_original')
                print(f"[INFO] Cuerpo original obtenido del contexto anterior")
            
            # Guardar nodos específicos detectados del contexto anterior si existen
            # Esto es importante porque los nodos específicos se detectan del texto del correo original
            nodos_especificos_anteriores = info_anterior.get('nodos_especificos_detectados')
            if nodos_especificos_anteriores:
                info_detectada['nodos_especificos_detectados'] = nodos_especificos_anteriores
                print(f"[INFO] Nodos específicos obtenidos del contexto anterior: {len(nodos_especificos_anteriores)} nodo(s)")
                for nodo in nodos_especificos_anteriores:
                    print(f"  - {nodo.get('nodeId')}: {nodo.get('name')}")
            
            # NOTA: mall_anterior ya fue obtenido arriba (línea ~2195), solo verificamos si existe aquí
            # Guardar mall detectado del contexto anterior si existe (importante para Parque Arauco)
            if mall_anterior:
                info_detectada['mall_detectado'] = mall_anterior
                print(f"[INFO] Mall detectado del contexto anterior: {mall_anterior.get('nombre_completo')}")
                # Si hay mall pero no hay empresa, asignar Parque Arauco automáticamente
                if not empresa_detectada or empresa_detectada.upper() != "PARQUE ARAUCO":
                    empresa_detectada = "Parque Arauco"
                    info_detectada['empresa'] = empresa_detectada
                    print(f"[INFO] Empresa asignada automáticamente: Parque Arauco (por mall del contexto anterior)")
            
            # Si se detecta un mall en la respuesta actual, también asignar Parque Arauco
            if info_detectada.get('mall_detectado') and not empresa_detectada:
                empresa_detectada = "Parque Arauco"
                info_detectada['empresa'] = empresa_detectada
                print(f"[INFO] Empresa asignada automáticamente: Parque Arauco (por mall detectado en respuesta)")
            
            # Guardar message_id para eliminar contexto después de procesar
            if in_reply_to:
                message_id_para_eliminar = in_reply_to.strip('<>')
            elif references:
                message_ids_refs = references.split()
                if message_ids_refs:
                    message_id_para_eliminar = message_ids_refs[-1].strip('<>')
    
    # VALIDAR PARÁMETROS NECESARIOS (después de combinar con contexto)
    es_valida, parametros_faltantes = validar_parametros_solicitud(info_detectada)
    
    if not es_valida:
        print(f"[ADVERTENCIA] Faltan parámetros en la solicitud: {', '.join(parametros_faltantes)}")
        print(f"[INFO] Enviando correo al remitente solicitando la información faltante...")
        
        # Si no hay contacto, crear uno básico con el email
        if not contacto:
            print(f"[INFO] No se encontró contacto, creando uno básico para {remitente_email}")
            contacto = {
                "email": remitente_email,
                "tratamiento": "Estimado/a",
                "nombre": remitente_email.split("@")[0],
                "apellido": ""
            }
        
        # Si hay contexto anterior, usar el message_id de la respuesta anterior para responder
        message_id_para_respuesta = None
        if contexto_anterior:
            # Buscar el message_id de la respuesta anterior en el contexto
            # El contexto se guarda con el message_id de la respuesta como clave
            # Necesitamos encontrar ese message_id desde in_reply_to o references
            if in_reply_to:
                message_id_para_respuesta = in_reply_to.strip('<>')
            elif references:
                # Tomar el último message-id de references (el más reciente)
                message_ids_refs = references.split()
                if message_ids_refs:
                    message_id_para_respuesta = message_ids_refs[-1].strip('<>')
        
        # Detectar mall y nodos específicos del correo original antes de guardar el contexto
        # Esto es importante para preservar esta información cuando el usuario solo responde con el periodo
        if empresa_detectada and asunto_original:
            try:
                cuerpo_original = info_detectada.get('cuerpo_original', '')
                texto_original = f"{asunto_original} {cuerpo_original}".lower()
                
                # Detectar mall específico de Parque Arauco
                if "parque arauco" in empresa_detectada.lower():
                    mall_detectado = detectar_mall_parque_arauco(texto_original)
                    if mall_detectado:
                        info_detectada['mall_detectado'] = mall_detectado
                        print(f"[INFO] Mall detectado del correo original: {mall_detectado.get('nombre_completo')}")
                
                # Obtener ID de empresa
                empresas_dict = obtener_todas_las_empresas()
                empresa_id = empresas_dict.get(empresa_detectada.upper())
                
                if empresa_id:
                    # Obtener nodos de la empresa
                    import requests
                    url = f"http://104.248.53.141:7001/wes/api/acl-entities/v1/companies/{empresa_id}"
                    response = requests.get(url, timeout=10)
                    if response.status_code == 200:
                        data = response.json()
                        todos_los_nodos = data.get('nodes', [])
                        
                        # Si hay mall detectado, filtrar nodos por mall
                        if info_detectada.get('mall_detectado'):
                            nodos_del_mall = info_detectada['mall_detectado'].get('nodos', [])
                            todos_los_nodos = [node for node in todos_los_nodos if node.get('nodeId') in nodos_del_mall]
                        
                        # Detectar nodos específicos del texto original
                        nodos_especificos_detectados = detectar_nodos_especificos_en_texto(texto_original, todos_los_nodos)
                        
                        if nodos_especificos_detectados:
                            info_detectada['nodos_especificos_detectados'] = nodos_especificos_detectados
                            print(f"[INFO] Nodos específicos detectados del correo original: {len(nodos_especificos_detectados)} nodo(s)")
                            for nodo in nodos_especificos_detectados:
                                print(f"  - {nodo.get('nodeId')}: {nodo.get('name')}")
            except Exception as e:
                print(f"[ADVERTENCIA] No se pudieron detectar nodos específicos antes de guardar contexto: {e}")
        
        resultado_envio = enviar_correo_solicitando_informacion(
            contacto, 
            parametros_faltantes, 
            asunto_original,
            message_id_original=message_id_para_respuesta,
            info_detectada=info_detectada
        )
        
        if resultado_envio:
            print(f"[OK] Correo de solicitud de información enviado correctamente")
            # Si había contexto anterior y ahora se completó, eliminar el contexto
            if contexto_anterior and message_id_para_respuesta:
                eliminar_contexto_solicitud(message_id_para_respuesta)
            return True
        else:
            print(f"[ERROR] No se pudo enviar el correo de solicitud de información")
            return False
    
    print(f"[OK] Correo autorizado: {remitente_email}")
    print(f"[INFO] Configuración del correo autorizado:")
    print(f"  - Puntos/Empresa: {config_autorizado.get('puntos_monitoreo', [])}")
    print(f"  - Periodo: {config_autorizado.get('periodo', 'Automático')}")
    print(f"  - Tipo de reporte: {config_autorizado.get('tipo_reporte', 'ambos')}")
    
    # Obtener información del contacto para personalización
    if not contacto:
        print(f"[ADVERTENCIA] No se encontró información del contacto {remitente_email}")
        # Crear contacto básico si no existe
        contacto = {
            "nombre_completo": info_detectada.get('remitente_email', 'Usuario'),
            "email": info_detectada.get('remitente_email'),
            "tratamiento": "Estimado/a",
            "despedida": "Quedo atento a tus comentarios.",
            "firma": "Tu Agente WES"
        }
    
    # Determinar qué reportes generar basado en la configuración
    puntos_monitoreo = config_autorizado.get("puntos_monitoreo", [])
    periodo = config_autorizado.get("periodo")
    tipo_reporte = config_autorizado.get("tipo_reporte", "ambos")
    empresa_id = config_autorizado.get("empresa_id")
    
    print(f"[INFO] Procesando solicitud de reporte:")
    print(f"  - Remitente: {contacto['nombre_completo']}")
    print(f"  - Puntos/Empresa configurados: {puntos_monitoreo}")
    print(f"  - Periodo: {periodo if periodo else 'Automático (según solicitud o últimos 30 días)'}")
    print(f"  - Tipo de reporte: {tipo_reporte}")
    print(f"  - Empresa ID: {empresa_id if empresa_id else 'A inferir'}")
    
    # Si se detectó una empresa en el correo, verificar si coincide con la configuración
    if empresa_detectada:
        print(f"  - Empresa detectada en correo: {empresa_detectada}")
        # Verificar si la empresa detectada coincide con algún punto configurado
        empresas_configuradas = [p.upper() for p in puntos_monitoreo if isinstance(p, str) and not p.startswith("000")]
        if empresa_detectada.upper() in empresas_configuradas:
            print(f"  [OK] Empresa detectada coincide con la configuración")
        else:
            print(f"  [INFO] Empresa detectada no coincide exactamente, usando configuración del correo")
    
    # GENERAR REPORTES
    print(f"[INFO] Iniciando generación de reportes...")
    
    try:
        # Importar funciones de generación de reportes
        from generar_reporte_word import generate_report, generate_aggregated_report
        import argparse
        from datetime import datetime, timedelta
        
        # Determinar periodo: usar el periodo detectado en el correo (info_detectada)
        periodo_detectado = info_detectada.get('periodo')
        if periodo_detectado and isinstance(periodo_detectado, dict):
            start_date = periodo_detectado.get('inicio')
            end_date = periodo_detectado.get('fin')
            if start_date and end_date:
                print(f"[INFO] Periodo detectado en correo: {start_date} - {end_date}")
            else:
                # Si el periodo detectado no tiene inicio/fin válidos, usar automático
                start_date = '01/12/2025'
                end_date = datetime.now().strftime('%d/%m/%Y')
                print(f"[INFO] Periodo detectado incompleto, usando automático: {start_date} - {end_date}")
        elif periodo and isinstance(periodo, dict):
            # Usar periodo de la configuración del remitente si existe
            start_date = periodo.get('inicio', '01/12/2025')
            end_date = periodo.get('fin', datetime.now().strftime('%d/%m/%Y'))
            print(f"[INFO] Periodo desde configuración: {start_date} - {end_date}")
        else:
            # Periodo automático: últimos 30 días
            from datetime import timedelta
            end_date = datetime.now().strftime('%d/%m/%Y')
            fecha_inicio_obj = datetime.now() - timedelta(days=30)
            start_date = fecha_inicio_obj.strftime('%d/%m/%Y')
            print(f"[INFO] Periodo automático (últimos 30 días): {start_date} - {end_date}")
        
        print(f"[INFO] Periodo final a usar: {start_date} - {end_date}")
        
        # Obtener todas las empresas desde la API (para remitentes autorizados con "Todas")
        print(f"[INFO] Obteniendo lista de empresas desde la API...")
        try:
            empresas_dict = obtener_todas_las_empresas()
            if empresas_dict and len(empresas_dict) > 0:
                print(f"[OK] Empresas disponibles en el sistema: {len(empresas_dict)}")
            else:
                print(f"[ADVERTENCIA] No se obtuvieron empresas desde la API, usando diccionario conocido como respaldo")
                # Usar diccionario conocido como respaldo
                empresas_dict = {
                    "PARQUE ARAUCO": "000025",
                    "Parque Arauco": "000025",
                    "DERCO": "000012",
                    "BUPA": "000029",
                    "FUNDO ZAPALLAR": "000027",
                }
        except Exception as e:
            print(f"[ERROR] Error al obtener empresas desde la API: {e}")
            print(f"[INFO] Usando diccionario conocido como respaldo para continuar")
            import traceback
            traceback.print_exc()
            # Usar diccionario conocido como respaldo para no detener el proceso
            empresas_dict = {
                "PARQUE ARAUCO": "000025",
                "Parque Arauco": "000025",
                "DERCO": "000012",
                "BUPA": "000029",
                "FUNDO ZAPALLAR": "000027",
            }
        
        # Determinar qué empresa procesar PRIMERO
        empresa_a_procesar = None
        
        # Si puntos_monitoreo contiene "Todas", usar la empresa detectada en el correo
        puede_solicitar_cualquier_empresa = "Todas" in puntos_monitoreo or (isinstance(puntos_monitoreo, list) and "Todas" in [p.upper() for p in puntos_monitoreo])
        
        if puede_solicitar_cualquier_empresa:
            if empresa_detectada:
                empresa_a_procesar = empresa_detectada
                print(f"[INFO] Configuración permite todas las empresas. Usando empresa detectada: {empresa_a_procesar}")
            else:
                print(f"[ERROR] Se requiere 'Todas' pero no se detectó ninguna empresa en el correo")
                return False
        else:
            # Usar la primera empresa de la configuración
            empresa_a_procesar = puntos_monitoreo[0] if puntos_monitoreo else empresa_detectada
        
        # Obtener ID de empresa usando la función dinámica
        if not empresa_id:
            print(f"[INFO] Buscando ID de empresa para: {empresa_a_procesar}")
            empresa_id = obtener_empresa_id_por_nombre(empresa_a_procesar, empresas_dict)
            
            if empresa_id:
                print(f"[OK] Empresa encontrada: {empresa_a_procesar} -> {empresa_id}")
            else:
                print(f"[ERROR] No se pudo determinar el ID de empresa para: {empresa_a_procesar}")
                if empresas_dict:
                    print(f"[INFO] Empresas disponibles en el sistema ({len(empresas_dict)}):")
                    # Mostrar solo las primeras 10 para no saturar el log
                    empresas_lista = list(empresas_dict.keys())[:10]
                    for nombre in empresas_lista:
                        print(f"  - {nombre} ({empresas_dict[nombre]})")
                    if len(empresas_dict) > 10:
                        print(f"  ... y {len(empresas_dict) - 10} más")
                return False
        
        print(f"[INFO] Empresa a procesar: {empresa_a_procesar}")
        print(f"[INFO] Empresa ID: {empresa_id}")
        
        # Obtener nodos de la empresa
        nodos_a_procesar = []
        import requests
        try:
            url = f"http://104.248.53.141:7001/wes/api/acl-entities/v1/companies/{empresa_id}"
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                data = response.json()
                todos_los_nodos = data.get('nodes', [])
                
                # Obtener texto COMPLETO del correo para buscar nodos específicos
                # IMPORTANTE: Se usa TODO el texto (asunto + cuerpo) para detectar nodos,
                # por lo que funciona independientemente del orden de los parámetros
                asunto_original = info_detectada.get('asunto_original', '')
                cuerpo_original = info_detectada.get('cuerpo_original', '')
                texto_correo = f"{asunto_original} {cuerpo_original}".lower()
                asunto_lower = asunto_original.lower()
                cuerpo_lower = cuerpo_original.lower()
                
                print(f"[INFO] Analizando texto completo del correo para detectar parámetros (orden independiente)")
                
                # DETECTAR MALL ESPECÍFICO DE PARQUE ARAUCO
                mall_detectado = None
                if empresa_a_procesar and "parque arauco" in empresa_a_procesar.lower():
                    # Buscar mall en el contexto anterior primero
                    if contexto_anterior:
                        info_anterior = contexto_anterior.get('info_detectada', {})
                        mall_anterior = info_anterior.get('mall_detectado')
                        if mall_anterior:
                            mall_detectado = mall_anterior
                            print(f"[INFO] Mall detectado del contexto anterior: {mall_detectado.get('nombre_completo')}")
                    
                    # Si no hay mall en el contexto anterior, detectarlo del texto actual
                    if not mall_detectado:
                        mall_detectado = detectar_mall_parque_arauco(texto_correo)
                        if mall_detectado:
                            print(f"[INFO] Mall de Parque Arauco detectado: {mall_detectado.get('nombre_completo')}")
                            print(f"[INFO] Nodos del mall: {', '.join(mall_detectado.get('nodos', []))}")
                            # Guardar en info_detectada para preservarlo en el contexto
                            info_detectada['mall_detectado'] = mall_detectado
                
                # Detectar si se menciona "todos" o "todos los puntos" - significa todos los puntos de la empresa
                # Buscar tanto en el asunto como en el cuerpo del correo
                palabras_todos = [
                    "todos", 
                    "todos los", 
                    "todos los puntos", 
                    "todos los nodos", 
                    "todos los puntos de", 
                    "todos los puntos del", 
                    "todos los puntos de la",
                    "todos los puntos de monitoreo",
                    "todos los puntos de la empresa",
                    "todos los nodos de",
                    "todos los nodos de la",
                    "todos los nodos de la empresa"
                ]
                
                # Buscar en el asunto
                menciona_todos_asunto = any(palabra in asunto_lower for palabra in palabras_todos)
                # Buscar en el cuerpo
                menciona_todos_cuerpo = any(palabra in cuerpo_lower for palabra in palabras_todos)
                # Buscar en el texto combinado (por compatibilidad)
                menciona_todos_texto = any(palabra in texto_correo for palabra in palabras_todos)
                
                menciona_todos = menciona_todos_asunto or menciona_todos_cuerpo or menciona_todos_texto
                
                # Si se detectó un mall específico de Parque Arauco, filtrar nodos por mall
                if mall_detectado:
                    nodos_del_mall = mall_detectado.get('nodos', [])
                    print(f"[INFO] Filtrando nodos por mall: {mall_detectado.get('nombre_completo')}")
                    print(f"[INFO] Nodos del mall: {', '.join(nodos_del_mall)}")
                    
                    # Filtrar todos_los_nodos para incluir solo los del mall
                    todos_los_nodos = [node for node in todos_los_nodos if node.get('nodeId') in nodos_del_mall]
                    print(f"[INFO] Nodos filtrados: {len(todos_los_nodos)} nodo(s) del mall {mall_detectado.get('nombre_completo')}")
                
                if menciona_todos:
                    donde_se_detecto = []
                    if menciona_todos_asunto:
                        donde_se_detecto.append("asunto")
                    if menciona_todos_cuerpo:
                        donde_se_detecto.append("cuerpo")
                    ubicacion = " y ".join(donde_se_detecto) if donde_se_detecto else "correo"
                    print(f"[INFO] Se detectó 'todos' en el {ubicacion} del correo. Se procesarán todos los puntos de {empresa_a_procesar}")
                    # Usar todos los nodos (ya filtrados por mall si aplica)
                    for node in todos_los_nodos:
                        nodos_a_procesar.append(node.get('nodeId'))
                    print(f"[OK] Se procesarán todos los {len(nodos_a_procesar)} nodo(s) para {empresa_a_procesar}")
                else:
                    # Detectar si se menciona un nodo específico en el correo
                    # Esta funcionalidad funciona para TODAS las empresas
                    nodos_especificos_mencionados = []
                    
                    # Palabras clave comunes que indican un nodo específico
                    palabras_filtro = ["solo", "solamente", "únicamente", "solo el", "solo la", "solo los", "solo las", "solo de", "punto", "puntos"]
                    
                    # Normalizar nombres: quitar acentos y caracteres especiales para comparación
                    import unicodedata
                    def normalizar(texto):
                        texto = unicodedata.normalize('NFD', texto)
                        texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
                        return texto.lower()
                    
                    texto_normalizado = normalizar(texto_correo)
                    
                    # Buscar si hay palabras de filtro que indiquen un nodo específico
                    tiene_filtro = any(palabra in texto_correo for palabra in palabras_filtro)
                    
                    # SIEMPRE buscar nodos específicos mencionados (para todas las empresas)
                    # Primero buscar con palabras de filtro si existen
                    if tiene_filtro:
                        print(f"[INFO] Detectando nodos específicos mencionados en el correo (palabra de filtro detectada)...")
                        
                        # Buscar coincidencias entre nombres de nodos y el texto del correo
                        for node in todos_los_nodos:
                            node_name = node.get('name', '').strip()
                            node_name_lower = node_name.lower()
                            
                            # Buscar coincidencia exacta del nombre completo del nodo
                            if node_name_lower in texto_correo:
                                nodos_especificos_mencionados.append({
                                    'nodeId': node.get('nodeId'),
                                    'name': node_name,
                                    'match': 'nombre completo'
                                })
                                print(f"[OK] Nodo específico detectado (coincidencia exacta): {node.get('nodeId')} - {node_name}")
                                continue
                            
                            # Si no hay coincidencia exacta, buscar por palabras significativas
                            palabras_nodo = [p for p in node_name_lower.split() if len(p) > 3 and p not in ['del', 'de', 'la', 'el', 'los', 'las', 'y', 'en', 'con', 'para', 'por']]
                            
                            # Si todas las palabras significativas del nodo aparecen en el correo
                            if palabras_nodo and all(palabra in texto_correo for palabra in palabras_nodo):
                                nodos_especificos_mencionados.append({
                                    'nodeId': node.get('nodeId'),
                                    'name': node_name,
                                    'match': 'palabras significativas'
                                })
                                print(f"[OK] Nodo específico detectado (palabras significativas): {node.get('nodeId')} - {node_name}")
                                continue
                            
                            # Buscar si alguna palabra clave del nodo aparece cerca de palabras de filtro
                            for palabra_filtro in palabras_filtro:
                                if palabra_filtro in texto_correo:
                                    # Buscar el contexto después de la palabra de filtro
                                    indice_filtro = texto_correo.find(palabra_filtro)
                                    contexto = texto_correo[indice_filtro:indice_filtro+100]  # 100 caracteres después del filtro
                                    
                                    # Buscar palabras del nodo en el contexto
                                    palabras_encontradas = [p for p in palabras_nodo if p in contexto]
                                    if len(palabras_encontradas) >= 2 or (len(palabras_nodo) == 1 and palabras_encontradas):  # Al menos 2 palabras o todas si solo hay 1
                                        nodos_especificos_mencionados.append({
                                            'nodeId': node.get('nodeId'),
                                            'name': node_name,
                                            'match': f'cerca de "{palabra_filtro}"'
                                        })
                                        print(f"[OK] Nodo específico detectado (cerca de filtro): {node.get('nodeId')} - {node_name}")
                                        break
                    
                    # Si no se detectaron nodos específicos pero el texto es muy corto, intentar buscar de todas formas
                    if not nodos_especificos_mencionados and len(texto_correo.split()) < 50:
                        print(f"[INFO] Texto corto detectado, buscando nodos específicos...")
                        for node in todos_los_nodos:
                            node_name = node.get('name', '').strip()
                            node_name_lower = node_name.lower()
                            
                            # Buscar coincidencia exacta del nombre completo
                            if node_name_lower in texto_correo:
                                nodos_especificos_mencionados.append({
                                    'nodeId': node.get('nodeId'),
                                    'name': node_name,
                                    'match': 'coincidencia exacta en texto corto'
                                })
                                print(f"[OK] Nodo específico detectado: {node.get('nodeId')} - {node_name}")
                    
                    # Búsqueda adicional: buscar nombres de nodos que aparezcan en el correo sin palabras de filtro
                    # Esto ayuda cuando el usuario menciona directamente el nombre del nodo sin usar "solo"
                    # Esta búsqueda funciona para TODAS las empresas
                    if not nodos_especificos_mencionados:
                        print(f"[INFO] Buscando nombres de nodos mencionados directamente en el correo (para todas las empresas)...")
                        for node in todos_los_nodos:
                            node_name = node.get('name', '').strip()
                            node_name_lower = node_name.lower()
                            nombre_normalizado = normalizar(node_name_lower)
                            
                            # Buscar coincidencia exacta del nombre normalizado
                            if nombre_normalizado in texto_normalizado:
                                nodos_especificos_mencionados.append({
                                    'nodeId': node.get('nodeId'),
                                    'name': node_name,
                                    'match': 'coincidencia directa (normalizada)'
                                })
                                print(f"[OK] Nodo específico detectado (mencionado directamente): {node.get('nodeId')} - {node_name}")
                                break  # Si encontramos uno, probablemente es el que se busca
                            
                            # También buscar por palabras clave significativas (al menos 2 palabras del nodo)
                            palabras_nodo = [p for p in nombre_normalizado.split() if len(p) > 3 and p not in ['del', 'de', 'la', 'el', 'los', 'las', 'y', 'en', 'con', 'para', 'por', 'impulsion', 'impulsi', 'matriz', 'principal']]
                            if len(palabras_nodo) >= 2:
                                palabras_encontradas = sum(1 for p in palabras_nodo if p in texto_normalizado)
                                # Si al menos 2 palabras significativas aparecen, considerar que es el nodo
                                if palabras_encontradas >= 2:
                                    nodos_especificos_mencionados.append({
                                        'nodeId': node.get('nodeId'),
                                        'name': node_name,
                                        'match': f'coincidencia por {palabras_encontradas} palabras clave'
                                    })
                                    print(f"[OK] Nodo específico detectado ({palabras_encontradas} palabras clave): {node.get('nodeId')} - {node_name}")
                                    break
                            
                            # Búsqueda especial para nombres cortos o palabras clave únicas (ej: "esval", "bancaria")
                            if len(palabras_nodo) == 1 and palabras_nodo[0] in texto_normalizado:
                                # Verificar que la palabra no sea muy común
                                palabra_unica = palabras_nodo[0]
                                if palabra_unica not in ['norte', 'sur', 'este', 'oeste', 'principal', 'matriz']:
                                    nodos_especificos_mencionados.append({
                                        'nodeId': node.get('nodeId'),
                                        'name': node_name,
                                        'match': f'coincidencia por palabra clave única: {palabra_unica}'
                                    })
                                    print(f"[OK] Nodo específico detectado (palabra clave única '{palabra_unica}'): {node.get('nodeId')} - {node_name}")
                                    break
                
                    # Si se detectaron nodos específicos, usar solo esos (funciona para TODAS las empresas)
                    if nodos_especificos_mencionados:
                        nodos_a_procesar = [n['nodeId'] for n in nodos_especificos_mencionados]
                        print(f"[INFO] Se filtraron {len(nodos_a_procesar)} nodo(s) específico(s) de {len(todos_los_nodos)} disponibles para {empresa_a_procesar}")
                        print(f"[INFO] Nodos a procesar: {', '.join([n['name'] for n in nodos_especificos_mencionados])}")
                    # Si hay nodos específicos del contexto anterior, usarlos (importante cuando el usuario solo responde con el periodo)
                    elif info_detectada.get('nodos_especificos_detectados'):
                        nodos_especificos_anteriores = info_detectada.get('nodos_especificos_detectados')
                        nodos_a_procesar = [n['nodeId'] for n in nodos_especificos_anteriores]
                        print(f"[INFO] Usando nodos específicos del contexto anterior: {len(nodos_a_procesar)} nodo(s)")
                        print(f"[INFO] Nodos a procesar: {', '.join([n['name'] for n in nodos_especificos_anteriores])}")
                    # Si es DERCO y se menciona "matriz principal" (lógica especial de respaldo)
                    elif empresa_a_procesar.upper() == "DERCO" and ("matriz principal" in texto_correo):
                        print(f"[INFO] Filtrando solo nodos 'matriz principal' para DERCO")
                        for node in todos_los_nodos:
                            node_name = node.get('name', '').lower()
                            if 'matriz principal' in node_name:
                                nodos_a_procesar.append(node.get('nodeId'))
                                print(f"[OK] Nodo matriz principal encontrado: {node.get('nodeId')} - {node.get('name')}")
                    else:
                        # Si hay un mall detectado, usar los nodos del mall automáticamente
                        if mall_detectado:
                            # Cuando se detecta un mall pero no se especifican nodos, usar los nodos del mall
                            for node in todos_los_nodos:
                                nodos_a_procesar.append(node.get('nodeId'))
                            print(f"[OK] Mall detectado sin especificación de nodos. Se procesarán todos los {len(nodos_a_procesar)} nodo(s) del mall {mall_detectado.get('nombre_completo')}")
                        else:
                            # Para otras empresas o si no se especifica, usar todos los nodos
                            for node in todos_los_nodos:
                                nodos_a_procesar.append(node.get('nodeId'))
                            print(f"[OK] No se detectaron nodos específicos. Se procesarán todos los {len(nodos_a_procesar)} nodo(s) para {empresa_a_procesar}")
        except Exception as e:
            print(f"[ERROR] No se pudieron obtener nodos de {empresa_a_procesar}: {e}")
            return False
        
        if not nodos_a_procesar:
            print(f"[ERROR] No se encontraron nodos para procesar")
            return False
        
        # Detectar formato solicitado (Word o PDF)
        formato_solicitado = info_detectada.get("formato_solicitado", "pdf")
        print(f"[INFO] Formato solicitado: {formato_solicitado.upper()}")
        
        # Detectar si se solicita presentación
        solicita_presentacion = info_detectada.get("solicita_presentacion", False)
        formato_presentacion = info_detectada.get("formato_presentacion", None)
        
        archivos_pdf = []
        archivos_word = []
        archivos_ppt = []  # Lista para archivos PPT
        archivos_ppt_pdf = []  # Lista para PDFs de PPT
        
        # Generar reportes individuales
        # IMPORTANTE: Cuando se solicita una empresa (especialmente un mall), 
        # SIEMPRE generar reportes individuales para todos los nodos
        # Si se detectó una empresa o se mencionó "todos", siempre generar individuales
        generar_individuales = (
            tipo_reporte in ["individual", "ambos"] or 
            mall_detectado is not None or 
            menciona_todos or
            empresa_detectada is not None  # Si se detectó una empresa, generar individuales
        )
        if generar_individuales and nodos_a_procesar:
            print(f"[INFO] Generando {len(nodos_a_procesar)} reporte(s) individual(es)...")
            for node_id in nodos_a_procesar:
                try:
                    print(f"  Generando reporte para {node_id}...")
                    args = argparse.Namespace(
                        company_id=empresa_id,
                        node_id=node_id,
                        start_date=start_date,
                        end_date=end_date,
                        output_dir="reports",
                        enviar_correo=False,
                        destinatario=None,
                        smtp_servidor=None,
                        smtp_puerto=None,
                        smtp_usuario=None,
                        smtp_password=None
                    )
                    reporte_path = generate_report(args)
                    
                    # Según el formato solicitado, guardar Word o convertir a PDF
                    if formato_solicitado == "word":
                        # Mantener el archivo Word original
                        if reporte_path and os.path.exists(reporte_path):
                            archivos_word.append(reporte_path)
                            print(f"  [OK] Reporte Word generado: {reporte_path}")
                    else:
                        # Convertir a PDF (comportamiento por defecto)
                        from generar_reporte_word import convertir_word_a_pdf
                        pdf_path = convertir_word_a_pdf(reporte_path)
                        if pdf_path:
                            archivos_pdf.append(pdf_path)
                            print(f"  [OK] Reporte PDF generado: {pdf_path}")
                    
                    # NO generar PPT individual - solo se genera presentación agregada
                    # La presentación agregada se genera más abajo y contiene todos los nodos
                except Exception as e:
                    print(f"  [ERROR] Error al generar reporte para {node_id}: {e}")
                    import traceback
                    traceback.print_exc()
        
        # Generar reporte agregado
        # IMPORTANTE: 
        # - Si la empresa tiene solo un nodo, NO generar agregado ni presentación
        # - Cuando se solicita una empresa sin nodo específico, generar TODOS los reportes de la empresa
        # - Solo generar agregado si hay más de 1 nodo
        reporte_agregado = None  # Inicializar variable
        if len(nodos_a_procesar) == 1:
            print(f"[INFO] La empresa tiene solo 1 nodo. Generando solo reporte individual (NO agregado, NO presentación)")
            generar_agregado = False
        else:
            # Generar agregado si:
            # 1. Se solicita explícitamente (tipo_reporte incluye "agregado" o "ambos"), O
            # 2. Se detectó una empresa/mall (mall_detectado, menciona_todos, o empresa_detectada), O
            # 3. Se mencionó "todos" los puntos
            # IMPORTANTE: Cuando se solicita una empresa sin nodo específico, generar TODOS los reportes (individuales + agregado)
            generar_agregado = (
                tipo_reporte in ["agregado", "ambos"] or 
                mall_detectado is not None or 
                menciona_todos or
                empresa_detectada is not None  # Si se detectó una empresa, generar agregado
            ) and len(nodos_a_procesar) > 1
            
            if generar_agregado:
                print(f"[INFO] Se generará reporte agregado para {len(nodos_a_procesar)} nodo(s)")
        
        if generar_agregado and nodos_a_procesar:
            print(f"[INFO] Generando reporte agregado para {len(nodos_a_procesar)} nodo(s)...")
            try:
                from generar_reporte_word import get_company_name
                company_name = get_company_name(empresa_id)
                
                reporte_agregado = generate_aggregated_report(
                    company_id=empresa_id,
                    node_ids=nodos_a_procesar,
                    start_date=start_date,
                    end_date=end_date
                )
                
                if reporte_agregado:
                    # Según el formato solicitado, guardar Word o convertir a PDF
                    if formato_solicitado == "word":
                        # Mantener el archivo Word original
                        if os.path.exists(reporte_agregado):
                            archivos_word.append(reporte_agregado)
                            print(f"  [OK] Reporte agregado Word generado: {reporte_agregado}")
                    else:
                        # Convertir a PDF (comportamiento por defecto)
                        from generar_reporte_word import convertir_word_a_pdf
                        pdf_path = convertir_word_a_pdf(reporte_agregado)
                        if pdf_path:
                            archivos_pdf.append(pdf_path)
                            print(f"  [OK] Reporte agregado PDF generado: {pdf_path}")
                    
                    # Generar PPT agregado SOLO si se solicita explícitamente la presentación
                    # IMPORTANTE: Solo generar PPT si el usuario lo solicita explícitamente
                    # Si la empresa tiene solo 1 nodo, NO generar presentación (ya se validó arriba)
                    if solicita_presentacion and reporte_agregado and len(nodos_a_procesar) > 1:
                        try:
                            from generar_reportes_y_ppt_mall_maipu import generar_ppt_desde_agregado
                            from generar_reporte_word import get_company_name
                            company_name = get_company_name(empresa_id)
                            
                            ppt_path = generar_ppt_desde_agregado(
                                company_id=empresa_id,
                                node_ids=nodos_a_procesar,
                                start_date=start_date,
                                end_date=end_date,
                                aggregated_report_path=reporte_agregado,
                                company_name=company_name
                            )
                            
                            if ppt_path:
                                print(f"  [OK] Presentación PPT agregada generada: {ppt_path}")
                                
                                # Según el formato solicitado de presentación
                                if formato_presentacion == "ppt":
                                    # Adjuntar PPT directamente
                                    if os.path.exists(ppt_path):
                                        archivos_ppt.append(ppt_path)
                                        print(f"  [OK] PPT agregado listo para adjuntar: {ppt_path}")
                                elif formato_presentacion == "pdf":
                                    # Intentar convertir PPT a PDF
                                    from generar_reportes_y_ppt_mall_maipu import convertir_ppt_a_pdf
                                    pdf_ppt_path = convertir_ppt_a_pdf(ppt_path)
                                    if pdf_ppt_path:
                                        archivos_ppt_pdf.append(pdf_ppt_path)
                                        print(f"  [OK] PDF de presentación agregada generado: {pdf_ppt_path}")
                                    else:
                                        # Si no se puede convertir a PDF, adjuntar PPT
                                        if os.path.exists(ppt_path):
                                            archivos_ppt.append(ppt_path)
                                            print(f"  [OK] PPT agregado adjuntado (no se pudo convertir a PDF): {ppt_path}")
                                else:
                                    # Por defecto, intentar PDF
                                    from generar_reportes_y_ppt_mall_maipu import convertir_ppt_a_pdf
                                    pdf_ppt_path = convertir_ppt_a_pdf(ppt_path)
                                    if pdf_ppt_path:
                                        archivos_ppt_pdf.append(pdf_ppt_path)
                                        print(f"  [OK] PDF de presentación agregada generado (por defecto): {pdf_ppt_path}")
                                    else:
                                        # Si no se puede convertir a PDF, adjuntar PPT
                                        if os.path.exists(ppt_path):
                                            archivos_ppt.append(ppt_path)
                                            print(f"  [OK] PPT agregado adjuntado (no se pudo convertir a PDF): {ppt_path}")
                        except Exception as e:
                            print(f"  [ADVERTENCIA] No se pudo generar PPT agregado: {e}")
                            import traceback
                            traceback.print_exc()
            except Exception as e:
                print(f"  [ERROR] Error al generar reporte agregado: {e}")
                import traceback
                traceback.print_exc()
        
        # Preparar archivos para enviar
        # Si se solicita presentación, agregar PPT/PDF según corresponda
        # Si NO se solicita presentación, enviar PDF por defecto junto con reportes
        archivos_a_enviar = archivos_word if formato_solicitado == "word" else archivos_pdf
        tipo_archivo = "Word" if formato_solicitado == "word" else "PDF"
        
        # Si se solicita presentación, agregar archivos PPT/PDF de presentación
        # IMPORTANTE: Solo agregar presentación si el usuario la solicita explícitamente
        if solicita_presentacion:
            if formato_presentacion == "ppt" and archivos_ppt:
                # Agregar PPTs a los archivos a enviar
                archivos_a_enviar.extend(archivos_ppt)
                print(f"[INFO] Se agregaron {len(archivos_ppt)} archivo(s) PPT a los adjuntos")
            elif formato_presentacion == "pdf" and archivos_ppt_pdf:
                # Agregar PDFs de PPT a los archivos a enviar
                archivos_a_enviar.extend(archivos_ppt_pdf)
                print(f"[INFO] Se agregaron {len(archivos_ppt_pdf)} archivo(s) PDF de presentación a los adjuntos")
            elif archivos_ppt_pdf:
                # Por defecto, intentar PDF primero
                archivos_a_enviar.extend(archivos_ppt_pdf)
                print(f"[INFO] Se agregaron {len(archivos_ppt_pdf)} archivo(s) PDF de presentación a los adjuntos")
            elif archivos_ppt:
                # Si no hay PDF, usar PPT
                archivos_a_enviar.extend(archivos_ppt)
                print(f"[INFO] Se agregaron {len(archivos_ppt)} archivo(s) PPT a los adjuntos")
        # IMPORTANTE: NO incluir presentación si NO se solicita explícitamente
        # La presentación SOLO se genera si el usuario la solicita explícitamente en el correo
        
        if archivos_a_enviar:
            print(f"[INFO] Preparando envío de {len(archivos_a_enviar)} reporte(s) en formato {tipo_archivo} por correo...")
            print(f"[INFO] Destinatario: {contacto.get('nombre_completo', contacto.get('email', 'N/A'))} ({contacto.get('email', 'N/A')})")
            info_reporte = {
                'empresa': empresa_detectada or puntos_monitoreo[0] if puntos_monitoreo else "Empresa",
                'periodo': f"{start_date} - {end_date}",
                'tipo': tipo_reporte,
                'formato': formato_solicitado
            }
            # Determinar qué archivos enviar según formato
            archivos_word_a_enviar = [a for a in archivos_a_enviar if str(a).endswith('.docx')]
            archivos_pdf_a_enviar = [a for a in archivos_a_enviar if str(a).endswith('.pdf')]
            archivos_ppt_a_enviar = [a for a in archivos_a_enviar if str(a).endswith('.pptx')]
            
            # Si se solicita presentación, incluir PPT/PDF según corresponda
            if solicita_presentacion:
                if formato_presentacion == "ppt" and archivos_ppt_a_enviar:
                    # Agregar PPTs a los archivos PDF si es PDF, o a Word si es Word
                    if formato_solicitado == "word":
                        archivos_word_a_enviar.extend(archivos_ppt_a_enviar)
                    else:
                        archivos_pdf_a_enviar.extend(archivos_ppt_a_enviar)
                elif formato_presentacion == "pdf" and archivos_pdf_a_enviar:
                    # Los PDFs de PPT ya están en archivos_pdf_a_enviar
                    pass
            
            # Enviar correo
            if formato_solicitado == "word":
                if enviar_correo_personalizado(contacto, info_reporte, archivos_word=archivos_word_a_enviar, archivos_ppt=archivos_ppt_a_enviar if solicita_presentacion else None):
                    print(f"[OK] Reportes Word enviados correctamente")
                    if solicita_presentacion and archivos_ppt_a_enviar:
                        print(f"[OK] Presentaciones PPT incluidas")
                    # Si había contexto anterior (respuesta a solicitud pendiente), eliminarlo
                    if contexto_anterior and message_id_para_eliminar:
                        eliminar_contexto_solicitud(message_id_para_eliminar)
                    return True
                else:
                    print(f"[ERROR] No se pudieron enviar los reportes Word")
                    return False
            else:
                if enviar_correo_personalizado(contacto, info_reporte, archivos_pdf=archivos_pdf_a_enviar, archivos_ppt=archivos_ppt_a_enviar if solicita_presentacion and formato_presentacion == "ppt" else None):
                    print(f"[OK] Reportes PDF enviados correctamente")
                    if solicita_presentacion:
                        if formato_presentacion == "ppt" and archivos_ppt_a_enviar:
                            print(f"[OK] Presentaciones PPT incluidas")
                        elif formato_presentacion == "pdf":
                            print(f"[OK] Presentaciones PDF incluidas")
                    # Si había contexto anterior (respuesta a solicitud pendiente), eliminarlo
                    if contexto_anterior and message_id_para_eliminar:
                        eliminar_contexto_solicitud(message_id_para_eliminar)
                    return True
                else:
                    print(f"[ERROR] No se pudieron enviar los reportes PDF")
                    return False
        else:
            print(f"[ADVERTENCIA] No se generaron reportes")
            return False
        
    except Exception as e:
        print(f"[ERROR] Error al generar reportes: {e}")
        import traceback
        traceback.print_exc()
        return False


def monitorear_y_procesar_correos(service=None):
    """
    Función principal que monitorea correos y procesa solicitudes de reportes.
    Si no hay OAuth (credentials_drive.json + token), usa IMAP con SMTP_USUARIO + contraseña de aplicación.
    """
    print("=" * 70)
    print("MONITOREO DE CORREOS Y GENERACIÓN DE REPORTES")
    print("=" * 70)
    print()

    if excel_abierto():
        print("[PAUSA] Excel detectado abierto: no se procesa este ciclo para evitar conflicto.")
        return service
    
    if service is None:
        print("Autenticando con Gmail API...")
        try:
            service = obtener_servicio_gmail()
            if service:
                print("  [OK] Autenticación exitosa")
        except Exception as e:
            print(f"  [ERROR] Error en autenticación: {e}")
            service = None
    else:
        print("  [INFO] Reutilizando sesión Gmail API")
    
    print()
    
    if not service:
        print("  [ADVERTENCIA] No se pudo obtener servicio de Gmail. Intentando fallback IMAP...")
        return monitorear_y_procesar_correos_imap()
    
    # Obtener correos no leídos
    # Limpiar contextos caducados (más de 1 día)
    print("Limpiando contextos caducados...")
    contextos_eliminados = limpiar_contextos_caducados(service, dias_limite=1)
    if contextos_eliminados > 0:
        print(f"  [OK] {contextos_eliminados} contexto(s) caducado(s) eliminado(s)")
    else:
        print("  [INFO] No hay contextos caducados")
    print()
    
    print("Buscando correos no leídos...")
    mensajes = obtener_correos_no_leidos(service, max_results=10)
    
    if not mensajes:
        print("  [INFO] No hay correos no leídos")
        return service
    
    print(f"  [INFO] Se encontraron {len(mensajes)} correo(s) no leído(s)")
    print()
    
    # Procesar cada correo (ordenados por fecha, más recientes primero)
    solicitudes_procesadas = 0
    correos_procesados_exitosamente = 0
    
    for i, mensaje in enumerate(mensajes, 1):
        msg_id = mensaje['id']
        print(f"[{i}/{len(mensajes)}] Procesando correo {msg_id}...")
        
        # Obtener cuerpo del correo
        resultado_correo = obtener_cuerpo_correo(service, msg_id)
        if not resultado_correo or not resultado_correo[0]:
            print("  [ERROR] No se pudo obtener el contenido del correo")
            continue
        
        asunto, cuerpo, remitente_email, remitente_nombre, message_id_header, in_reply_to, references = resultado_correo
        
        print(f"  Asunto: {asunto}")
        print(f"  Remitente: {remitente_email}")
        
        # Verificar si el remitente está autorizado ANTES de analizar
        config_autorizado = esta_autorizado(remitente_email)
        if not config_autorizado:
            print("  [INFO] Remitente no autorizado, saltando correo")
            # Marcar como leído para no procesarlo de nuevo
            marcar_como_leido(service, msg_id)
            print()
            continue
        
        # Analizar si es una solicitud de reporte
        info_detectada = analizar_correo(asunto, cuerpo, remitente_email)
        
        if info_detectada and info_detectada.get("es_solicitud"):
            print("  [OK] Solicitud de reporte detectada")
            
            # Agregar información de headers del correo para detectar respuestas
            info_detectada['message_id_original'] = message_id_header
            info_detectada['gmail_msg_id'] = msg_id  # ID del mensaje en Gmail API
            info_detectada['in_reply_to'] = in_reply_to
            info_detectada['references'] = references
            
            # Procesar la solicitud
            if procesar_solicitud_reporte(info_detectada):
                print("  [OK] Reporte procesado y correo enviado")
                solicitudes_procesadas += 1
                correos_procesados_exitosamente += 1
                
                # Marcar correo como leído INMEDIATAMENTE después de procesarlo exitosamente
                if marcar_como_leido(service, msg_id):
                    print("  [OK] Correo marcado como leído")
                else:
                    print("  [ADVERTENCIA] No se pudo marcar el correo como leído, pero el reporte fue enviado")
                
                # Continuar procesando otros correos (no hacer break)
                # Esto permite procesar múltiples solicitudes del mismo remitente
                print("  [INFO] Continuando con el procesamiento de correos adicionales...")
            else:
                print("  [ERROR] No se pudo procesar la solicitud")
                # No marcar como leído si falló, para reintentarlo después
        else:
            print("  [INFO] No es una solicitud de reporte")
            # Marcar como leído si no es una solicitud
            marcar_como_leido(service, msg_id)
        
        print()
    
    # Resumen
    timestamp_fin = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print("=" * 70)
    print("PROCESO COMPLETADO")
    print("=" * 70)
    print(f"  - Correos procesados: {len(mensajes)}")
    print(f"  - Solicitudes de reporte procesadas: {solicitudes_procesadas}")
    print(f"  - Última revisión: {timestamp_fin}")
    print()
    
    return service


def monitoreo_continuo(intervalo_minutos=5):
    """
    Función que ejecuta el monitoreo en bucle continuo.
    
    Args:
        intervalo_minutos: Minutos de espera entre cada revisión (default: 5)
    """
    import time
    
    print("=" * 70)
    print("INICIANDO MONITOREO CONTINUO DE CORREOS")
    print("=" * 70)
    print(f"Intervalo de revisión: {intervalo_minutos} minutos")
    print(f"[INFO] Los logs se guardan en: {LOG_FILE}")
    print("Presiona Ctrl+C para detener")
    print()
    
    service = None
    ciclo = 0
    
    try:
        while True:
            ciclo += 1
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"\n[{timestamp}] Ciclo #{ciclo} - Revisando correos...")
            print("-" * 70)
            
            try:
                # Pausa operativa si el usuario está editando el Excel
                while excel_abierto():
                    print(f"  [PAUSA] Excel abierto: esperando {TIEMPO_POLL_EXCEL_SEG}s para no generar reportes...")
                    time.sleep(TIEMPO_POLL_EXCEL_SEG)

                # Gmail API: reutilizar sesión entre ciclos si existe; si no hay OAuth, monitorear usa IMAP.
                if not service:
                    service = obtener_servicio_gmail()

                service = monitorear_y_procesar_correos(service)
                
                # Esperar antes de la próxima revisión
                timestamp_espera = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                print(f"[INFO] Última revisión completada: {timestamp_espera}")
                print(f"[INFO] Esperando {intervalo_minutos} minutos antes de la próxima revisión...")
                time.sleep(intervalo_minutos * 60)
                
            except KeyboardInterrupt:
                print("\n[INFO] Deteniendo monitoreo por solicitud del usuario...")
                break
            except Exception as e:
                print(f"[ERROR] Error en el ciclo #{ciclo}: {e}")
                import traceback
                traceback.print_exc()
                print(f"[INFO] Reintentando en {intervalo_minutos} minutos...")
                service = None  # Forzar nueva autenticación
                time.sleep(intervalo_minutos * 60)
    
    except KeyboardInterrupt:
        print("\n[INFO] Monitoreo detenido")
    except Exception as e:
        print(f"\n[ERROR] Error crítico: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Cerrar el archivo de log
        if 'log_file_handle' in globals():
            log_file_handle.close()
        # Restaurar stdout original
        sys.stdout = original_stdout


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Monitoreo de correos y generación de reportes")
    parser.add_argument(
        "--continuo",
        action="store_true",
        help="Ejecutar en modo continuo (bucle infinito)"
    )
    parser.add_argument(
        "--intervalo",
        type=int,
        default=5,
        help="Intervalo en minutos entre revisiones (default: 5)"
    )
    
    args = parser.parse_args()
    
    if args.continuo:
        # Modo continuo para servicio de Windows
        monitoreo_continuo(intervalo_minutos=args.intervalo)
    else:
        # Modo único (una sola ejecución)
        monitorear_y_procesar_correos()

