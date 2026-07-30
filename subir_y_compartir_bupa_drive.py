"""Script para subir reportes de BUPA a Google Drive, crear carpeta Bupa, compartirla y enviar correo con link."""

import os
import sys
from pathlib import Path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import pickle
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, make_msgid
import re

# Importar configuración de correos
from config_correos_equipo import (
    CORREOS_EQUIPO_WES,
    obtener_email,
    obtener_nombre
)

# Importar funciones para obtener nombres
from generar_reporte_word import get_node_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Scopes necesarios para Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# Rutas de archivos
CREDENTIALS_DIR = Path(r'C:\Users\joseo\Desktop\WES\2026\Agente Derco')
CREDENTIALS_FILE = CREDENTIALS_DIR / 'credentials_drive.json'
TOKEN_FILE = CREDENTIALS_DIR / 'token_drive.pickle'
REPORTS_DIR = Path('reports/BUPA')

# ID de la carpeta raíz de Google Drive
ROOT_FOLDER_ID = 'root'

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587


def obtener_servicio_drive():
    """Obtiene el servicio de Google Drive autenticado."""
    creds = None
    
    # Cargar token si existe
    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, 'rb') as token:
                creds = pickle.load(token)
        except Exception as e:
            print(f"  [ADVERTENCIA] Error al cargar token: {e}")
            creds = None
    
    # Si no hay credenciales válidas, solicitar autorización
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"  [ADVERTENCIA] Error al refrescar token: {e}")
                creds = None
        
        if not creds or not creds.valid:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"[ERROR] No se encontró el archivo de credenciales:")
                print(f"  {CREDENTIALS_FILE}")
                sys.exit(1)
            
            print("  [INFO] Iniciando flujo de autenticación OAuth2...")
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_console()
        
        # Guardar credenciales para la próxima vez
        try:
            with open(TOKEN_FILE, 'wb') as token:
                pickle.dump(creds, token)
        except Exception as e:
            print(f"  [ADVERTENCIA] No se pudo guardar el token: {e}")
    
    return build('drive', 'v3', credentials=creds)


def crear_carpeta(service, nombre_carpeta, parent_id='root'):
    """Crea una carpeta en Google Drive y retorna su ID y link."""
    # Verificar si la carpeta ya existe
    query = f"name='{nombre_carpeta}' and mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    
    if items:
        folder_id = items[0]['id']
        print(f"  [INFO] La carpeta '{nombre_carpeta}' ya existe (ID: {folder_id})")
    else:
        # Crear la carpeta
        file_metadata = {
            'name': nombre_carpeta,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_id] if parent_id != 'root' else []
        }
        
        folder = service.files().create(
            body=file_metadata,
            fields='id'
        ).execute()
        
        folder_id = folder.get('id')
        print(f"  [OK] Carpeta '{nombre_carpeta}' creada (ID: {folder_id})")
    
    # Obtener link de la carpeta
    folder_link = f"https://drive.google.com/drive/folders/{folder_id}"
    
    return folder_id, folder_link


def compartir_carpeta(service, folder_id, emails, permiso='writer'):
    """Comparte una carpeta en Google Drive con los emails especificados."""
    print(f"  [INFO] Compartiendo carpeta con {len(emails)} usuario(s)...")
    
    for email in emails:
        try:
            permission = {
                'type': 'user',
                'role': permiso,
                'emailAddress': email
            }
            
            service.permissions().create(
                fileId=folder_id,
                body=permission,
                sendNotificationEmail=False  # No enviar notificación automática
            ).execute()
            
            print(f"    [OK] Carpeta compartida con {email} (permiso: {permiso})")
        except Exception as e:
            print(f"    [ERROR] No se pudo compartir con {email}: {e}")


def subir_archivo(service, file_path, folder_id, nombre_archivo=None):
    """Sube un archivo a Google Drive en la carpeta especificada."""
    if not os.path.exists(file_path):
        print(f"  [ERROR] El archivo no existe: {file_path}")
        return None
    
    nombre = nombre_archivo or os.path.basename(file_path)
    
    # Verificar si el archivo ya existe
    query = f"name='{nombre}' and '{folder_id}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    
    if items:
        # Actualizar archivo existente
        file_id = items[0]['id']
        media = MediaFileUpload(file_path, resumable=True)
        file = service.files().update(
            fileId=file_id,
            media_body=media,
            fields='id, name'
        ).execute()
        print(f"  [OK] Archivo actualizado: {nombre}")
        return file.get('id')
    else:
        # Crear nuevo archivo
        file_metadata = {
            'name': nombre,
            'parents': [folder_id]
        }
        media = MediaFileUpload(file_path, resumable=True)
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, name'
        ).execute()
        print(f"  [OK] Archivo subido: {nombre}")
        return file.get('id')


def limpiar_nombre_archivo(nombre):
    """Limpia el nombre para que sea válido en Google Drive."""
    # Reemplazar caracteres problemáticos
    nombre = nombre.replace("/", "_").replace("\\", "_")
    nombre = nombre.replace(":", "_").replace("*", "_")
    nombre = nombre.replace("?", "_").replace('"', "_")
    nombre = nombre.replace("<", "_").replace(">", "_")
    nombre = nombre.replace("|", "_")
    # Eliminar espacios múltiples
    nombre = re.sub(r'\s+', ' ', nombre).strip()
    return nombre


def encontrar_carpetas_reporte_individual(directorio):
    """Encuentra todas las carpetas de reportes individuales con sus archivos."""
    carpetas_reporte = []
    
    reporte_dir = directorio / "REPORTE"
    if reporte_dir.exists():
        for carpeta in reporte_dir.iterdir():
            if carpeta.is_dir():
                # Extraer node_id del nombre de la carpeta o del archivo
                node_id = None
                archivos_en_carpeta = []
                
                # Buscar archivo .docx para extraer node_id
                for archivo in carpeta.glob("Reporte_*.docx"):
                    # Formato: Reporte_000029_000029-01_20251015_20251215.docx
                    match = re.search(r'Reporte_\d+_(\d+-\d+)_', archivo.name)
                    if match:
                        node_id = match.group(1)
                        break
                
                if node_id:
                    # Obtener nombre del punto
                    nombre_punto = get_node_name(node_id)
                    nombre_carpeta = limpiar_nombre_archivo(nombre_punto)
                    
                    # Recopilar todos los archivos de la carpeta
                    for archivo in carpeta.iterdir():
                        if archivo.is_file() and not archivo.name.startswith("~$"):
                            archivos_en_carpeta.append(archivo)
                    
                    if archivos_en_carpeta:
                        carpetas_reporte.append({
                            'node_id': node_id,
                            'nombre_punto': nombre_punto,
                            'nombre_carpeta': nombre_carpeta,
                            'archivos': archivos_en_carpeta,
                            'carpeta_local': carpeta
                        })
    
    return carpetas_reporte


def encontrar_carpeta_reporte_agregado(directorio):
    """Encuentra la carpeta del reporte agregado con todos sus archivos."""
    agregado_dir = directorio / "ABREGADO"
    if agregado_dir.exists():
        for carpeta in agregado_dir.iterdir():
            if carpeta.is_dir():
                archivos_en_carpeta = []
                
                # Recopilar todos los archivos de la carpeta
                for archivo in carpeta.iterdir():
                    if archivo.is_file() and not archivo.name.startswith("~$"):
                        archivos_en_carpeta.append(archivo)
                
                if archivos_en_carpeta:
                    return {
                        'archivos': archivos_en_carpeta,
                        'carpeta_local': carpeta
                    }
    
    return None


def enviar_correo_con_link(carpeta_link, total_archivos):
    """Envía un correo con el link de la carpeta compartida."""
    # Obtener correos del equipo
    email_anibal = obtener_email("anibal")
    email_diego = obtener_email("diego")
    email_juan = obtener_email("juan")
    email_benjamin = obtener_email("benjamin")
    
    nombre_diego = obtener_nombre("diego")
    
    # Destinatario: Diego (principal), CC: Juan, Benjamín, Aníbal
    to_recipients = [email_diego]
    cc_recipients = [email_juan, email_benjamin, email_anibal]
    
    # Construir correo
    msg = MIMEMultipart('alternative')
    
    # Headers básicos
    msg["From"] = f"Agente WES <{SMTP_USUARIO}>"
    msg["To"] = ", ".join(to_recipients)
    msg["Cc"] = ", ".join(cc_recipients)
    msg["Reply-To"] = SMTP_USUARIO
    msg["Subject"] = "Reportes BUPA - Carpeta compartida en Google Drive"
    
    # Headers importantes para evitar spam
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="wes.cl")
    msg["X-Mailer"] = "WES Report Generator"
    msg["X-Priority"] = "3"
    msg["MIME-Version"] = "1.0"
    
    fecha_generacion = datetime.now().strftime("%d de %B de %Y")
    
    # Cuerpo en texto plano
    cuerpo_texto = f"""Estimado {nombre_diego},

Te informo que se han generado y subido los reportes de BUPA a Google Drive.

ENLACE A LA CARPETA COMPARTIDA:
{carpeta_link}

La carpeta "Bupa" ha sido compartida con:
- {obtener_nombre("anibal")} ({email_anibal})
- {obtener_nombre("diego")} ({email_diego})
- {obtener_nombre("juan")} ({email_juan})
- {obtener_nombre("benjamin")} ({email_benjamin})

CONTENIDO DE LA CARPETA:
- {total_archivos} archivo(s) de reporte(s) subido(s)
- Reportes individuales por punto de monitoreo (cada uno en su propia carpeta)
- Reporte agregado consolidado (en la carpeta "Agregado")

METODOLOGÍA DE OBTENCIÓN DE DATOS:

Los datos fueron obtenidos mediante consultas automatizadas a la API de WES, accediendo a:
- Información de todos los puntos de monitoreo de BUPA
- Datos de consumo horario y diario del periodo analizado (15 de octubre al 15 de diciembre de 2025)
- Alertas y eventos registrados en el sistema
- Métricas de consumo efectivo y consumos nocturnos

El proceso incluyó la recopilación de datos históricos desde el 15 de octubre de 2025 hasta el 15 de diciembre de 2025, analizando el comportamiento de consumo de cada punto de monitoreo.

METODOLOGÍA DE ELABORACIÓN DEL INFORME:

1. Recopilación de datos: Se accedió a la API de WES para obtener:
   - Medidas horarias de consumo de agua
   - Alertas registradas en el periodo
   - Información de nodos y dispositivos

2. Procesamiento y análisis:
   - Cálculo de consumos totales, promedios diarios y mensuales
   - Identificación de consumos nocturnos (22:00-07:00) en los últimos 2 días
   - Proyección de consumos nocturnos basada en las últimas 2 alertas nocturnas
   - Cálculo de consumo efectivo (consumo total menos proyección de consumos nocturnos)

3. Generación de visualizaciones:
   - Gráficos de consumo diario y horario
   - Gráficos comparativos entre puntos
   - Tablas resumen con métricas clave
   - Análisis de ranking de consumo por punto

4. Consolidación:
   - Generación de reportes individuales por punto (cada uno con sus gráficas y documentos)
   - Generación de reporte agregado con análisis comparativo
   - Inclusión de narrativas explicativas y recomendaciones

Este correo y los reportes fueron generados automáticamente por mí, tu agente de IA al servicio de WES.

Quedo atento a tus comentarios y cualquier consulta que puedas tener.

Saludos,
Tu Agente WES
"""
    
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
        .note {{ background-color: #fff3cd; padding: 10px; border-left: 4px solid #ffc107; margin: 15px 0; }}
    </style>
</head>
<body>
    <div class="header">
        <h2>Reportes BUPA - Carpeta compartida en Google Drive</h2>
        <p><strong>Fecha de generación:</strong> {fecha_generacion}</p>
    </div>
    
    <div class="content">
        <p>Estimado <strong>{nombre_diego}</strong>,</p>
        
        <p>Te informo que se han generado y subido los reportes de BUPA a Google Drive.</p>
        
        <div class="link-box">
            <h3>🔗 Enlace a la carpeta compartida:</h3>
            <p><a href="{carpeta_link}" target="_blank">{carpeta_link}</a></p>
        </div>
        
        <div class="section">
            <h3>Usuarios con acceso a la carpeta:</h3>
            <ul>
                <li><strong>{obtener_nombre("anibal")}</strong> ({email_anibal})</li>
                <li><strong>{obtener_nombre("diego")}</strong> ({email_diego})</li>
                <li><strong>{obtener_nombre("juan")}</strong> ({email_juan})</li>
                <li><strong>{obtener_nombre("benjamin")}</strong> ({email_benjamin})</li>
            </ul>
        </div>
        
        <div class="section">
            <h3>Contenido de la carpeta:</h3>
            <ul>
                <li><strong>{total_archivos} archivo(s)</strong> de reporte(s) subido(s)</li>
                <li>Reportes individuales por punto de monitoreo (cada uno en su propia carpeta)</li>
                <li>Reporte agregado consolidado (en la carpeta "Agregado")</li>
            </ul>
        </div>
        
        <div class="section">
            <h3>Metodología de obtención de datos:</h3>
            <p>Los datos fueron obtenidos mediante consultas automatizadas a la API de WES, accediendo a:</p>
            <ul>
                <li>Información de todos los puntos de monitoreo de BUPA</li>
                <li>Datos de consumo horario y diario del periodo analizado</li>
                <li>Alertas y eventos registrados en el sistema</li>
                <li>Métricas de consumo efectivo y consumos nocturnos</li>
            </ul>
            <p>El proceso incluyó la recopilación de datos históricos desde el <strong>15 de octubre de 2025 hasta el 15 de diciembre de 2025</strong>, analizando el comportamiento de consumo de cada punto de monitoreo.</p>
        </div>
        
        <div class="section">
            <h3>Metodología de elaboración del informe:</h3>
            <ol>
                <li><strong>Recopilación de datos:</strong> Se accedió a la API de WES para obtener:
                    <ul>
                        <li>Medidas horarias de consumo de agua</li>
                        <li>Alertas registradas en el periodo</li>
                        <li>Información de nodos y dispositivos</li>
                    </ul>
                </li>
                <li><strong>Procesamiento y análisis:</strong>
                    <ul>
                        <li>Cálculo de consumos totales, promedios diarios y mensuales</li>
                        <li>Identificación de consumos nocturnos (22:00-07:00) en los últimos 2 días</li>
                        <li>Proyección de consumos nocturnos basada en las últimas 2 alertas nocturnas</li>
                        <li>Cálculo de consumo efectivo (consumo total menos proyección de consumos nocturnos)</li>
                    </ul>
                </li>
                <li><strong>Generación de visualizaciones:</strong>
                    <ul>
                        <li>Gráficos de consumo diario y horario</li>
                        <li>Gráficos comparativos entre puntos</li>
                        <li>Tablas resumen con métricas clave</li>
                        <li>Análisis de ranking de consumo por punto</li>
                    </ul>
                </li>
                <li><strong>Consolidación:</strong>
                    <ul>
                        <li>Generación de reportes individuales por punto</li>
                        <li>Generación de reporte agregado con análisis comparativo</li>
                        <li>Inclusión de narrativas explicativas y recomendaciones</li>
                    </ul>
                </li>
            </ol>
        </div>
        
        <p>Este correo y los reportes fueron generados automáticamente por mí, tu <strong>agente de IA al servicio de WES</strong>.</p>
        
        <p>Quedo atento a tus comentarios y cualquier consulta que puedas tener.</p>
    </div>
    
    <div class="footer">
        <p>Saludos,<br>
        <strong>Tu Agente WES</strong></p>
    </div>
</body>
</html>
"""
    
    # Agregar ambas versiones
    part1 = MIMEText(cuerpo_texto, "plain", "utf-8")
    part2 = MIMEText(cuerpo_html, "html", "utf-8")
    
    msg.attach(part1)
    msg.attach(part2)
    
    # Enviar correo
    try:
        print()
        print("=" * 70)
        print("ENVIANDO CORREO")
        print("=" * 70)
        print(f"[INFO] Conectando a {SMTP_SERVIDOR}:{SMTP_PUERTO}...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            print(f"[INFO] Autenticando como {SMTP_USUARIO}...")
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            print(f"[INFO] Enviando correo a: {', '.join(to_recipients)}")
            print(f"[INFO] Con copia a: {', '.join(cc_recipients)}")
            all_recipients = to_recipients + cc_recipients
            server.sendmail(SMTP_USUARIO, all_recipients, msg.as_string())
        print("[OK] Correo enviado correctamente.")
        return True
    except Exception as e:
        print(f"[ERROR] Falló el envío del correo: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    print("=" * 70)
    print("SUBIR REPORTES BUPA A GOOGLE DRIVE Y COMPARTIR")
    print("=" * 70)
    print()
    
    # Verificar que existe el directorio de reportes
    if not REPORTS_DIR.exists():
        print(f"[ERROR] No se encontró el directorio de reportes: {REPORTS_DIR}")
        sys.exit(1)
    
    # Obtener servicio de Google Drive
    print("Autenticando con Google Drive...")
    try:
        service = obtener_servicio_drive()
        print("  [OK] Autenticación exitosa")
    except Exception as e:
        print(f"  [ERROR] Error en autenticación: {e}")
        sys.exit(1)
    
    print()
    
    # Crear carpeta Bupa en la raíz
    print("Creando/verificando carpeta Bupa en Google Drive...")
    try:
        bupa_folder_id, bupa_folder_link = crear_carpeta(service, "Bupa", ROOT_FOLDER_ID)
    except Exception as e:
        print(f"  [ERROR] Error al crear carpeta: {e}")
        sys.exit(1)
    
    print()
    
    # Compartir carpeta con el equipo
    print("Compartiendo carpeta con el equipo...")
    emails_equipo = [
        obtener_email("anibal"),
        obtener_email("diego"),
        obtener_email("juan"),
        obtener_email("benjamin")
    ]
    compartir_carpeta(service, bupa_folder_id, emails_equipo, permiso='writer')
    
    print()
    
    # Encontrar carpetas de reportes individuales
    print("Buscando reportes individuales...")
    carpetas_individuales = encontrar_carpetas_reporte_individual(REPORTS_DIR)
    
    if not carpetas_individuales:
        print("  [ADVERTENCIA] No se encontraron reportes individuales")
    else:
        print(f"  [INFO] Se encontraron {len(carpetas_individuales)} reportes individuales")
    
    print()
    
    # Encontrar carpeta de reporte agregado
    print("Buscando reporte agregado...")
    carpeta_agregado = encontrar_carpeta_reporte_agregado(REPORTS_DIR)
    
    if not carpeta_agregado:
        print("  [ADVERTENCIA] No se encontró el reporte agregado")
    else:
        print(f"  [INFO] Se encontró el reporte agregado con {len(carpeta_agregado['archivos'])} archivos")
    
    print()
    
    # Subir reportes individuales (cada uno en su propia carpeta)
    total_archivos_subidos = 0
    total_archivos_fallidos = 0
    
    if carpetas_individuales:
        print("Subiendo reportes individuales...")
        for carpeta_info in carpetas_individuales:
            nombre_carpeta = carpeta_info['nombre_carpeta']
            print(f"Creando carpeta para: {nombre_carpeta}...")
            
            # Crear carpeta para este punto
            try:
                punto_folder_id, _ = crear_carpeta(service, nombre_carpeta, bupa_folder_id)
                
                # Subir todos los archivos de este punto
                print(f"  Subiendo {len(carpeta_info['archivos'])} archivo(s)...")
                for archivo in carpeta_info['archivos']:
                    nombre_archivo = archivo.name
                    print(f"    Subiendo: {nombre_archivo}...")
                    try:
                        subir_archivo(service, str(archivo), punto_folder_id, nombre_archivo)
                        total_archivos_subidos += 1
                    except Exception as e:
                        print(f"      [ERROR] Error al subir archivo: {e}")
                        total_archivos_fallidos += 1
                
                print()
            except Exception as e:
                print(f"  [ERROR] Error al crear carpeta para {nombre_carpeta}: {e}")
                print()
    
    # Subir reporte agregado (en carpeta "Agregado" dentro de Bupa)
    if carpeta_agregado:
        print("Subiendo reporte agregado...")
        print("Creando carpeta 'Agregado'...")
        
        try:
            agregado_folder_id, _ = crear_carpeta(service, "Agregado", bupa_folder_id)
            
            # Subir todos los archivos del reporte agregado
            print(f"  Subiendo {len(carpeta_agregado['archivos'])} archivo(s)...")
            for archivo in carpeta_agregado['archivos']:
                nombre_archivo = archivo.name
                print(f"    Subiendo: {nombre_archivo}...")
                try:
                    subir_archivo(service, str(archivo), agregado_folder_id, nombre_archivo)
                    total_archivos_subidos += 1
                except Exception as e:
                    print(f"      [ERROR] Error al subir archivo: {e}")
                    total_archivos_fallidos += 1
            
            print()
        except Exception as e:
            print(f"  [ERROR] Error al crear carpeta Agregado: {e}")
            print()
    
    # Enviar correo con link
    print()
    if total_archivos_subidos > 0:
        enviar_correo_con_link(bupa_folder_link, total_archivos_subidos)
    
    # Resumen
    print()
    print("=" * 70)
    print("PROCESO COMPLETADO")
    print("=" * 70)
    print()
    print("RESUMEN:")
    print(f"  - Carpeta creada/compartida: Bupa")
    print(f"  - Link de la carpeta: {bupa_folder_link}")
    print(f"  - Reportes individuales: {len(carpetas_individuales) if carpetas_individuales else 0}")
    print(f"  - Archivos subidos exitosamente: {total_archivos_subidos}")
    if total_archivos_fallidos > 0:
        print(f"  - Archivos fallidos: {total_archivos_fallidos}")
    print(f"  - Carpeta compartida con: {len(emails_equipo)} usuario(s)")
    print()
    print("Los reportes están disponibles en Google Drive en la carpeta 'Bupa'")
    print("  - Cada reporte individual está en su propia carpeta con el nombre del punto")
    print("  - El reporte agregado está en la carpeta 'Agregado' dentro de 'Bupa'")


if __name__ == "__main__":
    main()

