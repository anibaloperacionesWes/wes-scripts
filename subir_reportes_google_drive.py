"""Script para subir reportes de BUPA a Google Drive en la carpeta DERCO."""

import os
import sys
from pathlib import Path
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import pickle

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Scopes necesarios para Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# Rutas de archivos
# Buscar credenciales en la carpeta "Agente Derco"
CREDENTIALS_DIR = Path(r'C:\Users\joseo\Desktop\WES\2026\Agente Derco')
CREDENTIALS_FILE = CREDENTIALS_DIR / 'credentials_drive.json'
TOKEN_FILE = CREDENTIALS_DIR / 'token_drive.pickle'
REPORTS_DIR = Path('reports/BUPA')

# ID de la carpeta raíz de Google Drive (usualmente 'root')
ROOT_FOLDER_ID = 'root'


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
                print()
                print("Por favor, verifica que el archivo 'credentials_drive.json' existe en:")
                print(f"  {CREDENTIALS_DIR}")
                sys.exit(1)
            
            print("  [INFO] Iniciando flujo de autenticación OAuth2...")
            print("  [INFO] Se abrirá una ventana del navegador para autorizar el acceso")
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES)
            # Usar run_console para copiar/pegar código (más compatible)
            creds = flow.run_console()
        
        # Guardar credenciales para la próxima vez
        try:
            with open(TOKEN_FILE, 'wb') as token:
                pickle.dump(creds, token)
        except Exception as e:
            print(f"  [ADVERTENCIA] No se pudo guardar el token: {e}")
    
    return build('drive', 'v3', credentials=creds)


def crear_carpeta(service, nombre_carpeta, parent_id='root'):
    """Crea una carpeta en Google Drive y retorna su ID."""
    # Verificar si la carpeta ya existe
    query = f"name='{nombre_carpeta}' and mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    
    if items:
        print(f"  [INFO] La carpeta '{nombre_carpeta}' ya existe (ID: {items[0]['id']})")
        return items[0]['id']
    
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
    
    print(f"  [OK] Carpeta '{nombre_carpeta}' creada (ID: {folder.get('id')})")
    return folder.get('id')


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
        print(f"  [OK] Archivo actualizado: {nombre} (ID: {file.get('id')})")
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
        print(f"  [OK] Archivo subido: {nombre} (ID: {file.get('id')})")
        return file.get('id')


def encontrar_archivos_reporte(directorio):
    """Encuentra todos los archivos .docx de reportes en el directorio."""
    archivos = []
    
    # Buscar reportes individuales
    reporte_dir = directorio / "REPORTE"
    if reporte_dir.exists():
        for carpeta in reporte_dir.iterdir():
            if carpeta.is_dir():
                for archivo in carpeta.glob("Reporte_*.docx"):
                    archivos.append(archivo)
    
    # Buscar reporte agregado
    agregado_dir = directorio / "ABREGADO"
    if agregado_dir.exists():
        for carpeta in agregado_dir.iterdir():
            if carpeta.is_dir():
                for archivo in carpeta.glob("Reporte_Agregado_*.docx"):
                    archivos.append(archivo)
    
    return archivos


def main():
    print("=" * 70)
    print("SUBIR REPORTES DE BUPA A GOOGLE DRIVE")
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
    
    # Crear carpeta DERCO en la raíz
    print("Creando/verificando carpeta DERCO en Google Drive...")
    try:
        derco_folder_id = crear_carpeta(service, "DERCO", ROOT_FOLDER_ID)
    except Exception as e:
        print(f"  [ERROR] Error al crear carpeta: {e}")
        sys.exit(1)
    
    print()
    
    # Encontrar archivos de reportes
    print("Buscando archivos de reportes...")
    archivos = encontrar_archivos_reporte(REPORTS_DIR)
    
    if not archivos:
        print("  [ADVERTENCIA] No se encontraron archivos de reportes")
        sys.exit(0)
    
    print(f"  [INFO] Se encontraron {len(archivos)} archivos de reportes")
    print()
    
    # Subir archivos
    print("Subiendo archivos a Google Drive...")
    archivos_subidos = 0
    archivos_fallidos = 0
    
    for archivo in archivos:
        nombre_archivo = archivo.name
        print(f"Subiendo: {nombre_archivo}...")
        try:
            subir_archivo(service, str(archivo), derco_folder_id, nombre_archivo)
            archivos_subidos += 1
        except Exception as e:
            print(f"  [ERROR] Error al subir archivo: {e}")
            archivos_fallidos += 1
        print()
    
    # Resumen
    print("=" * 70)
    print("PROCESO COMPLETADO")
    print("=" * 70)
    print()
    print("RESUMEN:")
    print(f"  - Archivos subidos exitosamente: {archivos_subidos}")
    if archivos_fallidos > 0:
        print(f"  - Archivos fallidos: {archivos_fallidos}")
    print()
    print(f"Los reportes están disponibles en Google Drive en la carpeta 'DERCO'")


if __name__ == "__main__":
    main()

