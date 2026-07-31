"""
Sube el reporte de puntos en cero a Google Drive (carpeta «reporte en cero»).

Reutiliza el mismo flujo OAuth/API que ``subir_reportes_google_drive.py``,
con rutas de credenciales portables (``gmail_oauth/`` en el repo).

Uso:
  python subir_reporte_puntos_cero_drive.py
  python subir_reporte_puntos_cero_drive.py --archivo "reporte en cero/Reporte_....docx"
"""

from __future__ import annotations

import argparse
import os
import pickle
import sys
from pathlib import Path

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

from wes_paths import reporte_cero_dir, wes_scripts_root

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

SCOPES = ["https://www.googleapis.com/auth/drive.file"]

_REPO = Path(__file__).resolve().parent
_CANDIDATE_DIRS = [
    _REPO / "gmail_oauth",
    Path(os.environ.get("WES_DRIVE_CREDENTIALS_DIR", "").strip()) if os.environ.get("WES_DRIVE_CREDENTIALS_DIR") else None,
    Path(r"C:\Users\joseo\Desktop\WES\2026\Agente Derco"),
]
_CANDIDATE_DIRS = [d for d in _CANDIDATE_DIRS if d is not None]

FOLDER_CHAIN = ["Agente WES", "wes-scripts", "reporte en cero"]


def _credentials_dir() -> Path:
    for d in _CANDIDATE_DIRS:
        if (d / "credentials_drive.json").exists():
            return d
    return _REPO / "gmail_oauth"


def obtener_servicio_drive():
    """Obtiene el servicio de Google Drive autenticado (OAuth)."""
    cred_dir = _credentials_dir()
    credentials_file = cred_dir / "credentials_drive.json"
    token_file = cred_dir / "token_drive.pickle"

    creds = None
    if token_file.exists():
        try:
            with open(token_file, "rb") as token:
                creds = pickle.load(token)
        except Exception as e:
            print(f"  [ADVERTENCIA] Error al cargar token: {e}")
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and getattr(creds, "refresh_token", None):
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"  [ADVERTENCIA] Error al refrescar token: {e}")
                creds = None

        if not creds or not creds.valid:
            if not credentials_file.exists():
                print("[ERROR] No se encontró credentials_drive.json")
                print(f"  Buscado en: {cred_dir}")
                print()
                print("Copia credentials_drive.json (y token_drive.pickle si existe) a:")
                print(f"  {_REPO / 'gmail_oauth'}")
                print("O define WES_DRIVE_CREDENTIALS_DIR con esa carpeta.")
                sys.exit(1)

            print("  [INFO] Iniciando flujo OAuth2 (consola)...")
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_file), SCOPES)
            # Headless / cloud: run_console (pegar código). En desktop con browser: run_local_server.
            if os.environ.get("WES_DRIVE_OAUTH_LOCAL", "").strip() in {"1", "true", "yes"}:
                creds = flow.run_local_server(port=0)
            else:
                creds = flow.run_console()

        try:
            cred_dir.mkdir(parents=True, exist_ok=True)
            with open(token_file, "wb") as token:
                pickle.dump(creds, token)
        except Exception as e:
            print(f"  [ADVERTENCIA] No se pudo guardar el token: {e}")

    return build("drive", "v3", credentials=creds)


def crear_o_obtener_carpeta(service, nombre: str, parent_id: str = "root") -> str:
    query = (
        f"name='{nombre}' and mimeType='application/vnd.google-apps.folder' "
        f"and '{parent_id}' in parents and trashed=false"
    )
    results = service.files().list(q=query, fields="files(id, name)", pageSize=10).execute()
    items = results.get("files", [])
    if items:
        folder_id = items[0]["id"]
        print(f"  [INFO] Carpeta '{nombre}' ya existe (ID: {folder_id})")
        return folder_id

    meta = {
        "name": nombre,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [] if parent_id == "root" else [parent_id],
    }
    folder = service.files().create(body=meta, fields="id").execute()
    folder_id = folder["id"]
    print(f"  [OK] Carpeta '{nombre}' creada (ID: {folder_id})")
    return folder_id


def asegurar_cadena_carpetas(service, nombres: list[str], parent_id: str = "root") -> str:
    current = parent_id
    for nombre in nombres:
        current = crear_o_obtener_carpeta(service, nombre, current)
    return current


def subir_archivo(service, file_path: Path, folder_id: str, nombre_archivo: str | None = None) -> str | None:
    if not file_path.exists():
        print(f"  [ERROR] El archivo no existe: {file_path}")
        return None

    nombre = nombre_archivo or file_path.name
    query = f"name='{nombre}' and '{folder_id}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id, name)", pageSize=5).execute()
    items = results.get("files", [])
    media = MediaFileUpload(str(file_path), resumable=True)

    if items:
        file_id = items[0]["id"]
        file = service.files().update(fileId=file_id, media_body=media, fields="id, name, webViewLink").execute()
        print(f"  [OK] Archivo actualizado: {nombre}")
        return file.get("id")

    meta = {"name": nombre, "parents": [folder_id]}
    file = service.files().create(body=meta, media_body=media, fields="id, name, webViewLink").execute()
    print(f"  [OK] Archivo subido: {nombre}")
    return file.get("id")


def _resolver_archivo(explicit: Path | None) -> Path:
    if explicit is not None:
        p = explicit if explicit.is_absolute() else (_REPO / explicit)
        if not p.exists():
            raise FileNotFoundError(f"No existe: {p}")
        return p.resolve()

    cero = reporte_cero_dir()
    # Preferir el de hoy (20260731); si no, el .docx más reciente.
    hoy = list(cero.glob("Reporte_Puntos_En_Cero_20260731_*.docx"))
    hoy = [f for f in hoy if f.is_file() and not f.name.startswith("~$")]
    if hoy:
        hoy.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        return hoy[0].resolve()

    archivos = [
        f
        for f in cero.glob("Reporte_Puntos_En_Cero_*.docx")
        if f.is_file() and not f.name.startswith("~$")
    ]
    if not archivos:
        raise FileNotFoundError(f"No hay reportes .docx en {cero}")
    archivos.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return archivos[0].resolve()


def main() -> None:
    parser = argparse.ArgumentParser(description="Subir reporte de puntos en cero a Google Drive")
    parser.add_argument("--archivo", type=Path, default=None, help="Ruta al .docx (o .pdf) a subir")
    parser.add_argument(
        "--solo-carpeta",
        action="store_true",
        help="Crear/usar solo la carpeta «reporte en cero» en la raíz (sin Agente WES/wes-scripts)",
    )
    args = parser.parse_args()

    print("=" * 70)
    print("SUBIR REPORTE DE PUNTOS EN CERO A GOOGLE DRIVE")
    print("=" * 70)
    print(f"Repo / scripts root: {wes_scripts_root()}")
    print()

    archivo = _resolver_archivo(args.archivo)
    print(f"Archivo a subir: {archivo}")
    print()

    print("Autenticando con Google Drive...")
    service = obtener_servicio_drive()
    print("  [OK] Autenticación exitosa")
    print()

    cadena = ["reporte en cero"] if args.solo_carpeta else FOLDER_CHAIN
    print(f"Asegurando carpetas: {' / '.join(cadena)}")
    folder_id = asegurar_cadena_carpetas(service, cadena)
    folder_link = f"https://drive.google.com/drive/folders/{folder_id}"
    print(f"  Carpeta destino: {folder_link}")
    print()

    print("Subiendo archivo...")
    file_id = subir_archivo(service, archivo, folder_id)
    if not file_id:
        sys.exit(1)

    file_link = f"https://drive.google.com/file/d/{file_id}/view"
    print()
    print("=" * 70)
    print("LISTO")
    print("=" * 70)
    print(f"  Archivo: {archivo.name}")
    print(f"  Link archivo: {file_link}")
    print(f"  Link carpeta: {folder_link}")


if __name__ == "__main__":
    main()
