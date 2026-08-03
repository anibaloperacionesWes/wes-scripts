"""
Subida headless a Google Drive (sin navegador).

Pensado para Cloud Agents de Cursor (iPhone sin PC encendido):
usa secretos de entorno, no archivos locales interactivos.

Secretos / variables de entorno:
  GOOGLE_DRIVE_CLIENT_ID
  GOOGLE_DRIVE_CLIENT_SECRET
  GOOGLE_DRIVE_REFRESH_TOKEN
  GOOGLE_DRIVE_FOLDER_ID   (opcional; carpeta destino en Drive)

Uso:
  from wes_google_drive import subir_a_drive
  info = subir_a_drive(Path("reports/.../archivo.pdf"))
  print(info["web_view_link"])
"""

from __future__ import annotations

import mimetypes
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SCOPES = ["https://www.googleapis.com/auth/drive.file"]


def _env(name: str) -> str:
    return os.environ.get(name, "").strip()


def credenciales_configuradas() -> bool:
    return bool(
        _env("GOOGLE_DRIVE_CLIENT_ID")
        and _env("GOOGLE_DRIVE_CLIENT_SECRET")
        and _env("GOOGLE_DRIVE_REFRESH_TOKEN")
    )


def obtener_servicio_drive():
    """Autentica con refresh token (sin abrir navegador)."""
    if not credenciales_configuradas():
        raise RuntimeError(
            "Faltan secretos de Google Drive. Configurá en Cursor Dashboard → "
            "Cloud Agents → Secrets:\n"
            "  GOOGLE_DRIVE_CLIENT_ID\n"
            "  GOOGLE_DRIVE_CLIENT_SECRET\n"
            "  GOOGLE_DRIVE_REFRESH_TOKEN\n"
            "Opcional: GOOGLE_DRIVE_FOLDER_ID\n"
            "Generá el refresh token una vez en el PC con: "
            "python obtener_token_google_drive.py"
        )

    creds = Credentials(
        token=None,
        refresh_token=_env("GOOGLE_DRIVE_REFRESH_TOKEN"),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=_env("GOOGLE_DRIVE_CLIENT_ID"),
        client_secret=_env("GOOGLE_DRIVE_CLIENT_SECRET"),
        scopes=SCOPES,
    )
    creds.refresh(Request())
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def limpiar_nombre(nombre: str) -> str:
    nombre = re.sub(r'[<>:"/\\|?*]', "_", nombre)
    return re.sub(r"\s+", " ", nombre).strip()


def asegurar_carpeta(
    service,
    nombre: str,
    parent_id: Optional[str] = None,
) -> Dict[str, str]:
    """Busca o crea una carpeta. Retorna id y link."""
    parent = parent_id or _env("GOOGLE_DRIVE_FOLDER_ID") or "root"
    safe = limpiar_nombre(nombre).replace("'", "\\'")
    parent_clause = f"'{parent}' in parents" if parent != "root" else "'root' in parents"
    query = (
        f"name='{safe}' and mimeType='application/vnd.google-apps.folder' "
        f"and {parent_clause} and trashed=false"
    )
    items = (
        service.files()
        .list(q=query, fields="files(id, name)", pageSize=5)
        .execute()
        .get("files", [])
    )
    if items:
        folder_id = items[0]["id"]
    else:
        meta: Dict[str, Any] = {
            "name": limpiar_nombre(nombre),
            "mimeType": "application/vnd.google-apps.folder",
        }
        if parent and parent != "root":
            meta["parents"] = [parent]
        folder_id = (
            service.files().create(body=meta, fields="id").execute()["id"]
        )

    return {
        "id": folder_id,
        "web_view_link": f"https://drive.google.com/drive/folders/{folder_id}",
    }


def subir_a_drive(
    file_path: Path | str,
    *,
    folder_id: Optional[str] = None,
    nombre: Optional[str] = None,
    subcarpeta: Optional[str] = None,
) -> Dict[str, str]:
    """
    Sube un archivo a Drive y retorna id + links.

    Si `subcarpeta` se indica, la crea (o reutiliza) bajo folder_id / FOLDER_ID / root.
    """
    path = Path(file_path)
    if not path.is_file():
        raise FileNotFoundError(f"No existe el archivo: {path}")

    service = obtener_servicio_drive()
    dest = folder_id or _env("GOOGLE_DRIVE_FOLDER_ID") or "root"

    if subcarpeta:
        dest = asegurar_carpeta(service, subcarpeta, parent_id=dest)["id"]

    fname = limpiar_nombre(nombre or path.name)
    mime, _ = mimetypes.guess_type(str(path))
    mime = mime or "application/octet-stream"

    # Si ya existe mismo nombre en la carpeta, actualizar
    parent_clause = f"'{dest}' in parents" if dest != "root" else "'root' in parents"
    safe_name = fname.replace("'", "\\'")
    existing = (
        service.files()
        .list(
            q=f"name='{safe_name}' and {parent_clause} and trashed=false",
            fields="files(id)",
            pageSize=1,
        )
        .execute()
        .get("files", [])
    )

    media = MediaFileUpload(str(path), mimetype=mime, resumable=True)
    fields = "id, name, webViewLink, webContentLink"

    if existing:
        file = (
            service.files()
            .update(fileId=existing[0]["id"], media_body=media, fields=fields)
            .execute()
        )
    else:
        meta: Dict[str, Any] = {"name": fname}
        if dest and dest != "root":
            meta["parents"] = [dest]
        file = (
            service.files()
            .create(body=meta, media_body=media, fields=fields)
            .execute()
        )

    file_id = file["id"]
    web = file.get("webViewLink") or f"https://drive.google.com/file/d/{file_id}/view"
    return {
        "id": file_id,
        "name": file.get("name", fname),
        "web_view_link": web,
        "web_content_link": file.get("webContentLink") or "",
        "folder_id": dest,
    }


def subir_varios(
    paths: List[Path | str],
    *,
    subcarpeta: Optional[str] = None,
    folder_id: Optional[str] = None,
) -> List[Dict[str, str]]:
    return [
        subir_a_drive(p, subcarpeta=subcarpeta, folder_id=folder_id) for p in paths
    ]
