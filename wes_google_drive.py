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

# Acceso completo para consolidar en Agente WES/wes-scripts/reports
SCOPES = ["https://www.googleapis.com/auth/drive"]

# Carpeta Drive: Agente WES / wes-scripts / reports
# (misma que G:\Mi unidad\Agente WES\wes-scripts\reports)
DEFAULT_REPORTS_FOLDER_ID = "1r-eMj4SWkxJs045MDfky2x3jqrifalcj"

# Carpeta Drive: Agente WES / wes-scripts
# (misma que G:\Mi unidad\Agente WES\wes-scripts)
DEFAULT_WES_SCRIPTS_FOLDER_ID = "1WvEmtl3bexNvffxheGIQ6bUkN-iA4Lhg"

# Carpeta Drive: Agente WES / wes-scripts / mantenimiento wes
# (misma que G:\Mi unidad\Agente WES\wes-scripts\mantenimiento wes)
DEFAULT_MANTENIMIENTO_FOLDER_ID = "150GFVtGFlPXb_7bQfe7AS4SClKEXLEuX"

# Carpeta Drive: G:\Mi unidad\Actas de Mantencion
# Actas PDF: {Cliente}/{Año}/{mes}/
DEFAULT_ACTAS_MANTENCION_FOLDER_ID = "1-gDG2ND4beTpiqJqUG7d3dsT6wiHbKeQ"


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
    parent = (
        parent_id
        or _env("GOOGLE_DRIVE_FOLDER_ID")
        or DEFAULT_REPORTS_FOLDER_ID
        or "root"
    )
    # Permitir rutas tipo "UDD/ABREGADO"
    partes = [p for p in re.split(r"[/\\]+", nombre.strip()) if p]
    folder_id = parent
    last_link = f"https://drive.google.com/drive/folders/{folder_id}"
    for parte in partes:
        safe = limpiar_nombre(parte).replace("'", "\\'")
        parent_clause = (
            f"'{folder_id}' in parents"
            if folder_id != "root"
            else "'root' in parents"
        )
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
                "name": limpiar_nombre(parte),
                "mimeType": "application/vnd.google-apps.folder",
            }
            if folder_id and folder_id != "root":
                meta["parents"] = [folder_id]
            folder_id = (
                service.files().create(body=meta, fields="id").execute()["id"]
            )
        last_link = f"https://drive.google.com/drive/folders/{folder_id}"

    return {
        "id": folder_id,
        "web_view_link": last_link,
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
    dest = folder_id or _env("GOOGLE_DRIVE_FOLDER_ID") or DEFAULT_REPORTS_FOLDER_ID

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


def subir_a_mantenimiento_wes(
    file_path: Path | str,
    *,
    subcarpeta: Optional[str] = None,
    nombre: Optional[str] = None,
) -> Dict[str, str]:
    """
    Sube bajo Agente WES / wes-scripts / mantenimiento wes
    (nunca a la raíz de Mi unidad ni suelto en reports).
    """
    return subir_a_drive(
        file_path,
        folder_id=DEFAULT_MANTENIMIENTO_FOLDER_ID,
        subcarpeta=subcarpeta,
        nombre=nombre,
    )


_MESES_ES = (
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
)

_MAP_CLIENTE_ACTAS = {
    "COR. PUENTE": "CORP PUENTE ALTO",
    "GENCHI": "GENDARMERIA",
    "LA FLORIDA": "CORP LA FLORIDA",
    "LA REINA": "CORP LA REINA",
    "PROVIDENCIA": "CORP PROVIDENCIA",
    "LAS CONDES": "COLEGIO LAS CONDES",
    "NIDO": "NIDO DE AGUILAS",
    "BUPA ANTOFGASTA": "BUPA",
    "HEGC": "HOSPITAL EXEQUIEL GONZALEZ CORTES",
    "MADECCO": "MADECO",
    "MAE": "MADECO",
    "PAE": "PARQUE ARAUCO",
    "PAK": "PARQUE ARAUCO",
}


def carpeta_acta_cliente_mes(cliente: str, fecha: Optional[str] = None) -> str:
    """Ruta relativa bajo Actas de Mantencion: Cliente/Año/mes."""
    from datetime import datetime

    cli = _MAP_CLIENTE_ACTAS.get((cliente or "SIN_CLIENTE").strip(), (cliente or "SIN_CLIENTE").strip())
    d = None
    if fecha:
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
            try:
                d = datetime.strptime(str(fecha)[:10], fmt)
                break
            except ValueError:
                continue
    if d is None:
        d = datetime.now()
    return f"{cli}/{d.year}/{_MESES_ES[d.month - 1]}"


def subir_acta_mantencion(
    file_path: Path | str,
    *,
    cliente: str,
    fecha: Optional[str] = None,
    nombre: Optional[str] = None,
) -> Dict[str, str]:
    """Sube PDF de acta a G:\\Mi unidad\\Actas de Mantencion\\Cliente\\Año\\mes\\."""
    return subir_a_drive(
        file_path,
        folder_id=DEFAULT_ACTAS_MANTENCION_FOLDER_ID,
        subcarpeta=carpeta_acta_cliente_mes(cliente, fecha),
        nombre=nombre,
    )
