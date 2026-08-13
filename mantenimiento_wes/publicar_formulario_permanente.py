# -*- coding: utf-8 -*-
"""
Publica / actualiza el Formulario permanente (Google Apps Script) y la carpeta
de técnicos en Drive.

Uso:
  python3 mantenimiento_wes/publicar_formulario_permanente.py

Qué hace:
  1) Crea o actualiza el proyecto Apps Script «Formulario Visita WES Permanente»
  2) Crea la carpeta Drive «Tecnicos_WES_Formulario» con instrucciones
  3) Imprime el link de edición del Script (desde ahí se activa el /exec fijo)

Activación del link permanente (1 vez, desde el celular logueado en Google):
  1. Abrí el link de edición del Script
  2. Implementar → Nueva implementación → Aplicación web
  3. Ejecutar como: Yo
  4. Quién tiene acceso: Cualquiera
  5. Implementar → copiá la URL (/exec) y compartila con los técnicos
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any, Dict, Optional

ROOT = Path(__file__).resolve().parent
APPS = ROOT / "apps_script"
REPO = ROOT.parent
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from wes_google_drive import asegurar_carpeta, obtener_servicio_drive  # noqa: E402

SCRIPT_NAME = "Formulario Visita WES Permanente"
FOLDER_TECH = "Tecnicos_WES_Formulario"
STATE_FILE = ROOT / "apps_script" / "ESTADO_PUBLICACION.json"
SHEET_REGISTRO = "https://docs.google.com/spreadsheets/d/1GlRn7QXWEre7ziau29ojR5lTl-bZ8T3mCT3cD93HZgM/edit"


def _read_sources() -> Dict[str, str]:
    codigo = (APPS / "Codigo.gs").read_text(encoding="utf-8")
    formulario = (APPS / "Formulario.html").read_text(encoding="utf-8")
    catalogos = (APPS / "catalogos.html").read_text(encoding="utf-8")
    appsscript = json.dumps(
        {
            "timeZone": "America/Santiago",
            "exceptionLogging": "STACKDRIVER",
            "runtimeVersion": "V8",
            "webapp": {
                "executeAs": "USER_DEPLOYING",
                "access": "ANYONE_ANONYMOUS",
            },
        },
        indent=2,
    )
    return {
        "Codigo.gs": codigo,
        "Formulario.html": formulario,
        "catalogos.html": catalogos,
        "appsscript.json": appsscript,
    }


def _script_payload(sources: Dict[str, str]) -> Dict[str, Any]:
    files = [
        {
            "id": "appsscript",
            "name": "appsscript",
            "type": "json",
            "source": sources["appsscript.json"],
        },
        {
            "id": "codigo",
            "name": "Codigo",
            "type": "server_js",
            "source": sources["Codigo.gs"],
        },
        {
            "id": "formulario",
            "name": "Formulario",
            "type": "html",
            "source": sources["Formulario.html"],
        },
        {
            "id": "catalogos",
            "name": "catalogos",
            "type": "html",
            "source": sources["catalogos.html"],
        },
    ]
    return {"files": files}


def _find_script(service, name: str) -> Optional[Dict[str, str]]:
    safe = name.replace("'", "\\'")
    q = (
        f"name='{safe}' and mimeType='application/vnd.google-apps.script' "
        "and trashed=false"
    )
    files = (
        service.files()
        .list(q=q, fields="files(id,name,webViewLink)", pageSize=5)
        .execute()
        .get("files", [])
    )
    return files[0] if files else None


def publicar_script(service) -> Dict[str, str]:
    from googleapiclient.http import MediaInMemoryUpload

    sources = _read_sources()
    payload = json.dumps(_script_payload(sources), ensure_ascii=False).encode("utf-8")
    media = MediaInMemoryUpload(
        payload,
        mimetype="application/vnd.google-apps.script+json",
        resumable=False,
    )
    existing = _find_script(service, SCRIPT_NAME)
    if existing:
        updated = (
            service.files()
            .update(
                fileId=existing["id"],
                media_body=media,
                fields="id,name,webViewLink",
            )
            .execute()
        )
        return {
            "id": updated["id"],
            "web_view_link": updated.get("webViewLink")
            or f"https://script.google.com/d/{updated['id']}/edit",
            "updated": "true",
        }

    created = (
        service.files()
        .create(
            body={"name": SCRIPT_NAME, "mimeType": "application/vnd.google-apps.script"},
            media_body=media,
            fields="id,name,webViewLink",
        )
        .execute()
    )
    return {
        "id": created["id"],
        "web_view_link": created.get("webViewLink")
        or f"https://script.google.com/d/{created['id']}/edit",
        "updated": "false",
    }


def escribir_instrucciones(service, script_link: str, folder_id: str) -> Dict[str, str]:
    texto = f"""FORMULARIO PERMANENTE · TÉCNICOS WES
=====================================

1) ACTIVAR EL LINK FIJO (solo 1 vez, Anibal u operaciones)
   Abrí este proyecto Apps Script (con tu Google WES):
   {script_link}

   En el celular o PC:
   - Tocá «Implementar» (Deploy)
   - «Nueva implementación»
   - Tipo: «Aplicación web»
   - Ejecutar como: Yo
   - Quién tiene acceso: Cualquiera
   - Implementar
   - Copiá la URL que termina en /exec  ← ESE es el link permanente

2) USO DIARIO (técnicos)
   Abren el link /exec en el celular → completan acta + firma → Enviar.
   Eso:
   - genera el PDF
   - manda correo al cliente (acusar recibo)
   - agrega la fila en el Excel/Sheet de fallas

3) EXCEL / SHEET (ya permanente)
   {SHEET_REGISTRO}

4) Si cambian clientes/máquinas
   Actualizá el Sheet (Base1) y pedile al agente que vuelva a correr:
   python3 mantenimiento_wes/publicar_formulario_permanente.py
   (después: Implementar → Administrar implementaciones → Nueva versión)

Carpeta Drive de técnicos: Tecnicos_WES_Formulario
"""
    from googleapiclient.http import MediaInMemoryUpload

    media = MediaInMemoryUpload(texto.encode("utf-8"), mimetype="text/plain")
    # upsert INSTRUCCIONES_LINK_PERMANENTE.txt
    safe = "INSTRUCCIONES_LINK_PERMANENTE.txt"
    q = (
        f"name='{safe}' and '{folder_id}' in parents and trashed=false"
    )
    found = (
        service.files()
        .list(q=q, fields="files(id)", pageSize=1)
        .execute()
        .get("files", [])
    )
    if found:
        f = (
            service.files()
            .update(fileId=found[0]["id"], media_body=media, fields="id,webViewLink")
            .execute()
        )
    else:
        f = (
            service.files()
            .create(
                body={
                    "name": safe,
                    "parents": [folder_id],
                    "mimeType": "text/plain",
                },
                media_body=media,
                fields="id,webViewLink",
            )
            .execute()
        )
    return {
        "id": f["id"],
        "web_view_link": f.get("webViewLink")
        or f"https://drive.google.com/file/d/{f['id']}/view",
    }


def main() -> int:
    service = obtener_servicio_drive()
    folder = asegurar_carpeta(service, FOLDER_TECH)
    script = publicar_script(service)
    # mover script a la carpeta si se puede (add parent)
    try:
        meta = service.files().get(fileId=script["id"], fields="parents").execute()
        prev = ",".join(meta.get("parents") or [])
        service.files().update(
            fileId=script["id"],
            addParents=folder["id"],
            removeParents=prev or None,
            fields="id,parents",
        ).execute()
    except Exception as exc:  # noqa: BLE001
        print(f"[WARN] No se pudo mover el script a la carpeta: {exc}")

    instr = escribir_instrucciones(service, script["web_view_link"], folder["id"])

    state = {
        "script_id": script["id"],
        "script_edit_link": script["web_view_link"],
        "folder_id": folder["id"],
        "folder_link": folder["web_view_link"],
        "instrucciones_link": instr["web_view_link"],
        "sheet_registro": SHEET_REGISTRO,
        "nota": (
            "El link /exec permanente aparece después de Implementar → Aplicación web "
            "(acceso Cualquiera). Pegalo aquí cuando lo tengas."
        ),
        "webapp_exec_url": "",
    }
    STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

    # también bajo reports para Drive sync habitual
    reports = REPO / "reports" / "Mantenimientos" / "formulario_visita"
    reports.mkdir(parents=True, exist_ok=True)
    (reports / "LINK_PERMANENTE_INSTRUCCIONES.txt").write_text(
        json.dumps(state, ensure_ascii=False, indent=2)
        + "\n\nAbrí y activá (1 vez):\n"
        + f"{script['web_view_link']}\n",
        encoding="utf-8",
    )

    print("=" * 60)
    print("Formulario permanente preparado en Google")
    print("=" * 60)
    print(f"Carpeta técnicos: {folder['web_view_link']}")
    print(f"Script (activar /exec): {script['web_view_link']}")
    print(f"Instrucciones Drive: {instr['web_view_link']}")
    print()
    print("ACTIVACIÓN (1 minuto, con tu usuario Google WES):")
    print("  Implementar → Nueva implementación → Aplicación web")
    print("  Acceso: Cualquiera → Implementar → copiar URL /exec")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
