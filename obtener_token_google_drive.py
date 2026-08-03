"""
ONE-SHOT en el PC (con navegador): genera el refresh token de Google Drive.

Pasos:
  1. Google Cloud Console → proyecto → APIs → habilitar Google Drive API
  2. Credenciales → OAuth client ID tipo "Desktop"
  3. Descargar JSON a: gmail_oauth/credentials_drive.json
  4. Ejecutar: python obtener_token_google_drive.py
  5. Copiar CLIENT_ID, CLIENT_SECRET y REFRESH_TOKEN a
     Cursor Dashboard → Cloud Agents → Secrets

NO commitear credentials ni tokens al repo.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow

# drive (no solo drive.file) para poder escribir en la carpeta
# existente "Agente WES / wes-scripts / reports" del Drive sincronizado.
SCOPES = ["https://www.googleapis.com/auth/drive"]
CREDENTIALS = Path(__file__).resolve().parent / "gmail_oauth" / "credentials_drive.json"
OUT_LOCAL = Path(__file__).resolve().parent / "gmail_oauth" / "drive_token_local.json"


def main() -> int:
    if sys.platform == "win32":
        for s in (sys.stdout, sys.stderr):
            try:
                s.reconfigure(encoding="utf-8", line_buffering=True)
            except Exception:
                pass

    if not CREDENTIALS.is_file():
        print(f"[ERROR] Falta {CREDENTIALS}")
        print("Descargá el JSON de OAuth (Desktop) desde Google Cloud Console.")
        return 1

    print("[INFO] Se abrirá el navegador para autorizar acceso a Drive...")
    flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS), SCOPES)
    creds = flow.run_local_server(port=0)

    if not creds.refresh_token:
        print(
            "[ERROR] Google no devolvió refresh_token.\n"
            "En Google Cloud → OAuth consent: revocá el acceso de la app "
            "y volvé a correr este script, o usá access_type=offline."
        )
        return 1

    raw = json.loads(CREDENTIALS.read_text(encoding="utf-8"))
    installed = raw.get("installed") or raw.get("web") or {}
    client_id = installed.get("client_id", "")
    client_secret = installed.get("client_secret", "")

    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": creds.refresh_token,
        "token_uri": "https://oauth2.googleapis.com/token",
        "scopes": SCOPES,
    }
    OUT_LOCAL.parent.mkdir(parents=True, exist_ok=True)
    OUT_LOCAL.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    print()
    print("=" * 60)
    print("Listo. Pegá ESTOS secretos en Cursor Cloud Agents → Secrets:")
    print("=" * 60)
    print(f"GOOGLE_DRIVE_CLIENT_ID={client_id}")
    print(f"GOOGLE_DRIVE_CLIENT_SECRET={client_secret}")
    print(f"GOOGLE_DRIVE_REFRESH_TOKEN={creds.refresh_token}")
    print()
    print("Opcional (carpeta fija en Drive; ID desde la URL de la carpeta):")
    print("GOOGLE_DRIVE_FOLDER_ID=xxxxxxxxxxxxxxxxxxxxx")
    print()
    print(f"Copia local (NO subir a git): {OUT_LOCAL}")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
