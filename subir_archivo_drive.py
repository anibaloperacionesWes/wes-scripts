"""
Sube uno o más archivos a Google Drive (modo headless).

Requiere secretos GOOGLE_DRIVE_* (ver wes_google_drive.py).

Uso:
  python subir_archivo_drive.py reports/UDD/.../archivo.pdf
  python subir_archivo_drive.py archivo.docx --subcarpeta "Reportes_iPhone/UDD"
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from wes_google_drive import credenciales_configuradas, subir_a_drive


def main() -> int:
    if sys.platform == "win32":
        for s in (sys.stdout, sys.stderr):
            try:
                s.reconfigure(encoding="utf-8", line_buffering=True)
            except Exception:
                pass

    parser = argparse.ArgumentParser(description="Subir archivo(s) a Google Drive")
    parser.add_argument("archivos", nargs="+", type=Path, help="Rutas a subir")
    parser.add_argument(
        "--subcarpeta",
        default=None,
        help="Subcarpeta bajo GOOGLE_DRIVE_FOLDER_ID (o root)",
    )
    parser.add_argument(
        "--folder-id",
        default=None,
        help="ID de carpeta destino (override de GOOGLE_DRIVE_FOLDER_ID)",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Imprimir resultado en JSON",
    )
    args = parser.parse_args()

    if not credenciales_configuradas():
        print(
            "[ERROR] Faltan secretos GOOGLE_DRIVE_CLIENT_ID / "
            "CLIENT_SECRET / REFRESH_TOKEN.\n"
            "1) En el PC: python obtener_token_google_drive.py\n"
            "2) Pegá los valores en Cursor → Cloud Agents → Secrets",
            file=sys.stderr,
        )
        return 1

    resultados = []
    for path in args.archivos:
        if not path.is_file():
            print(f"[ERROR] No existe: {path}", file=sys.stderr)
            return 1
        info = subir_a_drive(
            path,
            folder_id=args.folder_id,
            subcarpeta=args.subcarpeta,
        )
        resultados.append(info)
        if not args.json:
            print(f"[OK] {info['name']}")
            print(f"     Link: {info['web_view_link']}")

    if args.json:
        print(json.dumps(resultados if len(resultados) > 1 else resultados[0], indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
