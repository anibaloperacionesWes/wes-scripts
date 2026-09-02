"""
Corrida de los lunes: consolidado semanal → Drive → Juan y Aníbal.

Cubre la semana lunes–domingo anterior. No enviar a Diego.
No reemplaza el informe de fin de mes.

Automatización Cursor (igual que puntos en cero diario):
  Nombre: Consolidado semanal gestión hídrica
  Horario: cada lunes 08:30 hora Chile
  Prompt:
    En wes-scripts correr `python correr_consolidado_semanal_lunes.py`.
    Genera el PDF consolidado de la semana lunes–domingo anterior, súbelo a
    Drive (CONSOLIDADO/SEMANAL) y envíalo a juanlopez@wes.cl y
    anibal.aoperaciones@wes.cl. No enviar a Diego. El correo debe listar
    los puntos a revisar. Si un cliente falla, seguir con el resto.

Uso:
  python correr_consolidado_semanal_lunes.py
  python correr_consolidado_semanal_lunes.py --sin-correo
  python correr_consolidado_semanal_lunes.py --hasta 31/08/2026 --sin-correo
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime

from enviar_consolidado_semanal_juan_anibal import enviar
from generar_consolidado_semanal_gestion_hidrica import generar_consolidado
from generar_informes_gestion_hidrica_semanal import _rango_es, _semana_completa
from wes_google_drive import credenciales_configuradas, subir_a_drive


def main() -> int:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
            sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
        except Exception:
            pass
    parser = argparse.ArgumentParser()
    parser.add_argument("--hasta", default=None, help="Último día dd/mm/YYYY")
    parser.add_argument(
        "--sin-correo",
        action="store_true",
        help="Genera y sube a Drive, no envía correo.",
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    hasta = datetime.strptime(args.hasta, "%d/%m/%Y") if args.hasta else None
    start, end = _semana_completa(hasta)
    pdf, filas, sin_alerta = generar_consolidado(start, end)
    periodo = _rango_es(start, end)
    drive = ""
    if credenciales_configuradas():
        info = subir_a_drive(pdf, subcarpeta="CONSOLIDADO/SEMANAL")
        drive = info.get("web_view_link") or ""
        print(f"[OK] Drive: {drive}", flush=True)
    else:
        print("[ADVERTENCIA] Sin credenciales Drive; se omite la subida.", flush=True)
    if args.sin_correo:
        print("[INFO] --sin-correo: no se envía.", flush=True)
        return 0
    enviar(
        pdf,
        periodo=periodo,
        filas=filas,
        sin_alerta=sin_alerta,
        drive=drive,
        dry_run=args.dry_run,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
