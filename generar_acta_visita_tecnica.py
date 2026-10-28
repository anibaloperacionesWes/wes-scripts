#!/usr/bin/env python3
"""
CLI — Formulario / acta de visita técnica WES (mismo PDF que el formulario web).

Formulario web oficial (terreno):
  https://script.google.com/a/macros/wes.cl/s/AKfycbx5DTddldWVkvK7QLxS85UbDDLQFG93_gJx3MYlL52gdPpq-FY3mRX9nxGQ7OOCbtxC/exec

Este script genera ese mismo PDF desde un JSON (campos del formulario).
Opcionalmente sube a Drive (Actas de Mantencion) y envía correo con acuse.

Uso:
  python generar_acta_visita_tecnica.py --json mantenimiento_wes/ejemplos/acta_ejemplo.json
  python generar_acta_visita_tecnica.py --json visita.json --subir-drive
  python generar_acta_visita_tecnica.py --json visita.json --subir-drive --enviar-correo
  python generar_acta_visita_tecnica.py --demo
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
from typing import Any, Dict, Optional

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "mantenimiento_wes"))

from generar_pdf_acta_visita import generar_pdf_acta  # noqa: E402

CHILE = ZoneInfo("America/Santiago")
# Carpeta Drive: Actas de Mantencion (regla acta-visita-tecnica)
ACTAS_DRIVE_FOLDER_ID = "1-gDG2ND4beTpiqJqUG7d3dsT6wiHbKeQ"

MESES_ES = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}


def _slug_path(text: str) -> str:
    import re

    t = (text or "cliente").strip()
    t = re.sub(r"[^\w\-]+", "_", t, flags=re.UNICODE)
    return (t.strip("_") or "cliente")[:60]


def _parse_fecha(fecha: str) -> datetime:
    fecha = (fecha or "").strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(fecha, fmt)
        except ValueError:
            continue
    return datetime.now(CHILE)


def nombre_pdf(data: Dict[str, Any], stamp: Optional[str] = None) -> str:
    stamp = stamp or datetime.now(CHILE).strftime("%Y%m%d_%H%M%S")
    folio = data.get("folio") or data.get("ot") or "sin_folio"
    fecha = str(data.get("fecha") or datetime.now(CHILE).date())
    cliente = _slug_path(str(data.get("cliente") or "CLIENTE")).upper()
    maquina = _slug_path(str(data.get("maquina") or "SITIO")).upper()
    return f"folio_{folio}_{fecha}_{cliente}_{maquina}_{stamp}.pdf"


def ruta_local(data: Dict[str, Any], out_dir: Optional[Path] = None) -> Path:
    stamp = datetime.now(CHILE).strftime("%Y%m%d_%H%M%S")
    cliente = _slug_path(str(data.get("cliente") or "Cliente"))
    if out_dir is None:
        out_dir = ROOT / "reports" / cliente / "Visitas_Tecnicas"
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir / nombre_pdf(data, stamp)


def subcarpeta_drive(data: Dict[str, Any]) -> str:
    """Ruta relativa bajo Actas de Mantencion: Cliente/año/mes."""
    cliente = str(data.get("cliente") or "Cliente").strip()
    dt = _parse_fecha(str(data.get("fecha") or ""))
    mes = MESES_ES.get(dt.month, f"{dt.month:02d}")
    return f"{cliente}/{dt.year}/{mes}"


def cargar_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def demo_payload() -> Dict[str, Any]:
    ahora = datetime.now(CHILE)
    return {
        "cliente": "DEMO",
        "maquina": "SITIO_EJEMPLO",
        "comuna": "Santiago",
        "fecha": ahora.strftime("%Y-%m-%d"),
        "hora": "10:00",
        "tecnico": "Anibal Aranda",
        "ot": "DEMO-001",
        "folio": "DEMO-001",
        "estado_visita": "cerrada",
        "motivos": ["Mantenimiento"],
        "tecnologias": ["CPA y CIR"],
        "tipo_mtto": "Mtto Preventivo",
        "tipo_falla": "Auditoría",
        "falla_especifica": "Validación Data",
        "lectura_medidor": "—",
        "solucion": (
            "Acta de demostración del generador oficial WES. "
            "Reemplazar este JSON con los datos reales de la visita."
        ),
        "observaciones": "Ejemplo — no enviar al cliente.",
        "recibido_por": "—",
        "cargo": "—",
        "email_cliente": "",
        "email_cc": "",
        "enviar_correo_cliente": False,
        "checklist_cir": [
            {"elemento": "Conectividad App", "estado": "OK", "obs": "Demo"}
        ],
        "checklist_cpa": [
            {"elemento": "Func. medidor", "estado": "OK", "obs": ""}
        ],
        "checklist_sab": [],
        "firma_png": "",
    }


def subir_drive(pdf_path: Path, data: Dict[str, Any]) -> Optional[Dict[str, str]]:
    from wes_google_drive import credenciales_configuradas, subir_a_drive

    if not credenciales_configuradas():
        print(
            "[ADVERTENCIA] Faltan secretos GOOGLE_DRIVE_*; se omite subida a Drive.",
            file=sys.stderr,
        )
        return None
    sub = subcarpeta_drive(data)
    info = subir_a_drive(
        pdf_path,
        folder_id=ACTAS_DRIVE_FOLDER_ID,
        subcarpeta=sub,
    )
    print(f"[OK] Drive ({sub}): {info.get('web_view_link')}")
    return info


def enviar_correo(pdf_path: Path, data: Dict[str, Any]) -> Dict[str, Any]:
    from enviar_acta_cliente_pdf import enviar_acta_pdf_cliente

    result = enviar_acta_pdf_cliente(pdf_path, data)
    print(f"[OK] Correo: {result}")
    return result


def main() -> int:
    if sys.platform == "win32":
        for s in (sys.stdout, sys.stderr):
            try:
                s.reconfigure(encoding="utf-8")
            except Exception:
                pass

    parser = argparse.ArgumentParser(
        description="Genera el acta oficial PDF de visita técnica WES"
    )
    parser.add_argument(
        "--json",
        type=Path,
        help="Ruta al JSON con los campos de la visita",
    )
    parser.add_argument(
        "--demo",
        action="store_true",
        help="Genera un PDF de demostración (sin Drive ni correo)",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help="Ruta de salida del PDF (opcional)",
    )
    parser.add_argument(
        "--subir-drive",
        action="store_true",
        help="Sube a Drive en Actas de Mantencion/{Cliente}/{año}/{mes}",
    )
    parser.add_argument(
        "--enviar-correo",
        action="store_true",
        help="Envía el PDF al email_cliente pidiendo acuse de recibo",
    )
    args = parser.parse_args()

    if args.demo:
        data = demo_payload()
    elif args.json:
        if not args.json.is_file():
            print(f"[ERROR] No existe: {args.json}", file=sys.stderr)
            return 1
        data = cargar_json(args.json)
    else:
        parser.error("Indicá --json RUTA o --demo")

    out_path = args.out or ruta_local(data)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Guardar copia del payload junto al PDF
    payload_path = out_path.with_name(out_path.stem + "_payload.json")
    payload_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    pdf = generar_pdf_acta(data, out_path=out_path)
    print(f"[OK] PDF: {pdf}")
    print(f"[OK] Payload: {payload_path}")

    drive_info = None
    if args.subir_drive and not args.demo:
        drive_info = subir_drive(pdf, data)

    if args.enviar_correo and not args.demo:
        data = dict(data)
        data["enviar_correo_cliente"] = True
        enviar_correo(pdf, data)

    result = {
        "pdf": str(pdf),
        "payload": str(payload_path),
        "drive": drive_info,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
