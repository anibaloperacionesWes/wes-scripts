"""
Gestión hídrica — Clínica Bupa Antofagasta (formato Fundo Zapallar).

Genera:
  - One_Pager_Gestion_Hidrica_Bupa_Antofagasta_....pdf
  - Informe_Mensual_Bupa_Antofagasta_....pdf

Periodo por defecto: desde el inicio del monitoreo (23/07/2026) hasta hoy.

Uso:
  python generar_informes_gestion_hidrica_bupa_antofagasta.py
  python generar_informes_gestion_hidrica_bupa_antofagasta.py --start 23/07/2026 --end 11/09/2026
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import replace
from datetime import datetime
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
    except Exception:
        pass

import generar_informes_gestion_hidrica_lote_agosto2026 as lote
from informe_gestion_hidrica_pdf import render_mensual, render_one_pager

DEFAULT_START = "23/07/2026"
FOLDER = "Bupa_Antofagasta"

_MESES = {
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


def _fecha_larga(d: datetime) -> str:
    return f"{d.day} de {_MESES[d.month]} de {d.year}"


def _cfg(start: str, end: str) -> dict:
    start_dt = datetime.strptime(start, "%d/%m/%Y")
    end_dt = datetime.strptime(end, "%d/%m/%Y")
    periodo_corto = (
        f"Clínica Bupa Antofagasta · {_fecha_larga(start_dt)} al {_fecha_larga(end_dt)}"
    )
    return {
        "key": "bupa_antofagasta",
        "company_id": "000029",
        "folder": FOLDER,
        "cliente": "Bupa Antofagasta",
        "sitio": "Clínica Bupa Antofagasta",
        "sujeto": "la clínica",
        "node_ids": [
            "000029-09",  # Matriz / medidor principal primero
            "000029-07",
            "000029-08",
            "000029-10",
        ],
        "start": start,
        "end": end,
        "periodo_corto": periodo_corto,
        "usar_kpi_ultimo_mes_6m": False,
        "nota_agosto": (
            f"El periodo se informa del {start_dt.strftime('%d/%m/%Y')} "
            f"al {end_dt.strftime('%d/%m/%Y')} (desde el inicio del monitoreo). "
            "No se extrapola el consumo."
        ),
        "panorama_nota": (
            f"El consumo se considera del {start_dt.strftime('%d/%m/%Y')} "
            f"al {end_dt.strftime('%d/%m/%Y')}. No se extrapola el resto del mes."
        ),
        "apply_exclusions": False,
        "matriz_id": "000029-09",
        "matriz_name": "Medidor Principal Sanitaria",
        "additive": False,
        # Sin perfil condominio: redacción genérica de clínica (no “condominio/Zapallar”).
        "nocturnal_explain": None,
        "kpi_label": "Consumo de entrada",
        "short_names": {
            "000029-09": "Medidor Principal Sanitaria",
            "000029-07": "Sala de Bomba Principal",
            "000029-08": "Sala de Bomba Sexto Piso",
            "000029-10": "Sala de Bomba N°2",
        },
        "leyenda": (
            "Medidor Principal Sanitaria (entrada real)",
            "Salas de bomba (aguas abajo / control interno)",
        ),
        "chart_nota": (
            "el Medidor Principal Sanitaria representa la entrada real a la clínica. "
            "Las salas de bomba miden tramos o impulsiones internas; por lo tanto, "
            "sus consumos no se suman al total de entrada."
        ),
        "nocturno_nota": (
            "Los volúmenes de las salas de bomba se expresan respecto de la entrada "
            "principal, pero no son aditivos porque corresponden a impulsiones internas."
        ),
    }


def main() -> int:
    ap = argparse.ArgumentParser(description="Gestión hídrica Bupa Antofagasta (formato Zapallar)")
    ap.add_argument("--start", default=DEFAULT_START, help="dd/mm/aaaa")
    ap.add_argument(
        "--end",
        default=datetime.now().strftime("%d/%m/%Y"),
        help="dd/mm/aaaa (default: hoy)",
    )
    args = ap.parse_args()

    cfg = _cfg(args.start, args.end)
    start_dt = datetime.strptime(args.start, "%d/%m/%Y")
    end_dt = datetime.strptime(args.end, "%d/%m/%Y")
    periodo_tag = f"{start_dt:%Y%m%d}_{end_dt:%Y%m%d}"
    label_periodo = f"{start_dt:%d/%m/%Y}–{end_dt:%d/%m/%Y}"

    # Cache propio para no mezclar con el lote de agosto.
    lote.CACHE_DIR = Path("/tmp/gh_bupa_antofagasta")
    cache_file = lote.CACHE_DIR / f"{cfg['key']}.json"
    if cache_file.exists():
        cache_file.unlink()

    print("=" * 70)
    print("GESTIÓN HÍDRICA — Bupa Antofagasta (formato Fundo Zapallar)")
    print(f"Periodo: {args.start} → {args.end}")
    print(f"Nodos: {', '.join(cfg['node_ids'])}")
    print("=" * 70)

    data = lote.fetch_cliente(cfg)
    spec = lote.build_spec(cfg, data, visitas=[])
    # Redacción clínica (evitar textos de condominio/Zapallar).
    motivo = (
        "consumo nocturno del 24 % sobre la entrada (Medidor Principal Sanitaria) "
        "que requiere seguimiento y validación frente a la operación habitual de la clínica "
        "(salas de bomba e impulsiones internas)."
    )
    if float(data["kpi"]["pct_nocturno"]) >= 18:
        spec = replace(spec, clasificacion="EN OBSERVACIÓN", motivo=motivo)
    spec = replace(
        spec,
        footer=f"Informe de gestión hídrica - {cfg['cliente']} | {label_periodo}",
        titulo_onepager="Resumen ejecutivo de gestión hídrica",
        titulo_mensual="Informe mensual de gestión hídrica",
    )

    out_dir = Path("reports") / FOLDER / "GESTION_HIDRICA"
    charts = out_dir / "_charts"
    charts.mkdir(parents=True, exist_ok=True)

    one = out_dir / f"One_Pager_Gestion_Hidrica_Bupa_Antofagasta_{periodo_tag}.pdf"
    monthly = out_dir / f"Informe_Mensual_Bupa_Antofagasta_{periodo_tag}.pdf"

    render_one_pager(spec, one)
    render_mensual(spec, monthly, charts)
    print(f"[OK] {one}")
    print(f"[OK] {monthly}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
