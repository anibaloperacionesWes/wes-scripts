"""
Informe de gestión hídrica (formato Zapallar) — CDUC.

Periodo: 1 al 31 de agosto de 2026.
Puntos operativos: Club House, Raimundo Tupper, Equitación,
Calle de Servicio y Canchas de Tenis. Edificio Deportivo y Rugby
quedan fuera (exclusión operativa habitual).

Uso:
    python generar_informes_gestion_hidrica_cduc_agosto2026.py
"""

from __future__ import annotations

import sys
from typing import Any, Dict, List

from generar_informes_gestion_hidrica_lote_agosto2026 import run_lote

CLIENTES: List[Dict[str, Any]] = [
    {
        "key": "cduc",
        "company_id": "000021",
        "folder": "CDUC",
        "cliente": "CDUC",
        "sitio": "CDUC",
        "sujeto": "el club",
        "node_ids": [
            "000021-01",
            "000021-03",
            "000021-04",
            "000021-05",
            "000021-07",
        ],
        "apply_exclusions": False,
        "matriz_id": None,
        "matriz_name": "",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "workers": 5,
        "short_names": {
            "000021-01": "Club House",
            "000021-03": "Raimundo Tupper",
            "000021-04": "Equitación",
            "000021-05": "Calle de Servicio",
            "000021-07": "Canchas de Tenis",
        },
        "leyenda": None,
        "chart_nota": (
            "Cada barra es un recinto. Club House, Raimundo Tupper, Equitación, "
            "Calle de Servicio y Canchas de Tenis se suman al total. "
            "Edificio Deportivo y Rugby no entran en este informe."
        ),
        "nocturno_nota": (
            "El consumo nocturno se suma entre recintos porque son puntos distintos."
        ),
        "panorama_nota": (
            "Agosto se evalúa del 1 al 31 y no se extrapola. "
            "Edificio Deportivo y Rugby quedan fuera del total."
        ),
    },
]


def main() -> None:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
            sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
        except Exception:
            pass
    run_lote(CLIENTES, "GESTIÓN HÍDRICA — CDUC — 01/08/2026 a 31/08/2026")


if __name__ == "__main__":
    main()
