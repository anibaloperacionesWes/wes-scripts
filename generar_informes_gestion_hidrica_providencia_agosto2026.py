"""
Informes de gestión hídrica (one-pager + mensual extendido) — Colegios Providencia.

Periodo: 1 al 31 de agosto de 2026.
Nodos: Liceo Lastarria, Carmela Carvajal, Liceo 7, Liceo Juan Pablo Duarte
(000006-03 Arturo Alessandri Palma excluido por criterio operativo).

Mismo motor que el lote de agosto (formato Zapallar).

Uso:
    python generar_informes_gestion_hidrica_providencia_agosto2026.py
"""

from __future__ import annotations

import sys
from typing import Any, Dict, List

from generar_informes_gestion_hidrica_lote_agosto2026 import run_lote

CLIENTES: List[Dict[str, Any]] = [
    {
        "key": "providencia",
        "company_id": "000006",
        "folder": "Providencia",
        "cliente": "Colegios Providencia",
        "sitio": "Colegios de Providencia",
        "sujeto": "los colegios",
        "verbo_registro": "registraron",
        "node_ids": ["000006-01", "000006-02", "000006-04", "000006-05"],
        "apply_exclusions": False,
        "matriz_id": None,
        "matriz_name": "",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "workers": 4,
        "short_names": {
            "000006-01": "Lastarria",
            "000006-02": "Carmela Carvajal",
            "000006-04": "Liceo 7",
            "000006-05": "J.P. Duarte",
        },
        "leyenda": None,
        "chart_nota": (
            "Lastarria, Carmela Carvajal, Liceo 7 y Juan Pablo Duarte son "
            "establecimientos distintos: sus consumos se suman al total de "
            "colegios de Providencia. Arturo Alessandri Palma (000006-03) "
            "no entra en este informe."
        ),
        "nocturno_nota": (
            "El consumo nocturno se suma entre los cuatro colegios porque no "
            "miden el mismo caudal."
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
    run_lote(
        CLIENTES,
        "GESTIÓN HÍDRICA — Colegios Providencia — 01/08/2026 a 31/08/2026",
    )


if __name__ == "__main__":
    main()
