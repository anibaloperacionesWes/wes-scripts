"""
Informes de gestión hídrica (one-pager + mensual) — Liceo 7 Luisa Saavedra.

Periodo: 1 al 31 de agosto de 2026.
Nodo: 000006-04 (Liceo 7), aislado del agregado de Colegios Providencia
para diagnosticar el salto de consumo concentrado en este establecimiento.

Uso:
    python generar_informes_gestion_hidrica_liceo7_agosto2026.py
"""

from __future__ import annotations

import sys
from typing import Any, Dict, List

from generar_informes_gestion_hidrica_lote_agosto2026 import run_lote

CLIENTES: List[Dict[str, Any]] = [
    {
        "key": "liceo7",
        "company_id": "000006",
        "folder": "Providencia/Liceo_7",
        "cliente": "Liceo 7",
        "sitio": "Liceo 7 Luisa Saavedra",
        "sujeto": "el liceo",
        "verbo_registro": "registró",
        "node_ids": ["000006-04"],
        "apply_exclusions": False,
        "matriz_id": "000006-04",
        "matriz_name": "Liceo 7",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "workers": 4,
        "short_names": {"000006-04": "Liceo 7"},
        "leyenda": None,
        "chart_nota": (
            "un único punto de monitoreo (000006-04) representa el consumo "
            "del Liceo 7 Luisa Saavedra."
        ),
        "nocturno_nota": (
            "El consumo nocturno corresponde al único medidor del Liceo 7."
        ),
        "panorama_nota": (
            "Informe acotado al Liceo 7 para explicar el salto de agosto "
            "visto en el agregado de Colegios Providencia (concentró ~84 % "
            "del consumo total del mes)."
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
        "GESTIÓN HÍDRICA — Liceo 7 — 01/08/2026 a 31/08/2026",
    )


if __name__ == "__main__":
    main()
