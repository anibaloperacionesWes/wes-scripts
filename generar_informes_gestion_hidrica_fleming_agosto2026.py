#!/usr/bin/env python3
"""
Informe de gestión hídrica (formato Zapallar) — Alexander Fleming.

Periodo: 1 al 31 de agosto de 2026.
Un solo medidor (000022-00). Juan Pablo II (000022-01) no entra.

Uso:
    python generar_informes_gestion_hidrica_fleming_agosto2026.py
"""
from __future__ import annotations

import sys
from typing import Any, Dict, List

from generar_informes_gestion_hidrica_lote_agosto2026 import run_lote

CLIENTES: List[Dict[str, Any]] = [
    {
        "key": "fleming",
        "company_id": "000022",
        "folder": "Alexander_Fleming",
        "cliente": "Alexander Fleming",
        "sitio": "Alexander Fleming",
        "sujeto": "el recinto",
        "node_ids": ["000022-00"],
        "apply_exclusions": False,
        "matriz_id": "000022-00",
        "matriz_name": "Alexander Fleming",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "workers": 4,
        "short_names": {"000022-00": "Alexander Fleming"},
        "leyenda": None,
        "chart_nota": (
            "un único punto de monitoreo representa el consumo del recinto."
        ),
        "nocturno_nota": (
            "El consumo nocturno corresponde al único medidor. "
            "La ventana es 00:00–06:59, hora de Chile."
        ),
        "panorama_nota": (
            "Agosto se evalúa del 1 al 31 y no se extrapola. "
            "La serie no registra el 29/08."
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
    run_lote(CLIENTES, "GESTIÓN HÍDRICA — Alexander Fleming — 01/08/2026 a 31/08/2026")


if __name__ == "__main__":
    main()
