"""
Informe de gestión hídrica (formato Zapallar) — COPEC.

Periodo: 1 al 31 de agosto de 2026.
La Matriz Principal es el consumo real de la estación; los demás medidores
no se suman (están aguas abajo o son reutilización).

Uso:
    python generar_informes_gestion_hidrica_copec_agosto2026.py
"""

from __future__ import annotations

import sys
from typing import Any, Dict, List

from generar_informes_gestion_hidrica_lote_agosto2026 import run_lote

CLIENTES: List[Dict[str, Any]] = [
    {
        "key": "copec",
        "company_id": "000009",
        "folder": "COPEC",
        "cliente": "COPEC",
        "sitio": "COPEC",
        "sujeto": "la estación",
        "node_ids": [
            "000009-06",
            "000009-05",
            "000009-04",
            "000009-09",
            "000009-10",
            "000009-11",
            "000009-08",
            "000009-01",
            "000009-00",
            "000009-03",
            "000009-02",
        ],
        "apply_exclusions": False,
        "matriz_id": "000009-06",
        "matriz_name": "Matriz Principal",
        "additive": False,
        "nocturnal_explain": None,
        "kpi_label": "Consumo de entrada",
        "workers": 8,
        "short_names": {
            "000009-06": "Matriz Principal",
            "000009-05": "Riego",
            "000009-04": "Lav. auto S",
            "000009-09": "Lav. autoserv. N",
            "000009-10": "Lav. autoserv. S",
            "000009-11": "Pronto Tienda",
            "000009-08": "Pronto Baños",
            "000009-01": "Oficina",
            "000009-00": "Costanera",
            "000009-03": "Lav. auto N",
            "000009-02": "Estanque reutiliz.",
        },
        "leyenda": (
            "Matriz Principal (entrada real)",
            "Puntos internos y reutilización (no se suman)",
        ),
        "chart_nota": (
            "la Matriz Principal representa el consumo real de red de la estación. "
            "Lavado, Pronto, riego y oficina están aguas abajo o en derivaciones del "
            "mismo caudal y no se suman. El estanque de reutilización no es agua de red."
        ),
        "nocturno_nota": (
            "El nocturno de la Matriz Principal es la referencia. Los puntos internos "
            "no se suman porque miden tramos del mismo caudal."
        ),
        "panorama_nota": (
            "Agosto se evalúa del 1 al 31 y no se extrapola. Costanera, Lavado automático "
            "Norte y Estanque de reutilización registraron 0 m³ en el periodo."
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
    run_lote(CLIENTES, "GESTIÓN HÍDRICA — COPEC — 01/08/2026 a 31/08/2026")


if __name__ == "__main__":
    main()
