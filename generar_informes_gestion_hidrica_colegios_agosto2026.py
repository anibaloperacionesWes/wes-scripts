"""
Informes de gestión hídrica (formato Zapallar) — lote pendiente de colegios.

Periodo: 1 al 31 de agosto de 2026.
Clientes: Renca (colegios), La Florida, La Reina, CORMUP Peñalolén.

Mismo motor que el lote comercial de 8. Incluye visitas del formulario
cuando hay alguna en el recinto.

Uso:
    python generar_informes_gestion_hidrica_colegios_agosto2026.py
"""

from __future__ import annotations

import sys
from typing import Any, Dict, List

from generar_informes_gestion_hidrica_lote_agosto2026 import run_lote

NODOS_CORMUP = [f"000008-{i:02d}" for i in range(1, 15)]

CLIENTES: List[Dict[str, Any]] = [
    {
        "key": "renca",
        "company_id": "000017",
        "folder": "Renca",
        "cliente": "Renca",
        "sitio": "Colegios de Renca",
        "sujeto": "los colegios",
        "node_ids": ["000017-04", "000017-07", "000017-08"],
        "apply_exclusions": False,
        "matriz_id": None,
        "matriz_name": "",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "workers": 6,
        "short_names": {
            "000017-04": "Lo Velásquez",
            "000017-07": "Cumbre Pte.",
            "000017-08": "ICCO",
        },
        "leyenda": None,
        "chart_nota": (
            "Lo Velásquez, Cumbre Pte. e ICCO son establecimientos distintos: "
            "sus consumos se suman al total de colegios de Renca. Gimnasio y "
            "Piscina Municipal no entran en este informe."
        ),
        "nocturno_nota": (
            "El consumo nocturno se suma entre los tres colegios porque no miden el mismo caudal."
        ),
    },
    {
        "key": "florida",
        "company_id": "000028",
        "folder": "La_Florida",
        "cliente": "La Florida",
        "sitio": "Liceo Alto Cordillera",
        "sujeto": "el liceo",
        "node_ids": ["000028-01"],
        "apply_exclusions": False,
        "matriz_id": "000028-01",
        "matriz_name": "Alto Cordillera",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "workers": 4,
        "short_names": {"000028-01": "Alto Cordillera"},
        "leyenda": None,
        "chart_nota": "un único punto de monitoreo representa el consumo del liceo.",
        "nocturno_nota": (
            "El consumo nocturno corresponde al único medidor del Liceo Alto Cordillera."
        ),
    },
    {
        "key": "reina",
        "company_id": "000024",
        "folder": "La_Reina",
        "cliente": "La Reina",
        "sitio": "Eugenio María De Hostos",
        "sujeto": "el colegio",
        "node_ids": ["000024-01"],
        "apply_exclusions": False,
        "matriz_id": "000024-01",
        "matriz_name": "De Hostos",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "workers": 4,
        "short_names": {"000024-01": "De Hostos"},
        "leyenda": None,
        "chart_nota": "un único punto de monitoreo representa el consumo del colegio.",
        "nocturno_nota": (
            "El consumo nocturno corresponde al único medidor de Eugenio María De Hostos."
        ),
        "panorama_nota": (
            "Agosto se evalúa del 1 al 31 y no se extrapola. El consumo de este colegio "
            "se considera operativo desde el 05/08 (los días 01 al 04 pueden corresponder "
            "a regularización)."
        ),
    },
    {
        "key": "cormup",
        "company_id": "000008",
        "folder": "CORMUP",
        "cliente": "CORMUP",
        "sitio": "CORMUP Peñalolén",
        "sujeto": "la corporación",
        "node_ids": list(NODOS_CORMUP),
        "apply_exclusions": False,
        "matriz_id": None,
        "matriz_name": "",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "workers": 8,
        "short_names": {
            "000008-01": "Hermida F.",
            "000008-02": "E. de la Barra",
            "000008-03": "C. Fernández",
            "000008-04": "Tobalaba",
            "000008-05": "Santa María",
            "000008-06": "Arrieta C.",
            "000008-07": "Erasmo Escala",
            "000008-08": "Alicura",
            "000008-09": "J.B. Pastene",
            "000008-10": "M. Huici",
            "000008-11": "Valle Hermoso",
            "000008-12": "Unión Árabe",
            "000008-13": "Likankura",
            "000008-14": "Juan Pablo II",
        },
        "leyenda": None,
        "chart_nota": (
            "Cada barra es un establecimiento. Los 14 colegios se suman al total de CORMUP."
        ),
        "nocturno_nota": (
            "El nocturno se suma entre establecimientos porque son recintos distintos. "
            "En CORMUP la ventana es UTC 00:00–07:00, igual que en la app."
        ),
        "ventana_nocturna": (
            "En los colegios CORMUP el nocturno se toma del CSV horario, marcas UTC 00:00 "
            "a 07:00 (misma ventana que la app). "
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
        "GESTIÓN HÍDRICA — lote pendiente colegios — 01/08/2026 a 31/08/2026",
    )


if __name__ == "__main__":
    main()
