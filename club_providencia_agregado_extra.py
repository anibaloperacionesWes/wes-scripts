"""Compatibilidad — Club Providencia delega en agregado_extendido_extra."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import List

from docx import Document

from agregado_extendido_extra import (
    agregar_analisis_nocturno_extendido,
    agregar_secciones_consumo_diario_y_max_dia as _secciones_diario,
    narrativa_consumo_total_extendido,
)

_COMPANY = "000031"


def agregar_secciones_consumo_diario_y_max_dia(
    doc: Document,
    nodes_data: List[dict],
    start_dt: datetime,
    end_dt: datetime,
    output_dir: Path,
) -> None:
    _secciones_diario(_COMPANY, doc, nodes_data, start_dt, end_dt, output_dir)


def agregar_analisis_nocturno_club_providencia(
    doc: Document,
    nodes_data: List[dict],
    start_dt: datetime,
    end_dt: datetime,
    output_dir: Path,
    price_per_m3: float,
) -> None:
    agregar_analisis_nocturno_extendido(
        _COMPANY, doc, nodes_data, start_dt, end_dt, output_dir, price_per_m3
    )


def narrativa_consumo_total_club_providencia(nodes_data: List[dict]) -> str:
    return narrativa_consumo_total_extendido(_COMPANY, nodes_data)
