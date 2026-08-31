"""
Perfiles de cliente para informes de gestión hídrica.

Cada recinto tiene un tipo de operación distinto: el mismo % nocturno no se lee
igual en un colegio, una sucursal o un mercado mayorista. Este módulo concentra
esa lectura para personalizar clasificación, hallazgos y acciones.

Empezar por Lo Valledor; ir sumando clientes.
"""
from __future__ import annotations

from typing import Any, Dict

# Mercado Mayorista Lo Valledor (https://mercadolovalledor.cl/).
# Principal centro hortofrutícola de Chile: ~300.000 m², ~20.000 visitantes/día,
# ~8.000 vehículos, ~2.700 puntos de venta. Vigilancia 24/7, 9 baños con ducha
# de agua caliente, foodtrucks. El agua se usa para lavado de producto, higiene,
# cocina y aseo del recinto — de día y de madrugada.
VALLEDOR: Dict[str, Any] = {
    "tipo": "mercado_mayorista_hortofruticola",
    "nombre": "Mercado Mayorista Lo Valledor",
    "sitio_oficial": "https://mercadolovalledor.cl/",
    # Horario publicado (accesos y horarios + FAQ).
    "horario_compra_lv": "Lun–vie 17:00 a 14:00 del día siguiente",
    "horario_compra_finde": "Sábado cierra 14:00; domingo reabre 16:00",
    "horario_camiones": "Domingo 11:00 a sábado 14:00 (continuo); cerrado sáb 14:00–dom 11:00",
    # Peak comercial informado por operación WES.
    "horario_full": "22:00–03:00",
    "cpa_estado": "instalado_pendiente",
    "cpa_nota": (
        "El equipo CPA está instalado y falta activarlo y programar. "
        "Sin ese control el recinto queda en Requiere atención."
    ),
}


def aplicar_perfil_valledor(cfg: Dict[str, Any]) -> Dict[str, Any]:
    """Completa el cfg del lote con el perfil operativo de Lo Valledor."""
    p = VALLEDOR
    cfg = dict(cfg)
    cfg["sitio"] = p["nombre"]
    cfg["sujeto"] = "el mercado"
    cfg["perfil"] = p["tipo"]
    cfg["cpa_estado"] = p["cpa_estado"]
    cfg["nocturnal_explain"] = "mercado"
    cfg["panorama_nota"] = (
        "Principal mercado mayorista hortofrutícola de Chile (~300.000 m², "
        "más de 20.000 visitantes/día). Compra L–V 17:00–14:00 del día siguiente; "
        "peak comercial 22:00–03:00. El consumo 24 h es esperable (lavado de producto, "
        "baños con ducha, foodtrucks). El equipo CPA está instalado y falta activar y programar."
    )
    cfg["nocturno_nota"] = (
        "En un mercado mayorista el caudal de madrugada es uso del recinto "
        "(compra, peak 22:00–03:00, lavado e higiene), no una pérdida. "
        "P1 y Barrio Norte se suman porque no miden el mismo caudal."
    )
    cfg["ventana_nocturna"] = (
        "La ventana 00:00–06:59 se reporta como referencia. En Lo Valledor cae "
        "dentro del horario de compra (hasta las 14:00) y se solapa con el peak "
        "comercial 22:00–03:00; no se interpreta como pérdida. "
    )
    cfg["lectura_nocturno"] = (
        ". Lo Valledor opera como mercado mayorista hortofrutícola: compra L–V "
        "desde las 17:00 hasta las 14:00 del día siguiente y peak 22:00–03:00. "
        "El volumen entre las 00:00 y las 06:59 (uso normal del recinto) alcanzó "
    )
    return cfg
