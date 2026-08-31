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
    """Completa el cfg del lote con el perfil operativo de Lo Valledor.

    La ficha del mercado (tipo de recinto, visitantes, usos de agua) queda en
    VALLEDOR para la lectura interna; no se copia al informe.
    """
    p = VALLEDOR
    cfg = dict(cfg)
    cfg["perfil"] = p["tipo"]
    cfg["cpa_estado"] = p["cpa_estado"]
    cfg["nocturnal_explain"] = "mercado"
    cfg["panorama_nota"] = (
        "Agosto se evalúa del 1 al 31 y no se extrapola. "
        "El equipo CPA está instalado y falta activar y programar."
    )
    cfg["nocturno_nota"] = (
        "El caudal en 00:00–06:59 incluye el peak 22:00–03:00 y no se lee como pérdida. "
        "P1 y Barrio Norte se suman porque no miden el mismo caudal."
    )
    cfg["ventana_nocturna"] = (
        "Se considera nocturno el volumen entre las 00:00 y las 06:59. "
        "En este recinto esa ventana incluye el peak 22:00–03:00; no se interpreta como pérdida. "
    )
    cfg["lectura_nocturno"] = (
        ". El volumen entre las 00:00 y las 06:59 (incluye el peak 22:00–03:00) alcanzó "
    )
    return cfg
