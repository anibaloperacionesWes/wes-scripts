"""
Registro operativo Parque Arauco (000025): observaciones por nodo para listados y trazabilidad.
"""

from __future__ import annotations

# Textos breves para columna Observaciones en Excel/CSV.
OBSERVACIONES_LISTADO: dict[str, str] = {
    "000025-11": (
        "No operativo (Matriz principal 1° piso). Reemplazado por 000025-30 (Matriz A.A). "
        "Conservar registro según gestión / JO del mall si aplica."
    ),
    "000025-14": (
        "No operativo (Red de Incendio). Reemplazado por 000025-34 (Alimentación Baños). "
        "Conservar registro según gestión / JO del mall si aplica."
    ),
    "000025-30": (
        "Equipo instalado en reemplazo de 000025-11 (Matriz principal 1° piso)."
    ),
    "000025-34": (
        "Equipo instalado en reemplazo de 000025-14 (Red de Incendio)."
    ),
    "000025-25": (
        "Equipo retirado el 30-03-2026. Para la evaluación debe considerarse con información histórica."
    ),
    "000025-26": (
        "Equipo retirado el 30-03-2026. Para la evaluación debe considerarse con información histórica."
    ),
    "000025-35": (
        "Equipo instalado en reemplazo de 000025-25 (retirado 30-03-2026)."
    ),
    "000025-36": (
        "Equipo instalado en reemplazo de 000025-26 (retirado 30-03-2026)."
    ),
}


# Fecha recepción de obras por mall (texto columna igual que en listado: «Arauco …»).
RECEPCION_OBRAS_POR_MALL: dict[str, str] = {
    "Arauco Estación": "09-12-2025",
    "Arauco Maipú": "10-11-2025",
    "Arauco El Bosque": "03-11-2025",
    "Arauco Quilicura": "11-11-2025",
    "Arauco Curauma": "27-11-2025",
    "Arauco Buenaventura": "21-11-2025",
    "Arauco Kennedy": "12-12-2025",
}


def observacion_para_nodo(node_id: str) -> str:
    return OBSERVACIONES_LISTADO.get(str(node_id).strip(), "")


def recepcion_obras_para_mall(mall: str) -> str:
    return RECEPCION_OBRAS_POR_MALL.get(str(mall).strip(), "")
