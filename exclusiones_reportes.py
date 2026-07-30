"""
Exclusiones compartidas para reportes (puntos en cero, individuales, agregados y PPT).

Registro auditable de exclusiones por pedido operativo / cliente: ver
registro_puntos_deshabilitados.txt (mismo directorio que este módulo).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable, List, Optional, Set

_PACKAGE_DIR = Path(__file__).resolve().parent
REGISTRO_PUNTOS_DESHABILITADOS_TXT = _PACKAGE_DIR / "registro_puntos_deshabilitados.txt"
_NODE_ID_LINE = re.compile(r"^(\d{6}-\d{2})\s*\|")


def _node_ids_desde_registro_puntos_deshabilitados() -> Set[str]:
    """Lee NODO_ID desde registro_puntos_deshabilitados.txt (primer campo de cada línea de datos)."""
    path = REGISTRO_PUNTOS_DESHABILITADOS_TXT
    if not path.is_file():
        return set()
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return set()
    found: Set[str] = set()
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        m = _NODE_ID_LINE.match(line)
        if m:
            found.add(m.group(1))
    return found

# IDs de empresas a excluir
EXCLUDED_COMPANY_IDS: Set[str] = {"000000", "000001", "000004"}

# Empresas excluidas solo del informe «puntos en cero» (creadas en app, sin instalación operativa).
EXCLUDED_COMPANY_IDS_SOLO_PUNTOS_EN_CERO: Set[str] = {
    "000029",  # BUPA — puntos pendientes de instalación
    "000010",  # Corporación Puente Alto — colegios fuera del reporte en cero
}

EXCLUDED_COMPANY_IDS_PUNTOS_EN_CERO: Set[str] = (
    EXCLUDED_COMPANY_IDS | EXCLUDED_COMPANY_IDS_SOLO_PUNTOS_EN_CERO
)

# Palabras clave de nombres de empresas a excluir (comparación case-insensitive)
EXCLUDED_COMPANY_NAME_KEYWORDS: Set[str] = {
    "MOP",
    "MINISTERIO DE OBRAS PUBLICAS",
    "MINISTERIO DE O.P",
    "MINISTERIO DE O. P",
    "MINISTERIO DE OBRAS PÚBLICAS",
    "LO BOZA",
    "TRES MONTES LUCCHETTI",
    "MADECO",
}

# IDs de nodos fijos a excluir (técnico / histórico). Los dados de baja por cliente o
# operación con detalle quedan en registro_puntos_deshabilitados.txt y se unen abajo.
_EXCLUDED_NODE_IDS_CORE: Set[str] = {
    "000013-01",  # Plaza Boulevard Pajaros Sur
    "000016-01",  # SCL Rebeca Matte Bello
    "000017-01",  # Rebeca Matte Bello
    "000025-03",  # Poniente 7
    "000025-05",  # Locales de Comida
    "000025-06",  # KFC
    "000011-01",  # Casa Juan Lopez
    "000012-01",  # Lo Boza Lavado de Vehículos
    "000012-02",  # Lo Boza Pozo
    "000012-03",  # Lo Boza Reutilización
    "000012-04",  # Lo Boza Edificio Principal Casino
    "000012-05",  # Lo Boza Matriz Principal
    "000012-20",  # La Cabaña (DERCO) — no opera / no incluir en puntos en cero
    "000019-01",  # Oficina WES
    "000021-08",  # Rugby CDUC
    "000006-03",  # Excluir solicitado
    "000017-03",  # Excluir solicitado
    "000017-02",  # Excluir solicitado
    "000007-08",  # Excluir solicitado
    "000002-02",  # Excluir solicitado
    "000021-02",  # Excluir solicitado
    "000025-11",  # Excluir solicitado
    "000025-14",  # Excluir solicitado
    "000027-05",  # Riego Fundo Zapallar
    "000007-09",  # Control Nido de Aguilas
}

EXCLUDED_NODE_IDS: Set[str] = set(_EXCLUDED_NODE_IDS_CORE) | _node_ids_desde_registro_puntos_deshabilitados()

# Solo «puntos en cero» (no afecta agregados ni individuales salvo que el ID también esté en el registro).
EXCLUDED_NODE_IDS_SOLO_PUNTOS_EN_CERO: Set[str] = {
    "000025-30",  # Matriz A.A — punto reubicado (Parque Arauco)
}

EXCLUDED_NODE_IDS_PUNTOS_EN_CERO: Set[str] = EXCLUDED_NODE_IDS | EXCLUDED_NODE_IDS_SOLO_PUNTOS_EN_CERO

# Fundo Zapallar (000027): 8 puntos de monitoreo en reportes WES.
# 000027-05 (Riego) sigue en EXCLUDED_NODE_IDS y no se incluye en agregados/individuales por esta lista.
FUNDO_ZAPALLAR_NODE_IDS: List[str] = [
    "000027-01",  # Matriz ESVAL
    "000027-02",  # Estanque Inferior
    "000027-03",  # Etapa N°5
    "000027-04",  # Etapa N°1 al 4
    "000027-06",  # Etapa N°1
    "000027-07",  # Etapa N°2
    "000027-08",  # Etapa N°3
    "000027-09",  # Riego Llenado de Estanque ESVAL
]


def is_company_excluded(company_id: Optional[str], company_name: Optional[str] = None) -> bool:
    if company_id and company_id in EXCLUDED_COMPANY_IDS:
        return True
    if company_name:
        company_name_upper = company_name.upper().strip()
        return any(keyword in company_name_upper for keyword in EXCLUDED_COMPANY_NAME_KEYWORDS)
    return False


def is_node_excluded(
    node_id: Optional[str],
    company_id: Optional[str] = None,
    company_name: Optional[str] = None,
) -> bool:
    if not node_id:
        return False
    if node_id in EXCLUDED_NODE_IDS:
        return True
    if not company_id and "-" in node_id:
        company_id = node_id.split("-", 1)[0]
    return is_company_excluded(company_id, company_name)


def filter_node_ids(
    node_ids: Iterable[str],
    company_id: Optional[str] = None,
    company_name: Optional[str] = None,
) -> List[str]:
    return [
        node_id
        for node_id in node_ids
        if not is_node_excluded(node_id, company_id=company_id, company_name=company_name)
    ]
