"""
Puntos Parque Arauco dados de baja por mall: no deben entrar en agregados Word,
gráficas comparativas ni datos agregados para PPT.

Claves de mall: mismas que get_mall_name_for_parque_arauco ("Quilicura", "El Bosque").
"""

from __future__ import annotations

from typing import Dict, List, Optional

NODOS_INACTIVOS_POR_MALL: Dict[str, frozenset[str]] = {
    "Quilicura": frozenset({"000025-14"}),  # Red de Incendio
    "El Bosque": frozenset({"000025-11"}),  # Matriz principal 1° piso
    "Kennedy": frozenset(
        {
            "000025-25",  # Baño N°5 Damas — retirado 30-03-2026
            "000025-26",  # Baño N°6 Varones — retirado 30-03-2026
        }
    ),
}

# Nombre de carpeta / slug -> clave en NODOS_INACTIVOS_POR_MALL y get_mall_name_for_parque_arauco
_SLUG_A_MALL: Dict[str, str] = {
    "El_Bosque": "El Bosque",
    "Quilicura": "Quilicura",
    "Maipú": "Maipú",
    "Maipu": "Maipú",
    "Estación": "Estación",
    "Estacion": "Estación",
    "Curauma": "Curauma",
    "Buenaventura": "Buenaventura",
    "Kennedy": "Kennedy",
}


def mall_clave_desde_argumento(mall_name_arg: Optional[str]) -> str:
    """Normaliza mall_name pasado a generate_aggregated_report (puede ser slug El_Bosque)."""
    if not mall_name_arg or not str(mall_name_arg).strip():
        return ""
    s = str(mall_name_arg).strip()
    if s in _SLUG_A_MALL:
        return _SLUG_A_MALL[s]
    return s


def filtrar_nodos_activos_mall(mall: str, node_ids: List[str]) -> List[str]:
    """Quita IDs inactivos para ese mall; conserva orden estable (ordenado)."""
    mall_k = mall_clave_desde_argumento(mall) or mall
    drop = NODOS_INACTIVOS_POR_MALL.get(mall_k, frozenset())
    return sorted({n for n in node_ids if n not in drop})


def aplicar_bajas_mall_pa(mall_name_arg: Optional[str], node_ids: List[str]) -> List[str]:
    """
    Usar en agregados PA (000025): quita siempre puntos dados de baja por mall.
    Resuelve mall desde el argumento (slug o nombre) o desde el primer nodo.
    """
    if not node_ids:
        return []

    mall_k = mall_clave_desde_argumento(mall_name_arg)
    if not mall_k:
        from generar_reporte_word import get_mall_name_for_parque_arauco, get_node_name

        mall_k = get_mall_name_for_parque_arauco(node_ids[0], get_node_name(node_ids[0])) or ""

    if not mall_k:
        return sorted(set(node_ids))

    antes = set(node_ids)
    out = filtrar_nodos_activos_mall(mall_k, list(node_ids))
    dropped = sorted(antes - set(out))
    if dropped:
        print(
            f"[INFO] Parque Arauco — excluidos del agregado (baja operativa, mall {mall_k}): "
            f"{', '.join(dropped)}"
        )
    return out
