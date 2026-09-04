"""
Puntos con control hídrico (CPA / WES).

Fuente principal: Excel de horarios de corte
(`reports/HORARIOS CONTROL NOCTURNO.xlsx`).
Si el Excel no está en el entorno, se usa el inventario operativo conocido
(mismos IDs que el reporte de control nocturno).

Un punto CON CONTROL tiene equipo CPA/WES (instalado). Aunque falte
activarlo, cuenta como con control: el recinto ya tiene la máquina.
Lo Valledor (P1) tiene CPA instalado → CON CONTROL.
"""

from __future__ import annotations

from functools import lru_cache
from typing import Dict, FrozenSet, Optional, Tuple

# Inventario de respaldo: IDs WES con horario de corte (control nocturno).
# CORMUP 000008-02 (E. de la Barra), 000008-08 (Alicura) y 000008-13
# (Likankura) no están: esos colegios operan sin control.
CONTROL_NOCTURNO_IDS: FrozenSet[str] = frozenset(
    {
        "000006-01",
        "000006-02",
        "000006-04",
        "000008-01",
        "000008-03",
        "000008-04",
        "000008-05",
        "000008-06",
        "000008-07",
        "000008-09",
        "000008-10",
        "000008-11",
        "000008-12",
        "000008-14",
        "000002-01",
        "000012-06",
        "000017-04",
        "000017-07",
        "000017-08",
        "000020-02",
        "000022-00",
        "000024-01",
        "000025-01",
        "000025-02",
        "000025-04",
        "000025-07",
        "000025-08",
        "000025-09",
        "000025-10",
        "000025-12",
        "000025-13",
        "000025-15",
        "000025-17",
        "000025-18",
        "000025-19",
        "000025-20",
        "000025-21",
        "000025-22",
        "000025-23",
        "000025-24",
        "000025-25",
        "000025-27",
        "000025-28",
        "000025-29",
        "000025-30",
        "000028-01",
    }
)


@lru_cache(maxsize=1)
def ids_con_control() -> FrozenSet[str]:
    try:
        from control_nocturno import cargar_targets_desde_excel, default_excel_path

        path = default_excel_path()
        if path.is_file():
            return frozenset(cargar_targets_desde_excel(path).keys())
    except Exception:
        pass
    return CONTROL_NOCTURNO_IDS


def estado_control(cfg: dict, node_id: Optional[str]) -> Tuple[str, str, bool]:
    """
    Returns (etiqueta, detalle, tiene_control_activo).

    CPA instalado y no operando cuenta como SIN CONTROL.
    """
    nid = str(node_id or "").strip()
    if cfg.get("cpa_estado") == "instalado_pendiente":
        return "SIN CONTROL", "CPA instalado, no opera", False
    if nid and nid in ids_con_control():
        return "CON CONTROL", "CPA/WES activo", True
    if cfg.get("nocturnal_explain") == "wes" and nid == (cfg.get("matriz_id") or ""):
        return "CON CONTROL", "WES en la matriz", True
    return "SIN CONTROL", "Sin CPA/WES", False


def nota_red_cliente(cfg: dict, node_id: Optional[str], tiene: bool) -> str:
    """Si el cliente tiene otros puntos con control y este no, lo dice."""
    if tiene:
        return ""
    ids = ids_con_control()
    siblings = [n for n in (cfg.get("node_ids") or []) if n in ids and n != node_id]
    if siblings:
        return "Otros puntos del mismo cliente sí tienen control."
    return ""
