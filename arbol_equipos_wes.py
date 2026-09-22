#!/usr/bin/env python3
"""
Árbol de equipos WES — modelo para el Dashboard.

Dos topologías de cliente:

1) puntos_separados
   Cada nodo es un punto independiente (colegio, clínica, local, etc.).
   No hay relación hidráulica padre→hijo entre ellos.
   Ejemplo: CORMUP, Providencia, La Florida, Renca colegios.

2) red_con_subredes
   Hay una red principal (matriz / medidor de entrada) y subredes o puntos
   aguas abajo. El agregado debe respetar la jerarquía para no doble-contar.
   Ejemplo: Fundo Zapallar (ESVAL → consumidores), Parque Arauco por mall.

Uso:
  python arbol_equipos_wes.py                  # imprime resumen
  python arbol_equipos_wes.py --export json    # escribe arbol_equipos_wes.json
  python arbol_equipos_wes.py --export json --salida dashboard/arbol_equipos_wes.json
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests

ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
ROOT = Path(__file__).resolve().parent
DEFAULT_OUT = ROOT / "arbol_equipos_wes.json"

# ---------------------------------------------------------------------------
# Topologías
# ---------------------------------------------------------------------------
TOPOLOGIA_PUNTOS = "puntos_separados"
TOPOLOGIA_RED = "red_con_subredes"

# Tipos de nodo en el árbol (UI / agregación)
TIPO_CLIENTE = "cliente"
TIPO_SITIO = "sitio"  # mall, fundo, sede
TIPO_RED = "red"  # red principal / matriz
TIPO_SUBRED = "subred"  # tramo o sector aguas abajo de una matriz
TIPO_PUNTO = "punto"  # hoja / medidor

# ---------------------------------------------------------------------------
# Overrides de jerarquía hidráulica (no vienen de la API hoy)
# parent = nodeId padre; si parent es None → raíz de esa red/sitio
# ---------------------------------------------------------------------------

# Fundo Zapallar: Matriz ESVAL alimenta el resto.
# Etapas 1–3 son detalle de Etapa N°1 al 4 (sub-sectorización).
HIERARCHY_FUNDO_ZAPALLAR: Dict[str, Optional[str]] = {
    "000027-01": None,  # Matriz ESVAL (raíz)
    "000027-02": "000027-01",  # Estanque Inferior
    "000027-03": "000027-01",  # Etapa N°5
    "000027-04": "000027-01",  # Etapa N°1 al 4
    "000027-05": "000027-01",  # Riego
    "000027-06": "000027-04",  # Etapa N°1
    "000027-07": "000027-04",  # Etapa N°2
    "000027-08": "000027-04",  # Etapa N°3
    "000027-09": "000027-01",  # Riego Llenado de Estanque ESVAL
}

# Nido de Águilas: Estanque C (sanitaria) carga B y A; B alimenta Teatro/HS/Elementary;
# Piscina sale del Estanque A (no monitoreado).
HIERARCHY_NIDO: Dict[str, Optional[str]] = {
    "000007-07": None,  # Estanque C (entrada desde medidor sanitaria)
    "000007-01": "000007-07",  # Estanque B
    "000007-02": "000007-01",  # Teatro
    "000007-03": "000007-01",  # High School
    "000007-04": "000007-01",  # Elementary
    "000007-05": "000007-07",  # Piscina (vía Estanque A no monitoreado)
    "000007-06": None,  # Pozo Profundo
    "000007-09": None,  # Control
}

NODE_NOTAS: Dict[str, str] = {
    "000007-07": (
        "Alimentado por medidor de la sanitaria; carga Estanque B y Estanque A "
        "(A no se monitorea)."
    ),
    "000007-05": (
        "Sale del Estanque A (no monitoreado), cargado desde Estanque C; "
        "otras dependencias no están claras."
    ),
    "000009-02": (
        "Solo alimenta el sistema de lavados cuando falla la ósmosis. "
        "El estanque de reutilización de ósmosis (90 %) no se monitorea."
    ),
    "000009-09": (
        "Se alimenta de la matriz; tiene ósmosis que recupera ~90 % hacia "
        "estanque de reutilización no monitoreado."
    ),
    "000009-10": (
        "Se alimenta de la matriz; tiene ósmosis que recupera ~90 % hacia "
        "estanque de reutilización no monitoreado."
    ),
    "000012-06": (
        "Matriz principal: alimenta toda la red Quilicura; "
        "los otros puntos representan una fracción mínima del total."
    ),
    "000025-22": "Alimenta Distrito de lujo (DL); a veces también Sandia Nueva.",
    "000025-28": "A veces abastece Distrito de lujo (DL) junto / en lugar de Sandia Antigua.",
    "000025-27": "Distrito de lujo — alimentado por Sandia Antigua (y a veces Nueva).",
}

# COPEC: Matriz Principal alimenta Costanera, admin, Pronto y lavados.
HIERARCHY_COPEC: Dict[str, Optional[str]] = {
    "000009-06": None,  # Matriz Principal
    "000009-00": "000009-06",  # Costanera
    "000009-01": "000009-06",  # Oficina Admin.
    "000009-08": "000009-06",  # Pronto Baños
    "000009-11": "000009-06",  # Pronto Tienda
    "000009-03": "000009-06",  # Lavado Automático Norte
    "000009-04": "000009-06",  # Lavado Automático Sur
    "000009-09": "000009-06",  # Lavado Auto servicio Norte
    "000009-10": "000009-06",  # Lavado Auto servicio Sur
    "000009-02": "000009-06",  # Estanque Reutilización (backup ósmosis)
    "000009-05": None,  # Riego (independiente / pendiente)
}

# Inchcape Quilicura: matriz alimenta el resto (fracción mínima cada uno).
HIERARCHY_INCHCAPE_QUILICURA: Dict[str, Optional[str]] = {
    "000012-06": None,
    "000012-07": "000012-06",
    "000012-08": "000012-06",
    "000012-09": "000012-06",
    "000012-10": "000012-06",
    "000012-11": "000012-06",
    "000012-12": "000012-06",
}

# Parque Arauco — Kennedy (PAK)
HIERARCHY_PA_KENNEDY: Dict[str, Optional[str]] = {
    "000025-20": None,  # Impulsión Anden 3-4 Matriz Principal
    "000025-21": "000025-20",  # Locales Gast.
    "000025-29": "000025-20",  # Restaurante
    # Sandia → DL → subredes
    "000025-22": None,  # Sala de Bomba Sandia Antigua
    "000025-27": "000025-22",  # Distrito de lujo DL
    "000025-35": "000025-27",  # PAK BAZAR GOURMET
    "000025-36": "000025-27",  # PAK DL KENNEDY
    "000025-28": None,  # Sandia Nueva (alternativo a Antigua)
    # Independientes
    "000025-23": None,  # Llenado Pileta
    "000025-24": None,  # Llenado Pileta Cascada
}

# Códigos operativos de mall Parque Arauco
PA_MALL_CODIGO: Dict[str, str] = {
    "Buenaventura": "BOM",
    "Curauma": "CUR",
    "Estación": "MAE",
    "El Bosque": "AEB",
    "Maipú": "MAM",
    "Kennedy": "PAK",
    "Quilicura": "MAQ",
}

DISPLAY_NAME_OVERRIDES: Dict[str, str] = {
    "000020-05": "Agunsa sucursal San Antonio",
}

# Mall → nodos (misma fuente que generar_reporte_word.get_mall_name_for_parque_arauco)
PA_MALL_BY_NODE: Dict[str, str] = {
    "000025-01": "Estación",
    "000025-02": "Estación",  # Abastecimiento Sur Terminal → MAE
    "000025-19": "Estación",
    "000025-03": "Estación",
    "000025-05": "Estación",
    "000025-06": "Estación",
    "000025-04": "Estación",
    "000025-07": "Estación",
    "000025-08": "Maipú",
    "000025-09": "Maipú",
    "000025-10": "Maipú",
    "000025-32": "Maipú",
    "000025-33": "Maipú",
    "000025-11": "El Bosque",
    "000025-12": "El Bosque",
    "000025-30": "El Bosque",  # Matriz A.A — NO es MAE; AEB
    "000025-13": "Quilicura",
    "000025-14": "Quilicura",
    "000025-34": "Quilicura",
    "000025-15": "Curauma",
    "000025-16": "Curauma",
    "000025-37": "Curauma",
    "000025-38": "Curauma",
    "000025-17": "Buenaventura",
    "000025-18": "Buenaventura",
    "000025-20": "Kennedy",
    "000025-21": "Kennedy",
    "000025-22": "Kennedy",
    "000025-23": "Kennedy",
    "000025-24": "Kennedy",
    "000025-25": "Kennedy",
    "000025-26": "Kennedy",
    "000025-27": "Kennedy",
    "000025-28": "Kennedy",
    "000025-29": "Kennedy",
    "000025-35": "Kennedy",
    "000025-36": "Kennedy",
}

# Empresa API con varios clientes operativos en el Dashboard (mismo companyId).
CLIENTES_MULTISITIO: Dict[str, Dict[str, Any]] = {
    "000029": {
        "expandir_como_clientes": True,
        "sitios": [
            {
                "id_suffix": "antofagasta",
                "name": "Bupa Antofagasta",
                "node_ids": [
                    "000029-07",
                    "000029-08",
                    "000029-09",
                    "000029-10",
                ],
                "topologia": TOPOLOGIA_PUNTOS,
                "estado_operativo": "activo",
                "nota": (
                    "Clínica Antofagasta operativa. Medidor Principal Sanitaria (09) "
                    "es referencia de cuenta; salas de bomba son puntos paralelos."
                ),
            },
        ],
    },
    "000012": {
        "nombre_dashboard": "Inchcape",
        "nombre_api": "DERCO",
        "expandir_como_clientes": False,
        "topologia": TOPOLOGIA_RED,
        "sitios": [
            {
                "id_suffix": "quilicura",
                "name": "Quilicura",
                "node_ids": [
                    "000012-06",
                    "000012-07",
                    "000012-08",
                    "000012-09",
                    "000012-10",
                    "000012-11",
                    "000012-12",
                ],
                "topologia": TOPOLOGIA_RED,
                "hierarchy": HIERARCHY_INCHCAPE_QUILICURA,
                "estado_operativo": "activo",
                "nota": "Matriz Principal alimenta la red; el resto es fracción mínima.",
            },
        ],
    },
    "000020": {
        "nombre_dashboard": "AGUNSA",
        "expandir_como_clientes": False,
        "topologia": TOPOLOGIA_PUNTOS,
        "sitios": [
            {
                "id_suffix": "lampa",
                "name": "Lampa",
                "node_ids": ["000020-01", "000020-02", "000020-03", "000020-04"],
                "estado_operativo": "activo",
            },
            {
                "id_suffix": "san-antonio",
                "name": "Agunsa sucursal San Antonio",
                "node_ids": ["000020-05"],
                "estado_operativo": "activo",
            },
        ],
    },
}

# Clientes cuyo default es red_con_subredes.
CLIENTES_RED: Dict[str, Dict[str, Any]] = {
    "000007": {
        "label": "Nido de Aguilas",
        "hierarchy": HIERARCHY_NIDO,
        "sitio_unico": "Nido de Águilas",
    },
    "000009": {
        "label": "COPEC",
        "hierarchy": HIERARCHY_COPEC,
        "sitio_unico": "COPEC Costanera",
    },
    "000027": {
        "label": "Fundo Zapallar",
        "hierarchy": HIERARCHY_FUNDO_ZAPALLAR,
        "sitio_unico": "Fundo Zapallar",
    },
    "000025": {
        "label": "Parque Arauco",
        "agrupar_por": "mall",
        "hierarchies_por_sitio": {
            "Kennedy": HIERARCHY_PA_KENNEDY,
        },
    },
    "000026": {
        "label": "UDD",
        "hierarchy": {},
        "sitio_unico": "UDD",
        "nota": "Red tipo anillo; jerarquía hidráulica pendiente de mapear.",
    },
}

CLIENTES_EXCLUIDOS_DASHBOARD = {
    "000000",  # Wes Spa
    "000001",  # Ejército de Chile
    "000004",  # Gendarmería
    "000005",  # MOP
    "000010",  # Corporación Puente Alto
    "000011",  # Sistemas Socios Wes
    "000013",  # Lo Barnechea
    "000014",  # Tres Montes Lucchetti
    "000016",  # Renca (SCL Rebeca Matte Bello)
    "000018",  # MADECO
    "000019",  # WESSPA
    "000023",  # MADECO
    "000030",  # Estadio Israelita Maccabi
}

# Nodos fuera del árbol Dashboard (revisión operativa).
NODOS_EXCLUIDOS_DASHBOARD = {
    "000002-02",  # Lo Valledor Pozo
    "000006-03",  # Arturo Alessandri Palma
    "000007-08",  # Nido Cancha
    "000017-01",  # Rebeca Matte Bello
    "000017-02",  # Juana Atala de Hirmas
    "000017-03",  # José Luis Araneda
    "000021-08",  # Rugby CDUC
    "000022-01",  # Juan Pablo II (Las Condes)
    # Parque Arauco — dados de baja / Curauma 15-16 (CUR anillos sí van)
    "000025-03",  # Poniente 7
    "000025-05",  # Locales de Comida
    "000025-06",  # KFC
    "000025-14",  # Quilicura Red de Incendio (no operativo)
    "000025-15",  # Curauma Matriz Principal
    "000025-16",  # Curauma Baños
    "000025-25",  # Baño N°5 Damas (retirado)
    "000025-26",  # Baño N°6 Varones (retirado)
    # Inchcape sitios fuera
    "000012-01",
    "000012-02",
    "000012-03",
    "000012-04",
    "000012-05",
    "000012-13",
    "000012-14",
    # Bupa Santiago
    "000029-01",
    "000029-02",
    "000029-03",
    "000029-04",
    "000029-05",
    "000029-06",
}

# Malls PA fuera del Dashboard (vacío: MAQ vuelve al árbol)
PA_MALLS_EXCLUIDOS: set = set()


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"Accept": "application/json"})
    return s


def fetch_companies(max_id: int = 80) -> List[Dict[str, Any]]:
    """Obtiene empresas y nodos desde la API de entidades."""
    sess = _session()
    out: List[Dict[str, Any]] = []
    for i in range(max_id):
        cid = f"{i:06d}"
        try:
            r = sess.get(f"{ENTITY_BASE_URL}/companies/{cid}", timeout=10)
            if r.status_code != 200:
                continue
            data = r.json()
            name = str(data.get("name") or "").strip()
            if not name:
                continue
            nodes = data.get("nodes") or []
            if not nodes:
                r2 = sess.get(f"{ENTITY_BASE_URL}/companies/{cid}/nodes", timeout=10)
                if r2.status_code == 200:
                    payload = r2.json()
                    nodes = payload if isinstance(payload, list) else payload.get("nodes", [])
            nlist = []
            for n in nodes:
                if not isinstance(n, dict):
                    continue
                nid = str(n.get("nodeId") or n.get("id") or "").strip()
                nname = str(n.get("name") or "").strip()
                if nid:
                    nlist.append({"nodeId": nid, "name": nname or nid})
            out.append({"companyId": cid, "name": name, "nodes": nlist})
        except requests.RequestException:
            continue
    return out


def _display_name(node_id: str, name: str) -> str:
    return DISPLAY_NAME_OVERRIDES.get(node_id, name)


def _attach_nota(node: Dict[str, Any]) -> Dict[str, Any]:
    nid = node.get("nodeId")
    if nid and nid in NODE_NOTAS:
        node["nota"] = NODE_NOTAS[nid]
    return node


def _filter_nodes(nodes: List[Dict[str, str]]) -> List[Dict[str, str]]:
    return [n for n in nodes if n.get("nodeId") not in NODOS_EXCLUIDOS_DASHBOARD]


def _punto_leaf(node_id: str, name: str, **extra: Any) -> Dict[str, Any]:
    leaf = {
        "id": node_id,
        "nodeId": node_id,
        "name": _display_name(node_id, name),
        "tipo": TIPO_PUNTO,
        "children": [],
        **extra,
    }
    return _attach_nota(leaf)


def _nodes_por_ids(
    company: Dict[str, Any],
    node_ids: List[str],
    nombres_fallback: Optional[Dict[str, str]] = None,
) -> List[Dict[str, str]]:
    """Filtra nodos de la empresa; usa fallback si el ID no viene en la API."""
    by_id = {n["nodeId"]: n for n in company.get("nodes", [])}
    out: List[Dict[str, str]] = []
    for nid in node_ids:
        if nid in NODOS_EXCLUIDOS_DASHBOARD:
            continue
        if nid in by_id:
            out.append(by_id[nid])
        elif nombres_fallback and nid in nombres_fallback:
            out.append({"nodeId": nid, "name": nombres_fallback[nid]})
    return out


def _sitio_desde_cfg(
    company_id: str,
    sitio_cfg: Dict[str, Any],
    nodos: List[Dict[str, str]],
) -> Dict[str, Any]:
    topo = sitio_cfg.get("topologia", TOPOLOGIA_PUNTOS)
    hier = sitio_cfg.get("hierarchy") or {}
    if topo == TOPOLOGIA_RED and hier:
        kids = _build_from_parent_map(nodos, hier)
    else:
        kids = [
            _punto_leaf(n["nodeId"], n["name"])
            for n in sorted(nodos, key=lambda x: x["nodeId"])
        ]
    sitio: Dict[str, Any] = {
        "id": f"{company_id}-{sitio_cfg['id_suffix']}",
        "name": sitio_cfg["name"],
        "tipo": TIPO_SITIO,
        "topologia": topo,
        "children": kids,
    }
    if sitio_cfg.get("nota"):
        sitio["nota"] = sitio_cfg["nota"]
    if sitio_cfg.get("estado_operativo"):
        sitio["estado_operativo"] = sitio_cfg["estado_operativo"]
    return sitio


def _cliente_multisitio(company: Dict[str, Any], cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Una o más entradas Dashboard para empresas API multi-sitio."""
    cid = company["companyId"]
    nombre_api = cfg.get("nombre_api") or company["name"]
    expandir = cfg.get("expandir_como_clientes", False)
    nombres_fb = cfg.get("nombres_fallback") or {}

    if expandir:
        clientes: List[Dict[str, Any]] = []
        for sitio_cfg in cfg["sitios"]:
            nodos = _nodes_por_ids(company, sitio_cfg["node_ids"], nombres_fb)
            sitio = _sitio_desde_cfg(cid, sitio_cfg, nodos)
            clientes.append(
                {
                    "id": f"{cid}-{sitio_cfg['id_suffix']}",
                    "companyId": cid,
                    "name": sitio_cfg["name"],
                    "nombre_api": nombre_api,
                    "tipo": TIPO_CLIENTE,
                    "topologia": sitio_cfg.get("topologia", TOPOLOGIA_PUNTOS),
                    "estado_operativo": sitio_cfg.get("estado_operativo", "activo"),
                    "descripcion": sitio_cfg.get("nota")
                    or f"Sitio operativo de {nombre_api}.",
                    "children": sitio["children"],
                }
            )
        return clientes

    nombre = cfg.get("nombre_dashboard") or company["name"]
    sitios = [
        _sitio_desde_cfg(
            cid,
            sitio_cfg,
            _nodes_por_ids(company, sitio_cfg["node_ids"], nombres_fb),
        )
        for sitio_cfg in cfg["sitios"]
    ]
    return [
        {
            "id": cid,
            "companyId": cid,
            "name": nombre,
            "nombre_api": nombre_api,
            "tipo": TIPO_CLIENTE,
            "topologia": cfg.get("topologia", TOPOLOGIA_PUNTOS),
            "descripcion": (
                f"Cliente {nombre} (empresa API {nombre_api}) con varios sitios."
            ),
            "children": sitios,
        }
    ]


def _build_from_parent_map(
    nodes: List[Dict[str, str]],
    parent_of: Dict[str, Optional[str]],
) -> List[Dict[str, Any]]:
    """Arma un bosque a partir de parent_of[nodeId] = parentNodeId | None."""
    nodes = _filter_nodes(nodes)
    by_id = {n["nodeId"]: n for n in nodes}
    allowed = set(by_id)

    # Si el padre quedó excluido, el nodo sube a raíz
    effective_parent: Dict[str, Optional[str]] = {}
    for nid in list(parent_of.keys()) + [n["nodeId"] for n in nodes]:
        if nid not in allowed:
            continue
        parent = parent_of.get(nid)
        while parent is not None and parent not in allowed:
            parent = parent_of.get(parent)
        if nid in parent_of or nid in allowed:
            effective_parent[nid] = parent if nid in parent_of else None

    for n in nodes:
        if n["nodeId"] not in effective_parent:
            effective_parent[n["nodeId"]] = None

    children_map: Dict[Optional[str], List[str]] = {}
    for nid, parent in effective_parent.items():
        children_map.setdefault(parent, []).append(nid)

    def build(nid: str) -> Dict[str, Any]:
        meta = by_id.get(nid, {"nodeId": nid, "name": nid})
        kids = children_map.get(nid, [])
        nombre = _display_name(nid, meta.get("name") or nid)
        if kids:
            node = {
                "id": nid,
                "nodeId": nid,
                "name": nombre,
                "tipo": TIPO_RED if effective_parent.get(nid) is None else TIPO_SUBRED,
                "children": [build(c) for c in sorted(kids)],
            }
            return _attach_nota(node)
        return _punto_leaf(nid, meta.get("name") or nid)

    roots = children_map.get(None, [])
    seen = set()
    ordered = []
    for nid in roots:
        if nid not in seen:
            seen.add(nid)
            ordered.append(nid)
    return [build(nid) for nid in sorted(ordered)]


def _cliente_puntos_separados(company: Dict[str, Any]) -> Dict[str, Any]:
    children = [
        _punto_leaf(n["nodeId"], n["name"])
        for n in sorted(_filter_nodes(company["nodes"]), key=lambda x: x["nodeId"])
    ]
    return {
        "id": company["companyId"],
        "companyId": company["companyId"],
        "name": company["name"],
        "tipo": TIPO_CLIENTE,
        "topologia": TOPOLOGIA_PUNTOS,
        "descripcion": (
            "Puntos separados: cada medidor es independiente; "
            "no hay subredes hidráulicas entre ellos."
        ),
        "children": children,
    }


def _cliente_fundo(company: Dict[str, Any], cfg: Dict[str, Any]) -> Dict[str, Any]:
    forest = _build_from_parent_map(_filter_nodes(company["nodes"]), cfg["hierarchy"])
    sitio = {
        "id": f"{company['companyId']}-sitio",
        "name": cfg.get("sitio_unico") or company["name"],
        "tipo": TIPO_SITIO,
        "children": forest,
    }
    return {
        "id": company["companyId"],
        "companyId": company["companyId"],
        "name": company["name"],
        "tipo": TIPO_CLIENTE,
        "topologia": TOPOLOGIA_RED,
        "descripcion": (
            "Red con subredes: hay matriz/entrada y puntos aguas abajo. "
            "Al agregar, no sumar hijos si ya está incluida la matriz padre."
        ),
        "children": [sitio],
    }


def _cliente_parque_arauco(company: Dict[str, Any], cfg: Dict[str, Any]) -> Dict[str, Any]:
    by_mall: Dict[str, List[Dict[str, str]]] = {}
    sin_mall: List[Dict[str, str]] = []
    for n in _filter_nodes(company["nodes"]):
        mall = PA_MALL_BY_NODE.get(n["nodeId"])
        if mall and mall in PA_MALLS_EXCLUIDOS:
            continue
        if mall:
            by_mall.setdefault(mall, []).append(n)
        else:
            sin_mall.append(n)

    sitios: List[Dict[str, Any]] = []
    hierarchies = cfg.get("hierarchies_por_sitio") or {}
    for mall in sorted(by_mall.keys()):
        nodos = by_mall[mall]
        hier = hierarchies.get(mall)
        codigo = PA_MALL_CODIGO.get(mall, "")
        mall_label = f"{codigo} — {mall}" if codigo else mall
        if hier:
            kids = _build_from_parent_map(nodos, hier)
            nota = "Incluye matriz y subredes hidráulicas mapeadas."
        else:
            kids = [
                _punto_leaf(n["nodeId"], n["name"], pendiente_jerarquia=True)
                for n in sorted(nodos, key=lambda x: x["nodeId"])
            ]
            nota = "Sitio con varios puntos; jerarquía hidráulica pendiente de mapear."
        sitios.append(
            {
                "id": f"{company['companyId']}-{mall.lower().replace(' ', '-')}",
                "name": mall_label,
                "codigo_mall": codigo or None,
                "tipo": TIPO_SITIO,
                "nota": nota,
                "children": kids,
            }
        )

    # sin_mall: no se incluye en Dashboard

    return {
        "id": company["companyId"],
        "companyId": company["companyId"],
        "name": company["name"],
        "tipo": TIPO_CLIENTE,
        "topologia": TOPOLOGIA_RED,
        "descripcion": (
            "Cliente multi-sitio (malls). Códigos: BOM, CUR, MAE, AEB, MAM, PAK, MAQ."
        ),
        "children": sitios,
    }


def _cliente_red_generico(company: Dict[str, Any], cfg: Dict[str, Any]) -> Dict[str, Any]:
    hier = cfg.get("hierarchy") or {}
    nodos = _filter_nodes(company["nodes"])
    if hier:
        kids = _build_from_parent_map(nodos, hier)
    else:
        kids = [
            _punto_leaf(n["nodeId"], n["name"], pendiente_jerarquia=True)
            for n in sorted(nodos, key=lambda x: x["nodeId"])
        ]
    sitio = {
        "id": f"{company['companyId']}-sitio",
        "name": cfg.get("sitio_unico") or company["name"],
        "tipo": TIPO_SITIO,
        "nota": cfg.get("nota"),
        "children": kids,
    }
    return {
        "id": company["companyId"],
        "companyId": company["companyId"],
        "name": company["name"],
        "tipo": TIPO_CLIENTE,
        "topologia": TOPOLOGIA_RED,
        "descripcion": (
            "Red con subredes (jerarquía parcial o pendiente). "
            "Completar parent map antes de agregar en Dashboard."
        ),
        "children": [sitio],
    }


def build_tree(companies: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
    """Construye el árbol completo listo para el Dashboard."""
    if companies is None:
        companies = fetch_companies()

    clientes: List[Dict[str, Any]] = []
    for company in sorted(companies, key=lambda c: c["companyId"]):
        cid = company["companyId"]
        if cid in CLIENTES_EXCLUIDOS_DASHBOARD:
            continue

        if cid in CLIENTES_MULTISITIO:
            clientes.extend(_cliente_multisitio(company, CLIENTES_MULTISITIO[cid]))
        elif cid in CLIENTES_RED:
            cfg = CLIENTES_RED[cid]
            if cid == "000025":
                clientes.append(_cliente_parque_arauco(company, cfg))
            elif cid in ("000027", "000007", "000009", "000026"):
                clientes.append(_cliente_red_generico(company, cfg) if cid != "000027" else _cliente_fundo(company, cfg))
            else:
                clientes.append(_cliente_red_generico(company, cfg))
        else:
            clientes.append(_cliente_puntos_separados(company))

    # Descartar entradas sin puntos hoja (todo excluido)
    clientes = [c for c in clientes if _count_leaves([c]) > 0]

    resumen = {
        "total_clientes": len(clientes),
        "puntos_separados": sum(1 for c in clientes if c["topologia"] == TOPOLOGIA_PUNTOS),
        "red_con_subredes": sum(1 for c in clientes if c["topologia"] == TOPOLOGIA_RED),
        "total_nodos_hoja": _count_leaves(clientes),
    }

    return {
        "version": 1,
        "generado": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "modelo": {
            "topologias": {
                TOPOLOGIA_PUNTOS: (
                    "Cliente con uno o varios medidores independientes "
                    "(sin relación hidráulica padre→hijo)."
                ),
                TOPOLOGIA_RED: (
                    "Cliente con red principal y subredes/puntos aguas abajo; "
                    "la agregación debe respetar la jerarquía."
                ),
            },
            "tipos_nodo": [TIPO_CLIENTE, TIPO_SITIO, TIPO_RED, TIPO_SUBRED, TIPO_PUNTO],
        },
        "resumen": resumen,
        "clientes": clientes,
    }


def _count_leaves(nodes: List[Dict[str, Any]]) -> int:
    total = 0
    for n in nodes:
        kids = n.get("children") or []
        if not kids and n.get("nodeId"):
            total += 1
        else:
            total += _count_leaves(kids)
    return total


def print_summary(tree: Dict[str, Any]) -> None:
    r = tree["resumen"]
    print("=" * 60)
    print("ÁRBOL DE EQUIPOS WES — Dashboard")
    print("=" * 60)
    print(f"Clientes:           {r['total_clientes']}")
    print(f"  puntos_separados: {r['puntos_separados']}")
    print(f"  red_con_subredes: {r['red_con_subredes']}")
    print(f"Nodos hoja:         {r['total_nodos_hoja']}")
    print("-" * 60)
    for c in tree["clientes"]:
        n_leaf = _count_leaves([c])
        flag = "RED " if c["topologia"] == TOPOLOGIA_RED else "PTO "
        print(f"  [{flag}] {c['companyId']}  {c['name']:<32}  hojas={n_leaf}")
    print("=" * 60)


def export_json(tree: Dict[str, Any], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(tree, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return path


def main() -> None:
    parser = argparse.ArgumentParser(description="Árbol de equipos WES para Dashboard")
    parser.add_argument(
        "--export",
        choices=["json"],
        help="Exportar árbol (json)",
    )
    parser.add_argument(
        "--salida",
        type=Path,
        default=DEFAULT_OUT,
        help=f"Ruta de salida (default: {DEFAULT_OUT.name})",
    )
    parser.add_argument(
        "--offline",
        type=Path,
        help="Usar JSON local de empresas/nodos en vez de llamar a la API",
    )
    args = parser.parse_args()

    if args.offline:
        companies = json.loads(args.offline.read_text(encoding="utf-8"))
    else:
        print("Consultando API de empresas/nodos...")
        companies = fetch_companies()
        print(f"Empresas con nombre: {len(companies)}")

    tree = build_tree(companies)
    print_summary(tree)

    if args.export == "json":
        out = export_json(tree, args.salida)
        print(f"\n[OK] Exportado: {out}")


if __name__ == "__main__":
    main()
