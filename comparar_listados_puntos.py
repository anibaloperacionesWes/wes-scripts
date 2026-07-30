"""
Compara el listado de puntos del proceso "puntos en cero" vs. reportes individuales.
Conecta a la API, detecta puntos nuevos y actualiza un cache local de nombres.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, List, Tuple

import requests

from reporte_puntos_en_cero import obtener_todos_los_nodos
from generar_reporte_word import NODE_NAMES


ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
CACHE_FILE = Path("node_names_api_cache.json")


def obtener_nodos_api() -> Dict[str, str]:
    """
    Obtiene todos los nodos desde la API de entidades.

    Returns:
        Dict[nodeId -> nodeName]
    """
    url = f"{ENTITY_BASE_URL}/configuration/companies"
    response = requests.get(url, timeout=20)
    response.raise_for_status()
    empresas = response.json()

    nodos_api: Dict[str, str] = {}
    for empresa in empresas:
        company_id = empresa.get("companyId")
        if not company_id:
            continue
        detail_url = f"{ENTITY_BASE_URL}/companies/{company_id}"
        detail_resp = requests.get(detail_url, timeout=20)
        if detail_resp.status_code != 200:
            continue
        data = detail_resp.json()
        for node in data.get("nodes", []):
            node_id = node.get("nodeId")
            node_name = (node.get("name") or "").strip()
            if node_id and node_name:
                nodos_api[node_id] = node_name

    return nodos_api


def cargar_cache() -> Dict[str, str]:
    if not CACHE_FILE.exists():
        return {}
    try:
        return json.loads(CACHE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def guardar_cache(cache: Dict[str, str]) -> None:
    CACHE_FILE.write_text(
        json.dumps(cache, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def main() -> None:
    print("=" * 70)
    print("COMPARACION DE LISTADOS - PUNTOS EN CERO vs INDIVIDUALES")
    print("=" * 70)

    print("[INFO] Conectando a la API para obtener nodos...")
    nodos_api = obtener_nodos_api()
    print(f"[OK] Nodos desde API: {len(nodos_api)}")

    print("[INFO] Obteniendo listado del proceso 'puntos en cero'...")
    nodos_puntos_cero = obtener_todos_los_nodos()
    ids_puntos_cero = {n["nodeId"] for n in nodos_puntos_cero}
    print(f"[OK] Nodos en puntos en cero (con exclusiones): {len(ids_puntos_cero)}")

    print("[INFO] Cargando listado de reportes individuales (NODE_NAMES + cache)...")
    cache = cargar_cache()
    listado_individuales = {**NODE_NAMES, **cache}
    ids_individuales = set(listado_individuales.keys())
    print(f"[OK] Nodos en individuales: {len(ids_individuales)}")

    # Comparaciones
    nuevos_para_individuales = sorted(ids_puntos_cero - ids_individuales)
    nuevos_para_puntos_cero = sorted(set(nodos_api.keys()) - ids_puntos_cero)

    # Actualizar cache con nuevos nodos para individuales (solo los que están en puntos en cero)
    nuevos_cache = False
    for node_id in nuevos_para_individuales:
        nombre = nodos_api.get(node_id)
        if nombre:
            cache[node_id] = nombre
            nuevos_cache = True

    if nuevos_cache:
        guardar_cache(cache)

    print()
    print("=" * 70)
    print("RESUMEN FINAL")
    print("=" * 70)
    print(f"Total nodos API: {len(nodos_api)}")
    print(f"Nodos proceso puntos en cero: {len(ids_puntos_cero)}")
    print(f"Nodos proceso individuales: {len(ids_individuales)}")
    print(f"Nuevos para individuales (agregados al cache): {len(nuevos_para_individuales)}")
    print(f"Nuevos para puntos en cero (API vs puntos en cero): {len(nuevos_para_puntos_cero)}")

    if nuevos_para_individuales:
        print("\nPuntos nuevos agregados a individuales (nodeId - nombre):")
        for node_id in nuevos_para_individuales:
            nombre = nodos_api.get(node_id, "(sin nombre en API)")
            print(f"  - {node_id} - {nombre}")

    if nuevos_para_puntos_cero:
        print("\nPuntos en API no incluidos en puntos en cero (revisar exclusiones):")
        for node_id in nuevos_para_puntos_cero:
            nombre = nodos_api.get(node_id, "(sin nombre en API)")
            print(f"  - {node_id} - {nombre}")

    if nuevos_cache:
        print(f"\n[OK] Cache actualizado: {CACHE_FILE}")
    else:
        print("\n[OK] Cache sin cambios (no hay puntos nuevos para individuales).")


if __name__ == "__main__":
    main()
