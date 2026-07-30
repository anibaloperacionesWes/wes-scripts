"""
Parque Arauco (000025): por cada mall, reporte Word agregado + bloque Q3 hallazgos
y un correo con todos los adjuntos (un .docx por mall).

Uso:
  python generar_pa_todos_malls_hallazgos_y_enviar.py
  python generar_pa_todos_malls_hallazgos_y_enviar.py --no-email
  python generar_pa_todos_malls_hallazgos_y_enviar.py --solo-mall "Maipú"
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from pathlib import Path
from typing import Dict, List

import requests

ENTITY_BASE = "http://104.248.53.141:7001/wes/api/acl-entities/v1"

from exclusiones_reportes import filter_node_ids
from generar_pa_agregado_todos_puntos import _parrafo_hallazgos_q3
from pa_nodos_inactivos_por_mall import filtrar_nodos_activos_mall
from generar_reporte_word import generate_aggregated_report, get_company_name, get_mall_name_for_parque_arauco
from generar_reportes_y_ppt_mall_maipu import obtener_datos_agregados
from pa_hallazgos_word_helpers import enviar_anibal_adjuntos, prepend_q3_hallazgos


def _cargar_nodos_pa() -> dict:
    r = requests.get(f"{ENTITY_BASE}/companies/000025", timeout=60)
    r.raise_for_status()
    return r.json()


def _nodos_por_mall(company_id: str, company_name: str) -> Dict[str, List[str]]:
    payload = _cargar_nodos_pa()
    raw = payload.get("nodes") or []
    por_mall: Dict[str, List[str]] = defaultdict(list)
    for n in raw:
        nid = str(n.get("nodeId", "")).strip()
        if not nid:
            continue
        nombre = str(n.get("name", "") or "").strip()
        mall = get_mall_name_for_parque_arauco(nid, nombre).strip()
        if not mall:
            print(f"[SKIP] Sin mall en mapeo: {nid} ({nombre})")
            continue
        por_mall[mall].append(nid)

    out: Dict[str, List[str]] = {}
    for mall, ids in por_mall.items():
        filtrados = filter_node_ids(sorted(set(ids)), company_id=company_id, company_name=company_name)
        filtrados = filtrar_nodos_activos_mall(mall, filtrados)
        if filtrados:
            out[mall] = sorted(filtrados)
        else:
            print(f"[SKIP] Mall {mall}: todos los nodos excluidos por configuración.")
    return dict(sorted(out.items(), key=lambda x: x[0].casefold()))


def main() -> int:
    parser = argparse.ArgumentParser(
        description="PA: hallazgos Q3 + Word agregado por mall (correo único con adjuntos)"
    )
    parser.add_argument("--desde", default="01/03/2026", help="DD/MM/YYYY")
    parser.add_argument("--hasta", default="27/03/2026", help="DD/MM/YYYY")
    parser.add_argument("--no-email", action="store_true", help="Solo generar Word")
    parser.add_argument(
        "--solo-mall",
        default="",
        help='Procesar solo un mall (ej. "Maipú", "Estación")',
    )
    args = parser.parse_args()

    company_id = "000025"
    cname = get_company_name(company_id)
    periodo = f"{args.desde} a {args.hasta}"

    por_mall = _nodos_por_mall(company_id, cname)
    if args.solo_mall.strip():
        clave = args.solo_mall.strip()
        if clave not in por_mall:
            print(f"[ERROR] Mall no encontrado o sin nodos: {clave!r}. Disponibles: {list(por_mall.keys())}")
            return 1
        por_mall = {clave: por_mall[clave]}

    if not por_mall:
        print("[ERROR] No hay malls con nodos tras exclusiones.")
        return 1

    generados: List[Path] = []
    lineas_cuerpo: List[str] = [
        "Reportes Word agregados Parque Arauco — un archivo por mall.",
        "Cada documento incluye al inicio «3) ¿Qué hallazgos encontramos? (Análisis de consumo)».",
        "",
        f"Período: {periodo}.",
        "",
    ]

    for mall, node_ids in por_mall.items():
        print(f"\n[INFO] === Mall {mall} — {len(node_ids)} nodos: {node_ids} ===")
        out = generate_aggregated_report(
            company_id=company_id,
            node_ids=node_ids,
            start_date=args.desde,
            end_date=args.hasta,
            output_dir="reports",
            fuente_agua_id=None,
            mall_name=mall,
            apply_exclusions=False,
            generate_ppt=False,
        )
        out_path = Path(out)
        print(f"[OK] Reporte base: {out_path}")

        datos = obtener_datos_agregados(node_ids, args.desde, args.hasta)
        sujeto = (
            f"Parque Arauco Mall {mall} (puntos de monitoreo del mall incluidos en este reporte)"
        )
        q3 = _parrafo_hallazgos_q3(datos, periodo, alcance_sujeto=sujeto)
        prepend_q3_hallazgos(out_path, q3)
        generados.append(out_path)
        lineas_cuerpo.append(f"• {mall}: {', '.join(node_ids)} → {out_path.name}")

    print(f"\n[OK] {len(generados)} documento(s) generado(s).")

    if not args.no_email:
        asunto = f"Parque Arauco — Hallazgos consumo por mall ({periodo})"
        cuerpo = "\n".join(lineas_cuerpo)
        enviar_anibal_adjuntos(generados, asunto, cuerpo)
        print("[OK] Correo enviado a anibal.aoperaciones@wes.cl (adjuntos: todos los malls).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
