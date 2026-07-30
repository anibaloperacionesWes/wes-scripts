import argparse
import requests
from collections import defaultdict
from exclusiones_reportes import filter_node_ids
from generar_reporte_word import generate_report, generate_aggregated_report, get_mall_name_for_parque_arauco
from pa_nodos_inactivos_por_mall import filtrar_nodos_activos_mall
from generar_reportes_y_ppt_mall_maipu import generar_ppt_desde_agregado

COMPANY_ID = "000025"  # Parque Arauco
START_DATE = "01/02/2026"
END_DATE = "28/02/2026"
ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"

print("Conectando a la API para obtener nodos de Parque Arauco...")
url = f"{ENTITY_BASE_URL}/companies/{COMPANY_ID}"
resp = requests.get(url, timeout=15)
resp.raise_for_status()
data = resp.json()

nodes = data.get("nodes", [])
mall_nodes = defaultdict(list)
for node in nodes:
    node_id = node.get("nodeId")
    node_name = (node.get("name") or "").strip()
    if not node_id:
        continue
    mall_name = get_mall_name_for_parque_arauco(node_id, node_name)
    if mall_name:
        mall_nodes[mall_name].append(node_id)

print(f"Malls detectados: {', '.join(sorted(mall_nodes.keys()))}")

for mall_name in sorted(mall_nodes.keys()):
    node_ids = filter_node_ids(mall_nodes[mall_name], company_id=COMPANY_ID, company_name="Parque Arauco")
    node_ids = filtrar_nodos_activos_mall(mall_name, node_ids)
    if not node_ids:
        print(f"[INFO] Mall {mall_name}: sin nodos tras exclusiones, se omite")
        continue

    print("\n" + "=" * 70)
    print(f"Procesando mall: {mall_name} ({len(node_ids)} nodos)")
    print("=" * 70)

    # Reportes individuales
    for node_id in node_ids:
        try:
            args = argparse.Namespace(
                company_id=COMPANY_ID,
                node_id=node_id,
                start_date=START_DATE,
                end_date=END_DATE,
                output_dir="reports",
                enviar_correo=False,
                destinatario=None,
                smtp_servidor=None,
                smtp_puerto=None,
                smtp_usuario=None,
                smtp_password=None,
            )
            path = generate_report(args)
            print(f"[OK] Individual: {path}")
        except Exception as e:
            print(f"[ADVERTENCIA] No se pudo generar individual {node_id}: {e}")

    # Reporte agregado + PPT
    try:
        agg_path = generate_aggregated_report(
            company_id=COMPANY_ID,
            node_ids=node_ids,
            start_date=START_DATE,
            end_date=END_DATE,
            output_dir="reports",
            mall_name=mall_name,
        )
        print(f"[OK] Agregado: {agg_path}")

        ppt_path = generar_ppt_desde_agregado(
            company_id=COMPANY_ID,
            node_ids=node_ids,
            start_date=START_DATE,
            end_date=END_DATE,
            aggregated_report_path=agg_path,
            mall_name=mall_name,
            company_name="Parque Arauco",
        )
        print(f"[OK] PPT: {ppt_path}")
    except Exception as e:
        print(f"[ADVERTENCIA] Error en agregado/PPT para {mall_name}: {e}")

print("\nPROCESO COMPLETADO")
