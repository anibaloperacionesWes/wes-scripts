"""

Parque Arauco — Mall Maipú: genera reporte Word agregado, antepone

«3) ¿Qué hallazgos encontramos? (Análisis de consumo)» y envía por correo.

"""



from __future__ import annotations



import argparse

from pathlib import Path



from exclusiones_reportes import filter_node_ids

from generar_pa_agregado_todos_puntos import _parrafo_hallazgos_q3

from generar_reporte_word import generate_aggregated_report, get_company_name

from generar_reportes_y_ppt_mall_maipu import get_maipu_nodes, obtener_datos_agregados

from pa_hallazgos_word_helpers import enviar_anibal_adjuntos, prepend_q3_hallazgos





def main() -> int:

    parser = argparse.ArgumentParser(

        description="PA Mall Maipú: hallazgos consumo + Word agregado + correo"

    )

    parser.add_argument("--desde", default="01/03/2026", help="DD/MM/YYYY")

    parser.add_argument("--hasta", default="27/03/2026", help="DD/MM/YYYY")

    parser.add_argument(

        "--no-email",

        action="store_true",

        help="Solo generar Word, no enviar correo",

    )

    args = parser.parse_args()



    company_id = "000025"

    cname = get_company_name(company_id)

    raw_nodes = get_maipu_nodes(company_id)

    node_ids = sorted({str(n.get("nodeId", "")).strip() for n in raw_nodes if n.get("nodeId")})

    node_ids = filter_node_ids(node_ids, company_id=company_id, company_name=cname)

    if not node_ids:

        print("[ERROR] Sin nodos Mall Maipú tras exclusiones.")

        return 1



    periodo = f"{args.desde} a {args.hasta}"

    print(f"[INFO] Maipú: {len(node_ids)} nodos {node_ids}, período {periodo}")



    out = generate_aggregated_report(

        company_id=company_id,

        node_ids=node_ids,

        start_date=args.desde,

        end_date=args.hasta,

        output_dir="reports",

        fuente_agua_id=None,

        mall_name="Maipú",

        apply_exclusions=False,

        generate_ppt=False,

    )

    out_path = Path(out)

    print(f"[OK] Reporte base: {out_path}")



    datos = obtener_datos_agregados(node_ids, args.desde, args.hasta)

    sujeto = (

        "Parque Arauco Mall Maipú (puntos de monitoreo del mall incluidos en este reporte)"

    )

    q3 = _parrafo_hallazgos_q3(datos, periodo, alcance_sujeto=sujeto)

    print("[INFO] Insertando bloque de hallazgos (Q3) al inicio del Word...")

    prepend_q3_hallazgos(out_path, q3)

    print(f"[OK] Documento actualizado: {out_path}")



    if not args.no_email:

        asunto = f"Parque Arauco Mall Maipú — Hallazgos consumo ({periodo})"

        cuerpo = (

            "Adjunto reporte Word agregado Mall Maipú con la sección "

            "«¿Qué hallazgos encontramos? (Análisis de consumo)» al inicio.\n\n"

            f"Puntos incluidos: {', '.join(node_ids)}.\n"

            f"Período: {periodo}."

        )

        enviar_anibal_adjuntos([out_path], asunto, cuerpo)

        print("[OK] Correo enviado a anibal.aoperaciones@wes.cl")

    return 0





if __name__ == "__main__":

    raise SystemExit(main())

