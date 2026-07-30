from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from types import SimpleNamespace

from generar_reporte_word import generate_aggregated_report, generate_report
from reporte_puente_alto_lxm import obtener_nodos_puente_alto


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Genera reportes por colegio + agregado para Corporacion Puente Alto."
    )
    ap.add_argument("--start-date", required=True, help="Fecha inicio YYYY-MM-DD")
    ap.add_argument("--end-date", required=True, help="Fecha fin YYYY-MM-DD")
    ap.add_argument("--output-dir", default="reports", help="Carpeta base de salida")
    ap.add_argument(
        "--workers",
        type=int,
        default=4,
        help="Cantidad de colegios en paralelo (default: 4).",
    )
    ap.add_argument(
        "--agg-workers",
        type=int,
        default=6,
        help="Workers internos para el reporte agregado (default: 6).",
    )
    args = ap.parse_args()

    company_id = "000010"
    nodos = sorted(obtener_nodos_puente_alto(), key=lambda n: n["nodeId"])
    node_ids = [n["nodeId"] for n in nodos]

    workers = max(1, min(int(args.workers), len(node_ids)))
    agg_workers = max(1, int(args.agg_workers))
    print(f"[INFO] Empresa {company_id}: {len(node_ids)} colegios")
    print(f"[INFO] Modo rápido: reportes por punto en paralelo (workers={workers})")

    def _one(node_id: str) -> tuple[str, str]:
        ns = SimpleNamespace(
            company_id=company_id,
            node_id=node_id,
            start_date=args.start_date,
            end_date=args.end_date,
            output_dir=args.output_dir,
            enviar_correo=False,
            destinatario=None,
            smtp_servidor="smtp.gmail.com",
            smtp_puerto=587,
            smtp_usuario=None,
            smtp_password=None,
        )
        out = generate_report(ns)
        return node_id, str(out)

    if workers == 1:
        for i, node_id in enumerate(node_ids, start=1):
            print(f"[{i}/{len(node_ids)}] Generando reporte {node_id} ...", flush=True)
            nid, out = _one(node_id)
            print(f"  [OK] {nid}: {out}", flush=True)
    else:
        with ThreadPoolExecutor(max_workers=workers) as ex:
            fut_map = {ex.submit(_one, nid): nid for nid in node_ids}
            done = 0
            for fut in as_completed(fut_map):
                nid = fut_map[fut]
                done += 1
                try:
                    _, out = fut.result()
                    print(f"[{done}/{len(node_ids)}] [OK] {nid}: {out}", flush=True)
                except Exception as e:
                    print(f"[{done}/{len(node_ids)}] [ERROR] {nid}: {e}", flush=True)

    print("[INFO] Generando reporte agregado ...", flush=True)
    out_ag = generate_aggregated_report(
        company_id=company_id,
        node_ids=node_ids,
        start_date=args.start_date,
        end_date=args.end_date,
        output_dir=args.output_dir,
        apply_exclusions=False,
        generate_ppt=False,
        parallel_node_fetch=True,
        max_parallel_workers=agg_workers,
    )
    print(f"[OK] Agregado: {out_ag}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
