"""
Descarga ``GET .../nodes/{id}/dates.measures.csv`` **mes a mes** por cada colegio
(Corporacion Puente Alto, empresa 000010). Orden de nodos: ``nodeId`` ascendente,
empezando por **000010-01** (Escuela Andes del Sur).

Salida por defecto::
  reports/proyeccion ahorre puente 2025/csv_mensual_por_nodo/<node_id>/<AAAA-MM>_dates.measures.csv

Ejemplos::

  python descargar_csv_mensual_puente_alto.py --year 2025 --mes-desde 1 --mes-fin 12

  python descargar_csv_mensual_puente_alto.py --year 2025 --solo-node-id 000010-08 --mes-fin 3

Variable opcional: ``WES_API_BASE_URL`` (misma que otros scripts acl-node).
"""
from __future__ import annotations

import argparse
import calendar
import sys
from datetime import date
from pathlib import Path

import requests

from generar_reporte_word import acl_node_base_url
from reporte_puente_alto_lxm import obtener_nodos_puente_alto

ROOT = Path(__file__).resolve().parent
OUT_DEFAULT = ROOT / "reports" / "proyeccion ahorre puente 2025" / "csv_mensual_por_nodo"


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Descarga CSV mensual dates.measures.csv por nodo Puente Alto (desde 000010-01)."
    )
    ap.add_argument("--year", type=int, default=2025)
    ap.add_argument("--mes-desde", type=int, default=1)
    ap.add_argument("--mes-fin", type=int, default=12)
    ap.add_argument(
        "--desde-node-id",
        default="000010-01",
        help="Incluir solo nodos con nodeId >= este (default 000010-01 = todos los PA en orden).",
    )
    ap.add_argument(
        "--solo-node-id",
        default=None,
        help="Si se indica, solo ese nodo (ej. 000010-08).",
    )
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=OUT_DEFAULT,
        help="Carpeta base (se crean subcarpetas por node_id).",
    )
    args = ap.parse_args()

    year = args.year
    m0 = max(1, min(12, args.mes_desde))
    m1 = max(1, min(12, args.mes_fin))
    if m0 > m1:
        print("[ERROR] mes-desde no puede ser mayor que mes-fin.", file=sys.stderr)
        return 1

    out_base = Path(args.out_dir).expanduser().resolve()
    out_base.mkdir(parents=True, exist_ok=True)

    nodos = obtener_nodos_puente_alto()
    nodos.sort(key=lambda x: x["nodeId"])
    pref = args.desde_node_id.strip()
    nodos = [n for n in nodos if n["nodeId"] >= pref]
    if args.solo_node_id:
        sid = args.solo_node_id.strip()
        nodos = [n for n in nodos if n["nodeId"] == sid]
        if not nodos:
            print(f"[ERROR] No existe el nodo {sid} en empresa 000010.", file=sys.stderr)
            return 1

    sess = requests.Session()
    base = acl_node_base_url()
    total_arch = 0
    print(
        f"[INFO] Ano={year} meses {m0}-{m1} | nodos={len(nodos)} | salida={out_base}",
        flush=True,
    )

    for n in nodos:
        nid = n["nodeId"]
        nombre = n["nodeName"]
        dir_n = out_base / nid
        dir_n.mkdir(parents=True, exist_ok=True)

        for mes in range(m0, m1 + 1):
            ult = calendar.monthrange(year, mes)[1]
            ini = date(year, mes, 1)
            fin = date(year, mes, ult)
            url = f"{base}/nodes/{nid}/dates.measures.csv"
            params = [("start", ini.strftime("%d%m%Y")), ("end", fin.strftime("%d%m%Y"))]
            try:
                r = sess.get(url, params=params, timeout=240)
                r.raise_for_status()
                body = r.text
            except Exception as ex:
                print(f"[ERROR] {nid} {year}-{mes:02d}: {ex}", flush=True)
                continue

            fname = f"{year}-{mes:02d}_dates.measures.csv"
            dest = dir_n / fname
            dest.write_text(body, encoding="utf-8")
            total_arch += 1
            nl = body.count("\n") + (1 if body and not body.endswith("\n") else 0)
            print(
                f"[OK] {nid} | {nombre[:40]:<40} | {year}-{mes:02d} | "
                f"{len(body)} bytes | ~{nl} lineas -> {dest.name}",
                flush=True,
            )

    print(f"[LISTO] Archivos escritos: {total_arch}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
