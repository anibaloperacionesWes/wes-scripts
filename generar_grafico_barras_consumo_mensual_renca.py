from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import matplotlib.pyplot as plt
import pandas as pd
import requests

from generar_consolidado_m3_mensual_puente_alto import consumo_mes_un_nodo

ENTITY_BASE = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
COMPANY_ID_RENCA = "000017"


def obtener_nodos_renca() -> List[Dict[str, str]]:
    r = requests.get(f"{ENTITY_BASE}/companies/{COMPANY_ID_RENCA}", timeout=30)
    r.raise_for_status()
    data = r.json()
    nodes = data.get("nodes") or []
    out: List[Dict[str, str]] = []
    for n in nodes:
        nid = str(n.get("nodeId", "")).strip()
        name = str(n.get("name", "")).strip()
        if nid and name:
            out.append({"nodeId": nid, "nodeName": name})
    out.sort(key=lambda x: x["nodeId"])
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="Grafico de barras consumo mensual puntos Renca")
    ap.add_argument("--year", type=int, default=datetime.now().year)
    ap.add_argument("--mes-desde", type=int, default=1)
    ap.add_argument("--mes-fin", type=int, default=datetime.now().month)
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=Path("reports") / "Renca",
        help="Carpeta de salida para PNG y CSV",
    )
    args = ap.parse_args()

    year = int(args.year)
    m0 = max(1, min(12, int(args.mes_desde)))
    m1 = max(1, min(12, int(args.mes_fin)))
    if m0 > m1:
        raise SystemExit("[ERROR] mes-desde no puede ser mayor que mes-fin.")

    out_dir = args.out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    nodos = obtener_nodos_renca()
    if not nodos:
        raise SystemExit("[ERROR] No se encontraron nodos de Renca.")

    meses = list(range(m0, m1 + 1))
    labels = [f"{year}-{m:02d}" for m in meses]

    sess = requests.Session()
    filas: List[Dict[str, object]] = []
    for n in nodos:
        row: Dict[str, object] = {"node_id": n["nodeId"], "punto": n["nodeName"]}
        total = 0.0
        for m in meses:
            m3, _, _ = consumo_mes_un_nodo(sess, n["nodeId"], year, m)
            row[f"{year}-{m:02d}"] = round(m3, 4)
            total += float(m3)
        row["total_periodo_m3"] = round(total, 4)
        filas.append(row)

    df = pd.DataFrame(filas)
    csv_path = out_dir / f"consumo_mensual_puntos_renca_{year}_{m0:02d}_{m1:02d}.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    n_series = max(1, len(df))
    x = list(range(len(labels)))
    group_w = 0.8
    w = group_w / n_series

    plt.figure(figsize=(max(10, len(labels) * 1.3), 6))
    for i, (_, r) in enumerate(df.iterrows()):
        vals = [float(r.get(lbl, 0.0) or 0.0) for lbl in labels]
        offs = [(xx - group_w / 2) + (i + 0.5) * w for xx in x]
        plt.bar(offs, vals, width=w, label=str(r["punto"]))

    plt.xticks(x, labels, rotation=0)
    plt.ylabel("Consumo mensual (m3)")
    plt.xlabel("Mes")
    plt.title(f"Renca - Consumo mensual por punto ({year})")
    plt.grid(axis="y", linestyle="--", alpha=0.35)
    plt.legend(fontsize=8, ncol=2)
    plt.tight_layout()

    png_path = out_dir / f"grafico_barras_consumo_mensual_puntos_renca_{year}_{m0:02d}_{m1:02d}.png"
    plt.savefig(png_path, dpi=180)
    plt.close()

    print(f"[OK] CSV: {csv_path}")
    print(f"[OK] PNG: {png_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
