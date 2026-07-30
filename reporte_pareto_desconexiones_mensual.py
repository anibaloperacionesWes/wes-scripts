"""
Genera un gráfico de Pareto mensual de nodos que más veces aparecen como "sin respuesta"
según el histórico `logs/desconexiones_diarias.csv` (generado por `reporte_incidentes_dia_anterior.py`).

Uso:
  python reporte_pareto_desconexiones_mensual.py --year 2026 --month 3

Salida:
  reports/pareto_desconexiones/PARETO_DESCONEXIONES_YYYYMM_YYYYMMDD_HHMM/
    - pareto_desconexiones_YYYYMM.png
    - resumen_pareto_YYYYMM.csv
"""

import argparse
import csv
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


LOGS_CSV = Path("logs") / "desconexiones_diarias.csv"


def _month_range(year: int, month: int) -> tuple[str, str]:
    start = datetime(year, month, 1).date().isoformat()
    if month == 12:
        end = datetime(year + 1, 1, 1).date().isoformat()
    else:
        end = datetime(year, month + 1, 1).date().isoformat()
    return start, end  # end is exclusive


def _load_rows_for_month(year: int, month: int) -> list[dict]:
    if not LOGS_CSV.exists():
        return []

    start_iso, end_iso = _month_range(year, month)
    rows = []
    with open(LOGS_CSV, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            fecha = (row.get("fecha") or "").strip()
            if not fecha:
                continue
            if start_iso <= fecha < end_iso:
                rows.append(row)
    return rows


def _build_pareto(rows: list[dict]) -> tuple[list[tuple[str, int]], dict[str, dict]]:
    """
    Returns:
      - counts_sorted: list of (nodeId, count) desc
      - meta: nodeId -> {nodeName, companyName}
    """
    counts = Counter()
    meta = {}
    for r in rows:
        node_id = (r.get("nodeId") or "").strip()
        if not node_id:
            continue
        counts[node_id] += 1
        meta[node_id] = {
            "nodeName": (r.get("nodeName") or "").strip(),
            "companyName": (r.get("companyName") or "").strip(),
        }
    counts_sorted = sorted(counts.items(), key=lambda x: x[1], reverse=True)
    return counts_sorted, meta


def _save_csv(output_dir: Path, year: int, month: int, counts_sorted: list[tuple[str, int]], meta: dict[str, dict]):
    out_csv = output_dir / f"resumen_pareto_{year}{month:02d}.csv"
    total = sum(c for _, c in counts_sorted) or 1
    cumulative = 0
    with open(out_csv, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["rank", "nodeId", "nodeName", "companyName", "desconexiones", "porcentaje", "porcentaje_acumulado"])
        for i, (node_id, c) in enumerate(counts_sorted, start=1):
            cumulative += c
            m = meta.get(node_id, {})
            pct = c / total * 100.0
            pct_cum = cumulative / total * 100.0
            writer.writerow([i, node_id, m.get("nodeName", ""), m.get("companyName", ""), c, f"{pct:.2f}", f"{pct_cum:.2f}"])
    return out_csv


def _plot_pareto(output_dir: Path, year: int, month: int, counts_sorted: list[tuple[str, int]], meta: dict[str, dict], top_n: int):
    if not counts_sorted:
        return None

    shown = counts_sorted[:max(1, top_n)]
    labels = []
    values = []
    for node_id, c in shown:
        m = meta.get(node_id, {})
        # Etiqueta compacta: nodeId - Empresa (si cabe)
        company = m.get("companyName", "")
        label = f"{node_id}" + (f" | {company}" if company else "")
        labels.append(label)
        values.append(c)

    total = sum(c for _, c in counts_sorted) or 1
    cum = []
    running = 0
    for v in values:
        running += v
        cum.append(running / total * 100.0)

    fig, ax1 = plt.subplots(figsize=(12, 6))

    ax1.bar(range(len(values)), values, color="#1f77b4")
    ax1.set_ylabel("Cantidad de desconexiones (apariciones en 'sin respuesta')")
    ax1.set_xlabel("Nodo (nodeId | empresa)")
    ax1.set_xticks(range(len(labels)))
    ax1.set_xticklabels(labels, rotation=60, ha="right", fontsize=8)

    ax2 = ax1.twinx()
    ax2.plot(range(len(cum)), cum, color="#d62728", marker="o", linewidth=2)
    ax2.set_ylabel("% acumulado (sobre el mes completo)")
    ax2.set_ylim(0, 105)
    ax2.axhline(80, color="#d62728", linestyle="--", linewidth=1, alpha=0.6)

    titulo = f"Pareto mensual de desconexiones - {year}-{month:02d}"
    ax1.set_title(titulo)
    fig.tight_layout()

    out_png = output_dir / f"pareto_desconexiones_{year}{month:02d}.png"
    fig.savefig(out_png, dpi=200)
    plt.close(fig)
    return out_png


def main():
    parser = argparse.ArgumentParser(description="Generar Pareto mensual de nodos desconectados (sin respuesta)")
    now = datetime.now()
    parser.add_argument("--year", type=int, default=now.year, help="Año (default: año actual)")
    parser.add_argument("--month", type=int, default=now.month, help="Mes 1-12 (default: mes actual)")
    parser.add_argument("--top", type=int, default=25, help="Top N nodos a mostrar en el gráfico (default: 25)")
    args = parser.parse_args()

    if args.month < 1 or args.month > 12:
        raise SystemExit("Mes inválido. Use 1-12.")

    rows = _load_rows_for_month(args.year, args.month)
    counts_sorted, meta = _build_pareto(rows)

    base_out = Path("reports") / "pareto_desconexiones"
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = base_out / f"PARETO_DESCONEXIONES_{args.year}{args.month:02d}_{stamp}"
    out_dir.mkdir(parents=True, exist_ok=True)

    if not rows or not counts_sorted:
        # Generar un CSV vacío informativo
        out_csv = _save_csv(out_dir, args.year, args.month, [], {})
        print(f"[INFO] No hay datos para {args.year}-{args.month:02d}.")
        print(f"[OK] CSV generado: {out_csv}")
        return 0

    out_csv = _save_csv(out_dir, args.year, args.month, counts_sorted, meta)
    out_png = _plot_pareto(out_dir, args.year, args.month, counts_sorted, meta, top_n=args.top)

    print(f"[OK] Datos leídos: {len(rows)} filas (mes {args.year}-{args.month:02d})")
    print(f"[OK] CSV resumen: {out_csv}")
    if out_png:
        print(f"[OK] Gráfico Pareto: {out_png}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

