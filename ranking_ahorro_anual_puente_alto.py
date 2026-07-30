"""
Ranking anual por **ahorro en m³**: (consumo sin WES estimado anual − consumo con WES medido anual).

Lee el CSV del consolidado ``generar_consolidado_m3_mensual_puente_alto.py`` (columnas
``total_anio_m3``, ``total_anio_sin_WES_est_m3``).

Genera:

- ``ranking_ahorro_anual_pa_<año>.csv``
- ``graficos/ranking_ahorro_anual_pa_<año>.png`` (barras horizontales, mayor arriba)

Ejemplo::

  python ranking_ahorro_anual_puente_alto.py --year 2025
"""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import pandas as pd

ROOT = Path(__file__).resolve().parent
OUT_REPORTS = ROOT / "reports" / "proyeccion ahorre puente 2025"
OUT_GRAFICOS = OUT_REPORTS / "graficos"


def _elegir_csv_consolidado(out_dir: Path, explicit: Path | None) -> Path:
    if explicit and explicit.is_file():
        return explicit.resolve()
    candidatos = sorted(
        out_dir.glob("consolidado_m3_mensual_colegios_puente_alto_*_desde_checkpoint.csv"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if candidatos:
        return candidatos[0]
    parcial = out_dir / "consolidado_m3_mensual_colegios_puente_alto_2025_PARCIAL.csv"
    if parcial.is_file():
        return parcial
    raise FileNotFoundError(
        f"No hay CSV consolidado en {out_dir}. Genere uno con generar_consolidado_m3_mensual_puente_alto.py "
        "o use --consolidado-csv."
    )


def _fmt_m3(v: float) -> str:
    return f"{v:.1f}".replace(".", ",")


def main() -> int:
    ap = argparse.ArgumentParser(description="Ranking ahorro anual m³ — Puente Alto")
    ap.add_argument("--year", type=int, default=2025, help="Solo para nombres de archivo (datos del consolidado cargado)")
    ap.add_argument(
        "--consolidado-csv",
        type=Path,
        default=None,
        help="CSV consolidado (default: el más reciente)",
    )
    ap.add_argument("--no-grafico", action="store_true", help="Solo escribe el CSV")
    args = ap.parse_args()

    year = args.year
    csv_in = _elegir_csv_consolidado(OUT_REPORTS, args.consolidado_csv)
    print(f"[INFO] Consolidado: {csv_in}", flush=True)

    df = pd.read_csv(csv_in, dtype=str)
    df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]
    if "total_anio_m3" not in df.columns or "total_anio_sin_WES_est_m3" not in df.columns:
        raise SystemExit("[ERROR] El CSV no tiene total_anio_m3 / total_anio_sin_WES_est_m3.")

    base = df[df["node_id"].astype(str).str.strip().str.upper() != "TOTAL"].copy()
    base["total_anio_m3"] = pd.to_numeric(base["total_anio_m3"].astype(str).str.replace(",", "."), errors="coerce")
    base["total_anio_sin_WES_est_m3"] = pd.to_numeric(
        base["total_anio_sin_WES_est_m3"].astype(str).str.replace(",", "."),
        errors="coerce",
    )
    base["ahorro_anual_m3"] = base["total_anio_sin_WES_est_m3"] - base["total_anio_m3"]
    base = base.sort_values("ahorro_anual_m3", ascending=False).reset_index(drop=True)
    base.insert(0, "ranking", range(1, len(base) + 1))

    tmp = base.rename(columns={"total_anio_m3": "total_anio_con_WES_m3"})
    orden = [
        "ranking",
        "node_id",
        "colegio",
        "total_anio_con_WES_m3",
        "total_anio_sin_WES_est_m3",
        "ahorro_anual_m3",
    ]
    if "pct_eficiencia_auditoria" in tmp.columns:
        orden.append("pct_eficiencia_auditoria")
    salida = tmp[orden]
    for c in ("total_anio_con_WES_m3", "total_anio_sin_WES_est_m3", "ahorro_anual_m3"):
        if c in salida.columns:
            salida[c] = salida[c].round(4)

    OUT_REPORTS.mkdir(parents=True, exist_ok=True)
    csv_out = OUT_REPORTS / f"ranking_ahorro_anual_pa_{year}.csv"
    salida.to_csv(csv_out, index=False, encoding="utf-8-sig")
    print(f"[OK] {csv_out}", flush=True)

    if args.no_grafico:
        return 0

    # Barras horizontales: mayor ahorro arriba → orden ascendente en el eje Y (último = arriba en barh default no)
    # barh: primera fila abajo. Queremos mayor arriba → ordenar por ahorro ascendente (pequeño abajo, grande arriba)
    plot_df = salida.sort_values("ahorro_anual_m3", ascending=True)
    nombres = plot_df["colegio"].astype(str).tolist()
    vals = plot_df["ahorro_anual_m3"].astype(float).tolist()
    y = range(len(nombres))

    fig, ax = plt.subplots(figsize=(11.0, max(5.0, 0.42 * len(nombres))), dpi=120)
    bars = ax.barh(y, vals, height=0.65, color="#0f766e", edgecolor="#115e59", linewidth=0.6)
    ax.set_yticks(list(y))
    ax.set_yticklabels(nombres, fontsize=9)
    ax.set_xlabel("Ahorro anual estimado (m³) = sin WES − con WES")
    ax.set_title(f"Ranking ahorro anual — Puente Alto (todos los colegios) — {year}")
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: _fmt_m3(float(v))))
    ax.grid(axis="x", linestyle="--", alpha=0.35)
    ax.set_axisbelow(True)
    lbl = [_fmt_m3(float(v)) for v in vals]
    ax.bar_label(bars, labels=lbl, padding=4, fontsize=8)
    fig.tight_layout()
    OUT_GRAFICOS.mkdir(parents=True, exist_ok=True)
    png_out = OUT_GRAFICOS / f"ranking_ahorro_anual_pa_{year}.png"
    fig.savefig(png_out, bbox_inches="tight")
    plt.close(fig)
    print(f"[OK] {png_out}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
