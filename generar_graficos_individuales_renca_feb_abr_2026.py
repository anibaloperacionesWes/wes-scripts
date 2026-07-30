from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def main() -> int:
    base = Path("reports") / "Renca" / "csv_puntos_feb_abr_2026"
    resumen_csv = base / "resumen_mensual_feb_abr_2026_nodos_filtrados.csv"
    out_dir = base / "graficos_individuales"
    out_dir.mkdir(parents=True, exist_ok=True)

    if not resumen_csv.is_file():
        print(f"[ERROR] No existe: {resumen_csv.resolve()}")
        return 1

    df = pd.read_csv(resumen_csv)
    meses = ["2026-02", "2026-03", "2026-04"]

    for _, row in df.iterrows():
        node_id = str(row.get("node_id", "")).strip()
        punto = str(row.get("punto", "")).strip()
        if not node_id:
            continue

        valores = [float(row.get(m, 0.0) or 0.0) for m in meses]
        plt.figure(figsize=(8, 5))
        bars = plt.bar(meses, valores, color="#1f77b4", width=0.55)

        for b, v in zip(bars, valores):
            plt.text(
                b.get_x() + b.get_width() / 2,
                b.get_height(),
                f"{v:,.1f}".replace(",", "."),
                ha="center",
                va="bottom",
                fontsize=10,
                color="#0b2f5b",
            )

        plt.title(f"{node_id} - {punto}\nConsumo mensual Feb-Abr 2026")
        plt.xlabel("Mes")
        plt.ylabel("Consumo (suma VALUE)")
        plt.grid(axis="y", linestyle="--", alpha=0.35)
        ymax = max(valores) if valores else 1.0
        plt.ylim(0, ymax * 1.18 if ymax > 0 else 1.0)
        plt.tight_layout()

        safe_name = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in f"{node_id}_{punto}")
        out_png = out_dir / f"{safe_name}_barras_azules_feb_abr_2026.png"
        plt.savefig(out_png, dpi=180)
        plt.close()
        print(f"[OK] {out_png.resolve()}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
