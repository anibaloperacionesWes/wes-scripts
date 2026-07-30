from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def main() -> int:
    base = Path("reports") / "Renca" / "csv_puntos_feb_abr_2026"
    nodes = sorted([p for p in base.iterdir() if p.is_dir()])
    rows = []

    for d in nodes:
        nid = d.name.split("_")[0]
        punto = d.name.split("_", 1)[1] if "_" in d.name else d.name
        acc = {"node_id": nid, "punto": punto, "2026-02": 0.0, "2026-03": 0.0, "2026-04": 0.0}
        for f in d.glob("*.csv"):
            day = f.stem
            if not (day.startswith("2026-02") or day.startswith("2026-03") or day.startswith("2026-04")):
                continue
            try:
                df = pd.read_csv(f)
            except Exception:
                continue
            if "VALUE" not in df.columns:
                continue
            val = pd.to_numeric(df["VALUE"], errors="coerce").fillna(0.0).sum()
            key = day[:7]
            if key in acc:
                acc[key] += float(val)
        for k in ("2026-02", "2026-03", "2026-04"):
            acc[k] = round(acc[k], 4)
        rows.append(acc)

    out_df = pd.DataFrame(rows)
    out_csv = base / "resumen_mensual_feb_abr_2026_nodos_filtrados.csv"
    out_df.to_csv(out_csv, index=False, encoding="utf-8-sig")

    labels = ["2026-02", "2026-03", "2026-04"]
    x = list(range(len(labels)))
    n = max(len(out_df), 1)
    group_w = 0.8
    bar_w = group_w / n

    blue_palette = ["#0B3C5D", "#1D5E89", "#2E7DB0", "#4A9AD1", "#76B7E5", "#A7D1F2"]

    plt.figure(figsize=(11, 6))
    for i, row in out_df.iterrows():
        vals = [float(row[c]) for c in labels]
        offsets = [(xx - group_w / 2) + (i + 0.5) * bar_w for xx in x]
        color = blue_palette[i % len(blue_palette)]
        bars = plt.bar(offsets, vals, width=bar_w, color=color, label=f"{row['node_id']} {row['punto']}")
        for b, v in zip(bars, vals):
            y_text = max(v * 0.92, v - 20)
            plt.text(
                b.get_x() + b.get_width() / 2,
                y_text,
                f"{v:,.1f} m3".replace(",", "."),
                ha="center",
                va="top",
                fontsize=8,
                color="white",
                rotation=90,
            )

    plt.xticks(x, labels)
    plt.ylabel("Consumo mensual (suma VALUE, escala symlog)")
    plt.xlabel("Mes")
    plt.title("Renca - Totales mensuales por nodo (Feb-Abr 2026)")
    plt.grid(axis="y", linestyle="--", alpha=0.35)
    # Escala no lineal suave para que se vean barras pequeñas y grandes a la vez.
    plt.yscale("symlog", linthresh=40)
    plt.legend(
        fontsize=8,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.16),
        ncol=2,
        frameon=True,
    )
    ymax = float(out_df[labels].to_numpy().max()) if not out_df.empty else 1.0
    plt.ylim(0, ymax * 1.35 if ymax > 0 else 1.0)
    plt.tight_layout(rect=[0, 0.06, 1, 1])

    out_png = base / "grafico_barras_totales_mensuales_renca_feb_abr_2026.png"
    plt.savefig(out_png, dpi=180)
    plt.close()

    print(f"[OK] PNG: {out_png.resolve()}")
    print(f"[OK] CSV: {out_csv.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
