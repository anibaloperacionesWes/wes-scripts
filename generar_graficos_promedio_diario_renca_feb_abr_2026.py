from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def main() -> int:
    base = Path("reports") / "Renca" / "csv_puntos_feb_abr_2026"
    out_dir = base / "graficos_individuales_promedio_diario"
    out_dir.mkdir(parents=True, exist_ok=True)

    meses = ["2026-02", "2026-03", "2026-04"]
    nodos = sorted(
        [p for p in base.iterdir() if p.is_dir() and p.name.startswith("000017-")]
    )
    if not nodos:
        print(f"[ERROR] No hay carpetas de nodos en {base.resolve()}")
        return 1

    for d in nodos:
        node_id = d.name.split("_")[0]
        punto = d.name.split("_", 1)[1] if "_" in d.name else d.name
        month_sum = {m: 0.0 for m in meses}
        month_days_with_data = {m: 0 for m in meses}

        for f in sorted(d.glob("*.csv")):
            day = f.stem
            mes = day[:7]
            if mes not in month_sum:
                continue
            try:
                df = pd.read_csv(f)
            except Exception:
                continue
            if "VALUE" not in df.columns:
                continue
            total_dia = pd.to_numeric(df["VALUE"], errors="coerce").fillna(0.0).sum()
            if total_dia > 0:
                month_days_with_data[mes] += 1
            month_sum[mes] += float(total_dia)

        prom_diario = []
        for m in meses:
            dias = month_days_with_data[m]
            prom = (month_sum[m] / dias) if dias > 0 else 0.0
            prom_diario.append(round(prom, 3))

        plt.figure(figsize=(8, 5))
        bars = plt.bar(meses, prom_diario, color="#1f77b4", width=0.55)
        for b, v in zip(bars, prom_diario):
            y_text = max(v * 0.92, v - 0.5) if v > 0 else 0.05
            plt.text(
                b.get_x() + b.get_width() / 2,
                y_text,
                f"{v:,.2f} m3/d".replace(",", "."),
                ha="center",
                va="top",
                fontsize=9,
                color="white",
                rotation=90,
            )

        plt.title(f"{node_id} - {punto}\nPromedio diario de consumo por mes (Feb-Abr 2026)")
        plt.xlabel("Mes")
        plt.ylabel("Promedio diario (m3/dia)")
        plt.grid(axis="y", linestyle="--", alpha=0.35)
        ymax = max(prom_diario) if prom_diario else 1.0
        plt.ylim(0, ymax * 1.25 if ymax > 0 else 1.0)
        plt.tight_layout()

        safe_name = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in f"{node_id}_{punto}")
        out_png = out_dir / f"{safe_name}_promedio_diario_feb_abr_2026.png"
        plt.savefig(out_png, dpi=180)
        plt.close()
        print(f"[OK] {out_png.resolve()}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
