from __future__ import annotations

from datetime import date
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def main() -> int:
    base = Path("reports") / "Renca" / "csv_puntos_feb_abr_2026"
    out_dir = base / "graficos_individuales_dia_semana"
    out_dir.mkdir(parents=True, exist_ok=True)

    desde = date(2026, 2, 1)
    hasta = date(2026, 4, 30)
    dias_labels = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo"]

    nodos = sorted([p for p in base.iterdir() if p.is_dir() and p.name.startswith("000017-")])
    if not nodos:
        print(f"[ERROR] No hay carpetas de nodos en {base.resolve()}")
        return 1

    resumen_rows = []
    for d in nodos:
        node_id = d.name.split("_")[0]
        punto = d.name.split("_", 1)[1] if "_" in d.name else d.name

        sum_by_wd = [0.0] * 7
        count_by_wd = [0] * 7

        for f in sorted(d.glob("*.csv")):
            try:
                dia = date.fromisoformat(f.stem)
            except ValueError:
                continue
            if dia < desde or dia > hasta:
                continue
            try:
                df = pd.read_csv(f)
            except Exception:
                continue
            if "VALUE" not in df.columns:
                continue
            total_dia = float(pd.to_numeric(df["VALUE"], errors="coerce").fillna(0.0).sum())
            wd = dia.weekday()  # 0=Lunes ... 6=Domingo
            sum_by_wd[wd] += total_dia
            count_by_wd[wd] += 1

        prom_by_wd = [
            round(sum_by_wd[i] / count_by_wd[i], 3) if count_by_wd[i] > 0 else 0.0
            for i in range(7)
        ]

        plt.figure(figsize=(9, 5))
        bars = plt.bar(dias_labels, prom_by_wd, color="#1f77b4", width=0.6)
        for b, v in zip(bars, prom_by_wd):
            y = max(v * 0.92, v - 0.4) if v > 0 else 0.06
            plt.text(
                b.get_x() + b.get_width() / 2,
                y,
                f"{v:,.2f} m3/d".replace(",", "."),
                ha="center",
                va="top",
                fontsize=9,
                color="white",
                rotation=90,
            )

        plt.title(f"{node_id} - {punto}\nPromedio diario por dia de semana (Feb-Abr 2026)")
        plt.xlabel("Dia de semana")
        plt.ylabel("Promedio diario (m3/dia)")
        plt.grid(axis="y", linestyle="--", alpha=0.35)
        ymax = max(prom_by_wd) if prom_by_wd else 1.0
        plt.ylim(0, ymax * 1.25 if ymax > 0 else 1.0)
        plt.tight_layout()

        safe_name = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in f"{node_id}_{punto}")
        out_png = out_dir / f"{safe_name}_promedio_por_dia_semana_feb_abr_2026.png"
        plt.savefig(out_png, dpi=180)
        plt.close()
        print(f"[OK] {out_png.resolve()}")

        row = {"node_id": node_id, "punto": punto}
        for i, lbl in enumerate(dias_labels):
            row[lbl] = prom_by_wd[i]
        resumen_rows.append(row)

    out_csv = out_dir / "resumen_promedio_por_dia_semana_feb_abr_2026.csv"
    pd.DataFrame(resumen_rows).to_csv(out_csv, index=False, encoding="utf-8-sig")
    print(f"[OK] CSV resumen: {out_csv.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
