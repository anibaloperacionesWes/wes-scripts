"""
Gráfico de barras agrupadas: consumo mensual **con WES** (medición API, azul) vs **sin WES estimado**
(rojo), por colegio Puente Alto.

Lee el CSV del consolidado generado por ``generar_consolidado_m3_mensual_puente_alto.py``.

Por defecto genera **un PNG por cada mes** (12 archivos). Los colegios con consumo **con WES**
(medición) en **cero** ese mes **no entran** en el gráfico.

Ejemplos::

  python grafico_con_vs_sin_wes_puente_alto.py --year 2025

  python grafico_con_vs_sin_wes_puente_alto.py --year 2025 --mes 3

  python grafico_con_vs_sin_wes_puente_alto.py --year 2025 --solo-agregado

  python grafico_con_vs_sin_wes_puente_alto.py --year 2025 --con-agregado
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import List, Optional, Tuple

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd

# Debajo de esto se considera consumo «cero» (no se incluye la barra en el gráfico)
_UMBRAL_CON_M3 = 1e-6

# Residuo numérico mínimo solo si hiciera falta dibujar (tras filtrar ceros no debería usarse)
_EPS_ALTO_BARRA = 1e-6


def _fmt_m3_etiqueta(v: float) -> str:
    """Un decimal máximo; coma como separador decimal (formato habitual en Chile)."""
    if abs(v) < 5e-4:
        return "0,0"
    s = f"{v:.1f}"
    return s.replace(".", ",")


def _linear_width_asinh(con_m3: List[float], sin_m3: List[float]) -> float:
    """Ancho de la zona casi lineal en escala asinh: mejora la lectura de consumos bajos frente a picos."""
    vals = [float(x) for x in con_m3 + sin_m3 if x is not None]
    if not vals:
        return 80.0
    mx = max(vals)
    p50 = float(np.percentile(vals, 50))
    # Heurística: ni tan estrecha (aplasta lo bajo) ni tan ancha (pierde efecto)
    lw = max(25.0, min(350.0, 0.12 * mx + 0.35 * p50))
    return float(lw)

ROOT = Path(__file__).resolve().parent
OUT_DIR_DEFAULT = ROOT / "reports" / "proyeccion ahorre puente 2025" / "graficos"


def _mes_clave(year: int, month: int) -> str:
    return f"{year}-{month:02d}"


def _columnas_mes(year: int, month: int) -> Tuple[str, str]:
    mk = _mes_clave(year, month)
    return mk, f"{mk}_sin_WES_est_m3"


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


def _cargar_datos_mes(
    csv_path: Path,
    year: int,
    month: int,
) -> Tuple[List[str], List[float], List[float]]:
    df = pd.read_csv(csv_path, dtype=str)
    df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]
    col_con, col_sin = _columnas_mes(year, month)
    if col_con not in df.columns or col_sin not in df.columns:
        raise KeyError(f"El CSV no tiene columnas {col_con} / {col_sin}: ¿rango de meses distinto?")

    d = df[df["node_id"].astype(str).str.upper() != "TOTAL"].copy()
    nombres: List[str] = []
    con_vals: List[float] = []
    sin_vals: List[float] = []
    for _, row in d.iterrows():
        nombre = str(row.get("colegio", row.get("node_id", ""))).strip()
        try:
            v_con = float(str(row[col_con]).replace(",", "."))
            v_sin = float(str(row[col_sin]).replace(",", "."))
        except (TypeError, ValueError):
            continue
        nombres.append(nombre)
        con_vals.append(v_con)
        sin_vals.append(v_sin)
    return _filtrar_filas_sin_consumo_cero(nombres, con_vals, sin_vals)


def _filtrar_filas_sin_consumo_cero(
    nombres: List[str],
    con_vals: List[float],
    sin_vals: List[float],
) -> Tuple[List[str], List[float], List[float]]:
    """Quita colegios cuyo consumo con WES del mes es (casi) cero."""
    nn: List[str] = []
    cc: List[float] = []
    ss: List[float] = []
    for n, c, s in zip(nombres, con_vals, sin_vals):
        if float(c) <= _UMBRAL_CON_M3:
            continue
        nn.append(n)
        cc.append(float(c))
        ss.append(float(s))
    return nn, cc, ss


def _meses_etiquetas() -> Tuple[str, ...]:
    return (
        "ene",
        "feb",
        "mar",
        "abr",
        "may",
        "jun",
        "jul",
        "ago",
        "sep",
        "oct",
        "nov",
        "dic",
    )


def _cargar_serie_agregada_mensual(
    csv_path: Path,
    year: int,
) -> Tuple[List[float], List[float]]:
    """
    Suma de todos los colegios por mes: fila TOTAL del consolidado si existe;
    si no, suma de todas las filas excepto TOTAL.
    """
    df = pd.read_csv(csv_path, dtype=str)
    df.columns = [str(c).replace("\ufeff", "").strip() for c in df.columns]
    total_row: Optional[pd.Series] = None
    mask_t = df["node_id"].astype(str).str.strip().str.upper() == "TOTAL"
    if mask_t.any():
        total_row = df.loc[mask_t].iloc[0]

    con_m: List[float] = []
    sin_m: List[float] = []
    for month in range(1, 13):
        col_con, col_sin = _columnas_mes(year, month)
        if col_con not in df.columns or col_sin not in df.columns:
            raise KeyError(f"Faltan columnas {col_con} / {col_sin}")
        if total_row is not None:
            try:
                vc = float(str(total_row[col_con]).replace(",", "."))
                vs = float(str(total_row[col_sin]).replace(",", "."))
            except (TypeError, ValueError):
                vc, vs = 0.0, 0.0
        else:
            sub = df[~mask_t].copy()
            vc = vs = 0.0
            for _, row in sub.iterrows():
                try:
                    vc += float(str(row[col_con]).replace(",", "."))
                    vs += float(str(row[col_sin]).replace(",", "."))
                except (TypeError, ValueError):
                    continue
        con_m.append(vc)
        sin_m.append(vs)
    return con_m, sin_m


def generar_grafico_agregado_anual(
    con_m3: List[float],
    sin_m3: List[float],
    year: int,
    out_png: Path,
    *,
    figsize: Tuple[float, float] = (13.0, 6.5),
    dpi: int = 120,
) -> None:
    """Un gráfico: eje X = meses del año; barras agrupadas total con WES vs total sin WES estimado."""
    if len(con_m3) != 12 or len(sin_m3) != 12:
        raise ValueError("Se esperan 12 valores por serie.")
    meses = _meses_etiquetas()
    x = np.arange(12)
    w = 0.36
    fig, ax = plt.subplots(figsize=figsize, dpi=dpi)
    alt_con = [max(float(v), _EPS_ALTO_BARRA) for v in con_m3]
    alt_sin = [max(float(v), _EPS_ALTO_BARRA) for v in sin_m3]
    lw = _linear_width_asinh(con_m3, sin_m3)
    bars_con = ax.bar(
        x - w / 2,
        alt_con,
        width=w,
        label="Con WES — suma todos los colegios (m³)",
        color="#2563eb",
    )
    bars_sin = ax.bar(
        x + w / 2,
        alt_sin,
        width=w,
        label="Sin WES estimado — suma todos los colegios (m³)",
        color="#dc2626",
    )
    ax.set_yscale("asinh", linear_width=lw)
    ax.set_ylabel("Consumo agregado (m³) · escala asinh")
    ax.set_title(
        f"Puente Alto — consumo mensual agregado (todos los colegios) — {year}\n"
        "Totales desde fila TOTAL del consolidado (o suma por colegio si no hay TOTAL)"
    )
    ax.set_xticks(x)
    ax.set_xticklabels(list(meses), fontsize=10)
    ax.legend(loc="upper left")

    def _fmt_eje_y(val: float, _: object) -> str:
        return f"{val:.1f}".replace(".", ",")

    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_eje_y))
    lbl_c = [_fmt_m3_etiqueta(float(v)) for v in con_m3]
    lbl_s = [_fmt_m3_etiqueta(float(v)) for v in sin_m3]
    ax.bar_label(bars_con, labels=lbl_c, fontsize=7, padding=2, rotation=70)
    ax.bar_label(bars_sin, labels=lbl_s, fontsize=7, padding=2, rotation=70)
    ax.grid(axis="y", linestyle="--", alpha=0.35)
    ax.set_axisbelow(True)
    fig.tight_layout()
    out_png.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_png, bbox_inches="tight")
    plt.close(fig)


def generar_grafico_mes(
    nombres: List[str],
    con_m3: List[float],
    sin_m3: List[float],
    year: int,
    month: int,
    out_png: Path,
    *,
    figsize: Tuple[float, float] = (14.0, 7.5),
    dpi: int = 120,
) -> None:
    x = range(len(nombres))
    w = 0.38
    fig, ax = plt.subplots(figsize=figsize, dpi=dpi)

    alt_con = [max(float(v), _EPS_ALTO_BARRA) for v in con_m3]
    alt_sin = [max(float(v), _EPS_ALTO_BARRA) for v in sin_m3]

    lw = _linear_width_asinh(con_m3, sin_m3)
    bars_con = ax.bar(
        [i - w / 2 for i in x],
        alt_con,
        width=w,
        label="Con WES (medición)",
        color="#2563eb",
    )
    bars_sin = ax.bar(
        [i + w / 2 for i in x],
        alt_sin,
        width=w,
        label="Sin WES (estimado)",
        color="#dc2626",
    )

    ax.set_yscale("asinh", linear_width=lw)
    ax.set_ylabel("Consumo (m³) · escala asinh (mejor si hay valores muy distintos)")
    meses_es = (
        "enero",
        "febrero",
        "marzo",
        "abril",
        "mayo",
        "junio",
        "julio",
        "agosto",
        "septiembre",
        "octubre",
        "noviembre",
        "diciembre",
    )
    titulo_mes = meses_es[month - 1] if 1 <= month <= 12 else str(month)
    ax.set_title(
        f"Puente Alto — con WES vs sin WES estimado — {titulo_mes} {year}\n"
        f"Solo colegios con medición con WES > 0 · n = {len(nombres)}"
    )
    ax.set_xticks(list(x))
    ax.set_xticklabels(nombres, rotation=35, ha="right", fontsize=9)
    ax.legend(loc="upper right")

    def _fmt_eje_y(val: float, _: object) -> str:
        return f"{val:.1f}".replace(".", ",")

    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_eje_y))

    lbl_con = [_fmt_m3_etiqueta(float(v)) for v in con_m3]
    lbl_sin = [_fmt_m3_etiqueta(float(v)) for v in sin_m3]
    fs = max(6, min(8, int(880 / max(len(nombres), 6))))
    ax.bar_label(bars_con, labels=lbl_con, fontsize=fs, padding=2, rotation=75)
    ax.bar_label(bars_sin, labels=lbl_sin, fontsize=fs, padding=2, rotation=75)

    ax.grid(axis="y", linestyle="--", alpha=0.35)
    ax.set_axisbelow(True)
    fig.tight_layout()
    out_png.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_png, bbox_inches="tight")
    plt.close(fig)


def main() -> int:
    ap = argparse.ArgumentParser(description="Gráfico barras Con WES vs Sin WES (Puente Alto)")
    ap.add_argument("--year", type=int, default=2025)
    ap.add_argument(
        "--mes",
        type=int,
        default=None,
        metavar="N",
        help="Solo el mes N (1-12). Si no se indica, se generan los 12 meses.",
    )
    ap.add_argument(
        "--todos-meses",
        action="store_true",
        help="Igual que el default: genera los 12 PNG (compatibilidad).",
    )
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=OUT_DIR_DEFAULT,
        help="Carpeta para los PNG",
    )
    ap.add_argument(
        "--consolidado-csv",
        type=Path,
        default=None,
        help="CSV consolidado (default: el más reciente en reports/proyeccion ahorre puente 2025)",
    )
    ap.add_argument(
        "--solo-agregado",
        action="store_true",
        help="Solo el gráfico agregado municipal (un PNG, todos los meses en un eje).",
    )
    ap.add_argument(
        "--con-agregado",
        action="store_true",
        help="Además de los gráficos por colegio/mes, genera el PNG agregado anual.",
    )
    args = ap.parse_args()

    base_reports = ROOT / "reports" / "proyeccion ahorre puente 2025"
    csv_path = _elegir_csv_consolidado(base_reports, args.consolidado_csv)
    print(f"[INFO] Consolidado: {csv_path}")

    out_dir = Path(args.out_dir).expanduser().resolve()
    year = args.year

    def _emitir_agregado() -> Path:
        con_a, sin_a = _cargar_serie_agregada_mensual(csv_path, year)
        png_a = out_dir / f"puente_alto_AGREGADO_con_vs_sin_wes_{year}.png"
        generar_grafico_agregado_anual(con_a, sin_a, year, png_a)
        print(f"[OK] {png_a} (agregado 12 meses)", flush=True)
        return png_a

    if args.solo_agregado:
        _emitir_agregado()
        return 0

    def _procesar_un_mes(m: int) -> Path | None:
        try:
            nombres, con_v, sin_v = _cargar_datos_mes(csv_path, year, m)
        except KeyError:
            print(f"[SKIP] Mes {m}: columnas ausentes en CSV.", flush=True)
            return None
        if not nombres:
            print(
                f"[SKIP] Mes {m}: ningún colegio con consumo con WES > 0 "
                f"(tras excluir ceros).",
                flush=True,
            )
            return None
        png = out_dir / f"puente_alto_con_vs_sin_wes_{year}_{m:02d}.png"
        generar_grafico_mes(nombres, con_v, sin_v, year, m, png)
        print(f"[OK] {png} ({len(nombres)} colegios)", flush=True)
        return png

    un_solo = args.mes is not None
    if un_solo:
        m = max(1, min(12, int(args.mes)))
        _procesar_un_mes(m)
        if args.con_agregado:
            _emitir_agregado()
        return 0

    generados: List[Path] = []
    for m in range(1, 13):
        p = _procesar_un_mes(m)
        if p is not None:
            generados.append(p)
    print(f"[INFO] Total archivos (por colegio/mes): {len(generados)}", flush=True)
    if args.con_agregado:
        _emitir_agregado()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
