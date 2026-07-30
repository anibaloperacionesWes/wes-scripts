"""
Lee la hoja ``Puente Alto`` del Excel de auditoría/habilitación y extrae el % de ahorro
(rendimiento) por establecimiento.

La celda del porcentaje suele ir como **fracción** (0–1) en la columna «Con WES»; se valida
contra el par Sin WES / Con WES de la misma lectura: (sin - con) / sin.

Salida por defecto: ``reports/proyeccion ahorre puente 2025/pct_auditoria_informe_pa.csv``
(mismo formato que consume ``generar_consolidado_m3_mensual_puente_alto.py``).

Ejemplo::

  python extraer_pct_auditoria_habilitacion_pa.py
  
  python extraer_pct_auditoria_habilitacion_pa.py --excel "R:\\...\\AUDITORIA...xlsx" --dry-run
"""

from __future__ import annotations

import argparse
import csv
import math
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import pandas as pd

from reporte_puente_alto_lxm import mapear_establecimiento_a_nodo, obtener_nodos_puente_alto

# El bloque Liceo Chiloé en la hoja «Puente Alto» puede traer lecturas antiguas (p. ej. 2024);
# la tabla resumen «AUDITORIAS» (sección Puente Alto, marzo 2025) tiene el %% oficial por colegio.
_NODE_USAR_FRACCION_AUDITORIAS: frozenset[str] = frozenset({"000010-08"})

ROOT = Path(__file__).resolve().parent
OUT_DIR_DEFAULT = ROOT / "reports" / "proyeccion ahorre puente 2025"
DEFAULT_EXCEL = Path(
    r"g:\Mi unidad\Colegios\Auditorias\Habilitacion Colegios 2025"
    r"\AUDITORIA CORMUP & PUENTE ALTO 2025.xlsx"
)


def _is_num(x: object) -> bool:
    return isinstance(x, (int, float)) and not (isinstance(x, float) and math.isnan(x))


def _school_header_rows(df: pd.DataFrame) -> List[Tuple[int, float, str]]:
    out: List[Tuple[int, float, str]] = []
    for i in range(len(df)):
        r = df.iloc[i]
        b1, b2 = r.iloc[1], r.iloc[2]
        if _is_num(b1) and pd.notna(b2):
            s = str(b2).strip()
            if len(s) > 6 and not s.startswith("Fecha") and "Lectura" not in s:
                out.append((i, float(b1), s))
    return out


def _extract_section_ratio(sec: pd.DataFrame) -> Optional[float]:
    """
    En el tramo del colegio (DataFrame ya recortado), encuentra la fracción de ahorro en
    columna índice 9, validada con el par Sin/Con (índices 8 y 9) de cada lectura.
    """
    pairs: List[Tuple[int, float, float]] = []
    for j in range(len(sec)):
        row = sec.iloc[j]
        sin_, con_ = row.iloc[8], row.iloc[9]
        if _is_num(sin_) and _is_num(con_) and float(sin_) > 0:
            pairs.append((j, float(sin_), float(con_)))

    fracs: List[Tuple[int, float]] = []
    for j in range(len(sec)):
        row = sec.iloc[j]
        v = row.iloc[9]
        if _is_num(v):
            fv = float(v)
            if 0 < fv < 1:
                fracs.append((j, fv))

    chosen: Optional[float] = None
    tol = 0.02
    for pj, sin_, con_ in pairs:
        exp = (sin_ - con_) / sin_
        best_err: Optional[float] = None
        best_f: Optional[float] = None
        for fj, fv in fracs:
            if fj <= pj:
                continue
            err = abs(fv - exp)
            if best_err is None or err < best_err:
                best_err = err
                best_f = fv
        if best_err is not None and best_err < tol and best_f is not None:
            chosen = best_f

    if chosen is None and fracs:
        chosen = fracs[-1][1]

    return chosen


def _auditorias_fraccion_por_node_id(path: Path, nodos: Sequence[dict]) -> Dict[str, float]:
    """
    Lee la tabla «Puente Alto» en la hoja AUDITORIAS (fracción 0–1 en columna índice 15).
    Devuelve ``node_id`` -> fracción de ahorro (no porcentaje).
    """
    raw = pd.read_excel(path, sheet_name="AUDITORIAS", header=None)
    start: Optional[int] = None
    for i in range(len(raw)):
        v = raw.iloc[i, 1]
        if isinstance(v, str) and v.strip().upper() == "PUENTE ALTO":
            start = i
            break
    if start is None:
        return {}
    out: Dict[str, float] = {}
    r = start + 2
    while r < len(raw):
        id_cell = raw.iloc[r, 0]
        name_cell = raw.iloc[r, 1]
        if pd.isna(id_cell) or name_cell is None or (isinstance(name_cell, float) and math.isnan(name_cell)):
            break
        nombre = str(name_cell).strip()
        if not nombre:
            break
        frac_cell = raw.iloc[r, 15]
        if _is_num(frac_cell):
            fv = float(frac_cell)
            if 0 <= fv <= 1.0:
                nid = mapear_establecimiento_a_nodo(nombre, nodos)
                if nid:
                    out[nid] = fv
        r += 1
    return out


def extraer_porcentajes_desde_excel(path: Path) -> List[Tuple[str, str, float]]:
    df = pd.read_excel(path, sheet_name="Puente Alto", header=None)
    starts = _school_header_rows(df)
    nodos = obtener_nodos_puente_alto()
    aud_frac = _auditorias_fraccion_por_node_id(path, nodos)
    out: List[Tuple[str, str, float]] = []
    for idx, (i, _num, name) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else len(df)
        sec = df.iloc[i:end]
        frac = _extract_section_ratio(sec)
        if frac is None:
            continue
        pct = frac * 100.0
        nid = mapear_establecimiento_a_nodo(name, nodos)
        if not nid:
            raise ValueError(f"No hay nodeId para el nombre Excel: {name!r}")
        out.append((nid, name, pct))
    out.sort(key=lambda x: x[0])
    for i, (nid, name, pct) in enumerate(out):
        if nid in _NODE_USAR_FRACCION_AUDITORIAS and nid in aud_frac:
            nuevo = aud_frac[nid] * 100.0
            if abs(nuevo - pct) > 0.01:
                print(
                    f"[INFO] {nid} {name}: Puente Alto detalle={pct:.4f}% -> "
                    f"AUDITORIAS resumen={nuevo:.4f}%",
                    flush=True,
                )
            out[i] = (nid, name, nuevo)
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="Extrae % auditoría Puente Alto desde Excel habilitación.")
    ap.add_argument("--excel", type=Path, default=DEFAULT_EXCEL, help="Ruta al .xlsx de auditoría")
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=OUT_DIR_DEFAULT,
        help="Carpeta de salida (pct_auditoria_informe_pa.csv)",
    )
    ap.add_argument("--dry-run", action="store_true", help="Solo imprimir tabla, no escribir CSV")
    args = ap.parse_args()

    if not args.excel.is_file():
        print(f"[ERROR] No existe el archivo: {args.excel}")
        return 1

    rows = extraer_porcentajes_desde_excel(args.excel)
    args.out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = args.out_dir / "pct_auditoria_informe_pa.csv"
    detalle = args.out_dir / "pct_auditoria_informe_pa_detalle.csv"

    for nid, name, pct in rows:
        print(f"{nid}  {pct:10.4f}  {name}")

    if args.dry_run:
        return 0

    with csv_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["node_id", "pct_eficiencia_auditoria"])
        for nid, _name, pct in rows:
            w.writerow([nid, round(pct, 4)])

    with detalle.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["node_id", "nombre_excel", "pct_eficiencia_auditoria", "fraccion"])
        for nid, name, pct in rows:
            w.writerow([nid, name, round(pct, 6), round(pct / 100.0, 12)])

    print(f"[OK] {csv_path}")
    print(f"[OK] {detalle}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
