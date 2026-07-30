"""
Descarga el CSV horario del API WES: GET .../nodes/{id}/dates.measures.csv

Por cada dia civil Chile puede hacerse una o dos peticiones (dias UTC distintos que cubren ese dia).

Ejemplos (PowerShell):
  python descargar_csv_medidas_wes.py --node-id 000017-08 --dia 2026-03-24
  python descargar_csv_medidas_wes.py --node-id 000017-08 --desde 2026-03-23 --hasta 2026-03-29
  python descargar_csv_medidas_wes.py --node-id 000017-08 --dia 2026-03-24 -o C:\\Descargas\\mi_dia.csv

Listo para Excel/informes (WES_MEDIDAS_CSV_DIR): un archivo por dia ``AAAA-MM-DD.csv``
  python descargar_csv_medidas_wes.py --node-id 000017-08 --desde 2026-03-23 --hasta 2026-03-29 --para-medidas-dir

Carpeta: ``--out-dir`` o variable de entorno ``WES_MEDIDAS_CSV_DIR`` o ``./csv_medidas_wes``.

Misma base URL que el resto de scripts; opcional: set WES_API_BASE_URL=...

**Duplicados TIME:** el API a veces devuelve dos filas con la misma marca ``TIME``; al interpretar
el CSV hay que **sumar** los ``VALUE`` por instante antes de armar horas Chile o totales, si no la
suma horaria queda ~la mitad del ``totalM3`` del día (ver ``_value_by_time_sum_duplicate_rows`` en
``generar_reporte_word``).
"""
from __future__ import annotations

import argparse
import os
from datetime import date, timedelta
from pathlib import Path

import requests

from generar_reporte_word import (
    _chile_hours_from_dates_measures_csv_text,
    _utc_calendar_dates_for_chile_day,
    acl_node_base_url,
    fetch_json,
    normalize_measures_payload,
    _requests_session,
)


def _iter_chile_days(desde: date, hasta: date):
    d = desde
    while d <= hasta:
        yield d
        d += timedelta(days=1)


def _get_csv_bloques_utc_dia_chile(
    node_id: str,
    dia_chile: date,
    *,
    session: requests.Session,
) -> list[tuple[date, str]]:
    """Devuelve (fecha_utc, texto_csv) por cada GET necesario para ese dia Chile."""
    base = acl_node_base_url()
    url = f"{base}/nodes/{node_id}/dates.measures.csv"
    out: list[tuple[date, str]] = []
    for ud in _utc_calendar_dates_for_chile_day(dia_chile):
        ddmmyyyy = ud.strftime("%d%m%Y")
        r = session.get(
            url,
            params=[("start", ddmmyyyy), ("end", ddmmyyyy)],
            timeout=120,
        )
        r.raise_for_status()
        out.append((ud, r.text))
    return out


def descargar_dia_chile(
    node_id: str,
    dia_chile: date,
    out_dir: Path,
    *,
    session: requests.Session,
) -> list[Path]:
    guardados: list[Path] = []
    for ud, text in _get_csv_bloques_utc_dia_chile(node_id, dia_chile, session=session):
        nombre = f"{node_id}_chile{dia_chile:%Y%m%d}_UTC{ud:%Y%m%d}.csv"
        dest = out_dir / nombre
        dest.write_text(text, encoding="utf-8")
        guardados.append(dest)
    return guardados


def _fusionar_csv_bloques(bloques: list[tuple[date, str]]) -> str:
    """Un solo texto: encabezado una vez, luego todas las filas de datos."""
    merged: list[str] = []
    for i, (_, text) in enumerate(bloques):
        lines = text.strip().split("\n")
        if i == 0:
            merged.extend(lines)
        else:
            merged.extend(lines[1:])
    return "\n".join(merged) + "\n"


def _dir_medidas_default() -> Path:
    env = os.environ.get("WES_MEDIDAS_CSV_DIR", "").strip()
    if env:
        return Path(env).resolve()
    return (Path(__file__).resolve().parent / "csv_medidas_wes").resolve()


def _resumen_dia_csv_fusionado(texto: str, dia_chile: date) -> float:
    """Suma m3/h por hora Chile en el dia (aprox. m3 del dia si cada valor es caudal horario)."""
    acc = _chile_hours_from_dates_measures_csv_text(texto, dia_chile)
    return float(sum(acc.values()))


def _total_m3_json_dia(node_id: str, dia_chile: date) -> float | None:
    """totalM3 del JSON para ese dia civil, si aparece en month."""
    target = dia_chile.strftime("%Y-%m-%d")
    for ud in _utc_calendar_dates_for_chile_day(dia_chile):
        date_str = ud.strftime("%d%m%Y")
        raw = fetch_json(
            f"{acl_node_base_url()}/nodes/measures/dates",
            params=[("id", node_id), ("start", date_str), ("end", date_str)],
        )
        pl = normalize_measures_payload(raw, node_id)
        for nm in pl.get("month", []):
            ds = nm.get("date", "")
            if target in ds or ds.startswith(target[:10]):
                t = nm.get("totalM3")
                if t is not None:
                    return float(t)
    return None


def main() -> int:
    p = argparse.ArgumentParser(description="Descarga dates.measures.csv desde el API WES acl-node.")
    p.add_argument("--node-id", required=True, help="ID del nodo, ej. 000017-08")
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument(
        "--dia",
        type=lambda s: date.fromisoformat(s),
        help="Un dia civil Chile (AAAA-MM-DD)",
    )
    g.add_argument("--desde", type=lambda s: date.fromisoformat(s), help="Inicio rango Chile (AAAA-MM-DD)")
    p.add_argument("--hasta", type=lambda s: date.fromisoformat(s), help="Fin rango Chile (AAAA-MM-DD), con --desde")
    p.add_argument(
        "--out-dir",
        type=Path,
        default=None,
        help="Carpeta de salida (default: ./csv_descargas_wes)",
    )
    p.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Solo valido con --dia: un solo archivo .csv juntando todos los bloques UTC del dia.",
    )
    p.add_argument(
        "--para-medidas-dir",
        action="store_true",
        help="Guarda cada dia Chile como AAAA-MM-DD.csv (fusionado), listo para WES_MEDIDAS_CSV_DIR.",
    )
    p.add_argument(
        "--sin-resumen",
        action="store_true",
        help="No imprimir tabla de suma por dia (solo con --para-medidas-dir).",
    )
    args = p.parse_args()

    if args.desde is not None and args.hasta is None:
        p.error("--desde requiere --hasta")
    if args.hasta is not None and args.desde is None:
        p.error("--hasta requiere --desde")
    if args.output is not None and args.dia is None:
        p.error("-o/--output solo se usa con --dia")
    if args.output is not None and args.para_medidas_dir:
        p.error("Use --para-medidas-dir o -o, no ambos")
    if args.para_medidas_dir and args.out_dir is None:
        pass  # se usa _dir_medidas_default() abajo

    out_dir = args.out_dir
    if out_dir is None:
        out_dir = _dir_medidas_default() if args.para_medidas_dir else (Path(__file__).resolve().parent / "csv_descargas_wes")
    out_dir = out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    sess = _requests_session()

    if args.dia is not None:
        dia = args.dia
        bloques = _get_csv_bloques_utc_dia_chile(args.node_id, dia, session=sess)
        fusion = _fusionar_csv_bloques(bloques)
        if args.para_medidas_dir:
            os.environ["WES_MEDIDAS_CSV_DIR"] = str(out_dir)
            dest = out_dir / f"{dia.isoformat()}.csv"
            dest.write_text(fusion, encoding="utf-8")
            print(dest.resolve())
            print(f"WES_MEDIDAS_CSV_DIR={out_dir}")
            if not args.sin_resumen:
                s = _resumen_dia_csv_fusionado(fusion, dia)
                tj = _total_m3_json_dia(args.node_id, dia)
                tjs = f"{tj:.3f}" if tj is not None else "N/A"
                print(f"  suma CSV horaria: {s:.3f} m3 | totalM3 JSON: {tjs}")
        elif args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(fusion, encoding="utf-8")
            print(args.output.resolve())
        else:
            for path in descargar_dia_chile(args.node_id, dia, out_dir, session=sess):
                print(path.resolve())
        return 0

    desde: date = args.desde
    hasta: date = args.hasta
    if args.para_medidas_dir:
        os.environ["WES_MEDIDAS_CSV_DIR"] = str(out_dir)
        print(f"Carpeta (WES_MEDIDAS_CSV_DIR): {out_dir}")
        print(
            "dia Chile    suma CSV horaria    totalM3 JSON    diff"
        )
        print("-" * 62)
        for d in _iter_chile_days(desde, hasta):
            bloques = _get_csv_bloques_utc_dia_chile(args.node_id, d, session=sess)
            fusion = _fusionar_csv_bloques(bloques)
            dest = out_dir / f"{d.isoformat()}.csv"
            dest.write_text(fusion, encoding="utf-8")
            if not args.sin_resumen:
                s = _resumen_dia_csv_fusionado(fusion, d)
                tj = _total_m3_json_dia(args.node_id, d)
                if tj is not None:
                    diff = s - tj
                    print(f"{d.isoformat()}  {s:10.3f}          {tj:10.3f}      {diff:+8.3f}")
                else:
                    print(f"{d.isoformat()}  {s:10.3f}          (sin total JSON)")
            else:
                print(dest.resolve())
        if not args.sin_resumen:
            print("-" * 62)
            print("Archivos: AAAA-MM-DD.csv | diff ~0 si CSV y JSON coinciden para el total del dia.")
        return 0

    todos: list[Path] = []
    for d in _iter_chile_days(desde, hasta):
        todos.extend(descargar_dia_chile(args.node_id, d, out_dir, session=sess))
    for path in todos:
        print(path.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
