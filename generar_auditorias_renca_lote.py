"""
Lote de auditorías Renca (mismo formato Word/PDF que abril 2026).

Los 5 puntos de la auditoría de abril:
  000017-08 Colegio ICCO Renca
  000017-04 Esc. Lo Velásquez
  000017-06 Piscina Municipal
  000017-05 Gimnasio
  000017-07 Cumbre de cóndores pte.

Ventana por defecto (agosto 2026, 7+7 lunes–domingo):
  Con WES:  10–16 ago 2026  (semana previa, con control)
  Sin WES:  17–23 ago 2026  (hoy lunes 17 → domingo 23; el lunes 24 vuelve el control)

Si un punto se queda CON control, excluirlo con ``--excluir 000017-XX``.

Uso:
  python generar_auditorias_renca_lote.py --dry-run
  python generar_auditorias_renca_lote.py --excluir 000017-05
  python generar_auditorias_renca_lote.py --excluir 000017-05 --solo-desde-csv
"""

from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

import generar_informe_auditoria_icco_renca_word as icco
from generar_auditoria_wes_cliente import (
    _etiqueta_cuadro_default,
    _resolver_carpeta,
    ejecutar_auditoria_completa,
    periodo_desde_rango,
    xlsx_nombre_consolidado,
)
from generar_reporte_word import get_node_name

ROOT = Path(__file__).resolve().parent
BASE_AGOSTO = (
    ROOT
    / "reports"
    / "reporte de auditoria"
    / "auditoria_puntos_renca_agosto_2026"
)

# Mismos 5 nodos del agregado de abril (reporte_agregado_auditorias_renca_pareto).
PUNTOS_RENCA_AUDITORIA: tuple[tuple[str, str], ...] = (
    ("000017-08", "Colegio ICCO Renca"),
    ("000017-04", "Esc. Lo Velásquez"),
    ("000017-06", "Piscina Municipal"),
    ("000017-05", "Gimnasio"),
    ("000017-07", "Cumbre de cóndores pte."),
)

CON_DESDE_DEFAULT = date(2026, 8, 10)
CON_HASTA_DEFAULT = date(2026, 8, 16)
SIN_DESDE_DEFAULT = date(2026, 8, 17)
SIN_HASTA_DEFAULT = date(2026, 8, 23)


def _fecha_iso(s: str) -> date:
    return date.fromisoformat(s)


def main() -> int:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    ap = argparse.ArgumentParser(
        description="Lote de auditorías Renca (formato ICCO / abril 2026)."
    )
    ap.add_argument(
        "--excluir",
        action="append",
        default=[],
        help="Nodo que se queda CON control (se puede repetir). Ej. --excluir 000017-05",
    )
    ap.add_argument(
        "--base",
        type=Path,
        default=BASE_AGOSTO,
        help="Carpeta base de salida (por defecto: auditoria_puntos_renca_agosto_2026).",
    )
    ap.add_argument("--con-desde", type=_fecha_iso, default=CON_DESDE_DEFAULT)
    ap.add_argument("--con-hasta", type=_fecha_iso, default=CON_HASTA_DEFAULT)
    ap.add_argument("--sin-desde", type=_fecha_iso, default=SIN_DESDE_DEFAULT)
    ap.add_argument("--sin-hasta", type=_fecha_iso, default=SIN_HASTA_DEFAULT)
    ap.add_argument(
        "--solo-desde-csv",
        action="store_true",
        help="No descargar por API: leer csv_descarga_api/*.csv ya existentes.",
    )
    ap.add_argument(
        "--omitir-validacion-csv-app",
        action="store_true",
    )
    ap.add_argument(
        "--dry-run",
        action="store_true",
        help="Solo imprime el plan (nodos, fechas, carpetas). No llama API ni escribe informes.",
    )
    args = ap.parse_args()

    excluir = {x.strip() for x in args.excluir if x and x.strip()}
    periodo_ref = periodo_desde_rango(
        f"Con control ({args.con_desde:%d-%m} a {args.con_hasta:%d-%m-%Y})",
        args.con_desde,
        args.con_hasta,
    )
    periodo_aud = periodo_desde_rango(
        f"Sin control ({args.sin_desde:%d-%m} a {args.sin_hasta:%d-%m-%Y})",
        args.sin_desde,
        args.sin_hasta,
    )
    xlsx_name = xlsx_nombre_consolidado(periodo_ref, periodo_aud)
    base = Path(args.base).expanduser().resolve()

    seleccion: list[tuple[str, str]] = []
    omitidos: list[tuple[str, str]] = []
    for nid, nombre in PUNTOS_RENCA_AUDITORIA:
        if nid in excluir:
            omitidos.append((nid, nombre))
        else:
            seleccion.append((nid, nombre))

    print("Auditoría Renca — formato abril 2026 (Word ICCO)")
    print(
        f"Con WES: {periodo_ref.dias[0]:%d-%m-%Y} al {periodo_ref.dias[-1]:%d-%m-%Y} "
        f"({len(periodo_ref.dias)} días)"
    )
    print(
        f"Sin WES: {periodo_aud.dias[0]:%d-%m-%Y} al {periodo_aud.dias[-1]:%d-%m-%Y} "
        f"({len(periodo_aud.dias)} días)"
    )
    print(f"Base: {base}")
    print(f"Excel: {xlsx_name}")
    if omitidos:
        print("Excluidos (siguen CON control):")
        for nid, nm in omitidos:
            print(f"  - {nid}  {nm}")
    print("A generar:")
    for nid, nm in seleccion:
        carpeta = _resolver_carpeta(base, nid, nm, None)
        print(f"  - {nid}  {nm}")
        print(f"      {carpeta}")

    if len(periodo_ref.dias) != len(periodo_aud.dias):
        print(
            f"[ERROR] Periodos de distinta longitud "
            f"({len(periodo_ref.dias)} vs {len(periodo_aud.dias)})."
        )
        return 2
    if not seleccion:
        print("[ERROR] No queda ningún punto tras --excluir.")
        return 2
    if args.dry_run:
        print("Dry-run: no se generaron informes.")
        return 0

    ok = 0
    fail = 0
    for i, (nid, nm) in enumerate(seleccion, start=1):
        nombre = (get_node_name(nid) or "").strip() or nm
        cliente_dir = _resolver_carpeta(base, nid, nombre, None)
        icco._ETIQUETA_COLEGIO_CUADRO_RESUMEN = _etiqueta_cuadro_default(nombre)
        print("-" * 72)
        print(f"[{i}/{len(seleccion)}] {nid} — {nombre}")
        try:
            docx, pdf = ejecutar_auditoria_completa(
                node_id=nid,
                nombre_mostrado=nombre,
                cliente_dir=cliente_dir,
                solo_desde_csv=args.solo_desde_csv,
                omitir_validacion_csv_app=args.omitir_validacion_csv_app,
                periodo_ref=periodo_ref,
                periodo_aud=periodo_aud,
            )
            ok += 1
            print(f"[OK] {docx.name}" + (f" + {pdf.name}" if pdf else " (sin PDF)"))
        except Exception as e:
            fail += 1
            print(f"[ERROR] {e}")

    print("-" * 72)
    print(f"Completado. OK: {ok} | ERROR: {fail} | omitidos: {len(omitidos)}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
