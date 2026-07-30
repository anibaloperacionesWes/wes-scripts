"""
Auditoría WES por nodo (cliente): descarga CSV, Excel consolidado, gráficos comparativos,
informe Word y PDF — **una carpeta por cliente** bajo el reporte Renca abril 2026.

Ventana de fechas: la definida en ``auditoria_cpa_icco_renca_grafico`` (7+7: Con WES 13–19 abr,
Sin WES 6–12 abr 2026).

Uso:
  python generar_auditoria_wes_cliente.py --node-id 000017-04 --nombre "Esc. Lo Velásquez"
  python generar_auditoria_wes_cliente.py --node-id 000017-08 --nombre "Colegio ICCO Renca" --carpeta "ICCO Renca 000017-08"
  python generar_auditoria_wes_cliente.py --node-id 000017-05
    (usa el nombre que devuelve la API WES si omites --nombre)

Opciones:
  --solo-desde-csv   No llama a la API: usa los YYYY-MM-DD.csv ya presentes en csv_descarga_api/
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import auditoria_cpa_icco_renca_grafico as graf
import generar_informe_auditoria_icco_renca_word as icco
from auditoria_cpa_icco_renca_grafico import PERIODO_AUDITORIA, PERIODO_REFERENCIA
from generar_excel_auditoria_consolidado_dos_periodos import (
    generar_excel_consolidado,
    validar_consolidado_csv_contra_app,
)
from generar_graficos_comparativos_desde_excel_consolidado import (
    _limpiar_pngs_carpeta_graficos,
    generar_pngs,
    leer_matriz_consolidado,
)
from generar_reporte_word import convertir_word_a_pdf, get_node_name

ROOT = Path(__file__).resolve().parent
BASE_REPORTES = (
    ROOT
    / "reports"
    / "reporte de auditoria"
    / "auditoria_puntos_renca_abril_2026"
)

XLSX_NOMBRE = "consumo_consolidado_parseo_filas_abr06-12_abr13-19_2026.xlsx"


def _safe_stem(nombre: str, max_len: int = 56) -> str:
    """Nombre seguro para archivo/carpeta (sin caracteres raros para Windows)."""
    s = "".join(ch for ch in nombre.strip() if ch.isalnum() or ch in (" ", "-", "_", "."))
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace(" ", "_")
    if len(s) > max_len:
        s = s[:max_len].rstrip("_")
    return s or "cliente"


def _etiqueta_cuadro_default(nombre: str) -> str:
    t = nombre.strip().upper()
    if len(t) <= 22:
        return t
    return t[:22] + "..."


def _resolver_carpeta(base: Path, node_id: str, nombre: str, carpeta_explicita: str | None) -> Path:
    if carpeta_explicita:
        sub = carpeta_explicita.strip()
        if not sub:
            raise ValueError("--carpeta no puede estar vacío.")
        return (base / sub).resolve()
    stem = _safe_stem(nombre, 48)
    return (base / f"Auditoria {stem} {node_id}").resolve()


def ejecutar_auditoria_completa(
    *,
    node_id: str,
    nombre_mostrado: str,
    cliente_dir: Path,
    solo_desde_csv: bool,
    omitir_validacion_csv_app: bool = False,
) -> tuple[Path, Path | None]:
    cliente_dir.mkdir(parents=True, exist_ok=True)
    xlsx_out = cliente_dir / XLSX_NOMBRE
    d1 = tuple(PERIODO_REFERENCIA.dias)
    d2 = tuple(PERIODO_AUDITORIA.dias)

    print("1) CSV + Excel consolidado (API o carpeta local)…")
    generar_excel_consolidado(
        xlsx_out,
        node_id,
        d1,
        d2,
        titulo_p1=f"Con WES: {d1[0]:%d-%m-%Y} al {d1[-1]:%d-%m-%Y}",
        titulo_p2=f"Sin WES: {d2[0]:%d-%m-%Y} al {d2[-1]:%d-%m-%Y}",
        solo_desde_csv=solo_desde_csv,
        csv_dir=cliente_dir / "csv_descarga_api",
    )
    print(f"   Excel: {xlsx_out}")
    print(f"   CSV:   {(cliente_dir / 'csv_descarga_api').resolve()}")

    if not omitir_validacion_csv_app:
        print("   Validación: consolidado (suma día) vs totalM3 JSON (referencia app)…")
        validar_consolidado_csv_contra_app(
            node_id,
            cliente_dir / "csv_descarga_api",
            list(d1) + list(d2),
        )

    gdir = cliente_dir / "graficos_comparativos"
    print("2) Gráficos comparativos…")
    _limpiar_pngs_carpeta_graficos(gdir)
    fechas, mats = leer_matriz_consolidado(xlsx_out)
    generar_pngs(fechas, mats, gdir)

    stem_doc = f"Auditoria_{_safe_stem(nombre_mostrado)}_{node_id}"
    docx_path = cliente_dir / f"{stem_doc}.docx"

    graf.NOMBRE_PUNTO = nombre_mostrado
    icco.NOMBRE_PUNTO = nombre_mostrado
    icco.PORTADA_TITULO = "Informe de Auditoría"
    icco.PORTADA_REFERENCIA_BORRADOR = stem_doc
    icco.PORTADA_ESTABLECIMIENTO_LINEA1 = nombre_mostrado
    icco.PORTADA_ESTABLECIMIENTO_LINEA2 = ""

    print("3) Word (figuras desde Excel / graficos_comparativos)…")
    p = icco.generar_informe_word(
        node_id=node_id,
        out_dir=cliente_dir,
        output_docx=docx_path,
        mantener_borrador_manual=False,
        solo_consolidado=False,
        periodo_ref=PERIODO_REFERENCIA,
        periodo_aud=PERIODO_AUDITORIA,
        figuras_desde_xlsx=xlsx_out,
    )

    print("4) PDF…")
    pdf = convertir_word_a_pdf(p)
    return p, pdf


def main() -> int:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    ap = argparse.ArgumentParser(
        description="Auditoría WES completa por nodo; salida en carpeta propia del cliente."
    )
    ap.add_argument("--node-id", required=True, help="ID nodo WES (ej. 000017-04).")
    ap.add_argument(
        "--nombre",
        default=None,
        help="Nombre del establecimiento en portada e informe. "
        "Si se omite, se usa el nombre en API (get_node_name).",
    )
    ap.add_argument(
        "--carpeta",
        default=None,
        help="Nombre de la subcarpeta bajo reporte Renca abril 2026 "
        '(ej. "Mi cliente 000017-04"). Por defecto: "Auditoria <nombre> <node-id>".',
    )
    ap.add_argument(
        "--base",
        type=Path,
        default=BASE_REPORTES,
        help="Carpeta base donde se crea la carpeta del cliente (por defecto: reporte Renca abril 2026).",
    )
    ap.add_argument(
        "--etiqueta-cuadro",
        default=None,
        help="Texto corto del cuadro resumen (MAYÚSCULAS recomendado). Por defecto: nombre abreviado.",
    )
    ap.add_argument(
        "--solo-desde-csv",
        action="store_true",
        help="No descargar por API: leer csv_descarga_api/*.csv ya existentes.",
    )
    ap.add_argument(
        "--omitir-validacion-csv-app",
        action="store_true",
        help="No comparar suma diaria del consolidado con totalM3 del JSON (salta chequeo anti-error).",
    )
    args = ap.parse_args()

    node_id = args.node_id.strip()
    nombre = (args.nombre or "").strip() or (get_node_name(node_id) or "").strip() or node_id
    etiqueta = (args.etiqueta_cuadro or "").strip() or _etiqueta_cuadro_default(nombre)

    base = Path(args.base).expanduser().resolve()
    cliente_dir = _resolver_carpeta(base, node_id, nombre, args.carpeta)

    icco._ETIQUETA_COLEGIO_CUADRO_RESUMEN = etiqueta

    print(f"Cliente: {nombre} | Nodo: {node_id}")
    print(f"Carpeta: {cliente_dir}")

    try:
        docx, pdf = ejecutar_auditoria_completa(
            node_id=node_id,
            nombre_mostrado=nombre,
            cliente_dir=cliente_dir,
            solo_desde_csv=args.solo_desde_csv,
            omitir_validacion_csv_app=args.omitir_validacion_csv_app,
        )
    except Exception as e:
        print(f"[ERROR] {e}")
        return 1

    print(f"Word: {docx.resolve()}")
    if pdf:
        print(f"PDF:  {pdf.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
