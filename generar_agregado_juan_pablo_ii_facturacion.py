"""
Reporte agregado Juan Pablo II (Las Condes) — respaldo para cobro de factura.

Genera Word + PDF + CSV/TXT con consumo total WES en el periodo de lectura
(sin exclusiones de nodo 000022-01).

Uso:
  python generar_agregado_juan_pablo_ii_facturacion.py
  python generar_agregado_juan_pablo_ii_facturacion.py --desde 01/05/2026 --hasta 12/05/2026
  python generar_agregado_juan_pablo_ii_facturacion.py --referencia-factura "Boleta N° 123456"
"""

from __future__ import annotations

import argparse
import csv
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from generar_reporte_word import (
    convertir_word_a_pdf,
    format_number_chilean,
    generate_aggregated_report,
    get_node_name,
    parse_date,
)

COMPANY_ID = "000022"
NODE_ID = "000022-01"
OUT_BASE = Path(__file__).resolve().parent / "reports" / "Las_Condes" / "Facturacion" / "Juan_Pablo_II"


def _slug_periodo(desde: str, hasta: str) -> str:
    d0 = parse_date(desde)
    d1 = parse_date(hasta)
    return f"{d0.strftime('%Y%m%d')}_{d1.strftime('%Y%m%d')}"


def main() -> int:
    ap = argparse.ArgumentParser(description="Agregado Juan Pablo II para facturación")
    ap.add_argument("--desde", default="01/04/2026", help="Inicio periodo (DD/MM/YYYY)")
    ap.add_argument("--hasta", default="12/04/2026", help="Fin periodo (DD/MM/YYYY)")
    ap.add_argument(
        "--referencia-factura",
        default="",
        help="Texto opcional (N° boleta, cuenta Aguas Andinas, etc.)",
    )
    args = ap.parse_args()

    node_name = get_node_name(NODE_ID)
    slug = _slug_periodo(args.desde, args.hasta)
    out_dir = OUT_BASE / slug
    out_dir.mkdir(parents=True, exist_ok=True)

    nota = (
        f"Documento de respaldo WES para cobro de servicio de monitoreo. "
        f"Punto: {node_name} ({NODE_ID}), empresa Las Condes. "
        f"Periodo de consumo medido en aplicación: {args.desde} al {args.hasta}."
    )
    if args.referencia_factura.strip():
        nota += f" Referencia facturación: {args.referencia_factura.strip()}."

    print("=" * 70)
    print("  AGREGADO JUAN PABLO II — FACTURACIÓN")
    print(f"  {args.desde} - {args.hasta}")
    print("=" * 70)

    docx_path = generate_aggregated_report(
        company_id=COMPANY_ID,
        node_ids=[NODE_ID],
        start_date=args.desde,
        end_date=args.hasta,
        output_dir="reports",
        apply_exclusions=False,
        generate_ppt=False,
        nota_contexto_periodo=nota,
    )

    # Copiar Word a carpeta de facturación con nombre claro
    dest_docx = out_dir / f"Respaldo_Facturacion_Juan_Pablo_II_{slug}.docx"
    shutil.copy2(docx_path, dest_docx)
    print(f"[OK] Word: {dest_docx}")

    pdf_path = convertir_word_a_pdf(dest_docx)
    if pdf_path and pdf_path.is_file():
        print(f"[OK] PDF:  {pdf_path}")
    else:
        print("[ADVERTENCIA] No se pudo convertir a PDF (revise Word instalado).")

    # Consumo total desde API (respaldo de cobro)
    from generar_reporte_word import (
        acl_node_base_url,
        fetch_json,
        flatten_measures,
        normalize_measures_payload,
        summarize_consumption,
    )

    start_dt = parse_date(args.desde)
    end_dt = parse_date(args.hasta, end_of_day=True)
    raw = fetch_json(
        f"{acl_node_base_url()}/nodes/measures/dates",
        params=[
            ("id", NODE_ID),
            ("start", start_dt.strftime("%d%m%Y")),
            ("end", end_dt.strftime("%d%m%Y")),
        ],
    )
    measures = flatten_measures(normalize_measures_payload(raw, NODE_ID))
    summary = summarize_consumption(measures)
    total_m3 = float(summary.get("total", 0) or 0)
    dias = int(summary.get("dias", 0) or 0)

    agregado_json = docx_path.parent / "datos_agregados.json"
    if agregado_json.is_file():
        shutil.copy2(agregado_json, out_dir / f"datos_wes_{slug}.json")

    resumen_txt = out_dir / f"Resumen_Cobro_Juan_Pablo_II_{slug}.txt"
    lines = [
        "RESUMEN PARA COBRO — MONITOREO WES",
        "=" * 50,
        f"Establecimiento: Las Condes — {node_name}",
        f"Nodo WES: {NODE_ID}",
        f"Periodo: {args.desde} al {args.hasta}",
    ]
    if args.referencia_factura.strip():
        lines.append(f"Referencia factura: {args.referencia_factura.strip()}")
    if total_m3 is not None:
        lines.append(f"Consumo total registrado en app WES: {format_number_chilean(float(total_m3), 1)} m³")
    if dias:
        lines.append(f"Días con registro en periodo: {dias}")
    lines.extend(
        [
            "",
            f"Informe Word: {dest_docx.name}",
            f"Generado: {datetime.now().strftime('%d-%m-%Y %H:%M')}",
            "",
            "Use el Word/PDF adjunto como respaldo del consumo medido en el periodo indicado.",
        ]
    )
    resumen_txt.write_text("\n".join(lines), encoding="utf-8")
    print(f"[OK] Resumen: {resumen_txt}")

    csv_path = out_dir / f"Resumen_Cobro_Juan_Pablo_II_{slug}.csv"
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["campo", "valor"])
        w.writerow(["empresa", "Las Condes"])
        w.writerow(["punto", node_name])
        w.writerow(["nodo_wes", NODE_ID])
        w.writerow(["periodo_desde", args.desde])
        w.writerow(["periodo_hasta", args.hasta])
        w.writerow(["referencia_factura", args.referencia_factura.strip()])
        w.writerow(
            ["consumo_total_m3_wes", f"{float(total_m3):.3f}" if total_m3 is not None else ""]
        )
        w.writerow(["archivo_word", dest_docx.name])
    print(f"[OK] CSV: {csv_path}")
    print(f"\n[CARPETA FACTURACIÓN] {out_dir.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
