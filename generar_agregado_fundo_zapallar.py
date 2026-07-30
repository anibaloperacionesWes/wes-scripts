"""Script rápido para generar el reporte agregado de Fundo Zapallar.

Por defecto:
- Genera SOLO el Word agregado (sin PPT) para acelerar.
- Descarga/arma datos en paralelo (más rápido que secuencial).

Uso:
  python generar_agregado_fundo_zapallar.py --start-date 2026-04-01 --end-date 2026-04-30
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, timedelta
from generar_reporte_word import generate_aggregated_report, get_node_name
from exclusiones_reportes import FUNDO_ZAPALLAR_NODE_IDS

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

NODOS_FUNDO_ZAPALLAR = list(FUNDO_ZAPALLAR_NODE_IDS)

COMPANY_ID = "000027"

# Periodo: 01 de diciembre hasta ayer
START_DATE = "01/12/2025"
END_DATE = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")

def main():
    ap = argparse.ArgumentParser(description="Reporte agregado Fundo Zapallar (rápido).")
    ap.add_argument("--start-date", default=None, help="Fecha inicio (YYYY-MM-DD o dd/mm/aaaa).")
    ap.add_argument("--end-date", default=None, help="Fecha fin (YYYY-MM-DD o dd/mm/aaaa).")
    ap.add_argument("--workers", type=int, default=8, help="Paralelismo interno del agregado (default: 8).")
    ap.add_argument(
        "--con-ppt",
        action="store_true",
        help="Genera PPT (más lento). Por defecto se omite para acelerar.",
    )
    args = ap.parse_args()

    start_date = args.start_date or START_DATE
    end_date = args.end_date or END_DATE

    print("=" * 60)
    print("GENERANDO REPORTE AGREGADO DE FUNDO ZAPALLAR")
    print("=" * 60)
    print(f"Empresa: Fundo Zapallar ({COMPANY_ID})")
    print(f"Periodo: {start_date} - {end_date}")
    print(f"Total de nodos: {len(NODOS_FUNDO_ZAPALLAR)}")
    print(f"Nodos: {', '.join([get_node_name(n) for n in NODOS_FUNDO_ZAPALLAR])}")
    print("=" * 60)
    print()
    print("NOTA: El reporte excluirá ESVAL del cálculo de consumo efectivo")
    print("y agregará narrativa especial sobre ESVAL como fuente de agua.")
    print(f"[INFO] Modo rápido: sin PPT, paralelo workers={max(1, int(args.workers))}")
    print()
    
    try:
        reporte_agregado = generate_aggregated_report(
            COMPANY_ID,
            NODOS_FUNDO_ZAPALLAR,
            start_date,
            end_date,
            apply_exclusions=False,
            generate_ppt=False,
            parallel_node_fetch=True,
            max_parallel_workers=max(1, int(args.workers)),
        )
        print("=" * 60)
        print("[OK] REPORTE AGREGADO GENERADO EXITOSAMENTE")
        print("=" * 60)
        print(f"Ubicación: {reporte_agregado}")
        print()

        if args.con_ppt:
            print("[INFO] Generando presentación PPT (modo lento)...")
            try:
                from pathlib import Path
                from generar_reportes_y_ppt_mall_maipu import generar_ppt_desde_agregado

                aggregated_report_path = (
                    Path(reporte_agregado) if isinstance(reporte_agregado, str) else reporte_agregado
                )
                if aggregated_report_path.is_file():
                    aggregated_report_path = aggregated_report_path.parent

                ppt_path = generar_ppt_desde_agregado(
                    company_id=COMPANY_ID,
                    node_ids=NODOS_FUNDO_ZAPALLAR,
                    start_date=start_date,
                    end_date=end_date,
                    aggregated_report_path=aggregated_report_path,
                    company_name="Fundo Zapallar",
                )
                print("[OK] PPT generada exitosamente")
                print(f"Ubicación PPT: {ppt_path}")
            except Exception as e:
                print(f"[ADVERTENCIA] No se pudo generar PPT: {e}")
                import traceback

                traceback.print_exc()
        
    except Exception as e:
        print("=" * 60)
        print("[ERROR] ERROR AL GENERAR REPORTE AGREGADO")
        print("=" * 60)
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()














