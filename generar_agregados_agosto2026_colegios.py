"""
Agregados agosto 2026 — lote colegios (fin de mes, mismo periodo que el lote comercial):

  Renca (colegios), La Florida, La Reina, CORMUP (Peñalolén).

Formato: generate_aggregated_report, sin PPT, fetch paralelo.

Renca: solo colegios operativos (Lo Velásquez, Cumbre Pte., Cumbre Ote. / ICCO).
Se excluyen Gimnasio, Piscina Municipal y puntos dados de baja.
La Reina: consumo válido desde el 05/08/2026 (se excluyen 01–04/08).
Sin puntos rojos de alertas ni sección de día de mayor consumo.

Uso:
  python generar_agregados_agosto2026_colegios.py
"""

from __future__ import annotations

import sys
import time
from typing import List, Optional

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
    except Exception:
        pass

from generar_reporte_word import generate_aggregated_report

# Mismo periodo que el lote comercial de agosto 2026
PERIODO = ("01/08/2026", "24/08/2026")

NODOS_RENCA_COLEGIOS = ["000017-04", "000017-07", "000017-08"]
NODOS_LA_FLORIDA = ["000028-01"]
NODOS_LA_REINA = ["000024-01"]
NODOS_CORMUP = [f"000008-{i:02d}" for i in range(1, 15)]


def _run(
    label: str,
    company_id: str,
    node_ids: List[str],
    *,
    apply_exclusions: bool = False,
    company_folder_override: Optional[str] = None,
    workers: Optional[int] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    nota_contexto_periodo: Optional[str] = None,
) -> None:
    if not node_ids:
        print(f"[ERROR] {label}: sin nodos.\n")
        return
    ini = start_date or PERIODO[0]
    fin = end_date or PERIODO[1]
    w = workers if workers is not None else max(4, min(8, len(node_ids)))
    print("=" * 70)
    print(f"[INFO] {label}")
    print(f"       Empresa {company_id} | {len(node_ids)} nodo(s): {', '.join(node_ids)}")
    print(f"       Periodo: {ini} – {fin}")
    t0 = time.perf_counter()
    out = generate_aggregated_report(
        company_id=company_id,
        node_ids=node_ids,
        start_date=ini,
        end_date=fin,
        output_dir="reports",
        apply_exclusions=apply_exclusions,
        generate_ppt=False,
        parallel_node_fetch=True,
        max_parallel_workers=w,
        company_folder_override=company_folder_override,
        nota_contexto_periodo=nota_contexto_periodo,
    )
    print(f"[OK] {label}: {out}")
    print(f"[INFO] Tiempo: {time.perf_counter() - t0:.1f} s\n")


def main() -> None:
    jobs = [
        (
            "Renca (colegios)",
            "000017",
            list(NODOS_RENCA_COLEGIOS),
            dict(apply_exclusions=False, workers=6),
        ),
        (
            "La Florida",
            "000028",
            list(NODOS_LA_FLORIDA),
            dict(apply_exclusions=False, workers=4),
        ),
        (
            "La Reina",
            "000024",
            list(NODOS_LA_REINA),
            dict(
                apply_exclusions=False,
                workers=4,
                start_date="05/08/2026",
                end_date="24/08/2026",
                nota_contexto_periodo=(
                    "Nota: el consumo de Eugenio María De Hostos se considera válido "
                    "desde el 05/08/2026. Los días 01 al 04 de agosto se excluyen del análisis."
                ),
            ),
        ),
        (
            "CORMUP (Peñalolén)",
            "000008",
            list(NODOS_CORMUP),
            dict(apply_exclusions=False, workers=10),
        ),
    ]

    print("GENERACIÓN AGREGADOS AGOSTO 2026 — LOTE COLEGIOS\n")
    print(f"Periodo: {PERIODO[0]} – {PERIODO[1]}\n")
    ok = 0
    errors: list[str] = []
    for label, cid, nids, kwargs in jobs:
        try:
            _run(label, cid, nids, **kwargs)
            ok += 1
        except Exception as e:
            errors.append(f"{label}: {e}")
            print(f"[ERROR] {label}: {e}\n")
            import traceback

            traceback.print_exc()
    print(f"[INFO] Completados: {ok}/{len(jobs)}")
    if errors:
        print("[INFO] Fallidos:")
        for e in errors:
            print(f"  - {e}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
