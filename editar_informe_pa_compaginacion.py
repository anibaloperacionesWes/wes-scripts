"""
Aplica al .docx informe Puente Alto (mismo post-proceso que al generar con plantilla):
- huecos antes de la sección 4; 1.4 con tarifa plana; sensibilidad en metodología; restos 3.3.

Uso::

  python editar_informe_pa_compaginacion.py
  python editar_informe_pa_compaginacion.py "Ruta\\Informe_PA_generado_compaginacion.docx"
  python editar_informe_pa_compaginacion.py --tarifa-clp-m3 1300
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document

from generar_borrador_resumen_ejecutivo_puente_alto import (
    TARIFA_PLANA_CLP_POR_M3,
    aplicar_post_procesado_informe_consolidado_pa,
)

ROOT = Path(__file__).resolve().parent
DEFAULT_DOCX = (
    ROOT
    / "reports"
    / "proyeccion ahorre puente 2025"
    / "Informe_PA_generado_compaginacion.docx"
)


def main() -> int:
    ap = argparse.ArgumentParser(description="Post-proceso Informe_PA_generado_compaginacion.docx")
    ap.add_argument(
        "docx",
        nargs="?",
        type=Path,
        default=None,
        help=f"Ruta .docx (por defecto: {DEFAULT_DOCX})",
    )
    ap.add_argument("--tarifa-clp-m3", type=int, default=TARIFA_PLANA_CLP_POR_M3)
    args = ap.parse_args()

    path = Path(args.docx).expanduser().resolve() if args.docx else DEFAULT_DOCX
    if not path.is_file():
        print(f"[ERROR] No existe el archivo: {path}")
        return 1

    doc = Document(str(path))
    n_gap, ok14, n_tar, n_p33 = aplicar_post_procesado_informe_consolidado_pa(
        doc, tarifa_clp=args.tarifa_clp_m3
    )
    doc.save(str(path))

    print(f"[OK] Guardado: {path}")
    if n_gap:
        print(f"     Eliminados {n_gap} párrafos hueco antes del bloque 4.")
    elif not n_p33 and not n_tar and not ok14:
        print("     Sin cambios de huecos/1.4/tarifa/3.3 (ya al día o sin coincidencias).")
    if ok14:
        print(f"     Punto 1.4 actualizado (tarifa plana {args.tarifa_clp_m3} CLP/m³).")
    else:
        print(
            "     [WARN] No se encontró el párrafo 1.4 (texto «1.4» + económico + chileno). Revise el Word."
        )
    if n_tar:
        print(f"     Eliminados {n_tar} párrafos de sensibilidad / escenarios tarifarios.")
    if n_p33:
        print(f"     Eliminados {n_p33} párrafos (restos apartado 3.3).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
