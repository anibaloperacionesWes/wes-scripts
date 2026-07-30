"""
Genera un PDF con las hojas **Propuesta N°1**, **N°2** y **N°3** desde el Excel de regulación
(``Detalle_consumo_*_regulaciones.xlsx``).

Requiere Microsoft Excel instalado (Windows, automatización COM).

Uso:
  python generar_pdf_propuestas_regulacion.py --excel "calculo de regulaciones/Detalle_consumo_000006_01_regulaciones.xlsx"
  python generar_pdf_propuestas_regulacion.py --todos
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CALC = ROOT / "calculo de regulaciones"

HOJAS_PROPUESTA = ("Propuesta N°1", "Propuesta N°2", "Propuesta N°3")

# xlTypePDF
XL_TYPE_PDF = 0
# xlLandscape
XL_LANDSCAPE = 2


def exportar_propuestas_a_pdf(excel_path: Path, pdf_path: Path | None = None) -> Path:
    """
    Abre ``excel_path``, copia las tres hojas Propuesta a un libro temporal y exporta a PDF.
    """
    import win32com.client

    excel_path = excel_path.resolve()
    if not excel_path.is_file():
        raise FileNotFoundError(excel_path)

    if pdf_path is None:
        pdf_path = excel_path.with_name(excel_path.stem + "_Propuestas.pdf")
    pdf_path = Path(pdf_path).resolve()
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    wb_src = None
    wb_dst = None
    try:
        wb_src = xl.Workbooks.Open(str(excel_path))

        for name in HOJAS_PROPUESTA:
            try:
                wb_src.Worksheets(name)
            except Exception as e:
                raise KeyError(
                    f"No está la hoja «{name}» en {excel_path.name}. ¿Generaste el Excel con el script actual?"
                ) from e

        wb_dst = xl.Workbooks.Add()
        # Copiar cada Propuesta *antes* de la primera hoja para que queden en orden 1→3 y Sheet1 al final
        for name in HOJAS_PROPUESTA:
            wb_src.Worksheets(name).Copy(Before=wb_dst.Sheets(1))
        wb_dst.Sheets(wb_dst.Sheets.Count).Delete()

        # Unificar cada propuesta en una sola página PDF (evita 2 páginas por hoja)
        for i in range(1, wb_dst.Sheets.Count + 1):
            ws = wb_dst.Sheets(i)
            ps = ws.PageSetup
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
            ps.Orientation = XL_LANDSCAPE

        wb_dst.ExportAsFixedFormat(XL_TYPE_PDF, str(pdf_path))
        return pdf_path
    finally:
        try:
            if wb_dst is not None:
                wb_dst.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if wb_src is not None:
                wb_src.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass
        del xl


def main() -> int:
    ap = argparse.ArgumentParser(description="PDF con las 3 hojas Propuesta desde Excel regulación")
    ap.add_argument(
        "--excel",
        type=Path,
        default=None,
        help="Ruta al .xlsx (p. ej. Detalle_consumo_000006_01_regulaciones.xlsx)",
    )
    ap.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Ruta del PDF de salida (por defecto: mismo nombre que el Excel + _Propuestas.pdf)",
    )
    ap.add_argument(
        "--todos",
        action="store_true",
        help=f"Genera PDF para {CALC / 'Detalle_consumo_000006_01_regulaciones.xlsx'} y ..._02_...",
    )
    args = ap.parse_args()

    if args.todos:
        pares = [
            (
                CALC / "Detalle_consumo_000006_01_regulaciones.xlsx",
                CALC / "Propuestas_regulacion_000006_01.pdf",
            ),
            (
                CALC / "Detalle_consumo_000006_02_regulaciones.xlsx",
                CALC / "Propuestas_regulacion_000006_02.pdf",
            ),
        ]
        for xlsx, pdf in pares:
            if not xlsx.is_file():
                print(f"[ERROR] No existe: {xlsx}", file=sys.stderr)
                return 1
            out = exportar_propuestas_a_pdf(xlsx, pdf)
            print(out)
        return 0

    if args.excel is None:
        print("[ERROR] Indique --excel archivo.xlsx o use --todos", file=sys.stderr)
        return 1

    p = exportar_propuestas_a_pdf(args.excel.resolve(), args.output)
    print(p)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
