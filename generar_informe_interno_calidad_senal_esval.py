"""
Informe interno ejecutivo — Validación Matriz ESVAL (Fundo Zapallar).
Estilo alineado a «Informe Validacion CUR»: corto, tablas claras, % error.

Uso:
  python generar_informe_interno_calidad_senal_esval.py
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
_HEADING = RGBColor(31, 71, 136)
_MUTED = RGBColor(89, 89, 89)

NODE_ID = "000027-01"
NODO_NOMBRE = "Matriz ESVAL"
EMPRESA = "Fundo Zapallar"

DIAMETRO_FINAL_MM = 127
ESPESOR_MORTERO_MM = 3
M45_MAXIMO = 1.5

VAL_FECHA = "22-09-2026"
HORA_INI, HORA_FIN = "11:50", "15:00"

ITRON_INI, ITRON_FIN = 383956.27, 383978.01
ITRON_DELTA = round(ITRON_FIN - ITRON_INI, 2)  # 21.74

US_NET_CRUDO_INI, US_NET_CRUDO_FIN = 1811906, 1814145
US_INI = round(US_NET_CRUDO_INI * 0.01, 2)  # 18119.06
US_FIN = round(US_NET_CRUDO_FIN * 0.01, 2)  # 18141.45
US_DELTA = round(US_FIN - US_INI, 2)  # 22.39

ERROR_PCT = round((1 - ITRON_DELTA / US_DELTA) * 100, 1)  # 2.9
TOL_US_PCT = 1
TOL_ITRON_PCT = 5

EVIDENCIAS = ROOT / "reports" / "Fundo_Zapallar" / "Informes_Tecnicos" / "_evidencias"
FOTOS = [
    (EVIDENCIAS / "itron_1150.jpg", f"Itron Flostar S — {HORA_INI}"),
    (EVIDENCIAS / "ultrasonido_1150.jpg", f"Ultrasónico WES — {HORA_INI}"),
    (EVIDENCIAS / "itron_1500.jpg", f"Itron Flostar S — {HORA_FIN}"),
    (EVIDENCIAS / "ultrasonido_1500.jpg", f"Ultrasónico WES — {HORA_FIN}"),
]


def _fmt(x: float, dec: int = 2) -> str:
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _font(run, *, size=11, bold=False, color=None) -> None:
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def _h(doc: Document, text: str, level: int = 1) -> None:
    p = doc.add_heading(text, level=level)
    for r in p.runs:
        r.font.color.rgb = _HEADING
        r.font.name = "Calibri"


def _p(doc: Document, text: str) -> None:
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(6)
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = para.add_run(text)
    _font(r, size=11)


def _bullet(doc: Document, text: str) -> None:
    para = doc.add_paragraph(style="List Bullet")
    para.paragraph_format.space_after = Pt(2)
    para.clear()
    r = para.add_run(text)
    _font(r, size=11)


def _merge_title(tbl, text: str, cols: int) -> None:
    cell = tbl.rows[0].cells[0]
    for j in range(1, cols):
        cell.merge(tbl.rows[0].cells[j])
    cell.text = ""
    r = cell.paragraphs[0].add_run(text)
    _font(r, size=11, bold=True, color=_HEADING)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER


def _set_cells(row, values: list[str], *, bold: bool = False, size: int = 9) -> None:
    for j, v in enumerate(values):
        cell = row.cells[j]
        cell.text = ""
        r = cell.paragraphs[0].add_run(v)
        _font(r, size=size, bold=bold, color=_HEADING if bold else None)


def _fig(doc: Document, path: Path, caption: str) -> None:
    if not path.is_file():
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(path), width=Cm(10.5))
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = c.add_run(caption)
    _font(r, size=9, color=_MUTED)


def _tabla_validacion(doc: Document, titulo: str, lectura_ini: float, lectura_fin: float, consumo: float) -> None:
    headers = [
        "ANÁLISIS",
        "HORA INICIAL",
        "LECTURA (m³)",
        "HORA FINAL",
        "LECTURA (m³)",
        "CONSUMO m³",
        "HORAS",
    ]
    tbl = doc.add_table(rows=3, cols=7)
    tbl.style = "Table Grid"
    _merge_title(tbl, titulo, 7)
    _set_cells(tbl.rows[1], headers, bold=True, size=9)
    _set_cells(
        tbl.rows[2],
        [
            NODO_NOMBRE,
            f"{VAL_FECHA} {HORA_INI}",
            _fmt(lectura_ini),
            f"{VAL_FECHA} {HORA_FIN}",
            _fmt(lectura_fin),
            _fmt(consumo),
            "3,2",
        ],
        size=9,
    )
    doc.add_paragraph()


def generar_informe(out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_docx = out_dir / f"Informe_Validacion_Matriz_ESVAL_{stamp}.docx"

    fotos_local: list[tuple[Path, str]] = []
    for src, cap in FOTOS:
        if src.is_file():
            dst = out_dir / src.name
            shutil.copy2(src, dst)
            fotos_local.append((dst, cap))

    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.8)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)

    try:
        from generar_reporte_word import add_logo_to_header

        add_logo_to_header(doc)
    except Exception as exc:
        print(f"[WARN] Logo: {exc}")

    title = doc.add_heading("Informe de validación WES v/s medidor Itron (ESVAL)", level=0)
    for r in title.runs:
        r.font.color.rgb = _HEADING
        r.font.name = "Calibri"

    meta = doc.add_paragraph()
    for line in (
        f"{EMPRESA} · {NODO_NOMBRE} ({NODE_ID})",
        f"Fecha de validación: {VAL_FECHA}  |  Ventana: {HORA_INI} → {HORA_FIN}",
        "Uso interno WES",
    ):
        r = meta.add_run(line + "\n")
        _font(r, size=10, color=_MUTED)

    _p(
        doc,
        "En las siguientes tablas se muestran las lecturas registradas en terreno, "
        "comparando el medidor de turbina Itron (referencia ESVAL) con el medidor ultrasónico WES "
        "en el mismo periodo, para determinar la exactitud del monitoreo.",
    )

    _h(doc, "Qué se revisó en terreno", level=1)
    _bullet(doc, "Calidad de señal del ultrasónico: dentro de lo recomendado por el fabricante.")
    _bullet(doc, "Cables en buen estado. Se limpió la tubería y se renovó la silicona de los sensores.")
    _bullet(
        doc,
        "Se ajustó la configuración de la cañería (diámetro y revestimiento interno de mortero) "
        f"y el factor de escala del equipo (tope {_fmt(M45_MAXIMO, 1)}).",
    )
    _bullet(
        doc,
        f"Configuración al cierre: diámetro {DIAMETRO_FINAL_MM} mm · mortero {ESPESOR_MORTERO_MM} mm.",
    )

    _h(doc, NODO_NOMBRE, level=1)

    _tabla_validacion(
        doc,
        "Validación medidor Itron (ESVAL)",
        ITRON_INI,
        ITRON_FIN,
        ITRON_DELTA,
    )
    _tabla_validacion(
        doc,
        "Validación medidor ultrasónico WES (NET × 0,01)",
        US_INI,
        US_FIN,
        US_DELTA,
    )

    t_err = doc.add_table(rows=3, cols=2)
    t_err.style = "Table Grid"
    _set_cells(t_err.rows[0], ["Total Itron (lectura)", f"{_fmt(ITRON_DELTA)} m³"], size=10)
    _set_cells(t_err.rows[1], ["Total ultrasónico WES", f"{_fmt(US_DELTA)} m³"], size=10)
    _set_cells(t_err.rows[2], ["% Error", f"{_fmt(ERROR_PCT, 1)} %"], bold=True, size=10)
    doc.add_paragraph()

    _p(
        doc,
        f"En base a las lecturas podemos determinar que el rango de error entre el medidor Itron "
        f"y el ultrasónico WES es de un {_fmt(ERROR_PCT, 1)} %. "
        f"Los valores no son idénticos: el fabricante del ultrasónico declara un rango de error "
        f"del {TOL_US_PCT} % y el de Itron del {TOL_ITRON_PCT} %. Un {_fmt(ERROR_PCT, 1)} % se "
        f"considera aceptable dentro de esos márgenes.",
    )

    if fotos_local:
        _h(doc, "Evidencia fotográfica", level=1)
        for path, cap in fotos_local:
            _fig(doc, path, cap)

    pie = doc.add_paragraph()
    pie.paragraph_format.space_before = Pt(12)
    r = pie.add_run(
        f"Informe interno WES · {EMPRESA} · generado {datetime.now().strftime('%d-%m-%Y %H:%M')}."
    )
    _font(r, size=9, color=_MUTED)

    doc.save(out_docx)
    print(f"[OK] Word: {out_docx}")
    print(f"[CIFRAS] Itron Δ={ITRON_DELTA} | US Δ={US_DELTA} | error={ERROR_PCT}%")
    return out_docx


def _intentar_pdf(docx_path: Path) -> Path | None:
    pdf_path = docx_path.with_suffix(".pdf")
    for cmd in ("soffice", "libreoffice"):
        bin_path = shutil.which(cmd)
        if not bin_path:
            continue
        try:
            subprocess.run(
                [
                    bin_path,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(docx_path.parent),
                    str(docx_path),
                ],
                check=True,
                capture_output=True,
                timeout=120,
            )
            if pdf_path.is_file():
                return pdf_path
        except Exception:
            pass
    return None


def main() -> int:
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = (
        ROOT
        / "reports"
        / "Fundo_Zapallar"
        / "Informes_Tecnicos"
        / f"validacion_esval_{stamp}"
    )
    docx_path = generar_informe(out_dir)
    _intentar_pdf(docx_path)
    print(f"[OK] Carpeta: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
