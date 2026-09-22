"""
Informe interno WES — Revisión calidad de señal / comparación ultrasónico vs turbina
Matriz ESVAL (Fundo Zapallar, 000027-01).

Uso:
  python generar_informe_interno_calidad_senal_esval.py
"""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
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
COMPANY_ID = "000027"

# Umbral fabricante (calidad de señal / Q)
UMBRAL_FABRICANTE = 60

# Lecturas iniciales
DN_UP_INICIAL = 76.7
Q_INICIAL = 93

# Tras cambio de silicona
DN_UP_POST = 80
Q_POST = 96

# Comparación por pulso (volumen de prueba)
LITROS_ESVAL_ITRON = 230
LITROS_ULTRASONIDO = 100


def _set_run_font(run, *, size: int = 11, bold: bool = False, color: RGBColor | None = None) -> None:
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def _add_heading(doc: Document, text: str, level: int = 1) -> None:
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = _HEADING
        run.font.name = "Calibri"


def _p(doc: Document, text: str, *, bold: bool = False, size: int = 11) -> None:
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(6)
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = para.add_run(text)
    _set_run_font(run, size=size, bold=bold)


def _bullet(doc: Document, text: str) -> None:
    para = doc.add_paragraph(style="List Bullet")
    para.paragraph_format.space_after = Pt(3)
    para.clear()
    run = para.add_run(text)
    _set_run_font(run, size=11)


def _tabla_calidad(doc: Document) -> None:
    tbl = doc.add_table(rows=4, cols=4)
    tbl.style = "Table Grid"
    headers = ["Parámetro", "Lectura inicial", "Tras silicona", "Recomendación fabricante"]
    rows = [
        ["Calidad DN / UP", f"{DN_UP_INICIAL}", f"{DN_UP_POST}", f"> {UMBRAL_FABRICANTE}"],
        ["Calidad Q (sonido)", f"{Q_INICIAL}", f"{Q_POST}", f"> {UMBRAL_FABRICANTE}"],
        ["Estado vs umbral", "Dentro de rango", "Dentro de rango (mejorado)", "—"],
    ]
    for j, h in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        cell.text = ""
        run = cell.paragraphs[0].add_run(h)
        _set_run_font(run, size=10, bold=True, color=_HEADING)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = tbl.rows[i].cells[j]
            cell.text = ""
            run = cell.paragraphs[0].add_run(val)
            _set_run_font(run, size=10)


def _tabla_comparacion(doc: Document) -> None:
    tbl = doc.add_table(rows=4, cols=3)
    tbl.style = "Table Grid"
    headers = ["Equipo", "Volumen registrado", "Observación"]
    delta = LITROS_ESVAL_ITRON - LITROS_ULTRASONIDO
    pct = (delta / LITROS_ESVAL_ITRON) * 100 if LITROS_ESVAL_ITRON else 0
    rows = [
        [
            "Medidor turbina Itron (referencia ESVAL)",
            f"{LITROS_ESVAL_ITRON} litros",
            "Referencia de facturación / red",
        ],
        [
            "Medidor ultrasónico WES (por pulso)",
            f"{LITROS_ULTRASONIDO} litros",
            "Misma ventana de prueba",
        ],
        [
            "Diferencia",
            f"{delta} litros ({pct:.0f} %)",
            "Desvío relevante; no atribuible a calidad de señal",
        ],
    ]
    for j, h in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        cell.text = ""
        run = cell.paragraphs[0].add_run(h)
        _set_run_font(run, size=10, bold=True, color=_HEADING)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = tbl.rows[i].cells[j]
            cell.text = ""
            run = cell.paragraphs[0].add_run(val)
            _set_run_font(run, size=10)


def generar_informe(out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_docx = out_dir / f"Informe_Interno_Calidad_Senal_ESVAL_{stamp}.docx"

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    try:
        from generar_reporte_word import add_logo_to_header

        add_logo_to_header(doc)
    except Exception as exc:
        print(f"[WARN] Logo en encabezado no aplicado: {exc}")

    title = doc.add_heading("Informe interno — Revisión calidad de señal y comparación de caudal", level=0)
    for run in title.runs:
        run.font.color.rgb = _HEADING
        run.font.name = "Calibri"

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for line in (
        f"Empresa: {EMPRESA} ({COMPANY_ID})",
        f"Punto: {NODO_NOMBRE} ({NODE_ID})",
        "Tipo: Informe técnico interno (terreno)",
        f"Fecha de elaboración: {datetime.now().strftime('%d-%m-%Y %H:%M')}",
        "Clasificación: Uso interno WES",
    ):
        run = meta.add_run(line + "\n")
        _set_run_font(run, size=10, color=_MUTED)

    _add_heading(doc, "1. Objetivo", level=1)
    _p(
        doc,
        "Documentar las acciones realizadas en terreno sobre el medidor ultrasónico asociado a "
        f"{NODO_NOMBRE}, ante un desvío significativo respecto del medidor de turbina Itron de ESVAL, "
        "y dejar registrada la hipótesis operativa para la siguiente intervención.",
    )

    _add_heading(doc, "2. Resumen ejecutivo", level=1)
    _p(
        doc,
        "Se revisó la calidad de señal entre transductores (DN/UP) y el indicador Q de calidad de "
        "sonido: ambos valores iniciales estaban por sobre el umbral recomendado por el fabricante "
        f"(> {UMBRAL_FABRICANTE}). Se contrastó el volumen por pulso del ultrasónico contra el medidor "
        f"de turbina Itron (ESVAL): {LITROS_ULTRASONIDO} L vs {LITROS_ESVAL_ITRON} L. Se verificó "
        "continuidad de cables (OK) y se renovó la silicona de acoplamiento, mejorando DN/UP a "
        f"{DN_UP_POST} y Q a {Q_POST}. Esas dos primeras líneas de trabajo no explican ni corrigen "
        "la diferencia de volumen; se plantea como hipótesis principal que los parámetros de "
        "configuración de cañería (diámetro / geometría) no corresponden a la tubería real instalada.",
    )

    _add_heading(doc, "3. Actividades realizadas", level=1)

    _add_heading(doc, "3.1 Revisión de calidad de señal (menú del equipo)", level=2)
    _p(
        doc,
        "En el menú de diagnóstico del medidor ultrasónico se evaluó la calidad de señal entre "
        "transductores (DN y UP) y, en el mismo menú, la calidad del sonido identificada como Q.",
    )
    _tabla_calidad(doc)
    doc.add_paragraph()
    _bullet(
        doc,
        f"Lectura inicial DN/UP: {DN_UP_INICIAL} (recomendación fabricante: sobre {UMBRAL_FABRICANTE}).",
    )
    _bullet(
        doc,
        f"Lectura inicial Q: {Q_INICIAL} (recomendación fabricante: sobre {UMBRAL_FABRICANTE}).",
    )
    _bullet(
        doc,
        "Conclusión parcial: la calidad de señal y de sonido no aparecen como causa primaria del "
        "desvío de caudal; ambos indicadores estaban dentro de rango aceptable desde el inicio.",
    )

    _add_heading(doc, "3.2 Revisión de diámetros de cañería", level=2)
    _p(
        doc,
        "Se revisaron los diámetros de cañería declarados / configurados en el equipo, como paso "
        "previo a la comparación volumétrica. El detalle numérico de la configuración queda sujeto "
        "a verificación en la siguiente visita (comparar diámetro interior real, material y "
        "separación de sensores contra los valores cargados en el medidor).",
    )

    _add_heading(doc, "3.3 Comparación por pulso: ultrasónico vs turbina Itron (ESVAL)", level=2)
    _p(
        doc,
        "Se realizó una prueba de volumen en la misma ventana, contrastando el medidor ultrasónico "
        "WES (conteo por pulso) contra el medidor de turbina marca Itron utilizado como referencia "
        "ESVAL.",
    )
    _tabla_comparacion(doc)
    doc.add_paragraph()
    _bullet(
        doc,
        f"Resultado: rango muy distante — ESVAL/Itron {LITROS_ESVAL_ITRON} L vs ultrasónico "
        f"{LITROS_ULTRASONIDO} L.",
    )
    _bullet(
        doc,
        "El ultrasónico registra aproximadamente un 43 % del volumen indicado por el medidor de "
        "referencia en la prueba realizada (100/230).",
    )

    _add_heading(doc, "3.4 Prueba 1 — Continuidad de cables", level=2)
    _p(
        doc,
        "Ante la diferencia de volumen, la primera prueba de campo fue asegurar la continuidad "
        "eléctrica de los cables de los transductores. Resultado: continuidad correcta. No se "
        "identificó falla de cableado que explique el desvío.",
    )

    _add_heading(doc, "3.5 Prueba 2 — Cambio / renovación de silicona de acoplamiento", level=2)
    _p(
        doc,
        "Se procedió a cambiar/renovar la silicona de acoplamiento acústico de los transductores. "
        "Tras la intervención, mejoraron los indicadores de calidad:",
    )
    _bullet(doc, f"DN / UP: de {DN_UP_INICIAL} a {DN_UP_POST}.")
    _bullet(doc, f"Q: de {Q_INICIAL} a {Q_POST}.")
    _p(
        doc,
        "La mejora confirma un mejor acoplamiento acústico, pero no resolvió la discrepancia "
        "volumétrica observada frente al medidor Itron. En consecuencia, las dos primeras pruebas "
        "(continuidad y silicona) no generan un cambio operativo en el problema de medición de caudal.",
    )

    _add_heading(doc, "4. Hipótesis operativa", level=1)
    _p(
        doc,
        "Con los resultados de las dos primeras pruebas (cables OK; silicona mejora señal pero no "
        "el volumen), se plantea la siguiente hipótesis:",
    )
    _p(
        doc,
        "El problema puede deberse a que la cañería configurada no es la que corresponde "
        "físicamente — es decir, los valores de configuración del medidor ultrasónico "
        "(diámetro interior, material u otros parámetros geométricos) no son los correctos "
        "respecto de la tubería real. Un error en esos parámetros distorsiona el cálculo de "
        "velocidad/área y, por tanto, el volumen acumulado, aun con buena calidad de señal.",
        bold=False,
    )

    _add_heading(doc, "5. Próximos pasos recomendados", level=1)
    _bullet(
        doc,
        "Verificar in situ el diámetro interior real de la tubería (medición física / ficha de "
        "instalación) y contrastarlo con el valor cargado en el menú de configuración del ultrasónico.",
    )
    _bullet(
        doc,
        "Revisar material de cañería, espesor de pared y distancia entre sensores según manual del "
        "fabricante; corregir configuración si hay inconsistencias.",
    )
    _bullet(
        doc,
        "Repetir la prueba de volumen por pulso (ultrasónico vs Itron/ESVAL) después de cualquier "
        "ajuste de parámetros, documentando lecturas antes/después.",
    )
    _bullet(
        doc,
        "Mantener registro fotográfico del menú de configuración y de la sección de tubería "
        "donde están montados los transductores.",
    )

    _add_heading(doc, "6. Conclusión", level=1)
    _p(
        doc,
        "La calidad de señal DN/UP y el indicador Q estaban y siguen por sobre el umbral del "
        "fabricante; la renovación de silicona los mejoró aún más. La continuidad de cables es "
        "correcta. Persiste un desvío importante de volumen respecto del medidor de turbina Itron "
        f"({LITROS_ULTRASONIDO} L vs {LITROS_ESVAL_ITRON} L). La línea de trabajo prioritaria pasa "
        "a ser la validación y corrección de los parámetros de configuración de cañería.",
    )

    pie = doc.add_paragraph()
    pie.paragraph_format.space_before = Pt(18)
    run = pie.add_run(
        "Documento generado para uso interno del equipo WES. No constituye informe de cliente "
        "ni certificado de calibración."
    )
    _set_run_font(run, size=9, color=_MUTED)

    doc.save(out_docx)
    print(f"[OK] Word: {out_docx}")
    return out_docx


def _intentar_pdf(docx_path: Path) -> Path | None:
    pdf_path = docx_path.with_suffix(".pdf")
    # LibreOffice (Linux / cloud)
    for cmd in ("soffice", "libreoffice"):
        import shutil
        import subprocess

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
                print(f"[OK] PDF: {pdf_path}")
                return pdf_path
        except Exception as exc:
            print(f"[WARN] {cmd} falló: {exc}")
    try:
        from generar_reporte_word import convertir_word_a_pdf

        p = convertir_word_a_pdf(docx_path)
        if p and Path(p).is_file():
            print(f"[OK] PDF: {p}")
            return Path(p)
    except Exception as exc:
        print(f"[WARN] convertir_word_a_pdf: {exc}")
    print("[INFO] PDF no disponible en este entorno; se entrega solo Word.")
    return None


def main() -> int:
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = ROOT / "reports" / "Fundo_Zapallar" / "Informes_Tecnicos" / f"calidad_senal_ultrasonido_{stamp}"
    docx_path = generar_informe(out_dir)
    _intentar_pdf(docx_path)
    print(f"[OK] Carpeta: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
