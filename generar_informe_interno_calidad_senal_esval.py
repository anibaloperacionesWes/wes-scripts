"""
Informe interno WES — Revisión calidad de señal / comparación ultrasónico vs turbina
Matriz ESVAL (Fundo Zapallar, 000027-01).

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

UMBRAL_FABRICANTE = 60

DN_UP_INICIAL = 76.7
Q_INICIAL = 93
DN_UP_POST = 80
Q_POST = 96

# Comparación inicial por pulso
LITROS_ESVAL_ITRON_1 = 230
LITROS_ULTRASONIDO_1 = 100

# Tras corrección de configuración (diámetro + mortero)
DIFERENCIA_LITROS_POST_CONFIG = 40

# Configuración cañería
DIAMETRO_CONFIG_ERRONEO_MM = 110
DIAMETRO_ESVAL_MM = 118
ESPESOR_MORTERO_MM = 3
MATERIAL = "Fierro dúctil (EN545)"
MARCADO_FISICO = "EN545 / DN100 / PN16"

FOTO_CANERIA = (
    ROOT
    / "reports"
    / "Fundo_Zapallar"
    / "Informes_Tecnicos"
    / "_evidencias"
    / "caneria_EN545_DN100_PN16.jpg"
)


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


def _fill_header_row(tbl, headers: list[str]) -> None:
    for j, h in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        cell.text = ""
        run = cell.paragraphs[0].add_run(h)
        _set_run_font(run, size=10, bold=True, color=_HEADING)


def _fill_row(tbl, i: int, values: list[str]) -> None:
    for j, val in enumerate(values):
        cell = tbl.rows[i].cells[j]
        cell.text = ""
        run = cell.paragraphs[0].add_run(val)
        _set_run_font(run, size=10)


def _tabla_calidad(doc: Document) -> None:
    tbl = doc.add_table(rows=4, cols=4)
    tbl.style = "Table Grid"
    _fill_header_row(
        tbl,
        ["Parámetro", "Lectura inicial", "Tras limpieza + silicona", "Recomendación fabricante"],
    )
    _fill_row(tbl, 1, ["Calidad DN / UP", f"{DN_UP_INICIAL}", f"{DN_UP_POST}", f"> {UMBRAL_FABRICANTE}"])
    _fill_row(tbl, 2, ["Calidad Q (sonido)", f"{Q_INICIAL}", f"{Q_POST}", f"> {UMBRAL_FABRICANTE}"])
    _fill_row(tbl, 3, ["Estado vs umbral", "Dentro de rango", "Dentro de rango (mejorado)", "—"])


def _tabla_comparacion_inicial(doc: Document) -> None:
    tbl = doc.add_table(rows=4, cols=3)
    tbl.style = "Table Grid"
    delta = LITROS_ESVAL_ITRON_1 - LITROS_ULTRASONIDO_1
    pct = (delta / LITROS_ESVAL_ITRON_1) * 100
    _fill_header_row(tbl, ["Equipo", "Volumen registrado", "Observación"])
    _fill_row(
        tbl,
        1,
        [
            "Medidor turbina Itron (referencia ESVAL)",
            f"{LITROS_ESVAL_ITRON_1} litros",
            "Referencia de red / facturación",
        ],
    )
    _fill_row(
        tbl,
        2,
        [
            "Medidor ultrasónico WES (por pulso)",
            f"{LITROS_ULTRASONIDO_1} litros",
            "Misma ventana de prueba",
        ],
    )
    _fill_row(
        tbl,
        3,
        [
            "Diferencia",
            f"{delta} litros ({pct:.0f} %)",
            "Desvío relevante; no atribuible a calidad de señal",
        ],
    )


def _tabla_config_caneria(doc: Document) -> None:
    tbl = doc.add_table(rows=5, cols=3)
    tbl.style = "Table Grid"
    _fill_header_row(tbl, ["Parámetro", "Configuración previa", "Corrección aplicada"])
    _fill_row(tbl, 1, ["Material", "No validado / incompleto", MATERIAL])
    _fill_row(
        tbl,
        2,
        [
            "Diámetro (mm)",
            f"{DIAMETRO_CONFIG_ERRONEO_MM} mm",
            f"{DIAMETRO_ESVAL_MM} mm (ficha sanitaria ESVAL)",
        ],
    )
    _fill_row(
        tbl,
        3,
        [
            "Menú de cobertura (lining)",
            "Sin mortero / no aplicado",
            f"Mortero — espesor {ESPESOR_MORTERO_MM} mm",
        ],
    )
    _fill_row(
        tbl,
        4,
        [
            "Evidencia física",
            "—",
            f"Marcado en obra: {MARCADO_FISICO}",
        ],
    )


def generar_informe(out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_docx = out_dir / f"Informe_Interno_Calidad_Senal_ESVAL_{stamp}.docx"

    # Copiar evidencia a la carpeta del informe
    foto_local: Path | None = None
    if FOTO_CANERIA.is_file():
        foto_local = out_dir / FOTO_CANERIA.name
        shutil.copy2(FOTO_CANERIA, foto_local)

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

    title = doc.add_heading(
        "Informe interno — Revisión calidad de señal y comparación de caudal",
        level=0,
    )
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
        f"{NODO_NOMBRE}, ante un desvío respecto del medidor de turbina Itron de ESVAL; registrar "
        "las correcciones de configuración de cañería aplicadas y el procedimiento de cálculo del "
        "factor de escala para alinear ambas lecturas.",
    )

    _add_heading(doc, "2. Resumen ejecutivo", level=1)
    _p(
        doc,
        "La calidad de señal DN/UP e indicador Q estaban desde el inicio por sobre el umbral del "
        f"fabricante (> {UMBRAL_FABRICANTE}). La primera comparación volumétrica mostró un desvío "
        f"importante: turbina Itron/ESVAL {LITROS_ESVAL_ITRON_1} L vs ultrasónico "
        f"{LITROS_ULTRASONIDO_1} L. Se verificó continuidad de cables (OK); se limpió la tubería y "
        f"se renovó la silicona, mejorando DN/UP a {DN_UP_POST} y Q a {Q_POST}, sin corregir el "
        "volumen. Al revisar en detalle el material de la cañería (fierro dúctil EN545 DN100) y "
        "contrastarlo con la información de la sanitaria ESVAL, se corrigió el diámetro configurado "
        f"de {DIAMETRO_CONFIG_ERRONEO_MM} mm a {DIAMETRO_ESVAL_MM} mm y se cargó el lining de mortero "
        f"({ESPESOR_MORTERO_MM} mm) en el menú de cobertura. La nueva prueba redujo la diferencia a "
        f"aproximadamente {DIFERENCIA_LITROS_POST_CONFIG} L (aún con desviación). Queda pendiente "
        "aplicar un factor de escala calculado confrontando el flujo del menú 00 del ultrasónico "
        "contra una medición tipo laboratorio (1 minuto con cronómetro sobre el medidor de turbina).",
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
        "Conclusión parcial: calidad de señal y de sonido dentro de rango; no son la causa primaria "
        "del desvío de caudal.",
    )

    _add_heading(doc, "3.2 Comparación inicial por pulso: ultrasónico vs turbina Itron (ESVAL)", level=2)
    _p(
        doc,
        "Se contrastó el volumen por pulso del medidor ultrasónico WES contra el medidor de turbina "
        "marca Itron (referencia ESVAL) en la misma ventana de prueba.",
    )
    _tabla_comparacion_inicial(doc)
    doc.add_paragraph()
    _bullet(
        doc,
        f"Resultado: rango muy distante — ESVAL/Itron {LITROS_ESVAL_ITRON_1} L vs ultrasónico "
        f"{LITROS_ULTRASONIDO_1} L (~43 % del volumen de referencia).",
    )

    _add_heading(doc, "3.3 Prueba 1 — Continuidad de cables", level=2)
    _p(
        doc,
        "Primera prueba de campo: asegurar la continuidad eléctrica de los cables de los "
        "transductores. Resultado: continuidad correcta. No se identificó falla de cableado que "
        "explique el desvío.",
    )

    _add_heading(doc, "3.4 Prueba 2 — Limpieza de tubería y renovación de silicona", level=2)
    _p(
        doc,
        "Antes de renovar la silicona de acoplamiento, se limpió la superficie de la tubería en "
        "la zona de montaje de los transductores, para asegurar un acoplamiento acústico correcto "
        "y una lectura confiable. Luego se cambió/renovó la silicona.",
    )
    _bullet(doc, f"DN / UP: de {DN_UP_INICIAL} a {DN_UP_POST}.")
    _bullet(doc, f"Q: de {Q_INICIAL} a {Q_POST}.")
    _p(
        doc,
        "La mejora confirma mejor acoplamiento acústico, pero no resolvió la discrepancia "
        "volumétrica frente al medidor Itron. Continuidad + limpieza/silicona no generaron cambio "
        "operativo en el problema de medición de caudal.",
    )

    _add_heading(doc, "3.5 Hipótesis y validación de configuración de cañería", level=2)
    _p(
        doc,
        "Ante lo anterior se levantó la hipótesis de que la cañería configurada no correspondía a "
        "la tubería real (valores de configuración incorrectos). Se revisó la cañería con mayor "
        "detalle, identificando el material y el marcado físico en obra.",
    )
    _bullet(doc, f"Marcado visible en la pieza: {MARCADO_FISICO} (fierro dúctil según EN545).")
    _bullet(
        doc,
        f"Validación con información publicada / ficha de la sanitaria ESVAL: diámetro de "
        f"{DIAMETRO_ESVAL_MM} mm (no {DIAMETRO_CONFIG_ERRONEO_MM} mm como estaba configurado).",
    )
    _bullet(
        doc,
        "Las cañerías de fierro dúctil de este tipo traen mortero inyectado (lining interno). En el "
        f"menú de cobertura del sensor debía seleccionarse mortero e indicar el espesor: "
        f"{ESPESOR_MORTERO_MM} mm en este caso.",
    )

    if foto_local and foto_local.is_file():
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cap.add_run()
        run.add_picture(str(foto_local), width=Cm(12.5))
        pie_foto = doc.add_paragraph()
        pie_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = pie_foto.add_run(
            f"Figura 1. Evidencia en terreno — cañería {MARCADO_FISICO} (Matriz ESVAL)."
        )
        _set_run_font(r, size=9, color=_MUTED)

    _p(doc, "Correcciones aplicadas en la configuración del medidor ultrasónico:")
    _tabla_config_caneria(doc)
    doc.add_paragraph()

    _add_heading(doc, "3.6 Nueva prueba volumétrica tras corrección de configuración", level=2)
    _p(
        doc,
        "Luego de corregir diámetro y cobertura (mortero), se repitió la comparación ultrasónico "
        "vs turbina Itron/ESVAL.",
    )
    _bullet(
        doc,
        f"La diferencia se redujo a aproximadamente {DIFERENCIA_LITROS_POST_CONFIG} litros: mejora "
        "clara respecto de los 130 L de desvío inicial, pero aún con desviación residual.",
    )
    _bullet(
        doc,
        "Conclusión parcial: la hipótesis de configuración incorrecta se confirma en parte "
        "(el ajuste de diámetro + mortero acercó las lecturas), pero no basta por sí solo para "
        "igualar ambos medidores.",
    )

    _add_heading(doc, "3.7 Factor de escala (próximo ajuste según manual del sensor)", level=2)
    _p(
        doc,
        "Con apoyo de IA y los manuales del sensor, se definió el procedimiento para calcular un "
        "factor de escala que permita alinear el caudal del ultrasónico con la referencia de "
        "turbina:",
    )
    _bullet(
        doc,
        "Valor A — flujo pasante del medidor ultrasónico leído en el menú 00 (caudal instantáneo / "
        "indicado por el equipo).",
    )
    _bullet(
        doc,
        "Valor B — medición tipo laboratorio: con cronómetro, durante 1 minuto, registrar cuántos "
        "litros pasaron por el medidor de turbina Itron (referencia ESVAL).",
    )
    _bullet(
        doc,
        "Con ambos valores se calcula el factor de escala (relación B/A o según fórmula del manual "
        "del fabricante) y se carga en el equipo para que las lecturas coincidan.",
    )
    _p(
        doc,
        "Este paso queda como acción inmediata de seguimiento: ejecutar la toma de 1 minuto, "
        "documentar A y B, aplicar el factor y repetir la prueba de contraste.",
    )

    _add_heading(doc, "4. Hallazgos", level=1)
    _bullet(
        doc,
        "Calidad de señal (DN/UP y Q) dentro de especificación desde el inicio; mejora adicional "
        "tras limpieza de tubería + silicona.",
    )
    _bullet(doc, "Continuidad de cables correcta.")
    _bullet(
        doc,
        f"Configuración previa errónea: diámetro {DIAMETRO_CONFIG_ERRONEO_MM} mm y sin lining de "
        f"mortero; correcto según ESVAL/obra: {DIAMETRO_ESVAL_MM} mm + mortero {ESPESOR_MORTERO_MM} mm "
        f"en cañería {MATERIAL}.",
    )
    _bullet(
        doc,
        f"Tras corrección: diferencia residual ~{DIFERENCIA_LITROS_POST_CONFIG} L → requiere factor "
        "de escala (menú 00 vs prueba cronometrada 1 min en turbina).",
    )

    _add_heading(doc, "5. Próximos pasos", level=1)
    _bullet(
        doc,
        "Ejecutar el procedimiento de factor de escala: menú 00 (ultrasónico) vs litros en 1 minuto "
        "en medidor de turbina; cargar el factor según manual.",
    )
    _bullet(
        doc,
        "Repetir contraste volumétrico post factor de escala y archivar lecturas antes/después.",
    )
    _bullet(
        doc,
        "Dejar registro fotográfico del menú de configuración (diámetro, cobertura/mortero y "
        "factor de escala) junto a esta evidencia de cañería.",
    )

    _add_heading(doc, "6. Conclusión", level=1)
    _p(
        doc,
        "El desvío inicial no se explica por calidad de señal ni por cableado. La limpieza de "
        "tubería y la silicona mejoraron DN/UP y Q, pero no el volumen. La revisión del material "
        f"(fierro dúctil EN545) y la validación con datos ESVAL permitieron corregir el diámetro "
        f"({DIAMETRO_CONFIG_ERRONEO_MM} → {DIAMETRO_ESVAL_MM} mm) y cargar mortero de "
        f"{ESPESOR_MORTERO_MM} mm en el menú de cobertura, reduciendo la diferencia a ~"
        f"{DIFERENCIA_LITROS_POST_CONFIG} L. El cierre de la desviación residual se aborda con el "
        "cálculo e ingreso del factor de escala según menú 00 y prueba cronometrada de 1 minuto "
        "sobre el medidor de turbina.",
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
    out_dir = (
        ROOT
        / "reports"
        / "Fundo_Zapallar"
        / "Informes_Tecnicos"
        / f"calidad_senal_ultrasonido_{stamp}"
    )
    docx_path = generar_informe(out_dir)
    _intentar_pdf(docx_path)
    print(f"[OK] Carpeta: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
