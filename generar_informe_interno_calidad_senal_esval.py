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

LITROS_ESVAL_ITRON_1 = 230
LITROS_ULTRASONIDO_1 = 100

DIFERENCIA_LITROS_POST_CONFIG = 40
DIFERENCIA_LITROS_POST_M45 = 30

DIAMETRO_CONFIG_ERRONEO_MM = 110
DIAMETRO_ESVAL_MM = 118
DIAMETRO_FINAL_MM = 127  # +8 % sobre 118
INCREMENTO_DIAMETRO_PCT = 8
ESPESOR_MORTERO_MM = 3
MATERIAL = "Fierro dúctil (EN545)"
MARCADO_FISICO = "EN545 / DN100 / PN16"

# Factor de escala M45
M45_FACTOR_ACTUAL = 0.5934
M45_CAUDAL_TURBINA_LPM = 460.0
M45_CAUDAL_ULTRASONIDO_LPM = 87.05
M45_CALCULADO = 3.1362
M45_MAXIMO_EQUIPO = 1.5

EVIDENCIAS = ROOT / "reports" / "Fundo_Zapallar" / "Informes_Tecnicos" / "_evidencias"
FOTO_CANERIA = EVIDENCIAS / "caneria_EN545_DN100_PN16.jpg"
FOTO_FORMULA_M45 = EVIDENCIAS / "formula_factor_escala_M45.jpg"


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


def _add_figure(doc: Document, path: Path, caption: str, width_cm: float = 12.5) -> None:
    if not path.is_file():
        return
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.add_run().add_picture(str(path), width=Cm(width_cm))
    pie = doc.add_paragraph()
    pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pie.add_run(caption)
    _set_run_font(r, size=9, color=_MUTED)


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
    tbl = doc.add_table(rows=6, cols=3)
    tbl.style = "Table Grid"
    _fill_header_row(tbl, ["Parámetro", "Configuración previa / intermedia", "Valor final"])
    _fill_row(tbl, 1, ["Material", "No validado / incompleto", MATERIAL])
    _fill_row(
        tbl,
        2,
        [
            "Diámetro (mm)",
            f"{DIAMETRO_CONFIG_ERRONEO_MM} → {DIAMETRO_ESVAL_MM} mm (ficha ESVAL)",
            f"{DIAMETRO_FINAL_MM} mm (+{INCREMENTO_DIAMETRO_PCT} % sobre {DIAMETRO_ESVAL_MM})",
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
            "Factor de escala M45",
            f"Actual {M45_FACTOR_ACTUAL} → calculado {M45_CALCULADO}",
            f"Seteado al máximo permitido: {M45_MAXIMO_EQUIPO}",
        ],
    )
    _fill_row(tbl, 5, ["Evidencia física", "—", f"Marcado en obra: {MARCADO_FISICO}"])


def _tabla_evolucion_diferencias(doc: Document) -> None:
    tbl = doc.add_table(rows=5, cols=3)
    tbl.style = "Table Grid"
    _fill_header_row(tbl, ["Etapa", "Diferencia aprox. (L)", "Observación"])
    _fill_row(
        tbl,
        1,
        [
            "Comparación inicial",
            f"{LITROS_ESVAL_ITRON_1 - LITROS_ULTRASONIDO_1} L",
            f"Itron {LITROS_ESVAL_ITRON_1} L vs ultrasónico {LITROS_ULTRASONIDO_1} L",
        ],
    )
    _fill_row(
        tbl,
        2,
        [
            "Tras diámetro 118 mm + mortero 3 mm",
            f"{DIFERENCIA_LITROS_POST_CONFIG} L",
            "Mejora, aún con desviación",
        ],
    )
    _fill_row(
        tbl,
        3,
        [
            f"Tras M45 = {M45_MAXIMO_EQUIPO} (máximo)",
            f"{DIFERENCIA_LITROS_POST_M45} L",
            "Mejora parcial; no cierra el error",
        ],
    )
    _fill_row(
        tbl,
        4,
        [
            f"Tras diámetro {DIAMETRO_FINAL_MM} mm (+{INCREMENTO_DIAMETRO_PCT} %)",
            "0 L / sin error",
            "Ambos medidores en el mismo ciclo de lectura",
        ],
    )


def generar_informe(out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_docx = out_dir / f"Informe_Interno_Calidad_Senal_ESVAL_{stamp}.docx"

    foto_caneria: Path | None = None
    foto_formula: Path | None = None
    if FOTO_CANERIA.is_file():
        foto_caneria = out_dir / FOTO_CANERIA.name
        shutil.copy2(FOTO_CANERIA, foto_caneria)
    if FOTO_FORMULA_M45.is_file():
        foto_formula = out_dir / FOTO_FORMULA_M45.name
        shutil.copy2(FOTO_FORMULA_M45, foto_formula)

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
        "Estado: Cerrado — medidores alineados",
    ):
        run = meta.add_run(line + "\n")
        _set_run_font(run, size=10, color=_MUTED)

    _add_heading(doc, "1. Objetivo", level=1)
    _p(
        doc,
        "Documentar las acciones realizadas en terreno sobre el medidor ultrasónico asociado a "
        f"{NODO_NOMBRE}, ante un desvío respecto del medidor de turbina Itron de ESVAL; registrar "
        "las correcciones de configuración (diámetro, mortero y factor de escala M45) hasta lograr "
        "que ambos medidores lean en el mismo ciclo, sin error entre ellos.",
    )

    _add_heading(doc, "2. Resumen ejecutivo", level=1)
    _p(
        doc,
        "La calidad de señal DN/UP e indicador Q estaban desde el inicio por sobre el umbral del "
        f"fabricante (> {UMBRAL_FABRICANTE}). La primera comparación volumétrica mostró un desvío "
        f"importante: turbina Itron/ESVAL {LITROS_ESVAL_ITRON_1} L vs ultrasónico "
        f"{LITROS_ULTRASONIDO_1} L. Se verificó continuidad de cables (OK); se limpió la tubería y "
        f"se renovó la silicona (DN/UP {DN_UP_POST}, Q {Q_POST}), sin corregir el volumen.",
    )
    _p(
        doc,
        f"Se validó material fierro dúctil EN545 DN100 y se corrigió el diámetro de "
        f"{DIAMETRO_CONFIG_ERRONEO_MM} a {DIAMETRO_ESVAL_MM} mm (ficha ESVAL), más lining de mortero "
        f"{ESPESOR_MORTERO_MM} mm: la diferencia bajó a ~{DIFERENCIA_LITROS_POST_CONFIG} L. Con la "
        f"prueba de 1 minuto se calculó M45 = {M45_CALCULADO}, pero el equipo admite máximo "
        f"{M45_MAXIMO_EQUIPO}; al setear ese tope la diferencia pasó a ~{DIFERENCIA_LITROS_POST_M45} L "
        f"(aún insuficiente). Siguiendo el manual, se aumentó el diámetro un "
        f"{INCREMENTO_DIAMETRO_PCT} % ({DIAMETRO_ESVAL_MM} → {DIAMETRO_FINAL_MM} mm). Tras ese ajuste, "
        "ambos medidores quedaron en el mismo ciclo de lectura, sin error entre ellos.",
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
        "volumétrica frente al medidor Itron.",
    )

    _add_heading(doc, "3.5 Validación de configuración de cañería (diámetro + mortero)", level=2)
    _p(
        doc,
        "Se levantó la hipótesis de que la cañería configurada no correspondía a la tubería real. "
        "Se revisó el material y el marcado físico en obra, y se contrastó con la información de "
        "la sanitaria ESVAL.",
    )
    _bullet(doc, f"Marcado visible: {MARCADO_FISICO} (fierro dúctil según EN545).")
    _bullet(
        doc,
        f"Diámetro según ficha ESVAL: {DIAMETRO_ESVAL_MM} mm (estaba configurado en "
        f"{DIAMETRO_CONFIG_ERRONEO_MM} mm).",
    )
    _bullet(
        doc,
        f"Lining: mortero inyectado — en el menú de cobertura se seleccionó mortero con espesor "
        f"{ESPESOR_MORTERO_MM} mm.",
    )
    if foto_caneria:
        _add_figure(
            doc,
            foto_caneria,
            f"Figura 1. Evidencia en terreno — cañería {MARCADO_FISICO} (Matriz ESVAL).",
        )
    _p(
        doc,
        f"Tras aplicar {DIAMETRO_ESVAL_MM} mm + mortero {ESPESOR_MORTERO_MM} mm, la nueva prueba "
        f"volumétrica dejó una diferencia residual de aproximadamente "
        f"{DIFERENCIA_LITROS_POST_CONFIG} litros (aún con desviación).",
    )

    _add_heading(doc, "3.6 Cálculo y aplicación del factor de escala M45", level=2)
    _p(
        doc,
        "Con apoyo de IA y los manuales del sensor se ejecutó la prueba tipo laboratorio de "
        "1 minuto: caudal del medidor de turbina versus caudal del ultrasónico (menú 00), para "
        "calcular el nuevo factor de escala M45 a partir del factor actual.",
    )
    _bullet(doc, f"Factor actual (M45): {M45_FACTOR_ACTUAL}.")
    _bullet(
        doc,
        f"Caudal turbina (prueba 1 min): {M45_CAUDAL_TURBINA_LPM:g} L/min "
        f"(≈ {M45_CAUDAL_TURBINA_LPM / 1000:.2f} m³/min).",
    )
    _bullet(doc, f"Caudal ultrasónico (menú 00): {M45_CAUDAL_ULTRASONIDO_LPM} L/min.")
    _bullet(
        doc,
        f"Fórmula aplicada: Nuevo M45 = Factor actual × (caudal turbina / caudal ultrasónico) "
        f"= {M45_FACTOR_ACTUAL} × ({M45_CAUDAL_TURBINA_LPM:g} / {M45_CAUDAL_ULTRASONIDO_LPM}) "
        f"= {M45_CALCULADO}.",
    )
    if foto_formula:
        _add_figure(
            doc,
            foto_formula,
            "Figura 2. Fórmula de cálculo del nuevo factor de escala M45.",
            width_cm=14.0,
        )
    _p(
        doc,
        f"El valor calculado ({M45_CALCULADO}) supera el máximo admitido por el equipo "
        f"({M45_MAXIMO_EQUIPO}). Se seteó M45 = {M45_MAXIMO_EQUIPO} (tope). Al repetir la "
        f"comparación volumétrica, la diferencia bajó de ~{DIFERENCIA_LITROS_POST_CONFIG} L a "
        f"~{DIFERENCIA_LITROS_POST_M45} L: mejora parcial, pero aún no «flotaba» / no se alineaba "
        "porque el rango de error seguía siendo mayor al aceptable.",
    )

    _add_heading(doc, "3.7 Ajuste final de diámetro (+8 %) según recomendación de manual", level=2)
    _p(
        doc,
        "Ante el tope del factor de escala, la recomendación del manual del sensor es aumentar el "
        "diámetro configurado. Se aplicó un incremento del "
        f"{INCREMENTO_DIAMETRO_PCT} % sobre {DIAMETRO_ESVAL_MM} mm, pasando a "
        f"{DIAMETRO_FINAL_MM} mm.",
    )
    _bullet(doc, f"Diámetro final configurado: {DIAMETRO_FINAL_MM} mm.")
    _bullet(doc, f"M45 permanece en el máximo seteado: {M45_MAXIMO_EQUIPO}.")
    _bullet(doc, f"Mortero de cobertura: {ESPESOR_MORTERO_MM} mm.")
    _p(
        doc,
        "Tras este ajuste se repitió la comparación: ambos medidores (ultrasónico y turbina "
        "Itron/ESVAL) quedaron en el mismo ciclo de lectura, sin error entre ellos.",
    )

    _add_heading(doc, "3.8 Resumen de configuración final", level=2)
    _tabla_config_caneria(doc)
    doc.add_paragraph()
    _p(doc, "Evolución de la diferencia volumétrica:")
    _tabla_evolucion_diferencias(doc)
    doc.add_paragraph()

    _add_heading(doc, "4. Hallazgos", level=1)
    _bullet(
        doc,
        "Calidad de señal (DN/UP y Q) dentro de especificación; mejora adicional tras limpieza + "
        "silicona.",
    )
    _bullet(doc, "Continuidad de cables correcta.")
    _bullet(
        doc,
        f"Configuración inicial errónea: diámetro {DIAMETRO_CONFIG_ERRONEO_MM} mm y sin mortero; "
        f"corrección intermedia {DIAMETRO_ESVAL_MM} mm + mortero {ESPESOR_MORTERO_MM} mm.",
    )
    _bullet(
        doc,
        f"M45 teórico {M45_CALCULADO} limitado por el máximo del equipo ({M45_MAXIMO_EQUIPO}); "
        f"con el tope la diferencia quedó en ~{DIFERENCIA_LITROS_POST_M45} L.",
    )
    _bullet(
        doc,
        f"Cierre: diámetro {DIAMETRO_FINAL_MM} mm (+{INCREMENTO_DIAMETRO_PCT} %) alineó ambos "
        "medidores en el mismo ciclo, sin error.",
    )

    _add_heading(doc, "5. Estado y seguimiento", level=1)
    _bullet(
        doc,
        "Intervención cerrada en terreno: ultrasónico y turbina Itron/ESVAL coinciden en lectura.",
    )
    _bullet(
        doc,
        "Mantener respaldo fotográfico de menús (diámetro 127 mm, mortero 3 mm, M45 = 1,5) junto "
        "a las evidencias de este informe.",
    )
    _bullet(
        doc,
        "Monitorear en plataforma WES que el punto 000027-01 conserve coherencia de caudal "
        "respecto de la referencia ESVAL en los días posteriores.",
    )

    _add_heading(doc, "6. Conclusión", level=1)
    _p(
        doc,
        "El desvío inicial no se explica por calidad de señal ni por cableado. La limpieza de "
        "tubería y la silicona mejoraron DN/UP y Q, pero no el volumen. La corrección a "
        f"{DIAMETRO_ESVAL_MM} mm + mortero {ESPESOR_MORTERO_MM} mm redujo la diferencia a ~"
        f"{DIFERENCIA_LITROS_POST_CONFIG} L. El factor de escala calculado "
        f"({M45_CALCULADO}) no pudo aplicarse completo por el máximo del equipo "
        f"({M45_MAXIMO_EQUIPO}); con el tope la diferencia bajó a ~{DIFERENCIA_LITROS_POST_M45} L. "
        f"El aumento de diámetro en {INCREMENTO_DIAMETRO_PCT} % "
        f"({DIAMETRO_ESVAL_MM} → {DIAMETRO_FINAL_MM} mm), según recomendación del manual, cerró "
        "la desviación: ambos medidores quedaron en el mismo ciclo de lectura y sin error entre ellos.",
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
