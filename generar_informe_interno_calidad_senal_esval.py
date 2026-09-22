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
FOTO_ITRON_1150 = EVIDENCIAS / "itron_1150.jpg"
FOTO_US_1150 = EVIDENCIAS / "ultrasonido_1150.jpg"
FOTO_ITRON_1500 = EVIDENCIAS / "itron_1500.jpg"
FOTO_US_1500 = EVIDENCIAS / "ultrasonido_1500.jpg"

# Validación corta foto vs app (22-09-2026) — cifras leídas de evidencia
VAL_FECHA = "22-09-2026"
HORA_INICIO = "11:50"
HORA_CIERRE = "15:00"

# Itron Flostar S (odómetro mecánico, 6 dígitos negros + 2 rojos)
ITRON_1150_M3 = 383956.27
ITRON_1500_M3 = 383978.01
ITRON_DELTA_M3 = round(ITRON_1500_M3 - ITRON_1150_M3, 2)  # 21.74

# Ultrasónico: NET crudo × 0,01 = m³
US_FLOW_1150 = 17.183
US_NET_CRUDO_1150 = 1811906
US_FLOW_1500 = 10.648
US_NET_CRUDO_1500 = 1814145
US_NET_FACTOR = 0.01
US_NET_1150 = round(US_NET_CRUDO_1150 * US_NET_FACTOR, 2)  # 18119.06
US_NET_1500 = round(US_NET_CRUDO_1500 * US_NET_FACTOR, 2)  # 18141.45
US_DELTA_M3 = round(US_NET_1500 - US_NET_1150, 2)  # 22.39

# Error relativo respecto del ultrasónico: 1 − (Itron / Ultrasonido)
ERROR_RELATIVO = 1.0 - (ITRON_DELTA_M3 / US_DELTA_M3)  # ≈ 0.02903 → 2,90 %
ERROR_RELATIVO_PCT = round(ERROR_RELATIVO * 100, 2)  # 2.90
DELTA_ITRON_VS_US_M3 = round(US_DELTA_M3 - ITRON_DELTA_M3, 2)  # 0.65

# App WES — consumo horario API (000027-01, 22-09-2026)
APP_HORA_11 = 0.25
APP_HORA_12 = 7.66
APP_HORA_13 = 10.01
APP_HORA_14 = 0.55
APP_HORA_15 = 9.87
# 11:50→15:00 = 10/60 de hora 11 + horas 12+13+14 (a las 15:00 exactas no se suma hora 15)
APP_DELTA_1150_1500 = round(
    APP_HORA_11 * (10 / 60) + APP_HORA_12 + APP_HORA_13 + APP_HORA_14, 3
)  # 18.262


def _fmt_m3(x: float, dec: int = 2) -> str:
    """Formato chileno: miles con punto, decimal con coma."""
    s = f"{x:,.{dec}f}"  # 18,119.06 (en-US)
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_pct(x: float, dec: int = 2) -> str:
    return f"{x:.{dec}f}".replace(".", ",")


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
            "Alineados en terreno",
            "Mismo ciclo de lectura post-ajuste; ver §6 para Δ fotográfico 11:50–15:00",
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
        "que ambos medidores lean en el mismo ciclo tras el ajuste final.",
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
        f"{M45_MAXIMO_EQUIPO}; al setear ese tope la diferencia pasó a ~{DIFERENCIA_LITROS_POST_M45} L. "
        f"Siguiendo el manual, se aumentó el diámetro un {INCREMENTO_DIAMETRO_PCT} % "
        f"({DIAMETRO_ESVAL_MM} → {DIAMETRO_FINAL_MM} mm). Tras ese ajuste, ambos medidores quedaron "
        "en el mismo ciclo de lectura en terreno.",
    )
    _p(
        doc,
        f"Validación fotográfica {HORA_INICIO}→{HORA_CIERRE} del {VAL_FECHA}: Itron "
        f"{_fmt_m3(ITRON_DELTA_M3)} m³; ultrasónico (NET×0,01) {_fmt_m3(US_DELTA_M3)} m³; "
        f"error relativo 1 − ({_fmt_m3(ITRON_DELTA_M3)}/{_fmt_m3(US_DELTA_M3)}) = "
        f"{_fmt_pct(ERROR_RELATIVO_PCT)} % (ese intervalo incluye el tramo previo a los ajustes).",
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
        "Itron/ESVAL) quedaron en el mismo ciclo de lectura.",
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
        f"Cierre en terreno: diámetro {DIAMETRO_FINAL_MM} mm (+{INCREMENTO_DIAMETRO_PCT} %) dejó "
        "ambos medidores en el mismo ciclo de lectura.",
    )
    _bullet(
        doc,
        f"Validación fotográfica {HORA_INICIO}→{HORA_CIERRE}: Itron {_fmt_m3(ITRON_DELTA_M3)} m³ vs "
        f"ultrasónico {_fmt_m3(US_DELTA_M3)} m³; error "
        f"1 − ({_fmt_m3(ITRON_DELTA_M3)}/{_fmt_m3(US_DELTA_M3)}) = {_fmt_pct(ERROR_RELATIVO_PCT)} % "
        "(intervalo mixto pre/post ajuste).",
    )

    _add_heading(doc, "5. Estado y seguimiento", level=1)
    _bullet(
        doc,
        "Intervención cerrada en terreno tras diámetro 127 mm + M45 = 1,5 + mortero 3 mm.",
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

    _add_heading(doc, "6. Validación cuantitativa: lecturas 11:50 vs 15:00", level=1)
    _p(
        doc,
        f"Fecha {VAL_FECHA}, nodo {NODE_ID} ({NODO_NOMBRE}). Se leyeron de fotografía el odómetro "
        f"Itron Flostar S y el totalizador NET del ultrasónico a las {HORA_INICIO} (inicio de la "
        f"jornada / pre-ajuste) y a las {HORA_CIERRE} (cierre). El NET del ultrasónico se convierte "
        f"a m³ multiplicando por {US_NET_FACTOR} (factor del display: «×0,01 m³»).",
    )
    _p(
        doc,
        "Importante: el intervalo 11:50–15:00 abarca el periodo de corrección (diámetro, mortero, "
        "M45 y +8 % de diámetro). Por tanto el Δ de ese tramo es un balance del día de trabajo, "
        "no una prueba de calibración post-ajuste puro.",
    )

    _add_heading(doc, "6.1 Medidor Itron Flostar S (referencia ESVAL)", level=2)
    tbl_it = doc.add_table(rows=4, cols=2)
    tbl_it.style = "Table Grid"
    _fill_header_row(tbl_it, ["Concepto", "Valor"])
    _fill_row(tbl_it, 1, [f"Lectura {HORA_INICIO}", f"{_fmt_m3(ITRON_1150_M3)} m³"])
    _fill_row(tbl_it, 2, [f"Lectura {HORA_CIERRE}", f"{_fmt_m3(ITRON_1500_M3)} m³"])
    _fill_row(
        tbl_it,
        3,
        [
            f"Δ Itron ({HORA_CIERRE} − {HORA_INICIO})",
            f"{_fmt_m3(ITRON_DELTA_M3)} m³",
        ],
    )
    doc.add_paragraph()
    _p(
        doc,
        f"Cálculo: {_fmt_m3(ITRON_1500_M3)} − {_fmt_m3(ITRON_1150_M3)} = "
        f"{_fmt_m3(ITRON_DELTA_M3)} m³. Lectura de odómetro: 6 dígitos negros (m³ enteros) + "
        "2 dígitos rojos (decimales).",
    )

    _add_heading(doc, "6.2 Medidor ultrasónico (NET × 0,01)", level=2)
    tbl_us = doc.add_table(rows=5, cols=2)
    tbl_us.style = "Table Grid"
    _fill_header_row(tbl_us, ["Concepto", "Valor"])
    _fill_row(
        tbl_us,
        1,
        [
            f"{HORA_INICIO} — Flow / NET crudo",
            f"{_fmt_m3(US_FLOW_1150, 3)} m³/h · NET {US_NET_CRUDO_1150:,}".replace(",", "."),
        ],
    )
    _fill_row(
        tbl_us,
        2,
        [
            f"{HORA_INICIO} — NET × 0,01",
            f"{US_NET_CRUDO_1150:,} × 0,01 = {_fmt_m3(US_NET_1150)} m³".replace(",", "."),
        ],
    )
    _fill_row(
        tbl_us,
        3,
        [
            f"{HORA_CIERRE} — Flow / NET crudo",
            f"{_fmt_m3(US_FLOW_1500, 3)} m³/h · NET {US_NET_CRUDO_1500:,}".replace(",", "."),
        ],
    )
    _fill_row(
        tbl_us,
        4,
        [
            f"{HORA_CIERRE} — NET × 0,01 y Δ",
            f"{US_NET_CRUDO_1500:,} × 0,01 = {_fmt_m3(US_NET_1500)} m³ → "
            f"Δ = {_fmt_m3(US_DELTA_M3)} m³".replace(",", "."),
        ],
    )
    doc.add_paragraph()
    _p(
        doc,
        f"Cálculo Δ ultrasónico: {_fmt_m3(US_NET_1500)} − {_fmt_m3(US_NET_1150)} = "
        f"{_fmt_m3(US_DELTA_M3)} m³.",
    )

    _add_heading(doc, "6.3 Porcentaje de error (Itron vs ultrasónico)", level=2)
    _p(
        doc,
        "Tomando el Δ del ultrasónico como referencia del tramo fotográfico:",
    )
    _p(
        doc,
        f"Error = 1 − (Δ Itron / Δ Ultrasónico) = 1 − ({_fmt_m3(ITRON_DELTA_M3)} / "
        f"{_fmt_m3(US_DELTA_M3)}) = 1 − {_fmt_pct(ITRON_DELTA_M3 / US_DELTA_M3, 4)} = "
        f"{_fmt_pct(ERROR_RELATIVO_PCT)} %.",
        bold=False,
    )
    tbl_err = doc.add_table(rows=4, cols=2)
    tbl_err.style = "Table Grid"
    _fill_header_row(tbl_err, ["Métrica", "Valor"])
    _fill_row(tbl_err, 1, ["Δ Itron", f"{_fmt_m3(ITRON_DELTA_M3)} m³"])
    _fill_row(tbl_err, 2, ["Δ Ultrasónico (NET×0,01)", f"{_fmt_m3(US_DELTA_M3)} m³"])
    _fill_row(
        tbl_err,
        3,
        [
            "Diferencia absoluta | US − Itron |",
            f"{_fmt_m3(DELTA_ITRON_VS_US_M3)} m³",
        ],
    )
    doc.add_paragraph()
    _bullet(
        doc,
        f"Error relativo: {_fmt_pct(ERROR_RELATIVO_PCT)} % "
        f"(Itron midió {_fmt_pct(ERROR_RELATIVO_PCT)} % menos que el ultrasónico en ese tramo).",
    )

    _add_heading(doc, "6.4 Contraste con app WES (mismo día)", level=2)
    tbl_inst = doc.add_table(rows=3, cols=4)
    tbl_inst.style = "Table Grid"
    _fill_header_row(tbl_inst, ["Hora", "Foto ultrasónico", "App WES (m³ en esa hora)", "Veredicto"])
    _fill_row(
        tbl_inst,
        1,
        [
            HORA_INICIO,
            f"Flow {_fmt_m3(US_FLOW_1150, 3)} m³/h · NET {_fmt_m3(US_NET_1150)} m³",
            f"Hora 11: {_fmt_m3(APP_HORA_11)} m³",
            "No cuadra (config aún incorrecta)",
        ],
    )
    _fill_row(
        tbl_inst,
        2,
        [
            HORA_CIERRE,
            f"Flow {_fmt_m3(US_FLOW_1500, 3)} m³/h · NET {_fmt_m3(US_NET_1500)} m³",
            f"Hora 15: {_fmt_m3(APP_HORA_15)} m³",
            "Cuadra (desvío ~7,3 % vs caudal de foto)",
        ],
    )
    doc.add_paragraph()
    tbl_vol = doc.add_table(rows=4, cols=2)
    tbl_vol.style = "Table Grid"
    _fill_header_row(tbl_vol, ["Fuente", "Δ volumen 11:50 → 15:00"])
    _fill_row(tbl_vol, 1, ["Itron (odómetro)", f"{_fmt_m3(ITRON_DELTA_M3)} m³"])
    _fill_row(tbl_vol, 2, ["Ultrasónico (NET × 0,01)", f"{_fmt_m3(US_DELTA_M3)} m³"])
    _fill_row(
        tbl_vol,
        3,
        [
            "App WES (10/60·h11 + h12 + h13 + h14)",
            f"{_fmt_m3(APP_DELTA_1150_1500, 3)} m³",
        ],
    )
    doc.add_paragraph()
    _p(
        doc,
        f"La app ({_fmt_m3(APP_DELTA_1150_1500, 3)} m³) queda bajo el NET del ultrasónico "
        f"({_fmt_m3(US_DELTA_M3)} m³) porque el tramo incluye la mañana mal configurada y la hora 14 "
        f"casi en cero ({_fmt_m3(APP_HORA_14)} m³), coherente con trabajo en terreno. A las "
        f"{HORA_INICIO} la app no reflejaba el caudal real ({_fmt_m3(US_FLOW_1150, 3)} m³/h en foto vs "
        f"{_fmt_m3(APP_HORA_11)} m³ en hora 11); a las {HORA_CIERRE} el caudal de foto "
        f"({_fmt_m3(US_FLOW_1500, 3)} m³/h) es coherente con la hora 15 de app ({_fmt_m3(APP_HORA_15)} m³).",
    )

    _add_heading(doc, "6.5 Evidencia fotográfica", level=2)
    fotos_val = [
        (
            FOTO_ITRON_1150,
            f"Figura 3. Itron Flostar S — {HORA_INICIO}. Lectura {_fmt_m3(ITRON_1150_M3)} m³.",
        ),
        (
            FOTO_US_1150,
            f"Figura 4. Ultrasónico FZ ESVAL — {HORA_INICIO}. Flow {_fmt_m3(US_FLOW_1150, 3)} m³/h · "
            f"NET {US_NET_CRUDO_1150} × 0,01 = {_fmt_m3(US_NET_1150)} m³.",
        ),
        (
            FOTO_ITRON_1500,
            f"Figura 5. Itron Flostar S — {HORA_CIERRE}. Lectura {_fmt_m3(ITRON_1500_M3)} m³.",
        ),
        (
            FOTO_US_1500,
            f"Figura 6. Ultrasónico FZ ESVAL — {HORA_CIERRE}. Flow {_fmt_m3(US_FLOW_1500, 3)} m³/h · "
            f"NET {US_NET_CRUDO_1500} × 0,01 = {_fmt_m3(US_NET_1500)} m³.",
        ),
    ]
    for src, caption in fotos_val:
        if not src.is_file():
            continue
        dst = out_dir / src.name
        if not dst.is_file():
            shutil.copy2(src, dst)
        _add_figure(doc, dst, caption, width_cm=11.5)

    _add_heading(doc, "7. Conclusión", level=1)
    _p(
        doc,
        "El desvío inicial no se explica por calidad de señal ni por cableado. La limpieza de "
        "tubería y la silicona mejoraron DN/UP y Q, pero no el volumen. La corrección a "
        f"{DIAMETRO_ESVAL_MM} mm + mortero {ESPESOR_MORTERO_MM} mm redujo la diferencia a ~"
        f"{DIFERENCIA_LITROS_POST_CONFIG} L. El factor de escala calculado "
        f"({_fmt_m3(M45_CALCULADO, 4)}) no pudo aplicarse completo por el máximo del equipo "
        f"({_fmt_m3(M45_MAXIMO_EQUIPO, 1)}); con el tope la diferencia bajó a ~"
        f"{DIFERENCIA_LITROS_POST_M45} L. El aumento de diámetro en {INCREMENTO_DIAMETRO_PCT} % "
        f"({DIAMETRO_ESVAL_MM} → {DIAMETRO_FINAL_MM} mm) alineó ambos medidores en el mismo ciclo "
        "de lectura en terreno.",
    )
    _p(
        doc,
        f"Sobre el tramo fotográfico {HORA_INICIO}–{HORA_CIERRE} del {VAL_FECHA}: pasaron "
        f"{_fmt_m3(ITRON_DELTA_M3)} m³ por Itron y {_fmt_m3(US_DELTA_M3)} m³ por ultrasónico "
        f"(NET×0,01), con error relativo "
        f"1 − ({_fmt_m3(ITRON_DELTA_M3)}/{_fmt_m3(US_DELTA_M3)}) = {_fmt_pct(ERROR_RELATIVO_PCT)} % "
        f"(diferencia absoluta {_fmt_m3(DELTA_ITRON_VS_US_M3)} m³). Ese porcentaje corresponde al "
        "intervalo completo de la visita (incluye pre-ajuste). La app WES a las 15:00 ya es "
        "coherente con el caudal instantáneo del ultrasónico.",
    )

    pie = doc.add_paragraph()
    pie.paragraph_format.space_before = Pt(18)
    run = pie.add_run(
        "Documento generado para uso interno del equipo WES. No constituye informe de cliente "
        "ni certificado de calibración. Cifras de §6 tomadas de evidencia fotográfica del "
        f"{VAL_FECHA}; consumo horario app vía API WES nodo {NODE_ID}."
    )
    _set_run_font(run, size=9, color=_MUTED)

    doc.save(out_docx)
    print(f"[OK] Word: {out_docx}")
    print(
        f"[CIFRAS] Itron Δ={ITRON_DELTA_M3} | US Δ={US_DELTA_M3} | "
        f"error={ERROR_RELATIVO_PCT}% | |US-Itron|={DELTA_ITRON_VS_US_M3}"
    )
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
