"""
Informe interno — Matriz ESVAL (Fundo Zapallar).
Versión completa + cálculo de validación (estilo CUR), sin fotos.

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
DN_UP_INICIAL, DN_UP_POST = 76.7, 80
Q_INICIAL, Q_POST = 93, 96

LITROS_ESVAL_ITRON_1 = 230
LITROS_ULTRASONIDO_1 = 100
DIFERENCIA_LITROS_POST_CONFIG = 40
DIFERENCIA_LITROS_POST_M45 = 30

DIAMETRO_CONFIG_ERRONEO_MM = 110
DIAMETRO_ESVAL_MM = 118
DIAMETRO_FINAL_MM = 127
INCREMENTO_DIAMETRO_PCT = 8
ESPESOR_MORTERO_MM = 3
MATERIAL = "Fierro dúctil (EN545)"
MARCADO_FISICO = "EN545 / DN100 / PN16"

M45_FACTOR_ACTUAL = 0.5934
M45_CAUDAL_TURBINA_LPM = 460.0
M45_CAUDAL_ULTRASONIDO_LPM = 87.05
M45_CALCULADO = 3.1362
M45_MAXIMO_EQUIPO = 1.5

# Validación 22-09-2026 — 11:50 → 15:00
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


def _fmt(x: float, dec: int = 2) -> str:
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _font(run, *, size: int = 11, bold: bool = False, color: RGBColor | None = None) -> None:
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
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = para.add_run(text)
    _font(r, size=11)


def _bullet(doc: Document, text: str) -> None:
    para = doc.add_paragraph(style="List Bullet")
    para.paragraph_format.space_after = Pt(3)
    para.clear()
    r = para.add_run(text)
    _font(r, size=11)


def _fill_header(tbl, headers: list[str]) -> None:
    for j, h in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        cell.text = ""
        r = cell.paragraphs[0].add_run(h)
        _font(r, size=10, bold=True, color=_HEADING)


def _fill_row(tbl, i: int, values: list[str], *, bold: bool = False) -> None:
    for j, val in enumerate(values):
        cell = tbl.rows[i].cells[j]
        cell.text = ""
        r = cell.paragraphs[0].add_run(val)
        _font(r, size=10, bold=bold)


def _merge_title(tbl, text: str, cols: int) -> None:
    cell = tbl.rows[0].cells[0]
    for j in range(1, cols):
        cell.merge(tbl.rows[0].cells[j])
    cell.text = ""
    r = cell.paragraphs[0].add_run(text)
    _font(r, size=11, bold=True, color=_HEADING)
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER


def _set_row(row, values: list[str], *, bold: bool = False, size: int = 9) -> None:
    for j, v in enumerate(values):
        cell = row.cells[j]
        cell.text = ""
        r = cell.paragraphs[0].add_run(v)
        _font(r, size=size, bold=bold, color=_HEADING if bold else None)


def _tabla_validacion(
    doc: Document,
    titulo: str,
    lectura_ini: float,
    lectura_fin: float,
    consumo: float,
) -> None:
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
    _set_row(tbl.rows[1], headers, bold=True, size=9)
    _set_row(
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
    out_docx = out_dir / f"Informe_Interno_Calidad_Senal_ESVAL_{stamp}.docx"

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)

    try:
        from generar_reporte_word import add_logo_to_header

        add_logo_to_header(doc)
    except Exception as exc:
        print(f"[WARN] Logo: {exc}")

    title = doc.add_heading(
        "Informe interno — Revisión y validación Matriz ESVAL",
        level=0,
    )
    for run in title.runs:
        run.font.color.rgb = _HEADING
        run.font.name = "Calibri"

    meta = doc.add_paragraph()
    for line in (
        f"Empresa: {EMPRESA} ({COMPANY_ID})",
        f"Punto: {NODO_NOMBRE} ({NODE_ID})",
        f"Validación: {VAL_FECHA} · {HORA_INI} → {HORA_FIN}",
        "Clasificación: Uso interno WES",
    ):
        r = meta.add_run(line + "\n")
        _font(r, size=10, color=_MUTED)

    # —— 1. Objetivo ——
    _h(doc, "1. Objetivo", level=1)
    _p(
        doc,
        "Documentar las acciones realizadas en terreno sobre el medidor ultrasónico de "
        f"{NODO_NOMBRE} ante un desvío respecto del medidor de turbina Itron de ESVAL, "
        "y presentar el cálculo de validación del periodo revisado.",
    )

    # —— 2. Resumen ——
    _h(doc, "2. Resumen", level=1)
    _p(
        doc,
        "La calidad de señal (DN/UP y Q) estaba dentro del umbral del fabricante. "
        f"La primera comparación volumétrica mostró un desvío importante "
        f"({LITROS_ESVAL_ITRON_1} L Itron vs {LITROS_ULTRASONIDO_1} L ultrasónico). "
        "Se verificaron cables, se limpió la tubería y se renovó la silicona; "
        "luego se corrigió la configuración de cañería y el factor de escala. "
        f"En la validación del {VAL_FECHA} el error entre ambos medidores fue de "
        f"{_fmt(ERROR_PCT, 1)} %, aceptable según las tolerancias de fabricante.",
    )

    # —— 3. Actividades ——
    _h(doc, "3. Actividades realizadas", level=1)

    _h(doc, "3.1 Calidad de señal", level=2)
    _p(
        doc,
        "En el menú de diagnóstico se evaluó la calidad de señal entre transductores (DN/UP) "
        "y la calidad de sonido (Q). Ambos valores iniciales superaban el umbral del fabricante "
        f"(> {UMBRAL_FABRICANTE}).",
    )
    tbl_q = doc.add_table(rows=3, cols=4)
    tbl_q.style = "Table Grid"
    _fill_header(tbl_q, ["Parámetro", "Inicial", "Tras limpieza + silicona", "Recomendación"])
    _fill_row(tbl_q, 1, ["DN / UP", f"{DN_UP_INICIAL}", f"{DN_UP_POST}", f"> {UMBRAL_FABRICANTE}"])
    _fill_row(tbl_q, 2, ["Q", f"{Q_INICIAL}", f"{Q_POST}", f"> {UMBRAL_FABRICANTE}"])
    doc.add_paragraph()

    _h(doc, "3.2 Comparación inicial por pulso", level=2)
    _p(
        doc,
        "Se contrastó el volumen del ultrasónico WES contra el medidor de turbina Itron "
        "(referencia ESVAL) en la misma ventana de prueba.",
    )
    tbl_c = doc.add_table(rows=4, cols=3)
    tbl_c.style = "Table Grid"
    delta0 = LITROS_ESVAL_ITRON_1 - LITROS_ULTRASONIDO_1
    _fill_header(tbl_c, ["Equipo", "Volumen", "Observación"])
    _fill_row(tbl_c, 1, ["Itron / ESVAL", f"{LITROS_ESVAL_ITRON_1} L", "Referencia"])
    _fill_row(tbl_c, 2, ["Ultrasónico WES", f"{LITROS_ULTRASONIDO_1} L", "Misma ventana"])
    _fill_row(tbl_c, 3, ["Diferencia", f"{delta0} L", "Desvío relevante"])
    doc.add_paragraph()

    _h(doc, "3.3 Continuidad de cables", level=2)
    _p(
        doc,
        "Se verificó la continuidad eléctrica de los cables de los transductores. "
        "Resultado: correcta. No explica el desvío de volumen.",
    )

    _h(doc, "3.4 Limpieza de tubería y silicona", level=2)
    _p(
        doc,
        "Antes de renovar la silicona se limpió la superficie de la tubería en la zona de "
        "los transductores. Tras el cambio de silicona mejoraron DN/UP y Q, pero no se resolvió "
        "la discrepancia volumétrica.",
    )
    _bullet(doc, f"DN / UP: {DN_UP_INICIAL} → {DN_UP_POST}")
    _bullet(doc, f"Q: {Q_INICIAL} → {Q_POST}")

    _h(doc, "3.5 Configuración de cañería", level=2)
    _p(
        doc,
        "Se revisó el material en obra y se contrastó con la información de la sanitaria ESVAL.",
    )
    _bullet(doc, f"Marcado físico: {MARCADO_FISICO} ({MATERIAL}).")
    _bullet(
        doc,
        f"Diámetro según ESVAL: {_fmt(DIAMETRO_ESVAL_MM, 0)} mm "
        f"(estaba en {_fmt(DIAMETRO_CONFIG_ERRONEO_MM, 0)} mm).",
    )
    _bullet(
        doc,
        f"Lining: mortero inyectado — menú de cobertura con espesor {ESPESOR_MORTERO_MM} mm.",
    )
    _p(
        doc,
        f"Tras aplicar {DIAMETRO_ESVAL_MM} mm + mortero {ESPESOR_MORTERO_MM} mm, "
        f"la diferencia bajó a aproximadamente {DIFERENCIA_LITROS_POST_CONFIG} L.",
    )

    _h(doc, "3.6 Factor de escala M45", level=2)
    _p(
        doc,
        "Con la prueba de 1 minuto (caudal turbina vs caudal ultrasónico menú 00) se calculó "
        "el nuevo factor de escala:",
    )
    _p(
        doc,
        f"Nuevo M45 = {_fmt(M45_FACTOR_ACTUAL, 4)} × "
        f"({_fmt(M45_CAUDAL_TURBINA_LPM, 0)} L/min turbina / "
        f"{_fmt(M45_CAUDAL_ULTRASONIDO_LPM, 2)} L/min ultrasónico) = "
        f"{_fmt(M45_CALCULADO, 4)}.",
    )
    _p(
        doc,
        f"El valor calculado supera el máximo del equipo ({_fmt(M45_MAXIMO_EQUIPO, 1)}). "
        f"Se seteó el tope. La diferencia bajó a ~{DIFERENCIA_LITROS_POST_M45} L, aún con desviación.",
    )

    _h(doc, "3.7 Ajuste final de diámetro", level=2)
    _p(
        doc,
        f"Según recomendación del manual, se aumentó el diámetro un {INCREMENTO_DIAMETRO_PCT} %: "
        f"{DIAMETRO_ESVAL_MM} → {DIAMETRO_FINAL_MM} mm. Tras ese ajuste ambos medidores "
        "quedaron en el mismo ciclo de lectura en terreno.",
    )

    tbl_cfg = doc.add_table(rows=5, cols=3)
    tbl_cfg.style = "Table Grid"
    _fill_header(tbl_cfg, ["Parámetro", "Antes / intermedio", "Final"])
    _fill_row(tbl_cfg, 1, ["Material", "Sin validar", MATERIAL])
    _fill_row(
        tbl_cfg,
        2,
        [
            "Diámetro (mm)",
            f"{DIAMETRO_CONFIG_ERRONEO_MM} → {DIAMETRO_ESVAL_MM}",
            f"{DIAMETRO_FINAL_MM}",
        ],
    )
    _fill_row(tbl_cfg, 3, ["Cobertura", "Sin mortero", f"Mortero {ESPESOR_MORTERO_MM} mm"])
    _fill_row(
        tbl_cfg,
        4,
        [
            "Factor M45",
            f"{_fmt(M45_FACTOR_ACTUAL, 4)} → calc. {_fmt(M45_CALCULADO, 4)}",
            f"Tope {_fmt(M45_MAXIMO_EQUIPO, 1)}",
        ],
    )
    doc.add_paragraph()

    # —— 4. Validación (cálculo, sin fotos) ——
    _h(doc, "4. Cálculo de validación", level=1)
    _p(
        doc,
        "En las siguientes tablas se muestran las lecturas del medidor Itron y del ultrasónico WES "
        f"entre las {HORA_INI} y las {HORA_FIN} del {VAL_FECHA}. El consumo de ese periodo se compara "
        "para determinar la exactitud del monitoreo. "
        "En el ultrasónico el totalizador NET se convierte a m³ multiplicando por 0,01.",
    )

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
    _set_row(t_err.rows[0], ["Total Itron (lectura)", f"{_fmt(ITRON_DELTA)} m³"], size=10)
    _set_row(t_err.rows[1], ["Total ultrasónico WES", f"{_fmt(US_DELTA)} m³"], size=10)
    _set_row(t_err.rows[2], ["% Error", f"{_fmt(ERROR_PCT, 1)} %"], bold=True, size=10)
    doc.add_paragraph()

    _p(
        doc,
        f"Cálculo: 1 − ({_fmt(ITRON_DELTA)} / {_fmt(US_DELTA)}) = {_fmt(ERROR_PCT, 1)} %.",
    )
    _p(
        doc,
        f"En base a las lecturas, el rango de error entre el medidor Itron y el ultrasónico WES "
        f"es de un {_fmt(ERROR_PCT, 1)} %. Los valores no son idénticos: el fabricante del "
        f"ultrasónico declara un rango de error del {TOL_US_PCT} % y el de Itron del "
        f"{TOL_ITRON_PCT} %. Un {_fmt(ERROR_PCT, 1)} % se considera aceptable dentro de esos márgenes.",
    )

    # —— 5. Conclusión ——
    _h(doc, "5. Conclusión", level=1)
    _p(
        doc,
        "El desvío inicial no se debió a calidad de señal ni a cableado. La limpieza y la silicona "
        "mejoraron los indicadores de señal, pero no el volumen. La corrección de diámetro, mortero "
        f"y factor de escala (tope {_fmt(M45_MAXIMO_EQUIPO, 1)}), más el ajuste final a "
        f"{DIAMETRO_FINAL_MM} mm, alineó ambos medidores en terreno. "
        f"La validación del {VAL_FECHA} arroja un error de {_fmt(ERROR_PCT, 1)} %, "
        "aceptable según las tolerancias de fabricante.",
    )

    pie = doc.add_paragraph()
    pie.paragraph_format.space_before = Pt(14)
    r = pie.add_run(
        f"Documento interno WES · {EMPRESA} · generado {datetime.now().strftime('%d-%m-%Y %H:%M')}."
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
        / f"calidad_senal_ultrasonido_{stamp}"
    )
    docx_path = generar_informe(out_dir)
    _intentar_pdf(docx_path)
    print(f"[OK] Carpeta: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
