"""
Comparativo vacaciones CORMUP / Peñalolén — septiembre 2026 (formato formal).

Semana sin control:     07–13/09/2026
Semana con control especial por vacaciones: 14–20/09/2026

Cohorte: colegios con control, excluyendo Tobalaba (pulso) y los tres sin control
(Eduardo de la Barra, Alicura, Likankura). Incluye valorización económica CLP/m³.

Uso:
  python generar_comparativo_cormup_vacaciones_sept2026.py
  python generar_comparativo_cormup_vacaciones_sept2026.py --sin-agregados
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
from zoneinfo import ZoneInfo

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from matplotlib.backends.backend_pdf import PdfPages

from auditoria_cpa_icco_renca_grafico import _vectores_m3h_por_dias
from generar_reporte_word import (
    add_formatted_heading,
    add_logo_to_header,
    add_picture_with_pagination,
    estilizar_tabla_wes,
    format_currency_chilean,
    format_number_chilean,
    generate_aggregated_report,
    get_water_price_per_m3,
)
from wes_estilo_graficos_app import COLOR_BARRA_WES, COLOR_CONSUMO, COLOR_NOCHE

CHILE = ZoneInfo("America/Santiago")
ROOT = Path(__file__).resolve().parent
COMPANY_ID = "000008"
PRECIO_DEFAULT_CLP = 1200.0

SIN_CONTROL_INI = date(2026, 9, 7)
SIN_CONTROL_FIN = date(2026, 9, 13)
CON_CONTROL_INI = date(2026, 9, 14)
CON_CONTROL_FIN = date(2026, 9, 20)

COLEGIOS: List[Tuple[str, str]] = [
    ("000008-01", "Lic. Antonio Hermida F"),
    ("000008-03", "Carlos Fernandez P."),
    ("000008-05", "Santa Maria"),
    ("000008-06", "Luis Arrieta Caña"),
    ("000008-07", "Erasmo Escala"),
    ("000008-09", "Juan Bautista Pasten"),
    ("000008-10", "Matilde Huici Navas"),
    ("000008-11", "CE Valle Hermoso"),
    ("000008-12", "Unión Nacional Árabe"),
    ("000008-14", "Juan Pablo II"),
]

SIN_CONTROL: List[Tuple[str, str]] = [
    ("000008-02", "Eduardo de la Barra"),
    ("000008-08", "Alicura"),
    ("000008-13", "Likankura"),
]

TOBALABA = ("000008-04", "Tobalaba")

HORARIOS_ESPECIALES = [
    "Lic. Antonio Hermida F: corte programado (24 h).",
    "Carlos Fernandez P.: obras — agua habilitada toda la semana (incluye sáb/dom) 09:00–18:00.",
    "Santa Maria: corte programado (24 h).",
    "Luis Arrieta Caña: patinaje lun–mar 17:30–21:00; martes 15/09 habilitación adicional por revisión del colegio.",
    "Erasmo Escala: corte programado (24 h).",
    "Juan Bautista Pasten: obras — agua habilitada toda la semana (incluye sáb/dom) 09:00–18:00.",
    "Matilde Huici Navas: corte programado (24 h).",
    "CE Valle Hermoso: patinaje lun y mié 17:30–21:00; miércoles 16/09 habilitación especial ≈10:00–tarde.",
    "Unión Nacional Árabe: corte programado (24 h).",
    "Juan Pablo II: corte programado (24 h).",
]


@dataclass
class ResultadoColegio:
    node_id: str
    nombre: str
    m3_sin: float
    m3_con: float
    por_dia_sin: List[float] = field(default_factory=list)
    por_dia_con: List[float] = field(default_factory=list)

    @property
    def ahorro_m3(self) -> float:
        return self.m3_sin - self.m3_con

    @property
    def ahorro_pct(self) -> Optional[float]:
        if self.m3_sin <= 1e-9:
            return None
        return 100.0 * (self.m3_sin - self.m3_con) / self.m3_sin


def _rango(ini: date, fin: date) -> List[date]:
    out: List[date] = []
    d = ini
    while d <= fin:
        out.append(d)
        d += timedelta(days=1)
    return out


def _fmt_pct(v: Optional[float]) -> str:
    if v is None:
        return "—"
    return f"{format_number_chilean(v, 1)} %"


def _sombrear(cell, hex_color: str) -> None:
    try:
        shading = parse_xml(
            f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            f'w:val="clear" w:fill="{hex_color}"/>'
        )
        tc_pr = cell._element.get_or_add_tcPr()
        if tc_pr.find(qn("w:shd")) is None:
            tc_pr.append(shading)
    except Exception:
        pass


def _set_cell(
    cell,
    texto: str,
    *,
    bold: bool = False,
    size: int = 8,
    color: Optional[RGBColor] = None,
) -> None:
    cell.text = texto
    for p in cell.paragraphs:
        for run in p.runs:
            run.bold = bold
            run.font.size = Pt(size)
            if color is not None:
                run.font.color.rgb = color


def _consumo_periodo(node_id: str, dias: Sequence[date]) -> Tuple[float, List[float]]:
    vecs = _vectores_m3h_por_dias(node_id, dias)
    por_dia = [float(sum(v)) for v in vecs]
    return float(sum(por_dia)), por_dia


def _evaluar_pares(
    pares: Sequence[Tuple[str, str]],
    max_workers: int,
    etiqueta: str,
) -> List[ResultadoColegio]:
    dias_sin = _rango(SIN_CONTROL_INI, SIN_CONTROL_FIN)
    dias_con = _rango(CON_CONTROL_INI, CON_CONTROL_FIN)
    resultados: Dict[str, ResultadoColegio] = {}

    def _uno(nid: str, nombre: str) -> ResultadoColegio:
        m3_sin, por_sin = _consumo_periodo(nid, dias_sin)
        m3_con, por_con = _consumo_periodo(nid, dias_con)
        return ResultadoColegio(
            node_id=nid,
            nombre=nombre,
            m3_sin=m3_sin,
            m3_con=m3_con,
            por_dia_sin=por_sin,
            por_dia_con=por_con,
        )

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = {ex.submit(_uno, nid, nom): nid for nid, nom in pares}
        for fut in as_completed(futs):
            r = fut.result()
            resultados[r.node_id] = r
            print(
                f"  [{etiqueta}] {r.nombre}: 7–13={r.m3_sin:.2f} m³ | "
                f"14–20={r.m3_con:.2f} m³ | Δ={r.ahorro_m3:.2f} m³ ({_fmt_pct(r.ahorro_pct)})"
            )
    return [resultados[nid] for nid, _ in pares]


def evaluar_colegios(max_workers: int = 8) -> List[ResultadoColegio]:
    return _evaluar_pares(COLEGIOS, max_workers, "OK")


def evaluar_fuera_comparativo(max_workers: int = 4) -> List[ResultadoColegio]:
    return _evaluar_pares([TOBALABA] + list(SIN_CONTROL), max_workers, "INFO fuera")


def precio_referencia_clp() -> float:
    try:
        p = float(get_water_price_per_m3(COMPANY_ID, COLEGIOS[0][0]))
        if p > 0:
            return p
    except Exception:
        pass
    return PRECIO_DEFAULT_CLP


def grafico_barras_colegios(filas: List[ResultadoColegio], out_png: Path) -> Path:
    nombres = [f.nombre for f in filas]
    sin_vals = [f.m3_sin for f in filas]
    con_vals = [f.m3_con for f in filas]
    x = list(range(len(filas)))
    w = 0.38

    fig, ax = plt.subplots(figsize=(12, 6.2))
    ax.bar([i - w / 2 for i in x], sin_vals, width=w, color=COLOR_NOCHE, label="Sin control (7–13/09)")
    ax.bar(
        [i + w / 2 for i in x],
        con_vals,
        width=w,
        color=COLOR_BARRA_WES,
        label="Con control vacaciones (14–20/09)",
    )
    ax.set_xticks(x)
    ax.set_xticklabels(nombres, rotation=35, ha="right", fontsize=8)
    ax.set_ylabel("Consumo (m³)")
    ax.set_title("CORMUP Peñalolén — consumo semanal por colegio (cohorte con control)")
    ax.legend(frameon=False)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    fig.tight_layout()
    out_png.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_png, dpi=150)
    plt.close(fig)
    return out_png


def grafico_totales_diarios(filas: List[ResultadoColegio], out_png: Path) -> Path:
    dias_sin = _rango(SIN_CONTROL_INI, SIN_CONTROL_FIN)
    dias_con = _rango(CON_CONTROL_INI, CON_CONTROL_FIN)
    tot_sin = [sum(f.por_dia_sin[i] for f in filas) for i in range(len(dias_sin))]
    tot_con = [sum(f.por_dia_con[i] for f in filas) for i in range(len(dias_con))]

    fig, ax = plt.subplots(figsize=(11, 5.2))
    ax.plot(
        [d.strftime("%d/%m") for d in dias_sin],
        tot_sin,
        marker="o",
        color=COLOR_NOCHE,
        label="Sin control (7–13/09)",
    )
    ax.plot(
        [d.strftime("%d/%m") for d in dias_con],
        tot_con,
        marker="o",
        color=COLOR_CONSUMO,
        label="Con control vacaciones (14–20/09)",
    )
    ax.set_ylabel("Consumo agregado (m³/día)")
    ax.set_title("CORMUP Peñalolén — total diario de los 10 colegios con control")
    ax.legend(frameon=False)
    ax.grid(True, linestyle=":", alpha=0.4)
    fig.tight_layout()
    fig.savefig(out_png, dpi=150)
    plt.close(fig)
    return out_png


def grafico_ahorro_clp(filas: List[ResultadoColegio], precio: float, out_png: Path) -> Path:
    orden = sorted(filas, key=lambda f: f.ahorro_m3, reverse=True)
    nombres = [f.nombre for f in orden]
    vals = [f.ahorro_m3 * precio for f in orden]
    colors = [COLOR_BARRA_WES if v >= 0 else COLOR_NOCHE for v in vals]

    fig, ax = plt.subplots(figsize=(12, 5.8))
    ax.barh(nombres[::-1], vals[::-1], color=colors[::-1])
    ax.set_xlabel("Ahorro valorizado (CLP)")
    ax.set_title(f"Ahorro económico por colegio (referencia {format_number_chilean(precio, 0)} CLP/m³)")
    ax.grid(axis="x", linestyle=":", alpha=0.4)
    fig.tight_layout()
    fig.savefig(out_png, dpi=150)
    plt.close(fig)
    return out_png


def crear_informe_word(
    filas: List[ResultadoColegio],
    fuera: List[ResultadoColegio],
    chart_barras: Path,
    chart_diario: Path,
    chart_clp: Path,
    out_docx: Path,
    precio: float,
) -> Path:
    tot_sin = sum(f.m3_sin for f in filas)
    tot_con = sum(f.m3_con for f in filas)
    ahorro = tot_sin - tot_con
    pct = (100.0 * ahorro / tot_sin) if tot_sin > 1e-9 else None
    clp_sin = tot_sin * precio
    clp_con = tot_con * precio
    clp_ahorro = ahorro * precio

    tobalaba = next((f for f in fuera if f.node_id == TOBALABA[0]), None)
    sin_ctrl = [f for f in fuera if f.node_id != TOBALABA[0]]

    doc = Document()
    add_logo_to_header(doc)

    title = doc.add_paragraph("Reporte Comparativo — CORMUP (Peñalolén)")
    title.style = "Title"
    for run in title.runs:
        run.font.size = Pt(22)
        run.font.color.rgb = RGBColor(0, 51, 102)

    gen = datetime.now(CHILE).strftime("%d-%m-%Y")
    sub = doc.add_paragraph(
        "MONITOREO WES\n"
        "Análisis consolidado de vacaciones de Fiestas Patrias — colegios con control hidráulico\n"
        f"Periodo sin control: {SIN_CONTROL_INI:%d-%m-%y} – {SIN_CONTROL_FIN:%d-%m-%y}\n"
        f"Periodo con control especial: {CON_CONTROL_INI:%d-%m-%y} – {CON_CONTROL_FIN:%d-%m-%y}\n"
        f"Precio de referencia: {format_number_chilean(precio, 0)} CLP/m³  |  Generado: {gen}"
    )
    sub.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for run in sub.runs:
        run.font.size = Pt(10)

    add_formatted_heading(doc, "1. Resumen ejecutivo", level=1)
    p = doc.add_paragraph(
        f"El presente informe compara el consumo hídrico de {len(filas)} colegios CORMUP de Peñalolén "
        f"durante dos semanas consecutivas de igual duración (7 días): una semana previa sin control "
        f"de vacaciones ({SIN_CONTROL_INI:%d/%m/%Y}–{SIN_CONTROL_FIN:%d/%m/%Y}) y la semana con control "
        f"especial por vacaciones de alumnos ({CON_CONTROL_INI:%d/%m/%Y}–{CON_CONTROL_FIN:%d/%m/%Y}). "
        f"La valorización económica utiliza el precio de referencia del servicio "
        f"({format_number_chilean(precio, 0)} CLP/m³)."
    )
    p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    p2 = doc.add_paragraph(
        f"Consumo semana sin control: {format_number_chilean(tot_sin, 1)} m³ "
        f"({format_currency_chilean(clp_sin)}).\n"
        f"Consumo semana con control vacaciones: {format_number_chilean(tot_con, 1)} m³ "
        f"({format_currency_chilean(clp_con)}).\n"
        f"Ahorro volumétrico: {format_number_chilean(ahorro, 1)} m³"
        + (f" ({_fmt_pct(pct)})." if pct is not None else ".")
        + f"\nAhorro económico estimado: {format_currency_chilean(clp_ahorro)}."
    )
    p2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    p3 = doc.add_paragraph(
        "La semana con control no correspondió a un corte 24 h uniforme: se mantuvieron "
        "habilitaciones por obras, uso de patinaje y revisiones puntuales acordadas con el cliente. "
        "Dichos consumos esperados forman parte del resultado y deben interpretarse como uso "
        "autorizado, no como falla del control."
    )
    p3.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    add_formatted_heading(doc, "2. Alcance del análisis", level=1)
    doc.add_paragraph(
        f"Se incluyen los {len(filas)} establecimientos con equipo de control hidráulico operativo "
        "durante el periodo de vacaciones:"
    )
    for _, nom in COLEGIOS:
        doc.add_paragraph(nom, style="List Bullet")
    p_alc = doc.add_paragraph(
        "Quedan fuera del total de ahorro: (i) Tobalaba, por anomalía de pulso del medidor — "
        "pendiente de revisión técnica; y (ii) Eduardo de la Barra, Alicura y Likankura, "
        "establecimientos que solo cuentan con monitoreo y no disponen de válvula de control. "
        "Su situación se detalla en las secciones 6 y 7."
    )
    p_alc.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    add_formatted_heading(doc, "3. Programa de control especial (14–20/09/2026)", level=1)
    p_h = doc.add_paragraph(
        "Según lo solicitado y confirmado para el periodo de vacaciones, el programa aplicado fue:"
    )
    p_h.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for txt in HORARIOS_ESPECIALES:
        doc.add_paragraph(txt, style="List Bullet")

    add_formatted_heading(doc, "4. Comparación por colegio — volumen y valorización", level=1)
    p4 = doc.add_paragraph(
        "La siguiente tabla ordena los colegios de mayor a menor ahorro volumétrico. "
        "Las columnas de costo expresan el consumo de cada semana valorizado al precio de "
        f"referencia ({format_number_chilean(precio, 0)} CLP/m³)."
    )
    p4.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    headers = [
        "#",
        "Colegio",
        "Sin control\n(m³)",
        "Con control\n(m³)",
        "Ahorro\n(m³)",
        "Ahorro\n(%)",
        "Costo sin\n(CLP)",
        "Costo con\n(CLP)",
        "Ahorro\n(CLP)",
    ]
    orden = sorted(filas, key=lambda f: f.ahorro_m3, reverse=True)
    table = doc.add_table(rows=1 + len(orden) + 1, cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        _set_cell(cell, h, bold=True, size=7, color=RGBColor(255, 255, 255))
        _sombrear(cell, "1F4E79")

    for r_i, f in enumerate(orden, start=1):
        vals = [
            str(r_i),
            f.nombre,
            format_number_chilean(f.m3_sin, 1),
            format_number_chilean(f.m3_con, 1),
            format_number_chilean(f.ahorro_m3, 1),
            _fmt_pct(f.ahorro_pct),
            format_currency_chilean(f.m3_sin * precio),
            format_currency_chilean(f.m3_con * precio),
            format_currency_chilean(f.ahorro_m3 * precio),
        ]
        col_ahorro = RGBColor(0, 97, 0) if f.ahorro_m3 >= 0 else RGBColor(156, 0, 6)
        for c_i, v in enumerate(vals):
            _set_cell(
                table.rows[r_i].cells[c_i],
                v,
                bold=(c_i in (4, 5, 8)),
                size=7,
                color=col_ahorro if c_i in (4, 5, 8) else None,
            )

    totales = [
        "",
        "TOTAL (10 colegios)",
        format_number_chilean(tot_sin, 1),
        format_number_chilean(tot_con, 1),
        format_number_chilean(ahorro, 1),
        _fmt_pct(pct),
        format_currency_chilean(clp_sin),
        format_currency_chilean(clp_con),
        format_currency_chilean(clp_ahorro),
    ]
    for c_i, v in enumerate(totales):
        cell = table.rows[1 + len(orden)].cells[c_i]
        _set_cell(cell, v, bold=True, size=7)
        _sombrear(cell, "D6EAF8")
    estilizar_tabla_wes(table, has_total_row=True)

    add_formatted_heading(doc, "5. Gráficos de consumo y ahorro económico", level=1)
    p5 = doc.add_paragraph(
        "Se presentan el ranking semanal por colegio, la evolución diaria agregada de ambas "
        "semanas y el ahorro valorizado en pesos chilenos."
    )
    p5.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    add_picture_with_pagination(doc, str(chart_barras), width=Inches(6.3))
    doc.add_paragraph("")
    add_picture_with_pagination(doc, str(chart_diario), width=Inches(6.3))
    doc.add_paragraph("")
    add_picture_with_pagination(doc, str(chart_clp), width=Inches(6.3))

    add_formatted_heading(doc, "6. Tobalaba — seguimiento pendiente (pulso)", level=1)
    if tobalaba:
        p_tob = doc.add_paragraph(
            f"El colegio Tobalaba ({tobalaba.node_id}) cuenta en condiciones normales con equipo "
            f"de control; sin embargo, durante este periodo se excluyó del programa de corte y del "
            f"cálculo de ahorro por problemas en el pulso del medidor, que invalidan la serie de "
            f"consumo para evaluación operativa.\n\n"
            f"Consumo registrado (referencial, no validado para control):\n"
            f"• Semana 7–13/09: {format_number_chilean(tobalaba.m3_sin, 1)} m³ "
            f"({format_currency_chilean(tobalaba.m3_sin * precio)}).\n"
            f"• Semana 14–20/09: {format_number_chilean(tobalaba.m3_con, 1)} m³ "
            f"({format_currency_chilean(tobalaba.m3_con * precio)}).\n\n"
            "Se deja constancia de que debe revisarse el tema de Tobalaba (diagnóstico y "
            "corrección del pulso / telemetría) para reincorporarlo al esquema de control "
            "hídrico en los próximos periodos. Mientras no se normalice la medición, no debe "
            "incluirse en rankings de cumplimiento ni en valorizaciones de ahorro."
        )
    else:
        p_tob = doc.add_paragraph(
            "El colegio Tobalaba se excluyó del análisis por problemas de pulso. "
            "Debe hacerse seguimiento técnico para reincorporarlo al control."
        )
    p_tob.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    add_formatted_heading(doc, "7. Colegios sin control hidráulico", level=1)
    p7 = doc.add_paragraph(
        "CORMUP / Peñalolén opera hoy con catorce puntos de monitoreo. Once de ellos disponen "
        "(en condiciones normales) de válvula de control; tres solo cuentan con medición y "
        "alertas, sin capacidad de corte remoto. Estos últimos no formaron parte del programa "
        "de vacaciones ni del cálculo de ahorro:"
    )
    p7.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    headers7 = [
        "Colegio",
        "ID",
        "Situación",
        "Consumo 7–13\n(m³)",
        "Consumo 14–20\n(m³)",
        "Costo 7–13\n(CLP)",
        "Costo 14–20\n(CLP)",
    ]
    table7 = doc.add_table(rows=1 + len(sin_ctrl), cols=len(headers7))
    table7.style = "Table Grid"
    for i, h in enumerate(headers7):
        cell = table7.rows[0].cells[i]
        _set_cell(cell, h, bold=True, size=8, color=RGBColor(255, 255, 255))
        _sombrear(cell, "1F4E79")
    for r_i, f in enumerate(sin_ctrl, start=1):
        vals = [
            f.nombre,
            f.node_id,
            "Solo monitoreo (sin válvula)",
            format_number_chilean(f.m3_sin, 1),
            format_number_chilean(f.m3_con, 1),
            format_currency_chilean(f.m3_sin * precio),
            format_currency_chilean(f.m3_con * precio),
        ]
        for c_i, v in enumerate(vals):
            _set_cell(table7.rows[r_i].cells[c_i], v, size=8)
    estilizar_tabla_wes(table7, has_total_row=False)

    p7b = doc.add_paragraph(
        "Eduardo de la Barra, Alicura y Likankura permanecen fuera del esquema de corte porque "
        "no poseen actuador/válvula WES. Su consumo se monitorea y puede alertarse, pero no es "
        "posible programar habilitaciones ni cortes remotos. Se recomienda mantenerlos visibles "
        "en los reportes de monitoreo y evaluar, con el cliente, la factibilidad de incorporar "
        "control en una etapa posterior."
    )
    p7b.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    add_formatted_heading(doc, "8. Lectura operativa del periodo con control", level=1)
    p8 = doc.add_paragraph(
        "• Hermida, Santa Maria, Erasmo Escala, Matilde Huici, Unión Nacional Árabe y Juan Pablo II "
        "tenían corte programado; residuales en esa semana deben revisarse como eventual fuga o "
        "habilitación no registrada.\n"
        "• Carlos Fernandez y Juan Bautista Pasten concentran consumo diurno esperado por obras "
        "(09:00–18:00 todos los días, incluido fin de semana).\n"
        "• Luis Arrieta y Valle Hermoso tuvieron ventanas de patinaje y habilitaciones puntuales "
        "(martes 15/09 y miércoles 16/09, respectivamente), coherentes con el programa acordado."
    )
    p8.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    add_formatted_heading(doc, "9. Conclusiones y recomendaciones", level=1)
    if ahorro > 0:
        p9 = doc.add_paragraph(
            f"En el conjunto de los {len(filas)} colegios con control operativo, la semana de "
            f"vacaciones con control especial registró {format_number_chilean(ahorro, 1)} m³ menos "
            f"que la semana previa sin control ({_fmt_pct(pct)}), equivalente a un ahorro "
            f"económico estimado de {format_currency_chilean(clp_ahorro)} al precio de referencia "
            f"de {format_number_chilean(precio, 0)} CLP/m³. El resultado incorpora las habilitaciones "
            f"de obra y patinaje acordadas con el cliente."
        )
    else:
        p9 = doc.add_paragraph(
            f"En el conjunto de los {len(filas)} colegios, la semana con control especial no muestra "
            f"ahorro neto frente a la semana sin control (variación "
            f"{format_number_chilean(ahorro, 1)} m³ / {format_currency_chilean(clp_ahorro)}). "
            f"Se recomienda revisar habilitaciones y residuales fuera de ventana."
        )
    p9.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    p9b = doc.add_paragraph(
        "Recomendaciones:\n"
        "1) Mantener el esquema de control en próximos periodos no lectivos, documentando "
        "habilitaciones excepcionales.\n"
        "2) Priorizar la revisión técnica de Tobalaba (pulso/telemetría) para reincorporarlo "
        "al control y a la cuantificación de ahorro.\n"
        "3) Dejar explícito ante el cliente que Eduardo de la Barra, Alicura y Likankura no "
        "tienen control hidráulico; cualquier reducción allí requiere intervención en terreno "
        "o ampliación de infraestructura WES.\n"
        "4) Revisar residuales en colegios con corte 24 h programado (en particular Hermida, "
        "si persistiera consumo fuera de ventana)."
    )
    p9b.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    nota = doc.add_paragraph(
        "Nota metodológica: la valorización en CLP es orientativa y utiliza el precio por m³ "
        f"obtenido desde la plataforma WES para CORMUP ({format_number_chilean(precio, 0)} CLP/m³). "
        "No sustituye la facturación del sanitário ni considera cargos fijos."
    )
    for run in nota.runs:
        run.italic = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(89, 89, 89)

    out_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_docx))
    return out_docx


def convertir_pdf(
    docx_path: Path,
    filas: List[ResultadoColegio],
    fuera: List[ResultadoColegio],
    chart_barras: Path,
    chart_diario: Path,
    chart_clp: Path,
    precio: float,
) -> Path:
    pdf = docx_path.with_suffix(".pdf")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(docx_path.parent),
                    str(docx_path),
                ],
                check=True,
                timeout=180,
                capture_output=True,
            )
            if pdf.exists():
                return pdf
        except Exception as exc:
            print(f"[AVISO] LibreOffice falló: {exc}")

    tot_sin = sum(f.m3_sin for f in filas)
    tot_con = sum(f.m3_con for f in filas)
    ahorro = tot_sin - tot_con
    pct = (100.0 * ahorro / tot_sin) if tot_sin > 1e-9 else None
    tobalaba = next((f for f in fuera if f.node_id == TOBALABA[0]), None)
    sin_ctrl = [f for f in fuera if f.node_id != TOBALABA[0]]

    with PdfPages(pdf) as pages:
        fig = plt.figure(figsize=(11.69, 8.27))
        ax = fig.add_subplot(111)
        ax.axis("off")
        ax.set_title(
            "CORMUP Peñalolén — Comparativo vacaciones (formal)\n"
            f"7–13 vs 14–20/09/2026  |  {precio:.0f} CLP/m³",
            fontsize=12,
            pad=14,
        )
        headers = ["Colegio", "Sin m³", "Con m³", "Ahorro m³", "Ahorro %", "Ahorro CLP"]
        data = [headers]
        for f in sorted(filas, key=lambda x: x.ahorro_m3, reverse=True):
            data.append(
                [
                    f.nombre[:24],
                    f"{f.m3_sin:.1f}",
                    f"{f.m3_con:.1f}",
                    f"{f.ahorro_m3:.1f}",
                    "—" if f.ahorro_pct is None else f"{f.ahorro_pct:.1f}%",
                    f"${f.ahorro_m3 * precio:,.0f}",
                ]
            )
        data.append(
            [
                "TOTAL",
                f"{tot_sin:.1f}",
                f"{tot_con:.1f}",
                f"{ahorro:.1f}",
                "—" if pct is None else f"{pct:.1f}%",
                f"${ahorro * precio:,.0f}",
            ]
        )
        table = ax.table(cellText=data, loc="center", cellLoc="center")
        table.auto_set_font_size(False)
        table.set_fontsize(7.5)
        table.scale(1.15, 1.3)
        for j in range(len(headers)):
            table[0, j].set_facecolor("#1F4E79")
            table[0, j].set_text_props(color="white", weight="bold")
        for i in range(1, len(data)):
            color = (
                "#D6EAF8"
                if i == len(data) - 1
                else ("#C6EFCE" if float(data[i][3]) >= 0 else "#FFC7CE")
            )
            for j in range(len(headers)):
                table[i, j].set_facecolor(color)
        nota_tob = ""
        if tobalaba:
            nota_tob = (
                f" Tobalaba excluido (pulso): 7–13={tobalaba.m3_sin:.1f} m³, "
                f"14–20={tobalaba.m3_con:.1f} m³ — seguimiento pendiente."
            )
        sin_txt = (
            ", ".join(f.nombre for f in sin_ctrl)
            if sin_ctrl
            else "Eduardo de la Barra, Alicura, Likankura"
        )
        ax.text(
            0.02,
            0.03,
            f"Sin control (solo monitoreo): {sin_txt}.{nota_tob}",
            transform=ax.transAxes,
            fontsize=7.5,
        )
        pages.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        for img in (chart_barras, chart_diario, chart_clp):
            fig2 = plt.figure(figsize=(11.69, 8.27))
            ax2 = fig2.add_subplot(111)
            ax2.axis("off")
            ax2.imshow(plt.imread(img))
            pages.savefig(fig2, bbox_inches="tight")
            plt.close(fig2)
    return pdf


def generar_agregados_apoyo() -> List[Path]:
    nodes = [nid for nid, _ in COLEGIOS]
    nota_sin = (
        "Periodo de referencia sin control operativo de vacaciones (07–13/09/2026). "
        "Excluidos: Tobalaba (pulso — seguimiento pendiente), Eduardo de la Barra, Alicura y "
        "Likankura (sin válvula de control)."
    )
    nota_con = (
        "Periodo con control especial por vacaciones (14–20/09/2026). "
        "Excluidos: Tobalaba (pulso — seguimiento pendiente), Eduardo de la Barra, Alicura y "
        "Likankura (sin válvula de control). "
        "Horarios: corte 24 h salvo Carlos Fernandez y Juan Bautista Pasten "
        "(obra agua 09:00–18:00 toda la semana); Luis Arrieta patinaje lun–mar 17:30–21:00 "
        "(+ habilitación mar 15 por revisión); Valle Hermoso patinaje lun y mié 17:30–21:00 "
        "(+ habilitación mié 16 ≈10:00–tarde)."
    )
    outs: List[Path] = []
    for start, end, nota in (
        ("07/09/2026", "13/09/2026", nota_sin),
        ("14/09/2026", "20/09/2026", nota_con),
    ):
        print(f"\n[INFO] Agregado apoyo {start}–{end}")
        path = generate_aggregated_report(
            company_id=COMPANY_ID,
            node_ids=list(nodes),
            start_date=start,
            end_date=end,
            apply_exclusions=False,
            generate_ppt=False,
            nota_contexto_periodo=nota,
            parallel_node_fetch=True,
            max_parallel_workers=8,
        )
        outs.append(Path(path))
        print(f"[OK] Agregado: {path}")
    return outs


def main() -> int:
    ap = argparse.ArgumentParser(description="Comparativo vacaciones CORMUP sept 2026")
    ap.add_argument("--sin-agregados", action="store_true", help="Solo el comparativo (sin agregados Word)")
    ap.add_argument(
        "--workers",
        type=int,
        default=2,
        help="Workers paralelos API (default 2; valores altos pueden devolver ceros)",
    )
    args = ap.parse_args()

    print("=" * 72)
    print("CORMUP Peñalolén — comparativo vacaciones 7–13 vs 14–20/09/2026 (formal)")
    print("=" * 72)

    precio = precio_referencia_clp()
    print(f"[INFO] Precio referencia: {precio:.0f} CLP/m³")

    filas = evaluar_colegios(max_workers=args.workers)
    fuera = evaluar_fuera_comparativo(max_workers=min(4, args.workers))

    ts = datetime.now(CHILE).strftime("%Y%m%d_%H%M")
    out_dir = ROOT / "reports" / "CORMUP" / "VACACIONES" / f"COMPARATIVO_{ts}"
    out_dir.mkdir(parents=True, exist_ok=True)

    payload = {
        "precio_clp_m3": precio,
        "colegios": [asdict(f) for f in filas],
        "fuera_comparativo": [asdict(f) for f in fuera],
    }
    (out_dir / "resultado.json").write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    chart_barras = grafico_barras_colegios(filas, out_dir / "chart_barras_colegios.png")
    chart_diario = grafico_totales_diarios(filas, out_dir / "chart_totales_diarios.png")
    chart_clp = grafico_ahorro_clp(filas, precio, out_dir / "chart_ahorro_clp.png")

    docx = crear_informe_word(
        filas,
        fuera,
        chart_barras,
        chart_diario,
        chart_clp,
        out_dir / "Comparativo_CORMUP_Vacaciones_20260907_20260920.docx",
        precio,
    )
    print(f"[OK] Word: {docx}")

    pdf = convertir_pdf(docx, filas, fuera, chart_barras, chart_diario, chart_clp, precio)
    print(f"[OK] PDF: {pdf}")

    agregados: List[Path] = []
    if not args.sin_agregados:
        try:
            agregados = generar_agregados_apoyo()
        except Exception as exc:
            print(f"[AVISO] No se generaron agregados de apoyo: {exc}")

    tot_sin = sum(f.m3_sin for f in filas)
    tot_con = sum(f.m3_con for f in filas)
    meta = {
        "out_dir": str(out_dir),
        "docx": str(docx),
        "pdf": str(pdf),
        "agregados": [str(p) for p in agregados],
        "precio_clp_m3": precio,
        "total_sin_m3": tot_sin,
        "total_con_m3": tot_con,
        "ahorro_m3": tot_sin - tot_con,
        "ahorro_clp": (tot_sin - tot_con) * precio,
    }
    (out_dir / "meta.json").write_text(json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(meta, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
