# -*- coding: utf-8 -*-
"""
Informe de ahorro Parque Arauco para reunión — estilo reporte agregado.

Destaca lo defendible:
  - MAE: ahorro ya logrado (Estanque Sur + Norte; Pizza Hut = control puesto, noche residual)
  - Buenaventura: ahorro ya logrado (San Ignacio 500, control 17/07)
  - Quilicura (MAQ): propuesta on/off 00–06 en Matriz + proyección
  - Kennedy (PAK): propuesta on/off 00–06 en Bazar Gourmet + proyección

  python3 generar_informe_ahorro_pa_reunion.py
  python3 generar_informe_ahorro_pa_reunion.py --hasta 08/09/2026
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

from generar_consolidado_pa_ppt import (
    JSON_DAILY,
    JSON_NIGHT,
    OUT_DIR,
    cargar_diario,
    refrescar_diario,
    refrescar_noches,
    sumar_rango,
)
from generar_ppt_recorrido_ejecutivo_pa import (
    BAZAR,
    CTRL_NORTE,
    CTRL_PIZZA,
    CTRL_SI500,
    DL_KENNEDY,
    MAQ_ALZA,
    PIZZA_NOCHES_ATIPICAS,
    SUR_REPARACION,
    TARIFA_CLP_M3,
    UMBRAL_MAE_BANOS_DIA,
    UMBRAL_MAE_NORTE_DIA,
    UMBRAL_MAE_PIZZA_DIA,
    UMBRAL_MAE_SUR_DIA,
    UMBRAL_MAQ_DIA,
    UMBRAL_PAK_BAZAR_DIA,
    UMBRAL_PAK_DL_DIA,
    UMBRAL_PAK_KEN_DIA,
    UMBRAL_SI300_DIA,
    UMBRAL_SI500_DIA,
)
from generar_reporte_word import add_logo_to_header, format_number_chilean

NAVY = RGBColor(13, 59, 102)
GOLD = RGBColor(201, 162, 39)
GRAY = RGBColor(80, 80, 80)
SI500 = "000025-18"
MATRIZ_MAQ = "000025-13"
CHARTS = OUT_DIR / "charts_informe_reunion"


def fn(v: float, d: int = 1) -> str:
    return format_number_chilean(float(v), d)


def clp(m3: float) -> str:
    return f"${fn(float(m3) * TARIFA_CLP_M3, 0)}"


def _mediana(vals: List[float]) -> float:
    xs = sorted(float(v) for v in vals)
    if not xs:
        return 0.0
    n = len(xs)
    mid = n // 2
    if n % 2:
        return xs[mid]
    return (xs[mid - 1] + xs[mid]) / 2.0


def _rango(d0: date, d1: date) -> List[date]:
    out = []
    d = d0
    while d <= d1:
        out.append(d)
        d += timedelta(days=1)
    return out


def _n06_serie(hourly: Dict[str, Dict[str, float]], nid: str) -> Dict[str, float]:
    return dict((hourly.get(nid) or {}))


def _stats_control(
    serie: Dict[str, float],
    ctrl: date,
    hasta: date,
    *,
    excluir: set | None = None,
    lookback_dias: int | None = 14,
) -> Dict[str, float]:
    """Mediana de noche 00–06 antes vs con control.

    Si ``lookback_dias`` está definido, el «antes» es solo esa ventana previa
    (como el PPT de SI500: 7 días). El «después» son todas las noches hasta ``hasta``.
    """
    pre, post = [], []
    pre0 = (ctrl - timedelta(days=lookback_dias)) if lookback_dias else date.min
    for iso, v in serie.items():
        d = date.fromisoformat(iso)
        if d > hasta:
            continue
        if excluir and d in excluir:
            continue
        fv = float(v)
        if pre0 <= d < ctrl:
            pre.append(fv)
        elif d >= ctrl:
            post.append(fv)
    pre_m = _mediana(pre)
    post_m = _mediana(post)
    ahor = max(0.0, pre_m - post_m)
    return {
        "pre": pre_m,
        "post": post_m,
        "ahorro_noche": ahor,
        "n_pre": float(len(pre)),
        "n_post": float(len(post)),
        "ahorro_mes": ahor * 30.0,
        "ahorro_acum": ahor * float(len(post)),
    }


def _stats_sur(daily: Dict[str, float], hasta: date) -> Dict[str, float]:
    pre, post = [], []
    for iso, v in daily.items():
        d = date.fromisoformat(iso)
        if d > hasta:
            continue
        fv = float(v or 0)
        if d < SUR_REPARACION:
            pre.append(fv)
        elif d > SUR_REPARACION:
            post.append(fv)
    pre_m = _mediana(pre)
    post_m = _mediana(post)
    baja = max(0.0, pre_m - post_m)
    n_post = float(len(post))
    return {
        "pre": pre_m,
        "post": post_m,
        "baja_dia": baja,
        "n_post": n_post,
        "ahorro_mes": baja * 30.0,
        "ahorro_acum": baja * n_post,
    }


def _stats_propuesta_noche(serie: Dict[str, float], d0: date, hasta: date) -> Dict[str, float]:
    vals = []
    for iso, v in serie.items():
        d = date.fromisoformat(iso)
        if d0 <= d <= hasta:
            vals.append(float(v))
    med = _mediana(vals)
    residual = 2.0 if med > 6 else 0.5
    ahor = max(0.0, med - residual)
    return {
        "noche": med,
        "residual": residual,
        "ahorro_noche": ahor,
        "ahorro_mes": ahor * 30.0,
        "ahorro_bruto_mes": med * 30.0,
        "n": float(len(vals)),
    }


def _shade(cell, hex_color: str) -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def _set_run(run, text: str, *, size=11, bold=False, color=NAVY, name="Calibri") -> None:
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = name


def _p(doc: Document, text: str, *, size=11, bold=False, color=NAVY, justify=True) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_run(p.add_run(text), text, size=size, bold=bold, color=color)


def _h(doc: Document, text: str, level: int = 1) -> None:
    p = doc.add_heading(text, level=level)
    for r in p.runs:
        r.font.color.rgb = NAVY
        r.font.name = "Calibri"


def _tabla(doc: Document, headers: List[str], rows: List[List[str]], col_w: List[float] | None = None):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = tbl.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        _set_run(p.add_run(h), h, size=10, bold=True, color=RGBColor(255, 255, 255))
        _shade(cell, "0D3B66")
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = tbl.rows[ri + 1].cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            _set_run(p.add_run(val), val, size=10, bold=ci == 0, color=NAVY)
            if ri % 2 == 1:
                _shade(cell, "F4F7FA")
    if col_w:
        for row in tbl.rows:
            for i, w in enumerate(col_w):
                row.cells[i].width = Cm(w)
    doc.add_paragraph("")
    return tbl


def chart_barras_ahorro(path: Path, filas: List[Tuple[str, float, str]]) -> None:
    labels = [a for a, _, _ in filas]
    vals = [b for _, b, _ in filas]
    cols = ["#2E7D32" if t == "logrado" else "#C9A227" for _, _, t in filas]
    fig, ax = plt.subplots(figsize=(9.2, 3.6), dpi=150)
    y = np.arange(len(labels))
    ax.barh(y, vals, color=cols, zorder=3)
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=9)
    ax.set_xlabel("m³ / mes", fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.xaxis.grid(True, linestyle=":", alpha=0.5, zorder=0)
    ax.set_axisbelow(True)
    xmax = max(vals + [1.0]) * 1.18
    ax.set_xlim(0, xmax)
    for yi, v in zip(y, vals):
        ax.text(v + xmax * 0.015, yi, fn(v, 0), va="center", fontsize=8, fontweight="bold", color="#0D3B66")
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def chart_antes_despues(path: Path, titulo: str, pre: float, post: float, etq_pre: str, etq_post: str) -> None:
    fig, ax = plt.subplots(figsize=(6.4, 3.1), dpi=150)
    x = np.arange(2)
    ax.bar(x, [pre, post], color=["#8FA4B8", "#C9A227"], zorder=3, width=0.55)
    ax.set_xticks(x)
    ax.set_xticklabels([etq_pre, etq_post], fontsize=9)
    ax.set_ylabel("m³", fontsize=9)
    ax.set_title(titulo, fontsize=10, loc="left", color="#0D3B66")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.yaxis.grid(True, linestyle=":", alpha=0.5, zorder=0)
    ymax = max(pre, post, 1.0) * 1.28
    ax.set_ylim(0, ymax)
    for xi, v in zip(x, [pre, post]):
        ax.text(xi, v + ymax * 0.03, fn(v, 1), ha="center", fontsize=9, fontweight="bold", color="#0D3B66")
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def _word_a_pdf(docx_path: Path) -> Path | None:
    pdf = docx_path.with_suffix(".pdf")
    for bin_name in ("soffice", "libreoffice"):
        try:
            subprocess.run(
                [bin_name, "--headless", "--convert-to", "pdf", "--outdir", str(docx_path.parent), str(docx_path)],
                check=True,
                capture_output=True,
                timeout=120,
            )
            if pdf.is_file():
                return pdf
        except (FileNotFoundError, subprocess.CalledProcessError, subprocess.TimeoutExpired):
            continue
    return None


def build_doc(ctx: Dict[str, Any], hasta: date) -> Path:
    CHARTS.mkdir(parents=True, exist_ok=True)
    sur = ctx["sur"]
    norte = ctx["norte"]
    pizza = ctx["pizza"]
    bom = ctx["bom"]
    maq = ctx["maq"]
    pak = ctx["pak"]
    pak_nom = ctx["pak_nom"]

    # Pizza Hut ya está en residual nocturno: el control se cita como evidencia, no como $.
    logrado_mes = sur["ahorro_mes"] + norte["ahorro_mes"] + bom["ahorro_mes"]
    propuesto_mes = maq["ahorro_mes"] + pak["ahorro_mes"]
    total_mes = logrado_mes + propuesto_mes
    logrado_acum = sur["ahorro_acum"] + norte["ahorro_acum"] + bom["ahorro_acum"]

    filas_chart = [
        (f"MAE Estanque Sur\n(presostatos 10/06)", sur["ahorro_mes"], "logrado"),
        (f"MAE Estanque Norte\n(control 05/08)", norte["ahorro_mes"], "logrado"),
        (f"Buenaventura SI500\n(control 17/07)", bom["ahorro_mes"], "logrado"),
        (f"Quilicura Matriz\n(propuesta 00–06)", maq["ahorro_mes"], "propuesto"),
        (f"Kennedy {pak_nom}\n(propuesta 00–06)", pak["ahorro_mes"], "propuesto"),
    ]
    p_bar = CHARTS / "ahorro_mensual_barras.png"
    chart_barras_ahorro(p_bar, filas_chart)
    p_sur = CHARTS / "mae_sur_antes_despues.png"
    chart_antes_despues(
        p_sur,
        "MAE Estanque Sur — m³/día (mediana)",
        sur["pre"],
        sur["post"],
        "Antes 10/06",
        "Después 11/06",
    )
    p_bom = CHARTS / "bom_noche_antes_despues.png"
    chart_antes_despues(
        p_bom,
        "Buenaventura SI500 — m³/noche 00–06 (mediana)",
        bom["pre"],
        bom["post"],
        "Antes 17/07",
        "Con control",
    )

    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.6)
    sec.left_margin = Cm(1.8)
    sec.right_margin = Cm(1.8)
    add_logo_to_header(doc)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_run(t.add_run("WES  ·  Parque Arauco"), "WES  ·  Parque Arauco", size=12, bold=True, color=GOLD)
    h = doc.add_paragraph()
    _set_run(
        h.add_run("Informe de ahorro y propuestas de control"),
        "Informe de ahorro y propuestas de control",
        size=22,
        bold=True,
        color=NAVY,
    )
    _p(
        doc,
        f"Documento para reunión. Período de datos 01/05/2026 – {hasta:%d/%m/%Y}. "
        f"Tarifa de referencia ${fn(TARIFA_CLP_M3, 0)}/m³ (la misma del PPT de 7 malls). "
        "Lo logrado es control o reparación ya operativa. Lo propuesto es on/off 00:00–06:00 "
        "aún no implementado, con proyección conservadora (noche típica menos un residual de ~2 m³).",
        size=11,
    )

    _h(doc, "1. Lo que hay que defender en la reunión", 1)
    _p(
        doc,
        "Tres mensajes, en este orden: (1) WES ya bajó consumo donde se intervino — Estación y "
        "Buenaventura son evidencia, no promesa. (2) Lo mismo se puede copiar en Quilicura y Kennedy, "
        "que son los dos recintos con mayor noche sin control. (3) Maipú no se vende como ahorro de "
        "esta lámina: Placa y Falabella son alimentación alternativa del mall (cuando se inyecta una, "
        "se corta la otra).",
    )

    _h(doc, "2. Resumen ejecutivo (m³/mes y $)", 1)
    _tabla(
        doc,
        ["Recinto / acción", "Estado", "Base", "Ahorro m³/mes", "$/mes", "Acumulado a la fecha"],
        [
            [
                "MAE Estanque Sur — presostatos 10/06",
                "Logrado",
                f"{fn(sur['pre'], 0)} → {fn(sur['post'], 0)} m³/día",
                fn(sur["ahorro_mes"], 0),
                clp(sur["ahorro_mes"]),
                f"{fn(sur['ahorro_acum'], 0)} m³ ({clp(sur['ahorro_acum'])})",
            ],
            [
                "MAE Estanque Norte — on/off 05/08 (00–05)",
                "Logrado",
                f"noche {fn(norte['pre'], 1)} → {fn(norte['post'], 1)} m³",
                fn(norte["ahorro_mes"], 0),
                clp(norte["ahorro_mes"]),
                f"{fn(norte['ahorro_acum'], 0)} m³ ({clp(norte['ahorro_acum'])})",
            ],
            [
                "MAE Pizza Hut — on/off 01/07",
                "Operativo (no suma $)",
                f"noche ya residual {fn(pizza['pre'], 1)} → {fn(pizza['post'], 1)} m³",
                "—",
                "—",
                "Control puesto; el día alto es ocupación de local",
            ],
            [
                "Buenaventura SI500 — on/off 17/07",
                "Logrado",
                f"noche {fn(bom['pre'], 1)} → {fn(bom['post'], 1)} m³",
                fn(bom["ahorro_mes"], 0),
                clp(bom["ahorro_mes"]),
                f"{fn(bom['ahorro_acum'], 0)} m³ ({clp(bom['ahorro_acum'])})",
            ],
            [
                "Quilicura Matriz — on/off 00–06",
                "Propuesta",
                f"noche típica {fn(maq['noche'], 1)} m³ (residual {fn(maq['residual'], 0)})",
                fn(maq["ahorro_mes"], 0),
                clp(maq["ahorro_mes"]),
                "Aún no operativo",
            ],
            [
                f"Kennedy {pak_nom} — on/off 00–06",
                "Propuesta",
                f"noche típica {fn(pak['noche'], 1)} m³ (residual {fn(pak['residual'], 0)})",
                fn(pak["ahorro_mes"], 0),
                clp(pak["ahorro_mes"]),
                "Aún no operativo",
            ],
            [
                "Total logrado",
                "Operativo",
                "MAE + Buenaventura",
                fn(logrado_mes, 0),
                clp(logrado_mes),
                f"{fn(logrado_acum, 0)} m³ ({clp(logrado_acum)})",
            ],
            [
                "Total si se aprueban MAQ + PAK",
                "Logrado + propuesto",
                "Cuatro recintos",
                fn(total_mes, 0),
                clp(total_mes),
                "Proyección a 30 días de operación",
            ],
        ],
    )
    if p_bar.is_file():
        doc.add_picture(str(p_bar), width=Inches(6.3))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(
            cap.add_run("Verde = ya operativo. Dorado = propuesta de control (proyección conservadora)."),
            "Verde = ya operativo. Dorado = propuesta de control (proyección conservadora).",
            size=9,
            color=GRAY,
        )

    _h(doc, "3. Mall Arauco Estación (MAE) — ahorro ya logrado", 1)
    _p(
        doc,
        f"Estanque Sur es el caso más limpio para la reunión: el 10/06 se repararon los presostatos. "
        f"La mediana diaria pasó de {fn(sur['pre'], 0)} a {fn(sur['post'], 0)} m³/día "
        f"(−{fn(sur['baja_dia'], 0)} m³/día). A 30 días son {fn(sur['ahorro_mes'], 0)} m³/mes "
        f"({clp(sur['ahorro_mes'])}). Acumulado desde el 11/06 hasta {hasta:%d/%m}: "
        f"{fn(sur['ahorro_acum'], 0)} m³ ({clp(sur['ahorro_acum'])}). "
        "No es un modelo: es el mall después de la reparación.",
    )
    if p_sur.is_file():
        doc.add_picture(str(p_sur), width=Inches(5.6))
    _p(
        doc,
        f"Encima de eso hay dos controles nocturnos operativos. Estanque Norte (desde el 05/08, "
        f"00:00–05:00): noche típica {fn(norte['pre'], 1)} → {fn(norte['post'], 1)} m³, "
        f"{fn(norte['ahorro_noche'], 1)} m³/noche = {fn(norte['ahorro_mes'], 0)} m³/mes "
        f"({clp(norte['ahorro_mes'])}). Es un ahorro menor frente a Sur; se cita para mostrar que "
        "el on/off también corre en MAE, no para inflar el total. "
        f"Pizza Hut (desde el 01/07) ya tenía la madrugada en residual "
        f"({fn(pizza['pre'], 1)} → {fn(pizza['post'], 1)} m³/noche): el control está puesto, "
        "pero no hay m³ extra que vender. Un día alto (p.ej. 28/08) es ocupación de local, no fuga.",
    )
    _p(
        doc,
        f"Umbrales 24 h a dejar instalados en MAE: Norte {fn(UMBRAL_MAE_NORTE_DIA, 0)} · "
        f"Sur {fn(UMBRAL_MAE_SUR_DIA, 0)} · Pizza Hut {fn(UMBRAL_MAE_PIZZA_DIA, 0)} · "
        f"Baños Públicos {fn(UMBRAL_MAE_BANOS_DIA, 0)} m³/día.",
        bold=True,
    )

    _h(doc, "4. Buenaventura / San Ignacio (BOM) — ahorro ya logrado", 1)
    _p(
        doc,
        f"San Ignacio 500 tiene on/off nocturno operativo desde el 17/07. La madrugada 00–06 pasó de "
        f"{fn(bom['pre'], 1)} a {fn(bom['post'], 1)} m³ (mediana). Ahorro "
        f"{fn(bom['ahorro_noche'], 1)} m³/noche = {fn(bom['ahorro_mes'], 0)} m³/mes "
        f"({clp(bom['ahorro_mes'])}). Acumulado 17/07–{hasta:%d/%m} "
        f"({int(bom['n_post'])} noches): {fn(bom['ahorro_acum'], 0)} m³ ({clp(bom['ahorro_acum'])}). "
        f"El residual típico ({fn(bom['post'], 1)} m³) no se lee como fuga: es lo que queda con el control puesto. "
        "San Ignacio 300 queda en monitoreo; no es el punto del corte.",
    )
    if p_bom.is_file():
        doc.add_picture(str(p_bom), width=Inches(5.6))
    _p(
        doc,
        f"Umbrales 24 h: San Ignacio 500 {fn(UMBRAL_SI500_DIA, 0)} m³/día · "
        f"San Ignacio 300 {fn(UMBRAL_SI300_DIA, 0)} m³/día.",
        bold=True,
    )

    _h(doc, "5. Mall Arauco Quilicura (MAQ) — propuesta de control", 1)
    _p(
        doc,
        f"La Matriz Principal concentra el recinto. Desde el 22/06 el día se duplicó y se sostuvo; "
        f"Alimentación Baños es uso hábil, no es el problema. La noche 00–06, desde esa alza, "
        f"anda en {fn(maq['noche'], 1)} m³ (mediana, {int(maq['n'])} noches hasta {hasta:%d/%m}).",
    )
    _p(
        doc,
        f"Propuesta: on/off 00:00–06:00 en Matriz Principal, igual que SI500 y Norte. "
        f"Proyección conservadora: {fn(maq['noche'], 1)} − {fn(maq['residual'], 0)} m³ de residual = "
        f"{fn(maq['ahorro_noche'], 1)} m³/noche × 30 = {fn(maq['ahorro_mes'], 0)} m³/mes "
        f"({clp(maq['ahorro_mes'])}). Si el corte fuera a cero, el techo sería "
        f"{fn(maq['ahorro_bruto_mes'], 0)} m³/mes ({clp(maq['ahorro_bruto_mes'])}); "
        "en la reunión se defiende el número conservador.",
        bold=True,
    )
    _p(
        doc,
        f"Umbral 24 h a activar: Matriz Principal {fn(UMBRAL_MAQ_DIA, 0)} m³/día.",
    )

    _h(doc, "6. Parque Arauco Kennedy (PAK) — propuesta de control", 1)
    _p(
        doc,
        "La cadena no se suma a la cabecera del mall: Sandía Antigua y Sandía Nueva alimentan "
        "Distrito de Lujo; de ahí sale a Bazar Gourmet y DL Kennedy. Cortar de noche en el eslabón "
        "que más gasta es el control, no apagar las Sandías (dejarían sin agua toda la cadena).",
    )
    _p(
        doc,
        f"El eslabón de mayor noche 00–06 es {pak_nom}: {fn(pak['noche'], 1)} m³/noche "
        f"(el otro queda más abajo). Propuesta: on/off 00:00–06:00 en {pak_nom}. "
        f"Proyección conservadora: {fn(pak['ahorro_noche'], 1)} m³/noche × 30 = "
        f"{fn(pak['ahorro_mes'], 0)} m³/mes ({clp(pak['ahorro_mes'])}). "
        f"Techo si el corte fuera a cero: {fn(pak['ahorro_bruto_mes'], 0)} m³/mes "
        f"({clp(pak['ahorro_bruto_mes'])}).",
        bold=True,
    )
    _p(
        doc,
        f"Umbrales 24 h: Distrito de Lujo {fn(UMBRAL_PAK_DL_DIA, 0)} · "
        f"Bazar Gourmet {fn(UMBRAL_PAK_BAZAR_DIA, 0)} · DL Kennedy {fn(UMBRAL_PAK_KEN_DIA, 0)} m³/día.",
    )

    _h(doc, "7. Pedido concreto de la reunión", 1)
    _p(
        doc,
        f"Se pide aprobar on/off 00:00–06:00 en (a) Matriz Principal de Quilicura y "
        f"(b) {pak_nom} de Kennedy, con el mismo criterio que ya opera en San Ignacio 500. "
        f"Con eso se proyectan {fn(propuesto_mes, 0)} m³/mes adicionales ({clp(propuesto_mes)}). "
        "Junto con el corte, activar los umbrales 24 h de este informe. "
        "No se pide intervención de ahorro en Maipú ni en las Sandías de Kennedy.",
        bold=True,
    )

    _h(doc, "8. Cómo responder en la reunión", 1)
    bullets = [
        f"¿Cuánto ahorramos ya? {fn(logrado_mes, 0)} m³/mes ({clp(logrado_mes)}), MAE + Buenaventura. "
        f"Casi todo es Estanque Sur + SI500. Acumulado a {hasta:%d/%m}: "
        f"{fn(logrado_acum, 0)} m³ ({clp(logrado_acum)}).",
        f"¿Cuánto más si aprueban Quilicura y Kennedy? {fn(propuesto_mes, 0)} m³/mes "
        f"({clp(propuesto_mes)}), con residual de ~2 m³. Suma total {fn(total_mes, 0)} m³/mes "
        f"({clp(total_mes)}).",
        "¿Por qué Quilicura? La Matriz concentra el mall y la noche se quedó alta desde junio. El control es el mismo que ya corre en SI500.",
        f"¿Por qué Kennedy? Hay que cortar en {pak_nom} (el de mayor noche de la cadena DL), no en las Sandías.",
        "¿Pizza Hut cuánto ahorra? El control está puesto; la noche ya era residual. No se vende en $.",
        "¿Y Maipú el 08/09? Placa 4,5 m³ y Falabella 92 m³ el mismo día es cambio de alimentación, no dos fallas. No se presenta como ahorro.",
        "Los umbrales 24 h (promedio operativo × 1,25) se piden junto con el on/off, para que el recinto vea el día completo, no solo la madrugada.",
    ]
    for b in bullets:
        p = doc.add_paragraph(style="List Bullet")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_run(p.add_run(b), b, size=11, color=NAVY)

    _h(doc, "9. Notas de método", 1)
    _p(
        doc,
        "Mediana, no promedio: un sábado alto no infla el ahorro. Noche = suma 00:00–06:00 hora Chile. "
        "Ahorro mensual = m³/noche (o m³/día en Sur) × 30. Acumulado = ese delta × noches/días con la "
        "medida ya operativa. Propuestas: se resta un residual de 2 m³ (lo visto en SI500 con control) "
        "para no vender el corte a cero. Fuente: API WES, mismos puntos del consolidado 7 malls.",
        size=10,
        color=GRAY,
    )

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"Informe_ahorro_PA_reunion_{hasta.strftime('%Y%m%d')}.docx"
    doc.save(str(path))
    print(f"[OK] Word {path}", flush=True)
    return path


def main() -> int:
    parser = argparse.ArgumentParser(description="Informe de ahorro PA para reunión")
    parser.add_argument("--hasta", default="08/09/2026")
    parser.add_argument("--skip-refresh", action="store_true")
    args = parser.parse_args()
    hasta = datetime.strptime(args.hasta, "%d/%m/%Y").date()
    desde = date(2026, 5, 1)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if args.skip_refresh and JSON_DAILY.is_file():
        by = cargar_diario()
    else:
        nodos = [
            "000025-01",
            "000025-04",
            "000025-07",
            "000025-19",
            "000025-13",
            "000025-17",
            "000025-18",
            "000025-27",
            "000025-35",
            "000025-36",
        ]
        by = refrescar_diario(nodos, desde, hasta)

    night_nodes = ["000025-01", "000025-07", "000025-18", "000025-13", "000025-27", "000025-35", "000025-36"]
    night_from = date(2026, 6, 1)
    if args.skip_refresh and JSON_NIGHT.is_file():
        hourly = json.loads(JSON_NIGHT.read_text(encoding="utf-8")).get("hourly") or {}
        # si faltan noches largas, igual refresca
        need = False
        for nid in night_nodes:
            if night_from.isoformat() not in (hourly.get(nid) or {}):
                need = True
                break
        if need:
            print("[INFO] Noches de junio no están en cache; se descargan.", flush=True)
            hourly = refrescar_noches(night_nodes, night_from, hasta)
    else:
        hourly = refrescar_noches(night_nodes, night_from, hasta)

    sur = _stats_sur((by.get("000025-19") or {}).get("daily") or {}, hasta)
    norte = _stats_control(_n06_serie(hourly, "000025-01"), CTRL_NORTE, hasta, lookback_dias=14)
    pizza = _stats_control(
        _n06_serie(hourly, "000025-07"),
        CTRL_PIZZA,
        hasta,
        excluir=PIZZA_NOCHES_ATIPICAS,
        lookback_dias=14,
    )
    bom = _stats_control(_n06_serie(hourly, SI500), CTRL_SI500, hasta, lookback_dias=7)
    maq = _stats_propuesta_noche(_n06_serie(hourly, MATRIZ_MAQ), MAQ_ALZA, hasta)
    bazar = _stats_propuesta_noche(_n06_serie(hourly, BAZAR), date(2026, 7, 1), hasta)
    ken = _stats_propuesta_noche(_n06_serie(hourly, DL_KENNEDY), date(2026, 7, 1), hasta)
    if bazar["noche"] >= ken["noche"]:
        pak, pak_nom = bazar, "Bazar Gourmet"
    else:
        pak, pak_nom = ken, "DL Kennedy"

    ctx = {
        "sur": sur,
        "norte": norte,
        "pizza": pizza,
        "bom": bom,
        "maq": maq,
        "pak": pak,
        "pak_nom": pak_nom,
    }
    print(
        "[INFO] MAE Sur",
        fn(sur["ahorro_mes"], 0),
        "Norte",
        fn(norte["ahorro_mes"], 0),
        "Pizza",
        fn(pizza["ahorro_mes"], 0),
        "BOM",
        fn(bom["ahorro_mes"], 0),
        "MAQ",
        fn(maq["ahorro_mes"], 0),
        "PAK",
        pak_nom,
        fn(pak["ahorro_mes"], 0),
        flush=True,
    )
    docx = build_doc(ctx, hasta)
    pdf = _word_a_pdf(docx)
    if pdf:
        print(f"[OK] PDF {pdf}", flush=True)
    else:
        print("[AVISO] No hay LibreOffice; se entrega Word (Drive lo abre).", flush=True)
    print("\n=== SALIDA ===")
    print(docx)
    if pdf:
        print(pdf)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
