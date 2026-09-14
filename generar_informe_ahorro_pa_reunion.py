# -*- coding: utf-8 -*-
"""
Informe de ahorro Parque Arauco para reunión — estilo reporte agregado.

Destaca lo demostrable:
  - MAE: Estanque Sur (presostatos) + Estanque Norte (01:00–05:00 → cero)
  - Buenaventura: SI500 01:00–05:00 pasó a cero (corte 00:30)
  - Quilicura / Kennedy / El Bosque 1° piso / Falabella Maipú: misma ventana, proyección a cero
  - Maipú: on/off en Falabella + relocalizar Arrow e instalar el punto 6 (Pasillo 2)

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
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

from generar_consolidado_pa_ppt import (
    JSON_DAILY,
    JSON_HOURS,
    JSON_NIGHT,
    OUT_DIR,
    cargar_diario,
    refrescar_diario,
    refrescar_horas_noche,
    refrescar_noches,
)
from generar_ppt_recorrido_ejecutivo_pa import (
    BAZAR,
    CTRL_NORTE,
    CTRL_PIZZA,
    CTRL_SI500,
    DL_KENNEDY,
    FALABELLA,
    MAQ_ALZA,
    MATRIZ_AEB,
    PIZZA_NOCHES_ATIPICAS,
    SUR_REPARACION,
    TARIFA_CLP_M3,
    UMBRAL_FALABELLA_DIA,
    UMBRAL_MAE_BANOS_DIA,
    UMBRAL_MAE_NORTE_DIA,
    UMBRAL_MAE_PIZZA_DIA,
    UMBRAL_MAE_SUR_DIA,
    UMBRAL_MAQ_DIA,
    UMBRAL_MATRIZ_AEB_DIA,
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
# Matriz 1° piso activa desde 15/05; la noche típica se toma de julio en adelante.
AEB_PROP_DESDE = date(2026, 7, 1)
# Desde el 15/08 el mall de Maipú sale por Falabella (reactivado 11/08).
FALABELLA_PROP_DESDE = date(2026, 8, 15)
CHARTS = OUT_DIR / "charts_informe_reunion"
DIAGRAMA_MAM = OUT_DIR / "mam_propuesta_localizacion.png"
# Corte a las 00:30: 01:00–05:00 demuestra el cero; el $ es la noche 00:30–06:00 (00–06 menos 00:00–00:30).
H_INI = 1  # 01:00
H_FIN = 5  # 05:00 (demostración a cero)
H_NOCHE_FIN = 6  # 00:00–06:00 para el monto


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


def _hval(rec: Dict[str, Any], h: int) -> float:
    if not rec:
        return 0.0
    return float(rec.get(str(h), rec.get(h, 0)) or 0)


def _serie_ventana(
    by_h: Dict[str, Dict[str, Dict[str, float]]],
    nid: str,
    h0: int = H_INI,
    h1: int = H_FIN,
) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for iso, rec in (by_h.get(nid) or {}).items():
        out[iso] = sum(_hval(rec, h) for h in range(h0, h1))
    return out


def _serie_hora(by_h: Dict[str, Dict[str, Dict[str, float]]], nid: str, h: int = 0) -> Dict[str, float]:
    return {iso: _hval(rec, h) for iso, rec in (by_h.get(nid) or {}).items()}


def _stats_control(
    serie: Dict[str, float],
    ctrl: date,
    hasta: date,
    *,
    excluir: set | None = None,
    lookback_dias: int | None = 14,
    serie_h0: Dict[str, float] | None = None,
) -> Dict[str, float]:
    """Mediana 01:00–05:00 antes vs con control. Meta = cero (no se resta residual)."""
    pre, post, h0_post = [], [], []
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
            if serie_h0 is not None:
                h0_post.append(float(serie_h0.get(iso, 0.0) or 0.0))
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
        "h0_post": _mediana(h0_post) if h0_post else 0.0,
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


def _stats_logrado(
    by_h: Dict[str, Dict[str, Dict[str, float]]],
    nid: str,
    ctrl: date,
    hasta: date,
    *,
    lookback_dias: int,
    excluir: set | None = None,
) -> Dict[str, float]:
    """$ = noche 00:00–06:00 menos el tramo 00:00–00:30. 01:00–05:00 prueba el cero."""
    n06 = _stats_control(
        _serie_ventana(by_h, nid, 0, H_NOCHE_FIN),
        ctrl,
        hasta,
        excluir=excluir,
        lookback_dias=lookback_dias,
        serie_h0=_serie_hora(by_h, nid, 0),
    )
    cero = _stats_control(
        _serie_ventana(by_h, nid, H_INI, H_FIN),
        ctrl,
        hasta,
        excluir=excluir,
        lookback_dias=lookback_dias,
    )
    leftover = n06.get("h0_post") or n06["post"]
    ahor = max(0.0, n06["pre"] - leftover)
    n06["cero_pre"] = cero["pre"]
    n06["cero_post"] = cero["post"]
    n06["ahorro_noche"] = ahor
    n06["ahorro_mes"] = ahor * 30.0
    n06["ahorro_acum"] = ahor * n06["n_post"]
    return n06


def _stats_propuesta_noche(
    serie: Dict[str, float], d0: date, hasta: date, *, min_noche: float = 0.0
) -> Dict[str, float]:
    """Proyección a cero desde las 00:30: el ahorro es la noche típica 00:00–06:00 completa."""
    vals = []
    for iso, v in serie.items():
        d = date.fromisoformat(iso)
        if d0 <= d <= hasta and float(v) >= min_noche:
            vals.append(float(v))
    med = _mediana(vals)
    return {
        "noche": med,
        "residual": 0.0,
        "ahorro_noche": med,
        "ahorro_mes": med * 30.0,
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


def _p_lead(doc: Document, lead: str, rest: str, *, size=11) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_run(p.add_run(lead), lead, size=size, bold=True, color=NAVY)
    _set_run(p.add_run(" " + rest), " " + rest, size=size, bold=False, color=NAVY)


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
    tbl.autofit = False
    tbl.allow_autofit = False
    widths = col_w or []
    total_cm = sum(widths) if widths else 17.4
    tblPr = tbl._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(int(total_cm * 567)))
    tblW.set(qn("w:type"), "dxa")
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)

    def _cell(cell, text: str, *, header: bool, bold: bool, zebra: bool) -> None:
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        _set_run(
            p.add_run(text),
            text,
            size=8,
            bold=bold or header,
            color=RGBColor(255, 255, 255) if header else NAVY,
        )
        if header:
            _shade(cell, "0D3B66")
        elif zebra:
            _shade(cell, "F4F7FA")

    for i, h in enumerate(headers):
        _cell(tbl.rows[0].cells[i], h, header=True, bold=True, zebra=False)
    for ri, row in enumerate(rows):
        tr_pr = tbl.rows[ri + 1]._tr.get_or_add_trPr()
        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))
        last = str(row[0]).startswith("Total")
        for ci, val in enumerate(row):
            cell = tbl.rows[ri + 1].cells[ci]
            _cell(cell, val, header=False, bold=ci == 0 or last, zebra=ri % 2 == 1)
            if last:
                _shade(cell, "D9E1F2")
    if widths:
        for row in tbl.rows:
            tr_pr = row._tr.get_or_add_trPr()
            if tr_pr.find(qn("w:cantSplit")) is None:
                tr_pr.append(OxmlElement("w:cantSplit"))
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph("")
    return tbl


def chart_barras_ahorro(path: Path, filas: List[Tuple[str, float, str]]) -> None:
    labels = [a for a, _, _ in filas]
    vals = [b for _, b, _ in filas]
    cols = ["#2E7D32" if t == "logrado" else "#C9A227" for _, _, t in filas]
    fig, ax = plt.subplots(figsize=(9.2, 4.6), dpi=150)
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
    aeb = ctx["aeb"]
    fala = ctx["fala"]

    # Pizza Hut ya está en cero 01:00–05:00: el control se cita como evidencia, no como $.
    logrado_mes = sur["ahorro_mes"] + norte["ahorro_mes"] + bom["ahorro_mes"]
    propuesto_mes = maq["ahorro_mes"] + pak["ahorro_mes"] + aeb["ahorro_mes"] + fala["ahorro_mes"]
    total_mes = logrado_mes + propuesto_mes
    logrado_acum = sur["ahorro_acum"] + norte["ahorro_acum"] + bom["ahorro_acum"]

    filas_chart = [
        (f"MAE Estanque Sur\n(presostatos 10/06)", sur["ahorro_mes"], "logrado"),
        (f"MAE Estanque Norte\n(control 05/08)", norte["ahorro_mes"], "logrado"),
        (f"Buenaventura SI500\n(control 17/07)", bom["ahorro_mes"], "logrado"),
        (f"Quilicura Matriz\n(propuesta 00:30 → 0)", maq["ahorro_mes"], "propuesto"),
        (f"Kennedy {pak_nom}\n(propuesta 00:30 → 0)", pak["ahorro_mes"], "propuesto"),
        (f"El Bosque Matriz 1° piso\n(propuesta 00:30 → 0)", aeb["ahorro_mes"], "propuesto"),
        (f"Maipú Falabella\n(propuesta 00:30 → 0)", fala["ahorro_mes"], "propuesto"),
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
    p_norte = CHARTS / "mae_norte_antes_despues.png"
    chart_antes_despues(
        p_norte,
        "MAE Estanque Norte — m³/noche desde 00:30 (mediana)",
        norte["pre"],
        0.0,
        "Antes 05/08",
        "A cero",
    )
    p_bom = CHARTS / "bom_noche_antes_despues.png"
    chart_antes_despues(
        p_bom,
        "Buenaventura SI500 — m³/noche desde 00:30 (mediana)",
        bom["pre"],
        0.0,
        "Antes 17/07",
        "A cero",
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
        f"Tarifa de referencia ${fn(TARIFA_CLP_M3, 0)}/m³ (la misma del PPT de 7 malls).",
        size=11,
    )
    _p(
        doc,
        "El corte se activa a las 00:30. La noche que se demuestra —y la que tiene que ir a cero— "
        "es de 01:00 a 05:00.",
        size=11,
    )
    _p(
        doc,
        "Entre 00:00 y 00:30 el medidor todavía registra agua: es el tramo anterior al corte, "
        "no consumo de noche.",
        size=11,
    )

    _h(doc, "1. Lo que hay que demostrar en la reunión", 1)
    _p(doc, "En este orden:")
    _p_lead(
        doc,
        "Ya operativo — Estación y Buenaventura.",
        "Estanque Sur bajó el día completo (presostatos 10/06). Estanque Norte y San Ignacio 500 "
        "pasaron de consumo nocturno a cero entre 01:00 y 05:00.",
    )
    _p_lead(
        doc,
        "A copiar — Quilicura, Kennedy, El Bosque 1° piso y Falabella Maipú.",
        "El mismo corte se propone en Matriz Principal, Bazar Gourmet, Matriz 1° piso de El Bosque "
        "y Falabella de Maipú (el mall sale por ahí desde el 15/08). Esa noche hoy no es cero; "
        "esa es la proyección.",
    )
    _p_lead(
        doc,
        "Pendiente — Maipú (puntos 6 y Arrow).",
        "El on/off de Falabella sí entra en la simulación. Sigue pendiente concretar la relocalización "
        "de Pasillo 1 Arrow y la instalación del punto 6 (Pasillo 2), según la propuesta enviada a Don Miguel.",
    )

    p_br = doc.add_paragraph()
    p_br.add_run().add_break(WD_BREAK.PAGE)
    _h(doc, "2. Resumen ejecutivo (m³/mes y $)", 1)
    _tabla(
        doc,
        ["Recinto", "Estado", "Noche → cero", "m³/mes", "$/mes"],
        [
            [
                "Estanque Sur (presostatos 10/06)",
                "Logrado",
                f"{fn(sur['pre'], 0)} → {fn(sur['post'], 0)} m³/día",
                fn(sur["ahorro_mes"], 0),
                clp(sur["ahorro_mes"]),
            ],
            [
                "Estanque Norte",
                "Logrado",
                f"{fn(norte['pre'], 1)} → 0 m³ desde 00:30",
                fn(norte["ahorro_mes"], 0),
                clp(norte["ahorro_mes"]),
            ],
            [
                "Pizza Hut",
                "Operativo",
                "Ya en cero (no suma $)",
                "—",
                "—",
            ],
            [
                "Buenaventura SI500",
                "Logrado",
                f"{fn(bom['pre'], 1)} → 0 m³ desde 00:30",
                fn(bom["ahorro_mes"], 0),
                clp(bom["ahorro_mes"]),
            ],
            [
                "Quilicura Matriz",
                "Propuesta",
                f"{fn(maq['noche'], 1)} m³ → 0 desde 00:30",
                fn(maq["ahorro_mes"], 0),
                clp(maq["ahorro_mes"]),
            ],
            [
                f"Kennedy {pak_nom}",
                "Propuesta",
                f"{fn(pak['noche'], 1)} m³ → 0 desde 00:30",
                fn(pak["ahorro_mes"], 0),
                clp(pak["ahorro_mes"]),
            ],
            [
                "El Bosque Matriz 1° piso",
                "Propuesta",
                f"{fn(aeb['noche'], 1)} m³ → 0 desde 00:30",
                fn(aeb["ahorro_mes"], 0),
                clp(aeb["ahorro_mes"]),
            ],
            [
                "Maipú Falabella",
                "Propuesta",
                f"{fn(fala['noche'], 1)} m³ → 0 desde 00:30",
                fn(fala["ahorro_mes"], 0),
                clp(fala["ahorro_mes"]),
            ],
            [
                "Total logrado",
                "Operativo",
                "MAE + Buenaventura",
                fn(logrado_mes, 0),
                clp(logrado_mes),
            ],
            [
                "Total si se aprueban las 4",
                "Logrado + propuesto",
                "MAQ + PAK + AEB 1° piso + Falabella",
                fn(total_mes, 0),
                clp(total_mes),
            ],
        ],
        col_w=[5.4, 2.6, 4.2, 2.4, 2.8],
    )
    _p(
        doc,
        f"Acumulado ya operativo a {hasta:%d/%m}: {fn(logrado_acum, 0)} m³ ({clp(logrado_acum)}). "
        "El $ es la noche desde las 00:30 a cero (00:00–00:30 queda en el medidor y no se resta del ahorro).",
        size=10,
        color=GRAY,
    )
    if p_bar.is_file():
        doc.add_picture(str(p_bar), width=Inches(6.3))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(
            cap.add_run("Verde = ya operativo (noche a cero desde las 00:30). Dorado = propuesta, misma ventana a cero."),
            "Verde = ya operativo (noche a cero desde las 00:30). Dorado = propuesta, misma ventana a cero.",
            size=9,
            color=GRAY,
        )

    _h(doc, "Control nocturno ya operativo", 1)
    _p(
        doc,
        "El corte se activa a las 00:30. De 01:00 a 05:00 el consumo quedó en cero. "
        "Estos son los controles que ya corren (los mismos de la última lámina del consolidado):",
    )
    _tabla(
        doc,
        ["Punto", "Desde", "Control nocturno", "m³/mes", "$/mes"],
        [
            [
                "MAE Estanque Norte",
                "05/08",
                f"{fn(norte['pre'], 1)} → 0 m³ desde las 00:30 (01:00–05:00 en cero)",
                fn(norte["ahorro_mes"], 0),
                clp(norte["ahorro_mes"]),
            ],
            [
                "MAE Pizza Hut",
                "01/07",
                "01:00–05:00 ya en cero; el control está puesto, no suma m³ extra",
                "—",
                "—",
            ],
            [
                "BOM San Ignacio 500",
                "17/07",
                f"{fn(bom['pre'], 1)} → 0 m³ desde las 00:30. El ~2 m³ de 00:00–00:30 no es noche",
                fn(bom["ahorro_mes"], 0),
                clp(bom["ahorro_mes"]),
            ],
            [
                "MAE Estanque Sur",
                "10/06",
                f"No es on/off: presostatos. Día {fn(sur['pre'], 0)} → {fn(sur['post'], 0)} m³",
                fn(sur["ahorro_mes"], 0),
                clp(sur["ahorro_mes"]),
            ],
            [
                "Total ya operativo",
                "",
                "MAE + Buenaventura",
                fn(logrado_mes, 0),
                clp(logrado_mes),
            ],
        ],
        col_w=[4.6, 1.6, 6.2, 2.4, 2.6],
    )

    _h(doc, "A proponer — mismo corte 00:30", 1)
    _p(
        doc,
        "Copiar el on/off que ya corre en Norte y SI500. Meta: esa noche a cero.",
    )
    _tabla(
        doc,
        ["Punto", "Acción", "Noche → cero", "m³/mes", "$/mes"],
        [
            [
                "MAQ Matriz Principal",
                "On/off 00:30",
                f"{fn(maq['noche'], 1)} m³ → 0  ·  umbral 240 m³/día",
                fn(maq["ahorro_mes"], 0),
                clp(maq["ahorro_mes"]),
            ],
            [
                f"PAK {pak_nom}",
                "On/off 00:30 (no las Sandías)",
                f"{fn(pak['noche'], 1)} m³ → 0  ·  umbrales DL 390 · Bazar 250 · DL Kennedy 20",
                fn(pak["ahorro_mes"], 0),
                clp(pak["ahorro_mes"]),
            ],
            [
                "AEB Matriz 1° piso",
                "On/off 00:30",
                f"{fn(aeb['noche'], 1)} m³ → 0  ·  umbral 75 m³/día",
                fn(aeb["ahorro_mes"], 0),
                clp(aeb["ahorro_mes"]),
            ],
            [
                "MAM Falabella",
                "On/off 00:30 (desde 15/08 sale por Falabella)",
                f"{fn(fala['noche'], 1)} m³ → 0  ·  umbral 140 m³/día",
                fn(fala["ahorro_mes"], 0),
                clp(fala["ahorro_mes"]),
            ],
            [
                "Total si se aprueban las 4",
                "Logrado + propuesto",
                "MAQ + PAK + AEB 1° piso + Falabella. CUR: noche chica, sin on/off",
                fn(total_mes, 0),
                clp(total_mes),
            ],
        ],
        col_w=[4.6, 2.8, 5.0, 2.4, 2.6],
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
        f"Estanque Norte (desde el 05/08): el corte se activa a las 00:30. De 01:00 a 05:00 el "
        f"consumo quedó en cero (de {fn(norte.get('cero_pre', norte['pre']), 1)} a "
        f"{fn(norte.get('cero_post', 0), 1)} m³). La noche completa desde las 00:30 era "
        f"{fn(norte['pre'], 1)} m³: eso es {fn(norte['ahorro_noche'], 1)} m³/noche = "
        f"{fn(norte['ahorro_mes'], 0)} m³/mes ({clp(norte['ahorro_mes'])}). "
        f"Lo que aparece entre 00:00 y 00:30 ({fn(norte.get('h0_post', 0), 1)} m³) no es noche "
        "y no se resta del ahorro.",
    )
    if p_norte.is_file():
        doc.add_picture(str(p_norte), width=Inches(5.6))
    _p(
        doc,
        "Pizza Hut (desde el 01/07) ya tenía 01:00–05:00 en cero: el control está puesto, "
        "pero no hay m³ extra que sumar. Un día alto (p.ej. 28/08) es ocupación de local, no fuga.",
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
        f"San Ignacio 500 tiene corte nocturno desde el 17/07, armado a las 00:30. De 01:00 a "
        f"05:00 pasó de {fn(bom.get('cero_pre', bom['pre']), 1)} m³ a cero. La noche completa "
        f"desde las 00:30 era {fn(bom['pre'], 1)} m³ y quedó en cero: "
        f"{fn(bom['ahorro_noche'], 1)} m³/noche = {fn(bom['ahorro_mes'], 0)} m³/mes "
        f"({clp(bom['ahorro_mes'])}). Acumulado 17/07–{hasta:%d/%m} ({int(bom['n_post'])} noches): "
        f"{fn(bom['ahorro_acum'], 0)} m³ ({clp(bom['ahorro_acum'])}).",
    )
    _p(
        doc,
        f"Si alguien mira el total 00:00–06:00 va a ver ~{fn(bom.get('h0_post', 0), 1)} m³. "
        "Eso no es un residual de noche ni una fuga: es el agua que corre entre las 00:00 y las "
        "00:30, hasta que el corte queda armado. No se resta del ahorro. San Ignacio 300 queda "
        "en monitoreo; no es el punto del corte.",
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
        f"Alimentación Baños es uso hábil, no es el problema. La noche desde las 00:30, desde esa alza, "
        f"anda en {fn(maq['noche'], 1)} m³ (mediana, {int(maq['n'])} noches hasta {hasta:%d/%m}).",
    )
    _p(
        doc,
        f"Propuesta: mismo corte que SI500 y Norte (se arma a las 00:30) en Matriz Principal. "
        f"La meta es pasar {fn(maq['noche'], 1)} m³ a cero desde las 00:30 = "
        f"{fn(maq['ahorro_noche'], 1)} m³/noche × 30 = {fn(maq['ahorro_mes'], 0)} m³/mes "
        f"({clp(maq['ahorro_mes'])}). No se resta el tramo 00:00–00:30: el control es ir a cero.",
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
        f"El eslabón de mayor noche desde las 00:30 es {pak_nom}: {fn(pak['noche'], 1)} m³ "
        f"(el otro queda más abajo). Propuesta: mismo corte a las 00:30 en {pak_nom}. "
        f"Meta: {fn(pak['noche'], 1)} m³ → 0 = {fn(pak['ahorro_mes'], 0)} m³/mes "
        f"({clp(pak['ahorro_mes'])}). Tampoco se resta el tramo 00:00–00:30.",
        bold=True,
    )
    _p(
        doc,
        f"Umbrales 24 h: Distrito de Lujo {fn(UMBRAL_PAK_DL_DIA, 0)} · "
        f"Bazar Gourmet {fn(UMBRAL_PAK_BAZAR_DIA, 0)} · DL Kennedy {fn(UMBRAL_PAK_KEN_DIA, 0)} m³/día.",
    )

    _h(doc, "7. Mall Arauco El Bosque (AEB) — propuesta de control", 1)
    _p(
        doc,
        "Desde el 15/05 la Matriz A.A. está desactivada: el recinto se lee en Matriz 1° piso "
        f"({MATRIZ_AEB}) y Anillo Plaza. Esta simulación corta solo el primer piso, que es el "
        f"caudal de noche: {fn(aeb['noche'], 1)} m³ (mediana, {int(aeb['n'])} noches desde julio "
        f"hasta {hasta:%d/%m}). Anillo Plaza queda con umbral de día, sin $ en esta proyección.",
    )
    _p(
        doc,
        f"Propuesta: mismo corte a las 00:30 en Matriz 1° piso. Meta: {fn(aeb['noche'], 1)} m³ → 0 = "
        f"{fn(aeb['ahorro_mes'], 0)} m³/mes ({clp(aeb['ahorro_mes'])}). No se resta el tramo "
        "00:00–00:30: el control es ir a cero.",
        bold=True,
    )
    _p(
        doc,
        f"Umbral 24 h a activar: Matriz 1° piso {fn(UMBRAL_MATRIZ_AEB_DIA, 0)} m³/día.",
    )

    _h(doc, "8. Mall Arauco Maipú (MAM) — Falabella + pendiente Arrow / punto 6", 1)
    _p(
        doc,
        "Placa Bancaria y Falabella son alimentación alternativa del mall (cuando se inyecta una, "
        "se corta la otra). El 08/09 Placa 4,5 m³ y Falabella 92 m³ es cambio de alimentación, no "
        "dos fallas. Desde el 15/08 el mall sale por Falabella: el on/off se propone en esa línea.",
    )
    _p(
        doc,
        f"Noche típica de Falabella desde el 15/08 (solo noches con caudal, no los días que inyecta "
        f"Placa): {fn(fala['noche'], 1)} m³ (mediana, {int(fala['n'])} noches hasta {hasta:%d/%m}). "
        f"Propuesta: mismo corte a las 00:30. Meta: {fn(fala['noche'], 1)} m³ → 0 = "
        f"{fn(fala['ahorro_mes'], 0)} m³/mes ({clp(fala['ahorro_mes'])}). Tampoco se resta el tramo "
        "00:00–00:30.",
        bold=True,
    )
    _p(
        doc,
        f"Umbral 24 h a activar: Falabella {fn(UMBRAL_FALABELLA_DIA, 0)} m³/día.",
    )
    _p(
        doc,
        "Sigue pendiente —y hay que pedirlo en la reunión— concretar la relocalización "
        "del punto de Pasillo 1 Arrow y la instalación del punto 6 (Pasillo 2). Se envió a Don Miguel "
        "la propuesta de localización: dos sectores hoy no están monitoreados y pueden concentrar "
        "una parte importante del volumen distribuido. Pasillo Técnico Boulevard y Arrow casi no "
        "miden; no sirven donde están.",
    )
    _p(
        doc,
        "En la visita técnica del miércoles 19/08 se conversó que, para instalar y reubicar con "
        "seguridad, hace falta acceso al entretecho (escotilla o gatera). En paralelo WES fabrica "
        "el CIR del punto 6, que va en Pasillo 2. Falta validación de la propuesta y la habilitación "
        "de esos accesos para programar la faena.",
    )
    if DIAGRAMA_MAM.is_file():
        doc.add_picture(str(DIAGRAMA_MAM), width=Inches(6.1))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run(
            cap.add_run(
                "Propuesta MAM: relocalizar Pasillo 1 Arrow e instalar el punto 6 en Pasillo 2."
            ),
            "Propuesta MAM: relocalizar Pasillo 1 Arrow e instalar el punto 6 en Pasillo 2.",
            size=9,
            color=GRAY,
        )

    _h(doc, "9. Pedido concreto de la reunión", 1)
    _p(
        doc,
        f"Se pide aprobar el mismo corte (se arma a las 00:30; 01:00–05:00 a cero) en "
        f"(a) Matriz Principal de Quilicura, (b) {pak_nom} de Kennedy, "
        f"(c) Matriz 1° piso de El Bosque y (d) Falabella de Maipú. "
        f"Con eso se proyectan {fn(propuesto_mes, 0)} m³/mes adicionales ({clp(propuesto_mes)}). "
        "Junto con el corte, activar los umbrales 24 h de este informe.",
        bold=True,
    )
    _p(
        doc,
        "En Maipú, además del on/off de Falabella, se pide validar la propuesta enviada a Don Miguel "
        "(relocalizar Arrow e instalar el punto 6 en Pasillo 2) y habilitar la gatera al entretecho "
        "para programar la faena. No se pide corte en las Sandías de Kennedy ni on/off en Curauma.",
        bold=True,
    )

    _h(doc, "Cómo responder en la reunión", 1)
    bullets = [
        f"¿Cuánto se demuestra ya? {fn(logrado_mes, 0)} m³/mes ({clp(logrado_mes)}), MAE + Buenaventura. "
        f"Casi todo es Estanque Sur + SI500 a cero en 01:00–05:00. Acumulado a {hasta:%d/%m}: "
        f"{fn(logrado_acum, 0)} m³ ({clp(logrado_acum)}).",
        f"¿Cuánto más si aprueban las 4? {fn(propuesto_mes, 0)} m³/mes "
        f"({clp(propuesto_mes)}): Quilicura, Kennedy, El Bosque 1° piso y Falabella Maipú, "
        f"pasando esa misma ventana a cero. Suma total {fn(total_mes, 0)} m³/mes "
        f"({clp(total_mes)}).",
        "¿Por qué no restan 2 m³ de residual? Porque el control nocturno es ir a cero. Esos ~2 m³ "
        "son el tramo 00:00–00:30, antes de que el corte quede armado. No es noche.",
        "¿Por qué Quilicura? La Matriz concentra el mall y 01:00–05:00 se quedó alta desde junio. El corte es el mismo que ya corre en SI500.",
        f"¿Por qué Kennedy? Hay que cortar en {pak_nom} (el de mayor noche de la cadena DL), no en las Sandías.",
        "¿Por qué El Bosque 1° piso? La Matriz A.A. está desactivada desde el 15/05; el caudal de noche está en Matriz 1° piso.",
        "¿Pizza Hut cuánto ahorra? El control está puesto; 01:00–05:00 ya era cero. No se suma en $.",
        "¿Y Maipú? On/off en Falabella (el mall sale por ahí desde el 15/08). Placa/Falabella el 08/09 es cambio de alimentación, no dos fallas. Arrow + punto 6 sigue pendiente con Don Miguel y la gatera.",
        "Los umbrales 24 h (promedio operativo × 1,25) se piden junto con el corte, para que el recinto vea el día completo, no solo la madrugada.",
    ]
    for b in bullets:
        p = doc.add_paragraph(style="List Bullet")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_run(p.add_run(b), b, size=11, color=NAVY)

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
    hour_nodes = [
        "000025-01",
        "000025-07",
        "000025-18",
        "000025-13",
        "000025-35",
        "000025-36",
        MATRIZ_AEB,
        FALABELLA,
    ]
    night_from = date(2026, 6, 1)
    if args.skip_refresh and JSON_NIGHT.is_file():
        hourly = json.loads(JSON_NIGHT.read_text(encoding="utf-8")).get("hourly") or {}
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

    if args.skip_refresh and JSON_HOURS.is_file():
        by_h = json.loads(JSON_HOURS.read_text(encoding="utf-8")).get("by_h") or {}
        need_h = False
        check_iso = {
            MATRIZ_AEB: AEB_PROP_DESDE.isoformat(),
            FALABELLA: FALABELLA_PROP_DESDE.isoformat(),
        }
        for nid in hour_nodes:
            iso = check_iso.get(nid, night_from.isoformat())
            rec = (by_h.get(nid) or {}).get(iso) or {}
            if "1" not in rec and 1 not in rec:
                need_h = True
                break
        if need_h:
            print("[INFO] Horas 01–05 no están en cache; se descargan.", flush=True)
            by_h = refrescar_horas_noche(hour_nodes, night_from, hasta)
    else:
        by_h = refrescar_horas_noche(hour_nodes, night_from, hasta)

    sur = _stats_sur((by.get("000025-19") or {}).get("daily") or {}, hasta)
    norte = _stats_logrado(by_h, "000025-01", CTRL_NORTE, hasta, lookback_dias=14)
    pizza = _stats_logrado(
        by_h, "000025-07", CTRL_PIZZA, hasta, lookback_dias=14, excluir=PIZZA_NOCHES_ATIPICAS
    )
    bom = _stats_logrado(by_h, SI500, CTRL_SI500, hasta, lookback_dias=7)
    maq = _stats_propuesta_noche(_serie_ventana(by_h, MATRIZ_MAQ, 0, H_NOCHE_FIN), MAQ_ALZA, hasta)
    bazar = _stats_propuesta_noche(_serie_ventana(by_h, BAZAR, 0, H_NOCHE_FIN), date(2026, 7, 1), hasta)
    ken = _stats_propuesta_noche(_serie_ventana(by_h, DL_KENNEDY, 0, H_NOCHE_FIN), date(2026, 7, 1), hasta)
    aeb = _stats_propuesta_noche(
        _serie_ventana(by_h, MATRIZ_AEB, 0, H_NOCHE_FIN), AEB_PROP_DESDE, hasta
    )
    fala = _stats_propuesta_noche(
        _serie_ventana(by_h, FALABELLA, 0, H_NOCHE_FIN),
        FALABELLA_PROP_DESDE,
        hasta,
        min_noche=1.0,
    )
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
        "aeb": aeb,
        "fala": fala,
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
        "AEB 1° piso",
        fn(aeb["ahorro_mes"], 0),
        "Falabella",
        fn(fala["ahorro_mes"], 0),
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
