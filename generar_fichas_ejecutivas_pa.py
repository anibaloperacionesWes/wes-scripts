# -*- coding: utf-8 -*-
"""
Fichas ejecutivas Parque Arauco — una por recinto (mall), 5 variables:

  1) Equipos instalados (puntos activos WES)
  2) Consumo mensualizado (junio / julio / agosto a la fecha + proyección)
  3) Hallazgos / conclusiones  (controles nocturnos DESCARTADOS como fuga)
  4) Solicitudes / mensajes al recinto
  5) Controles nocturnos — estado (fuera del análisis de fugas)

Período: 01/06/2026 – 14/08/2026.
Fuente de datos: API WES. Narrativa alineada al PPT 7 Malls
(reports/.../entrega_diego_anibal) y a la gestión operativa de controles.

Uso:
  python generar_fichas_ejecutivas_pa.py
"""

from __future__ import annotations

import json
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PptRGB
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn as ppt_qn
from pptx.util import Emu, Inches as PptInches, Pt as PptPt

from generar_reporte_word import format_number_chilean

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "reports" / "Parque_Arauco" / "TMP_7MALLS" / "entrega_diego_anibal"
CHARTS = OUT_DIR / "charts_fichas"
JSON_DATOS = OUT_DIR / "datos_fichas_jun_ago.json"
LOGO = ROOT / "logo wes.bmp"
FONDO = ROOT / "Parque arauco fondo.jpg"

DESDE = date(2026, 6, 1)
HASTA = date(2026, 8, 14)
AGO_DIAS = 14
AGO_MES = 31
PERIODO = "01/06/2026 – 14/08/2026"
FECHA_EMISION = "14 agosto 2026"

NAVY = (13, 59, 102)
GOLD = (201, 162, 39)
TEAL = (31, 119, 180)
GREEN = (39, 124, 91)
RED = (153, 45, 48)
GRAY = (90, 90, 90)
LIGHT = (245, 247, 250)
WHITE = (255, 255, 255)

# Puntos activos por mall (WES + PPT 7 Malls). Inactivos/relocalizados fuera.
MALLS: List[Dict[str, Any]] = [
    {
        "code": "MAE",
        "nombre": "Estación",
        "titulo": "Mall Arauco Estación",
        "nodes": ["000025-01", "000025-04", "000025-07", "000025-19"],
        "extra_nodes": ["000025-02"],
        "pendiente": [],
        "recepcion": "20/10/2025",
        "capacitacion": "18/02/2025",
        "usuarios": [
            "Equipo medioambiente.dcl@parauco.com",
            "Sala de monitores MAE (salademonitoresmae@gmail.com)",
            "Sergio Fuenzalida — Analista Gestión Ambiental",
        ],
    },
    {
        "code": "MAM",
        "nombre": "Maipú",
        "titulo": "Mall Arauco Maipú",
        "nodes": ["000025-08", "000025-10", "000025-32", "000025-33"],
        "extra_nodes": [],
        "pendiente": ["000025-09"],
        "recepcion": "06/11/2025",
        "capacitacion": "14/11/2025",
        "usuarios": [
            "Miguel Rupayan — Encargado de Operaciones",
            "Constanza Vilches — Analista Ambiental",
            "Equipo Mantención: C. Bustamante, O. Cuevas y Supervisor Eléctrico",
        ],
    },
    {
        "code": "MAQ",
        "nombre": "Quilicura",
        "titulo": "Mall Arauco Quilicura",
        "nodes": ["000025-13", "000025-34"],
        "extra_nodes": [],
        "pendiente": [],
        "recepcion": "06/11/2025 (original) / 17/02/2026 (relocalizado)",
        "capacitacion": "14/11/2025",
        "usuarios": [
            "Mario Freitez — Jefe de Operaciones",
            "Tomás Saba — Center Manager",
            "Sebastián Araneda — Analista Ambiental",
            "Equipo Mantención: Iván Dustan, Katihuska Varas, Lucas Méndez, Carlos Leyto",
        ],
    },
    {
        "code": "BOM",
        "nombre": "Buenaventura",
        "titulo": "Buenaventura (San Ignacio)",
        "nodes": ["000025-17", "000025-18"],
        "extra_nodes": [],
        "pendiente": [],
        "recepcion": "18/11/2025",
        "capacitacion": "11/12/2025",
        "usuarios": [
            "Aliro Cortés — Jefe de Operaciones",
            "Tomás Saba — Center Manager",
            "Sebastián Araneda — Analista Ambiental",
        ],
    },
    {
        "code": "AEB",
        "nombre": "El Bosque",
        "titulo": "Arauco El Bosque",
        "nodes": ["000025-11", "000025-12"],
        "extra_nodes": [],
        "pendiente": [],
        "recepcion": "29/10/2025 (original) / 16/01/2026 (relocalizado)",
        "capacitacion": "20/11/2025",
        "usuarios": [
            "Tamara Martínez — Jefa de Operaciones",
            "Tomás Saba — Center Manager",
            "Sebastián Araneda — Analista Ambiental",
        ],
    },
    {
        "code": "CUR",
        "nombre": "Curauma",
        "titulo": "Arauco Curauma",
        "nodes": ["000025-37", "000025-38"],
        "extra_nodes": [],
        "pendiente": [],
        "recepcion": "29/04/2026 (Anillo Norte / Anillo Sur)",
        "capacitacion": "12/12/2025",
        "usuarios": [
            "Joceline Lazo — Jefe de Operaciones",
            "Constanza Vilches — Analista Ambiental",
        ],
    },
    {
        "code": "PAK",
        "nombre": "Kennedy",
        "titulo": "Parque Arauco Kennedy",
        "nodes": [
            "000025-20",
            "000025-21",
            "000025-22",
            "000025-23",
            "000025-24",
            "000025-27",
            "000025-28",
            "000025-29",
            "000025-35",
            "000025-36",
        ],
        "extra_nodes": [],
        "pendiente": [],
        "cabecera": [
            "000025-20",
            "000025-21",
            "000025-22",
            "000025-23",
            "000025-24",
            "000025-28",
            "000025-29",
        ],
        "recepcion": "12/12/2025",
        "capacitacion": "17/12/2025",
        "usuarios": [
            "Francisco Jeldres — Jefe de Operaciones",
            "Paula Azolas — Analista Gestión Ambiental",
            "Equipo Mantención (C. Naranjo, M. Jara, R. Moreno, R. Díaz, J. Gutiérrez, H. Fierro)",
        ],
    },
]


def fn(v: float, dec: int = 1) -> str:
    return format_number_chilean(float(v), dec)


def cargar_datos() -> Tuple[Dict[str, str], Dict[str, Dict[str, float]]]:
    raw = json.loads(JSON_DATOS.read_text(encoding="utf-8"))
    names: Dict[str, str] = {}
    by: Dict[str, Dict[str, float]] = {}
    for ns in raw["nodes_summary"]:
        nid = ns["node_id"]
        names[nid] = ns["node_name"]
        row = {"jun": 0.0, "jul": 0.0, "ago": 0.0, "total": 0.0, "dias": 0.0}
        days = set()
        for m in ns["measures"]:
            d = m["date"][:10]
            month = d[5:7]
            v = float(m["total_m3"])
            if month == "06":
                row["jun"] += v
            elif month == "07":
                row["jul"] += v
            elif month == "08":
                row["ago"] += v
            row["total"] += v
            days.add(d)
        row["dias"] = float(len(days))
        by[nid] = row
        names[nid] = ns["node_name"]
    return names, by


def sum_mes(by: Dict[str, Dict[str, float]], nids: List[str]) -> Dict[str, float]:
    out = {"jun": 0.0, "jul": 0.0, "ago": 0.0, "total": 0.0}
    for nid in nids:
        r = by.get(nid) or {}
        for k in out:
            out[k] += float(r.get(k) or 0)
    out["ago_d"] = out["ago"] / AGO_DIAS if AGO_DIAS else 0.0
    out["jun_d"] = out["jun"] / 30.0
    out["jul_d"] = out["jul"] / 31.0
    out["ago_proy"] = out["ago_d"] * AGO_MES
    return out


def chart_mensual(path: Path, jun: float, jul: float, ago: float, proy: float) -> None:
    fig, ax = plt.subplots(figsize=(5.4, 2.15), dpi=140)
    labels = ["Junio", "Julio", "Agosto\n(1–14)", "Ago. proy.\n(31 d)"]
    vals = [jun, jul, ago, proy]
    colors = ["#1F77B4", "#0D3B66", "#C9A227", "#8FA4B8"]
    bars = ax.bar(labels, vals, color=colors, width=0.62, zorder=3)
    ax.set_ylabel("m³", fontsize=8, color="#0D3B66")
    ax.tick_params(labelsize=7.5, colors="#0D3B66")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.yaxis.grid(True, linestyle=":", alpha=0.5, zorder=0)
    ax.set_axisbelow(True)
    ymax = max(vals) * 1.22 if max(vals) > 0 else 1
    ax.set_ylim(0, ymax)
    for b, v in zip(bars, vals):
        ax.text(
            b.get_x() + b.get_width() / 2,
            b.get_height() + ymax * 0.02,
            fn(v, 0),
            ha="center",
            va="bottom",
            fontsize=7,
            color="#0D3B66",
            fontweight="bold",
        )
    fig.tight_layout(pad=0.25)
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def _equipos_lineas(mall: Dict[str, Any], names: Dict[str, str], by: Dict[str, Dict[str, float]]) -> List[str]:
    lines: List[str] = []
    for nid in mall["nodes"]:
        nm = names.get(nid, nid)
        lines.append(f"• {nm} ({nid})")
    for nid in mall.get("extra_nodes") or []:
        nm = names.get(nid, nid)
        lines.append(f"• {nm} ({nid}) — activo, fuera del deck 7 Malls")
    for nid in mall.get("pendiente") or []:
        nm = names.get(nid, nid)
        dias = int((by.get(nid) or {}).get("dias") or 0)
        if dias:
            lines.append(f"• {nm} ({nid}) — pendiente OC; {dias} día(s) con dato en agosto")
        else:
            lines.append(f"• {nm} ({nid}) — a la espera de OC / cambio de equipo")
    return lines


def _consumo_lineas(mall: Dict[str, Any], tot: Dict[str, float], by: Dict[str, Dict[str, float]], names: Dict[str, str]) -> List[str]:
    lines = [
        f"Junio: {fn(tot['jun'])} m³  ({fn(tot['jun_d'])} m³/día)",
        f"Julio: {fn(tot['jul'])} m³  ({fn(tot['jul_d'])} m³/día)",
        f"Agosto 1–14: {fn(tot['ago'])} m³  ({fn(tot['ago_d'])} m³/día)",
        f"Proyección agosto: {fn(tot['ago_proy'])} m³",
    ]
    ranked = sorted(mall["nodes"], key=lambda n: -float((by.get(n) or {}).get("total") or 0))
    if mall["code"] == "PAK":
        lines.append("Cifras de cabecera (20+21+22+23+24+28+29). DL/Bazar/Kennedy no se suman.")
        top = "000025-22"
        t = by[top]
        lines.append(
            f"Mayor cabecera: {names.get(top, top)} — "
            f"jun {fn(t['jun'])} / jul {fn(t['jul'])} / ago {fn(t['ago'])} m³"
        )
    else:
        top = ranked[0]
        t = by[top]
        lines.append(
            f"Mayor aporte: {names.get(top, top)} — "
            f"jun {fn(t['jun'])} / jul {fn(t['jul'])} / ago {fn(t['ago'])} m³"
        )
    var_jj = ((tot["jul_d"] / tot["jun_d"]) - 1) * 100 if tot["jun_d"] else 0
    var_aj = ((tot["ago_d"] / tot["jul_d"]) - 1) * 100 if tot["jul_d"] else 0
    lines.append(f"Variación m³/día: jul vs jun {fn(var_jj, 0)}%  ·  ago vs jul {fn(var_aj, 0)}%")
    if mall["code"] == "MAE":
        lines.append("Suma de los 4 puntos del deck (sin 000025-02).")
    return lines


def contenidos(names: Dict[str, str], by: Dict[str, Dict[str, float]]) -> List[Dict[str, Any]]:
    """Arma las 5 variables por mall con cifras reales y controles descartados."""
    out: List[Dict[str, Any]] = []
    for mall in MALLS:
        nids_consumo = mall.get("cabecera") or mall["nodes"]
        tot = sum_mes(by, nids_consumo)
        extra = sum_mes(by, mall["nodes"] + (mall.get("extra_nodes") or []))
        equipos = _equipos_lineas(mall, names, by)
        consumo = _consumo_lineas(mall, tot, by, names)
        if mall["code"] == "MAE":
            hallazgos = [
                "Reparación 10/06 en red Sur (validada con mantención): Estanque Sur pasó de 83,1 a 29,4 m³/día y se mantiene (~27–29 m³/día jul–ago). No es fuga residual.",
                "Pizza Hut: control nocturno desde 01/07 funciona (noche ~0 m³ el 10/08). El alza de julio (23,8 → 37,7 m³/día) es diurna; agosto baja a 14,1 m³/día. Noche descartada.",
                "Estanque Norte: control desde 05/08. La noche no explica el alza (ago 1–4: 41 m³/día; 5–14: 50 m³/día). Revisar locales mall en horario hábil.",
                "Baños Públicos: 16,3 → 7,0 → 3,5 m³/día. Noches ya ~0. Control instalado sin funcionamiento: no es prioridad.",
            ]
            # 000025-02 queda en equipos; no saturar hallazgos.
            solicitudes = [
                "Confirmar que el corte on/off de Estanque Sur queda a cargo permanente de mantención nocturna.",
                "No abrir orden de fuga por noches de Norte / Pizza Hut / Sur: controles activos o corte operativo.",
                "Revisar alza diurna de Estanque Norte en agosto (locales mall).",
                "Incorporar Abastecimiento Sur Terminal (000025-02) al tablero del recinto si corresponde a la cuenta del mall.",
            ]
            controles = [
                "Estanque Sur (000025-19): corte on/off a cargo de personal de mantención nocturno — DESCARTADO.",
                "Pizza Hut (000025-07): control nocturno desde 01/07/2026 (00:00–06:00) — DESCARTADO. Verificado 10/08: 0 m³ en 0–6 h.",
                "Estanque Norte (000025-01): control nocturno desde 05/08/2026 — DESCARTADO.",
                "Baños Públicos: control instalado sin funcionamiento; noches ya ~0 — no se interpreta como fuga.",
            ]
        elif mall["code"] == "MAM":
            hallazgos = [
                f"Placa Bancaria concentra ~{fn(by['000025-08']['total']/tot['total']*100,0)}% del volumen monitoreado.",
                "Auditoría 18/06: Placa subió de 160,5 a 347,2 m³/día. Julio 228,5 m³/día. Agosto 1–14 vuelve a 137,0 m³/día (bajo el nivel pre-auditoría). Noche 10/08 = 0 m³.",
                "Ripley: noches tendiendo a cero (lámina 9 del deck); volumen jul 3.245 m³, estable en agosto (~97 m³/día).",
                "Pasillo Técnico Boulevard y salida ARROW: consumo residual (69,8 y 1,1 m³ en el período).",
                "Impulsión Falabella: sin dato jun–jul (OC pendiente). Aparecen 4 días desde el 11/08 (24,6 / 55,3 / 66,1 / 12,0 m³) — posible rehabilitación, no consolidada.",
            ]
            solicitudes = [
                "Confirmar si la OC / cambio de equipo de Impulsión Falabella ya se ejecutó (hay registros desde el 11/08).",
                "Mantener seguimiento de Placa Bancaria: el alza del 18/06 se revirtió en agosto; no reabrir como fuga nocturna.",
                "Pasillo y ARROW: dejar como referencia de red (no priorizar).",
            ]
            controles = [
                "Sin control nocturno WES declarado en este recinto.",
                "Ripley y Placa: patrón nocturno del deck ya era ~0; se descarta fuga de madrugada.",
            ]
        elif mall["code"] == "MAQ":
            hallazgos = [
                f"Matriz Principal = {fn(by['000025-13']['total']/tot['total']*100,1)}% del recinto. Baños = {fn(by['000025-34']['total']/tot['total']*100,1)}%.",
                "Alza clara: 131 m³/día en junio → 193 m³/día en julio y 195 m³/día en agosto (1–14).",
                "10/08 Matriz: 23,7 m³ en 0–6 h (de 198 m³ del día). Sigue el patrón del deck (21,7 m³/noche; sin noches en cero).",
                "Alimentación Baños: bajo y estable (jun 103 / jul 62 / ago 44 m³); uso hábil.",
                "Red de Incendio (000025-14) relocalizada: 0 m³. No forma parte del activo.",
            ]
            solicitudes = [
                "Implementar control on/off 00:00–08:00 en Matriz Principal (igual que estanques MAE).",
                "Oportunidad de orden de magnitud: ~24 m³/noche × 30 ≈ 720 m³/mes si se corta el caudal inhábil.",
                "No hay control nocturno que descartar: este es el hallazgo principal del recinto.",
            ]
            controles = [
                "No hay control nocturno activo en Quilicura.",
                "El consumo de madrugada de Matriz Principal NO se descarta: es la variable a gestionar.",
            ]
        elif mall["code"] == "BOM":
            hallazgos = [
                "San Ignacio 500 = ~80% del recinto. Alza desde el 26/06 (36,7 → 99,6 m³/día) y 1–15/07 en 148,7 m³/día.",
                "Control nocturno 500 desde 16/07 FUNCIONA: 15/07 noche 42,4 m³ → 20/07 3,3 m³ → 10/08 2,3 m³. Noche DESCARTADA como fuga.",
                "El volumen diurno no volvió a la base de junio: 16–31/07 114,8 m³/día y 1–14/08 111,2 m³/día vs 36,7 m³/día (1–25/06). Queda alza operacional de día.",
                "San Ignacio 300 (solo monitoreo): 11,5 → 33,8 → 38,8 m³/día. El 10/08, 11,6 m³ de 22,6 m³ fueron en 0–6 h (~51% nocturno). Sin control.",
            ]
            solicitudes = [
                "500: no reabrir fuga nocturna. Pedir a operaciones la causa del caudal diurno/vespertino que quedó alto desde el 26/06.",
                "300: revisar consumo de madrugada sostenido e implementar control on/off (no está cubierto).",
                "Mensaje al JO (Aliro Cortés): el control de 500 ya aporta; el ahorro grande ahora está en el día y en el 300.",
            ]
            controles = [
                "San Ignacio 500 (000025-18): control nocturno desde 16/07/2026 — DESCARTADO (verificado: noche 42 → 2 m³).",
                "San Ignacio 300: sin control. No se descarta.",
            ]
        elif mall["code"] == "AEB":
            hallazgos = [
                "Puntos activos: Matriz principal 1° piso (000025-11) y Anillo Plaza (000025-12). Matriz A.A. (000025-30) = 0 m³ en el período (no operativo).",
                f"Matriz 11: jun {fn(by['000025-11']['jun'])} / jul {fn(by['000025-11']['jul'])} / ago {fn(by['000025-11']['ago'])} m³ (~55 m³/día, estable).",
                "10/08 Matriz: 7,6 m³ en 0–6 h (base inhábil del deck: 7,9 m³/noche). Sin control.",
                f"Anillo Plaza: jul {fn(by['000025-12']['jul'])} m³ (alza vs jun {fn(by['000025-12']['jun'])}); 10/08 noche 2,7 de 12,2 m³.",
                "La portada del deck 7 Malls aún nombra Matriz A.A.; el dato activo es 000025-11.",
            ]
            solicitudes = [
                "Replicar control on/off 00:00–08:00 en Matriz (como estanques MAE). Oportunidad ~8 m³/noche.",
                "Revisar llaves / equipos del Anillo Plaza para bajar la base nocturna.",
                "Actualizar el listado del recinto: activo 11+12; 30 en cero.",
            ]
            controles = [
                "No hay control nocturno activo en El Bosque.",
                "La noche de Matriz y Anillo NO se descarta: es la oportunidad de ahorro del recinto.",
            ]
        elif mall["code"] == "CUR":
            hallazgos = [
                "Red reconfigurada: Anillo Norte (000025-38) y Anillo Sur (000025-37). Matriz/Baños 15–16 = 0 m³ (no comparables).",
                f"Volumen conjunto estable: jun {fn(tot['jun'])} / jul {fn(tot['jul'])} / ago {fn(tot['ago'])} m³ (~20–22 m³/día).",
                "Ambos anillos aportan en partes similares (Sur levemente mayor).",
                "En el deck se cuadró WES vs boleta Esval N° 2368516 (18/05–16/06); la diferencia se explicó por desfase de lecturas, no por pérdida.",
            ]
            solicitudes = [
                "Enviar la(s) boleta(s) Esval de junio–agosto para repetir el cuadre WES vs cuenta.",
                "No hay control nocturno que activar como prioridad: el recinto está estable y acotado.",
                "Mantener a Joceline Lazo / medioambiente.dcl en el tablero de los dos anillos.",
            ]
            controles = [
                "Sin control nocturno WES en Curauma.",
                "No se declara fuga nocturna en este recinto para el período.",
            ]
        else:  # PAK
            hallazgos = [
                "10 puntos activos. Baños 5/6 (25–26) relocalizados a Bazar Gourmet (35) y DL Kennedy (36).",
                f"Cadena: Sandía Antigua (22) y Sandía Nueva (28) alimentan DL (27), que se reparte en Bazar (35) y DL Kennedy (36). No sumar la cadena para facturar.",
                f"Consumo de cabecera (sin doble conteo): jun {fn(tot['jun'])} / jul {fn(tot['jul'])} / ago {fn(tot['ago'])} m³.",
                "10/08 DL: 47,5 m³ en 0–6 h (de 221 m³ del día). El deck ya marcaba al DL como el mayor nocturno de la cadena.",
                f"Andén Locales Gastronómicos (21) sube: jun {fn(by['000025-21']['jun'])} → jul {fn(by['000025-21']['jul'])} m³. Piletas: volumen menor.",
            ]
            solicitudes = [
                "Evaluar control on/off 00:00–08:00 en Distrito de Lujo y salas de bomba Sandía (mismo criterio MAE).",
                "No usar la suma de los 10 puntos como consumo del mall: hay doble conteo de la cadena DL.",
                "Mensaje a JO (Francisco Jeldres) / mantención: la noche del DL es la variable de mayor impacto.",
            ]
            controles = [
                "Sin control nocturno WES activo en Kennedy.",
                "El patrón 0–8 h del DL NO se descarta: es el hallazgo a gestionar.",
            ]

        out.append(
            {
                "mall": mall,
                "tot": tot,
                "extra": extra,
                "equipos": equipos,
                "consumo": consumo,
                "hallazgos": hallazgos,
                "solicitudes": solicitudes,
                "controles": controles,
            }
        )
    return out


# ---------------------------------------------------------------------------
# PPT
# ---------------------------------------------------------------------------
def _rgb(t: Tuple[int, int, int]) -> PptRGB:
    return PptRGB(*t)


def _set_run(run, text: str, size: int, bold: bool = False, color=NAVY) -> None:
    run.text = text
    run.font.size = PptPt(size)
    run.font.bold = bold
    run.font.color.rgb = _rgb(color)
    run.font.name = "Calibri"


def _caja(slide, l, t, w, h, fill=LIGHT, line=None):
    sh = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        PptInches(l),
        PptInches(t),
        PptInches(w),
        PptInches(h),
    )
    sh.fill.solid()
    sh.fill.fore_color.rgb = _rgb(fill)
    sh.line.fill.background()
    if line:
        sh.line.fill.solid()
        sh.line.color.rgb = _rgb(line)
        sh.line.width = Emu(6350)
    try:
        sh.adjustments[0] = 0.08
    except Exception:
        pass
    return sh


def _tb(slide, l, t, w, h, lines: List[Tuple[str, int, bool, Tuple[int, int, int]]], align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(PptInches(l), PptInches(t), PptInches(w), PptInches(h))
    tf = box.text_frame
    tf.word_wrap = True
    for i, (text, size, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = PptPt(2)
        p.level = 0
        run = p.add_run()
        _set_run(run, text, size, bold, color)
    return box


def _header_bar(slide, prs, titulo: str, sub: str) -> None:
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, PptInches(0.92)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = _rgb(NAVY)
    bar.line.fill.background()
    gold = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, PptInches(0.92), prs.slide_width, PptInches(0.06)
    )
    gold.fill.solid()
    gold.fill.fore_color.rgb = _rgb(GOLD)
    gold.line.fill.background()
    if LOGO.is_file():
        slide.shapes.add_picture(str(LOGO), PptInches(12.15), PptInches(0.18), height=PptInches(0.58))
    _tb(slide, 0.28, 0.12, 11.6, 0.42, [(titulo, 20, True, WHITE)])
    _tb(slide, 0.28, 0.50, 11.6, 0.36, [(sub, 11, False, (220, 230, 240))])


def _card_title(slide, l, t, w, titulo: str, color=TEAL) -> None:
    _tb(slide, l, t, w, 0.28, [(titulo.upper(), 11, True, color)])


def build_ppt(fichas: List[Dict[str, Any]], names: Dict[str, str], by: Dict[str, Dict[str, float]]) -> Path:
    prs = Presentation()
    prs.slide_width = PptInches(13.333)
    prs.slide_height = PptInches(7.5)

    # Portada
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    if FONDO.is_file():
        pic = sl.shapes.add_picture(str(FONDO), 0, 0, width=prs.slide_width, height=prs.slide_height)
        spTree = sl.shapes._spTree
        spTree.remove(pic.element)
        spTree.insert(2, pic.element)
    veil = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, PptInches(3.55), prs.slide_width, PptInches(3.95))
    veil.fill.solid()
    veil.fill.fore_color.rgb = _rgb((13, 59, 102))
    try:
        veil.fill.fore_color.brightness = 0  # type: ignore
    except Exception:
        pass
    # transparencia aproximada no crítica
    veil.line.fill.background()
    _tb(sl, 0.6, 3.75, 12, 0.5, [("WES  ·  Parque Arauco", 16, True, GOLD)])
    _tb(sl, 0.6, 4.20, 12, 0.7, [("Fichas ejecutivas por recinto", 32, True, WHITE)])
    _tb(
        sl,
        0.6,
        4.95,
        12,
        1.2,
        [
            (f"Período {PERIODO}   |   Emisión {FECHA_EMISION}", 16, False, WHITE),
            ("5 variables por mall  ·  controles nocturnos descartados del análisis de fugas", 14, False, (220, 230, 240)),
            ("MAE · MAM · MAQ · BOM · AEB · CUR · PAK   |   Puntos activos WES", 14, False, GOLD),
        ],
    )

    # Metodología
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _header_bar(
        sl, prs,
        "Cómo leer las fichas",
        "Misma lógica del PPT 7 Malls (entrega Diego / Aníbal), actualizado a agosto",
    )
    bloques = [
        (
            "1. Equipos instalados",
            "Puntos WES activos del recinto. Se excluyen relocalizados en cero (Red de Incendio, Matriz/Baños CUR antiguos, Baños 5/6 PAK, KFC/Poniente 7/Locales comida). Falabella se declara pendiente de OC.",
        ),
        (
            "2. Consumo mensualizado",
            "Suma de puntos WES del recinto: junio, julio y agosto 1–14, más proyección lineal a 31 días. En Kennedy no se suma la cadena DL (doble conteo).",
        ),
        (
            "3. Hallazgos",
            "Tendencia jun–ago y puntos de mayor aporte. La noche de un punto CON control no se lee como fuga.",
        ),
        (
            "4. Solicitudes / mensajes",
            "Lo que queremos pasar al JO / mantención / medioambiente del mall.",
        ),
        (
            "5. Controles nocturnos (descartados)",
            "San Ignacio 500 desde 16/07 · Pizza Hut desde 01/07 · Estanque Norte desde 05/08 · Estanque Sur: corte on/off de mantención nocturna. Verificados con perfil 0–6 h.",
        ),
    ]
    y = 1.15
    for tit, txt in bloques:
        _caja(sl, 0.35, y, 12.6, 1.05)
        _tb(sl, 0.5, y + 0.08, 12.3, 0.28, [(tit, 13, True, NAVY)])
        _tb(sl, 0.5, y + 0.38, 12.3, 0.58, [(txt, 12, False, GRAY)])
        y += 1.12

    # Una ficha por mall
    for item in fichas:
        mall = item["mall"]
        tot = item["tot"]
        sl = prs.slides.add_slide(prs.slide_layouts[6])
        n_eq = len(mall["nodes"])
        extra_n = len(mall.get("extra_nodes") or [])
        pend_n = len(mall.get("pendiente") or [])
        pts = f"{n_eq} puntos activos"
        if extra_n:
            pts += f" + {extra_n} adicional"
        if pend_n:
            pts += f" + {pend_n} pendiente OC"
        _header_bar(
            sl, prs,
            f"{mall['code']}  ·  {mall['titulo']}",
            f"{pts}   |   {PERIODO}   |   Recepción {mall['recepcion']}",
        )

        chart_path = CHARTS / f"mensual_{mall['code']}.png"
        chart_mensual(chart_path, tot["jun"], tot["jul"], tot["ago"], tot["ago_proy"])

        # 5 cards
        _caja(sl, 0.25, 1.12, 4.35, 2.55)
        _card_title(sl, 0.38, 1.18, 4.1, "1. Equipos instalados")
        eq_lines = [(ln, 10 if len(item["equipos"]) > 8 else 11, False, NAVY) for ln in item["equipos"]]
        _tb(sl, 0.38, 1.48, 4.1, 2.1, eq_lines)

        _caja(sl, 4.70, 1.12, 8.35, 2.55)
        _card_title(sl, 4.85, 1.18, 4.0, "2. Consumo mensualizado (puntos WES)")
        cons_lines = [(ln, 11, False, NAVY) for ln in item["consumo"]]
        _tb(sl, 4.85, 1.48, 3.95, 2.1, cons_lines)
        sl.shapes.add_picture(str(chart_path), PptInches(8.90), PptInches(1.45), width=PptInches(3.95))

        _caja(sl, 0.25, 3.78, 6.40, 2.05)
        _card_title(sl, 0.38, 3.84, 6.1, "3. Hallazgos / conclusiones", RED)
        hall = [(f"• {h}" if not h.startswith("•") else h, 10, False, NAVY) for h in item["hallazgos"][:5]]
        _tb(sl, 0.38, 4.14, 6.12, 1.62, hall)

        _caja(sl, 6.80, 3.78, 6.25, 2.05)
        _card_title(sl, 6.93, 3.84, 6.0, "4. Solicitudes / mensajes", GREEN)
        sol = [(f"• {h}", 10, False, NAVY) for h in item["solicitudes"][:4]]
        _tb(sl, 6.93, 4.14, 5.98, 1.62, sol)

        _caja(sl, 0.25, 5.95, 12.80, 1.38, fill=(255, 249, 235), line=GOLD)
        _card_title(sl, 0.38, 6.00, 12.5, "5. Controles nocturnos — descartados del análisis de fugas", GOLD)
        ctrl = [(f"• {h}", 11, False, NAVY) for h in item["controles"]]
        _tb(sl, 0.38, 6.28, 12.5, 0.98, ctrl)

    # Consolidado
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _header_bar(sl, prs, "Consolidado  ·  mensaje por recinto", f"Puntos WES  |  {PERIODO}")
    rows = [["Mall", "Puntos", "Junio m³", "Julio m³", "Ago 1–14 m³", "m³/día ago", "Mensaje"]]
    mensajes_corto = {
        "MAE": "Noches Norte/Pizza/Sur descartadas. Norte: alza diurna ago.",
        "MAM": "Placa se revirtió en ago. Confirmar OC Falabella (dato 11/08).",
        "MAQ": "Matriz ~24 m³/noche — pedir control on/off.",
        "BOM": "500 noche OK desde 16/07; queda alza diurna + noche del 300.",
        "AEB": "Activo 11+12. Pedir control inhábil en Matriz.",
        "CUR": "Estable. Pedir boletas Esval jun–ago.",
        "PAK": "Cifras de cabecera (sin DL). Noche DL ~47 m³ (10/08) — control.",
    }
    for item in fichas:
        m = item["mall"]
        t = item["tot"]
        n = len(m["nodes"]) + len(m.get("extra_nodes") or [])
        rows.append(
            [
                f"{m['code']} {m['nombre']}",
                str(n),
                fn(t["jun"], 0),
                fn(t["jul"], 0),
                fn(t["ago"], 0),
                fn(t["ago_d"], 1),
                mensajes_corto[m["code"]],
            ]
        )

    table_shape = sl.shapes.add_table(len(rows), 7, PptInches(0.28), PptInches(1.20), PptInches(12.75), PptInches(5.6))
    table = table_shape.table
    widths = [1.9, 0.85, 1.35, 1.35, 1.5, 1.2, 4.6]
    for i, w in enumerate(widths):
        table.columns[i].width = PptInches(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = val
            for p in cell.text_frame.paragraphs:
                p.font.size = PptPt(10 if r else 11)
                p.font.bold = r == 0 or c == 0
                p.font.name = "Calibri"
                p.font.color.rgb = _rgb(WHITE if r == 0 else NAVY)
            fill = cell.fill
            fill.solid()
            fill.fore_color.rgb = _rgb(NAVY if r == 0 else (LIGHT if r % 2 else WHITE))

    _tb(
        sl,
        0.28,
        6.95,
        12.7,
        0.4,
        [("Fuente: API WES. Controles informados por operación (16/07 SI500, 01/07 Pizza Hut, 05/08 Estanque Norte, corte Sur por mantención nocturna).", 10, False, GRAY)],
    )

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f"Fichas_ejecutivas_Parque_Arauco_{HASTA.strftime('%Y%m%d')}.pptx"
    prs.save(str(path))
    print(f"[OK] PPT {path}")
    return path


# ---------------------------------------------------------------------------
# Word
# ---------------------------------------------------------------------------
def _shade(cell, hex_color: str) -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def _set_cell(cell, text: str, *, bold=False, color=NAVY, size=10, fill=None, center=False) -> None:
    if fill:
        _shade(cell, fill)
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(*color)
    run.font.name = "Calibri"


def _add_bullets(doc: Document, items: List[str]) -> None:
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.space_before = Pt(0)
        run = p.add_run(it)
        run.font.size = Pt(11)
        run.font.name = "Calibri"
        run.font.color.rgb = RGBColor(*NAVY)


def build_word(fichas: List[Dict[str, Any]]) -> Path:
    doc = Document()
    for s in doc.sections:
        s.top_margin = Cm(1.6)
        s.bottom_margin = Cm(1.6)
        s.left_margin = Cm(1.8)
        s.right_margin = Cm(1.8)

    t = doc.add_paragraph()
    r = t.add_run("WES  ·  Parque Arauco")
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(*GOLD)
    r.font.name = "Calibri"

    h = doc.add_paragraph()
    r = h.add_run("Fichas ejecutivas por recinto")
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = RGBColor(*NAVY)
    r.font.name = "Calibri"

    p = doc.add_paragraph()
    r = p.add_run(
        f"Período {PERIODO}. Emisión {FECHA_EMISION}. "
        "Cinco variables por mall: equipos instalados, consumo mensualizado de los puntos WES, "
        "hallazgos, solicitudes, y controles nocturnos (descartados como fuga). "
        "Alineado al PPT 7 Malls de la carpeta entrega_diego_anibal, con datos junio–agosto."
    )
    r.font.size = Pt(11)
    r.font.name = "Calibri"

    p = doc.add_paragraph()
    r = p.add_run("Controles nocturnos descartados en todo el informe: ")
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(*NAVY)
    r = p.add_run(
        "San Ignacio 500 desde el 16/07/2026; Pizza Hut desde el 01/07/2026; "
        "Estanque Norte desde el 05/08/2026; Estanque Sur con corte on/off a cargo de "
        "personal de mantención nocturno. Esas madrugadas no se interpretan como fuga."
    )
    r.font.size = Pt(11)

    for item in fichas:
        mall = item["mall"]
        tot = item["tot"]
        doc.add_page_break()
        h = doc.add_paragraph()
        r = h.add_run(f"{mall['code']}  —  {mall['titulo']}")
        r.bold = True
        r.font.size = Pt(18)
        r.font.color.rgb = RGBColor(*NAVY)

        meta = doc.add_paragraph()
        r = meta.add_run(
            f"Recepción {mall['recepcion']}  ·  Capacitación {mall['capacitacion']}  ·  "
            f"Usuarios: {'; '.join(mall['usuarios'])}"
        )
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(*GRAY)
        r.italic = True

        for titulo, lineas in (
            ("1. Equipos instalados", item["equipos"]),
            ("2. Consumo mensualizado del recinto (puntos WES)", item["consumo"]),
        ):
            hh = doc.add_paragraph()
            rr = hh.add_run(titulo)
            rr.bold = True
            rr.font.size = Pt(13)
            rr.font.color.rgb = RGBColor(*TEAL)
            _add_bullets(doc, lineas)

        # tabla mensual
        table = doc.add_table(rows=2, cols=5)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr = ["Junio", "Julio", "Agosto 1–14", "m³/día ago", "Proyección agosto"]
        vals = [
            f"{fn(tot['jun'])} m³",
            f"{fn(tot['jul'])} m³",
            f"{fn(tot['ago'])} m³",
            f"{fn(tot['ago_d'])} m³/d",
            f"{fn(tot['ago_proy'])} m³",
        ]
        for i, x in enumerate(hdr):
            _set_cell(table.rows[0].cells[i], x, bold=True, color=WHITE, fill="0D3B66", center=True, size=10)
        for i, x in enumerate(vals):
            _set_cell(table.rows[1].cells[i], x, bold=True, fill="F5F7FA", center=True, size=11)

        chart = CHARTS / f"mensual_{mall['code']}.png"
        if chart.is_file():
            doc.add_paragraph().add_run().add_picture(str(chart), width=Inches(5.8))

        for titulo, lineas, col in (
            ("3. Hallazgos / conclusiones", item["hallazgos"], TEAL),
            ("4. Solicitudes / mensajes a pasar", item["solicitudes"], GREEN),
            ("5. Controles nocturnos (descartados del análisis de fugas)", item["controles"], GOLD),
        ):
            hh = doc.add_paragraph()
            rr = hh.add_run(titulo)
            rr.bold = True
            rr.font.size = Pt(13)
            rr.font.color.rgb = RGBColor(*col)
            _add_bullets(doc, lineas)

    # consolidado
    doc.add_page_break()
    h = doc.add_paragraph()
    r = h.add_run("Consolidado")
    r.bold = True
    r.font.size = Pt(18)
    r.font.color.rgb = RGBColor(*NAVY)
    table = doc.add_table(rows=1 + len(fichas), cols=6)
    table.style = "Table Grid"
    for i, x in enumerate(["Mall", "Junio m³", "Julio m³", "Ago 1–14 m³", "m³/día ago", "Mensaje clave"]):
        _set_cell(table.rows[0].cells[i], x, bold=True, color=WHITE, fill="0D3B66", center=True, size=9)
    msgs = {
        "MAE": "Descartar noches Norte/Pizza/Sur. Norte: alza diurna agosto.",
        "MAM": "Placa se revirtió. Confirmar OC Falabella (dato desde 11/08).",
        "MAQ": "Pedir control nocturno en Matriz (~24 m³/noche).",
        "BOM": "500 noche OK; queda alza diurna y noche del 300.",
        "AEB": "Activos 11 y 12. Pedir control inhábil en Matriz.",
        "CUR": "Estable. Pedir boletas Esval junio–agosto.",
        "PAK": "Cifras de cabecera (sin DL). Control nocturno en DL.",
    }
    for r_i, item in enumerate(fichas, 1):
        m = item["mall"]
        t = item["tot"]
        vals = [
            f"{m['code']} {m['nombre']}",
            fn(t["jun"], 0),
            fn(t["jul"], 0),
            fn(t["ago"], 0),
            fn(t["ago_d"], 1),
            msgs[m["code"]],
        ]
        fill = "F5F7FA" if r_i % 2 else "FFFFFF"
        for c, x in enumerate(vals):
            _set_cell(table.rows[r_i].cells[c], x, fill=fill, size=9, bold=c == 0)

    p = doc.add_paragraph()
    r = p.add_run(
        "Fuente: API de medidas WES, 01/06/2026–14/08/2026. "
        "La proyección de agosto es lineal (14 días → 31). "
        "Los controles nocturnos se verificaron con perfil horario 0–6 h en fechas de corte."
    )
    r.font.size = Pt(9)
    r.italic = True
    r.font.color.rgb = RGBColor(*GRAY)

    path = OUT_DIR / f"Fichas_ejecutivas_Parque_Arauco_{HASTA.strftime('%Y%m%d')}.docx"
    doc.save(str(path))
    print(f"[OK] Word {path}")
    return path


def convertir_pdf(docx_path: Path) -> Path | None:
    import shutil
    import subprocess

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("[AVISO] LibreOffice no está disponible; se omite PDF.")
        return None
    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        str(docx_path.parent),
        str(docx_path),
    ]
    print(f"[INFO] PDF via {soffice}", flush=True)
    subprocess.run(cmd, check=True, timeout=180)
    pdf = docx_path.with_suffix(".pdf")
    if pdf.is_file():
        print(f"[OK] PDF {pdf}")
        return pdf
    return None


def main() -> int:
    if not JSON_DATOS.is_file():
        print(f"[ERROR] Falta {JSON_DATOS}. Ejecutar primero la descarga de medidas.")
        return 1
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CHARTS.mkdir(parents=True, exist_ok=True)
    names, by = cargar_datos()
    fichas = contenidos(names, by)
    ppt = build_ppt(fichas, names, by)
    docx = build_word(fichas)
    pdf = convertir_pdf(docx)
    print("\n=== SALIDA ===")
    print(ppt)
    print(docx)
    if pdf:
        print(pdf)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
