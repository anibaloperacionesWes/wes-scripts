"""Pizza Hut (000025-07): noches en cero entre 00:30 y 05:00 desde 2025.

Serie horaria API WES (dates.measures.csv, hora civil Chile).
Una noche cuenta si las horas 01:00, 02:00, 03:00 y 04:00 están en 0,00 m³/h
(núcleo de la ventana 00:30–05:00). El inicio y el fin del ciclo son el primer
y el último instante de la tira continua en cero que cubre esa ventana.

Uso:
  python analizar_pizza_hut_cero_nocturno.py
"""

from __future__ import annotations

import csv
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

from generar_reporte_word import _chile_hours_from_dates_measures_csv_text, acl_node_base_url
from wes_estilo_graficos_app import COLOR_BARRA_WES, COLOR_NOCHE, guardar_grafico_horario_24h_app

NODE_ID = "000025-07"
NODE_NAME = "PIZZA HUT"
COMPANY = "Parque Arauco — Estación"
# Núcleo 00:30–05:00 con resolución horaria: 01:00, 02:00, 03:00, 04:00.
VENTANA_HORAS = (1, 2, 3, 4)
EPS = 1e-9
MAX_WORKERS = 16

COLOR_WES = "1F4E79"
COLOR_INICIO = "C6EFCE"
COLOR_FIN = "FCE4D6"
COLOR_UNICO = "FFF2CC"
COLOR_ALT = "F2F2F2"

DIAS_ES = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
MESES_ES = [
    "",
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
]


@dataclass
class Noche:
    fecha: str
    tiene_datos: bool
    en_cero_ventana: bool
    tipo: str  # sin_datos | corte_nocturno | dia_completo | con_consumo
    inicio: str
    fin: str
    inicio_desde_tarde_previa: bool
    h00: float
    h01: float
    h02: float
    h03: float
    h04: float
    h05: float
    h06: float
    max_ventana: float
    consumo_diurno: float
    consumo_dia: float
    perfil_00_06: str


@dataclass
class Ciclo:
    n: int
    inicio: date
    fin: date
    dias: int
    n_corte: int
    n_completo: int
    inicio_hora_tipica: str
    fin_hora_tipica: str


def _fmt_fecha(d: date) -> str:
    return f"{d.day:02d}-{d.month:02d}-{d.year}"


def _fmt_fecha_larga(d: date) -> str:
    return f"{DIAS_ES[d.weekday()]} {_fmt_fecha(d)}"


def _hora_txt(h: int) -> str:
    if h >= 24:
        return "24:00+"
    return f"{h:02d}:00"


def _es_cero(v: float) -> bool:
    return float(v) <= EPS


def _fetch_dia(dia: date) -> Tuple[date, Dict[int, float]]:
    """CSV horario del día. Dict vacío = sin filas (no es consumo cero)."""
    url = f"{acl_node_base_url()}/nodes/{NODE_ID}/dates.measures.csv"
    ds = dia.strftime("%d%m%Y")
    try:
        r = requests.get(url, params=[("start", ds), ("end", ds)], timeout=60)
        if not r.ok:
            return dia, {}
        text = (r.text or "").strip()
        if not text or text.count("\n") < 1:
            return dia, {}
        horas = _chile_hours_from_dates_measures_csv_text(text, dia)
        out = {int(h): float(v) for h, v in horas.items()} if horas else {}
        # Hueco de 1 h entre ceros (p. ej. CSV con T01 duplicado y T02 ausente).
        for h in range(7):
            if h in out:
                continue
            left = out.get(h - 1)
            right = out.get(h + 1)
            if left is not None and right is not None and _es_cero(left) and _es_cero(right):
                out[h] = 0.0
        return dia, out
    except Exception:
        return dia, {}


def descargar_serie(desde: date, hasta: date) -> Dict[date, Dict[int, float]]:
    dias: List[date] = []
    cur = desde
    while cur <= hasta:
        dias.append(cur)
        cur += timedelta(days=1)
    out: Dict[date, Dict[int, float]] = {}
    print(f"[INFO] Descargando {len(dias)} días de {NODE_ID} ({desde} → {hasta})...")
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futs = [ex.submit(_fetch_dia, d) for d in dias]
        done = 0
        for fut in as_completed(futs):
            dia, horas = fut.result()
            out[dia] = horas
            done += 1
            if done % 50 == 0 or done == len(dias):
                print(f"  [{done}/{len(dias)}] ...", flush=True)
    return out


def analizar_noche(
    dia: date,
    horas: Dict[int, float],
    h23_prev: Optional[float],
) -> Noche:
    ventana_presente = all(h in horas for h in VENTANA_HORAS)
    if not horas or not ventana_presente:
        return Noche(
            fecha=dia.isoformat(),
            tiene_datos=bool(horas),
            en_cero_ventana=False,
            tipo="sin_datos",
            inicio="",
            fin="",
            inicio_desde_tarde_previa=False,
            h00=round(float(horas.get(0, 0.0)), 3) if horas else 0,
            h01=round(float(horas.get(1, 0.0)), 3) if horas else 0,
            h02=round(float(horas.get(2, 0.0)), 3) if horas else 0,
            h03=round(float(horas.get(3, 0.0)), 3) if horas else 0,
            h04=round(float(horas.get(4, 0.0)), 3) if horas else 0,
            h05=round(float(horas.get(5, 0.0)), 3) if horas else 0,
            h06=round(float(horas.get(6, 0.0)), 3) if horas else 0,
            max_ventana=0,
            consumo_diurno=0,
            consumo_dia=round(sum(horas.values()), 3) if horas else 0,
            perfil_00_06="",
        )

    vals_win = [float(horas[h]) for h in VENTANA_HORAS]
    en_cero = all(_es_cero(v) for v in vals_win)
    diurno = sum(float(horas.get(h, 0.0)) for h in range(7, 24))
    total = sum(float(horas.get(h, 0.0)) for h in range(24))
    max_win = max(vals_win) if vals_win else 0.0

    inicio = ""
    fin = ""
    desde_prev = False
    if en_cero:
        start_h = 1
        while start_h > 0 and _es_cero(float(horas.get(start_h - 1, 1.0))):
            start_h -= 1
        desde_prev = start_h == 0 and h23_prev is not None and _es_cero(h23_prev)
        end_h = 4
        while end_h < 23 and _es_cero(float(horas.get(end_h + 1, 1.0))):
            end_h += 1
        inicio = "23:00 (día anterior)" if desde_prev else _hora_txt(start_h)
        # Fin = primera hora con consumo tras la tira en cero.
        fin = _hora_txt(end_h + 1)
        tipo = "dia_completo" if _es_cero(diurno) else "corte_nocturno"
    else:
        tipo = "con_consumo"

    perfil = " | ".join(f"{h:02d}:{float(horas.get(h, 0.0)):.2f}" for h in range(7))
    return Noche(
        fecha=dia.isoformat(),
        tiene_datos=True,
        en_cero_ventana=en_cero,
        tipo=tipo,
        inicio=inicio,
        fin=fin,
        inicio_desde_tarde_previa=desde_prev,
        h00=round(float(horas.get(0, 0.0)), 3),
        h01=round(float(horas.get(1, 0.0)), 3),
        h02=round(float(horas.get(2, 0.0)), 3),
        h03=round(float(horas.get(3, 0.0)), 3),
        h04=round(float(horas.get(4, 0.0)), 3),
        h05=round(float(horas.get(5, 0.0)), 3),
        h06=round(float(horas.get(6, 0.0)), 3),
        max_ventana=round(max_win, 3),
        consumo_diurno=round(diurno, 3),
        consumo_dia=round(total, 3),
        perfil_00_06=perfil,
    )


def agrupar_ciclos(noches: Sequence[Noche]) -> List[Ciclo]:
    ciclos: List[Ciclo] = []
    bloque: List[Noche] = []

    def cerrar() -> None:
        if not bloque:
            return
        ini = date.fromisoformat(bloque[0].fecha)
        fin = date.fromisoformat(bloque[-1].fecha)
        inicios = [n.inicio for n in bloque if n.inicio]
        fines = [n.fin for n in bloque if n.fin]
        n_corte = sum(1 for n in bloque if n.tipo == "corte_nocturno")
        n_comp = sum(1 for n in bloque if n.tipo == "dia_completo")
        ciclos.append(
            Ciclo(
                n=len(ciclos) + 1,
                inicio=ini,
                fin=fin,
                dias=len(bloque),
                n_corte=n_corte,
                n_completo=n_comp,
                inicio_hora_tipica=_moda(inicios),
                fin_hora_tipica=_moda(fines),
            )
        )

    for n in noches:
        if n.en_cero_ventana:
            if bloque:
                prev = date.fromisoformat(bloque[-1].fecha)
                cur = date.fromisoformat(n.fecha)
                if (cur - prev).days != 1:
                    cerrar()
                    bloque = []
            bloque.append(n)
        else:
            cerrar()
            bloque = []
    cerrar()
    return ciclos


def _moda(vals: Sequence[str]) -> str:
    if not vals:
        return "—"
    counts: Dict[str, int] = {}
    for v in vals:
        counts[v] = counts.get(v, 0) + 1
    return sorted(counts.items(), key=lambda x: (-x[1], x[0]))[0][0]


def _shade(cell, hex_color: str) -> None:
    shading = parse_xml(
        f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'w:val="clear" w:fill="{hex_color}"/>'
    )
    cell._tc.get_or_add_tcPr().append(shading)


def _header_row(table, fill: str = COLOR_WES) -> None:
    for cell in table.rows[0].cells:
        for p in cell.paragraphs:
            for run in p.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(9)
                run.font.name = "Calibri"
        _shade(cell, fill)


def _cell(cell, text: str, *, size: int = 8, bold: bool = False, fill: Optional[str] = None, center: bool = False) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    if center:
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run(str(text))
    run.font.size = Pt(size)
    run.bold = bold
    run.font.name = "Calibri"
    if fill:
        _shade(cell, fill)


def _add_heading(doc: Document, text: str, level: int = 1) -> None:
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor.from_string(COLOR_WES)


def guardar_graficos(
    noches: List[Noche],
    ciclos: List[Ciclo],
    serie: Dict[date, Dict[int, float]],
    out_dir: Path,
) -> Dict[str, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob("*.png"):
        old.unlink()
    paths: Dict[str, Path] = {}

    fechas = [date.fromisoformat(n.fecha) for n in noches if n.tiene_datos]
    flags = [1 if n.en_cero_ventana else 0 for n in noches if n.tiene_datos]
    if fechas:
        fig, ax = plt.subplots(figsize=(11.2, 3.6))
        fig.patch.set_facecolor("white")
        ax.set_facecolor("white")
        colors = [COLOR_NOCHE if f else COLOR_BARRA_WES for f in flags]
        ax.bar(fechas, flags, color=colors, width=1.0, align="center")
        ax.set_ylim(0, 1.35)
        ax.set_yticks([0, 1])
        ax.set_yticklabels(["con consumo", "en cero"])
        ax.set_title("Pizza Hut — noches en cero 00:30–05:00 (1 = ventana en 0 m³/h)", fontsize=11, fontweight="bold")
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%Y"))
        fig.autofmt_xdate(rotation=30)
        ax.grid(axis="y", alpha=0.3)
        p = out_dir / "timeline_noches_cero.png"
        fig.savefig(p, dpi=150, bbox_inches="tight")
        plt.close(fig)
        paths["timeline"] = p

    # Barras por mes
    by_mes: Dict[str, List[int]] = {}
    for n in noches:
        if not n.tiene_datos:
            continue
        d = date.fromisoformat(n.fecha)
        k = f"{d.year}-{d.month:02d}"
        by_mes.setdefault(k, [0, 0])
        if n.en_cero_ventana:
            by_mes[k][0] += 1
        else:
            by_mes[k][1] += 1
    if by_mes:
        keys = sorted(by_mes)
        x = np.arange(len(keys))
        cero = [by_mes[k][0] for k in keys]
        cons = [by_mes[k][1] for k in keys]
        fig, ax = plt.subplots(figsize=(11.2, 4.0))
        ax.bar(x, cero, color=COLOR_NOCHE, label="Noches en cero (00:30–05:00)")
        ax.bar(x, cons, bottom=cero, color=COLOR_BARRA_WES, label="Noches con consumo en la ventana")
        ax.set_xticks(x)
        ax.set_xticklabels(keys, rotation=45, ha="right", fontsize=8)
        ax.set_ylabel("Días")
        ax.set_title("Pizza Hut — noches en cero vs con consumo, por mes", fontsize=11, fontweight="bold")
        ax.legend(frameon=False, fontsize=8)
        ax.grid(axis="y", alpha=0.3)
        p = out_dir / "mensual_noches_cero.png"
        fig.savefig(p, dpi=150, bbox_inches="tight")
        plt.close(fig)
        paths["mensual"] = p

    # Perfiles: inicio y fin de cada ciclo (máx. 8 ciclos más recientes + el primero)
    ejemplos: List[Tuple[str, date]] = []
    if ciclos:
        ejemplos.append(("Inicio primer ciclo", ciclos[0].inicio))
        ejemplos.append(("Fin primer ciclo", ciclos[0].fin))
        if len(ciclos) > 1:
            ejemplos.append(("Inicio último ciclo", ciclos[-1].inicio))
            vigente = ciclos[-1].fin >= datetime.now().date()
            ejemplos.append(
                ("Último día (ciclo vigente)" if vigente else "Fin último ciclo", ciclos[-1].fin)
            )
        # Un ciclo intermedio si hay
        if len(ciclos) >= 3:
            mid = ciclos[len(ciclos) // 2]
            ejemplos.append((f"Inicio ciclo {mid.n}", mid.inicio))
            ejemplos.append((f"Fin ciclo {mid.n}", mid.fin))

    vistos = set()
    for titulo, dia in ejemplos:
        key = dia.isoformat()
        if key in vistos:
            continue
        vistos.add(key)
        horas = serie.get(dia) or {}
        if not horas:
            continue
        p = out_dir / f"perfil_{dia.isoformat()}.png"
        guardar_grafico_horario_24h_app(
            horas,
            p,
            titulo=f"Pizza Hut — {titulo} ({_fmt_fecha(dia)})",
            subtitulo="Consumo (m³/h) — hora Chile. Rojo: 00:00–06:00. Ventana analizada 00:30–05:00.",
        )
        paths[f"perfil_{key}"] = p

    return paths


def construir_word(
    noches: List[Noche],
    ciclos: List[Ciclo],
    charts: Dict[str, Path],
    out_path: Path,
    desde: date,
    hasta: date,
    primer_dato: Optional[date],
) -> Path:
    n_datos = sum(1 for n in noches if n.tiene_datos)
    n_cero = sum(1 for n in noches if n.en_cero_ventana)
    n_corte = sum(1 for n in noches if n.tipo == "corte_nocturno")
    n_comp = sum(1 for n in noches if n.tipo == "dia_completo")
    n_cons = sum(1 for n in noches if n.tipo == "con_consumo")
    n_sin = sum(1 for n in noches if n.tipo == "sin_datos")

    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(1.6)
        section.bottom_margin = Cm(1.6)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("INFORME WES — PUNTO EN CERO NOCTURNO")
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor.from_string(COLOR_WES)

    st = doc.add_paragraph()
    st.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = st.add_run(f"{NODE_NAME}  ·  {NODE_ID}  ·  {COMPANY}")
    r.bold = True
    r.font.size = Pt(13)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(
        f"Ventana 00:30–05:00  |  Desde 2025  |  Periodo con datos: "
        f"{_fmt_fecha(primer_dato) if primer_dato else '—'} a {_fmt_fecha(hasta)}"
    )
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    _add_heading(doc, "1. Criterio", 1)
    doc.add_paragraph(
        "Se revisó la serie horaria del punto Pizza Hut (000025-07) desde el 01-01-2025. "
        "La API entrega un valor por hora (m³/h, hora civil Chile). En esa resolución, "
        "la ventana 00:30–05:00 queda cubierta por las horas 01:00, 02:00, 03:00 y 04:00."
    )
    doc.add_paragraph(
        "Una noche se clasifica «en cero» si esas cuatro horas están en 0,00 m³/h. "
        "El inicio del ciclo es la primera hora en cero de la tira continua que cubre la ventana "
        "(00:00 si la hora 00 también está en cero; 01:00 si a las 00:00 aún había consumo; "
        "o 23:00 del día anterior si el corte partió la tarde previa). "
        "El fin del ciclo es la primera hora con consumo > 0 tras esa tira (habitualmente 06:00 o 07:00)."
    )
    doc.add_paragraph(
        "Los ciclos agrupados son rachas de noches consecutivas en cero. "
        "En las tablas se destaca el primer día (inicio, verde) y el último día (fin, naranja) de cada racha. "
        "Si la racha dura un solo día, esa noche es a la vez inicio y fin (amarillo)."
    )
    doc.add_paragraph(
        "Lectura operativa: hasta mayo de 2026 las noches en cero son esporádicas (1 a 3 días). "
        "Desde junio de 2026 aparece el corte automático habitual (cerca de 00:00–07:00). "
        "Desde el 09-08-2026 el control se mantiene casi todas las noches hasta hoy."
    )

    _add_heading(doc, "2. Resumen", 1)
    table = doc.add_table(rows=7, cols=2)
    table.style = "Table Grid"
    filas = [
        ("Punto", f"{NODE_NAME} ({NODE_ID})"),
        ("Primer día con datos", _fmt_fecha(primer_dato) if primer_dato else "sin datos en 2025"),
        ("Días con serie horaria", str(n_datos)),
        ("Noches en cero 00:30–05:00", f"{n_cero}  ({n_corte} con consumo diurno + {n_comp} día completo en cero)"),
        ("Noches con consumo en la ventana", str(n_cons)),
        ("Días sin datos (2025 previo a operación / huecos)", str(n_sin)),
        ("Ciclos (rachas consecutivas)", str(len(ciclos))),
    ]
    for i, (a, b) in enumerate(filas):
        _cell(table.rows[i].cells[0], a, size=10, bold=True, fill="D6E3F0")
        _cell(table.rows[i].cells[1], b, size=10)

    if charts.get("timeline"):
        doc.add_paragraph()
        pic = doc.add_paragraph()
        pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic.add_run().add_picture(str(charts["timeline"]), width=Inches(6.4))
    if charts.get("mensual"):
        pic = doc.add_paragraph()
        pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic.add_run().add_picture(str(charts["mensual"]), width=Inches(6.4))

    _add_heading(doc, "3. Ciclos: inicio y fin de cada racha", 1)
    doc.add_paragraph(
        "Cada fila es una racha de noches seguidas en cero. Verde = día de inicio. "
        "Naranja = día de fin. La hora típica resume la marca más frecuente de cierre/apertura en esa racha."
    )

    if not ciclos:
        doc.add_paragraph("No se detectaron noches en cero en la ventana 00:30–05:00.")
    else:
        hdr = [
            "Ciclo",
            "Inicio (1.er día)",
            "Fin (último día)",
            "Noches",
            "Hora inicio típica",
            "Hora fin típica",
            "Tipo",
        ]
        t2 = doc.add_table(rows=1 + len(ciclos), cols=len(hdr))
        t2.style = "Table Grid"
        for i, h in enumerate(hdr):
            _cell(t2.rows[0].cells[i], h, size=8, bold=True, center=True)
        _header_row(t2)
        for i, c in enumerate(ciclos, start=1):
            tipo = []
            if c.n_corte:
                tipo.append(f"{c.n_corte} corte nocturno")
            if c.n_completo:
                tipo.append(f"{c.n_completo} día completo")
            fin_txt = _fmt_fecha_larga(c.fin)
            if i == len(ciclos) and c.fin >= hasta:
                fin_txt += " (vigente)"
            vals = [
                str(c.n),
                _fmt_fecha_larga(c.inicio),
                fin_txt,
                str(c.dias),
                c.inicio_hora_tipica,
                c.fin_hora_tipica,
                " + ".join(tipo) or "—",
            ]
            for j, v in enumerate(vals):
                fill = None
                if j == 1:
                    fill = COLOR_UNICO if c.inicio == c.fin else COLOR_INICIO
                elif j == 2:
                    fill = COLOR_UNICO if c.inicio == c.fin else COLOR_FIN
                _cell(t2.rows[i].cells[j], v, size=8, fill=fill)

    # Detalle destacado de cada ciclo (primer y último día)
    _add_heading(doc, "4. Detalle del inicio y del fin de cada ciclo", 1)
    if ciclos:
        by_fecha = {n.fecha: n for n in noches}
        t3 = doc.add_table(rows=1, cols=8)
        t3.style = "Table Grid"
        headers = [
            "Ciclo / rol",
            "Fecha",
            "Inicio ciclo",
            "Fin ciclo",
            "00–06 (m³/h)",
            "Diurno m³",
            "Día m³",
            "Tipo",
        ]
        for i, h in enumerate(headers):
            _cell(t3.rows[0].cells[i], h, size=8, bold=True, center=True)
        _header_row(t3)

        def add_detalle(ciclo: Ciclo, rol: str, d: date, fill: str) -> None:
            n = by_fecha.get(d.isoformat())
            if n is None:
                return
            row = t3.add_row()
            tipo_txt = {
                "corte_nocturno": "Corte nocturno (hay consumo de día)",
                "dia_completo": "Día completo en cero",
            }.get(n.tipo, n.tipo)
            vals = [
                f"C{ciclo.n} · {rol}",
                _fmt_fecha_larga(d),
                n.inicio or "—",
                n.fin or "—",
                n.perfil_00_06,
                f"{n.consumo_diurno:.2f}".replace(".", ","),
                f"{n.consumo_dia:.2f}".replace(".", ","),
                tipo_txt,
            ]
            for j, v in enumerate(vals):
                _cell(row.cells[j], v, size=7, fill=fill, bold=(j == 0))

        for c in ciclos:
            if c.inicio == c.fin:
                add_detalle(c, "inicio y fin", c.inicio, COLOR_UNICO)
            else:
                add_detalle(c, "INICIO", c.inicio, COLOR_INICIO)
                add_detalle(c, "FIN", c.fin, COLOR_FIN)

    perfiles = [(k, p) for k, p in charts.items() if k.startswith("perfil_")]
    if perfiles:
        _add_heading(doc, "5. Perfiles horarios — inicio y fin de ciclo", 1)
        doc.add_paragraph(
            "Barras rojas = madrugada 00:00–06:00. Se muestran el primer y el último día "
            "del primer ciclo, de un ciclo intermedio y del ciclo más reciente."
        )
        for _, pth in perfiles:
            pic = doc.add_paragraph()
            pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pic.add_run().add_picture(str(pth), width=Inches(6.3))

    _add_heading(doc, "6. Listado de todas las noches en cero", 1)
    doc.add_paragraph(
        "Solo días en que la ventana 00:30–05:00 estuvo en 0,00 m³/h. "
        "Verde = inicio de ciclo. Naranja = fin de ciclo. Amarillo = ciclo de un solo día."
    )
    cero_nights = [n for n in noches if n.en_cero_ventana]
    extremos: Dict[str, str] = {}
    for c in ciclos:
        if c.inicio == c.fin:
            extremos[c.inicio.isoformat()] = "unico"
        else:
            extremos[c.inicio.isoformat()] = "inicio"
            extremos[c.fin.isoformat()] = "fin"

    if cero_nights:
        hdr = ["Fecha", "Rol", "Inicio", "Fin", "00", "01", "02", "03", "04", "05", "06", "Tipo"]
        t4 = doc.add_table(rows=1 + len(cero_nights), cols=len(hdr))
        t4.style = "Table Grid"
        for i, h in enumerate(hdr):
            _cell(t4.rows[0].cells[i], h, size=7, bold=True, center=True)
        _header_row(t4)
        for i, n in enumerate(cero_nights, start=1):
            rol = extremos.get(n.fecha, "")
            rol_txt = {"inicio": "INICIO ciclo", "fin": "FIN ciclo", "unico": "inicio y fin"}.get(rol, "")
            fill = {"inicio": COLOR_INICIO, "fin": COLOR_FIN, "unico": COLOR_UNICO}.get(rol)
            if fill is None and i % 2 == 0:
                fill = COLOR_ALT
            d = date.fromisoformat(n.fecha)
            tipo_txt = "día completo" if n.tipo == "dia_completo" else "corte nocturno"
            vals = [
                _fmt_fecha(d),
                rol_txt,
                n.inicio,
                n.fin,
                f"{n.h00:.2f}",
                f"{n.h01:.2f}",
                f"{n.h02:.2f}",
                f"{n.h03:.2f}",
                f"{n.h04:.2f}",
                f"{n.h05:.2f}",
                f"{n.h06:.2f}",
                tipo_txt,
            ]
            for j, v in enumerate(vals):
                _cell(
                    t4.rows[i].cells[j],
                    v,
                    size=7,
                    fill=fill,
                    bold=bool(rol_txt) and j <= 1,
                    center=j >= 2,
                )

    _add_heading(doc, "7. Notas", 1)
    doc.add_paragraph(
        "• Resolución horaria: no hay marca a las 00:30 exactas. Si a las 00:00 hay consumo y "
        "desde las 01:00 está en cero, el corte ocurrió entre 00:00 y 01:00 (compatible con 00:30)."
    )
    doc.add_paragraph(
        "• El control automático de Pizza Hut en Parque Arauco Estación se opera habitualmente "
        "cerca de 00:00–06:00. Este informe aísla el núcleo 00:30–05:00 pedido."
    )
    doc.add_paragraph(
        f"• Informe generado {datetime.now().strftime('%d-%m-%Y %H:%M')} (hora del agente). "
        f"Periodo solicitado: desde 2025 hasta {_fmt_fecha(hasta)}."
    )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f"[OK] Word: {out_path}")
    return out_path


def construir_pdf(
    noches: List[Noche],
    ciclos: List[Ciclo],
    charts: Dict[str, Path],
    out_path: Path,
    desde: date,
    hasta: date,
    primer_dato: Optional[date],
) -> Path:
    from matplotlib.backends.backend_pdf import PdfPages

    n_cero = sum(1 for n in noches if n.en_cero_ventana)
    n_corte = sum(1 for n in noches if n.tipo == "corte_nocturno")
    n_comp = sum(1 for n in noches if n.tipo == "dia_completo")
    n_cons = sum(1 for n in noches if n.tipo == "con_consumo")
    n_datos = sum(1 for n in noches if n.tiene_datos)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with PdfPages(out_path) as pdf:
        fig = plt.figure(figsize=(8.27, 11.69))
        ax = fig.add_subplot(111)
        ax.axis("off")
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.text(0.5, 0.92, "INFORME WES", ha="center", fontsize=12, color="#1F4E79", fontweight="bold")
        ax.text(0.5, 0.87, "Pizza Hut — noches en cero 00:30 a 05:00", ha="center", fontsize=16, color="#1F4E79", fontweight="bold")
        ax.text(0.5, 0.83, f"{NODE_ID}  ·  {COMPANY}", ha="center", fontsize=11)
        lineas = [
            f"Periodo: 01-01-2025 a {_fmt_fecha(hasta)}",
            f"Primer día con datos: {_fmt_fecha(primer_dato) if primer_dato else '—'}",
            f"Días con serie horaria: {n_datos}",
            f"Noches en cero (01:00–04:00 = 0 m³/h): {n_cero}",
            f"  · corte nocturno (hay consumo de día): {n_corte}",
            f"  · día completo en cero: {n_comp}",
            f"Noches con consumo en la ventana: {n_cons}",
            f"Ciclos (rachas consecutivas): {len(ciclos)}",
            "",
            "Criterio: ventana 00:30–05:00, resolución horaria → horas 01, 02, 03 y 04 en 0,00 m³/h.",
            "Inicio del ciclo = primera hora en cero de la tira continua.",
            "Fin del ciclo = primera hora con consumo > 0 después de esa tira.",
            "En tablas: verde = inicio de racha; naranja = fin de racha; amarillo = racha de 1 día.",
        ]
        y = 0.76
        for ln in lineas:
            ax.text(0.08, y, ln, ha="left", va="top", fontsize=10, family="DejaVu Sans")
            y -= 0.032
        pdf.savefig(fig)
        plt.close(fig)

        for key in ("timeline", "mensual"):
            pth = charts.get(key)
            if not pth or not pth.exists():
                continue
            img = plt.imread(str(pth))
            fig = plt.figure(figsize=(8.27, 11.69))
            ax = fig.add_subplot(111)
            ax.imshow(img)
            ax.axis("off")
            fig.tight_layout()
            pdf.savefig(fig)
            plt.close(fig)

        # Tabla de ciclos
        if ciclos:
            headers = ["N", "Inicio", "Fin", "Noches", "Hora ini.", "Hora fin"]
            page = 22
            for start in range(0, len(ciclos), page):
                chunk = ciclos[start : start + page]
                data = [headers]
                for c in chunk:
                    data.append(
                        [
                            str(c.n),
                            _fmt_fecha(c.inicio),
                            _fmt_fecha(c.fin),
                            str(c.dias),
                            c.inicio_hora_tipica,
                            c.fin_hora_tipica,
                        ]
                    )
                fig = plt.figure(figsize=(8.27, 11.69))
                ax = fig.add_subplot(111)
                ax.axis("off")
                ax.set_title("Ciclos — inicio y fin de cada racha", fontsize=12, color="#1F4E79", pad=16)
                tbl = ax.table(cellText=data, loc="upper center", cellLoc="center")
                tbl.auto_set_font_size(False)
                tbl.set_fontsize(8)
                tbl.scale(1.15, 1.45)
                for j in range(len(headers)):
                    tbl[0, j].set_facecolor("#1F4E79")
                    tbl[0, j].set_text_props(color="white", fontweight="bold")
                for i, c in enumerate(chunk, start=1):
                    tbl[i, 1].set_facecolor("#FFF2CC" if c.inicio == c.fin else "#C6EFCE")
                    tbl[i, 2].set_facecolor("#FFF2CC" if c.inicio == c.fin else "#FCE4D6")
                pdf.savefig(fig)
                plt.close(fig)

        # Detalle inicio/fin
        if ciclos:
            by_fecha = {n.fecha: n for n in noches}
            rows = [["Ciclo / rol", "Fecha", "Inicio", "Fin", "00-06", "Tipo"]]
            for c in ciclos:
                pares = [("inicio y fin", c.inicio)] if c.inicio == c.fin else [("INICIO", c.inicio), ("FIN", c.fin)]
                for rol, d in pares:
                    n = by_fecha.get(d.isoformat())
                    if not n:
                        continue
                    rows.append(
                        [
                            f"C{c.n} {rol}",
                            _fmt_fecha(d),
                            n.inicio,
                            n.fin,
                            n.perfil_00_06.replace(" | ", " "),
                            "día completo" if n.tipo == "dia_completo" else "corte nocturno",
                        ]
                    )
            page = 18
            for start in range(0, len(rows) - 1, page):
                chunk = [rows[0]] + rows[1 + start : 1 + start + page]
                fig = plt.figure(figsize=(8.27, 11.69))
                ax = fig.add_subplot(111)
                ax.axis("off")
                ax.set_title("Detalle inicio y fin de cada ciclo", fontsize=12, color="#1F4E79", pad=16)
                tbl = ax.table(cellText=chunk, loc="upper center", cellLoc="left")
                tbl.auto_set_font_size(False)
                tbl.set_fontsize(6.5)
                tbl.scale(1.15, 1.55)
                for j in range(len(chunk[0])):
                    tbl[0, j].set_facecolor("#1F4E79")
                    tbl[0, j].set_text_props(color="white", fontweight="bold")
                pdf.savefig(fig)
                plt.close(fig)

        perfiles = [(k, p) for k, p in charts.items() if k.startswith("perfil_")]
        for _, pth in perfiles:
            img = plt.imread(str(pth))
            fig = plt.figure(figsize=(8.27, 11.69))
            ax = fig.add_subplot(111)
            ax.imshow(img)
            ax.axis("off")
            fig.tight_layout()
            pdf.savefig(fig)
            plt.close(fig)

        # Listado completo
        cero_nights = [n for n in noches if n.en_cero_ventana]
        extremos: Dict[str, str] = {}
        for c in ciclos:
            if c.inicio == c.fin:
                extremos[c.inicio.isoformat()] = "unico"
            else:
                extremos[c.inicio.isoformat()] = "inicio"
                extremos[c.fin.isoformat()] = "fin"
        headers = ["Fecha", "Rol", "Inicio", "Fin", "00", "01", "02", "03", "04", "05", "06"]
        page = 28
        for start in range(0, len(cero_nights), page):
            chunk = cero_nights[start : start + page]
            data = [headers]
            fills: List[Optional[str]] = []
            for n in chunk:
                rol = extremos.get(n.fecha, "")
                rol_txt = {"inicio": "INICIO", "fin": "FIN", "unico": "inicio y fin"}.get(rol, "")
                data.append(
                    [
                        date.fromisoformat(n.fecha).strftime("%d-%m-%Y"),
                        rol_txt,
                        n.inicio,
                        n.fin,
                        f"{n.h00:.2f}",
                        f"{n.h01:.2f}",
                        f"{n.h02:.2f}",
                        f"{n.h03:.2f}",
                        f"{n.h04:.2f}",
                        f"{n.h05:.2f}",
                        f"{n.h06:.2f}",
                    ]
                )
                fills.append({"inicio": "#C6EFCE", "fin": "#FCE4D6", "unico": "#FFF2CC"}.get(rol))
            fig = plt.figure(figsize=(8.27, 11.69))
            ax = fig.add_subplot(111)
            ax.axis("off")
            ax.set_title("Todas las noches en cero 00:30–05:00", fontsize=12, color="#1F4E79", pad=16)
            tbl = ax.table(cellText=data, loc="upper center", cellLoc="center")
            tbl.auto_set_font_size(False)
            tbl.set_fontsize(7)
            tbl.scale(1.15, 1.35)
            for j in range(len(headers)):
                tbl[0, j].set_facecolor("#1F4E79")
                tbl[0, j].set_text_props(color="white", fontweight="bold")
            for i, fill in enumerate(fills, start=1):
                if fill:
                    for j in range(len(headers)):
                        tbl[i, j].set_facecolor(fill)
            pdf.savefig(fig)
            plt.close(fig)

    print(f"[OK] PDF: {out_path}")
    return out_path


def guardar_csv(noches: List[Noche], ciclos: List[Ciclo], out_dir: Path) -> Tuple[Path, Path]:
    p1 = out_dir / "pizza_hut_noches_cero_0030_0500.csv"
    extremos: Dict[str, str] = {}
    for c in ciclos:
        if c.inicio == c.fin:
            extremos[c.inicio.isoformat()] = "inicio y fin"
        else:
            extremos[c.inicio.isoformat()] = "INICIO ciclo"
            extremos[c.fin.isoformat()] = "FIN ciclo"
    with p1.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(
            [
                "fecha",
                "rol_ciclo",
                "en_cero_0030_0500",
                "tipo",
                "inicio_ciclo",
                "fin_ciclo",
                "h00",
                "h01",
                "h02",
                "h03",
                "h04",
                "h05",
                "h06",
                "consumo_diurno_m3",
                "consumo_dia_m3",
            ]
        )
        for n in noches:
            if not n.en_cero_ventana:
                continue
            w.writerow(
                [
                    n.fecha,
                    extremos.get(n.fecha, ""),
                    "SI",
                    n.tipo,
                    n.inicio,
                    n.fin,
                    n.h00,
                    n.h01,
                    n.h02,
                    n.h03,
                    n.h04,
                    n.h05,
                    n.h06,
                    n.consumo_diurno,
                    n.consumo_dia,
                ]
            )
    p2 = out_dir / "pizza_hut_ciclos_cero_nocturno.csv"
    with p2.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(
            [
                "ciclo",
                "inicio",
                "fin",
                "noches",
                "hora_inicio_tipica",
                "hora_fin_tipica",
                "n_corte_nocturno",
                "n_dia_completo",
            ]
        )
        for c in ciclos:
            w.writerow(
                [
                    c.n,
                    c.inicio.isoformat(),
                    c.fin.isoformat(),
                    c.dias,
                    c.inicio_hora_tipica,
                    c.fin_hora_tipica,
                    c.n_corte,
                    c.n_completo,
                ]
            )
    return p1, p2


def main() -> int:
    hoy = datetime.now().date()
    desde = date(2025, 1, 1)
    hasta = hoy
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = Path("reports/Puntos_En_Cero/Pizza_Hut")
    out_dir.mkdir(parents=True, exist_ok=True)
    charts_dir = out_dir / "graficos"

    serie = descargar_serie(desde, hasta)
    primer_dato = next((d for d in sorted(serie) if serie[d]), None)

    noches: List[Noche] = []
    prev_h23: Optional[float] = None
    for d in sorted(serie):
        horas = serie[d]
        noches.append(analizar_noche(d, horas, prev_h23))
        prev_h23 = float(horas[23]) if horas and 23 in horas else None

    ciclos = agrupar_ciclos(noches)
    n_cero = sum(1 for n in noches if n.en_cero_ventana)
    print(f"[INFO] Noches en cero: {n_cero}  |  ciclos: {len(ciclos)}")
    for c in ciclos:
        print(
            f"  Ciclo {c.n:02d}: {_fmt_fecha(c.inicio)} → {_fmt_fecha(c.fin)}  "
            f"({c.dias} noches)  inicio {c.inicio_hora_tipica}  fin {c.fin_hora_tipica}"
        )

    charts = guardar_graficos(noches, ciclos, serie, charts_dir)
    csv_n, csv_c = guardar_csv(noches, ciclos, out_dir)
    print(f"[OK] CSV: {csv_n}")
    print(f"[OK] CSV: {csv_c}")

    resumen = {
        "node_id": NODE_ID,
        "desde": desde.isoformat(),
        "hasta": hasta.isoformat(),
        "primer_dato": primer_dato.isoformat() if primer_dato else None,
        "noches_en_cero": n_cero,
        "n_ciclos": len(ciclos),
        "ciclos": [
            {
                "n": c.n,
                "inicio": c.inicio.isoformat(),
                "fin": c.fin.isoformat(),
                "dias": c.dias,
                "n_corte": c.n_corte,
                "n_completo": c.n_completo,
                "inicio_hora_tipica": c.inicio_hora_tipica,
                "fin_hora_tipica": c.fin_hora_tipica,
            }
            for c in ciclos
        ],
    }
    (out_dir / "resumen_ciclos.json").write_text(json.dumps(resumen, indent=2, ensure_ascii=False), encoding="utf-8")

    docx_path = out_dir / f"Pizza_Hut_cero_nocturno_0030_0500_{stamp}.docx"
    pdf_path = out_dir / f"Pizza_Hut_cero_nocturno_0030_0500_{stamp}.pdf"
    construir_word(noches, ciclos, charts, docx_path, desde, hasta, primer_dato)
    construir_pdf(noches, ciclos, charts, pdf_path, desde, hasta, primer_dato)

    print(f"DOCX={docx_path}")
    print(f"PDF={pdf_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
