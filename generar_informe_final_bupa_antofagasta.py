"""
Informe final — Estudio de diagnóstico de gestión hídrica
Clínica Bupa Antofagasta (solo Antofagasta; sin contenido de otros recintos).

Modelo: Informe final WES - BPA (15.10.2026)
Datos: cache diario WES + facturas Aguas Antofagasta + fotos oficiales Drive.
"""

from __future__ import annotations

import json
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from generar_reporte_word import format_number_chilean

ROOT = Path("reports/Bupa_Antofagasta/INFORME_FINAL")
CHARTS = ROOT / "charts"
FOTOS = ROOT / "fotos"
CACHE = Path(
    "reports/Bupa_Antofagasta/ANALISIS_PILOTO/cache_diario_bupa_antofagasta.json"
)

START = date(2026, 7, 23)
END = date(2026, 10, 11)
NUM_DIAS = (END - START).days + 1  # 81
DAY_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
NODOS = {
    "000029-09": ("Sanitaria", "Medidor Principal Sanitaria"),
    "000029-07": ("Bomba Principal", "Sala de Bomba Principal"),
    "000029-08": ("Sexto Piso", "Sala de Bomba Sexto Piso"),
}

# Facturas Aguas Antofagasta — G:\\Mi unidad\\Bupa Antofagasta\\Facturacion
FACTURAS = [
    {
        "ciclo": "Junio 2026",
        "emision": "02-06-2026",
        "vence": "17-06-2026",
        "m3": 7102.0,
        "clp": 19_662_850.0,
        "nro": "1453711",
        "lect_ant": 87562,
        "lect_act": 94664,
    },
    {
        "ciclo": "Julio 2026",
        "emision": "03-07-2026",
        "vence": "20-07-2026",
        "m3": 6696.0,
        "clp": 18_538_860.0,
        "nro": "1460583",
        "lect_ant": 94664,
        "lect_act": 101360,
    },
    {
        "ciclo": "Agosto 2026",
        "emision": "04-08-2026",
        "vence": "19-08-2026",
        "m3": 7627.0,
        "clp": 21_116_260.0,
        "nro": "1467321",
        "lect_ant": 101360,
        "lect_act": 108987,
    },
    {
        "ciclo": "Septiembre 2026",
        "emision": "27-08-2026",
        "vence": "11-09-2026",
        "m3": 2873.0,
        "clp": 7_955_190.0,
        "nro": "1474021",
        "lect_ant": 108987,
        "lect_act": 111860,
    },
]


def _load_stats_and_charts() -> dict:
    """Lee cache diario, calcula KPIs y regenera gráficos propios."""
    CHARTS.mkdir(parents=True, exist_ok=True)
    cache = json.loads(CACHE.read_text(encoding="utf-8"))
    stats: dict = {}

    def series(nid: str):
        s = cache.get(nid, {})
        out = []
        cur = START
        while cur <= END:
            if cur.isoformat() in s:
                out.append((cur, float(s[cur.isoformat()])))
            cur += timedelta(days=1)
        return out

    for nid, (corto, nombre) in NODOS.items():
        pts = series(nid)
        vals = [v for _, v in pts]
        slug = corto.lower().replace(" ", "_")
        stats[nid] = {
            "corto": corto,
            "nombre": nombre,
            "total": sum(vals),
            "n": len(vals),
            "prom": sum(vals) / len(vals) if vals else 0.0,
            "mn": min(vals) if vals else 0.0,
            "mx": max(vals) if vals else 0.0,
            "dmin": min(pts, key=lambda x: x[1])[0].isoformat() if pts else None,
            "dmax": max(pts, key=lambda x: x[1])[0].isoformat() if pts else None,
        }
        # diario
        fig, ax = plt.subplots(figsize=(9.2, 2.6))
        dates = [datetime.combine(d, datetime.min.time()) for d, _ in pts]
        totals = [v for _, v in pts]
        ax.plot(dates, totals, marker="o", markersize=2.2, color="#0050b3", lw=1.4)
        ax.fill_between(dates, totals, alpha=0.22, color="#0050b3")
        ax.set_title(f"Perfil de consumo diario (m³) — {corto}", fontsize=11, fontweight="bold")
        ax.set_ylabel("m³")
        ax.set_ylim(bottom=0)
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(dates) // 10)))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%b"))
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=35, ha="right", fontsize=7)
        ax.grid(True, ls="--", alpha=0.3, axis="y")
        fig.tight_layout(pad=0.3)
        fig.savefig(CHARTS / f"chart_diario_{slug}.png", dpi=150, bbox_inches="tight")
        plt.close(fig)
        # lun-dom
        buckets: dict = defaultdict(list)
        for d, v in pts:
            buckets[d.weekday()].append(v)
        avgs = [sum(buckets[i]) / len(buckets[i]) if buckets[i] else 0 for i in range(7)]
        fig, ax = plt.subplots(figsize=(9.0, 2.4))
        bars = ax.bar(DAY_ES, avgs, color="#0050b3", alpha=0.85, edgecolor="#003d8c")
        for b, v in zip(bars, avgs):
            ax.text(
                b.get_x() + b.get_width() / 2,
                b.get_height(),
                f"{v:.1f}".replace(".", ","),
                ha="center",
                va="bottom",
                fontsize=8,
                fontweight="bold",
            )
        ax.set_title(f"Promedio lun–dom (m³) — {corto}", fontsize=11, fontweight="bold")
        ax.set_ylabel("m³")
        ax.set_ylim(bottom=0)
        ax.grid(True, ls="--", alpha=0.3, axis="y")
        fig.tight_layout(pad=0.3)
        fig.savefig(CHARTS / f"chart_dow_{slug}.png", dpi=150, bbox_inches="tight")
        plt.close(fig)

    san = stats["000029-09"]["total"]
    p = stats["000029-07"]["total"]
    x = stats["000029-08"]["total"]
    gap = max(0.0, san - p - x)
    fig, ax = plt.subplots(figsize=(4.2, 3.2))
    ax.pie(
        [p, x, gap],
        labels=["Principal", "Sexto", "No monitoreado"],
        colors=["#0050b3", "#40a9ff", "#c0392b"],
        autopct=lambda pct: f"{pct:.0f}%",
        startangle=90,
        textprops={"fontsize": 9},
    )
    ax.set_title("% sobre Sanitaria (periodo)", fontsize=10, fontweight="bold")
    fig.tight_layout()
    fig.savefig(CHARTS / "chart_participacion.png", dpi=140, bbox_inches="tight")
    plt.close(fig)

    ciclos = [("Jun-26", 7102), ("Jul-26", 6696), ("Ago-26", 7627), ("Sep-26", 2873)]
    fig, ax = plt.subplots(figsize=(7.5, 3.0))
    ax.bar([c[0] for c in ciclos], [c[1] for c in ciclos], color="#0050b3", edgecolor="#003d8c")
    for i, (_, v) in enumerate(ciclos):
        ax.text(i, v, f"{v:,}".replace(",", "."), ha="center", va="bottom", fontsize=9, fontweight="bold")
    ax.set_ylabel("m³")
    ax.set_title("Consumo facturado Aguas Antofagasta (m³)", fontsize=11, fontweight="bold")
    ax.set_ylim(bottom=0)
    ax.grid(True, ls="--", alpha=0.3, axis="y")
    fig.tight_layout()
    fig.savefig(CHARTS / "chart_facturas_m3.png", dpi=140, bbox_inches="tight")
    plt.close(fig)
    return stats


def _fmt_clp(v: float) -> str:
    return f"${format_number_chilean(v, 0)}"


def _fmt_m3(v: float, nd: int = 1) -> str:
    return format_number_chilean(v, nd)


def _fmt_fecha(d: date | str) -> str:
    if isinstance(d, str):
        d = date.fromisoformat(d)
    return d.strftime("%d/%m/%Y")


def _set_run(run, *, size=10, bold=False, color=None, highlight=None) -> None:
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color or RGBColor(0, 0, 0)
    if highlight is not None:
        run.font.highlight_color = highlight


def _add_para(
    doc: Document,
    text: str,
    *,
    size: int = 10,
    bold: bool = False,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    space_after: int = 6,
    space_before: int = 0,
    highlight: bool = False,
    bullet: bool = False,
) -> None:
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    if bullet:
        p.style = "List Bullet"
    prefix = ""
    r = p.add_run(prefix + text)
    _set_run(
        r,
        size=size,
        bold=bold,
        highlight=WD_COLOR_INDEX.YELLOW if highlight else None,
    )


def _add_title(doc: Document, text: str, *, size: int = 14) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    _set_run(r, size=size, bold=True, color=RGBColor(0, 51, 102))


def _add_subtitle(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    _set_run(r, size=11, bold=True, color=RGBColor(0, 51, 102))


def _page_break(doc: Document) -> None:
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def _add_picture(doc: Document, path: Path, width_in: float, caption: str = "") -> None:
    if not path.is_file():
        _add_para(doc, f"[Sin imagen: {path.name}]", size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(str(path), width=Inches(width_in))
    if caption:
        c = doc.add_paragraph()
        c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraph_format.space_after = Pt(8)
        r = c.add_run(caption)
        _set_run(r, size=8, bold=False, color=RGBColor(80, 80, 80))


def _center_table(table) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _shade_header(row) -> None:
    for cell in row.cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "003366")
        shd.set(qn("w:val"), "clear")
        tcPr.append(shd)
        for para in cell.paragraphs:
            for run in para.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.bold = True
                run.font.size = Pt(9)


def _add_table(doc: Document, headers: list[str], rows: list[list[str]], *, compact=False) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for j, h in enumerate(headers):
        table.rows[0].cells[j].text = h
    _shade_header(table.rows[0])
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = table.rows[i].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            r = p.add_run(val)
            _set_run(r, size=8 if compact else 9, bold=(i == len(rows) and "Total" in str(row[0])))
    _center_table(table)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)


def _setup(doc: Document) -> None:
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)


def build() -> Path:
    ROOT.mkdir(parents=True, exist_ok=True)
    stats = _load_stats_and_charts()
    san = stats["000029-09"]
    prin = stats["000029-07"]
    sexto = stats["000029-08"]
    tot_salas = prin["total"] + sexto["total"]
    gap = max(0.0, san["total"] - tot_salas)
    pct_prin = 100.0 * prin["total"] / san["total"]
    pct_sexto = 100.0 * sexto["total"] / san["total"]
    pct_salas = 100.0 * tot_salas / san["total"]
    pct_gap = 100.0 * gap / san["total"]
    tot_fact_m3 = sum(f["m3"] for f in FACTURAS)
    tot_fact_clp = sum(f["clp"] for f in FACTURAS)

    doc = Document()
    _setup(doc)

    # ===== PORTADA =====
    for _ in range(4):
        doc.add_paragraph()
    _add_para(
        doc,
        "Estudio de diagnóstico de gestión hídrica",
        size=20,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=12,
    )
    _add_para(
        doc,
        "Clínica Bupa Antofagasta",
        size=16,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=8,
    )
    _add_para(
        doc,
        "WES — Monitoreo Inteligente",
        size=12,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=20,
    )
    _add_para(
        doc,
        "Octubre 2026",
        size=12,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    _add_para(
        doc,
        f"Periodo de monitoreo: {_fmt_fecha(START)} – {_fmt_fecha(END)} ({NUM_DIAS} días)",
        size=10,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_before=8,
    )

    # ===== ÍNDICE =====
    _page_break(doc)
    _add_title(doc, "Índice", size=16)
    indice = [
        ("Resumen Ejecutivo", "2"),
        ("I. Antecedentes", "3"),
        ("II. Objetivos y resultados esperados", "4"),
        ("III. Definición de sectores a monitorear", "4"),
        ("IV. Fotografías de instalación", "5"),
        ("V. Análisis de sectores monitoreados", "6"),
        ("    1) Medidor Principal Sanitaria", "6"),
        ("    2) Sala de Bomba Principal", "7"),
        ("    3) Sala de Bomba Sexto Piso", "8"),
        ("    4) Sala de Bomba N°2 (nota operativa)", "8"),
        ("VI. Facturación sanitaria (Aguas Antofagasta)", "9"),
        ("VII. Conclusiones y hallazgos", "10"),
        ("VIII. Recomendaciones y próximos pasos", "11"),
    ]
    for titulo, pag in indice:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run(titulo)
        _set_run(r1, size=10)
        tab = p.add_run("\t" + pag)
        _set_run(tab, size=10)

    # ===== RESUMEN EJECUTIVO (pág. 3 del pedido) =====
    _page_break(doc)
    _add_title(doc, "Resumen Ejecutivo")
    _add_subtitle(doc, "Propósito y Alcance")
    _add_para(
        doc,
        "El Estudio tuvo como objetivo principal establecer una línea de base detallada del consumo "
        "hídrico de la Clínica Bupa Antofagasta y validar la solución de monitoreo inteligente de WES "
        "para potenciar la capacidad de gestión eficiente y sostenible del agua.",
        bullet=True,
    )
    _add_para(
        doc,
        f"El estudio contempló el monitoreo en línea de 4 puntos críticos durante un periodo de "
        f"{NUM_DIAS} días ({_fmt_fecha(START)} al {_fmt_fecha(END)}), permitiendo identificar perfiles "
        f"de consumo diario y semanal, y detectar anomalías operacionales.",
        bullet=True,
    )

    _add_subtitle(doc, "Hallazgos Principales de Consumo")
    _add_para(
        doc,
        f"Brecha de visibilidad (agua no contabilizada en salas monitoreadas): existe una diferencia "
        f"de {_fmt_m3(gap)} m³ ({_fmt_m3(pct_gap)}% del total Sanitaria) entre el Medidor Principal "
        f"Sanitaria ({_fmt_m3(san['total'])} m³) y la suma de Bomba Principal + Sexto Piso "
        f"({_fmt_m3(tot_salas)} m³). Las hipótesis apuntan a tramos de red no instrumentados "
        f"(p. ej. empalmes previos a impulsión / otros consumos del recinto) y a la exclusión "
        f"analítica de Sala de Bomba N°2 por lecturas inconsistentes desde septiembre.",
        highlight=True,
        bullet=True,
    )
    _add_para(
        doc,
        f"Sanitaria concentra el 100% de la cuenta: promedio {_fmt_m3(san['prom'])} m³/día "
        f"(rango {_fmt_m3(san['mn'])}–{_fmt_m3(san['mx'])} m³/día). Bomba Principal representa "
        f"{_fmt_m3(pct_prin)}% y Sexto Piso {_fmt_m3(pct_sexto)}% del volumen Sanitaria.",
        highlight=True,
        bullet=True,
    )
    _add_para(
        doc,
        f"Facturación Aguas Antofagasta (jun–sep 2026, 4 ciclos): {_fmt_m3(tot_fact_m3, 0)} m³ y "
        f"{_fmt_clp(tot_fact_clp)}. El salto de volumen respecto del ritmo histórico se confirma "
        f"en lecturas encadenadas del medidor 534 (sin cambio de medidor ni estimación).",
        highlight=True,
        bullet=True,
    )

    _add_subtitle(doc, "Recomendaciones Estratégicas")
    _add_para(
        doc,
        "Monitoreo incremental: extender la medición a tramos no visualizados (incluye revisar "
        "Sala N°2 con validación en terreno y eventual recalibración) para cerrar la brecha "
        f"del {_fmt_m3(pct_gap)}%.",
        highlight=True,
        bullet=True,
    )
    _add_para(
        doc,
        "Apoyo tecnológico: implementar Control Inteligente WES en sectores de alto consumo "
        "para gestionar demanda en horarios inhábiles (prioritario) y hábiles.",
        highlight=True,
        bullet=True,
    )
    _add_para(
        doc,
        "Seguimiento factura vs WES: mantener contraste ciclo a ciclo entre boletas de Aguas "
        "Antofagasta (medidor 534) y el Medidor Principal Sanitaria WES.",
        highlight=True,
        bullet=True,
    )

    # ===== I. ANTECEDENTES =====
    _page_break(doc)
    _add_title(doc, "I. Antecedentes")
    _add_para(
        doc,
        "La Clínica Bupa tiene interés en avanzar en la gestión hídrica de su recinto y solicitó "
        "a WES el desarrollo de un estudio de diagnóstico en Clínica Bupa Antofagasta para el "
        "monitoreo en línea de sus redes hídricas.",
        bullet=True,
    )
    _add_para(
        doc,
        "El objetivo del Estudio es construir una línea de base detallada del consumo hídrico "
        "de los distintos sectores de la red del recinto.",
        bullet=True,
    )
    _add_para(
        doc,
        f"El proyecto se extendió entre julio y octubre 2026; por un periodo de {NUM_DIAS} días "
        f"se llevó adelante monitoreo en línea de 4 puntos seleccionados en Clínica Bupa Antofagasta.",
        bullet=True,
    )

    _add_subtitle(doc, "a. Hitos y Carta Gantt")
    _add_para(doc, "Previo a la implementación de los equipos, se ejecutaron hitos preparativos:")
    for hito in [
        "Reunión de Kick-off / 3 de junio",
        "Visita de inspección técnica / 24 de junio",
        "Habilitación empalmes eléctricos / junio – julio",
        "Instalación y habilitación equipos WES / 20 y 21 julio",
        "Capacitación equipo BUPA Plataforma WES / 31 julio",
    ]:
        _add_para(doc, hito, bullet=True, space_after=2)

    _add_subtitle(doc, "Ubicación")
    _add_table(
        doc,
        ["Recinto", "Dirección"],
        [["Clínica Bupa Antofagasta", "Manuel Antonio Matta 1945, Antofagasta"]],
    )

    # ===== II + III =====
    _page_break(doc)
    _add_title(doc, "II. Objetivos y resultados esperados")
    _add_para(
        doc,
        "El diagnóstico de gestión hídrica en Clínica Bupa Antofagasta se estructuró sobre tres ejes técnicos centrales:",
    )
    _add_para(
        doc,
        "(A) Línea de base: construir una línea de base hídrica respaldada por datos empíricos "
        "de alta frecuencia temporal.",
        bullet=True,
    )
    _add_para(
        doc,
        "(B) Perfil de consumo: elaborar perfiles diario y semanal (lun–dom) por sector, "
        "identificando variaciones en días hábiles e inhábiles.",
        bullet=True,
    )
    _add_para(
        doc,
        "(C) Detección y cuantificación: identificar ineficiencias, huecos de data y desacoples "
        "operacionales en las redes monitoreadas.",
        bullet=True,
    )

    _add_title(doc, "III. Definición de sectores a monitorear")
    _add_para(
        doc,
        "Tras visita técnica y reuniones Bupa–WES (junio–julio), y con base en facturación "
        "sanitaria y prioridades operacionales, se definieron los siguientes puntos:",
    )
    _add_para(
        doc,
        "1. Medidor Principal Sanitaria: matriz de cobre de 2\" a la entrada del recinto. "
        "Mide el 100% del volumen de agua potable entregado por Aguas Antofagasta.",
        bullet=True,
    )
    _add_para(
        doc,
        "2. Sala de Bomba Principal (Primer Piso): impulsión de cobre de 2\" en planta baja. "
        "Distribuye agua potable a gran parte de la infraestructura y alimenta acumulación "
        "hacia Sala de Bomba N°2.",
        bullet=True,
    )
    _add_para(
        doc,
        "3. Sala de Bomba N°2 (Primer Piso): impulsión en cobre de 2\" cuya función de diseño "
        "es traspasar agua hacia la sala del sexto piso. En este informe se mantiene como punto "
        "instalado, pero se excluye del balance cuantitativo por lecturas inconsistentes desde septiembre.",
        bullet=True,
    )
    _add_para(
        doc,
        "4. Sala de Bomba Sexto Piso: red de 2\" en cobre del sexto nivel; presuriza y abastece "
        "el resto del recinto.",
        bullet=True,
    )

    # ===== IV. FOTOS (pedido pág. 6) =====
    _page_break(doc)
    _add_title(doc, "IV. Fotografías de instalación")
    _add_para(
        doc,
        "Entre el 20 y 21 de julio se ejecutaron trabajos de instalación de equipos WES de "
        "Monitoreo Inteligente. A continuación, fotografías de los puntos instrumentados "
        "(fuente: Bupa Antofagasta / Fotografias).",
    )

    fotos = [
        ("Matriz Sanitaria.jpeg", "Medidor Principal Sanitaria — matriz de entrada"),
        ("Sala de bomba matriz principal.jpeg", "Sala de Bomba Principal (1er piso)"),
        ("Sala de bomna N°2.jpeg", "Sala de Bomba N°2 (1er piso)"),
        ("Sala de bomba Sexto Piso.jpeg", "Sala de Bomba Sexto Piso"),
    ]
    # Tabla 2×2 para mantener las 4 fotos en una sola página
    table = doc.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, (fname, cap) in enumerate(fotos):
        cell = table.rows[idx // 2].cells[idx % 2]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        img = FOTOS / fname
        if img.is_file():
            p.add_run().add_picture(str(img), width=Inches(2.85))
        c = cell.add_paragraph()
        c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = c.add_run(cap)
        _set_run(r, size=7, color=RGBColor(80, 80, 80))
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    _add_para(
        doc,
        "Usuarios BUPA habilitados en plataforma WES App: Catherine Velásquez (Supervisora de "
        "Proyectos e Infraestructura), Joel Cartes (Jefe de Proyecto e Infraestructura) y "
        "Jonathan Chaparro (Subgerente de Operaciones).",
        size=9,
    )

    # ===== V. ANÁLISIS (págs 7-8) =====
    _page_break(doc)
    _add_title(doc, "V. Análisis de sectores monitoreados")
    _add_para(
        doc,
        f"Consumos en m³ según data WES del periodo {_fmt_fecha(START)} al {_fmt_fecha(END)}. "
        "Gráficos propios (sin pantallazos de la aplicación).",
    )

    # --- Sanitaria ---
    _add_subtitle(doc, "1) Medidor Principal Sanitaria")
    _add_table(
        doc,
        ["Indicador", "Valor"],
        [
            ["Total (m³)", _fmt_m3(san["total"])],
            ["Promedio diario (m³)", _fmt_m3(san["prom"])],
            ["Días con data", f"{san['n']} / {NUM_DIAS}"],
            [
                "Rango min – max diario (m³)",
                f"{_fmt_m3(san['mn'])} – {_fmt_m3(san['mx'])} "
                f"({_fmt_fecha(san['dmin'])} / {_fmt_fecha(san['dmax'])})",
            ],
            ["% sobre Sanitaria", "100,0%"],
        ],
    )
    _add_picture(
        doc,
        CHARTS / "chart_diario_sanitaria.png",
        6.0,
        "Perfil de consumo diario (m³) — Medidor Sanitaria · Fuente: WES",
    )
    _add_para(
        doc,
        f"Patrón con picos hasta {_fmt_m3(san['mx'])} m³/día y mínimos del orden de "
        f"{_fmt_m3(san['mn'])} m³/día. El promedio del periodo es {_fmt_m3(san['prom'])} m³/día.",
        highlight=True,
        bullet=True,
    )
    _add_picture(
        doc,
        CHARTS / "chart_dow_sanitaria.png",
        5.8,
        "Promedio lun–dom (m³) — Medidor Sanitaria · Fuente: WES",
    )
    _add_para(
        doc,
        "En el promedio lun–dom se observa menor consumo relativo en fin de semana, "
        "consistente con menor afluencia de público y de servicios.",
        highlight=True,
        bullet=True,
    )

    # --- Principal ---
    _page_break(doc)
    _add_subtitle(doc, "2) Sala de Bomba Principal (Primer Piso)")
    _add_table(
        doc,
        ["Indicador", "Valor"],
        [
            ["Total (m³)", _fmt_m3(prin["total"])],
            ["Promedio diario (m³)", _fmt_m3(prin["prom"])],
            ["Días con data", f"{prin['n']} / {NUM_DIAS}"],
            [
                "Rango min – max diario (m³)",
                f"{_fmt_m3(prin['mn'])} – {_fmt_m3(prin['mx'])} "
                f"({_fmt_fecha(prin['dmin'])} / {_fmt_fecha(prin['dmax'])})",
            ],
            ["% sobre Sanitaria", f"{_fmt_m3(pct_prin)}%"],
        ],
    )
    _add_picture(
        doc,
        CHARTS / "chart_diario_bomba_principal.png",
        6.0,
        "Perfil de consumo diario (m³) — Bomba Principal · Fuente: WES",
    )
    _add_picture(
        doc,
        CHARTS / "chart_dow_bomba_principal.png",
        5.8,
        "Promedio lun–dom (m³) — Bomba Principal · Fuente: WES",
    )

    # --- Sexto ---
    _page_break(doc)
    _add_subtitle(doc, "3) Sala de Bomba Sexto Piso")
    _add_table(
        doc,
        ["Indicador", "Valor"],
        [
            ["Total (m³)", _fmt_m3(sexto["total"])],
            ["Promedio diario (m³)", _fmt_m3(sexto["prom"])],
            ["Días con data", f"{sexto['n']} / {NUM_DIAS}"],
            ["Hueco sin registro", "14/08/2026 – 17/08/2026 (lluvias / conectividad)"],
            [
                "Rango min – max diario (m³)",
                f"{_fmt_m3(sexto['mn'])} – {_fmt_m3(sexto['mx'])} "
                f"({_fmt_fecha(sexto['dmin'])} / {_fmt_fecha(sexto['dmax'])})",
            ],
            ["% sobre Sanitaria", f"{_fmt_m3(pct_sexto)}%"],
        ],
    )
    _add_picture(
        doc,
        CHARTS / "chart_diario_sexto_piso.png",
        6.0,
        "Perfil de consumo diario (m³) — Sexto Piso · Fuente: WES",
    )
    _add_picture(
        doc,
        CHARTS / "chart_dow_sexto_piso.png",
        5.8,
        "Promedio lun–dom (m³) — Sexto Piso · Fuente: WES",
    )

    _add_subtitle(doc, "4) Sala de Bomba N°2 — nota operativa")
    _add_para(
        doc,
        "El punto está instalado y fotografiado; sin embargo, desde septiembre las lecturas "
        "presentan inconsistencias que impiden usarlo en el balance primario del periodo. "
        "Se recomienda validación en terreno y, de confirmarse el desvío, recalibración o "
        "reemplazo del sensor antes de reincorporarlo al análisis.",
    )

    # Participación
    _add_subtitle(doc, "Participación sobre Sanitaria (periodo)")
    _add_table(
        doc,
        ["Punto", "Consumo (m³)", "% sobre Sanitaria"],
        [
            ["Medidor Principal Sanitaria", _fmt_m3(san["total"]), "100,0%"],
            ["Sala de Bomba Principal", _fmt_m3(prin["total"]), f"{_fmt_m3(pct_prin)}%"],
            ["Sala de Bomba Sexto Piso", _fmt_m3(sexto["total"]), f"{_fmt_m3(pct_sexto)}%"],
            ["Principal + Sexto", _fmt_m3(tot_salas), f"{_fmt_m3(pct_salas)}%"],
            ["No monitoreado (brecha)", _fmt_m3(gap), f"{_fmt_m3(pct_gap)}%"],
        ],
    )
    _add_picture(doc, CHARTS / "chart_participacion.png", 3.6, "Distribución % vs Sanitaria")

    # ===== VI. FACTURACIÓN (pedido pág. 13) =====
    _page_break(doc)
    _add_title(doc, "VI. Facturación sanitaria — Aguas Antofagasta")
    _add_para(
        doc,
        "Consolidación de las boletas del cliente Centro Médico Antofagasta S.A. "
        "(N° cliente 1110080842 · medidor 534 · diámetro 50 mm · dirección Manuel Antonio Matta 1957). "
        "Fuente: carpeta Bupa Antofagasta / Facturacion (4 PDF).",
    )
    _add_picture(
        doc,
        CHARTS / "chart_facturas_m3.png",
        5.5,
        "Consumo facturado por ciclo (m³)",
    )
    rows = []
    for f in FACTURAS:
        rows.append(
            [
                f["ciclo"],
                f["emision"],
                _fmt_m3(f["m3"], 0),
                _fmt_clp(f["clp"]),
                f["nro"],
            ]
        )
    rows.append(
        [
            "Total",
            "—",
            _fmt_m3(tot_fact_m3, 0),
            _fmt_clp(tot_fact_clp),
            "—",
        ]
    )
    _add_table(
        doc,
        ["Periodo", "Emisión", "Consumo (m³)", "Monto facturado", "N° factura"],
        rows,
        compact=True,
    )
    _add_subtitle(doc, "Lecturas encadenadas del medidor 534")
    lect_rows = [
        [
            f["ciclo"],
            str(f["lect_ant"]),
            str(f["lect_act"]),
            _fmt_m3(f["m3"], 0),
            f["nro"],
        ]
        for f in FACTURAS
    ]
    _add_table(
        doc,
        ["Ciclo", "Lectura anterior", "Lectura actual", "m³", "N° factura"],
        lect_rows,
        compact=True,
    )
    _add_para(
        doc,
        "No hay cambio de medidor, m³ por abonar en la factura siguiente ni lectura estimada: "
        "el salto de volumen está en el medidor. Tarifa publicada (El Mercurio de Antofagasta, "
        "15-11-2025) implica del orden de $2.769 CLP/m³ con IVA (agua + alcantarillado + tratamiento + cargo fijo prorrateado).",
    )

    # ===== VII. CONCLUSIONES =====
    _page_break(doc)
    _add_title(doc, "VII. Conclusiones y hallazgos")
    _add_para(
        doc,
        f"1) Periodo WES {_fmt_fecha(START)}–{_fmt_fecha(END)}: Sanitaria {_fmt_m3(san['total'])} m³ "
        f"({_fmt_m3(san['prom'])} m³/día).",
        bullet=True,
    )
    _add_para(
        doc,
        f"2) Principal + Sexto = {_fmt_m3(pct_salas)}% de Sanitaria; brecha no monitoreada = "
        f"{_fmt_m3(pct_gap)}% ({_fmt_m3(gap)} m³).",
        bullet=True,
    )
    _add_para(
        doc,
        "3) Sexto Piso: único hueco de data 14–17/08/2026 (asociado a lluvias / conectividad). "
        "Sanitaria y Principal sin huecos en el periodo.",
        bullet=True,
    )
    _add_para(
        doc,
        "4) Sala N°2 queda fuera del balance cuantitativo hasta validación/recalibración.",
        bullet=True,
    )
    _add_para(
        doc,
        f"5) Facturación jun–sep 2026: {_fmt_m3(tot_fact_m3, 0)} m³ / {_fmt_clp(tot_fact_clp)} "
        "(4 ciclos Aguas Antofagasta, lecturas encadenadas).",
        bullet=True,
    )

    # ===== VIII. RECOMENDACIONES =====
    _add_title(doc, "VIII. Recomendaciones y próximos pasos")
    _add_para(
        doc,
        "Acción 1 — Cerrar brecha de visibilidad: priorizar instrumentación de tramos no medidos "
        "y resolver el estado de Sala N°2.",
        bullet=True,
    )
    _add_para(
        doc,
        "Acción 2 — Control en horarios inhábiles: evaluar Control Inteligente WES donde el "
        "consumo nocturno/fin de semana no se justifique operativamente.",
        bullet=True,
    )
    _add_para(
        doc,
        "Acción 3 — Contraste factura vs WES: al cierre de cada ciclo de Aguas Antofagasta, "
        "comparar m³ boleta (medidor 534) con Sanitaria WES y documentar diferencias.",
        bullet=True,
    )
    _add_para(
        doc,
        "Acción 4 — Mantener usuarios BUPA activos en la App y revisión periódica de alertas "
        "con el equipo de infraestructura.",
        bullet=True,
    )

    _add_para(
        doc,
        f"Documento generado {datetime.now().strftime('%d/%m/%Y %H:%M')} · WES · Clínica Bupa Antofagasta",
        size=8,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_before=18,
    )

    out = ROOT / f"Informe_final_WES_Bupa_Antofagasta_{START.strftime('%Y%m%d')}_{END.strftime('%Y%m%d')}.docx"
    doc.save(str(out))
    return out


if __name__ == "__main__":
    path = build()
    print(f"[OK] {path}")
