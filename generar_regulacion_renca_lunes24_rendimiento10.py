"""
Regulación Renca desde el lunes 24/08/2026 (vuelve el periodo con control).

Objetivo: ~10 % de rendimiento vs el perfil sin WES (17–20 ago).
Programación WES en tres niveles sobre un caudal máximo (100 %):
  100 % = tope m³/h a cargar en el horario
   60 % = 0,60 × ese tope
   30 % = 0,30 × ese tope

Semáforo (vs el pico del día de línea base):
  ≥ 74 % del pico → 100 %   |  40–74 % → 60 %   |  < 40 % → 30 %

Uso:
  python generar_regulacion_renca_lunes24_rendimiento10.py
"""

from __future__ import annotations

import statistics
import subprocess
import shutil
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
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
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from auditoria_cpa_icco_renca_grafico import _vector_m3h_24_desde_api
from generar_reporte_word import add_logo_to_header, format_number_chilean

ROOT = Path(__file__).resolve().parent
TZ_CL = ZoneInfo("America/Santiago")
OBJETIVO_PCT = 10.0  # rendimiento = (sin − con) / sin
FACTOR_OBJ = 1.0 - OBJETIVO_PCT / 100.0
RATIO_100 = 0.74
RATIO_60 = 0.40
COLOR_HEAD = RGBColor(31, 71, 136)
FILL_100 = PatternFill("solid", fgColor="C6EFCE")
FILL_60 = PatternFill("solid", fgColor="FFF2CC")
FILL_30 = PatternFill("solid", fgColor="FFC7CE")
COLOR_100 = "#548235"
COLOR_60 = "#BF8F00"
COLOR_30 = "#C0504D"

PUNTOS: Tuple[Tuple[str, str], ...] = (
    ("000017-07", "ICCP (Cumbre de Cóndores pte.)"),
    ("000017-04", "Esc. Lo Velásquez"),
    ("000017-05", "Gimnasio municipal"),
    ("000017-06", "Piscina municipal"),
)
DIAS_SIN = (date(2026, 8, 17), date(2026, 8, 18), date(2026, 8, 19), date(2026, 8, 20))
LUNES_CONTROL = date(2026, 8, 24)


def _m3_a_lmin(m3h: float) -> float:
    return float(m3h) * 1000.0 / 60.0


def _nivel(v: float, pico: float) -> int:
    if pico <= 1e-9 or v <= 1e-9:
        return 30
    r = v / pico
    if r >= RATIO_100:
        return 100
    if r >= RATIO_60:
        return 60
    return 30


def _consumo_esperado(q100: float, sin_h: Sequence[float], fracs: Sequence[float]) -> float:
    return float(sum(min(float(sin_h[h]), q100 * fracs[h]) for h in range(24)))


def _q100_para_objetivo(sin_h: Sequence[float], fracs: Sequence[float], factor: float) -> float:
    """Busca el 100 % (m³/h) tal que el tope 30/60/100 deje factor × Σ sin WES."""
    base = float(sum(sin_h))
    if base <= 1e-9:
        return 0.0
    target = factor * base
    lo, hi = 0.0, max(max(sin_h), 0.01) * 8.0
    # Si ni con tope enorme se llega al target (el patrón 30/60 recorta de más), devolver hi.
    if _consumo_esperado(hi, sin_h, fracs) + 1e-9 < target:
        return hi
    for _ in range(48):
        mid = (lo + hi) / 2.0
        got = _consumo_esperado(mid, sin_h, fracs)
        if got >= target:
            hi = mid
        else:
            lo = mid
    return hi


def _vector_mediana(vecs: Sequence[Sequence[float]]) -> List[float]:
    out: List[float] = []
    for h in range(24):
        vals = [float(v[h]) for v in vecs]
        out.append(float(statistics.median(vals)))
    return out


def _set_shading(cell, hex_fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in tc_pr.findall(qn("w:shd")):
        tc_pr.remove(old)
    tc_pr.append(
        parse_xml(
            "<w:shd xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
            f'w:val="clear" w:fill="{hex_fill}"/>'
        )
    )


def _grafico_dia(
    sin_h: Sequence[float],
    cap_h: Sequence[float],
    niveles: Sequence[int],
    titulo: str,
    out_png: Path,
) -> None:
    x = list(range(24))
    colors = [{100: COLOR_100, 60: COLOR_60, 30: COLOR_30}[n] for n in niveles]
    fig, ax = plt.subplots(figsize=(12.2, 4.8))
    ax.bar([i - 0.18 for i in x], list(sin_h), width=0.36, color="#C0504D", label="Sin WES (línea base)", zorder=2)
    ax.bar([i + 0.18 for i in x], list(cap_h), width=0.36, color=colors, label="Tope 30/60/100 %", zorder=2)
    ax.set_xticks(x)
    ax.set_xticklabels([f"{h:02d}" for h in x], fontsize=7)
    ax.set_ylabel("m³/h")
    ax.set_title(titulo, fontweight="bold", fontsize=11, color="#1F4788")
    ax.grid(axis="y", alpha=0.3, zorder=0)
    ax.legend(fontsize=8, loc="upper right")
    fig.tight_layout()
    fig.savefig(out_png, dpi=150, bbox_inches="tight")
    plt.close(fig)


def _pdf(docx: Path) -> Optional[Path]:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    subprocess.run(
        [
            soffice,
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
            "--outdir",
            str(docx.parent),
            str(docx),
        ],
        check=False,
        capture_output=True,
        text=True,
        timeout=180,
    )
    p = docx.with_suffix(".pdf")
    return p if p.is_file() else None


def _plan_para_perfil(sin_h: List[float]) -> dict:
    pico = max(sin_h) if sin_h else 0.0
    niveles = [_nivel(sin_h[h], pico) for h in range(24)]
    fracs = [n / 100.0 for n in niveles]
    q100 = _q100_para_objetivo(sin_h, fracs, FACTOR_OBJ)
    cap = [q100 * f for f in fracs]
    esperado = _consumo_esperado(q100, sin_h, fracs)
    base = float(sum(sin_h))
    rend = ((base - esperado) / base * 100.0) if base > 1e-9 else 0.0
    return {
        "sin": sin_h,
        "pico": pico,
        "niveles": niveles,
        "q100": q100,
        "q60": q100 * 0.60,
        "q30": q100 * 0.30,
        "cap": cap,
        "base": base,
        "esperado": esperado,
        "rend": rend,
    }


def _excel(path: Path, planes_lun: Dict[str, dict], planes_tipo: Dict[str, dict]) -> None:
    wb = Workbook()
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    head = Font(bold=True, color="FFFFFF", name="Calibri")
    fill_h = PatternFill("solid", fgColor="1F4788")

    def _hdr(ws, row, vals):
        for i, v in enumerate(vals, 1):
            c = ws.cell(row, i, v)
            c.font = head
            c.fill = fill_h
            c.alignment = Alignment(wrap_text=True, horizontal="center")
            c.border = thin

    ws = wb.active
    ws.title = "Resumen_100_60_30"
    _hdr(
        ws,
        1,
        [
            "Punto",
            "Perfil",
            "Σ sin WES (m³)",
            "Σ con tope 10% (m³)",
            "Rendimiento %",
            "100 % (m³/h)",
            "60 % (m³/h)",
            "30 % (m³/h)",
            "100 % (L/min)",
            "60 % (L/min)",
            "30 % (L/min)",
            "Horas 100",
            "Horas 60",
            "Horas 30",
        ],
    )
    r = 2
    for nid, nom in PUNTOS:
        for etiqueta, plan in (("Lunes 24 (base lun 17)", planes_lun[nid]), ("Día tipo lun–jue (mediana)", planes_tipo[nid])):
            n = plan["niveles"]
            ws.cell(r, 1, nom)
            ws.cell(r, 2, etiqueta)
            ws.cell(r, 3, round(plan["base"], 2))
            ws.cell(r, 4, round(plan["esperado"], 2))
            ws.cell(r, 5, round(plan["rend"], 1))
            ws.cell(r, 6, round(plan["q100"], 3))
            ws.cell(r, 7, round(plan["q60"], 3))
            ws.cell(r, 8, round(plan["q30"], 3))
            ws.cell(r, 9, round(_m3_a_lmin(plan["q100"]), 2))
            ws.cell(r, 10, round(_m3_a_lmin(plan["q60"]), 2))
            ws.cell(r, 11, round(_m3_a_lmin(plan["q30"]), 2))
            ws.cell(r, 12, n.count(100))
            ws.cell(r, 13, n.count(60))
            ws.cell(r, 14, n.count(30))
            for c in range(1, 15):
                ws.cell(r, c).border = thin
            r += 1

    ws2 = wb.create_sheet("Horario_lunes_24")
    _hdr(
        ws2,
        1,
        ["Hora"]
        + [f"{nom} % / m³/h / L·min" for _, nom in PUNTOS],
    )
    # Better: one block per point
    ws2.delete_rows(1)
    col = 1
    for nid, nom in PUNTOS:
        plan = planes_lun[nid]
        ws2.cell(1, col, nom).font = head
        ws2.cell(1, col).fill = fill_h
        ws2.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 3)
        for i, h in enumerate(["Hora", "% WES", "Tope m³/h", "Tope L/min"]):
            c = ws2.cell(2, col + i, h)
            c.font = head
            c.fill = fill_h
            c.border = thin
        for h in range(24):
            rr = 3 + h
            n = plan["niveles"][h]
            fill = {100: FILL_100, 60: FILL_60, 30: FILL_30}[n]
            vals = [f"{h:02d}:00–{h:02d}:59", n, round(plan["cap"][h], 3), round(_m3_a_lmin(plan["cap"][h]), 2)]
            for i, v in enumerate(vals):
                c = ws2.cell(rr, col + i, v)
                c.border = thin
                c.fill = fill
        col += 5

    ws3 = wb.create_sheet("Horario_dia_tipo")
    col = 1
    for nid, nom in PUNTOS:
        plan = planes_tipo[nid]
        ws3.cell(1, col, nom).font = head
        ws3.cell(1, col).fill = fill_h
        ws3.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 3)
        for i, h in enumerate(["Hora", "% WES", "Tope m³/h", "Tope L/min"]):
            c = ws3.cell(2, col + i, h)
            c.font = head
            c.fill = fill_h
            c.border = thin
        for h in range(24):
            rr = 3 + h
            n = plan["niveles"][h]
            fill = {100: FILL_100, 60: FILL_60, 30: FILL_30}[n]
            vals = [f"{h:02d}:00–{h:02d}:59", n, round(plan["cap"][h], 3), round(_m3_a_lmin(plan["cap"][h]), 2)]
            for i, v in enumerate(vals):
                c = ws3.cell(rr, col + i, v)
                c.border = thin
                c.fill = fill
        col += 5

    ws4 = wb.create_sheet("Criterio")
    lines = [
        ("Objetivo", f"{OBJETIVO_PCT:.0f} % de rendimiento vs sin WES (consumo esperado = {FACTOR_OBJ:.0%} de la línea base)."),
        ("Inicio control", "Lunes 24/08/2026"),
        ("Línea base", "Sin WES 17–20/08/2026 (hora Chile). Lunes 24 usa el perfil del lunes 17."),
        ("Día tipo", "Mediana hora a hora de lun–jue 17–20 (el miércoles anómalo de ICCP no arrastra el máximo)."),
        ("100 %", "Caudal máximo a programar (m³/h). Es el tope de las horas pico."),
        ("60 % / 30 %", "0,60 y 0,30 de ese mismo 100 %."),
        ("Semáforo", f"vs pico del perfil: ≥{RATIO_100:.0%} → 100 %; {RATIO_60:.0%}–{RATIO_100:.0%} → 60 %; <{RATIO_60:.0%} o ~0 → 30 %."),
        ("Tope", "La válvula limita; no fuerza consumo. El 10 % se cumple si la demanda se parece a la línea base."),
    ]
    for i, (a, b) in enumerate(lines, 1):
        ws4.cell(i, 1, a).font = Font(bold=True)
        ws4.cell(i, 2, b)
    ws4.column_dimensions["A"].width = 18
    ws4.column_dimensions["B"].width = 92

    for wsx in wb.worksheets:
        for col in wsx.columns:
            letter = get_column_letter(col[0].column)
            wsx.column_dimensions[letter].width = max(12, min(28, wsx.column_dimensions[letter].width or 12))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _word(
    path: Path,
    hora_corte: int,
    pngs: Dict[str, Path],
    planes_lun: Dict[str, dict],
    planes_tipo: Dict[str, dict],
) -> None:
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    try:
        add_logo_to_header(doc)
    except Exception:
        pass
    h = doc.add_heading("Renca — regulación para el lunes 24/08 (10 % de rendimiento)", level=0)
    if h.runs:
        h.runs[0].font.color.rgb = COLOR_HEAD
    p = doc.add_paragraph()
    p.add_run("Vuelve el periodo con control. ").bold = True
    p.add_run(
        "Línea base = sin WES 17–20/08. Objetivo = gastar ~90 % de esa línea "
        f"(10 % de rendimiento). Jueves 20 incluido hasta {hora_corte:02d}:59 Chile. "
        "En WES se carga un caudal máximo (100 %) y el horario usa solo 30 %, 60 % o 100 % de ese máximo."
    )

    doc.add_heading("Cómo programar", level=1)
    doc.add_paragraph(
        "1) Cargar el 100 % = m³/h (o L/min) de la tabla. "
        "2) Pintar el día: verde 100 % en horas pico, amarillo 60 %, rojo 30 %. "
        "3) El 60 % y el 30 % salen solos: son 0,60 y 0,30 del mismo 100 %."
    )

    tbl = doc.add_table(rows=1 + len(PUNTOS), cols=7)
    tbl.style = "Table Grid"
    headers = ["Punto", "100 % m³/h", "60 % m³/h", "30 % m³/h", "100 % L/min", "Σ día (m³)", "Rend. %"]
    for j, hd in enumerate(headers):
        cell = tbl.rows[0].cells[j]
        cell.text = hd
        _set_shading(cell, "1F4788")
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9)
    for i, (nid, nom) in enumerate(PUNTOS, start=1):
        pl = planes_lun[nid]
        vals = [
            nom,
            format_number_chilean(pl["q100"], 2),
            format_number_chilean(pl["q60"], 2),
            format_number_chilean(pl["q30"], 2),
            format_number_chilean(_m3_a_lmin(pl["q100"]), 1),
            format_number_chilean(pl["esperado"], 1),
            format_number_chilean(pl["rend"], 1) + " %",
        ]
        for j, v in enumerate(vals):
            tbl.rows[i].cells[j].text = v
            for run in tbl.rows[i].cells[j].paragraphs[0].runs:
                run.font.size = Pt(9)
        _set_shading(tbl.rows[i].cells[1], "C6EFCE")
        _set_shading(tbl.rows[i].cells[2], "FFF2CC")
        _set_shading(tbl.rows[i].cells[3], "FFC7CE")

    doc.add_paragraph(
        "Números de la tabla = perfil del lunes 17 (el homólogo del lunes 24). "
        "La hoja «día tipo» del Excel usa la mediana lun–jue por si el miércoles 19 de ICCP no se quiere repetir."
    )

    for nid, nom in PUNTOS:
        pl = planes_lun[nid]
        doc.add_heading(nom, level=1)
        doc.add_paragraph(
            f"Línea base lun 17: {format_number_chilean(pl['base'], 1)} m³. "
            f"Con este horario: {format_number_chilean(pl['esperado'], 1)} m³ "
            f"({format_number_chilean(pl['rend'], 1)} %). "
            f"Pico medido: {format_number_chilean(pl['pico'], 2)} m³/h. "
            f"100 % a programar: {format_number_chilean(pl['q100'], 2)} m³/h "
            f"({format_number_chilean(_m3_a_lmin(pl['q100']), 1)} L/min)."
        )
        if f"lun_{nid}" in pngs:
            doc.add_picture(str(pngs[f"lun_{nid}"]), width=Cm(16.0))
        t = doc.add_table(rows=25, cols=4)
        t.style = "Table Grid"
        for j, hd in enumerate(["Hora", "%", "Tope m³/h", "Tope L/min"]):
            cell = t.rows[0].cells[j]
            cell.text = hd
            _set_shading(cell, "1F4788")
            for run in cell.paragraphs[0].runs:
                run.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(8)
        for h in range(24):
            n = pl["niveles"][h]
            hexf = {100: "C6EFCE", 60: "FFF2CC", 30: "FFC7CE"}[n]
            vals = [
                f"{h:02d}:00",
                str(n),
                format_number_chilean(pl["cap"][h], 2),
                format_number_chilean(_m3_a_lmin(pl["cap"][h]), 1),
            ]
            for j, v in enumerate(vals):
                t.rows[h + 1].cells[j].text = v
                _set_shading(t.rows[h + 1].cells[j], hexf)
                for run in t.rows[h + 1].cells[j].paragraphs[0].runs:
                    run.font.size = Pt(8)

    doc.add_heading("Notas", level=1)
    doc.add_paragraph(
        "La piscina en 17–19 ago estuvo casi plana (~56–58 m³/día). Si el perfil es plano, "
        "casi todas las horas quedan en 100 % y el 10 % sale bajando el caudal máximo "
        "(100 % ≈ 90 % del m³/h medido). ICCP no: el miércoles 19 duplicó; el lunes 24 "
        "usa el lunes 17, no ese pico."
    )
    doc.add_paragraph(
        "Escuela y gimnasio en 17–20 gastaron menos que la semana con WES. Ahí el 10 % "
        "es un tope suave para no dispararlos al reactivar control, no un recorte agresivo."
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


def main() -> int:
    ahora = datetime.now(TZ_CL)
    hora_corte = max(0, min(23, ahora.hour - 1)) if ahora.date() == date(2026, 8, 20) else 23
    if ahora.date() > date(2026, 8, 20):
        hora_corte = 23
    print(f"Chile {ahora:%Y-%m-%d %H:%M} | jueves 20 hasta {hora_corte:02d}:59")

    jobs = [(nid, d) for nid, _ in PUNTOS for d in DIAS_SIN]
    vecs: Dict[Tuple[str, date], List[float]] = {}

    def _job(item: Tuple[str, date]) -> Tuple[Tuple[str, date], List[float]]:
        nid, d = item
        v = _vector_m3h_24_desde_api(nid, d)
        if d == date(2026, 8, 20) and hora_corte < 23:
            v = [v[h] if h <= hora_corte else 0.0 for h in range(24)]
        return item, v

    print(f"Descargando {len(jobs)} series horarias…")
    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(_job, it) for it in jobs]
        for i, fut in enumerate(as_completed(futs), 1):
            k, v = fut.result()
            vecs[k] = v
            if i % 4 == 0 or i == len(jobs):
                print(f"  {i}/{len(jobs)}")

    planes_lun: Dict[str, dict] = {}
    planes_tipo: Dict[str, dict] = {}
    for nid, nom in PUNTOS:
        lun = vecs[(nid, date(2026, 8, 17))]
        tipo_src = [vecs[(nid, d)] for d in DIAS_SIN]
        tipo = _vector_mediana(tipo_src)
        planes_lun[nid] = _plan_para_perfil(lun)
        planes_tipo[nid] = _plan_para_perfil(tipo)
        pl = planes_lun[nid]
        print(
            f"  {nom}: lun17 {pl['base']:.1f} m³ | 100%={pl['q100']:.2f} m³/h "
            f"({_m3_a_lmin(pl['q100']):.1f} L/min) | esperado {pl['esperado']:.1f} m³ "
            f"({pl['rend']:.1f} %) | h100={pl['niveles'].count(100)} "
            f"h60={pl['niveles'].count(60)} h30={pl['niveles'].count(30)}"
        )

    ts = ahora.strftime("%Y%m%d_%H%M")
    out_dir = ROOT / "reports" / "reporte de auditoria" / f"regulacion_renca_lunes24_rend10_{ts}"
    png_dir = out_dir / "graficos"
    png_dir.mkdir(parents=True, exist_ok=True)
    pngs: Dict[str, Path] = {}
    for nid, nom in PUNTOS:
        p = png_dir / f"lunes24_{nid}.png"
        pl = planes_lun[nid]
        _grafico_dia(
            pl["sin"],
            pl["cap"],
            pl["niveles"],
            f"{nom} — lunes 24: sin WES lun 17 vs tope 30/60/100 (10 %)",
            p,
        )
        pngs[f"lun_{nid}"] = p

    xlsx = out_dir / "Regulacion_Renca_lunes24_rendimiento10.xlsx"
    _excel(xlsx, planes_lun, planes_tipo)
    docx = out_dir / f"Regulacion_Renca_lunes24_rendimiento10_{ts}.docx"
    _word(docx, hora_corte, pngs, planes_lun, planes_tipo)
    pdf = _pdf(docx)
    print(f"XLSX {xlsx}")
    print(f"DOCX {docx}")
    print(f"PDF  {pdf or '(no convertido)'}")
    return 0


if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except Exception:
            pass
    raise SystemExit(main())
