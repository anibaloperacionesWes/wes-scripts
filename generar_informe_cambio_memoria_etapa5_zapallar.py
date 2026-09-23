"""
Informe corto — cambio de memoria placa / Etapa N°5 Fundo Zapallar.

Validación lectura mecánica vs consumo (API WES + hueco leído de placa):
  - Inicio: lectura 22/09/2026 14:30 → mitad del consumo hora 14.
  - Fin: lectura 23/09/2026 16:54 → consumo hasta las 16:00.

Uso:
  python generar_informe_cambio_memoria_etapa5_zapallar.py
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "reports" / "Fundo_Zapallar" / "Informes_Tecnicos"
FOTOS_DIR = OUT_DIR / "fotos_etapa5"

try:
    from generar_reporte_word import CHILE_TZ, acl_node_base_url
except Exception:
    CHILE_TZ = ZoneInfo("America/Santiago")

    def acl_node_base_url() -> str:
        return "http://104.248.53.141:7003/wes/api/acl-node/v1"

COMPANY = "Fundo Zapallar"
COMPANY_ID = "000027"
NODE_ID = "000027-03"
NODE_NAME = "Etapa N°5"
FECHA_INTERVENCION = date(2026, 9, 22)
DIAMETRO = "DN90 fierro dúctil"
CAUDAL_MAX_REF_M3H = 60.0
COLOR_TITULO = RGBColor(31, 71, 136)

# Lecturas mecánicas (fotos terreno)
LECTURA_AYER_M3 = 5144.0
LECTURA_AYER_DT = datetime(2026, 9, 22, 14, 30, tzinfo=CHILE_TZ)
LECTURA_HOY_M3 = 5177.0
LECTURA_HOY_DT = datetime(2026, 9, 23, 16, 54, tzinfo=CHILE_TZ)
WES_HASTA_DT = datetime(2026, 9, 23, 16, 0, tzinfo=CHILE_TZ)

# Hueco 22/09 leído directo de la placa (m³/h).
# Extracción listó H:15:00 dos veces (0,60 y 2,90): se interpreta 0,60 como hora 14
# (coherente con lectura a las 14:30) y 2,90 como hora 15.
PLACA_HUECO_22: Dict[int, float] = {
    14: 0.60,
    15: 2.90,
    16: 5.50,
    17: 2.60,
    18: 0.90,
    19: 0.00,
    20: 0.00,
    21: 0.00,
    22: 0.00,
    23: 0.00,
}

MEDIDOR_MARCA = "Sensus"
MEDIDOR_MODELO = "MeiStream Plus 100"
MEDIDOR_SERIE = "8 SEN01 2370 9030"
MEDIDOR_Q3_M3H = 100.0
HRI_MODELO = "HRI-Mei B4 500 ms"
HRI_SERIE = "31730485"
HRI_PULSO_DN40_125_L = 100

FOTO_AYER = FOTOS_DIR / "lectura_20260922_1430_sensus_5144.png"
FOTO_HOY = FOTOS_DIR / "lectura_20260923_1654_sensus_5177.png"


@dataclass
class Validacion:
    lectura_ayer: float
    lectura_hoy: float
    delta_mecanico: float
    wes_m3: float
    detalle: List[Tuple[str, float, str]]
    hueco_horas_api: List[str]
    diferencia_m3: float
    diferencia_pct: float
    estado: str


def _set_run_font(run, *, bold: bool = False, size: int = 11, color: RGBColor | None = None) -> None:
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    if color is not None:
        run.font.color.rgb = color


def _add_bullet(doc: Document, text: str) -> None:
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(text)
    _set_run_font(run, size=11)


def _shade_cell(cell, hex_color: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def _fmt(n: float, dec: int = 1) -> str:
    return f"{n:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def obtener_api_por_hora() -> Dict[datetime, float]:
    base = acl_node_base_url()
    out: Dict[datetime, float] = {}
    for d in (date(2026, 9, 21), date(2026, 9, 22), date(2026, 9, 23)):
        ds = d.strftime("%d%m%Y")
        url = f"{base}/nodes/{NODE_ID}/dates.measures.csv"
        r = requests.get(url, params=[("start", ds), ("end", ds)], timeout=45)
        r.raise_for_status()
        lines = [ln for ln in r.text.strip().splitlines() if ln.strip()]
        start = 1 if lines and "TIME" in lines[0].upper() else 0
        for line in lines[start:]:
            parts = line.split(",")
            if len(parts) < 2:
                continue
            raw = parts[0].strip().replace("Z", "+00:00")
            dt = datetime.fromisoformat(raw)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            out[dt.astimezone(CHILE_TZ)] = float(parts[1])
    return out


def valor_hora(dt: datetime, api: Dict[datetime, float]) -> Tuple[float, str]:
    """Prioriza valores de placa en el hueco del 22/09; si no, API."""
    if dt.date() == date(2026, 9, 22) and dt.hour in PLACA_HUECO_22:
        return PLACA_HUECO_22[dt.hour], "placa (hueco)"
    if dt in api:
        return api[dt], "API WES"
    return 0.0, "sin dato (0)"


def calcular_validacion(lectura_ayer: float, lectura_hoy: float) -> Validacion:
    api = obtener_api_por_hora()
    detalle: List[Tuple[str, float, str]] = []
    total = 0.0
    hueco_api: List[str] = []

    inicio_hora = LECTURA_AYER_DT.replace(minute=0, second=0, microsecond=0)
    v14, src14 = valor_hora(inicio_hora, api)
    contrib14 = 0.5 * v14
    detalle.append(
        (
            inicio_hora.strftime("%d/%m/%Y %H:%M"),
            contrib14,
            f"mitad hora 14 (lectura 14:30); bruto {_fmt(v14, 2)} m³/h [{src14}]",
        )
    )
    total += contrib14
    if "API" not in src14 and inicio_hora not in api:
        hueco_api.append(inicio_hora.strftime("%d/%m/%Y %H:%M"))

    cur = inicio_hora + timedelta(hours=1)
    while cur < WES_HASTA_DT:
        v, src = valor_hora(cur, api)
        detalle.append((cur.strftime("%d/%m/%Y %H:%M"), v, src))
        total += v
        if src.startswith("placa") or src.startswith("sin"):
            if cur not in api:
                hueco_api.append(cur.strftime("%d/%m/%Y %H:%M"))
        cur += timedelta(hours=1)

    delta = lectura_hoy - lectura_ayer
    dif = delta - total
    pct = (dif / delta * 100.0) if delta else 0.0
    if abs(pct) <= 5:
        estado = "OK — diferencia ≤ 5 %"
    elif abs(pct) <= 15:
        estado = "ACEPTABLE — diferencia moderada (hueco placa + corte 16:00 vs foto 16:54)"
    else:
        estado = "REVISAR — diferencia relevante"

    return Validacion(
        lectura_ayer=lectura_ayer,
        lectura_hoy=lectura_hoy,
        delta_mecanico=round(delta, 2),
        wes_m3=round(total, 2),
        detalle=detalle,
        hueco_horas_api=hueco_api,
        diferencia_m3=round(dif, 2),
        diferencia_pct=round(pct, 1),
        estado=estado,
    )


def build_doc(lectura_ayer: float, lectura_hoy: float) -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    ahora = datetime.now(CHILE_TZ)
    stamp = ahora.strftime("%Y%m%d_%H%M")
    out = OUT_DIR / f"Informe_Cambio_Memoria_Etapa5_Zapallar_{stamp}.docx"
    val = calcular_validacion(lectura_ayer, lectura_hoy)

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("INFORME TÉCNICO CORTO")
    _set_run_font(r, bold=True, size=16, color=COLOR_TITULO)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run(
        f"Cambio de memoria de placa — {NODE_NAME} ({NODE_ID})\n"
        f"{COMPANY} ({COMPANY_ID})"
    )
    _set_run_font(r, bold=True, size=12, color=COLOR_TITULO)

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = meta.add_run(
        f"Fecha intervención: {FECHA_INTERVENCION.strftime('%d/%m/%Y')}  ·  "
        f"Generado: {ahora.strftime('%d/%m/%Y %H:%M')} (Chile)"
    )
    _set_run_font(r, size=10, color=RGBColor(90, 90, 90))

    h = doc.add_heading("1. Resumen", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO
    p = doc.add_paragraph()
    r = p.add_run(
        "En terreno se detectaron lecturas de caudal anómalas (pulsos del orden de "
        "92 y 200 m³/h) en la matriz de Etapa N°5. Tras descartar falla del sensor "
        "y de la cadena de pulsos, y verificar voltajes correctos, se concluyó que "
        "el error provenía de la memoria de la placa. Se reemplazó la memoria para "
        "normalizar el punto y evitar recurrencia. Se valida además el consumo con "
        "lecturas mecánicas del medidor Sensus versus serie WES/placa."
    )
    _set_run_font(r, size=11)

    h = doc.add_heading("2. Revisión en terreno", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO
    _add_bullet(doc, "Sensor Census / sensor inductivo: probado y en buen estado.")
    _add_bullet(doc, "Pulsos de revisión: realizados; respuesta correcta.")
    _add_bullet(doc, "Voltajes de alimentación / alimentación de placa: correctos.")
    _add_bullet(doc, "Diagnóstico: memoria de la placa arrojaba el error (lecturas irreales).")
    _add_bullet(doc, "Acción correctiva: cambio de memoria de la placa; punto reparado/normalizado.")

    h = doc.add_heading("3. Medidor y lecturas mecánicas", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO
    _add_bullet(doc, f"Marca/modelo: {MEDIDOR_MARCA} {MEDIDOR_MODELO}")
    _add_bullet(doc, f"Serie medidor: {MEDIDOR_SERIE} (2023)")
    _add_bullet(doc, f"Q3 medidor: {_fmt(MEDIDOR_Q3_M3H, 0)} m³/h")
    _add_bullet(doc, f"Módulo pulsos: {HRI_MODELO}, serie {HRI_SERIE}")
    _add_bullet(
        doc,
        f"Peso de pulso DN 40–125: {HRI_PULSO_DN40_125_L} L/pulso "
        f"(={HRI_PULSO_DN40_125_L/1000:.1f} m³/pulso), aplicable a DN90",
    )
    _add_bullet(
        doc,
        f"Lectura ayer: {_fmt(lectura_ayer, 0)} m³ "
        f"({LECTURA_AYER_DT.strftime('%d/%m/%Y %H:%M')} Chile)",
    )
    _add_bullet(
        doc,
        f"Lectura hoy: {_fmt(lectura_hoy, 0)} m³ "
        f"({LECTURA_HOY_DT.strftime('%d/%m/%Y %H:%M')} Chile)",
    )
    _add_bullet(doc, f"Δ mecánico: {_fmt(val.delta_mecanico, 0)} m³")

    if FOTO_AYER.is_file():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(FOTO_AYER), width=Inches(4.8))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = cap.add_run(f"Foto ayer — {_fmt(lectura_ayer, 0)} m³ ({LECTURA_AYER_DT.strftime('%d/%m/%Y %H:%M')})")
        _set_run_font(r, size=9, color=RGBColor(90, 90, 90))

    if FOTO_HOY.is_file():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(FOTO_HOY), width=Inches(4.8))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = cap.add_run(f"Foto hoy — {_fmt(lectura_hoy, 0)} m³ ({LECTURA_HOY_DT.strftime('%d/%m/%Y %H:%M')})")
        _set_run_font(r, size=9, color=RGBColor(90, 90, 90))

    h = doc.add_heading("4. Criterio hidráulico (DN90)", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO
    p = doc.add_paragraph()
    r = p.add_run(
        f"La red es {DIAMETRO}. Techo práctico de red ≈ {CAUDAL_MAX_REF_M3H:.0f} m³/h "
        f"(~2,5 m/s). El medidor admite Q3 = {_fmt(MEDIDOR_Q3_M3H, 0)} m³/h, pero la red "
        "no puede entregar de forma realista 92–200 m³/h (error de memoria)."
    )
    _set_run_font(r, size=11)

    table = doc.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    filas = [
        ("Velocidad de referencia", "Caudal teórico DN90"),
        ("1,5 m/s", "≈ 34 m³/h"),
        ("2,0 m/s", "≈ 46 m³/h"),
        ("2,5 m/s (techo práctico)", "≈ 57–60 m³/h"),
    ]
    for i, (a, b) in enumerate(filas):
        table.rows[i].cells[0].text = a
        table.rows[i].cells[1].text = b
        for cell in table.rows[i].cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    _set_run_font(run, bold=(i == 0), size=10)
        if i == 0:
            for cell in table.rows[i].cells:
                _shade_cell(cell, "1F4788")
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(255, 255, 255)

    h = doc.add_heading("5. Validación lecturas vs WES/placa", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO

    p = doc.add_paragraph()
    r = p.add_run(
        "Ventana: lectura ayer 14:30 → mitad del consumo hora 14; lectura hoy "
        f"(foto {LECTURA_HOY_DT.strftime('%H:%M')}) → consumo hasta las "
        f"{WES_HASTA_DT.strftime('%H:%M')}. El hueco de API del 22/09 (tarde) se "
        "completa con valores leídos directo de la placa en terreno."
    )
    _set_run_font(r, size=11)

    p = doc.add_paragraph()
    r = p.add_run(
        "Nota de extracción placa: H:15:00 apareció dos veces (0,60 y 2,90). "
        "Se interpreta 0,60 como hora 14:00 y 2,90 como hora 15:00."
    )
    _set_run_font(r, size=9, color=RGBColor(100, 80, 0))

    rows_data = [
        ("Concepto", "Valor"),
        ("Lectura ayer", f"{_fmt(val.lectura_ayer, 0)} m³ @ {LECTURA_AYER_DT.strftime('%d/%m/%Y %H:%M')}"),
        ("Lectura hoy", f"{_fmt(val.lectura_hoy, 0)} m³ @ {LECTURA_HOY_DT.strftime('%d/%m/%Y %H:%M')}"),
        ("Δ mecánico", f"{_fmt(val.delta_mecanico, 0)} m³"),
        ("Consumo WES+placa ventana", f"{_fmt(val.wes_m3, 1)} m³"),
        (
            "Diferencia (mecánico − serie)",
            f"{_fmt(val.diferencia_m3, 1)} m³ ({_fmt(val.diferencia_pct, 1)} %)",
        ),
        ("Estado", val.estado),
    ]
    t2 = doc.add_table(rows=len(rows_data), cols=2)
    t2.style = "Table Grid"
    for i, (a, b) in enumerate(rows_data):
        t2.rows[i].cells[0].text = a
        t2.rows[i].cells[1].text = b
        for cell in t2.rows[i].cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    _set_run_font(run, bold=(i == 0 or a == "Estado"), size=10)
        if i == 0:
            for cell in t2.rows[i].cells:
                _shade_cell(cell, "1F4788")
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(255, 255, 255)

    p = doc.add_paragraph()
    r = p.add_run("Detalle horario de la ventana:")
    _set_run_font(r, bold=True, size=10)

    det_rows = [("Hora Chile", "m³ aporte", "Fuente / nota")]
    for etq, contrib, nota in val.detalle:
        if contrib == 0 and "mitad" not in nota and "placa" not in nota:
            continue
        det_rows.append((etq, _fmt(contrib, 2), nota))
    t3 = doc.add_table(rows=len(det_rows), cols=3)
    t3.style = "Table Grid"
    for i, row in enumerate(det_rows):
        for j, txt in enumerate(row):
            t3.rows[i].cells[j].text = txt
            for para in t3.rows[i].cells[j].paragraphs:
                for run in para.runs:
                    _set_run_font(run, bold=(i == 0), size=8)
        if i == 0:
            for cell in t3.rows[i].cells:
                _shade_cell(cell, "1F4788")
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(255, 255, 255)

    h = doc.add_heading("6. Conclusión", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO
    p = doc.add_paragraph()
    r = p.add_run(
        "Se confirma falla de memoria de placa (sensor Sensus/HRI y voltajes OK). "
        f"Validación post-cambio: Δ mecánico {_fmt(val.delta_mecanico, 0)} m³ vs "
        f"serie {_fmt(val.wes_m3, 1)} m³ (diferencia {_fmt(val.diferencia_m3, 1)} m³ / "
        f"{_fmt(val.diferencia_pct, 1)} %). {val.estado}. "
        f"Seguimiento diario matutino del nodo {NODE_ID} con umbral "
        f"{CAUDAL_MAX_REF_M3H:.0f} m³/h (techo DN90)."
    )
    _set_run_font(r, size=11)

    p = doc.add_paragraph()
    r = p.add_run(f"Nota: informe intervención + validación nodo {NODE_ID}.")
    _set_run_font(r, size=9, color=RGBColor(100, 100, 100))

    doc.save(out)
    print(
        f"[VALIDACION] Δ mec={val.delta_mecanico} | serie={val.wes_m3} | "
        f"dif={val.diferencia_m3} ({val.diferencia_pct}%) | {val.estado}"
    )
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--lectura-ayer", type=float, default=LECTURA_AYER_M3)
    parser.add_argument("--lectura-hoy", type=float, default=LECTURA_HOY_M3)
    args = parser.parse_args()
    path = build_doc(args.lectura_ayer, args.lectura_hoy)
    print(f"[OK] Informe generado: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
