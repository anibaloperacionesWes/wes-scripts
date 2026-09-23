"""
Informe corto — cambio de memoria placa / Etapa N°5 Fundo Zapallar.

Incluye validación lectura mecánica vs consumo WES:
  - Inicio: lectura terreno 22/09/2026 14:30 → mitad del consumo hora 14.
  - Fin: lectura terreno 23/09/2026 16:54 → consumo WES hasta las 16:00.

Uso:
  python generar_informe_cambio_memoria_etapa5_zapallar.py
  python generar_informe_cambio_memoria_etapa5_zapallar.py --lectura-ayer 5160
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import List, Optional, Tuple
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

# Lecturas / medidor (foto terreno 23-09-2026 16:54, Zapallar)
LECTURA_HOY_M3 = 5177.0
LECTURA_HOY_DT = datetime(2026, 9, 23, 16, 54, tzinfo=CHILE_TZ)
LECTURA_AYER_DT = datetime(2026, 9, 22, 14, 30, tzinfo=CHILE_TZ)
# Corte WES al cierre de validación (según instrucción: hasta las 16:00 de hoy)
WES_HASTA_DT = datetime(2026, 9, 23, 16, 0, tzinfo=CHILE_TZ)

MEDIDOR_MARCA = "Sensus"
MEDIDOR_MODELO = "MeiStream Plus 100"
MEDIDOR_SERIE = "8 SEN01 2370 9030"
MEDIDOR_Q3_M3H = 100.0
HRI_MODELO = "HRI-Mei B4 500 ms"
HRI_SERIE = "31730485"
HRI_PULSO_DN40_125_L = 100  # 1 pulso = 100 L = 0,1 m³


@dataclass
class HoraWes:
    dt_chile: datetime
    m3h: float


@dataclass
class Validacion:
    lectura_ayer: Optional[float]
    lectura_hoy: float
    delta_mecanico: Optional[float]
    wes_m3: float
    detalle: List[Tuple[str, float, str]]  # etiqueta, contrib, nota
    hueco_horas: List[str]
    diferencia_m3: Optional[float]
    diferencia_pct: Optional[float]
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


def obtener_horarios(dia: date) -> List[HoraWes]:
    base = acl_node_base_url()
    ds = dia.strftime("%d%m%Y")
    url = f"{base}/nodes/{NODE_ID}/dates.measures.csv"
    r = requests.get(url, params=[("start", ds), ("end", ds)], timeout=45)
    r.raise_for_status()
    out: List[HoraWes] = []
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
        out.append(HoraWes(dt.astimezone(CHILE_TZ), float(parts[1])))
    return out


def calcular_validacion(lectura_ayer: Optional[float]) -> Validacion:
    """
    Ventana WES:
      - 22/09 14:30 → 0,5 × consumo hora 14:00–15:00
      - horas completas posteriores hasta < 23/09 16:00
    """
    series: List[HoraWes] = []
    # Pedir 21–23 para cubrir bordes UTC→Chile
    for d in (date(2026, 9, 21), date(2026, 9, 22), date(2026, 9, 23)):
        series.extend(obtener_horarios(d))

    # Deduplicar por timestamp
    by_ts = {h.dt_chile: h for h in series}
    series = [by_ts[k] for k in sorted(by_ts)]

    inicio_hora = LECTURA_AYER_DT.replace(minute=0, second=0, microsecond=0)
    detalle: List[Tuple[str, float, str]] = []
    total = 0.0
    presentes = set()

    for h in series:
        if h.dt_chile == inicio_hora:
            contrib = 0.5 * h.m3h
            detalle.append(
                (
                    h.dt_chile.strftime("%d/%m/%Y %H:%M"),
                    contrib,
                    f"mitad hora 14 (lectura 14:30); bruto {_fmt(h.m3h, 2)} m³/h",
                )
            )
            total += contrib
            presentes.add(h.dt_chile)
        elif inicio_hora < h.dt_chile < WES_HASTA_DT:
            detalle.append(
                (h.dt_chile.strftime("%d/%m/%Y %H:%M"), h.m3h, "hora completa en ventana")
            )
            total += h.m3h
            presentes.add(h.dt_chile)

    # Detectar huecos horarios esperados en la ventana
    huecos: List[str] = []
    cursor = inicio_hora
    while cursor < WES_HASTA_DT:
        if cursor not in presentes and cursor != inicio_hora:
            # la hora de inicio se maneja aparte (mitad); si falta, también es hueco
            huecos.append(cursor.strftime("%d/%m/%Y %H:%M"))
        elif cursor == inicio_hora and cursor not in presentes:
            huecos.append(cursor.strftime("%d/%m/%Y %H:%M") + " (hora 14, se asumió 0)")
            detalle.insert(
                0,
                (
                    cursor.strftime("%d/%m/%Y %H:%M"),
                    0.0,
                    "sin dato WES en hora 14 → mitad = 0 (hueco post intervención)",
                ),
            )
        cursor += timedelta(hours=1)

    delta = None
    dif = None
    pct = None
    if lectura_ayer is not None:
        delta = LECTURA_HOY_M3 - lectura_ayer
        dif = delta - total
        pct = (dif / delta * 100.0) if delta else None
        if pct is not None and abs(pct) <= 5:
            estado = "OK — diferencia ≤ 5 %"
        elif pct is not None and abs(pct) <= 15:
            estado = "REVISAR — diferencia moderada"
        else:
            estado = "ALERTA — diferencia relevante o hueco de datos"
    else:
        estado = "PENDIENTE — falta lectura mecánica de ayer 14:30"

    return Validacion(
        lectura_ayer=lectura_ayer,
        lectura_hoy=LECTURA_HOY_M3,
        delta_mecanico=delta,
        wes_m3=round(total, 2),
        detalle=detalle,
        hueco_horas=huecos,
        diferencia_m3=None if dif is None else round(dif, 2),
        diferencia_pct=None if pct is None else round(pct, 1),
        estado=estado,
    )


def build_doc(lectura_ayer: Optional[float]) -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    ahora = datetime.now(CHILE_TZ)
    stamp = ahora.strftime("%Y%m%d_%H%M")
    out = OUT_DIR / f"Informe_Cambio_Memoria_Etapa5_Zapallar_{stamp}.docx"
    val = calcular_validacion(lectura_ayer)

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
        "normalizar el punto y evitar recurrencia."
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

    h = doc.add_heading("3. Medidor en terreno", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO

    _add_bullet(doc, f"Marca/modelo: {MEDIDOR_MARCA} {MEDIDOR_MODELO}")
    _add_bullet(doc, f"Serie medidor: {MEDIDOR_SERIE} (fabricación 2023)")
    _add_bullet(doc, f"Q3 medidor: {_fmt(MEDIDOR_Q3_M3H, 0)} m³/h (capacidad del instrumento)")
    _add_bullet(doc, f"Módulo pulsos: {HRI_MODELO}, serie {HRI_SERIE}")
    _add_bullet(
        doc,
        f"Peso de pulso DN 40–125: {HRI_PULSO_DN40_125_L} L/pulso "
        f"(= {HRI_PULSO_DN40_125_L/1000:.1f} m³/pulso), aplicable a DN90",
    )
    _add_bullet(
        doc,
        f"Lectura mecánica hoy: {_fmt(LECTURA_HOY_M3, 0)} m³ "
        f"({LECTURA_HOY_DT.strftime('%d/%m/%Y %H:%M')} Chile, foto terreno Zapallar)",
    )

    foto = FOTOS_DIR / "lectura_20260923_1654_sensus.png"
    if foto.is_file():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(foto), width=Inches(5.2))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = cap.add_run(
            f"Foto lectura {LECTURA_HOY_DT.strftime('%d/%m/%Y %H:%M')} — "
            f"{MEDIDOR_MODELO} = {_fmt(LECTURA_HOY_M3, 0)} m³"
        )
        _set_run_font(r, size=9, color=RGBColor(90, 90, 90))

    h = doc.add_heading("4. Criterio hidráulico (DN90)", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO

    p = doc.add_paragraph()
    r = p.add_run(
        f"La red del punto es tubería de {DIAMETRO}. Como referencia de capacidad "
        f"máxima razonable de red se adopta ≈ {CAUDAL_MAX_REF_M3H:.0f} m³/h "
        f"(~2,5 m/s en DI ≈ 90 mm). El medidor admite Q3 = {_fmt(MEDIDOR_Q3_M3H, 0)} m³/h, "
        f"pero la red DN90 no puede entregar de forma realista caudales de 92–200 m³/h."
    )
    _set_run_font(r, size=11)

    table = doc.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    filas = [
        ("Velocidad de referencia", "Caudal teórico DN90"),
        ("1,5 m/s (diseño habitual)", "≈ 34 m³/h"),
        ("2,0 m/s", "≈ 46 m³/h"),
        ("2,5 m/s (techo práctico de red)", "≈ 57–60 m³/h"),
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

    p = doc.add_paragraph()
    r = p.add_run(
        "Los pulsos observados de ~92 m³/h y ~200 m³/h superan el techo hidráulico "
        "de la tubería (velocidades ~4–9 m/s); corresponden a error de registro en memoria."
    )
    _set_run_font(r, size=11)

    h = doc.add_heading("5. Validación lecturas vs WES", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO

    p = doc.add_paragraph()
    r = p.add_run(
        "Criterio de ventana (acordado): lectura de ayer a las 14:30 → se toma la "
        "mitad del consumo de la hora 14; lectura de hoy → consumo WES hasta las 16:00 "
        f"({WES_HASTA_DT.strftime('%d/%m/%Y %H:%M')} Chile), aunque la foto sea a las "
        f"{LECTURA_HOY_DT.strftime('%H:%M')}."
    )
    _set_run_font(r, size=11)

    # Tabla resumen validación
    rows_data = [
        ("Concepto", "Valor"),
        (
            "Lectura ayer (terreno)",
            (
                f"{_fmt(val.lectura_ayer, 0)} m³ @ {LECTURA_AYER_DT.strftime('%d/%m/%Y %H:%M')}"
                if val.lectura_ayer is not None
                else f"PENDIENTE @ {LECTURA_AYER_DT.strftime('%d/%m/%Y %H:%M')}"
            ),
        ),
        (
            "Lectura hoy (terreno)",
            f"{_fmt(val.lectura_hoy, 0)} m³ @ {LECTURA_HOY_DT.strftime('%d/%m/%Y %H:%M')}",
        ),
        (
            "Δ mecánico (hoy − ayer)",
            f"{_fmt(val.delta_mecanico, 1)} m³" if val.delta_mecanico is not None else "—",
        ),
        ("Consumo WES ventana", f"{_fmt(val.wes_m3, 1)} m³"),
        (
            "Diferencia (mecánico − WES)",
            (
                f"{_fmt(val.diferencia_m3, 1)} m³ ({_fmt(val.diferencia_pct, 1)} %)"
                if val.diferencia_m3 is not None
                else "—"
            ),
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
    r = p.add_run(
        f"Consumo WES en ventana = {_fmt(val.wes_m3, 1)} m³ "
        f"(mitad hora 14 del {LECTURA_AYER_DT.strftime('%d/%m')} + horas completas "
        f"hasta antes de las {WES_HASTA_DT.strftime('%H:%M')} del {WES_HASTA_DT.strftime('%d/%m')})."
    )
    _set_run_font(r, size=11)

    if val.hueco_horas:
        p = doc.add_paragraph()
        r = p.add_run(
            "Hueco de datos WES en la ventana (probable durante falla/cambio de memoria): "
            + ", ".join(val.hueco_horas[:12])
            + ("…" if len(val.hueco_horas) > 12 else "")
            + f" ({len(val.hueco_horas)} hora(s) sin serie)."
        )
        _set_run_font(r, size=10, color=RGBColor(120, 60, 0))

    # Detalle horario (solo contribuciones > 0 o notas especiales)
    p = doc.add_paragraph()
    r = p.add_run("Detalle horario WES usado en la validación:")
    _set_run_font(r, bold=True, size=10)

    det_rows = [("Hora Chile", "m³ aporte", "Nota")]
    for etq, contrib, nota in val.detalle:
        if contrib == 0 and "sin dato" not in nota:
            continue
        det_rows.append((etq, _fmt(contrib, 2), nota))
    if len(det_rows) == 1:
        det_rows.append(("—", "0,00", "Sin aportes > 0 en ventana"))

    t3 = doc.add_table(rows=len(det_rows), cols=3)
    t3.style = "Table Grid"
    for i, row in enumerate(det_rows):
        for j, txt in enumerate(row):
            t3.rows[i].cells[j].text = txt
            for para in t3.rows[i].cells[j].paragraphs:
                for run in para.runs:
                    _set_run_font(run, bold=(i == 0), size=9)
        if i == 0:
            for cell in t3.rows[i].cells:
                _shade_cell(cell, "1F4788")
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(255, 255, 255)

    if val.lectura_ayer is None:
        p = doc.add_paragraph()
        r = p.add_run(
            "Para cerrar el % de diferencia falta la lectura mecánica de ayer "
            f"({LECTURA_AYER_DT.strftime('%d/%m/%Y')} 14:30). "
            "Re-generar con: python generar_informe_cambio_memoria_etapa5_zapallar.py "
            "--lectura-ayer NNNN"
        )
        _set_run_font(r, size=10, color=RGBColor(150, 0, 0))

    h = doc.add_heading("6. Conclusión", level=1)
    for run in h.runs:
        run.font.color.rgb = COLOR_TITULO

    p = doc.add_paragraph()
    r = p.add_run(
        "Se confirma falla de memoria de placa (no del sensor Sensus/HRI ni de la red). "
        "Con el reemplazo de memoria el punto queda normalizado. "
        f"Consumo WES post-ventana de validación: {_fmt(val.wes_m3, 1)} m³ "
        f"hasta las {WES_HASTA_DT.strftime('%H:%M')} del {WES_HASTA_DT.strftime('%d/%m/%Y')}. "
        "Se mantiene seguimiento diario matutino del nodo 000027-03 con umbral de alerta "
        f"en {CAUDAL_MAX_REF_M3H:.0f} m³/h."
    )
    _set_run_font(r, size=11)

    p = doc.add_paragraph()
    r = p.add_run(
        "Nota: informe de intervención en terreno + validación de lecturas vs API WES "
        f"(nodo {NODE_ID})."
    )
    _set_run_font(r, size=9, color=RGBColor(100, 100, 100))

    doc.save(out)
    print(f"[VALIDACION] WES ventana = {val.wes_m3} m³ | estado = {val.estado}")
    if val.hueco_horas:
        print(f"[VALIDACION] Huecos WES: {len(val.hueco_horas)} horas")
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--lectura-ayer",
        type=float,
        default=None,
        help="Lectura mecánica de ayer a las 14:30 (m³). Si se omite, queda pendiente en el informe.",
    )
    args = parser.parse_args()
    path = build_doc(args.lectura_ayer)
    print(f"[OK] Informe generado: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
