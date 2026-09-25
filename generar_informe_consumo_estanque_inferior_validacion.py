"""
Consumo Estanque Inferior (000027-02) en la misma ventana de validación
de Matriz ESVAL (informe calidad de señal / Itron vs App).

Ventana principal (Matriz §4.2):
  22-09-2026 15:00 → 23-09-2026 17:00 (Chile)
  Horas app: 22-09 h16→23 + 23-09 h00→16

Metodología: idéntica al informe Matriz ESVAL — CSV
  /nodes/{id}/dates.measures.csv con hora del campo TIME (etiqueta),
  sin convertir UTC→Chile.
"""

from __future__ import annotations

import csv
import io
import json
from datetime import date, datetime
from pathlib import Path

import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

ROOT = Path(__file__).resolve().parent
BASE = "http://104.248.53.141:7003/wes/api/acl-node/v1"

MATRIZ_ID = "000027-01"
INFERIOR_ID = "000027-02"
ETAPA5_ID = "000027-03"

# Referencias del informe Matriz ESVAL (validación terreno)
ITRON_B_M3 = 170.89
APP_MATRIZ_B_REF = 175.41


def _fmt(x: float, dec: int = 2) -> str:
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _font(run, *, size: int = 11, bold: bool = False) -> None:
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    run.font.size = Pt(size)
    run.bold = bold


def hours_app(node_id: str, dia: date) -> dict[int, float]:
    r = requests.get(
        f"{BASE}/nodes/{node_id}/dates.measures.csv",
        params={"start": dia.strftime("%d%m%Y"), "end": dia.strftime("%d%m%Y")},
        timeout=60,
    )
    r.raise_for_status()
    acc: dict[int, float] = {}
    for row in csv.DictReader(io.StringIO(r.text)):
        t = row["TIME"].strip()
        v = float(row["VALUE"].strip())
        dt = datetime.fromisoformat(t.replace("Z", "+00:00"))
        acc[dt.hour] = acc.get(dt.hour, 0.0) + v
    return acc


def sum_ventana_b(node_id: str) -> tuple[float, list[tuple[str, int, float]]]:
    """15:00→17:00 exactos → horas 16–23 del 22 + 00–16 del 23."""
    h22 = hours_app(node_id, date(2026, 9, 22))
    h23 = hours_app(node_id, date(2026, 9, 23))
    detail: list[tuple[str, int, float]] = []
    total = 0.0
    for h in range(16, 24):
        v = h22.get(h, 0.0)
        detail.append(("2026-09-22", h, v))
        total += v
    for h in range(0, 17):
        v = h23.get(h, 0.0)
        detail.append(("2026-09-23", h, v))
        total += v
    return total, detail


def sum_ventana_a(node_id: str) -> float:
    h22 = hours_app(node_id, date(2026, 9, 22))
    return sum(h22.get(h, 0.0) for h in (12, 13, 14))


def sum_ventana_etapa5(node_id: str) -> float:
    """14:00→17:00 → horas 14–23 del 22 + 00–16 del 23."""
    h22 = hours_app(node_id, date(2026, 9, 22))
    h23 = hours_app(node_id, date(2026, 9, 23))
    return sum(h22.get(h, 0.0) for h in range(14, 24)) + sum(
        h23.get(h, 0.0) for h in range(0, 17)
    )


def _add_heading(doc: Document, text: str, level: int = 1) -> None:
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        _font(run, size=14 if level == 1 else 12, bold=True)


def _add_para(doc: Document, text: str, *, bold: bool = False) -> None:
    p = doc.add_paragraph()
    run = p.add_run(text)
    _font(run, bold=bold)
    p.paragraph_format.space_after = Pt(6)


def _add_table(doc: Document, headers: list[str], rows: list[list[str]]) -> None:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for p in cell.paragraphs:
            for run in p.runs:
                _font(run, bold=True, size=10)
    for r_i, row in enumerate(rows):
        for c_i, val in enumerate(row):
            cell = table.rows[r_i + 1].cells[c_i]
            cell.text = val
            for p in cell.paragraphs:
                for run in p.runs:
                    _font(run, size=10)
    doc.add_paragraph()


def build_report() -> tuple[Path, dict]:
    matriz_b, _ = sum_ventana_b(MATRIZ_ID)
    inferior_b, detail_inf = sum_ventana_b(INFERIOR_ID)
    etapa5_b, _ = sum_ventana_b(ETAPA5_ID)

    matriz_a = sum_ventana_a(MATRIZ_ID)
    inferior_a = sum_ventana_a(INFERIOR_ID)

    matriz_e5 = sum_ventana_etapa5(MATRIZ_ID)
    inferior_e5 = sum_ventana_etapa5(INFERIOR_ID)
    etapa5_e5 = sum_ventana_etapa5(ETAPA5_ID)

    assert abs(matriz_b - APP_MATRIZ_B_REF) < 0.05, (
        f"Matriz B={matriz_b} no coincide con referencia {APP_MATRIZ_B_REF}"
    )

    ratio = 100.0 * inferior_b / matriz_b if matriz_b else 0.0
    dif = matriz_b - inferior_b

    payload = {
        "ventana_matriz_B": {
            "inicio": "2026-09-22 15:00",
            "fin": "2026-09-23 17:00",
            "horas_app": "22-09 h16-23 + 23-09 h00-16",
            "matriz_app_m3": round(matriz_b, 2),
            "itron_matriz_m3": ITRON_B_M3,
            "estanque_inferior_m3": round(inferior_b, 2),
            "etapa5_m3": round(etapa5_b, 2),
            "inferior_sobre_matriz_pct": round(ratio, 1),
            "diferencia_matriz_menos_inferior_m3": round(dif, 2),
        },
        "ventana_A_h12_14": {
            "matriz_m3": round(matriz_a, 2),
            "estanque_inferior_m3": round(inferior_a, 2),
        },
        "ventana_etapa5_14_17": {
            "matriz_m3": round(matriz_e5, 2),
            "estanque_inferior_m3": round(inferior_e5, 2),
            "etapa5_m3": round(etapa5_e5, 2),
            "nota": "Etapa 5 el 22-09 h14-23 puede estar en hueco de placa; cifra API pura.",
        },
        "detalle_horario_inferior_B": [
            {"fecha": d, "hora": h, "m3": round(v, 2)} for d, h, v in detail_inf
        ],
    }

    out_dir = ROOT / "reports" / "Fundo_Zapallar" / "Informes_Tecnicos"
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    docx_path = out_dir / f"Consumo_Estanque_Inferior_validacion_Matriz_ESVAL_2209_2309_{stamp}.docx"
    json_path = out_dir / f"Consumo_Estanque_Inferior_validacion_Matriz_ESVAL_2209_2309_{stamp}.json"
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    doc = Document()
    title = doc.add_paragraph()
    run = title.add_run(
        "Fundo Zapallar — Consumo Estanque Inferior\n"
        "misma ventana de validación Matriz ESVAL"
    )
    _font(run, size=16, bold=True)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    _add_para(
        doc,
        "Se toma la ventana de validación del informe de Matriz ESVAL "
        "(Itron vs App WES, §4.2) y se calcula cuánta agua registró el "
        "Estanque Inferior (000027-02) en ese mismo periodo, con la misma "
        "metodología de horas app.",
    )

    _add_heading(doc, "1. Periodo (validación Matriz ESVAL)", level=1)
    _add_para(doc, "Inicio: 22-09-2026 15:00 (Chile)")
    _add_para(doc, "Fin: 23-09-2026 17:00 (Chile)")
    _add_para(
        doc,
        "Horas app sumadas: 22-09 horas 16→23 + 23-09 horas 00→16 "
        "(límites exactos HH:00: se excluye la hora de inicio y la hora de fin).",
    )
    _add_para(
        doc,
        "Fuente: CSV WES dates.measures — hora del campo TIME (misma regla "
        "que el total App 175,41 m³ del informe Matriz).",
    )

    _add_heading(doc, "2. Resultado — Estanque Inferior", level=1)
    _add_para(
        doc,
        f"Consumo Estanque Inferior en la ventana: {_fmt(inferior_b)} m³.",
        bold=True,
    )
    _add_table(
        doc,
        ["Punto", "Nodo", "Consumo ventana (m³)"],
        [
            ["Matriz ESVAL (App)", MATRIZ_ID, _fmt(matriz_b)],
            ["Matriz ESVAL (Itron, validación)", MATRIZ_ID, _fmt(ITRON_B_M3)],
            ["Estanque Inferior (App)", INFERIOR_ID, _fmt(inferior_b)],
            ["Etapa N°5 (App, misma ventana)", ETAPA5_ID, _fmt(etapa5_b)],
        ],
    )
    _add_para(
        doc,
        f"Inferior / Matriz App = {_fmt(ratio, 1)} % "
        f"(diferencia Matriz − Inferior = {_fmt(dif)} m³). "
        "Es coherente con el esquema ESVAL → Estanque Inferior "
        "(parte del caudal de matriz puede ir a riego u otras salidas "
        "antes/durante el llenado).",
    )

    _add_heading(doc, "3. Detalle horario Estanque Inferior", level=1)
    _add_table(
        doc,
        ["Fecha", "Hora", "m³"],
        [[d, f"{h:02d}:00", _fmt(v)] for d, h, v in detail_inf],
    )

    _add_heading(doc, "4. Ventanas de contexto", level=1)
    _add_para(
        doc,
        f"Ventana A Matriz (22-09 h12+h13+h14, Itron vs US): "
        f"Matriz {_fmt(matriz_a)} m³ · Inferior {_fmt(inferior_a)} m³.",
    )
    _add_para(
        doc,
        f"Ventana alineada a validación Etapa 5 (22-09 14:00 → 23-09 17:00): "
        f"Matriz {_fmt(matriz_e5)} m³ · Inferior {_fmt(inferior_e5)} m³ · "
        f"Etapa 5 API {_fmt(etapa5_e5)} m³ "
        f"(en terreno Etapa 5 midió ~33 m³ mecánicos; el 22-09 h14–23 puede "
        f"estar en hueco de placa).",
    )

    _add_heading(doc, "5. Conclusión", level=1)
    _add_para(
        doc,
        f"En el mismo periodo de la validación de Matriz ESVAL "
        f"(22-09 15:00 → 23-09 17:00), el Estanque Inferior gastó / registró "
        f"{_fmt(inferior_b)} m³ (App WES), equivalentes al {_fmt(ratio, 1)} % "
        f"del caudal de Matriz App ({_fmt(matriz_b)} m³; Itron {_fmt(ITRON_B_M3)} m³).",
        bold=True,
    )

    doc.save(docx_path)
    print(f"[OK] DOCX: {docx_path}")
    print(f"[OK] JSON: {json_path}")
    print(f"[CIFRA] Estanque Inferior ventana B = {_fmt(inferior_b)} m³")
    return docx_path, payload


if __name__ == "__main__":
    build_report()
