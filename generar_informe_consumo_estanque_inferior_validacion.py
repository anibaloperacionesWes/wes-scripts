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

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import requests
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parent
BASE = "http://104.248.53.141:7003/wes/api/acl-node/v1"

MATRIZ_ID = "000027-01"
INFERIOR_ID = "000027-02"
ETAPA5_ID = "000027-03"

ITRON_B_M3 = 170.89
APP_MATRIZ_B_REF = 175.41

# Paleta WES
COLOR_MATRIZ = "#1F4E79"
COLOR_INFERIOR = "#2E7D32"
COLOR_ETAPA5 = "#C62828"
COLOR_GRID = "#E0E0E0"


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


def ventana_b_slots() -> list[tuple[str, int, str]]:
    """Lista (fecha, hora, etiqueta) de la ventana B."""
    slots: list[tuple[str, int, str]] = []
    for h in range(16, 24):
        slots.append(("2026-09-22", h, f"22/{h:02d}"))
    for h in range(0, 17):
        slots.append(("2026-09-23", h, f"23/{h:02d}"))
    return slots


def sum_ventana_b(node_id: str) -> tuple[float, list[tuple[str, int, float]]]:
    h22 = hours_app(node_id, date(2026, 9, 22))
    h23 = hours_app(node_id, date(2026, 9, 23))
    by_day = {"2026-09-22": h22, "2026-09-23": h23}
    detail: list[tuple[str, int, float]] = []
    total = 0.0
    for fecha, h, _ in ventana_b_slots():
        v = by_day[fecha].get(h, 0.0)
        detail.append((fecha, h, v))
        total += v
    return total, detail


def sum_ventana_a(node_id: str) -> float:
    h22 = hours_app(node_id, date(2026, 9, 22))
    return sum(h22.get(h, 0.0) for h in (12, 13, 14))


def sum_ventana_etapa5(node_id: str) -> float:
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
                _font(run, bold=True, size=9)
    for r_i, row in enumerate(rows):
        for c_i, val in enumerate(row):
            cell = table.rows[r_i + 1].cells[c_i]
            cell.text = val
            for p in cell.paragraphs:
                for run in p.runs:
                    _font(run, size=9, bold=(r_i == len(rows) - 1))
    doc.add_paragraph()


def build_charts(
    out_dir: Path,
    stamp: str,
    labels: list[str],
    matriz_vals: list[float],
    inferior_vals: list[float],
    etapa5_vals: list[float],
    matriz_total: float,
    inferior_total: float,
    etapa5_total: float,
) -> tuple[Path, Path]:
    # --- Gráfico horario Matriz vs Inferior ---
    fig, ax = plt.subplots(figsize=(12, 4.8))
    x = list(range(len(labels)))
    w = 0.38
    ax.bar([i - w / 2 for i in x], matriz_vals, width=w, label="Matriz ESVAL", color=COLOR_MATRIZ)
    ax.bar(
        [i + w / 2 for i in x],
        inferior_vals,
        width=w,
        label="Estanque Inferior",
        color=COLOR_INFERIOR,
    )
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=90, fontsize=8)
    ax.set_ylabel("m³ / hora")
    ax.set_title(
        "Consumo horario — ventana validación Matriz ESVAL\n"
        "22-09 15:00 → 23-09 17:00 (h16–23 + h00–16)"
    )
    ax.legend(loc="upper right")
    ax.grid(axis="y", color=COLOR_GRID, linestyle="--", linewidth=0.7)
    ax.set_axisbelow(True)
    # Separador visual entre 22 y 23
    ax.axvline(7.5, color="#9E9E9E", linestyle=":", linewidth=1)
    ax.text(3.5, ax.get_ylim()[1] * 0.95, "22-09", ha="center", fontsize=9, color="#616161")
    ax.text(16, ax.get_ylim()[1] * 0.95, "23-09", ha="center", fontsize=9, color="#616161")
    fig.tight_layout()
    chart_horario = out_dir / f"chart_horario_matriz_vs_inferior_{stamp}.png"
    fig.savefig(chart_horario, dpi=150)
    plt.close(fig)

    # --- Gráfico totales ---
    fig2, ax2 = plt.subplots(figsize=(7.2, 4.2))
    cats = ["Matriz ESVAL\n(App)", "Matriz ESVAL\n(Itron)", "Estanque\nInferior", "Etapa N°5"]
    vals = [matriz_total, ITRON_B_M3, inferior_total, etapa5_total]
    colors = [COLOR_MATRIZ, "#5B9BD5", COLOR_INFERIOR, COLOR_ETAPA5]
    bars = ax2.bar(cats, vals, color=colors, width=0.65)
    ax2.set_ylabel("m³ en la ventana")
    ax2.set_title("Totales en ventana de validación (22-09 15:00 → 23-09 17:00)")
    ax2.grid(axis="y", color=COLOR_GRID, linestyle="--", linewidth=0.7)
    ax2.set_axisbelow(True)
    for bar, v in zip(bars, vals):
        ax2.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 2,
            _fmt(v),
            ha="center",
            va="bottom",
            fontsize=10,
            fontweight="bold",
        )
    # Anotar ratio Inferior/Matriz
    ratio = 100.0 * inferior_total / matriz_total if matriz_total else 0.0
    ax2.annotate(
        f"Inferior = {_fmt(ratio, 1)} % de Matriz App",
        xy=(2, inferior_total),
        xytext=(1.2, max(vals) * 0.72),
        arrowprops=dict(arrowstyle="->", color="#424242"),
        fontsize=9,
        color="#424242",
    )
    fig2.tight_layout()
    chart_totales = out_dir / f"chart_totales_validacion_{stamp}.png"
    fig2.savefig(chart_totales, dpi=150)
    plt.close(fig2)

    return chart_horario, chart_totales


def build_report() -> tuple[Path, dict]:
    _, detail_matriz = sum_ventana_b(MATRIZ_ID)
    inferior_b, detail_inf = sum_ventana_b(INFERIOR_ID)
    _, detail_e5 = sum_ventana_b(ETAPA5_ID)

    matriz_b = sum(v for _, _, v in detail_matriz)
    etapa5_b = sum(v for _, _, v in detail_e5)

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

    slots = ventana_b_slots()
    labels = [lab for _, _, lab in slots]
    matriz_vals = [v for _, _, v in detail_matriz]
    inferior_vals = [v for _, _, v in detail_inf]
    etapa5_vals = [v for _, _, v in detail_e5]

    out_dir = ROOT / "reports" / "Fundo_Zapallar" / "Informes_Tecnicos"
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    chart_horario, chart_totales = build_charts(
        out_dir,
        stamp,
        labels,
        matriz_vals,
        inferior_vals,
        etapa5_vals,
        matriz_b,
        inferior_b,
        etapa5_b,
    )

    # CSV tabla horaria
    csv_path = out_dir / f"tabla_horaria_validacion_matriz_inferior_{stamp}.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(
            [
                "Fecha",
                "Hora",
                "Etiqueta",
                "Matriz_ESVAL_m3",
                "Estanque_Inferior_m3",
                "Etapa_N5_m3",
                "Dif_Matriz_menos_Inferior_m3",
            ]
        )
        for (fecha, hora, lab), vm, vi, ve in zip(
            slots, matriz_vals, inferior_vals, etapa5_vals
        ):
            w.writerow(
                [
                    fecha,
                    f"{hora:02d}:00",
                    lab,
                    f"{vm:.2f}".replace(".", ","),
                    f"{vi:.2f}".replace(".", ","),
                    f"{ve:.2f}".replace(".", ","),
                    f"{vm - vi:.2f}".replace(".", ","),
                ]
            )
        w.writerow(
            [
                "TOTAL",
                "",
                "",
                f"{matriz_b:.2f}".replace(".", ","),
                f"{inferior_b:.2f}".replace(".", ","),
                f"{etapa5_b:.2f}".replace(".", ","),
                f"{dif:.2f}".replace(".", ","),
            ]
        )

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
        },
        "tabla_horaria": [
            {
                "fecha": fecha,
                "hora": hora,
                "etiqueta": lab,
                "matriz_m3": round(vm, 2),
                "inferior_m3": round(vi, 2),
                "etapa5_m3": round(ve, 2),
                "dif_m3": round(vm - vi, 2),
            }
            for (fecha, hora, lab), vm, vi, ve in zip(
                slots, matriz_vals, inferior_vals, etapa5_vals
            )
        ],
        "archivos": {
            "csv": str(csv_path),
            "chart_horario": str(chart_horario),
            "chart_totales": str(chart_totales),
        },
    }

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
        "Ventana de validación Matriz ESVAL (Itron vs App, §4.2): "
        "22-09-2026 15:00 → 23-09-2026 17:00 (Chile). "
        "Horas app: 22-09 h16→23 + 23-09 h00→16. "
        "Misma metodología CSV TIME del informe Matriz.",
    )

    _add_heading(doc, "1. Totales en la ventana", level=1)
    _add_para(
        doc,
        f"Estanque Inferior: {_fmt(inferior_b)} m³ "
        f"({_fmt(ratio, 1)} % de Matriz App; diferencia {_fmt(dif)} m³).",
        bold=True,
    )
    _add_table(
        doc,
        ["Punto", "Nodo", "Consumo ventana (m³)"],
        [
            ["Matriz ESVAL (App)", MATRIZ_ID, _fmt(matriz_b)],
            ["Matriz ESVAL (Itron)", MATRIZ_ID, _fmt(ITRON_B_M3)],
            ["Estanque Inferior (App)", INFERIOR_ID, _fmt(inferior_b)],
            ["Etapa N°5 (App)", ETAPA5_ID, _fmt(etapa5_b)],
        ],
    )
    doc.add_picture(str(chart_totales), width=Cm(16))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    _add_heading(doc, "2. Gráfico horario Matriz vs Estanque Inferior", level=1)
    doc.add_picture(str(chart_horario), width=Cm(17))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    _add_heading(doc, "3. Tabla horaria comparativa", level=1)
    rows = [
        [
            lab,
            f"{hora:02d}:00",
            _fmt(vm),
            _fmt(vi),
            _fmt(ve),
            _fmt(vm - vi),
        ]
        for (fecha, hora, lab), vm, vi, ve in zip(
            slots, matriz_vals, inferior_vals, etapa5_vals
        )
    ]
    rows.append(
        ["TOTAL", "", _fmt(matriz_b), _fmt(inferior_b), _fmt(etapa5_b), _fmt(dif)]
    )
    _add_table(
        doc,
        [
            "Etiqueta",
            "Hora",
            "Matriz (m³)",
            "Inferior (m³)",
            "Etapa 5 (m³)",
            "Dif M−I (m³)",
        ],
        rows,
    )

    _add_heading(doc, "4. Conclusión", level=1)
    _add_para(
        doc,
        f"En la misma ventana de validación de Matriz ESVAL, el Estanque Inferior "
        f"registró {_fmt(inferior_b)} m³ (App), frente a {_fmt(matriz_b)} m³ de Matriz App "
        f"y {_fmt(ITRON_B_M3)} m³ de Itron.",
        bold=True,
    )

    doc.save(docx_path)
    print(f"[OK] DOCX: {docx_path}")
    print(f"[OK] JSON: {json_path}")
    print(f"[OK] CSV:  {csv_path}")
    print(f"[OK] CHART horario: {chart_horario}")
    print(f"[OK] CHART totales: {chart_totales}")
    print(f"[CIFRA] Estanque Inferior = {_fmt(inferior_b)} m³")
    return docx_path, payload


if __name__ == "__main__":
    build_report()
