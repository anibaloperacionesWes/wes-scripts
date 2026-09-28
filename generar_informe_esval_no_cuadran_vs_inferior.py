"""
Respuesta puntual: los meses en que la cuenta ESVAL no calza con Matriz WES,
¿llegan a cuadrar con Estanque Inferior?

Fuente: mensual_esval_matriz_inferior.json (día completo Chile, delta turbina).
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

from generar_tabla_zapallar_matriz_vs_inferior import (
    HEADER_FILL,
    ROW_ALERTA_FILL,
    TOTAL_FILL,
    WES_BLUE,
    _fmt_m3,
    _set_cell_text,
    _set_landscape,
)
from generar_tabla_zapallar_mensual_esval_inferior import _fmt_pct_s

SRC = Path(
    "reports/Fundo_Zapallar/ANALISIS/tabla_mensual_esval_inferior_20260922_1510/"
    "mensual_esval_matriz_inferior.json"
)

# Meses donde Matriz no cuadra (±5 %) con la cuenta ESVAL
NO_CUADRAN_MATRIZ = {"Mar-26", "Abr-26", "May-26", "Jun-26", "Jul-26"}
# Agosto queda en "cerca" (−13 %); se muestra aparte como borde


HEADERS = [
    "Mes",
    "Ciclo cuenta",
    "N° boleta",
    "Cuenta ESVAL\n(m³)",
    "Matriz WES\n(m³)",
    "Dif Matriz\nvs ESVAL",
    "¿Matriz\ncuadra?",
    "Estanque Inf.\n(m³)",
    "Dif Inf.\nvs ESVAL",
    "¿Inf.\ncuadra?",
    "¿Inf. rescata\nla cuenta?",
]


def _color_chk(chk: str) -> str:
    if chk == "SI":
        return "C6E0B4"
    if chk == "cerca":
        return "FFF2CC"
    return "F4C7C3"


def _rescata(chk_mat: str, chk_inf: str) -> str:
    """¿El Inferior cuadra con ESVAL cuando Matriz no lo hace?"""
    if chk_mat in ("SI", "cerca") and chk_inf == "SI":
        return "n/a (Matriz ya ok/cerca)"
    if chk_mat == "NO" and chk_inf == "SI":
        return "SI — Inf. cuadra"
    if chk_mat == "NO" and chk_inf == "cerca":
        return "parcial (cerca)"
    if chk_mat == "NO" and chk_inf == "NO":
        return "NO"
    if chk_mat == "cerca" and chk_inf == "SI":
        return "SI — Inf. mejor"
    return "NO"


def main() -> None:
    data = json.loads(SRC.read_text(encoding="utf-8"))
    ciclos = data["ciclos"]

    # Todos los ciclos desde marzo (y febrero para contraste)
    foco = [c for c in ciclos if c["mes"] in NO_CUADRAN_MATRIZ or c["mes"] in ("Feb-26", "Ago-26")]
    no_cuadran = [c for c in ciclos if c["chk_mat_esval"] == "NO"]

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = Path("reports/Fundo_Zapallar/ANALISIS") / f"esval_no_cuadran_vs_inferior_{stamp}"
    out_dir.mkdir(parents=True, exist_ok=True)

    rows_out = []
    for c in ciclos:
        rescate = _rescata(c["chk_mat_esval"], c["chk_inf_esval"])
        rows_out.append({**c, "rescate": rescate})

    # --- Gráfico: solo meses que no cuadran Matriz vs ESVAL ---
    labels = [c["mes"] for c in no_cuadran]
    x = np.arange(len(labels))
    w = 0.28
    fig, ax = plt.subplots(figsize=(12.5, 5.0))
    ax.bar(x - w, [c["esval"] for c in no_cuadran], w, label="Cuenta ESVAL (turbina)", color="#7F7F7F")
    ax.bar(x, [c["matriz"] for c in no_cuadran], w, label="Matriz WES", color=WES_BLUE)
    ax.bar(x + w, [c["inferior"] for c in no_cuadran], w, label="Estanque Inferior", color="#70AD47")
    ax.set_ylabel("m³")
    ax.set_title(
        "Meses que NO cuadran: cuenta ESVAL vs Matriz vs Estanque Inferior",
        fontsize=12,
        fontweight="bold",
        color=WES_BLUE,
    )
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10)
    ax.legend(frameon=False, fontsize=9)
    ax.grid(axis="y", linestyle="--", alpha=0.35)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    chart = out_dir / "chart_meses_no_cuadran_tres_vias.png"
    fig.savefig(chart, dpi=160, bbox_inches="tight")
    plt.close(fig)

    # --- Word ---
    doc = Document()
    _set_landscape(doc)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = t.add_run(
        "¿Los meses que no cuadran con Matriz cuadran con Estanque Inferior?"
    )
    run.bold = True
    run.font.size = Pt(15)
    run.font.color.rgb = RGBColor(31, 78, 121)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = sub.add_run("Fundo Zapallar — cuentas ESVAL (delta medidor turbina) vs WES")
    r2.font.size = Pt(11)
    r2.font.color.rgb = RGBColor(89, 89, 89)

    n = doc.add_paragraph()
    n.add_run(
        "Pregunta: cuando costamos las cuentas ESVAL contra las lecturas de Matriz "
        "(000027-01), desde marzo en adelante no calzan. ¿Esos mismos meses llegan a "
        "cuadrar con Estanque Inferior (000027-02)?\n\n"
        "Criterio: ±5 % = SI, ±15 % = cerca, >15 % = NO. "
        "Cuenta = delta del medidor de turbina ESVAL (no el ítem “facturado”). "
        "WES = día completo 00:00–24:00 Chile por ciclo de boleta. "
        "Marzo–abril = transición (cambio de medidor); no calibrar con esos dos."
    ).font.size = Pt(10)

    # Veredicto grande
    verd = doc.add_paragraph()
    verd.alignment = WD_ALIGN_PARAGRAPH.CENTER
    vr = verd.add_run(
        "VEREDICTO: NO — ninguno de los 5 meses con Matriz≠ESVAL (Mar–Jul 2026) "
        "cuadra con Estanque Inferior."
    )
    vr.bold = True
    vr.font.size = Pt(12)
    vr.font.color.rgb = RGBColor(192, 0, 0)

    detalle = doc.add_paragraph()
    detalle.add_run(
        "Detalle: en May y Ago el Inferior sí se parece a Matriz (±7–8 %), pero ambos "
        "quedan bajos vs la turbina ESVAL. En Mar–Abr (transición) el Inferior se dispara "
        "por encima de la cuenta. Solo Feb-26 el Inferior cuadra con ESVAL (±4 %) mientras "
        "Matriz queda “cerca” (−10 %)."
    ).font.size = Pt(10)

    # Tabla completa 12 ciclos con columna rescate
    body = []
    for c in rows_out:
        body.append(
            [
                c["mes"],
                c["ciclo"],
                c["boleta"],
                _fmt_m3(c["esval"], 0),
                _fmt_m3(c["matriz"], 1),
                _fmt_pct_s(c["pct_mat_esval"]),
                c["chk_mat_esval"],
                _fmt_m3(c["inferior"], 1),
                _fmt_pct_s(c["pct_inf_esval"]),
                c["chk_inf_esval"],
                c["rescate"] if c["chk_mat_esval"] == "NO" else (
                    "SI — Inf. mejor" if c["rescate"].startswith("SI") else "—"
                ),
            ]
        )

    table = doc.add_table(rows=1 + len(body), cols=len(HEADERS))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(HEADERS):
        _set_cell_text(
            table.rows[0].cells[i],
            h.replace("\n", " "),
            bold=True,
            size=8,
            color=RGBColor(255, 255, 255),
            fill=HEADER_FILL,
        )
    for r_i, vals in enumerate(body):
        c = rows_out[r_i]
        highlight = c["chk_mat_esval"] == "NO"
        for c_i, val in enumerate(vals):
            if c_i == 6:
                fill = _color_chk(c["chk_mat_esval"])
            elif c_i == 9:
                fill = _color_chk(c["chk_inf_esval"])
            elif c_i == 10:
                if "SI" in str(val) and "n/a" not in str(val):
                    fill = "C6E0B4"
                elif val == "NO":
                    fill = "F4C7C3"
                elif "parcial" in str(val):
                    fill = "FFF2CC"
                else:
                    fill = "F2F2F2"
            elif highlight:
                fill = "FCE4D6"
            elif c.get("transicion"):
                fill = ROW_ALERTA_FILL
            else:
                fill = "FFFFFF"
            _set_cell_text(table.rows[r_i + 1].cells[c_i], val, size=8, fill=fill)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(chart), width=Inches(10.2))

    # Tabla resumen solo NO-cuadran
    doc.add_paragraph()
    h2 = doc.add_paragraph()
    hr = h2.add_run("Solo meses donde Matriz NO cuadra con la cuenta ESVAL")
    hr.bold = True
    hr.font.size = Pt(12)
    hr.font.color.rgb = RGBColor(31, 78, 121)

    headers2 = [
        "Mes",
        "Cuenta ESVAL",
        "Matriz",
        "Dif % Mat",
        "Inferior",
        "Dif % Inf",
        "¿Inf. rescata?",
    ]
    body2 = []
    for c in no_cuadran:
        body2.append(
            [
                c["mes"],
                _fmt_m3(c["esval"], 0),
                _fmt_m3(c["matriz"], 1),
                _fmt_pct_s(c["pct_mat_esval"]),
                _fmt_m3(c["inferior"], 1),
                _fmt_pct_s(c["pct_inf_esval"]),
                "NO",
            ]
        )
    t2 = doc.add_table(rows=1 + len(body2), cols=len(headers2))
    t2.style = "Table Grid"
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers2):
        _set_cell_text(
            t2.rows[0].cells[i], h, bold=True, size=9, color=RGBColor(255, 255, 255), fill=HEADER_FILL
        )
    for r_i, vals in enumerate(body2):
        for c_i, val in enumerate(vals):
            fill = "F4C7C3" if c_i in (3, 5, 6) else "FFFFFF"
            _set_cell_text(t2.rows[r_i + 1].cells[c_i], val, size=9, fill=fill)

    concl = doc.add_paragraph()
    concl.add_run(
        "Conclusión operativa: el desfase post-marzo no se explica “cambiando” el "
        "punto de cotejo a Estanque Inferior. La brecha es WES (Matriz e Inferior) "
        "frente al medidor de turbina de la boleta ESVAL. En post-transición (May–Ago) "
        "Inferior ≈ Matriz; ambos quedan por debajo de la cuenta."
    ).font.size = Pt(10)

    docx_path = out_dir / "Esval_meses_no_cuadran_vs_Estanque_Inferior.docx"
    doc.save(str(docx_path))

    # --- PDF resumen ---
    fig = plt.figure(figsize=(15.5, 10.0))
    fig.suptitle(
        "Meses ESVAL ≠ Matriz — ¿cuadran con Estanque Inferior?",
        fontsize=15,
        fontweight="bold",
        color=WES_BLUE,
        y=0.97,
        x=0.03,
        ha="left",
    )
    fig.text(
        0.03,
        0.93,
        "VEREDICTO: NO. Ninguno de Mar–Jul 2026 (Matriz≠ESVAL) cuadra con Inferior. "
        "Fuente: delta turbina ESVAL vs WES día completo Chile.",
        fontsize=10,
        color="#C00000",
        fontweight="bold",
    )
    ax = fig.add_axes([0.03, 0.42, 0.94, 0.48])
    ax.axis("off")
    headers_pdf = [
        "Mes",
        "Cuenta\nESVAL",
        "Matriz\nWES",
        "Dif %\nMat",
        "¿Mat\ncuadra?",
        "Estanque\nInf.",
        "Dif %\nInf",
        "¿Inf\ncuadra?",
        "¿Inf.\nrescata?",
    ]
    cell = [headers_pdf]
    for c in no_cuadran:
        cell.append(
            [
                c["mes"],
                _fmt_m3(c["esval"], 0),
                _fmt_m3(c["matriz"], 1),
                _fmt_pct_s(c["pct_mat_esval"]),
                c["chk_mat_esval"],
                _fmt_m3(c["inferior"], 1),
                _fmt_pct_s(c["pct_inf_esval"]),
                c["chk_inf_esval"],
                "NO",
            ]
        )
    # Totales de esos meses
    te = sum(c["esval"] for c in no_cuadran)
    tm = sum(c["matriz"] for c in no_cuadran)
    ti = sum(c["inferior"] for c in no_cuadran)
    cell.append(
        [
            "SUMA",
            _fmt_m3(te, 0),
            _fmt_m3(tm, 1),
            _fmt_pct_s((tm - te) / te * 100 if te else 0),
            "",
            _fmt_m3(ti, 1),
            _fmt_pct_s((ti - te) / te * 100 if te else 0),
            "",
            "NO",
        ]
    )
    col_w = [0.08, 0.11, 0.11, 0.10, 0.10, 0.12, 0.10, 0.10, 0.12]
    tb = ax.table(cellText=cell, loc="center", cellLoc="center", colWidths=col_w)
    tb.auto_set_font_size(False)
    tb.set_fontsize(9)
    tb.scale(1, 1.7)
    n_rows, n_cols = len(cell), len(headers_pdf)
    for c in range(n_cols):
        h = tb[0, c]
        h.set_facecolor(WES_BLUE)
        h.get_text().set_color("white")
        h.get_text().set_fontweight("bold")
    for r in range(1, n_rows - 1):
        for c in range(n_cols):
            if c in (3, 4, 6, 7, 8):
                tb[r, c].set_facecolor("#F4C7C3")
            else:
                tb[r, c].set_facecolor("white")
    for c in range(n_cols):
        tb[n_rows - 1, c].set_facecolor("#D6E3F0")
        tb[n_rows - 1, c].get_text().set_fontweight("bold")

    axc = fig.add_axes([0.06, 0.04, 0.88, 0.35])
    axc.imshow(plt.imread(str(chart)))
    axc.axis("off")
    pdf_path = out_dir / "Esval_meses_no_cuadran_vs_Estanque_Inferior.pdf"
    fig.savefig(pdf_path, dpi=170)
    plt.close(fig)

    payload = {
        "pregunta": (
            "Meses en que cuenta ESVAL no calza con Matriz — "
            "¿cuadran con Estanque Inferior?"
        ),
        "veredicto": "NO",
        "fuente": str(SRC),
        "meses_matriz_no_cuadra": [
            {
                "mes": c["mes"],
                "boleta": c["boleta"],
                "esval": c["esval"],
                "matriz": c["matriz"],
                "pct_mat_esval": c["pct_mat_esval"],
                "inferior": c["inferior"],
                "pct_inf_esval": c["pct_inf_esval"],
                "chk_inf_esval": c["chk_inf_esval"],
                "rescata": False,
            }
            for c in no_cuadran
        ],
        "nota_feb": (
            "Único mes donde Inferior cuadra con ESVAL (±4 %) "
            "mientras Matriz queda cerca (−10 %): Feb-26."
        ),
        "suma_no_cuadran": {"esval": te, "matriz": tm, "inferior": ti},
    }
    (out_dir / "esval_no_cuadran_vs_inferior.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )

    print(f"[OK] {pdf_path}")
    print(f"[OK] {docx_path}")
    print("Meses Matriz≠ESVAL:")
    for c in no_cuadran:
        print(
            f"  {c['mes']}: ESVAL {c['esval']:.0f}  Mat {c['matriz']:.1f} "
            f"({c['pct_mat_esval']:+.1f}%)  Inf {c['inferior']:.1f} "
            f"({c['pct_inf_esval']:+.1f}%)  rescata=NO"
        )
    print(f"OUT_DIR={out_dir}")


if __name__ == "__main__":
    main()
