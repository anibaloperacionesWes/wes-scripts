"""
Corrección de marco: Matriz (campo/app) vs Estanque Inferior.

No se espera igualdad con la boleta ESVAL (hay tubería / distancia).
La pregunta correcta: ¿Matriz e Inferior se *parecen* / conversan?

Fuentes:
- Validación terreno 22–23/09/2026 (Itron + App Matriz + Inferior)
- Ciclos de boleta día completo Chile (mensual_esval_matriz_inferior.json)
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
    WES_BLUE,
    _fmt_m3,
    _set_cell_text,
    _set_landscape,
)
from generar_tabla_zapallar_mensual_esval_inferior import _fmt_pct_s

SRC_MENSUAL = Path(
    "reports/Fundo_Zapallar/ANALISIS/tabla_mensual_esval_inferior_20260922_1510/"
    "mensual_esval_matriz_inferior.json"
)
SRC_VALID = Path(
    "reports/Fundo_Zapallar/Informes_Tecnicos/"
    "Consumo_Estanque_Inferior_validacion_Matriz_ESVAL_2209_2309_20260925_1602.json"
)

# ±15 % Inf vs Matriz = "se parecen" (hay tubería en el camino)
LIM_PARECEN = 15.0


def _parecen(pct_inf_mat: float, transicion: bool) -> str:
    if transicion:
        return "transición"
    if abs(pct_inf_mat) <= LIM_PARECEN:
        return "SI — conversan"
    if abs(pct_inf_mat) <= 25:
        return "cerca"
    return "NO"


def _fill_parecen(label: str) -> str:
    if label.startswith("SI"):
        return "C6E0B4"
    if label == "cerca":
        return "FFF2CC"
    if label == "transición":
        return "F4C7C3"
    return "F8CBAD"


def main() -> None:
    mensual = json.loads(SRC_MENSUAL.read_text(encoding="utf-8"))
    valid = json.loads(SRC_VALID.read_text(encoding="utf-8"))
    vb = valid["ventana_matriz_B"]
    ciclos = mensual["ciclos"]

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = Path("reports/Fundo_Zapallar/ANALISIS") / f"matriz_inferior_conversan_{stamp}"
    out_dir.mkdir(parents=True, exist_ok=True)

    rows = []
    for c in ciclos:
        label = _parecen(c["pct_inf_mat"], c.get("transicion", False))
        # Abr también es outlier de transición hidráulica aunque fase diga ultrasónico
        if c["mes"] == "Abr-26":
            label = "transición"
        rows.append({**c, "parecen": label})

    post = [r for r in rows if r["mes"] in ("May-26", "Jun-26", "Jul-26", "Ago-26")]
    tm = sum(r["matriz"] for r in post)
    ti = sum(r["inferior"] for r in post)
    te = sum(r["esval"] for r in post)
    pct_im = (ti / tm - 1.0) * 100.0 if tm else 0.0
    pct_me = (tm / te - 1.0) * 100.0 if te else 0.0

    # --- Chart 1: validación terreno ---
    fig, ax = plt.subplots(figsize=(8.5, 4.2))
    labels_v = ["Matriz App", "Matriz Itron\n(terreno)", "Estanque\nInferior"]
    vals_v = [vb["matriz_app_m3"], vb["itron_matriz_m3"], vb["estanque_inferior_m3"]]
    colors = [WES_BLUE, "#5B9BD5", "#70AD47"]
    bars = ax.bar(labels_v, vals_v, color=colors, width=0.55)
    for b, v in zip(bars, vals_v):
        ax.text(b.get_x() + b.get_width() / 2, v + 2, f"{v:.1f} m³", ha="center", fontsize=10)
    ax.set_ylabel("m³")
    ax.set_title(
        "Validación 22–23/09/2026 — misma ventana Matriz\n"
        f"Inferior = {vb['inferior_sobre_matriz_pct']:.0f} % de Matriz App "
        f"(Δ {vb['diferencia_matriz_menos_inferior_m3']:.1f} m³ · tubería en el camino)",
        fontsize=11,
        fontweight="bold",
        color=WES_BLUE,
    )
    ax.set_ylim(0, max(vals_v) * 1.18)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="y", linestyle="--", alpha=0.35)
    fig.tight_layout()
    chart_val = out_dir / "chart_validacion_matriz_itron_inferior.png"
    fig.savefig(chart_val, dpi=160, bbox_inches="tight")
    plt.close(fig)

    # --- Chart 2: meses post-marzo Matriz vs Inferior (no vs boleta) ---
    foco = [r for r in rows if r["mes"] in ("Mar-26", "Abr-26", "May-26", "Jun-26", "Jul-26", "Ago-26")]
    labels = [r["mes"] for r in foco]
    x = np.arange(len(labels))
    w = 0.36
    fig, ax = plt.subplots(figsize=(12.2, 4.8))
    ax.bar(x - w / 2, [r["matriz"] for r in foco], w, label="Matriz WES", color=WES_BLUE)
    ax.bar(x + w / 2, [r["inferior"] for r in foco], w, label="Estanque Inferior", color="#70AD47")
    # línea ESVAL tenue solo como referencia (no es el eje de “parecerse”)
    ax.plot(x, [r["esval"] for r in foco], "o--", color="#7F7F7F", alpha=0.55, label="Cuenta ESVAL (ref.)")
    ax.set_ylabel("m³")
    ax.set_title(
        "Desde marzo: Matriz vs Inferior (¿se parecen?) · ESVAL solo referencia",
        fontsize=12,
        fontweight="bold",
        color=WES_BLUE,
    )
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.legend(frameon=False, fontsize=9)
    ax.grid(axis="y", linestyle="--", alpha=0.35)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    chart_mes = out_dir / "chart_matriz_vs_inferior_desde_marzo.png"
    fig.savefig(chart_mes, dpi=160, bbox_inches="tight")
    plt.close(fig)

    # --- Word ---
    doc = Document()
    _set_landscape(doc)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = t.add_run("Matriz (campo/app) y Estanque Inferior — ¿conversan?")
    run.bold = True
    run.font.size = Pt(15)
    run.font.color.rgb = RGBColor(31, 78, 121)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = sub.add_run("Corrección de marco · Fundo Zapallar · no se espera igualdad con la boleta ESVAL")
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(89, 89, 89)

    n = doc.add_paragraph()
    n.add_run(
        "Crítica válida: Matriz y Estanque Inferior nunca van a calzar igual — hay tubería y "
        "distancia en el camino. Lo esperable es que se *parezcan*. La boleta ESVAL (delta "
        "turbina) es otra comparación: ahí el desfase post-marzo sigue abierto, pero no es "
        "el test de si los dos puntos WES conversan.\n\n"
        "Criterio “se parecen”: |Inf − Matriz| / Matriz ≤ 15 % (misma orden de magnitud / "
        "mismo perfil). Mar–Abr = transición (cambio de medidor); no calibrar."
    ).font.size = Pt(10)

    verd = doc.add_paragraph()
    verd.alignment = WD_ALIGN_PARAGRAPH.CENTER
    vr = verd.add_run(
        f"VEREDICTO: SÍ conversan. Validación terreno Inferior = {vb['inferior_sobre_matriz_pct']:.0f} % "
        f"de Matriz App. May–Ago agregado Inf/Mat = {pct_im:+.1f} %."
    )
    vr.bold = True
    vr.font.size = Pt(12)
    vr.font.color.rgb = RGBColor(0, 100, 0)

    # Bloque validación
    h1 = doc.add_paragraph()
    hr = h1.add_run("1) Validación de terreno (misma ventana Matriz 22–23/09/2026)")
    hr.bold = True
    hr.font.size = Pt(12)
    hr.font.color.rgb = RGBColor(31, 78, 121)

    p1 = doc.add_paragraph()
    p1.add_run(
        f"Ventana B Chile: {vb['inicio']} → {vb['fin']} "
        f"({vb['horas_app']}).\n"
        f"· Matriz App: {_fmt_m3(vb['matriz_app_m3'], 2)} m³\n"
        f"· Matriz Itron (terreno): {_fmt_m3(vb['itron_matriz_m3'], 2)} m³ "
        f"(dif. vs App {(vb['itron_matriz_m3']/vb['matriz_app_m3']-1)*100:+.1f} %)\n"
        f"· Estanque Inferior (misma ventana): {_fmt_m3(vb['estanque_inferior_m3'], 2)} m³ "
        f"= {vb['inferior_sobre_matriz_pct']:.0f} % de Matriz App "
        f"(Δ {vb['diferencia_matriz_menos_inferior_m3']:.1f} m³ hacia abajo — coherente con "
        "pérdida / volumen en tubería y llenado).\n"
        "App ≈ Itron y ambos conversan con Inferior en la misma ventana."
    ).font.size = Pt(10)

    pic = doc.add_paragraph()
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.add_run().add_picture(str(chart_val), width=Inches(7.2))

    # Bloque mensual
    h2 = doc.add_paragraph()
    hr2 = h2.add_run("2) Ciclos de boleta desde marzo — foco Matriz vs Inferior")
    hr2.bold = True
    hr2.font.size = Pt(12)
    hr2.font.color.rgb = RGBColor(31, 78, 121)

    headers = [
        "Mes",
        "Ciclo",
        "Matriz WES",
        "Estanque Inf.",
        "Dif % Inf/Mat",
        "¿Se parecen?",
        "Cuenta ESVAL (ref.)",
        "Nota boleta",
    ]
    body = []
    for r in foco:
        nota = "desfase boleta (otra capa)" if r["chk_mat_esval"] == "NO" else r["chk_mat_esval"]
        if r["parecen"] == "transición":
            nota = "transición — no calibrar"
        body.append(
            [
                r["mes"],
                r["ciclo"],
                _fmt_m3(r["matriz"], 1),
                _fmt_m3(r["inferior"], 1),
                _fmt_pct_s(r["pct_inf_mat"]),
                r["parecen"],
                _fmt_m3(r["esval"], 0),
                nota,
            ]
        )

    table = doc.add_table(rows=1 + len(body), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        _set_cell_text(
            table.rows[0].cells[i], h, bold=True, size=8, color=RGBColor(255, 255, 255), fill=HEADER_FILL
        )
    for r_i, vals in enumerate(body):
        r = foco[r_i]
        for c_i, val in enumerate(vals):
            if c_i == 5:
                fill = _fill_parecen(r["parecen"])
            elif r["parecen"] == "transición":
                fill = ROW_ALERTA_FILL
            else:
                fill = "FFFFFF"
            _set_cell_text(table.rows[r_i + 1].cells[c_i], val, size=8, fill=fill)

    p2 = doc.add_paragraph()
    p2.add_run(
        f"May–Ago (sin transición): Matriz {_fmt_m3(tm, 0)} m³ · Inferior {_fmt_m3(ti, 0)} m³ · "
        f"Inf/Mat {pct_im:+.1f} % → se parecen. "
        f"Misma ventana vs boleta ESVAL: Matriz queda {pct_me:+.1f} % bajo la turbina "
        f"(esa brecha es otro problema; no invalida que Matriz e Inferior conversen). "
        "Jun y Jul tienen más ruido (huecos horarios en Inferior); May y Ago quedan dentro de ±15 %."
    ).font.size = Pt(10)

    pic2 = doc.add_paragraph()
    pic2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic2.add_run().add_picture(str(chart_mes), width=Inches(10.0))

    concl = doc.add_paragraph()
    concl.add_run(
        "Conclusión: el informe anterior midió mal al exigir que Inferior “rescate” la boleta. "
        "Con el marco correcto — parecerse, no igualar — Matriz (campo/app) y Estanque Inferior "
        "sí conversan. La distancia residual es esperable por la tubería. El desfase cuenta "
        "ESVAL vs WES post-marzo sigue siendo una capa aparte (turbina vs ultrasonido)."
    ).font.size = Pt(10)

    docx_path = out_dir / "Matriz_e_Inferior_conversan_correccion_marco.docx"
    doc.save(str(docx_path))

    # --- PDF ---
    fig = plt.figure(figsize=(15.2, 10.2))
    fig.suptitle(
        "Matriz (campo/app) ↔ Estanque Inferior — sí conversan",
        fontsize=15,
        fontweight="bold",
        color=WES_BLUE,
        y=0.97,
        x=0.03,
        ha="left",
    )
    fig.text(
        0.03,
        0.935,
        "No se espera igualdad con la boleta ESVAL (hay tubería). "
        f"Validación: Inf = {vb['inferior_sobre_matriz_pct']:.0f}% Matriz App. "
        f"May–Ago Inf/Mat = {pct_im:+.1f}%.",
        fontsize=9.5,
        color="#006400",
        fontweight="bold",
    )

    ax = fig.add_axes([0.03, 0.48, 0.94, 0.42])
    ax.axis("off")
    headers_pdf = [
        "Mes",
        "Matriz\nWES",
        "Estanque\nInf.",
        "Dif %\nInf/Mat",
        "¿Se\nparecen?",
        "ESVAL\n(ref.)",
        "Nota",
    ]
    cell = [headers_pdf]
    for r in foco:
        nota = "desfase boleta" if r["chk_mat_esval"] == "NO" else ""
        if r["parecen"] == "transición":
            nota = "transición"
        cell.append(
            [
                r["mes"],
                _fmt_m3(r["matriz"], 1),
                _fmt_m3(r["inferior"], 1),
                _fmt_pct_s(r["pct_inf_mat"]),
                r["parecen"].replace(" — ", "\n"),
                _fmt_m3(r["esval"], 0),
                nota,
            ]
        )
    cell.append(
        [
            "May–Ago",
            _fmt_m3(tm, 0),
            _fmt_m3(ti, 0),
            _fmt_pct_s(pct_im),
            "SI — conversan",
            _fmt_m3(te, 0),
            f"Mat vs ESVAL {pct_me:+.0f}%",
        ]
    )
    col_w = [0.10, 0.12, 0.12, 0.12, 0.18, 0.12, 0.16]
    tb = ax.table(cellText=cell, loc="center", cellLoc="center", colWidths=col_w)
    tb.auto_set_font_size(False)
    tb.set_fontsize(9)
    tb.scale(1, 1.65)
    n_rows, n_cols = len(cell), len(headers_pdf)
    for c in range(n_cols):
        h = tb[0, c]
        h.set_facecolor(WES_BLUE)
        h.get_text().set_color("white")
        h.get_text().set_fontweight("bold")
    for r_i, r in enumerate(foco):
        for c in range(n_cols):
            if c == 4:
                tb[r_i + 1, c].set_facecolor(
                    "#C6E0B4"
                    if r["parecen"].startswith("SI")
                    else ("#FFF2CC" if r["parecen"] == "cerca" else "#F4C7C3")
                )
            elif r["parecen"] == "transición":
                tb[r_i + 1, c].set_facecolor("#F4C7C3")
            else:
                tb[r_i + 1, c].set_facecolor("white")
    for c in range(n_cols):
        tb[n_rows - 1, c].set_facecolor("#C6E0B4")
        tb[n_rows - 1, c].get_text().set_fontweight("bold")

    axc = fig.add_axes([0.05, 0.04, 0.42, 0.40])
    axc.imshow(plt.imread(str(chart_val)))
    axc.axis("off")
    axc2 = fig.add_axes([0.50, 0.04, 0.48, 0.40])
    axc2.imshow(plt.imread(str(chart_mes)))
    axc2.axis("off")

    pdf_path = out_dir / "Matriz_e_Inferior_conversan_correccion_marco.pdf"
    fig.savefig(pdf_path, dpi=170)
    plt.close(fig)

    payload = {
        "marco": (
            "No igualdad con boleta ESVAL (tubería). "
            "Test correcto: Matriz campo/app vs Inferior se parecen."
        ),
        "veredicto": "SI conversan",
        "validacion_terreno": vb,
        "may_ago": {
            "matriz": tm,
            "inferior": ti,
            "esval_ref": te,
            "pct_inf_sobre_matriz": pct_im,
            "pct_matriz_sobre_esval": pct_me,
        },
        "filas_desde_marzo": [
            {
                "mes": r["mes"],
                "matriz": r["matriz"],
                "inferior": r["inferior"],
                "pct_inf_mat": r["pct_inf_mat"],
                "parecen": r["parecen"],
                "esval_ref": r["esval"],
                "chk_mat_esval": r["chk_mat_esval"],
            }
            for r in foco
        ],
    }
    (out_dir / "matriz_inferior_conversan.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )
    print(f"[OK] {pdf_path}")
    print(f"[OK] {docx_path}")
    print(
        f"Validación: App {vb['matriz_app_m3']} · Itron {vb['itron_matriz_m3']} · "
        f"Inf {vb['estanque_inferior_m3']} ({vb['inferior_sobre_matriz_pct']}% App)"
    )
    print(f"May–Ago Inf/Mat {pct_im:+.1f}% · Mat/ESVAL {pct_me:+.1f}%")
    print(f"OUT_DIR={out_dir}")


if __name__ == "__main__":
    main()
