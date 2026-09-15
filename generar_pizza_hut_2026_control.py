"""Pizza Hut 2026: ¿hubo control? Calendario febrero + detalle 16-02."""

from __future__ import annotations

import calendar
import csv
from datetime import date, datetime, timedelta
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import FancyBboxPatch, Patch
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Cm, Inches, Pt, RGBColor

from analizar_pizza_hut_cero_nocturno import _fetch_dia, _fmt_fecha
from generar_pizza_hut_enero_control import DOW, DOW_LARGO, clasificar
from wes_estilo_graficos_app import guardar_grafico_horario_24h_app

COLOR_WES = "1F4E79"
MESES = [
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


def calendario_mes(dias: list, year: int, month: int, out: Path, titulo: str) -> Path:
    nd = calendar.monthrange(year, month)[1]
    first = date(year, month, 1)
    offset = first.weekday()
    rows_needed = (offset + nd - 1) // 7 + 1
    fig, ax = plt.subplots(figsize=(10.2, 3.2 + rows_needed * 0.95))
    fig.patch.set_facecolor("white")
    ax.set_xlim(0, 7)
    ax.set_ylim(0, rows_needed + 1.3)
    ax.axis("off")
    ax.set_title(titulo, fontsize=14, fontweight="bold", color="#1F4E79", pad=8)
    colores = {"completo": "#2E7D32", "parcial": "#EF6C00", "no": "#90A4AE", "sin_datos": "#EEEEEE"}
    by = {d["fecha"]: d for d in dias}
    top = rows_needed + 0.55
    for i, name in enumerate(DOW):
        ax.text(i + 0.5, top, name, ha="center", va="center", fontsize=10, fontweight="bold")
    for day in range(1, nd + 1):
        d = date(year, month, day)
        cell = offset + day - 1
        col = cell % 7
        row = rows_needed - 1 - cell // 7
        info = by.get(d, {"estado": "sin_datos", "inicio": "", "fin": ""})
        color = colores[info["estado"]]
        ax.add_patch(
            FancyBboxPatch(
                (col + 0.06, row + 0.08),
                0.88,
                0.84,
                boxstyle="round,pad=0.02,rounding_size=0.08",
                facecolor=color,
                edgecolor="white",
                linewidth=1.5,
            )
        )
        txt = "white" if info["estado"] in {"completo", "parcial"} else "#333333"
        ax.text(col + 0.5, row + 0.58, str(day), ha="center", va="center", fontsize=13, fontweight="bold", color=txt)
        if info["estado"] == "completo":
            ax.text(col + 0.5, row + 0.28, "SÍ cortó", ha="center", va="center", fontsize=7, color="white")
        elif info["estado"] == "parcial":
            ax.text(col + 0.5, row + 0.28, f"{info['inicio']}–{info['fin']}", ha="center", va="center", fontsize=6.5, color="white")
    ax.legend(
        handles=[
            Patch(facecolor="#2E7D32", label="SÍ — 00:30 a 05:00 en cero"),
            Patch(facecolor="#EF6C00", label="A medias — cortó más tarde, se mantuvo hasta la mañana"),
            Patch(facecolor="#90A4AE", label="NO — de madrugada siguió el agua"),
        ],
        loc="lower center",
        bbox_to_anchor=(0.5, -0.06),
        ncol=1,
        frameon=False,
        fontsize=9,
    )
    fig.tight_layout()
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out, dpi=160, bbox_inches="tight")
    plt.close(fig)
    return out


def barras_mes(dias: list, titulo: str, out: Path) -> Path:
    xs = [d["fecha"].day for d in dias]
    ys = [{"completo": 2, "parcial": 1, "no": 0, "sin_datos": 0}[d["estado"]] for d in dias]
    colors = [{"completo": "#2E7D32", "parcial": "#EF6C00", "no": "#90A4AE", "sin_datos": "#EEEEEE"}[d["estado"]] for d in dias]
    fig, ax = plt.subplots(figsize=(11.2, 3.8))
    ax.bar(xs, ys, color=colors, width=0.85)
    ax.set_xticks(xs)
    ax.set_xticklabels(xs, fontsize=8)
    ax.set_yticks([0, 1, 2])
    ax.set_yticklabels(["No cortó", "Cortó a medias", "Cortó 00:30–05:00"])
    ax.set_xlabel("Día")
    ax.set_title(titulo, fontweight="bold", color="#1F4E79")
    ax.set_ylim(0, 2.4)
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(out, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return out


def barras_2026(conteo: dict, out: Path) -> Path:
    keys = [f"2026-{m:02d}" for m in range(1, 10)]
    xs = range(len(keys))
    vals = [conteo.get(k, 0) for k in keys]
    labels = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep"]
    fig, ax = plt.subplots(figsize=(10.5, 4.0))
    colors = ["#90A4AE"] * 5 + ["#EF6C00"] * 2 + ["#2E7D32"] * 2
    ax.bar(list(xs), vals, color=colors)
    ax.set_xticks(list(xs))
    ax.set_xticklabels(labels)
    ax.set_ylabel("Noches con corte 00:30–05:00 completo")
    ax.set_title("Pizza Hut 2026 — noches con control completo, por mes", fontweight="bold", color="#1F4E79")
    ax.grid(axis="y", alpha=0.3)
    for i, v in enumerate(vals):
        ax.text(i, v + 0.4, str(v), ha="center", fontsize=9)
    fig.tight_layout()
    fig.savefig(out, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return out


def _shade(cell, hex_color: str) -> None:
    cell._tc.get_or_add_tcPr().append(
        parse_xml(
            f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            f'w:val="clear" w:fill="{hex_color}"/>'
        )
    )


def _cell(cell, text, *, fill=None, bold=False, size=10, color=None) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(str(text))
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = "Calibri"
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    if fill:
        _shade(cell, fill)


def cargar_mes(year: int, month: int) -> list:
    nd = calendar.monthrange(year, month)[1]
    out = []
    d = date(year, month, 1)
    while d.day <= nd and d.month == month:
        _, horas = _fetch_dia(d)
        out.append(clasificar(d, horas))
        d += timedelta(days=1)
    return out


def main() -> int:
    out_dir = Path("reports/Puntos_En_Cero/Pizza_Hut/2026")
    gra = out_dir / "graficos"
    gra.mkdir(parents=True, exist_ok=True)

    print("[INFO] Enero y febrero 2026...")
    ene = cargar_mes(2026, 1)
    feb = cargar_mes(2026, 2)

    cal_e = calendario_mes(ene, 2026, 1, gra / "calendario_enero_2026.png", "Pizza Hut — enero 2026")
    cal_f = calendario_mes(feb, 2026, 2, gra / "calendario_febrero_2026.png", "Pizza Hut — febrero 2026")
    bar_f = barras_mes(feb, "Febrero 2026: ¿operó el control esa noche?", gra / "barras_febrero_2026.png")

    # 16 feb vs vecinos que SÍ
    perfiles = []
    for dia, titulo in [
        (date(2026, 2, 16), "16-02-2026 — NO completo: a las 00:30 todavía había agua (cortó ~03:00)"),
        (date(2026, 2, 13), "13-02-2026 — SÍ cortó 00:00 a 08:00"),
        (date(2026, 2, 17), "17-02-2026 — SÍ cortó 01:00 a 08:00"),
        (date(2026, 1, 25), "25-01-2026 — SÍ (uno de los 2 días de enero)"),
    ]:
        horas = next(x["horas"] for x in (ene + feb) if x["fecha"] == dia)
        p = gra / f"perfil_{dia.isoformat()}.png"
        guardar_grafico_horario_24h_app(horas, p, titulo=f"Pizza Hut — {titulo}", subtitulo="m³/h hora Chile. Rojo = 00:00–06:00.")
        perfiles.append(p)

    conteo = {}
    csv_cero = Path("reports/Puntos_En_Cero/Pizza_Hut/pizza_hut_noches_cero_0030_0500.csv")
    with csv_cero.open(encoding="utf-8") as f:
        for row in csv.DictReader(f, delimiter=";"):
            if row["fecha"].startswith("2026-"):
                k = row["fecha"][:7]
                conteo[k] = conteo.get(k, 0) + 1
    bar_y = barras_2026(conteo, gra / "barras_2026_completo.png")

    f16 = next(x for x in feb if x["fecha"] == date(2026, 2, 16))
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    docx_path = out_dir / f"Pizza_Hut_control_2026_{stamp}.docx"
    pdf_path = out_dir / f"Pizza_Hut_control_2026_{stamp}.pdf"

    doc = Document()
    for s in doc.sections:
        s.top_margin = Cm(1.5)
        s.bottom_margin = Cm(1.5)
        s.left_margin = Cm(1.7)
        s.right_margin = Cm(1.7)
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("PIZZA HUT 2026 — ¿hubo control nocturno?")
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor.from_string(COLOR_WES)

    doc.add_paragraph(
        "Sí hay control en 2026, pero no desde enero. "
        "Enero y febrero: casi no. El 16 de febrero no cortó 00:30–05:00: "
        "a las 00:00 había 1,11 m³/h y a la 01:00 0,85 m³/h. Recién a las 03:00 quedó en cero hasta las 08:00. "
        "El control estable (todas las noches) parte el 9 de agosto 2026."
    )

    h = doc.add_heading("16 de febrero 2026", 1)
    for run in h.runs:
        run.font.color.rgb = RGBColor.from_string(COLOR_WES)
    tbl = doc.add_table(rows=2, cols=4)
    tbl.style = "Table Grid"
    for i, x in enumerate(["Día", "¿Control 00:30–05:00?", "Corte real", "Agua a las 00–02"]):
        _cell(tbl.rows[0].cells[i], x, fill=COLOR_WES, bold=True, size=10, color="FFFFFF")
    _cell(tbl.rows[1].cells[0], "lunes 16-02-2026", fill="FFE0B2", bold=True)
    _cell(tbl.rows[1].cells[1], "NO (a medias)", fill="FFE0B2")
    _cell(tbl.rows[1].cells[2], f"{f16['inicio']} → {f16['fin']}", fill="FFE0B2")
    _cell(tbl.rows[1].cells[3], "1,11 / 0,85 / 0,03 m³/h", fill="FFE0B2")

    doc.add_paragraph("Al lado: el 13 y el 17 de febrero SÍ cortaron toda la madrugada.")

    pic = doc.add_paragraph()
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.add_run().add_picture(str(cal_f), width=Inches(6.3))
    pic = doc.add_paragraph()
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.add_run().add_picture(str(bar_f), width=Inches(6.3))

    h = doc.add_heading("Enero 2026 (el ejemplo que recordabas)", 1)
    for run in h.runs:
        run.font.color.rgb = RGBColor.from_string(COLOR_WES)
    doc.add_paragraph("Solo 2 noches completas: domingo 25 y viernes 30. El resto no, o cortó tarde.")
    pic = doc.add_paragraph()
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.add_run().add_picture(str(cal_e), width=Inches(6.3))

    h = doc.add_heading("2026 completo", 1)
    for run in h.runs:
        run.font.color.rgb = RGBColor.from_string(COLOR_WES)
    doc.add_paragraph(
        "Ene–may: no operativo (2 a 6 noches sueltas). "
        "Jun–jul: empieza. "
        "Desde 09-08-2026: sí, todas las noches ~00:00 a 07:00."
    )
    pic = doc.add_paragraph()
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.add_run().add_picture(str(bar_y), width=Inches(6.3))

    h = doc.add_heading("Perfiles horarios", 1)
    for run in h.runs:
        run.font.color.rgb = RGBColor.from_string(COLOR_WES)
    for pth in perfiles:
        pic = doc.add_paragraph()
        pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic.add_run().add_picture(str(pth), width=Inches(6.2))

    doc.save(str(docx_path))

    with PdfPages(pdf_path) as pdf:
        fig = plt.figure(figsize=(8.27, 11.69))
        ax = fig.add_subplot(111)
        ax.axis("off")
        ax.text(0.5, 0.95, "PIZZA HUT 2026 — control nocturno", ha="center", fontsize=15, color="#1F4E79", fontweight="bold")
        ax.text(
            0.08,
            0.88,
            "16-02-2026: NO cortó 00:30–05:00.\n"
            "  00:00 = 1,11 m³/h   01:00 = 0,85   02:00 = 0,03\n"
            "  Recién 03:00–08:00 en cero.\n\n"
            "Enero: solo 25 y 30 con corte completo.\n"
            "Febrero: solo 5, 13 y 17 con corte completo.\n\n"
            "Control estable: desde el 09-08-2026 hasta hoy.",
            ha="left",
            va="top",
            fontsize=11,
            family="DejaVu Sans",
        )
        img = plt.imread(str(cal_f))
        ax2 = fig.add_axes([0.06, 0.04, 0.88, 0.48])
        ax2.imshow(img)
        ax2.axis("off")
        pdf.savefig(fig)
        plt.close(fig)

        for pth in [cal_e, bar_f, bar_y, *perfiles]:
            fig = plt.figure(figsize=(8.27, 11.69))
            ax = fig.add_subplot(111)
            ax.imshow(plt.imread(str(pth)))
            ax.axis("off")
            fig.tight_layout()
            pdf.savefig(fig)
            plt.close(fig)

    print(f"DOCX={docx_path}")
    print(f"PDF={pdf_path}")
    print(f"CAL_FEB={cal_f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
