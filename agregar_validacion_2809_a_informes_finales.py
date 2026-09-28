"""
Agrega la validación con fotos del 28-09-2026 a los informes FINAL
(Matriz ESVAL + Etapa 5), con estilo visual alineado al informe FINAL:
  - Navy #1F4788 / light #D6E3F0
  - KPI verde #E8F5E9 / #1B5E20
  - Filas alternadas #F5F8FB
  - Fotos con caption azul
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
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from googleapiclient.http import MediaFileUpload

from wes_google_drive import obtener_servicio_drive

ROOT = Path("reports/Fundo_Zapallar/Informes_Tecnicos")
SRC_DIR = ROOT / "_drive_edit"
VAL_DIR = ROOT / "validacion_terreno_2809_20260928_1726"
OUT_DIR = ROOT / f"_final_cliente_2809_{datetime.now().strftime('%Y%m%d_%H%M')}"

DRIVE_MATRIZ = "16IlboR89inKpImJSL4j4vNt9Lt0H4MHsf112l5Jjkh4"
DRIVE_ETAPA5 = "1bJqXMrPc8zSuhfx70YGgFmgxbrMTISipP6rn4Gw4BRU"

# Paleta FINAL (igual que generar_informe_cambio_memoria_etapa5 / calidad señal)
COLOR_TITULO = RGBColor(0x1F, 0x47, 0x88)
COLOR_META = RGBColor(0x64, 0x6E, 0x78)
COLOR_TEXTO = RGBColor(0x28, 0x28, 0x28)
COLOR_KPI = RGBColor(0x1B, 0x5E, 0x20)
COLOR_WHITE = RGBColor(255, 255, 255)
HEX_NAVY = "1F4788"
HEX_NAVY_LIGHT = "D6E3F0"
HEX_KPI = "E8F5E9"
HEX_ALT = "F5F8FB"
HEX_FOTO = "FAFBFC"
HEX_OK_ROW = "E8F5E9"

# Charts
CHART_NAVY = "#1F4788"
CHART_BLUE = "#5B9BD5"
CHART_GREEN = "#2E7D32"
CHART_ORANGE = "#E67E22"
CHART_GRID = "#E8EEF4"


def _fmt(x: float, dec: int = 2) -> str:
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _shade(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    existing = tc_pr.find(qn("w:shd"))
    if existing is not None:
        tc_pr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    tc_pr.append(shd)


def _set_table_borders(table, color: str = "D0D5DD") -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl.insert(0, tbl_pr)
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is not None:
        tbl_pr.remove(borders)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tbl_pr.append(borders)


def _run(p, text: str, *, bold: bool = False, size: float = 11, color: RGBColor | None = None):
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = "Calibri"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    if color:
        r.font.color.rgb = color
    return r


def _heading(doc: Document, text: str, *, size: float = 12) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    _run(p, text, bold=True, size=size, color=COLOR_TITULO)


def _para(doc: Document, text: str, *, size: float = 10.5, color: RGBColor | None = None) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    _run(p, text, size=size, color=color or COLOR_TEXTO)


def _fix_section_margins(doc: Document) -> None:
    for section in doc.sections:
        pg_mar = section._sectPr.find(qn("w:pgMar"))
        if pg_mar is None:
            continue
        for attr, val in list(pg_mar.attrib.items()):
            try:
                int(val)
            except ValueError:
                try:
                    pg_mar.set(attr, str(int(round(float(val)))))
                except ValueError:
                    pg_mar.set(attr, "720")


def _kpi_banner(doc: Document, cards: list[tuple[str, str, str]]) -> None:
    """Franja KPI: header navy + valor verde (#E8F5E9), estilo FINAL."""
    _fix_section_margins(doc)
    t = doc.add_table(rows=2, cols=len(cards))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(t, color=HEX_NAVY)
    for j, (titulo, linea1, linea2) in enumerate(cards):
        c0 = t.rows[0].cells[j]
        c0.paragraphs[0].clear()
        c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(c0.paragraphs[0], titulo, bold=True, size=9, color=COLOR_WHITE)
        _shade(c0, HEX_NAVY)

        c1 = t.rows[1].cells[j]
        c1.paragraphs[0].clear()
        c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = _run(c1.paragraphs[0], linea1 + "\n", bold=True, size=12, color=COLOR_KPI)
        _run(c1.paragraphs[0], linea2, bold=False, size=8, color=COLOR_META)
        _shade(c1, HEX_KPI)
    doc.add_paragraph()


def _table(
    doc: Document,
    headers: list[str],
    rows: list[list[str]],
    *,
    ok_rows: set[int] | None = None,
) -> None:
    ok_rows = ok_rows or set()
    _fix_section_margins(doc)
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(t, color="D0D5DD")
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.paragraphs[0].clear()
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(cell.paragraphs[0], h, bold=True, size=9, color=COLOR_WHITE)
        _shade(cell, HEX_NAVY)
    for r_i, vals in enumerate(rows):
        if r_i in ok_rows:
            fill = HEX_OK_ROW
        else:
            fill = HEX_ALT if r_i % 2 == 1 else "FFFFFF"
        for c_i, val in enumerate(vals):
            cell = t.rows[r_i + 1].cells[c_i]
            cell.paragraphs[0].clear()
            bold = r_i in ok_rows and c_i in (0, 1)
            color = COLOR_KPI if r_i in ok_rows else COLOR_TEXTO
            _run(cell.paragraphs[0], val, bold=bold, size=9, color=color)
            _shade(cell, fill)
    doc.add_paragraph()


def _fotos(doc: Document, izq: Path, cap_izq: str, der: Path, cap_der: str, *, ancho: float = 3.55) -> None:
    _fix_section_margins(doc)
    t = doc.add_table(rows=2, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(t, color=HEX_NAVY_LIGHT)
    for col, (path, cap) in enumerate(((izq, cap_izq), (der, cap_der))):
        cell = t.rows[0].cells[col]
        _shade(cell, HEX_FOTO)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if path.is_file():
            p.add_run().add_picture(str(path), width=Inches(ancho))
        else:
            _run(p, "(foto no disponible)", size=9, color=COLOR_META)

        cell_c = t.rows[1].cells[col]
        _shade(cell_c, HEX_NAVY_LIGHT)
        pc = cell_c.paragraphs[0]
        pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(pc, cap, bold=True, size=9, color=COLOR_TITULO)
    doc.add_paragraph()


def _chart_barras(
    path: Path,
    labels: list[str],
    vals: list[float],
    colors: list[str],
    title: str,
    subtitle: str,
) -> Path:
    fig, ax = plt.subplots(figsize=(8.8, 4.0), facecolor="white")
    ax.set_facecolor("white")
    x = np.arange(len(labels))
    bars = ax.bar(x, vals, color=colors, width=0.52, edgecolor="white", linewidth=0.8)
    for b, v in zip(bars, vals):
        ax.text(
            b.get_x() + b.get_width() / 2,
            v + max(vals) * 0.03,
            f"{v:,.1f}".replace(",", "X").replace(".", ",").replace("X", "."),
            ha="center",
            va="bottom",
            fontsize=11,
            fontweight="bold",
            color="#282828",
        )
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10, color="#282828")
    ax.set_ylabel("m³", fontsize=10, color="#646E78")
    ax.set_title(title, fontsize=12, fontweight="bold", color=CHART_NAVY, pad=8)
    ax.text(0.5, 1.02, subtitle, transform=ax.transAxes, ha="center", fontsize=9, color="#646E78")
    ax.set_ylim(0, max(vals) * 1.22)
    ax.yaxis.grid(True, linestyle="--", alpha=0.45, color=CHART_GRID)
    ax.set_axisbelow(True)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    for spine in ("left", "bottom"):
        ax.spines[spine].set_color("#D0D5DD")
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def _regenerar_charts(data: dict) -> tuple[Path, Path]:
    m, e = data["matriz"], data["etapa5"]
    c_m = _chart_barras(
        VAL_DIR / "chart_validacion_matriz_2809_estilo.png",
        ["Itron\n(terreno)", "Matriz App\nWES", "Estanque\nInferior"],
        [m["itron_delta"], m["app_m3"], m["inferior_m3"]],
        [CHART_BLUE, CHART_NAVY, CHART_GREEN],
        "Matriz ESVAL — continuidad 23/09 → 28/09",
        f"Error Itron vs App {m['error_pct']:.1f}% · Inferior = {m['inferior_sobre_app_pct']:.0f}% App",
    )
    c_e = _chart_barras(
        VAL_DIR / "chart_validacion_etapa5_2809_estilo.png",
        ["Sensus\n(terreno)", "App WES\nEtapa N°5"],
        [e["sensus_delta"], e["app_m3"]],
        [CHART_ORANGE, CHART_NAVY],
        "Etapa N°5 — continuidad 23/09 → 28/09",
        f"Error Sensus vs App {e['error_pct']:.1f}% · {e['estado']}",
    )
    return c_m, c_e


def _insert_before_conclusion(doc: Document, build_section) -> None:
    conclusion_idx = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith("5. Conclusión"):
            conclusion_idx = i
            break
    if conclusion_idx is None:
        raise RuntimeError("No se encontró '5. Conclusión'")

    concl_body = []
    pie = ""
    for p in doc.paragraphs[conclusion_idx + 1 :]:
        t = p.text.strip()
        if t.startswith("WES ·"):
            pie = t
            break
        if t:
            concl_body.append(t)

    clearing = False
    for p in doc.paragraphs:
        if p.text.strip().startswith("5. Conclusión"):
            clearing = True
        if clearing:
            p.clear()

    build_section(doc)

    _heading(doc, "5. Conclusión", size=14)
    # Evitar duplicar el extra de continuidad si ya está
    for t in concl_body:
        if "Continuidad 23–28/09" in t:
            continue
        _para(doc, t, size=10.5)
    stamp = datetime.now().strftime("%d-%m-%Y %H:%M")
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    _run(
        p,
        (pie.rsplit("·", 1)[0] + f"· {stamp}") if pie else f"WES · Informe final · Fundo Zapallar · {stamp}",
        size=9,
        color=COLOR_META,
    )


def _update_header_validaciones(doc: Document, extra: str) -> None:
    for p in doc.paragraphs:
        t = p.text.strip()
        if t.startswith("Validaciones:"):
            # si ya tiene el extra, no duplicar
            base = t.split(" | 23-09-2026")[0]
            p.clear()
            _run(p, f"{base} | {extra}", size=10, color=COLOR_META)
            return


def agregar_matriz(data: dict, chart: Path) -> Path:
    src = SRC_DIR / "Informe_Validacion_Matriz_ESVAL_Fundo_Zapallar_FINAL.docx"
    doc = Document(str(src))
    _fix_section_margins(doc)
    m = data["matriz"]
    _update_header_validaciones(
        doc,
        "23-09-2026 17:00 → 28-09-2026 09:05 (foto Itron vs App)",
    )

    def section(d: Document) -> None:
        _heading(d, "4.3 Continuidad Itron vs app WES (23-09 17:00 → 28-09 09:05)", size=12)
        _kpi_banner(
            d,
            [
                ("Δ Itron (terreno)", f"{_fmt(m['itron_delta'])} m³", "Foto 23/09 → 28/09"),
                ("Δ Matriz App WES", f"{_fmt(m['app_m3'])} m³", "000027-01 · 112 h"),
                ("Error", f"{m['error_pct']:.1f} %", m["estado"]),
                ("Estanque Inferior", f"{_fmt(m['inferior_m3'])} m³", f"{m['inferior_sobre_app_pct']:.0f}% de Matriz"),
            ],
        )
        _para(
            d,
            "Validación de continuidad con lectura fotográfica del Itron del 28-09-2026 09:05 "
            "y la lectura de cierre de la validación previa (23-09 ~17:05). "
            "App WES: horas 17–23 del 23-09 + días 24–27 completos + horas 00–08 del 28-09 "
            "(misma metodología TIME etiqueta). "
            f"Lecturas Itron: {_fmt(m['itron_ini'], 1)} → {_fmt(m['itron_fin'], 1)} m³ "
            f"(Δ {_fmt(m['itron_delta'])} m³). Escala odómetro: 6 ruedas negras + 1 roja.",
            size=10.5,
        )
        _fotos(
            d,
            VAL_DIR / "foto_matriz_itron_2309_1705.jpg",
            f"Itron · {_fmt(m['itron_ini'], 1)} m³ · 23-09-2026 17:05",
            VAL_DIR / "foto_matriz_itron_2809_0905.jpg",
            f"Itron · {_fmt(m['itron_fin'], 1)} m³ · 28-09-2026 09:05",
        )
        _table(
            d,
            ["Concepto", "Valor", "Nota"],
            [
                ["Lectura Itron inicio", f"{_fmt(m['itron_ini'], 1)} m³", "23-09-2026 17:05"],
                ["Lectura Itron fin", f"{_fmt(m['itron_fin'], 1)} m³", "28-09-2026 09:05"],
                ["Δ Itron (terreno)", f"{_fmt(m['itron_delta'])} m³", "Referencia mecánica"],
                ["Δ Matriz App WES", f"{_fmt(m['app_m3'])} m³", "000027-01"],
                ["Error Itron vs App", f"{m['error_pct']:.1f} %", m["estado"]],
                [
                    "Δ Estanque Inferior",
                    f"{_fmt(m['inferior_m3'])} m³",
                    f"{m['inferior_sobre_app_pct']:.0f}% de Matriz App — conversan",
                ],
            ],
            ok_rows={4},
        )
        if chart.is_file():
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(str(chart), width=Inches(6.6))
        _para(
            d,
            f"Error Itron vs app: {m['error_pct']:.1f}% "
            f"(1 − {_fmt(m['itron_delta'])}/{_fmt(m['app_m3'])}). {m['estado']}. "
            f"En la misma ventana el Estanque Inferior registra {_fmt(m['inferior_m3'])} m³ "
            f"({m['inferior_sobre_app_pct']:.0f}% de Matriz App): conversan (hay tubería en el camino).",
            size=10.5,
        )

    _insert_before_conclusion(doc, section)

    for p in doc.paragraphs:
        if "El desvío inicial no se debió" in p.text and "Continuidad 23–28/09" not in p.text:
            _run(
                p,
                f" Continuidad 23–28/09: Itron {_fmt(m['itron_delta'])} m³ vs App "
                f"{_fmt(m['app_m3'])} m³ (error {m['error_pct']:.1f}%, {m['estado']}).",
                size=10.5,
                color=COLOR_TEXTO,
            )
            break

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "Informe_Validacion_Matriz_ESVAL_Fundo_Zapallar_FINAL.docx"
    doc.save(str(out))
    return out


def agregar_etapa5(data: dict, chart: Path) -> Path:
    src = SRC_DIR / "Informe_Validacion_Etapa5_Zapallar_20260924_1136.docx"
    doc = Document(str(src))
    _fix_section_margins(doc)
    e = data["etapa5"]
    _update_header_validaciones(
        doc,
        "23-09-2026 17:00 → 28-09-2026 08:57 (foto Sensus vs App)",
    )

    def section(d: Document) -> None:
        _heading(d, "4.3 Continuidad Sensus vs app WES (23-09 17:00 → 28-09 08:57)", size=12)
        _kpi_banner(
            d,
            [
                ("Δ Sensus (terreno)", f"{_fmt(e['sensus_delta'], 0)} m³", "Foto 23/09 → 28/09"),
                ("Δ App WES Etapa N°5", f"{_fmt(e['app_m3'])} m³", "000027-03 · 112 h"),
                ("Error", f"{e['error_pct']:.1f} %", e["estado"]),
            ],
        )
        _para(
            d,
            "Validación de continuidad con lectura fotográfica Sensus del 28-09-2026 08:57 "
            "y la lectura de cierre previa (23-09 16:54 = 5.177 m³). "
            "App WES: misma ventana horaria que Matriz (h17 del 23-09 → h08 del 28-09). "
            f"Lecturas: {_fmt(e['sensus_ini'], 0)} → {_fmt(e['sensus_fin'], 0)} m³ "
            f"(Δ {_fmt(e['sensus_delta'], 0)} m³; odómetro 005225).",
            size=10.5,
        )
        _fotos(
            d,
            VAL_DIR / "foto_etapa5_sensus_2309_1654.jpg",
            f"Sensus · {_fmt(e['sensus_ini'], 0)} m³ · 23-09-2026 16:54",
            VAL_DIR / "foto_etapa5_sensus_2809_0857.jpg",
            f"Sensus · {_fmt(e['sensus_fin'], 0)} m³ · 28-09-2026 08:57",
        )
        _table(
            d,
            ["Concepto", "Valor", "Nota"],
            [
                ["Lectura Sensus inicio", f"{_fmt(e['sensus_ini'], 0)} m³", "23-09-2026 16:54"],
                ["Lectura Sensus fin", f"{_fmt(e['sensus_fin'], 0)} m³", "28-09-2026 08:57"],
                ["Δ Sensus (terreno)", f"{_fmt(e['sensus_delta'], 0)} m³", "Referencia mecánica"],
                ["Δ App WES Etapa N°5", f"{_fmt(e['app_m3'])} m³", "000027-03"],
                ["Error Sensus vs App", f"{e['error_pct']:.1f} %", e["estado"]],
            ],
            ok_rows={4},
        )
        if chart.is_file():
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(str(chart), width=Inches(5.6))
        _para(
            d,
            f"Error Sensus vs app: {e['error_pct']:.1f}% "
            f"(1 − {_fmt(e['sensus_delta'], 0)}/{_fmt(e['app_m3'])}). {e['estado']}.",
            size=10.5,
        )

    _insert_before_conclusion(doc, section)

    for p in doc.paragraphs:
        if "Se confirma falla de memoria" in p.text and "Continuidad 23–28/09" not in p.text:
            _run(
                p,
                f" Continuidad 23–28/09: Sensus {_fmt(e['sensus_delta'], 0)} m³ vs App "
                f"{_fmt(e['app_m3'])} m³ (error {e['error_pct']:.1f}%, {e['estado']}).",
                size=10.5,
                color=COLOR_TEXTO,
            )
            break

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "Informe_Validacion_Etapa5_Zapallar_FINAL.docx"
    doc.save(str(out))
    return out


def actualizar_google_doc(file_id: str, docx_path: Path) -> dict:
    svc = obtener_servicio_drive()
    media = MediaFileUpload(
        str(docx_path),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        resumable=True,
    )
    meta = (
        svc.files()
        .update(
            fileId=file_id,
            media_body=media,
            fields="id,name,webViewLink,mimeType,modifiedTime",
        )
        .execute()
    )
    return {
        "id": meta["id"],
        "name": meta["name"],
        "web_view_link": meta.get("webViewLink")
        or f"https://docs.google.com/document/d/{meta['id']}/edit",
        "mimeType": meta.get("mimeType"),
        "modifiedTime": meta.get("modifiedTime"),
    }


def main() -> None:
    data = json.loads((VAL_DIR / "validacion_terreno_2809.json").read_text(encoding="utf-8"))
    chart_m, chart_e = _regenerar_charts(data)
    out_m = agregar_matriz(data, chart_m)
    out_e = agregar_etapa5(data, chart_e)
    print(f"[OK] {out_m}")
    print(f"[OK] {out_e}")

    r_m = actualizar_google_doc(DRIVE_MATRIZ, out_m)
    print(f"[DRIVE Matriz] {r_m['web_view_link']}")
    r_e = actualizar_google_doc(DRIVE_ETAPA5, out_e)
    print(f"[DRIVE Etapa5] {r_e['web_view_link']}")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUT_DIR / "drive_update.json").write_text(
        json.dumps({"matriz": r_m, "etapa5": r_e}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
