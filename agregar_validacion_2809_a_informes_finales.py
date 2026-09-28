"""
Agrega la validación con fotos del 28-09-2026 a los informes FINAL:
  - Informe_Validacion_Matriz_ESVAL_Fundo_Zapallar_FINAL
  - Informe_Validacion_Etapa5_Zapallar_20260924_1136

y actualiza los Google Docs correspondientes en Drive.
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

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

WES_BLUE = RGBColor(31, 78, 121)


def _fmt(x: float, dec: int = 2) -> str:
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _shade(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
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


def _run(p, text: str, *, bold: bool = False, size: int = 11, color: RGBColor | None = None):
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = "Calibri"
    if color:
        r.font.color.rgb = color
    return r


def _heading(doc: Document, text: str, *, size: int = 13) -> None:
    p = doc.add_paragraph()
    _run(p, text, bold=True, size=size, color=WES_BLUE)


def _para(doc: Document, text: str, *, size: int = 10) -> None:
    p = doc.add_paragraph()
    _run(p, text, size=size)


def _table(doc: Document, headers: list[str], rows: list[list[str]], highlight_rows: set[int] | None = None) -> None:
    highlight_rows = highlight_rows or set()
    _fix_section_margins(doc)
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    try:
        t.style = "Table Grid"
    except KeyError:
        pass
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_borders(t)
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.paragraphs[0].clear()
        _run(cell.paragraphs[0], h, bold=True, size=9, color=RGBColor(255, 255, 255))
        _shade(cell, "1F4788")
    for r_i, vals in enumerate(rows):
        fill = "C6E0B4" if r_i in highlight_rows else "FFFFFF"
        for c_i, val in enumerate(vals):
            cell = t.rows[r_i + 1].cells[c_i]
            cell.paragraphs[0].clear()
            _run(cell.paragraphs[0], val, size=9)
            _shade(cell, fill)
    doc.add_paragraph()


def _fix_section_margins(doc: Document) -> None:
    """Algunos exports de Google Docs dejan márgenes float que rompen python-docx."""
    for section in doc.sections:
        sect_pr = section._sectPr
        pg_mar = sect_pr.find(qn("w:pgMar"))
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


def _fotos(doc: Document, izq: Path, cap_izq: str, der: Path, cap_der: str, *, ancho: float = 3.6) -> None:
    _fix_section_margins(doc)
    t = doc.add_table(rows=2, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for col, (path, cap) in enumerate(((izq, cap_izq), (der, cap_der))):
        cell = t.rows[0].cells[col]
        _shade(cell, "FAFBFC")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if path.is_file():
            p.add_run().add_picture(str(path), width=Inches(ancho))
        else:
            _run(p, "(foto no disponible)", size=9)
        cell_c = t.rows[1].cells[col]
        _shade(cell_c, "D6E3F0")
        pc = cell_c.paragraphs[0]
        pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(pc, cap, bold=True, size=9, color=WES_BLUE)
    doc.add_paragraph()


def _insert_before_conclusion(doc: Document, build_section) -> None:
    """Quita pie + conclusión, agrega sección nueva, reescribe conclusión y pie."""
    # Capturar textos de conclusión y pie
    conclusion_idx = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith("5. Conclusión"):
            conclusion_idx = i
            break
    if conclusion_idx is None:
        raise RuntimeError("No se encontró '5. Conclusión'")

    # Guardar conclusión body (párrafos entre 5. y pie WES)
    concl_body = []
    pie = ""
    for p in doc.paragraphs[conclusion_idx + 1 :]:
        t = p.text.strip()
        if t.startswith("WES ·"):
            pie = t
            break
        if t:
            concl_body.append(t)

    # Eliminar desde conclusión hasta el final (python-docx: clear text of those paras)
    # Mejor: construir doc nuevo copiando hasta antes de conclusión
    # Approach: mark by clearing conclusion onward, then append
    to_clear = []
    clearing = False
    for p in doc.paragraphs:
        if p.text.strip().startswith("5. Conclusión"):
            clearing = True
        if clearing:
            to_clear.append(p)
    for p in to_clear:
        p.clear()

    # Quitar párrafos vacíos finales dejando uno
    build_section(doc)

    _heading(doc, "5. Conclusión", size=13)
    for t in concl_body:
        _para(doc, t, size=10)
    stamp = datetime.now().strftime("%d-%m-%Y %H:%M")
    _para(
        doc,
        (pie.rsplit("·", 1)[0] + f"· {stamp}") if pie else f"WES · Informe final · Fundo Zapallar · {stamp}",
        size=9,
    )


def _update_header_validaciones(doc: Document, extra: str) -> None:
    for p in doc.paragraphs:
        t = p.text.strip()
        if t.startswith("Validaciones:"):
            # reemplazar run
            p.clear()
            _run(p, f"{t} | {extra}", size=10, color=RGBColor(89, 89, 89))
            return


def agregar_matriz(data: dict) -> Path:
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
        _para(
            d,
            "Validación de continuidad con lectura fotográfica del Itron del 28-09-2026 09:05 "
            "y la lectura de cierre de la validación previa (23-09 ~17:05). "
            "App WES: horas 17–23 del 23-09 + días 24–27 completos + horas 00–08 del 28-09 "
            "(misma metodología TIME etiqueta). "
            f"Lecturas Itron: {_fmt(m['itron_ini'], 1)} → {_fmt(m['itron_fin'], 1)} m³ "
            f"(Δ {_fmt(m['itron_delta'])} m³). Escala odómetro: 6 ruedas negras + 1 roja.",
            size=10,
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
                    "Δ Estanque Inferior (misma ventana)",
                    f"{_fmt(m['inferior_m3'])} m³",
                    f"{m['inferior_sobre_app_pct']:.0f}% de Matriz App",
                ],
            ],
            highlight_rows={4, 5},
        )
        chart = VAL_DIR / "chart_validacion_matriz_2809.png"
        if chart.is_file():
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(str(chart), width=Inches(6.8))
        _para(
            d,
            f"Error Itron vs app: {m['error_pct']:.1f}% "
            f"(1 − {_fmt(m['itron_delta'])}/{_fmt(m['app_m3'])}). {m['estado']}. "
            f"En la misma ventana el Estanque Inferior registra {_fmt(m['inferior_m3'])} m³ "
            f"({m['inferior_sobre_app_pct']:.0f}% de Matriz App): conversan (hay tubería en el camino).",
            size=10,
        )

    _insert_before_conclusion(doc, section)

    # Ampliar conclusión
    for p in doc.paragraphs:
        if "El desvío inicial no se debió" in p.text:
            extra = (
                f" Continuidad 23–28/09: Itron {_fmt(m['itron_delta'])} m³ vs App "
                f"{_fmt(m['app_m3'])} m³ (error {m['error_pct']:.1f}%, {m['estado']})."
            )
            _run(p, extra, size=10)
            break

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "Informe_Validacion_Matriz_ESVAL_Fundo_Zapallar_FINAL.docx"
    doc.save(str(out))
    return out


def agregar_etapa5(data: dict) -> Path:
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
        _para(
            d,
            "Validación de continuidad con lectura fotográfica Sensus del 28-09-2026 08:57 "
            "y la lectura de cierre previa (23-09 16:54 = 5.177 m³). "
            "App WES: misma ventana horaria que Matriz (h17 del 23-09 → h08 del 28-09). "
            f"Lecturas: {_fmt(e['sensus_ini'], 0)} → {_fmt(e['sensus_fin'], 0)} m³ "
            f"(Δ {_fmt(e['sensus_delta'], 0)} m³; odómetro 005225).",
            size=10,
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
            highlight_rows={4},
        )
        chart = VAL_DIR / "chart_validacion_etapa5_2809.png"
        if chart.is_file():
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(str(chart), width=Inches(5.8))
        _para(
            d,
            f"Error Sensus vs app: {e['error_pct']:.1f}% "
            f"(1 − {_fmt(e['sensus_delta'], 0)}/{_fmt(e['app_m3'])}). {e['estado']}.",
            size=10,
        )

    _insert_before_conclusion(doc, section)

    for p in doc.paragraphs:
        if "Se confirma falla de memoria" in p.text:
            extra = (
                f" Continuidad 23–28/09: Sensus {_fmt(e['sensus_delta'], 0)} m³ vs App "
                f"{_fmt(e['app_m3'])} m³ (error {e['error_pct']:.1f}%, {e['estado']})."
            )
            _run(p, extra, size=10)
            break

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "Informe_Validacion_Etapa5_Zapallar_FINAL.docx"
    doc.save(str(out))
    return out


def actualizar_google_doc(file_id: str, docx_path: Path) -> dict:
    """Reemplaza el contenido de un Google Doc nativo con un .docx (Drive convierte)."""
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
    out_m = agregar_matriz(data)
    out_e = agregar_etapa5(data)
    print(f"[OK] {out_m}")
    print(f"[OK] {out_e}")

    r_m = actualizar_google_doc(DRIVE_MATRIZ, out_m)
    print(f"[DRIVE Matriz] {r_m['web_view_link']}")
    r_e = actualizar_google_doc(DRIVE_ETAPA5, out_e)
    print(f"[DRIVE Etapa5] {r_e['web_view_link']}")

    (OUT_DIR / "drive_update.json").write_text(
        json.dumps({"matriz": r_m, "etapa5": r_e}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
