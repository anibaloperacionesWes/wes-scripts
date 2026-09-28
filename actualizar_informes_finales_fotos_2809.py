"""
Actualiza los informes FINAL según pedido:
  - Etapa 5: SOLO agrega la foto de lectura del 28-09 (sin nuevo cálculo).
  - Matriz: actualiza con lecturas de hoy (sección 4.3 continuidad + fotos).

Fuente: docx originales (_drive_edit) = mismo contenido que los PDF subidos.
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
from PIL import Image, ImageOps

from wes_google_drive import obtener_servicio_drive

ROOT = Path("reports/Fundo_Zapallar/Informes_Tecnicos")
SRC = ROOT / "_drive_edit"
VAL = ROOT / "validacion_terreno_2809_20260928_1726"
OUT = ROOT / f"_final_cliente_foto_2809_{datetime.now().strftime('%Y%m%d_%H%M')}"

DRIVE_MATRIZ = "16IlboR89inKpImJSL4j4vNt9Lt0H4MHsf112l5Jjkh4"
DRIVE_ETAPA5 = "1bJqXMrPc8zSuhfx70YGgFmgxbrMTISipP6rn4Gw4BRU"

FOTO_ETAPA5_HOY = VAL / "foto_etapa5_sensus_2809_0857.jpg"
FOTO_MATRIZ_PREV = VAL / "foto_matriz_itron_2309_1705.jpg"
FOTO_MATRIZ_HOY = VAL / "foto_matriz_itron_2809_0905.jpg"
# fallback assets
ASSET_ETAPA5 = Path("/home/ubuntu/.cursor/projects/workspace/assets/bf785120-aea0-404f-9fa4-0a4df3dccd62.png")
ASSET_MATRIZ = Path("/home/ubuntu/.cursor/projects/workspace/assets/2c90844e-579b-4a89-8e18-0e424d5e64cf.png")

DATA = json.loads((VAL / "validacion_terreno_2809.json").read_text(encoding="utf-8"))

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


def _fmt(x: float, dec: int = 2) -> str:
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _fix_margins(doc: Document) -> None:
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


def _shade(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    old = tc_pr.find(qn("w:shd"))
    if old is not None:
        tc_pr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    tc_pr.append(shd)


def _borders(table, color: str = "D0D5DD") -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl.insert(0, tbl_pr)
    old = tbl_pr.find(qn("w:tblBorders"))
    if old is not None:
        tbl_pr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tbl_pr.append(borders)


def _run(p, text, *, bold=False, size=11, color=None):
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = "Calibri"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    if color:
        r.font.color.rgb = color
    return r


def _ensure_foto(path: Path, asset: Path) -> Path:
    if path.is_file():
        return path
    OUT.mkdir(parents=True, exist_ok=True)
    dest = OUT / path.name
    img = ImageOps.exif_transpose(Image.open(asset)).convert("RGB")
    img.save(dest, quality=92)
    return dest


def _prep_foto_reloj(src: Path, dest: Path, *, max_side: int = 1200) -> Path:
    img = ImageOps.exif_transpose(Image.open(src)).convert("RGB")
    w, h = img.size
    # recorte centrado de la relojería (mitad superior-central)
    cx, cy = w // 2, int(h * 0.42)
    half_w, half_h = int(w * 0.38), int(h * 0.32)
    box = (max(0, cx - half_w), max(0, cy - half_h), min(w, cx + half_w), min(h, cy + half_h))
    crop = img.crop(box)
    cw, ch = crop.size
    scale = min(1.0, max_side / max(cw, ch))
    if scale < 1:
        crop = crop.resize((int(cw * scale), int(ch * scale)), Image.LANCZOS)
    dest.parent.mkdir(parents=True, exist_ok=True)
    crop.save(dest, quality=92)
    return dest


def _clear_from_conclusion(doc: Document) -> tuple[list[str], str]:
    idx = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith("5. Conclusión"):
            idx = i
            break
    if idx is None:
        raise RuntimeError("No hay sección 5. Conclusión")
    body, pie = [], ""
    for p in doc.paragraphs[idx + 1 :]:
        t = p.text.strip()
        if t.startswith("WES ·"):
            pie = t
            break
        if t:
            body.append(t)
    clearing = False
    for p in doc.paragraphs:
        if p.text.strip().startswith("5. Conclusión"):
            clearing = True
        if clearing:
            p.clear()
    return body, pie


def _write_conclusion(doc: Document, body: list[str], pie: str, extra: str | None = None) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    _run(p, "5. Conclusión", bold=True, size=14, color=COLOR_TITULO)
    text = " ".join(body)
    if extra and extra not in text:
        text = text.rstrip() + " " + extra
    _p = doc.add_paragraph()
    _run(_p, text, size=10.5, color=COLOR_TEXTO)
    stamp = datetime.now().strftime("%d-%m-%Y %H:%M")
    foot = doc.add_paragraph()
    foot.paragraph_format.space_before = Pt(8)
    base = pie.rsplit("·", 1)[0] if pie else "WES · Informe final · Fundo Zapallar · "
    _run(foot, f"{base}· {stamp}", size=9, color=COLOR_META)


def _foto_unica(doc: Document, path: Path, caption: str, *, ancho: float = 3.2) -> None:
    """Una foto centrada con caption (estilo FINAL)."""
    _fix_margins(doc)
    t = doc.add_table(rows=2, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(t, HEX_NAVY_LIGHT)
    c0 = t.rows[0].cells[0]
    _shade(c0, HEX_FOTO)
    p = c0.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(path), width=Inches(ancho))
    c1 = t.rows[1].cells[0]
    _shade(c1, HEX_NAVY_LIGHT)
    pc = c1.paragraphs[0]
    pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(pc, caption, bold=True, size=9, color=COLOR_TITULO)
    doc.add_paragraph()


def _fotos_lado(doc: Document, a: Path, ca: str, b: Path, cb: str, *, ancho: float = 3.4) -> None:
    _fix_margins(doc)
    t = doc.add_table(rows=2, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(t, HEX_NAVY_LIGHT)
    for col, (path, cap) in enumerate(((a, ca), (b, cb))):
        cell = t.rows[0].cells[col]
        _shade(cell, HEX_FOTO)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(path), width=Inches(ancho))
        cell_c = t.rows[1].cells[col]
        _shade(cell_c, HEX_NAVY_LIGHT)
        pc = cell_c.paragraphs[0]
        pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(pc, cap, bold=True, size=9, color=COLOR_TITULO)
    doc.add_paragraph()


def _kpi(doc: Document, cards: list[tuple[str, str, str]]) -> None:
    _fix_margins(doc)
    t = doc.add_table(rows=2, cols=len(cards))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(t, HEX_NAVY)
    for j, (titulo, v1, v2) in enumerate(cards):
        c0 = t.rows[0].cells[j]
        c0.paragraphs[0].clear()
        c0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(c0.paragraphs[0], titulo, bold=True, size=9, color=COLOR_WHITE)
        _shade(c0, HEX_NAVY)
        c1 = t.rows[1].cells[j]
        c1.paragraphs[0].clear()
        c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(c1.paragraphs[0], v1 + "\n", bold=True, size=12, color=COLOR_KPI)
        _run(c1.paragraphs[0], v2, size=8, color=COLOR_META)
        _shade(c1, HEX_KPI)
    doc.add_paragraph()


def _tabla_val(doc: Document, titulo: str, rows: list[list[str]]) -> None:
    """Tabla estilo 4.2 FINAL: header navy + filas."""
    _fix_margins(doc)
    # título de bloque
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    _run(p, titulo, bold=True, size=10, color=COLOR_TITULO)

    headers = ["ANÁLISIS", "FECHA INICIAL", "LECTURA (m³)", "FECHA FINAL", "LECTURA (m³)", "CONSUMO m³", "HORAS"]
    t = doc.add_table(rows=1 + len(rows), cols=7)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(t, "D0D5DD")
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.paragraphs[0].clear()
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(cell.paragraphs[0], h, bold=True, size=8, color=COLOR_WHITE)
        _shade(cell, HEX_NAVY)
    for r_i, vals in enumerate(rows):
        fill = HEX_ALT if r_i % 2 else "FFFFFF"
        for c_i, val in enumerate(vals):
            cell = t.rows[r_i + 1].cells[c_i]
            cell.paragraphs[0].clear()
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run(cell.paragraphs[0], val, size=9, color=COLOR_TEXTO)
            _shade(cell, fill)
    doc.add_paragraph()


def _tabla_totales(doc: Document, items: list[tuple[str, str]], ok_last: bool = True) -> None:
    _fix_margins(doc)
    t = doc.add_table(rows=len(items), cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(t, HEX_NAVY)
    for i, (k, v) in enumerate(items):
        fill = HEX_KPI if (ok_last and i == len(items) - 1) else (HEX_ALT if i % 2 else "FFFFFF")
        c0, c1 = t.rows[i].cells[0], t.rows[i].cells[1]
        c0.paragraphs[0].clear()
        c1.paragraphs[0].clear()
        color = COLOR_KPI if (ok_last and i == len(items) - 1) else COLOR_TEXTO
        _run(c0.paragraphs[0], k, bold=True, size=10, color=color)
        _run(c1.paragraphs[0], v, bold=True, size=10, color=color)
        _shade(c0, fill)
        _shade(c1, fill)
    doc.add_paragraph()


# ─── Etapa 5: SOLO foto ───────────────────────────────────────────────

def actualizar_etapa5() -> Path:
    doc = Document(str(SRC / "Informe_Validacion_Etapa5_Zapallar_20260924_1136.docx"))
    _fix_margins(doc)
    foto_src = _ensure_foto(FOTO_ETAPA5_HOY, ASSET_ETAPA5)
    foto = _prep_foto_reloj(foto_src, OUT / "foto_etapa5_2809_reloj.jpg")

    body, pie = _clear_from_conclusion(doc)

    # Insertar solo la foto nueva antes de la conclusión (tras 4.2 existente)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    _run(p, "4.3 Lectura fotográfica adicional (28-09-2026)", bold=True, size=12, color=COLOR_TITULO)
    p2 = doc.add_paragraph()
    _run(
        p2,
        "Foto de relojería enviada el 28-09-2026 a las 08:57. "
        "Lectura mecánica Sensus: 5.225 m³ (odómetro 005225).",
        size=10.5,
        color=COLOR_TEXTO,
    )
    _foto_unica(doc, foto, "Sensus · 5.225 m³ · 28-09-2026 08:57", ancho=3.0)

    _write_conclusion(doc, body, pie, extra=None)

    OUT.mkdir(parents=True, exist_ok=True)
    out = OUT / "Informe_Validacion_Etapa5_Zapallar_20260924_1136.docx"
    doc.save(str(out))
    return out


# ─── Matriz: actualizar con lecturas de hoy ───────────────────────────

def actualizar_matriz() -> Path:
    doc = Document(str(SRC / "Informe_Validacion_Matriz_ESVAL_Fundo_Zapallar_FINAL.docx"))
    _fix_margins(doc)
    m = DATA["matriz"]

    # Header validaciones
    for p in doc.paragraphs:
        if p.text.strip().startswith("Validaciones:"):
            p.clear()
            _run(
                p,
                "Validaciones: 22-09-2026 11:50–15:00  |  22-09-2026 15:00 → 23-09-2026 17:00  |  "
                "23-09-2026 17:00 → 28-09-2026 09:05",
                size=10,
                color=COLOR_META,
            )
            break

    foto_prev = FOTO_MATRIZ_PREV if FOTO_MATRIZ_PREV.is_file() else _ensure_foto(FOTO_MATRIZ_PREV, ASSET_MATRIZ)
    foto_hoy = _ensure_foto(FOTO_MATRIZ_HOY, ASSET_MATRIZ)
    # para hoy usar foto completa (ya es close-up); prev ya está horizontalizada
    foto_hoy_r = _prep_foto_reloj(foto_hoy, OUT / "foto_matriz_2809_reloj.jpg")

    body, pie = _clear_from_conclusion(doc)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    _run(p, "4.3 Continuidad Itron vs app WES (23-09 17:00 → 28-09 09:05)", bold=True, size=12, color=COLOR_TITULO)

    _kpi(
        doc,
        [
            ("Δ Itron (terreno)", f"{_fmt(m['itron_delta'])} m³", f"{_fmt(m['itron_ini'], 1)} → {_fmt(m['itron_fin'], 1)}"),
            ("Δ Matriz App WES", f"{_fmt(m['app_m3'])} m³", "000027-01"),
            ("Error", f"{m['error_pct']:.1f} %", m["estado"]),
        ],
    )

    p2 = doc.add_paragraph()
    _run(
        p2,
        "Validación con lecturas fotográficas del medidor Itron y el consumo de la app WES. "
        "App: horas 17→23 del 23-09 + días 24–27 + horas 00→08 del 28-09 "
        "(continúa tras la validación 4.2).",
        size=10.5,
        color=COLOR_TEXTO,
    )

    _fotos_lado(
        doc,
        foto_prev,
        f"Itron · {_fmt(m['itron_ini'], 1)} m³ · 23-09-2026 17:05",
        foto_hoy_r,
        f"Itron · {_fmt(m['itron_fin'], 1)} m³ · 28-09-2026 09:05",
    )

    horas = DATA["ventana_app"]["horas"]
    _tabla_val(
        doc,
        "Validación Matriz ESVAL — medidor Itron (continuidad)",
        [
            [
                "Matriz ESVAL",
                "23-09-2026\n17:00",
                _fmt(m["itron_ini"], 1),
                "28-09-2026\n09:05",
                _fmt(m["itron_fin"], 1),
                _fmt(m["itron_delta"]),
                str(horas),
            ]
        ],
    )
    _tabla_totales(
        doc,
        [
            ("Total App WES", f"{_fmt(m['app_m3'])} m³"),
            ("Total Lectura Itron", f"{_fmt(m['itron_delta'])} m³"),
            ("% Error", f"{m['error_pct']:.1f} %"),
        ],
    )
    p3 = doc.add_paragraph()
    _run(
        p3,
        f"Error entre lectura Itron y app WES: {m['error_pct']:.1f}% "
        f"(1 − {_fmt(m['itron_delta'])}/{_fmt(m['app_m3'])}). {m['estado']}.",
        size=10.5,
        color=COLOR_TEXTO,
    )

    extra = (
        f"Continuidad 23–28/09: Itron {_fmt(m['itron_delta'])} m³ vs app "
        f"{_fmt(m['app_m3'])} m³ (error {m['error_pct']:.1f}%, {m['estado']})."
    )
    _write_conclusion(doc, body, pie, extra=extra)

    OUT.mkdir(parents=True, exist_ok=True)
    out = OUT / "Informe_Validacion_Matriz_ESVAL_Fundo_Zapallar_FINAL.docx"
    doc.save(str(out))
    return out


def _docx_to_pdf(docx_path: Path) -> Path | None:
    """Intenta LibreOffice; si no, omite."""
    import shutil
    import subprocess

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(docx_path.parent), str(docx_path)],
        check=True,
        capture_output=True,
    )
    pdf = docx_path.with_suffix(".pdf")
    return pdf if pdf.is_file() else None


def actualizar_drive(file_id: str, docx_path: Path) -> dict:
    svc = obtener_servicio_drive()
    media = MediaFileUpload(
        str(docx_path),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        resumable=True,
    )
    meta = svc.files().update(
        fileId=file_id, media_body=media, fields="id,name,webViewLink,modifiedTime"
    ).execute()
    return {
        "id": meta["id"],
        "name": meta["name"],
        "web_view_link": meta.get("webViewLink")
        or f"https://docs.google.com/document/d/{meta['id']}/edit",
        "modifiedTime": meta.get("modifiedTime"),
    }


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    out_e5 = actualizar_etapa5()
    out_m = actualizar_matriz()
    print(f"[OK] {out_e5}")
    print(f"[OK] {out_m}")

    r_e = actualizar_drive(DRIVE_ETAPA5, out_e5)
    r_m = actualizar_drive(DRIVE_MATRIZ, out_m)
    print(f"[DRIVE Etapa5] {r_e['web_view_link']}")
    print(f"[DRIVE Matriz] {r_m['web_view_link']}")

    # PDF locales si hay LibreOffice
    for d in (out_e5, out_m):
        pdf = _docx_to_pdf(d)
        if pdf:
            print(f"[PDF] {pdf}")

    (OUT / "drive_update.json").write_text(
        json.dumps({"etapa5": r_e, "matriz": r_m}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
