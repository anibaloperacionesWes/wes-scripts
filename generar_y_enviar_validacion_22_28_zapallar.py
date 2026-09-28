"""
Informes FINAL Zapallar — UNA sola validación 22→28/09.
Colores del PDF original (#1F4788 / #E8F5E9 / #F5F8FB).
PDF vía Drive + correo a Juan y Aníbal.
"""

from __future__ import annotations

import csv
import io
import json
import os
import smtplib
from datetime import date, datetime, timedelta, timezone
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from PIL import Image, ImageOps

from wes_google_drive import obtener_servicio_drive

CHILE_TZ = ZoneInfo("America/Santiago")
# Hueco API del 22/09 tras cambio de placa (m³/h) — mismo criterio del informe Etapa 5 original
PLACA_HUECO_22: dict[int, float] = {
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

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "reports/Fundo_Zapallar/Informes_Tecnicos/_drive_edit"
EVID = ROOT / "reports/Fundo_Zapallar/Informes_Tecnicos/_evidencias"
FOTOS_E5 = ROOT / "reports/Fundo_Zapallar/Informes_Tecnicos/fotos_etapa5"
ASSET_M = Path("/home/ubuntu/.cursor/projects/workspace/assets/2c90844e-579b-4a89-8e18-0e424d5e64cf.png")
ASSET_E = Path("/home/ubuntu/.cursor/projects/workspace/assets/bf785120-aea0-404f-9fa4-0a4df3dccd62.png")
OUT = ROOT / "reports/Fundo_Zapallar/Informes_Tecnicos" / "validacion_22_28_out"

DRIVE_MATRIZ = "16IlboR89inKpImJSL4j4vNt9Lt0H4MHsf112l5Jjkh4"
DRIVE_ETAPA5 = "1bJqXMrPc8zSuhfx70YGgFmgxbrMTISipP6rn4Gw4BRU"
BASE = "http://104.248.53.141:7003/wes/api/acl-node/v1"

ITRON_INI, ITRON_FIN = 383978.01, 384736.5
ITRON_INI_DT, ITRON_FIN_DT = "22-09-2026 15:00", "28-09-2026 09:05"
SENSUS_INI, SENSUS_FIN = 5144.0, 5225.0
SENSUS_INI_DT, SENSUS_FIN_DT = "22-09-2026 14:30", "28-09-2026 08:57"

SMTP_USUARIO = "agente.ia@wes.cl"
TO = ["juanlopez@wes.cl", "anibal.aoperaciones@wes.cl"]

# Paleta exacta del PDF original que pasó el usuario
NAVY = "1F4788"          # header azul
OK = "E8F5E9"            # KPI verde claro
ALT = "F5F8FB"           # fila alterna
FOTO_BG = "FAFBFC"
COLOR_TITULO = RGBColor(0x1F, 0x47, 0x88)
COLOR_META = RGBColor(0x64, 0x6E, 0x78)
COLOR_TEXTO = RGBColor(0x28, 0x28, 0x28)
COLOR_KPI = RGBColor(0x1B, 0x5E, 0x20)
COLOR_WHITE = RGBColor(255, 255, 255)


def _fmt(x: float, dec: int = 2) -> str:
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _smtp_password() -> str:
    p = (os.environ.get("WES_GMAIL_APP_PASSWORD") or os.environ.get("WES_SMTP_PASSWORD") or "").strip()
    if p:
        return p.replace(" ", "")
    f = ROOT / "gmail_oauth" / "app_password.txt"
    if f.is_file():
        for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
            if line.strip() and not line.strip().startswith("#"):
                return line.strip().replace(" ", "")
    return "vxbynfpoehbweelj"


def _hours(node_id: str, dia: date) -> dict[int, float]:
    r = requests.get(
        f"{BASE}/nodes/{node_id}/dates.measures.csv",
        params={"start": dia.strftime("%d%m%Y"), "end": dia.strftime("%d%m%Y")},
        timeout=60,
    )
    r.raise_for_status()
    acc: dict[int, float] = {}
    for row in csv.DictReader(io.StringIO(r.text)):
        dt = datetime.fromisoformat(row["TIME"].strip().replace("Z", "+00:00"))
        acc[dt.hour] = acc.get(dt.hour, 0.0) + float(row["VALUE"].strip())
    return acc


def _sum(node_id: str, slots: list[tuple[date, int]]) -> float:
    by = {d: _hours(node_id, d) for d in sorted({x for x, _ in slots})}
    return round(sum(by[d].get(h, 0.0) for d, h in slots), 2)


def _slots_m():
    out = [(date(2026, 9, 22), h) for h in range(16, 24)]
    for d in range(23, 28):
        out += [(date(2026, 9, d), h) for h in range(24)]
    out += [(date(2026, 9, 28), h) for h in range(0, 9)]
    return out


def _slots_e():
    """Horas Chile 22-09 14:00 → 28-09 09:00 (excl.)."""
    out = [(date(2026, 9, 22), h) for h in range(14, 24)]
    for d in range(23, 28):
        out += [(date(2026, 9, d), h) for h in range(24)]
    out += [(date(2026, 9, 28), h) for h in range(0, 9)]
    return out


def _sum_etapa5_chile() -> float:
    """App Etapa 5: TIME en UTC → hora Chile + hueco placa 22/09.

    Igual que generar_informe_cambio_memoria_etapa5_zapallar.py.
    Sin esto el 22/09 queda truncado en la API y el total app sale bajo.
    """
    api: dict[datetime, float] = {}
    for d in range(22, 29):
        dia = date(2026, 9, d)
        r = requests.get(
            f"{BASE}/nodes/000027-03/dates.measures.csv",
            params={"start": dia.strftime("%d%m%Y"), "end": dia.strftime("%d%m%Y")},
            timeout=60,
        )
        r.raise_for_status()
        for row in csv.DictReader(io.StringIO(r.text)):
            dt = datetime.fromisoformat(row["TIME"].strip().replace("Z", "+00:00"))
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            api[dt.astimezone(CHILE_TZ)] = float(row["VALUE"].strip())

    start = datetime(2026, 9, 22, 14, 0, tzinfo=CHILE_TZ)
    end = datetime(2026, 9, 28, 9, 0, tzinfo=CHILE_TZ)
    total = 0.0
    cur = start
    while cur < end:
        if cur.date() == date(2026, 9, 22) and cur.hour in PLACA_HUECO_22:
            total += PLACA_HUECO_22[cur.hour]
        elif cur in api:
            total += api[cur]
        else:
            for k, v in api.items():
                if k.date() == cur.date() and k.hour == cur.hour:
                    total += v
                    break
        cur += timedelta(hours=1)
    return round(total, 2)


def _err(a, b):
    return abs(1 - min(a, b) / max(a, b)) * 100 if a > 0 and b > 0 else 0.0


def _estado(pct):
    if pct <= 5:
        return "Aceptable"
    if pct <= 15:
        return "Aceptable (diferencia moderada)"
    return "Revisar — diferencia relevante"


def _fix_margins(doc):
    for section in doc.sections:
        pg = section._sectPr.find(qn("w:pgMar"))
        if pg is None:
            continue
        for attr, val in list(pg.attrib.items()):
            try:
                int(val)
            except ValueError:
                try:
                    pg.set(attr, str(int(round(float(val)))))
                except ValueError:
                    pg.set(attr, "720")


def _shade(cell, fill: str):
    tc = cell._tc.get_or_add_tcPr()
    old = tc.find(qn("w:shd"))
    if old is not None:
        tc.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    tc.append(shd)


def _borders(table, color="D0D5DD"):
    tbl = table._tbl
    pr = tbl.tblPr
    if pr is None:
        pr = OxmlElement("w:tblPr")
        tbl.insert(0, pr)
    old = pr.find(qn("w:tblBorders"))
    if old is not None:
        pr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    pr.append(borders)


def _run(p, text, *, bold=False, size=11, color=None):
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = "Calibri"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    if color:
        r.font.color.rgb = color
    return r


def _clear_cell(cell):
    """Borra todo el contenido de la celda (todos los párrafos, no solo el primero)."""
    for p in cell.paragraphs:
        p.clear()
    # Eliminar párrafos extra dejando solo el primero
    tc = cell._tc
    ps = tc.xpath("./w:p")
    for el in ps[1:]:
        tc.remove(el)


def _set_cell(cell, text, *, bold=False, size=9, color=None, fill=None, center=False):
    _clear_cell(cell)
    if center:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(cell.paragraphs[0], text, bold=bold, size=size, color=color)
    if fill:
        _shade(cell, fill)


def _recolor_all_tables(doc: Document):
    """Restaura colores del PDF original en tablas existentes (Google export las pierde)."""
    for ti, table in enumerate(doc.tables):
        _borders(table, "D0D5DD")
        n = len(table.rows)
        for ri, row in enumerate(table.rows):
            for cell in row.cells:
                if ri == 0:
                    _shade(cell, NAVY)
                    for p in cell.paragraphs:
                        for r in p.runs:
                            r.font.color.rgb = COLOR_WHITE
                            r.bold = True
                else:
                    # KPI strip (tabla 0): fila valores en verde
                    if ti == 0:
                        _shade(cell, OK)
                        for p in cell.paragraphs:
                            runs = p.runs
                            if runs:
                                runs[0].font.color.rgb = COLOR_KPI
                                runs[0].bold = True
                                for r in runs[1:]:
                                    r.font.color.rgb = COLOR_META
                    else:
                        _shade(cell, ALT if ri % 2 == 0 else "FFFFFF")


def _prep(src: Path, dest: Path, *, rotate_cw90=False, crop=False) -> Path:
    img = ImageOps.exif_transpose(Image.open(src)).convert("RGB")
    if rotate_cw90:
        img = img.transpose(Image.ROTATE_270)
    if crop:
        w, h = img.size
        cx, cy = w // 2, int(h * 0.42)
        hw, hh = int(w * 0.36), int(h * 0.30)
        img = img.crop((max(0, cx - hw), max(0, cy - hh), min(w, cx + hw), min(h, cy + hh)))
    w, h = img.size
    s = min(1.0, 1100 / max(w, h))
    if s < 1:
        img = img.resize((int(w * s), int(h * s)), Image.LANCZOS)
    dest.parent.mkdir(parents=True, exist_ok=True)
    img.save(dest, quality=92)
    return dest


def _clear_from(doc, start_prefix: str):
    """Elimina del body XML párrafos Y tablas desde start_prefix hasta el final.

    Antes solo se borraba el texto de los párrafos y quedaban las tablas de las
    dos validaciones viejas (4.1 + 4.2). Ahora se remueven de verdad.
    """
    body = doc.element.body
    start = None
    for child in list(body):
        tag = child.tag.split("}")[-1]
        if tag != "p":
            continue
        text = "".join(n.text or "" for n in child.iter(qn("w:t"))).strip()
        if text.startswith(start_prefix):
            start = child
            break
    if start is None:
        return
    removing = False
    for child in list(body):
        if child is start:
            removing = True
        if not removing:
            continue
        tag = child.tag.split("}")[-1]
        if tag == "sectPr":
            continue
        body.remove(child)


def _kpi3(doc, cards: list[tuple[str, str, str]]):
    """Franja KPI 3 columnas — estilo PDF original."""
    _fix_margins(doc)
    t = doc.add_table(rows=2, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(t, NAVY)
    for j, (h, v1, v2) in enumerate(cards):
        _set_cell(t.rows[0].cells[j], h, bold=True, size=9, color=COLOR_WHITE, fill=NAVY, center=True)
        c = t.rows[1].cells[j]
        _clear_cell(c)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(c.paragraphs[0], v1 + "\n", bold=True, size=12, color=COLOR_KPI)
        _run(c.paragraphs[0], v2, size=8, color=COLOR_META)
        _shade(c, OK)
    doc.add_paragraph()


def _fotos(doc, a, ca, b, cb, ancho=3.35):
    _fix_margins(doc)
    t = doc.add_table(rows=2, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(t, "D0D5DD")
    for col, (path, cap) in enumerate(((a, ca), (b, cb))):
        c0 = t.rows[0].cells[col]
        _shade(c0, FOTO_BG)
        p = c0.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(path), width=Inches(ancho))
        c1 = t.rows[1].cells[col]
        _set_cell(c1, cap, bold=True, size=9, color=COLOR_TITULO, fill="D6E3F0", center=True)
    doc.add_paragraph()


def _tabla_val(doc, titulo, row7):
    _fix_margins(doc)
    p = doc.add_paragraph()
    _run(p, titulo, bold=True, size=10, color=COLOR_TITULO)
    headers = ["ANÁLISIS", "FECHA INICIAL", "LECTURA (m³)", "FECHA FINAL", "LECTURA (m³)", "CONSUMO m³", "HORAS"]
    t = doc.add_table(rows=2, cols=7)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(t, "D0D5DD")
    for i, h in enumerate(headers):
        _set_cell(t.rows[0].cells[i], h, bold=True, size=8, color=COLOR_WHITE, fill=NAVY, center=True)
    for i, v in enumerate(row7):
        _set_cell(t.rows[1].cells[i], v, size=9, color=COLOR_TEXTO, fill="FFFFFF", center=True)
    doc.add_paragraph()


def _totales(doc, items):
    _fix_margins(doc)
    t = doc.add_table(rows=len(items), cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(t, NAVY)
    for i, (k, v) in enumerate(items):
        fill = OK if i == len(items) - 1 else (ALT if i % 2 else "FFFFFF")
        color = COLOR_KPI if i == len(items) - 1 else COLOR_TEXTO
        _set_cell(t.rows[i].cells[0], k, bold=True, size=10, color=color, fill=fill)
        _set_cell(t.rows[i].cells[1], v, bold=True, size=10, color=color, fill=fill)
    doc.add_paragraph()


def _rewrite_cover_kpi(doc: Document, cards: list[tuple[str, str, str]]):
    """Reescribe la franja KPI de portada (tabla 0) con UNA validación + config."""
    if not doc.tables:
        return
    tbl = doc.tables[0]
    # Asegurar 2 filas x 3 cols
    while len(tbl.rows) < 2:
        tbl.add_row()
    _borders(tbl, NAVY)
    for j in range(min(3, len(tbl.columns))):
        h, v1, v2 = cards[j]
        _set_cell(tbl.rows[0].cells[j], h, bold=True, size=9, color=COLOR_WHITE, fill=NAVY, center=True)
        c = tbl.rows[1].cells[j]
        _clear_cell(c)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(c.paragraphs[0], v1 + "\n", bold=True, size=12, color=COLOR_KPI)
        _run(c.paragraphs[0], v2, size=8, color=COLOR_META)
        _shade(c, OK)


def build_matriz(app_m, err, estado, horas) -> Path:
    doc = Document(str(SRC / "Informe_Validacion_Matriz_ESVAL_Fundo_Zapallar_FINAL.docx"))
    _fix_margins(doc)
    delta = round(ITRON_FIN - ITRON_INI, 2)

    for p in doc.paragraphs:
        t = p.text.strip()
        if t.startswith("Validaciones:"):
            p.clear()
            _run(p, "Validación: 22-09-2026 15:00 → 28-09-2026 09:05", size=10, color=COLOR_META)
        if t.startswith("2. Resumen"):
            # next para will be updated below
            pass

    # Actualizar resumen (párrafo después de "2. Resumen")
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith("2. Resumen") and i + 1 < len(doc.paragraphs):
            nxt = doc.paragraphs[i + 1]
            nxt.clear()
            _run(
                nxt,
                "La calidad de señal (DN/UP y Q) estaba dentro del umbral del fabricante. "
                "Tras limpieza, silicona, corrección de diámetro/mortero y factor de escala, "
                f"la validación única 22–28/09 da error {err:.1f}% (Itron {_fmt(delta)} m³ vs app "
                f"{_fmt(app_m)} m³), {estado.lower()}.",
                size=10.5,
                color=COLOR_TEXTO,
            )
            break

    # Portada: UNA sola validación (no dos)
    _rewrite_cover_kpi(
        doc,
        [
            ("Itron vs app WES", f"Error {err:.1f} %", f"{_fmt(delta)} / {_fmt(app_m)} m³"),
            ("Lecturas Itron", f"{_fmt(ITRON_INI, 0)} → {_fmt(ITRON_FIN, 0)}", f"Δ {_fmt(delta)} m³ · 22→28/09"),
            ("Config. final", "Ø 127 mm · mortero 3 mm", "M45 tope 1,5"),
        ],
    )
    _recolor_all_tables(doc)

    _clear_from(doc, "4. Cálculo")

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    _run(p, "4. Cálculo de validación", bold=True, size=14, color=COLOR_TITULO)
    p2 = doc.add_paragraph()
    _run(p2, "4.1 Itron vs app WES (22-09 15:00 → 28-09 09:05)", bold=True, size=12, color=COLOR_TITULO)

    _kpi3(
        doc,
        [
            ("Itron vs app WES", f"Error {err:.1f} %", f"{_fmt(delta)} / {_fmt(app_m)} m³"),
            ("Lecturas Itron", f"{_fmt(ITRON_INI, 0)} → {_fmt(ITRON_FIN, 0)}", f"Δ {_fmt(delta)} m³"),
            ("Ventana", "22 → 28/09/2026", f"{horas} horas app"),
        ],
    )

    p3 = doc.add_paragraph()
    _run(
        p3,
        "Una sola validación con lecturas fotográficas del Itron y el consumo de la app WES "
        "(horas 16→23 del 22-09 + días 23–27 + horas 00→08 del 28-09).",
        size=10.5,
        color=COLOR_TEXTO,
    )

    f_ini = _prep(EVID / "itron_1500.jpg", OUT / "foto_matriz_ini.jpg", rotate_cw90=True)
    f_fin = _prep(ASSET_M, OUT / "foto_matriz_fin.jpg", crop=True)
    _fotos(
        doc,
        f_ini,
        f"Itron · {_fmt(ITRON_INI, 1)} m³ · {ITRON_INI_DT}",
        f_fin,
        f"Itron · {_fmt(ITRON_FIN, 1)} m³ · {ITRON_FIN_DT}",
    )

    _tabla_val(
        doc,
        "Validación Matriz ESVAL — medidor Itron",
        [
            "Matriz ESVAL",
            "22-09-2026\n15:00",
            _fmt(ITRON_INI, 1),
            "28-09-2026\n09:05",
            _fmt(ITRON_FIN, 1),
            _fmt(delta),
            str(horas),
        ],
    )
    _totales(
        doc,
        [
            ("Total App WES", f"{_fmt(app_m)} m³"),
            ("Total Lectura Itron", f"{_fmt(delta)} m³"),
            ("% Error", f"{err:.1f} %"),
        ],
    )
    p4 = doc.add_paragraph()
    _run(
        p4,
        f"Error entre lectura Itron y app WES: {err:.1f}% "
        f"(1 − {_fmt(delta)}/{_fmt(app_m)}). {estado}.",
        size=10.5,
        color=COLOR_TEXTO,
    )

    p5 = doc.add_paragraph()
    p5.paragraph_format.space_before = Pt(10)
    _run(p5, "5. Conclusión", bold=True, size=14, color=COLOR_TITULO)
    p6 = doc.add_paragraph()
    _run(
        p6,
        "El desvío inicial no se debió a calidad de señal ni a cableado. "
        "La corrección de diámetro, mortero y factor de escala alineó ambos medidores. "
        f"Validación única 22–28/09: Itron {_fmt(delta)} m³ vs app {_fmt(app_m)} m³ "
        f"(error {err:.1f}%, {estado}).",
        size=10.5,
        color=COLOR_TEXTO,
    )
    foot = doc.add_paragraph()
    foot.paragraph_format.space_before = Pt(8)
    _run(
        foot,
        f"WES · Informe final · Fundo Zapallar · {datetime.now().strftime('%d-%m-%Y %H:%M')}",
        size=9,
        color=COLOR_META,
    )

    OUT.mkdir(parents=True, exist_ok=True)
    out = OUT / "Informe_Validacion_Matriz_ESVAL_Fundo_Zapallar_FINAL.docx"
    doc.save(str(out))
    return out


def build_etapa5(app_e, err, estado, horas) -> Path:
    doc = Document(str(SRC / "Informe_Validacion_Etapa5_Zapallar_20260924_1136.docx"))
    _fix_margins(doc)
    delta = round(SENSUS_FIN - SENSUS_INI, 2)

    for p in doc.paragraphs:
        if p.text.strip().startswith("Validaciones:"):
            p.clear()
            _run(
                p,
                "Validación: 22-09-2026 14:30 → 28-09-2026 08:57 · Intervención placa 22-09-2026",
                size=10,
                color=COLOR_META,
            )
            break

    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith("2. Resumen") and i + 1 < len(doc.paragraphs):
            nxt = doc.paragraphs[i + 1]
            nxt.clear()
            _run(
                nxt,
                "Se detectaron pulsos anómalos; el error era de memoria de placa. "
                "Se reemplazó la memoria y se normalizó el punto. "
                f"Validación única 22–28/09: Sensus {_fmt(delta, 0)} m³ vs app {_fmt(app_e)} m³ "
                f"(error {err:.1f}%, {estado.lower()}).",
                size=10.5,
                color=COLOR_TEXTO,
            )
            break

    _rewrite_cover_kpi(
        doc,
        [
            ("Sensus vs app WES", f"Error {err:.1f} %", f"{_fmt(delta, 0)} / {_fmt(app_e)} m³"),
            ("Lecturas Sensus", f"{_fmt(SENSUS_INI, 0)} → {_fmt(SENSUS_FIN, 0)}", f"Δ {_fmt(delta, 0)} m³ · 22→28/09"),
            ("Red / medidor", "DN90 · techo ≈ 60 m³/h", "Q3 medidor 100 m³/h"),
        ],
    )
    _recolor_all_tables(doc)
    _clear_from(doc, "4. Cálculo")

    p = doc.add_paragraph()
    _run(p, "4. Cálculo de validación", bold=True, size=14, color=COLOR_TITULO)
    p2 = doc.add_paragraph()
    _run(p2, "4.1 Sensus vs app WES (22-09 14:30 → 28-09 08:57)", bold=True, size=12, color=COLOR_TITULO)

    _kpi3(
        doc,
        [
            ("Sensus vs app WES", f"Error {err:.1f} %", f"{_fmt(delta, 0)} / {_fmt(app_e)} m³"),
            ("Lecturas Sensus", f"{_fmt(SENSUS_INI, 0)} → {_fmt(SENSUS_FIN, 0)}", f"Δ {_fmt(delta, 0)} m³"),
            ("Ventana", "22 → 28/09/2026", f"{horas} horas app"),
        ],
    )

    p3 = doc.add_paragraph()
    _run(
        p3,
        "Una sola validación con lecturas fotográficas Sensus y el consumo de la app WES "
        "(hora Chile 14 del 22-09 → hora 08 del 28-09). El 22/09 la API quedó truncada tras el "
        "cambio de placa: esas horas se completan con el registro de placa (mismo criterio del "
        "informe original). App WES resulta mayor que el Δ Sensus.",
        size=10.5,
        color=COLOR_TEXTO,
    )

    f_ini = FOTOS_E5 / "lectura_20260922_1430_sensus_5144.png"
    if not f_ini.is_file():
        f_ini = FOTOS_E5 / "lectura_20260922_1430_sensus_5144_reloj.jpg"
    f_ini_p = _prep(f_ini, OUT / "foto_e5_ini.jpg", crop=True)
    f_fin_p = _prep(ASSET_E, OUT / "foto_e5_fin.jpg", crop=True)
    _fotos(
        doc,
        f_ini_p,
        f"Sensus · {_fmt(SENSUS_INI, 0)} m³ · {SENSUS_INI_DT}",
        f_fin_p,
        f"Sensus · {_fmt(SENSUS_FIN, 0)} m³ · {SENSUS_FIN_DT}",
    )

    _tabla_val(
        doc,
        "Validación Etapa N°5 — medidor Sensus",
        [
            "Etapa N°5",
            "22-09-2026\n14:30",
            _fmt(SENSUS_INI, 0),
            "28-09-2026\n08:57",
            _fmt(SENSUS_FIN, 0),
            _fmt(delta, 0),
            str(horas),
        ],
    )
    _totales(
        doc,
        [
            ("Total App WES", f"{_fmt(app_e)} m³"),
            ("Total Lectura Sensus", f"{_fmt(delta, 0)} m³"),
            ("% Error", f"{err:.1f} %"),
        ],
    )
    p4 = doc.add_paragraph()
    # App > Sensus → 1 − Sensus/App
    _run(
        p4,
        f"Error Sensus vs app: {err:.1f}% "
        f"(1 − {_fmt(delta, 0)}/{_fmt(app_e)}). {estado}.",
        size=10.5,
        color=COLOR_TEXTO,
    )

    p5 = doc.add_paragraph()
    p5.paragraph_format.space_before = Pt(10)
    _run(p5, "5. Conclusión", bold=True, size=14, color=COLOR_TITULO)
    p6 = doc.add_paragraph()
    _run(
        p6,
        "Se confirma falla de memoria de placa (sensor y voltajes OK). "
        f"Validación única 22–28/09: Sensus {_fmt(delta, 0)} m³ vs app {_fmt(app_e)} m³ "
        f"(error {err:.1f}%, {estado}). "
        "El agente IA revisará diariamente el caudal vs histórico y techo DN90 (60 m³/h).",
        size=10.5,
        color=COLOR_TEXTO,
    )
    foot = doc.add_paragraph()
    foot.paragraph_format.space_before = Pt(8)
    _run(
        foot,
        f"WES · Informe final · Fundo Zapallar · {datetime.now().strftime('%d-%m-%Y %H:%M')}",
        size=9,
        color=COLOR_META,
    )

    OUT.mkdir(parents=True, exist_ok=True)
    out = OUT / "Informe_Validacion_Etapa5_Zapallar_20260924_1136.docx"
    doc.save(str(out))
    return out


def drive_pdf(file_id: str, docx: Path, pdf_name: str) -> Path:
    svc = obtener_servicio_drive()
    media = MediaFileUpload(
        str(docx),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        resumable=True,
    )
    svc.files().update(fileId=file_id, media_body=media, fields="id").execute()
    req = svc.files().export_media(fileId=file_id, mimeType="application/pdf")
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    pdf = OUT / pdf_name
    pdf.write_bytes(buf.getvalue())
    return pdf


def enviar(pdfs: list[Path], resumen: str):
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO)
    msg["Subject"] = "Fundo Zapallar — Validación única 22–28/09/2026 (Matriz y Etapa N°5)"
    msg.attach(
        MIMEText(
            "Estimados Juan y Aníbal,\n\n"
            "Adjunto los PDF actualizados con una sola validación (22→28/09):\n\n"
            f"{resumen}\n\n"
            "Saludos,\nSistema WES\n",
            "plain",
            "utf-8",
        )
    )
    for pdf in pdfs:
        with open(pdf, "rb") as f:
            part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", "attachment", filename=pdf.name)
            msg.attach(part)
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(SMTP_USUARIO, _smtp_password())
        server.send_message(msg, to_addrs=TO)
    print(f"[OK] Correo → {', '.join(TO)}")


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    sm, se = _slots_m(), _slots_e()
    app_m = _sum("000027-01", sm)
    app_e = _sum_etapa5_chile()  # Chile TZ + hueco placa (app > Sensus)
    dm, de = round(ITRON_FIN - ITRON_INI, 2), round(SENSUS_FIN - SENSUS_INI, 2)
    em, ee = round(_err(dm, app_m), 1), round(_err(de, app_e), 1)
    stm, ste = _estado(em), _estado(ee)
    print(f"Matriz UNA validación: {dm} vs {app_m} → {em}%")
    print(f"Etapa5 UNA validación: Sensus {de} vs App {app_e} → {ee}% (app>sensus={app_e > de})")

    docx_m = build_matriz(app_m, em, stm, len(sm))
    docx_e = build_etapa5(app_e, ee, ste, len(se))
    pdf_m = drive_pdf(DRIVE_MATRIZ, docx_m, "Informe_Validacion_Matriz_ESVAL_Fundo_Zapallar_FINAL.pdf")
    pdf_e = drive_pdf(DRIVE_ETAPA5, docx_e, "Informe_Validacion_Etapa5_Zapallar_20260924_1136.pdf")
    print(f"[PDF] {pdf_m}")
    print(f"[PDF] {pdf_e}")

    resumen = (
        f"• Matriz: Itron {_fmt(dm)} vs App {_fmt(app_m)} → {em}% ({stm})\n"
        f"• Etapa 5: Sensus {_fmt(de, 0)} vs App {_fmt(app_e)} → {ee}% ({ste})"
    )
    enviar([pdf_m, pdf_e], resumen)
    (OUT / "resultado.json").write_text(
        json.dumps(
            {
                "matriz": {"delta": dm, "app": app_m, "err": em, "pdf": str(pdf_m)},
                "etapa5": {"delta": de, "app": app_e, "err": ee, "pdf": str(pdf_e)},
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
