"""
Reporte de puntos con horas perdidas (huecos en la serie horaria),
ordenado de mayor a menor, con gráficos por semana ISO.

Periodo operativo 2026 (hora Chile, lunes–domingo):
  - Semana 32: 03–09 ago
  - Semana 33: 10–16 ago
  - Semana 34: 17–23 ago
  - Semana 35 parcial: 24 y 25 ago (2 primeros días)

Hora perdida = hora Chile esperada sin registro en dates.measures.csv.
Una hora con valor 0 SÍ cuenta como dato (no es hora perdida).

El ranking se parte en dos:
  - desconexión: lastUpdate de la app ≥ 2 h (mismo criterio que puntos en cero)
  - conectado con huecos: lastUpdate fresco, pero faltan horas en la serie
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import requests
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from matplotlib.patches import Patch

from reporte_puntos_en_cero import (
    HORAS_UMBRAL_CONEXION_APP,
    _fmt_antiguedad,
    obtener_estado_conexion_nodo,
    obtener_todos_los_nodos,
)
from wes_paths import wes_scripts_root

try:
    from zoneinfo import ZoneInfo

    CHILE_TZ = ZoneInfo("America/Santiago")
except Exception:
    CHILE_TZ = timezone(timedelta(hours=-4))

BASE_URL = os.environ.get(
    "WES_API_BASE_URL", "http://104.248.53.141:7003/wes/api/acl-node/v1"
).rstrip("/")

MAX_WORKERS = max(4, int(os.environ.get("WES_HORAS_PERDIDAS_WORKERS", "24")))
REQUEST_TIMEOUT = 25

# Semanas ISO 2026 pedidas por operación.
SEMANA_DEFS: List[Tuple[int, date, date, str]] = [
    (32, date(2026, 8, 3), date(2026, 8, 9), "Semana 32 (03–09 ago)"),
    (33, date(2026, 8, 10), date(2026, 8, 16), "Semana 33 (10–16 ago)"),
    (34, date(2026, 8, 17), date(2026, 8, 23), "Semana 34 (17–23 ago)"),
    (35, date(2026, 8, 24), date(2026, 8, 25), "Semana 35 (24–25 ago, 2 días)"),
]
DIA_24 = date(2026, 8, 24)
DIA_25 = date(2026, 8, 25)

COLOR_S32 = "#1F4E79"
COLOR_S33 = "#2E86AB"
COLOR_S34 = "#E67E22"
COLOR_S35 = "#C0392B"
COLOR_DIA24 = "#1F4E79"
COLOR_DIA25 = "#C0392B"
COLOR_DESC = "#C0392B"
COLOR_HUECOS = "#2E86AB"
_SESSION = requests.Session()
_ADAPTER = requests.adapters.HTTPAdapter(pool_connections=32, pool_maxsize=32, max_retries=2)
_SESSION.mount("http://", _ADAPTER)
_SESSION.mount("https://", _ADAPTER)


def _fmt_int(n: float | int) -> str:
    return f"{int(round(n)):,}".replace(",", ".")


def _fmt_pct(part: float, total: float) -> str:
    if total <= 0:
        return "—"
    return f"{(part / total) * 100:.1f}".replace(".", ",") + "%"


def _trunc(texto: str, n: int = 36) -> str:
    t = (texto or "").strip()
    return t if len(t) <= n else t[: n - 1] + "…"


def _iter_days(inicio: date, fin: date) -> List[date]:
    out: List[date] = []
    cur = inicio
    while cur <= fin:
        out.append(cur)
        cur += timedelta(days=1)
    return out


TODOS_LOS_DIAS: List[date] = _iter_days(date(2026, 8, 3), date(2026, 8, 25))
SEMANA_POR_DIA: Dict[date, int] = {}
for num, ini, fin, _lbl in SEMANA_DEFS:
    for d in _iter_days(ini, fin):
        SEMANA_POR_DIA[d] = num


def horas_esperadas_dia(dia: date, ahora_chile: datetime) -> List[int]:
    """Horas 0–23 que ya deberían tener registro a la hora del reporte."""
    hoy = ahora_chile.date()
    if dia < hoy:
        return list(range(24))
    if dia > hoy:
        return []
    # La marca HH:00 del día en curso aparece cerca de esa hora; no exigir la hora actual.
    corte = max(0, ahora_chile.hour)
    return list(range(corte))


def _parse_horas_csv(csv_content: str, dia: date) -> Dict[int, float]:
    """Hora civil del TIME (convención app WES) para un día."""
    target = dia.strftime("%Y-%m-%d")
    acc: Dict[int, float] = {}
    lines = (csv_content or "").strip().splitlines()
    if len(lines) <= 1:
        return acc
    for line in lines[1:]:
        line = line.strip()
        if not line or "," not in line:
            continue
        time_str, value_str = line.split(",", 1)
        ts = time_str.strip()
        if not ts.startswith(target) or "T" not in ts:
            continue
        try:
            hi = int(ts[11:13])
            val = float(value_str.strip().replace(" ", "").replace(",", "."))
        except (ValueError, TypeError, IndexError):
            continue
        if 0 <= hi < 24:
            acc[hi] = acc.get(hi, 0.0) + val
    return acc


def _parse_dias_rango_csv(csv_content: str) -> Dict[date, float]:
    """CSV de rango (una fila por día, TIME=YYYY-MM-DD)."""
    out: Dict[date, float] = {}
    lines = (csv_content or "").strip().splitlines()
    if len(lines) <= 1:
        return out
    for line in lines[1:]:
        line = line.strip()
        if not line or "," not in line:
            continue
        time_str, value_str = line.split(",", 1)
        ts = time_str.strip()
        if "T" in ts:
            ts = ts.split("T", 1)[0]
        try:
            dia = date.fromisoformat(ts[:10])
            val = float(value_str.strip().replace(" ", "").replace(",", "."))
        except (ValueError, TypeError):
            continue
        out[dia] = out.get(dia, 0.0) + val
    return out


def _get_csv(node_id: str, start: date, end: date) -> Tuple[str, bool]:
    url = f"{BASE_URL}/nodes/{node_id}/dates.measures.csv"
    params = [("start", start.strftime("%d%m%Y")), ("end", end.strftime("%d%m%Y"))]
    try:
        r = _SESSION.get(url, params=params, timeout=REQUEST_TIMEOUT)
        if r.status_code != 200:
            return "", False
        return r.text, True
    except requests.RequestException:
        return "", False


@dataclass
class ResultadoNodo:
    node_id: str
    node_name: str
    company_id: str
    company_name: str
    perdidas_por_dia: Dict[str, int] = field(default_factory=dict)
    esperadas_por_dia: Dict[str, int] = field(default_factory=dict)
    error: str = ""
    last_update: str = ""
    wes_status: str = ""
    last_update_dt: Optional[datetime] = None
    horas_sin_conexion: Optional[float] = None
    desconectado: bool = False

    def perdidas_dia(self, dia: date) -> int:
        return int(self.perdidas_por_dia.get(dia.isoformat(), 0))

    def esperadas_dia(self, dia: date) -> int:
        return int(self.esperadas_por_dia.get(dia.isoformat(), 0))

    def perdidas_semana(self, semana: int) -> int:
        return sum(
            self.perdidas_dia(d) for d, s in SEMANA_POR_DIA.items() if s == semana
        )

    def esperadas_semana(self, semana: int) -> int:
        return sum(
            self.esperadas_dia(d) for d, s in SEMANA_POR_DIA.items() if s == semana
        )

    def perdidas_total(self) -> int:
        return sum(int(v) for v in self.perdidas_por_dia.values())

    def esperadas_total(self) -> int:
        return sum(int(v) for v in self.esperadas_por_dia.values())

    def etiqueta_estado(self) -> str:
        if not self.desconectado:
            return "Conectado (huecos)"
        if self.horas_sin_conexion is None:
            return "Desconectado (sin lastUpdate)"
        return f"Desconectado ({_fmt_antiguedad(self.horas_sin_conexion)})"


def analizar_nodo(nodo: Dict[str, str], ahora_chile: datetime) -> ResultadoNodo:
    res = ResultadoNodo(
        node_id=nodo["nodeId"],
        node_name=nodo.get("nodeName", ""),
        company_id=nodo.get("companyId", ""),
        company_name=nodo.get("companyName", ""),
    )
    texto, ok = _get_csv(res.node_id, TODOS_LOS_DIAS[0], TODOS_LOS_DIAS[-1])
    dias_con_total: Dict[date, float] = _parse_dias_rango_csv(texto) if ok else {}

    for dia in TODOS_LOS_DIAS:
        esperadas = horas_esperadas_dia(dia, ahora_chile)
        res.esperadas_por_dia[dia.isoformat()] = len(esperadas)
        if not esperadas:
            res.perdidas_por_dia[dia.isoformat()] = 0
            continue

        if not ok:
            # Sin respuesta del rango: intentar el día puntual.
            csv_dia, ok_dia = _get_csv(res.node_id, dia, dia)
            if not ok_dia:
                res.perdidas_por_dia[dia.isoformat()] = len(esperadas)
                res.error = "sin respuesta API"
                continue
            presentes = set(_parse_horas_csv(csv_dia, dia).keys())
            res.perdidas_por_dia[dia.isoformat()] = sum(
                1 for h in esperadas if h not in presentes
            )
            continue

        if dia not in dias_con_total:
            res.perdidas_por_dia[dia.isoformat()] = len(esperadas)
            continue

        csv_dia, ok_dia = _get_csv(res.node_id, dia, dia)
        if not ok_dia:
            # El total diario existe; no pudimos ver huecos intra-día → 0 perdidas
            # (no inflar el ranking por un fallo puntual).
            res.perdidas_por_dia[dia.isoformat()] = 0
            continue
        presentes = set(_parse_horas_csv(csv_dia, dia).keys())
        res.perdidas_por_dia[dia.isoformat()] = sum(
            1 for h in esperadas if h not in presentes
        )
    return res


def enriquecer_conexion(resultados: Sequence[ResultadoNodo], ahora_chile: datetime) -> None:
    """Marca desconectado si lastUpdate de la app tiene ≥ umbral (default 2 h)."""

    def _one(r: ResultadoNodo) -> None:
        try:
            st = obtener_estado_conexion_nodo(r.node_id)
        except Exception:
            r.desconectado = True
            r.last_update = "—"
            r.wes_status = "—"
            return
        lu = st.get("lastUpdate")
        r.wes_status = str(st.get("wesStatus") or "—")
        if hasattr(lu, "strftime"):
            r.last_update_dt = lu  # type: ignore[assignment]
            r.last_update = lu.strftime("%d-%m-%Y %H:%M")
        else:
            r.last_update_dt = None
            r.last_update = str(st.get("lastUpdateRaw") or "—")
        if r.last_update_dt is None:
            r.desconectado = True
            r.horas_sin_conexion = None
            return
        dt = r.last_update_dt
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=CHILE_TZ)
        r.horas_sin_conexion = (ahora_chile - dt.astimezone(CHILE_TZ)).total_seconds() / 3600.0
        r.desconectado = r.horas_sin_conexion >= float(HORAS_UMBRAL_CONEXION_APP)

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        list(ex.map(_one, resultados))


def particionar_perdidas(
    items: Sequence[ResultadoNodo],
) -> Tuple[List[ResultadoNodo], List[ResultadoNodo]]:
    """(desconectados, conectados-con-huecos), ambos mayor a menor por total."""
    desc = [r for r in items if r.desconectado]
    huec = [r for r in items if not r.desconectado]
    desc.sort(key=lambda r: r.perdidas_total(), reverse=True)
    huec.sort(key=lambda r: r.perdidas_total(), reverse=True)
    return desc, huec


def cargar_desde_json(path: Path) -> List[ResultadoNodo]:
    raw = json.loads(path.read_text(encoding="utf-8"))
    out: List[ResultadoNodo] = []
    for d in raw:
        r = ResultadoNodo(
            node_id=str(d.get("nodeId", "")),
            node_name=str(d.get("nodeName", "")),
            company_id=str(d.get("companyId", "")),
            company_name=str(d.get("companyName", "")),
            perdidas_por_dia={k: int(v) for k, v in (d.get("perdidasPorDia") or {}).items()},
            esperadas_por_dia={k: int(v) for k, v in (d.get("esperadasPorDia") or {}).items()},
            error=str(d.get("error") or ""),
        )
        out.append(r)
    return out


def _set_cell_shading(cell, hex_color: str) -> None:
    shading = parse_xml(
        f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'w:val="clear" w:fill="{hex_color}"/>'
    )
    tc_pr = cell._tc.get_or_add_tcPr()
    old = tc_pr.find(qn("w:shd"))
    if old is not None:
        tc_pr.remove(old)
    tc_pr.append(shading)


def _set_run_white_bold(cell, size: int = 8) -> None:
    for p in cell.paragraphs:
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        for run in p.runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(size)


def _fill_table(
    table,
    rows: List[Sequence[str]],
    header_fill: str = "1F4E79",
    highlight_col: Optional[int] = None,
) -> None:
    for j, header in enumerate(rows[0]):
        cell = table.rows[0].cells[j]
        cell.text = str(header)
        _set_cell_shading(cell, header_fill)
        _set_run_white_bold(cell, 8)
    for i, row in enumerate(rows[1:], start=1):
        for j, val in enumerate(row):
            cell = table.rows[i].cells[j]
            cell.text = str(val)
            for p in cell.paragraphs:
                p.alignment = (
                    WD_PARAGRAPH_ALIGNMENT.CENTER if j != 1 else WD_PARAGRAPH_ALIGNMENT.LEFT
                )
                for run in p.runs:
                    run.font.size = Pt(8)
                    run.font.color.rgb = RGBColor(0, 0, 0)
            if i % 2 == 0:
                _set_cell_shading(cell, "F2F2F2")
            if highlight_col is not None and j == highlight_col and i > 0:
                try:
                    n = int(str(val).split()[0].replace(".", ""))
                except ValueError:
                    n = 0
                if n >= 100:
                    _set_cell_shading(cell, "F4CCCC")
                elif n >= 24:
                    _set_cell_shading(cell, "FCE4D6")


def _add_picture(doc: Document, path: Path, width: float = 6.4) -> None:
    if not path.exists():
        return
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.add_run().add_picture(str(path), width=Inches(width))


def grafico_totales_semana(
    resultados: Sequence[ResultadoNodo],
    esperados: Dict[int, int],
    out: Path,
) -> Path:
    labels = [f"S{n}" for n, *_ in SEMANA_DEFS]
    desc = [
        sum(r.perdidas_semana(n) for r in resultados if r.desconectado)
        for n, *_ in SEMANA_DEFS
    ]
    huec = [
        sum(r.perdidas_semana(n) for r in resultados if not r.desconectado)
        for n, *_ in SEMANA_DEFS
    ]
    fig, ax = plt.subplots(figsize=(8.8, 4.4))
    fig.patch.set_facecolor("white")
    x = range(len(labels))
    ax.bar(x, desc, color=COLOR_DESC, width=0.62, label="Desconexión", zorder=3)
    ax.bar(x, huec, bottom=desc, color=COLOR_HUECOS, width=0.62, label="Conectado (huecos)", zorder=3)
    ax.set_xticks(list(x), labels)
    ax.set_ylabel("Horas perdidas (flota)")
    ax.set_title("Horas sin dato por semana — desconexión vs huecos")
    ax.legend(loc="upper right", fontsize=8)
    ax.grid(axis="y", linestyle=":", alpha=0.5, zorder=0)
    ax.set_axisbelow(True)
    for i, n in enumerate([32, 33, 34, 35]):
        total = desc[i] + huec[i]
        ax.text(
            i,
            total,
            f"{_fmt_int(total)}\n({_fmt_pct(total, esperados[n])})",
            ha="center",
            va="bottom",
            fontsize=8,
        )
    ymax = max((max(d + h for d, h in zip(desc, huec)) if desc else 0) * 1.22, 1.0)
    ax.set_ylim(0, ymax)
    fig.tight_layout()
    fig.savefig(out, dpi=160, bbox_inches="tight")
    plt.close(fig)
    return out


def grafico_ranking(
    items: Sequence[ResultadoNodo],
    value_fn,
    titulo: str,
    out: Path,
    color: Optional[str] = None,
    top: int = 18,
    leyenda_estado: bool = True,
) -> Path:
    subset = [r for r in items if value_fn(r) > 0][:top]
    if not subset:
        fig, ax = plt.subplots(figsize=(8.8, 3.2))
        ax.text(0.5, 0.5, "Sin horas perdidas en este recorte", ha="center", va="center")
        ax.axis("off")
        fig.savefig(out, dpi=160, bbox_inches="tight")
        plt.close(fig)
        return out
    orden = list(reversed(subset))
    labels = [f"{_trunc(r.node_name, 28)} ({r.node_id})" for r in orden]
    vals = [value_fn(r) for r in orden]
    if color:
        colors = [color] * len(orden)
    else:
        colors = [COLOR_DESC if r.desconectado else COLOR_HUECOS for r in orden]
    fig, ax = plt.subplots(figsize=(8.8, max(3.4, 0.32 * len(subset) + 1.4)))
    fig.patch.set_facecolor("white")
    ax.barh(labels, vals, color=colors, zorder=3)
    ax.set_xlabel("Horas perdidas")
    ax.set_title(titulo)
    ax.grid(axis="x", linestyle=":", alpha=0.5, zorder=0)
    ax.set_axisbelow(True)
    if leyenda_estado and color is None:
        ax.legend(
            handles=[
                Patch(facecolor=COLOR_DESC, label="Desconectado"),
                Patch(facecolor=COLOR_HUECOS, label="Conectado (huecos)"),
            ],
            loc="lower right",
            fontsize=8,
        )
    for y, v in enumerate(vals):
        ax.text(v, y, f" {_fmt_int(v)}", va="center", fontsize=7)
    ax.set_xlim(0, max((max(vals) if vals else 0) * 1.18, 1.0))
    fig.tight_layout()
    fig.savefig(out, dpi=160, bbox_inches="tight")
    plt.close(fig)
    return out


def grafico_apilado_top(
    items: Sequence[ResultadoNodo],
    out: Path,
    top: int = 15,
    titulo: str = "Top puntos — horas perdidas apiladas por semana",
) -> Path:
    subset = list(items)[:top]
    if not subset:
        fig, ax = plt.subplots(figsize=(8.8, 3.2))
        ax.axis("off")
        fig.savefig(out, dpi=160, bbox_inches="tight")
        plt.close(fig)
        return out
    labels = [_trunc(r.node_name, 22) for r in reversed(subset)]
    s32 = [r.perdidas_semana(32) for r in reversed(subset)]
    s33 = [r.perdidas_semana(33) for r in reversed(subset)]
    s34 = [r.perdidas_semana(34) for r in reversed(subset)]
    s35 = [r.perdidas_semana(35) for r in reversed(subset)]
    fig, ax = plt.subplots(figsize=(8.8, max(3.8, 0.34 * len(subset) + 1.6)))
    fig.patch.set_facecolor("white")
    ax.barh(labels, s32, color=COLOR_S32, label="S32", zorder=3)
    ax.barh(labels, s33, left=s32, color=COLOR_S33, label="S33", zorder=3)
    left34 = [a + b for a, b in zip(s32, s33)]
    ax.barh(labels, s34, left=left34, color=COLOR_S34, label="S34", zorder=3)
    left35 = [a + b for a, b in zip(left34, s34)]
    ax.barh(labels, s35, left=left35, color=COLOR_S35, label="S35 (2 d)", zorder=3)
    ax.set_xlabel("Horas perdidas")
    ax.set_title(titulo)
    ax.legend(loc="lower right", fontsize=8)
    ax.grid(axis="x", linestyle=":", alpha=0.5, zorder=0)
    ax.set_axisbelow(True)
    fig.tight_layout()
    fig.savefig(out, dpi=160, bbox_inches="tight")
    plt.close(fig)
    return out


def grafico_semana35_dias(
    items: Sequence[ResultadoNodo],
    out: Path,
    top: int = 15,
) -> Path:
    subset = [r for r in items if r.perdidas_semana(35) > 0][:top]
    if not subset:
        fig, ax = plt.subplots(figsize=(8.8, 3.2))
        ax.text(0.5, 0.5, "Sin horas perdidas el 24–25 ago", ha="center", va="center")
        ax.axis("off")
        fig.savefig(out, dpi=160, bbox_inches="tight")
        plt.close(fig)
        return out
    labels = [_trunc(r.node_name, 22) for r in reversed(subset)]
    h24 = [r.perdidas_dia(DIA_24) for r in reversed(subset)]
    h25 = [r.perdidas_dia(DIA_25) for r in reversed(subset)]
    y = range(len(subset))
    fig, ax = plt.subplots(figsize=(8.8, max(3.8, 0.36 * len(subset) + 1.6)))
    fig.patch.set_facecolor("white")
    h = 0.38
    ax.barh([i + h / 2 for i in y], h24, height=h, color=COLOR_DIA24, label="Lun 24 ago", zorder=3)
    ax.barh([i - h / 2 for i in y], h25, height=h, color=COLOR_DIA25, label="Mar 25 ago (parcial)", zorder=3)
    ax.set_yticks(list(y), labels)
    ax.set_xlabel("Horas perdidas")
    ax.set_title("Semana 35 — 24 vs 25 agosto (hora Chile)")
    ax.legend(loc="lower right", fontsize=8)
    ax.grid(axis="x", linestyle=":", alpha=0.5, zorder=0)
    ax.set_axisbelow(True)
    fig.tight_layout()
    fig.savefig(out, dpi=160, bbox_inches="tight")
    plt.close(fig)
    return out


def convertir_a_pdf(docx_path: Path) -> Optional[Path]:
    pdf_path = docx_path.with_suffix(".pdf")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("[ADVERTENCIA] No hay LibreOffice; se deja solo el Word.")
        return None
    subprocess.run(
        [
            soffice,
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
            "--outdir",
            str(docx_path.parent),
            str(docx_path),
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    return pdf_path if pdf_path.exists() else None


def _filas_ranking_periodo(items: Sequence[ResultadoNodo]) -> List[Sequence[str]]:
    rows: List[Sequence[str]] = [
        ("#", "Punto", "Empresa", "S32", "S33", "S34", "S35*", "Total", "%", "Estado", "Última conexión"),
    ]
    for i, r in enumerate(items, start=1):
        rows.append(
            (
                str(i),
                f"{r.node_name}\n{r.node_id}",
                r.company_name,
                _fmt_int(r.perdidas_semana(32)),
                _fmt_int(r.perdidas_semana(33)),
                _fmt_int(r.perdidas_semana(34)),
                _fmt_int(r.perdidas_semana(35)),
                _fmt_int(r.perdidas_total()),
                _fmt_pct(r.perdidas_total(), r.esperadas_total()),
                r.etiqueta_estado(),
                r.last_update or "—",
            )
        )
    return rows


def _agregar_tabla_ranking(
    doc: Document,
    items: Sequence[ResultadoNodo],
    header_fill: str,
    highlight_col: int = 7,
) -> None:
    rows = _filas_ranking_periodo(items)
    if len(rows) == 1:
        doc.add_paragraph("Ningún punto en este grupo.")
        return
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    _fill_table(t, rows, header_fill=header_fill, highlight_col=highlight_col)


def crear_reporte_word(
    resultados: List[ResultadoNodo],
    ahora_chile: datetime,
    output_dir: Path,
    charts: Dict[str, Path],
    top_chart: int,
) -> Path:
    con_perdidas = [r for r in resultados if r.perdidas_total() > 0]
    con_perdidas.sort(key=lambda r: r.perdidas_total(), reverse=True)
    desc, huec = particionar_perdidas(con_perdidas)
    n_total = len(resultados)
    n_con = len(con_perdidas)
    h_desc = sum(r.perdidas_total() for r in desc)
    h_huec = sum(r.perdidas_total() for r in huec)

    tot_sem = {n: sum(r.perdidas_semana(n) for r in resultados) for n, *_ in SEMANA_DEFS}
    esp_sem = {n: sum(r.esperadas_semana(n) for r in resultados) for n, *_ in SEMANA_DEFS}
    tot_all = sum(tot_sem.values())
    esp_all = sum(esp_sem.values())

    tot_24 = sum(r.perdidas_dia(DIA_24) for r in resultados)
    tot_25 = sum(r.perdidas_dia(DIA_25) for r in resultados)
    esp_24 = sum(r.esperadas_dia(DIA_24) for r in resultados)
    esp_25 = sum(r.esperadas_dia(DIA_25) for r in resultados)

    # Continuidad S32-34 → S35
    peores_3234 = sorted(
        resultados,
        key=lambda r: r.perdidas_semana(32) + r.perdidas_semana(33) + r.perdidas_semana(34),
        reverse=True,
    )
    peores_3234 = [r for r in peores_3234 if (r.perdidas_semana(32) + r.perdidas_semana(33) + r.perdidas_semana(34)) > 0][:12]
    nuevos_s35 = [
        r
        for r in con_perdidas
        if r.perdidas_semana(35) > 0
        and (r.perdidas_semana(32) + r.perdidas_semana(33) + r.perdidas_semana(34)) == 0
    ]
    siguen_s35 = [r for r in peores_3234 if r.perdidas_semana(35) > 0]
    mejoran_s35 = [r for r in peores_3234 if r.perdidas_semana(35) == 0]

    doc = Document()
    title = doc.add_heading("REPORTE DE PUNTOS CON HORAS PERDIDAS", 0)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    title.runs[0].font.color.rgb = RGBColor(192, 0, 0)
    title.runs[0].bold = True

    sub = doc.add_paragraph(
        "Rankings separados: desconexión vs conectado con huecos · "
        "Semanas ISO 32, 33 y 34 de 2026 y los 2 primeros días de la semana 35"
    )
    sub.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for run in sub.runs:
        run.font.size = Pt(11)
        run.italic = True

    gen = doc.add_paragraph(
        f"Generado: {ahora_chile.strftime('%d-%m-%Y %H:%M')} hora Chile · "
        f"Serie horaria API dates.measures.csv · Universo: {n_total} puntos "
        f"(mismas exclusiones que puntos en cero)"
    )
    gen.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    gen.runs[0].font.size = Pt(9)
    gen.runs[0].font.color.rgb = RGBColor(80, 80, 80)

    doc.add_heading("1. Criterio", 1)
    doc.add_paragraph(
        "Se cuenta como hora perdida cada hora Chile que ya debería tener un registro "
        "en la serie horaria y no lo tiene. Una hora con consumo 0 no se considera perdida: "
        "el punto sí reportó. El 25 de agosto es día en curso: solo se exigen las horas "
        f"00:00 a {max(0, ahora_chile.hour - 1):02d}:00 (corte {ahora_chile.strftime('%H:%M')} Chile)."
    )
    doc.add_paragraph(
        "El ranking se parte en dos según el estado de conexión actual de la app "
        f"(lastUpdate ≥ {HORAS_UMBRAL_CONEXION_APP:.0f} h → desconectado, aunque wesStatus siga ON). "
        "1) Horas perdidas por desconexión: el punto está caído ahora; esas horas del periodo "
        "se listan aparte. 2) Conectado con huecos: lastUpdate fresco, pero la serie horaria "
        "tiene faltantes (no recupera histórico, o ya se reconectó y quedan huecos de semanas previas)."
    )
    doc.add_paragraph(
        "Semana 32: lun 03 – dom 09 ago (7 días, 168 h/punto). "
        "Semana 33: lun 10 – dom 16 ago. "
        "Semana 34: lun 17 – dom 23 ago. "
        "Semana 35: lun 24 y mar 25 ago (parcial; no es semana cerrada)."
    )

    doc.add_heading("2. Resumen ejecutivo", 1)
    doc.add_paragraph(
        f"Puntos con al menos 1 h perdida: {n_con} de {n_total} ({_fmt_pct(n_con, n_total)}). "
        f"Horas perdidas de la flota: {_fmt_int(tot_all)} de {_fmt_int(esp_all)} esperadas "
        f"({_fmt_pct(tot_all, esp_all)})."
    )
    p_split = doc.add_paragraph()
    p_split.add_run("Separación del ranking. ").bold = True
    p_split.add_run(
        f"Por desconexión (caídos ahora): {len(desc)} puntos, {_fmt_int(h_desc)} h "
        f"({_fmt_pct(h_desc, tot_all)} del total perdido). "
        f"Conectados con huecos: {len(huec)} puntos, {_fmt_int(h_huec)} h "
        f"({_fmt_pct(h_huec, tot_all)})."
    )
    resumen_rows = [
        ("Semana", "Desde", "Hasta", "Horas perdidas", "% de las esperadas"),
        (
            "32",
            "03-08-2026",
            "09-08-2026",
            _fmt_int(tot_sem[32]),
            _fmt_pct(tot_sem[32], esp_sem[32]),
        ),
        (
            "33",
            "10-08-2026",
            "16-08-2026",
            _fmt_int(tot_sem[33]),
            _fmt_pct(tot_sem[33], esp_sem[33]),
        ),
        (
            "34",
            "17-08-2026",
            "23-08-2026",
            _fmt_int(tot_sem[34]),
            _fmt_pct(tot_sem[34], esp_sem[34]),
        ),
        (
            "35 (2 d)",
            "24-08-2026",
            "25-08-2026",
            _fmt_int(tot_sem[35]),
            _fmt_pct(tot_sem[35], esp_sem[35]),
        ),
    ]
    t = doc.add_table(rows=len(resumen_rows), cols=5)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _fill_table(t, resumen_rows)

    doc.add_paragraph("")
    _add_picture(doc, charts["totales"], 6.3)

    # Tendencia
    doc.add_heading("Lectura de las semanas 32–34", 2)
    orden_sem = [tot_sem[32], tot_sem[33], tot_sem[34]]
    if orden_sem[2] > orden_sem[0] and orden_sem[2] > orden_sem[1]:
        tendencia = (
            "Las horas perdidas de la flota suben hacia la semana 34 respecto de 32 y 33: "
            "hay más huecos de telemetría en la última semana completa."
        )
    elif orden_sem[2] < orden_sem[0] and orden_sem[2] < orden_sem[1]:
        tendencia = (
            "La semana 34 cierra con menos horas perdidas de flota que 32 y 33: "
            "mejora relativa de cobertura horaria."
        )
    else:
        tendencia = (
            "Entre las semanas 32, 33 y 34 la pérdida de horas de la flota se mantiene "
            "en un rango similar, sin un quiebre fuerte de una semana a otra."
        )
    doc.add_paragraph(tendencia)

    doc.add_heading("3. Rankings separados (mayor a menor)", 1)
    doc.add_paragraph(
        "Primero los puntos que están desconectados ahora; después los que están "
        "conectados pero acumulan horas sin dato. El detalle por semana va en gráficos "
        "aparte (secciones 4 y 5), sin tablas de cantidades."
    )

    doc.add_heading("3.1 Ranking — horas perdidas por desconexión", 2)
    doc.add_paragraph(
        f"{len(desc)} puntos con lastUpdate ≥ {HORAS_UMBRAL_CONEXION_APP:.0f} h "
        f"(o sin lastUpdate). Total {_fmt_int(h_desc)} h. Prioridad de visita / reconexión."
    )
    _add_picture(doc, charts["ranking_desconexion"], 6.3)

    doc.add_heading("3.2 Ranking — conectados con huecos (no es desconexión actual)", 2)
    doc.add_paragraph(
        f"{len(huec)} puntos con lastUpdate fresco y horas faltantes en la serie. "
        f"Total {_fmt_int(h_huec)} h. Incluye reconectados que no recuperan histórico "
        "y huecos de semanas previas."
    )
    _add_picture(doc, charts["ranking_huecos"], 6.3)

    for num, ini, fin, label in SEMANA_DEFS[:3]:
        doc.add_page_break()
        doc.add_heading(f"4.{num - 31} {label}", 1)
        orden = sorted(resultados, key=lambda r: r.perdidas_semana(num), reverse=True)
        con = [r for r in orden if r.perdidas_semana(num) > 0]
        n_desc = sum(1 for r in con if r.desconectado)
        n_huec = sum(1 for r in con if not r.desconectado)
        doc.add_paragraph(
            f"Puntos con horas perdidas: {len(con)} de {n_total} "
            f"({n_desc} por desconexión, {n_huec} conectados con huecos). "
            f"Total flota: {_fmt_int(tot_sem[num])} h "
            f"({_fmt_pct(tot_sem[num], esp_sem[num])} de las esperadas). "
            "Un gráfico por grupo; sin tablas de cantidades."
        )
        doc.add_heading("Desconexión", 2)
        _add_picture(doc, charts[f"s{num}_desc"], 6.3)
        doc.add_heading("Conectados con huecos", 2)
        _add_picture(doc, charts[f"s{num}_huecos"], 6.3)

    # Semana 35
    doc.add_page_break()
    doc.add_heading("5. Semana 35 — qué pasa el 24 y el 25 de agosto", 1)
    doc.add_paragraph(
        f"La semana 35 ISO 2026 empieza el lunes 24 de agosto. Al cierre de este informe "
        f"solo hay {DIA_24.strftime('%d/%m')} (día completo, 24 h) y {DIA_25.strftime('%d/%m')} "
        f"(día en curso, {esp_25 // max(n_total, 1)} h exigidas por punto). "
        "No se debe comparar el total de S35 con 168 h de una semana cerrada."
    )
    doc.add_paragraph(
        f"Lunes 24: {_fmt_int(tot_24)} h de flota ({_fmt_pct(tot_24, esp_24)}). "
        f"Martes 25 (hasta {ahora_chile.strftime('%H:%M')}): {_fmt_int(tot_25)} h "
        f"({_fmt_pct(tot_25, esp_25)})."
    )
    daily_32_34 = (tot_sem[32] + tot_sem[33] + tot_sem[34]) / 21.0 if n_total else 0
    if daily_32_34:
        factor = 24.0 / max(esp_25 / max(n_total, 1), 1)
        proy_25 = tot_25 * factor
        ritmo = doc.add_paragraph()
        ritmo.add_run("Ritmo diario. ").bold = True
        ritmo.add_run(
            f"Promedio S32–S34: {_fmt_int(daily_32_34)} h/día. "
            f"Lunes 24: {_fmt_int(tot_24)} h "
            f"({((tot_24 / daily_32_34) - 1) * 100:+.0f}% vs ese promedio). "
            f"Martes 25 parcial {_fmt_int(tot_25)} h (proyección a 24 h ≈ {_fmt_int(proy_25)} h, referencial)."
        )

    doc.add_heading("Desconexión", 2)
    _add_picture(doc, charts["s35_desc"], 6.3)
    doc.add_heading("Conectados con huecos", 2)
    _add_picture(doc, charts["s35_huecos"], 6.3)
    doc.add_heading("Lunes 24 vs martes 25", 2)
    _add_picture(doc, charts["s35_dias"], 6.3)

    orden35 = sorted(resultados, key=lambda r: r.perdidas_semana(35), reverse=True)
    con35 = [r for r in orden35 if r.perdidas_semana(35) > 0]
    n35_desc = sum(1 for r in con35 if r.desconectado)
    doc.add_paragraph(
        f"{len(con35)} puntos con horas perdidas el 24–25: "
        f"{n35_desc} desconectados ahora y {len(con35) - n35_desc} conectados con huecos."
    )

    doc.add_heading("Continuidad respecto de S32–S34", 2)
    doc.add_paragraph(
        f"De los 12 puntos con más horas perdidas en las tres semanas cerradas, "
        f"{len(siguen_s35)} siguen perdiendo horas el 24–25 ago y "
        f"{len(mejoran_s35)} aparecen en cero pérdidas en estos dos días "
        f"(recuperaron serie horaria o no acumulan huecos aún). "
        f"Puntos nuevos (0 h perdidas en S32–S34 y sí en S35): {len(nuevos_s35)}."
    )
    if siguen_s35:
        nombres_siguen = ", ".join(f"{r.node_name} ({r.node_id})" for r in siguen_s35)
        p = doc.add_paragraph()
        p.add_run("Siguen mal (desconexión o huecos el 24–25): ").bold = True
        p.add_run(nombres_siguen + ".")
    if mejoran_s35:
        nombres = ", ".join(f"{r.node_name} ({r.node_id})" for r in mejoran_s35)
        p = doc.add_paragraph()
        p.add_run("Top S32–S34 sin horas perdidas el 24–25: ").bold = True
        p.add_run(nombres + ".")
    if nuevos_s35:
        nombres_n = ", ".join(
            f"{r.node_name} ({r.node_id}, 24={_fmt_int(r.perdidas_dia(DIA_24))} h, "
            f"25={_fmt_int(r.perdidas_dia(DIA_25))} h)"
            for r in nuevos_s35
        )
        p = doc.add_paragraph()
        p.add_run("Incidentes nuevos en S35: ").bold = True
        p.add_run(nombres_n + ".")
    else:
        doc.add_paragraph(
            "No aparecen puntos nuevos: quienes pierden horas el 24–25 ya venían "
            "con huecos en S32, S33 o S34."
        )

    doc.add_heading("24 vs 25 a nivel flota", 2)
    if tot_24 == 0 and tot_25 == 0:
        doc.add_paragraph("La flota no registra horas perdidas en ninguno de los dos días.")
    elif tot_25 == 0 and tot_24 > 0:
        doc.add_paragraph(
            "El martes 25 (tramo ya transcurrido) no suma horas perdidas de flota, "
            "frente a un lunes 24 que sí tuvo huecos. Lectura: recuperación en el arranque de S35, "
            "con la salvedad de que el martes aún no cierra."
        )
    elif tot_24 == 0 and tot_25 > 0:
        doc.add_paragraph(
            "El lunes 24 cierra sin horas perdidas de flota y el martes 25 (parcial) sí acumula huecos. "
            "Hay que re-chequear al cierre del día para confirmar si es un incidente nuevo o un atraso de carga."
        )
    elif tot_25 > tot_24:
        doc.add_paragraph(
            f"El martes 25 ya supera al lunes 24 en horas perdidas de flota "
            f"({_fmt_int(tot_25)} vs {_fmt_int(tot_24)}), con menos horas exigidas. "
            "El arranque de S35 empeora respecto del lunes."
        )
    else:
        doc.add_paragraph(
            f"El lunes 24 concentra más horas perdidas ({_fmt_int(tot_24)}) que el tramo "
            f"transcurrido del martes 25 ({_fmt_int(tot_25)}). El martes, hasta el corte, "
            "no agrava el lunes; falta el resto del día."
        )

    doc.add_heading("6. Conclusión operativa", 1)
    if desc:
        lista_d = "; ".join(
            f"{r.node_name} ({r.node_id}, {_fmt_int(r.perdidas_total())} h, {r.etiqueta_estado()})"
            for r in desc[:5]
        )
        prio = doc.add_paragraph()
        prio.add_run("Prioridad 1 — desconectados. ").bold = True
        prio.add_run(
            f"{len(desc)} puntos caídos ahora. Atacar por horas perdidas (mayor a menor): "
            + lista_d
            + "."
        )
    if huec:
        lista_h = "; ".join(
            f"{r.node_name} ({r.node_id}, {_fmt_int(r.perdidas_total())} h)"
            for r in huec[:5]
        )
        prio2 = doc.add_paragraph()
        prio2.add_run("Prioridad 2 — conectados con huecos. ").bold = True
        prio2.add_run(
            "No están desconectados ahora; revisar por qué no hay serie completa "
            "(histórico no recuperado o corte ya cerrado). Primeros: "
            + lista_h
            + "."
        )
    doc.add_paragraph(
        "S35 aún no es comparable con S32–S34 en magnitud. Usar el 24 y el 25 como "
        "control de continuidad: si un punto del ranking de las semanas cerradas sigue "
        "sin serie el lunes/martes, el incidente no se cerró."
    )
    nota = doc.add_paragraph(
        "Nota: exclusiones iguales al reporte de puntos en cero (empresas/nodos fuera de "
        "operación o dados de baja). Fuente: GET /nodes/{id}/dates.measures.csv."
    )
    nota.runs[0].font.size = Pt(8)
    nota.runs[0].italic = True
    nota.runs[0].font.color.rgb = RGBColor(120, 120, 120)

    stamp = ahora_chile.strftime("%Y%m%d_%H%M")
    out = output_dir / f"Reporte_Horas_Perdidas_S32_S35_{stamp}.docx"
    doc.save(str(out))
    return out


def serializar(resultados: Iterable[ResultadoNodo]) -> List[Dict]:
    out = []
    for r in resultados:
        out.append(
            {
                "nodeId": r.node_id,
                "nodeName": r.node_name,
                "companyId": r.company_id,
                "companyName": r.company_name,
                "perdidasPorDia": r.perdidas_por_dia,
                "esperadasPorDia": r.esperadas_por_dia,
                "s32": r.perdidas_semana(32),
                "s33": r.perdidas_semana(33),
                "s34": r.perdidas_semana(34),
                "s35": r.perdidas_semana(35),
                "total": r.perdidas_total(),
                "lastUpdate": r.last_update,
                "wesStatus": r.wes_status,
                "desconectado": r.desconectado,
                "horasSinConexion": (
                    round(r.horas_sin_conexion, 1) if r.horas_sin_conexion is not None else None
                ),
                "estado": r.etiqueta_estado() if r.perdidas_total() > 0 else "ok",
                "error": r.error,
            }
        )
    return out


def main() -> int:
    if sys.platform == "win32":
        for stream in (sys.stdout, sys.stderr):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

    parser = argparse.ArgumentParser(description="Reporte semanal de horas perdidas")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Carpeta de salida (default: reports/Puntos_En_Cero)",
    )
    parser.add_argument("--top", type=int, default=25, help="Puntos en gráficos de ranking")
    parser.add_argument("--max-nodos", type=int, default=0, help="Limitar nodos (prueba)")
    parser.add_argument("--workers", type=int, default=MAX_WORKERS)
    parser.add_argument(
        "--desde-json",
        type=Path,
        default=None,
        help="Reusar horas ya calculadas (solo refresca estado de conexión)",
    )
    args = parser.parse_args()

    workers = max(4, int(args.workers))
    ahora = datetime.now(CHILE_TZ)
    output_dir = args.output_dir
    if output_dir is None:
        output_dir = Path(wes_scripts_root()) / "reports" / "Puntos_En_Cero"
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    charts_dir = output_dir / "_charts_horas_perdidas"
    charts_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 70)
    print("REPORTE DE HORAS PERDIDAS — S32 / S33 / S34 / S35 (2 días)")
    print(f"Corte Chile: {ahora.strftime('%d-%m-%Y %H:%M')}")
    print("=" * 70)

    if args.desde_json:
        json_in = args.desde_json.resolve()
        if not json_in.is_file():
            print(f"[ERROR] No existe --desde-json: {json_in}")
            return 1
        print(f"[INFO] Reusando horas desde {json_in}")
        resultados = cargar_desde_json(json_in)
    else:
        nodos = obtener_todos_los_nodos()
        if args.max_nodos and args.max_nodos > 0:
            nodos = nodos[: args.max_nodos]
            print(f"[INFO] --max-nodos={args.max_nodos}")
        if not nodos:
            print("[ERROR] No hay nodos para analizar.")
            return 1

        print(f"[INFO] Analizando {len(nodos)} puntos con {workers} workers...")
        resultados = []
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = {ex.submit(analizar_nodo, n, ahora): n for n in nodos}
            done = 0
            for fut in as_completed(futs):
                resultados.append(fut.result())
                done += 1
                if done % 20 == 0 or done == len(nodos):
                    print(f"  ... {done}/{len(nodos)} puntos")

    resultados.sort(key=lambda r: r.perdidas_total(), reverse=True)
    con = [r for r in resultados if r.perdidas_total() > 0]
    print(f"[INFO] Puntos con horas perdidas: {len(con)} / {len(resultados)}")
    print("[INFO] Clasificando desconexión (lastUpdate app)...")
    enriquecer_conexion(con, ahora)
    desc, huec = particionar_perdidas(con)
    print(
        f"[INFO] Desconectados: {len(desc)} ({sum(r.perdidas_total() for r in desc)} h) · "
        f"Conectados con huecos: {len(huec)} ({sum(r.perdidas_total() for r in huec)} h)"
    )

    top = max(5, int(args.top))
    charts = {
        "totales": charts_dir / "totales_por_semana.png",
        "ranking_desconexion": charts_dir / "ranking_desconexion.png",
        "ranking_huecos": charts_dir / "ranking_huecos.png",
        "s35_dias": charts_dir / "s35_24_vs_25.png",
    }
    for num in (32, 33, 34, 35):
        charts[f"s{num}_desc"] = charts_dir / f"ranking_s{num}_desconexion.png"
        charts[f"s{num}_huecos"] = charts_dir / f"ranking_s{num}_huecos.png"
    esp_sem = {n: sum(r.esperadas_semana(n) for r in resultados) for n, *_ in SEMANA_DEFS}
    grafico_totales_semana(resultados, esp_sem, charts["totales"])
    grafico_ranking(
        desc,
        lambda r: r.perdidas_total(),
        "Ranking — horas perdidas por desconexión",
        charts["ranking_desconexion"],
        COLOR_DESC,
        top,
        leyenda_estado=False,
    )
    grafico_ranking(
        huec,
        lambda r: r.perdidas_total(),
        "Ranking — conectados con huecos (no desconectados)",
        charts["ranking_huecos"],
        COLOR_HUECOS,
        top,
        leyenda_estado=False,
    )
    labels_sem = {
        32: "Semana 32 (03–09 ago)",
        33: "Semana 33 (10–16 ago)",
        34: "Semana 34 (17–23 ago)",
        35: "Semana 35 (24–25 ago, 2 días)",
    }
    for num in (32, 33, 34, 35):
        orden = sorted(resultados, key=lambda r: r.perdidas_semana(num), reverse=True)
        grafico_ranking(
            [r for r in orden if r.desconectado],
            lambda r, n=num: r.perdidas_semana(n),
            f"{labels_sem[num]} — desconexión",
            charts[f"s{num}_desc"],
            COLOR_DESC,
            top,
            leyenda_estado=False,
        )
        grafico_ranking(
            [r for r in orden if not r.desconectado],
            lambda r, n=num: r.perdidas_semana(n),
            f"{labels_sem[num]} — conectados con huecos",
            charts[f"s{num}_huecos"],
            COLOR_HUECOS,
            top,
            leyenda_estado=False,
        )
    grafico_semana35_dias(
        sorted(resultados, key=lambda r: r.perdidas_semana(35), reverse=True),
        charts["s35_dias"],
        min(top, 15),
    )

    json_path = output_dir / "horas_perdidas_s32_s35.json"
    json_path.write_text(
        json.dumps(serializar(resultados), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"[OK] JSON: {json_path}")

    docx_path = crear_reporte_word(resultados, ahora, output_dir, charts, top)
    print(f"[OK] DOCX: {docx_path}")
    pdf_path = convertir_a_pdf(docx_path)
    if pdf_path:
        print(f"[OK] PDF: {pdf_path}")

    print()
    print("Ranking desconexión:")
    for i, r in enumerate(desc[:10], start=1):
        print(
            f"  {i:2d}. {r.node_id}  {r.node_name}  T={r.perdidas_total()}  "
            f"{r.etiqueta_estado()}  lu={r.last_update}"
        )
    print("Ranking conectados con huecos:")
    for i, r in enumerate(huec[:10], start=1):
        print(
            f"  {i:2d}. {r.node_id}  {r.node_name}  T={r.perdidas_total()}  lu={r.last_update}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
