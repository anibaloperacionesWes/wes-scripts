#!/usr/bin/env python3
"""
Exporta el árbol de equipos WES a Word para revisión en Office / Drive.

Uso:
  python exportar_arbol_equipos_word.py
  python exportar_arbol_equipos_word.py --json arbol_equipos_wes.json
  python exportar_arbol_equipos_word.py --regen   # regenera JSON desde API primero
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

ROOT = Path(__file__).resolve().parent
DEFAULT_JSON = ROOT / "arbol_equipos_wes.json"
OUT_DIR = ROOT / "reports" / "Arbol_Equipos"

# Empresas fuera del Dashboard operativo (mismas que exclusiones + internos)
FUERA_IDS = {
    "000000",
    "000001",
    "000004",
    "000005",
    "000011",
    "000014",
    "000018",
    "000019",
    "000023",
}

WES_BLUE = RGBColor(0x00, 0x5C, 0xA8)
WES_DARK = RGBColor(0x1A, 0x1A, 0x2E)
WES_GRAY = RGBColor(0x55, 0x55, 0x55)
WES_GREEN = RGBColor(0x1B, 0x7A, 0x3D)
WES_ORANGE = RGBColor(0xC4, 0x5C, 0x00)
WES_RED = RGBColor(0xA3, 0x1A, 0x1A)


def _set_run(run, *, size: int = 11, bold: bool = False, color: RGBColor = WES_DARK) -> None:
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Calibri"
    r = run._element.get_or_add_rPr()
    rFonts = r.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), "Calibri")


def _add_para(
    doc: Document,
    text: str,
    *,
    size: int = 11,
    bold: bool = False,
    color: RGBColor = WES_DARK,
    space_after: int = 4,
    space_before: int = 0,
) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    run = p.add_run(text)
    _set_run(run, size=size, bold=bold, color=color)


def _topo_label(topo: Optional[str]) -> str:
    if topo == "red_con_subredes":
        return "Red con subredes"
    if topo == "puntos_separados":
        return "Puntos separados"
    return topo or "—"


def _estado_label(estado: Optional[str]) -> str:
    return {
        "activo": "Activo",
        "pendiente": "Pendiente (sin instalación)",
        "fuera": "Fuera / excluido",
    }.get(estado or "", estado or "Activo")


def _count_leaves(node: Dict[str, Any]) -> int:
    kids = node.get("children") or []
    if not kids and node.get("nodeId"):
        return 1
    return sum(_count_leaves(k) for k in kids)


def _indent_prefix(depth: int) -> str:
    if depth <= 0:
        return ""
    return "    " * (depth - 1) + "└── "


def _walk_lines(node: Dict[str, Any], depth: int = 0) -> List[tuple]:
    """Lista de (depth, texto, estilo) para el árbol."""
    tipo = node.get("tipo", "")
    name = node.get("name") or node.get("nodeId") or "?"
    nid = node.get("nodeId") or ""
    estado = node.get("estado_operativo")
    topo = node.get("topologia")

    if tipo == "cliente":
        label = f"{name}"
        if node.get("nombre_api") and node["nombre_api"] != name:
            label += f"  (API: {node['nombre_api']})"
        label += f"  [{node.get('companyId', '')}]"
        style = "cliente"
    elif tipo == "sitio":
        label = f"Sitio: {name}"
        if estado:
            label += f"  · {_estado_label(estado)}"
        if topo:
            label += f"  · {_topo_label(topo)}"
        style = "sitio"
    elif tipo == "red":
        label = f"RED  {name}"
        if nid:
            label += f"  ({nid})"
        style = "red"
    elif tipo == "subred":
        label = f"Subred  {name}"
        if nid:
            label += f"  ({nid})"
        style = "subred"
    else:
        label = name
        if nid:
            label += f"  ({nid})"
        style = "punto"

    lines = [(depth, label, style, estado)]
    for ch in node.get("children") or []:
        lines.extend(_walk_lines(ch, depth + 1))
    return lines


def _style_color(style: str, estado: Optional[str]) -> RGBColor:
    if estado == "fuera":
        return WES_RED
    if estado == "pendiente":
        return WES_ORANGE
    return {
        "cliente": WES_BLUE,
        "sitio": WES_DARK,
        "red": WES_GREEN,
        "subred": WES_GREEN,
        "punto": WES_GRAY,
    }.get(style, WES_DARK)


def _add_tree_block(doc: Document, cliente: Dict[str, Any]) -> None:
    topo = cliente.get("topologia")
    estado = cliente.get("estado_operativo")
    n_hojas = _count_leaves(cliente)

    title = f"{cliente.get('name')}  ·  {cliente.get('companyId', '')}"
    _add_para(doc, title, size=14, bold=True, color=WES_BLUE, space_before=10, space_after=2)

    meta = f"Topología: {_topo_label(topo)}  |  Puntos: {n_hojas}  |  Estado: {_estado_label(estado)}"
    _add_para(doc, meta, size=10, color=WES_GRAY, space_after=2)

    if cliente.get("descripcion"):
        _add_para(doc, cliente["descripcion"], size=9, color=WES_GRAY, space_after=4)

    for depth, label, style, est in _walk_lines(cliente):
        if depth == 0:
            continue  # ya está el título del cliente
        prefix = _indent_prefix(depth)
        color = _style_color(style, est)
        bold = style in ("sitio", "red", "subred")
        size = 11 if style == "sitio" else 10
        _add_para(doc, prefix + label, size=size, bold=bold, color=color, space_after=1)


def build_document(tree: Dict[str, Any]) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Pt(56)
    section.bottom_margin = Pt(56)
    section.left_margin = Pt(64)
    section.right_margin = Pt(56)

    # Portada / resumen
    _add_para(doc, "WES — Árbol de equipos", size=22, bold=True, color=WES_BLUE, space_after=6)
    _add_para(
        doc,
        "Documento de revisión para el Dashboard",
        size=12,
        color=WES_GRAY,
        space_after=8,
    )
    gen = tree.get("generado", "")
    r = tree.get("resumen") or {}
    _add_para(
        doc,
        f"Generado: {gen}  ·  Clientes: {r.get('total_clientes', '—')}  ·  "
        f"Puntos hoja: {r.get('total_nodos_hoja', '—')}  ·  "
        f"Redes: {r.get('red_con_subredes', '—')}  ·  "
        f"Puntos separados: {r.get('puntos_separados', '—')}",
        size=10,
        color=WES_GRAY,
        space_after=12,
    )

    _add_para(doc, "Leyenda", size=12, bold=True, color=WES_DARK, space_after=4)
    _add_para(doc, "• Puntos separados — cada medidor es independiente", size=10, color=WES_GRAY)
    _add_para(
        doc,
        "• Red con subredes — matriz/entrada y puntos aguas abajo (no doble-contar)",
        size=10,
        color=WES_GRAY,
    )
    _add_para(doc, "• Activo / Pendiente / Fuera — estado operativo del sitio o cliente", size=10, color=WES_GRAY)
    _add_para(
        doc,
        "• Empresa API distinta del nombre Dashboard (ej. DERCO → Inchcape) se indica entre paréntesis",
        size=10,
        color=WES_GRAY,
        space_after=14,
    )

    clientes = tree.get("clientes") or []
    fuera = [c for c in clientes if c.get("companyId") in FUERA_IDS]
    activos = [c for c in clientes if c.get("companyId") not in FUERA_IDS]

    # Índice rápido
    _add_para(doc, "1. Índice — clientes en el árbol", size=14, bold=True, color=WES_BLUE, space_after=6)
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    for i, h in enumerate(["#", "ID", "Cliente", "Topología", "Puntos"]):
        hdr[i].text = h
        for p in hdr[i].paragraphs:
            for run in p.runs:
                _set_run(run, size=9, bold=True, color=WES_DARK)

    for i, c in enumerate(activos, 1):
        row = table.add_row().cells
        vals = [
            str(i),
            c.get("companyId", ""),
            c.get("name", ""),
            _topo_label(c.get("topologia")),
            str(_count_leaves(c)),
        ]
        for j, v in enumerate(vals):
            row[j].text = v
            for p in row[j].paragraphs:
                for run in p.runs:
                    _set_run(run, size=9, color=WES_DARK)

    doc.add_paragraph()
    if fuera:
        _add_para(
            doc,
            "Empresas que siguen en la API pero están fuera del Dashboard operativo "
            "(Ejército, Gendarmería, MOP, MADECO, Lucchetti, internos WES): "
            + ", ".join(f"{c.get('name')} ({c.get('companyId')})" for c in fuera),
            size=9,
            color=WES_RED,
            space_after=12,
        )

    # Detalle activos
    _add_para(
        doc,
        "2. Detalle por cliente (solo activos / parciales)",
        size=14,
        bold=True,
        color=WES_BLUE,
        space_before=8,
        space_after=8,
    )
    for c in activos:
        _add_tree_block(doc, c)

    # Anexo fuera
    if fuera:
        doc.add_page_break()
        _add_para(
            doc,
            "Anexo — empresas fuera del Dashboard",
            size=14,
            bold=True,
            color=WES_RED,
            space_after=8,
        )
        _add_para(
            doc,
            "Listadas solo como referencia. No deben navegarse como clientes del Dashboard.",
            size=10,
            color=WES_GRAY,
            space_after=8,
        )
        for c in fuera:
            _add_tree_block(doc, c)

    return doc


def main() -> int:
    if sys.platform == "win32":
        for s in (sys.stdout, sys.stderr):
            try:
                s.reconfigure(encoding="utf-8", line_buffering=True)
            except Exception:
                pass

    parser = argparse.ArgumentParser(description="Exportar árbol de equipos a Word")
    parser.add_argument("--json", type=Path, default=DEFAULT_JSON)
    parser.add_argument("--regen", action="store_true", help="Regenerar JSON desde API")
    parser.add_argument("--salida", type=Path, default=None)
    args = parser.parse_args()

    if args.regen:
        from arbol_equipos_wes import build_tree, export_json, print_summary

        print("Regenerando árbol desde API...")
        tree = build_tree()
        print_summary(tree)
        export_json(tree, args.json)
    else:
        if not args.json.is_file():
            print(f"[ERROR] No existe {args.json}. Usá --regen o generá el JSON primero.", file=sys.stderr)
            return 1
        import json

        tree = json.loads(args.json.read_text(encoding="utf-8"))

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out = args.salida or (OUT_DIR / f"Arbol_Equipos_WES_{stamp}.docx")
    out.parent.mkdir(parents=True, exist_ok=True)

    doc = build_document(tree)
    doc.save(str(out))
    print(f"[OK] Word: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
