"""
Solo actualiza el gráfico de proyección/cuenta (barras de mayor a menor).
No mueve el cuadro PROYECCIÓN / CUENTA MENSUAL ni el resto del Word.
"""

from __future__ import annotations

import io
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.oxml.ns import qn

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from generar_reporte_word import format_number_chilean

DOC_PATH = Path(
    "reports/Bupa_Antofagasta/ABREGADO/AGREGADO_20260728_1725/"
    "Reporte_Agregado_BUPA_20260723_20260728.docx"
)
CHART_PATH = DOC_PATH.parent / "chart_proyeccion_mensual_total_por_punto.png"

# Valores mensuales actuales (factura / proyección WES)
ROWS = [
    {"nombre": "Medidor Principal Sanitaria", "m3": 6696.0},
    {"nombre": "Sala de Bomba Principal", "m3": 1139.5},
    {"nombre": "Sala de Bomba Sexto Piso", "m3": 180.5},
    {"nombre": "Sala de Bomba N°2", "m3": 31.0},
]


def _para_text(elem) -> str:
    return "".join(t.text or "" for t in elem.iter(qn("w:t"))).strip()


def _build_chart(path: Path) -> Path:
    ordered = sorted(ROWS, key=lambda r: r["m3"], reverse=True)
    nombres = [r["nombre"] for r in ordered]
    valores = [r["m3"] for r in ordered]
    fig, ax = plt.subplots(figsize=(8, 4.8))
    bars = ax.bar(nombres, valores, color="#0050b3", alpha=0.85, edgecolor="#003a80", linewidth=1.1)
    ax.set_ylabel("m³ / mes", fontsize=12, fontweight="bold")
    ax.set_title("Proyección / cuenta mensual por punto", fontsize=14, fontweight="bold")
    ax.set_ylim(bottom=0)
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=25, ha="right", fontsize=10)
    for bar, val in zip(bars, valores):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height(),
            f"{format_number_chilean(val, 1)} m³",
            ha="center",
            va="bottom",
            fontsize=10,
            fontweight="bold",
        )
    plt.tight_layout()
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def _replace_embedded_chart(doc: Document, chart_bytes: bytes) -> str:
    """Reemplaza la imagen que sigue al título/tabla de proyección, sin mover elementos."""
    body = doc.element.body
    children = list(body)
    target_rid = None
    for i, child in enumerate(children):
        if child.tag.split("}")[-1] != "p":
            continue
        t = _para_text(child).upper()
        if "PROYECCIÓN / CUENTA MENSUAL" in t or "PROYECCION / CUENTA MENSUAL" in t:
            # buscar la primera imagen después de este título (tabla + img)
            for nxt in children[i + 1 :]:
                blips = list(nxt.iter(qn("a:blip")))
                if blips:
                    target_rid = blips[0].get(qn("r:embed"))
                    break
                # detenerse si empieza otra sección
                nt = _para_text(nxt).upper()
                if nt.startswith("PARTICIPACI"):
                    break
            break

    if not target_rid:
        # fallback: imagen cuyo partname/contenido es el chart de proyección
        for rel in doc.part.rels.values():
            if "image" not in rel.reltype:
                continue
            # comparar con archivo actual en disco por tamaño aproximado no es fiable;
            # usar la imagen justo antes de PARTICIPACIÓN
        raise RuntimeError("No se encontró la imagen del gráfico de proyección")

    if target_rid not in doc.part.rels:
        raise RuntimeError(f"Relación {target_rid} no existe")

    part = doc.part.rels[target_rid].target_part
    part._blob = chart_bytes
    return target_rid


def main() -> None:
    _build_chart(CHART_PATH)
    chart_bytes = CHART_PATH.read_bytes()
    doc = Document(str(DOC_PATH))
    rid = _replace_embedded_chart(doc, chart_bytes)
    try:
        doc.save(str(DOC_PATH))
        out = DOC_PATH
    except PermissionError:
        out = DOC_PATH.with_name(DOC_PATH.stem + "_ordenado.docx")
        doc.save(str(out))
        print(f"[ADVERTENCIA] Guardado como {out.name}")
    print("[OK] Gráfico ordenado mayor→menor; cuadro no movido")
    print(f"     Imagen actualizada: {rid}")
    for r in sorted(ROWS, key=lambda x: x["m3"], reverse=True):
        print(f"       {r['nombre']}: {format_number_chilean(r['m3'], 1)} m³")


if __name__ == "__main__":
    main()
