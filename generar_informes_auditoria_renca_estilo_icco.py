"""
Genera reportes de auditoría en el mismo formato del borrador ICCO
para todos los puntos de Renca (companies 000016 y 000017),
usando las mismas fechas ICCO:

- Con control: 23-03-2026 a 26-03-2026
- Sin control: 06-04-2026 a 09-04-2026

Exclusiones solicitadas:
- Juana Atala de Hirmas
- Rebeca Matte Bello (incluye el nodo de 000016-01 si aparece)

Salida (Word + PDF por punto):
  reports/reporte de auditoria/auditorias_renca_estilo_icco_20260323_20260409/

Uso:
  python generar_informes_auditoria_renca_estilo_icco.py
"""

from __future__ import annotations

import sys
from pathlib import Path

from generar_reporte_word import ENTITY_BASE_URL, convertir_word_a_pdf, fetch_json, get_node_name

import generar_informe_auditoria_icco_renca_word as icco


ROOT = Path(__file__).resolve().parent
OUT_DIR = (
    ROOT
    / "reports"
    / "reporte de auditoria"
    / "auditorias_renca_estilo_icco_20260323_20260409"
)

COMPANIES_RENCA = ("000016", "000017")

# Exclusiones por nombre y/o id (lo que el usuario pidió).
EXCLUDE_NAMES = {
    "juana atala de hirmas",
    "rebeca matte bello",
    "scl rebeca matte bello",
}
EXCLUDE_NODE_IDS = {
    "000017-01",
    "000017-02",
    "000016-01",
}


def _safe_filename(name: str) -> str:
    s = "".join(ch for ch in name if ch.isalnum() or ch in (" ", "-", "_")).strip()
    return (s.replace(" ", "_") or "punto")


def _company_nodes(company_id: str) -> list[dict]:
    data = fetch_json(f"{ENTITY_BASE_URL}/companies/{company_id}")
    nodes = []
    if isinstance(data, dict) and isinstance(data.get("nodes"), list):
        nodes = data["nodes"]
    if not nodes:
        data2 = fetch_json(f"{ENTITY_BASE_URL}/companies/{company_id}/nodes")
        if isinstance(data2, list):
            nodes = data2
        elif isinstance(data2, dict) and isinstance(data2.get("nodes"), list):
            nodes = data2["nodes"]
    out = []
    for n in nodes or []:
        node_id = (n.get("nodeId") or n.get("node_id") or "").strip()
        name = (n.get("name") or n.get("nodeName") or "").strip()
        if node_id:
            out.append({"nodeId": node_id, "name": name})
    return out


def main() -> int:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    # List nodes
    nodes: list[tuple[str, str]] = []
    for cid in COMPANIES_RENCA:
        for n in _company_nodes(cid):
            nid = n["nodeId"]
            nm = (n.get("name") or "").strip() or get_node_name(nid)
            nodes.append((nid, nm))

    # de-dup
    seen: set[str] = set()
    uniq: list[tuple[str, str]] = []
    for nid, nm in nodes:
        if nid in seen:
            continue
        seen.add(nid)
        uniq.append((nid, nm))

    ok = 0
    fail = 0
    skip = 0

    for i, (nid, nm) in enumerate(sorted(uniq), start=1):
        nm_l = (nm or "").strip().lower()
        if nid in EXCLUDE_NODE_IDS or nm_l in EXCLUDE_NAMES:
            print(f"[{i}/{len(uniq)}] {nid} — {nm} ... [SKIP] excluido")
            skip += 1
            continue

        # Patch portada + etiquetas del módulo ICCO para este punto
        icco.PORTADA_ESTABLECIMIENTO_LINEA1 = (nm or "Renca").strip()
        icco.PORTADA_ESTABLECIMIENTO_LINEA2 = "Renca"
        icco.PORTADA_REFERENCIA_BORRADOR = f"Borrador_auditoria_{_safe_filename(nm)}_{nid}_2026"
        icco._ETIQUETA_COLEGIO_CUADRO_RESUMEN = "RENCA"
        icco.NOMBRE_PUNTO = (nm or nid).strip()

        out_docx = OUT_DIR / f"Auditoria_Renca_{_safe_filename(nm)}_{nid}.docx"
        try:
            print(f"[{i}/{len(uniq)}] {nid} — {nm} ...", end=" ", flush=True)
            p = icco.generar_informe_word(
                node_id=nid,
                out_dir=OUT_DIR,
                output_docx=out_docx,
                mantener_borrador_manual=False,
                solo_consolidado=False,
            )
            pdf = convertir_word_a_pdf(p)
            ok += 1
            print(f"[OK] {p.name}" + (f" + {pdf.name}" if pdf else " (sin PDF)"))
        except Exception as e:
            fail += 1
            print(f"[ERROR] {e}")

    print("-" * 72)
    print(f"Completado. OK: {ok} | SKIP: {skip} | ERROR: {fail}")
    return 0 if fail == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())

