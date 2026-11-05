"""
Exporta medidas de todos los puntos WES a CSV + Excel (una hoja por nodo).

Desde la nube el CSV horario de la app (dates.measures.csv) exige
``WES_X_AUTH_TOKEN``. Si no hay token, usa el JSON diario measures/dates
(fecha + totalM3), disponible sin login vía BFF.

Ejemplo:
  python exportar_csv_todos_puntos_excel.py
  python exportar_csv_todos_puntos_excel.py --desde 2025-01-01 --hasta 2026-11-05
"""

from __future__ import annotations

import argparse
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from wes_api_bff import listar_nodos_companias, measures_dates

# Empresas internas / fuera de operación cliente (mismo criterio base del repo).
EXCLUDE_COMPANY_IDS = {"000000", "000001", "000004", "100000"}


def _parse_day(s: str) -> date:
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise argparse.ArgumentTypeError(f"Fecha inválida: {s!r}")


def sanitize_sheet_name(name: str, used: Dict[str, int]) -> str:
    bad = "[]:*?/\\"
    out = "".join("_" if c in bad else c for c in name).strip() or "sheet"
    out = out[:31]
    if out not in used:
        used[out] = 0
        return out
    used[out] += 1
    suffix = f"_{used[out]}"
    return (out[: 31 - len(suffix)] + suffix)


def sanitize_filename(name: str) -> str:
    out = re.sub(r"[^\w.\-]+", "_", name, flags=re.UNICODE).strip("._")
    return (out or "nodo")[:120]


def chunked(items: List[str], size: int) -> List[List[str]]:
    return [items[i : i + size] for i in range(0, len(items), size)]


def fetch_daily_by_node(
    node_ids: List[str],
    desde: date,
    hasta: date,
    *,
    batch_size: int = 15,
) -> Dict[str, List[dict]]:
    """Devuelve nodeId -> filas {fecha, totalM3, ...}."""
    start = desde.strftime("%d%m%Y")
    end = hasta.strftime("%d%m%Y")
    by_node: Dict[str, List[dict]] = {nid: [] for nid in node_ids}
    batches = chunked(node_ids, batch_size)
    for i, batch in enumerate(batches, start=1):
        print(f"  lote {i}/{len(batches)} ({len(batch)} nodos)...", flush=True)
        try:
            payload = measures_dates(batch, start, end, timeout=180.0)
        except Exception as exc:
            print(f"  [WARN] lote falló, reintento 1x1: {exc}", flush=True)
            payload = []
            for nid in batch:
                try:
                    payload.extend(measures_dates([nid], start, end, timeout=120.0))
                except Exception as exc2:
                    print(f"  [WARN] {nid}: {exc2}", flush=True)
        for item in payload:
            nid = str(item.get("nodeId", "")).strip()
            if nid not in by_node:
                continue
            for m in item.get("month") or []:
                fecha = str(m.get("date", ""))[:10]
                if not fecha:
                    continue
                try:
                    d = date.fromisoformat(fecha)
                except ValueError:
                    continue
                if d < desde or d > hasta:
                    continue
                total = m.get("totalM3")
                by_node[nid].append(
                    {
                        "fecha": fecha,
                        "totalM3": float(total) if total is not None else None,
                        "expenses": (m.get("nodeKpi") or {}).get("expenses"),
                        "efficiency": (m.get("nodeKpi") or {}).get("efficiency"),
                    }
                )
    for nid, rows in by_node.items():
        rows.sort(key=lambda r: r["fecha"])
        # dedupe por fecha (última gana)
        seen: Dict[str, dict] = {}
        for r in rows:
            seen[r["fecha"]] = r
        by_node[nid] = [seen[k] for k in sorted(seen)]
    return by_node


def last_day_with_data(by_node: Dict[str, List[dict]], fallback: date) -> date:
    last: Optional[date] = None
    for rows in by_node.values():
        for r in reversed(rows):
            if r.get("totalM3") is None:
                continue
            try:
                d = date.fromisoformat(r["fecha"])
            except ValueError:
                continue
            if last is None or d > last:
                last = d
            break
    return last or fallback


def main() -> int:
    today = date.today()
    p = argparse.ArgumentParser(description="CSV + Excel por nodo (todos los puntos WES)")
    p.add_argument("--desde", type=_parse_day, default=date(2025, 1, 1))
    p.add_argument("--hasta", type=_parse_day, default=today)
    p.add_argument(
        "--out-dir",
        type=Path,
        default=None,
        help="Carpeta de salida (default reports/Todos_Puntos/csv_diario_<rango>)",
    )
    p.add_argument("--batch-size", type=int, default=15)
    args = p.parse_args()

    desde: date = args.desde
    hasta: date = args.hasta
    if hasta < desde:
        raise SystemExit("--hasta debe ser >= --desde")

    root = Path(__file__).resolve().parent
    out_dir = args.out_dir
    if out_dir is None:
        out_dir = (
            root
            / "reports"
            / "Todos_Puntos"
            / f"csv_diario_{desde.isoformat()}_a_{hasta.isoformat()}"
        )
    out_dir = out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    csv_dir = out_dir / "csv_por_nodo"
    csv_dir.mkdir(parents=True, exist_ok=True)

    print("Listando nodos...", flush=True)
    nodos = listar_nodos_companias(exclude_company_ids=EXCLUDE_COMPANY_IDS)
    if not nodos:
        print("[ERROR] No se encontraron nodos.")
        return 1
    print(f"Nodos: {len(nodos)} (excluidas {sorted(EXCLUDE_COMPANY_IDS)})", flush=True)

    node_ids = [n["nodeId"] for n in nodos]
    print(f"Descargando medidas diarias {desde} → {hasta} ...", flush=True)
    by_node = fetch_daily_by_node(node_ids, desde, hasta, batch_size=args.batch_size)

    effective_end = last_day_with_data(by_node, hasta)
    if effective_end < hasta:
        print(
            f"[INFO] Último día con data global: {effective_end} (pedido hasta {hasta})",
            flush=True,
        )
        # recortar filas posteriores (por si API devolvió algo futuro vacío)
        for nid, rows in by_node.items():
            by_node[nid] = [r for r in rows if r["fecha"] <= effective_end.isoformat()]

    xlsx_path = out_dir / (
        f"todos_puntos_diario_{desde.isoformat()}_a_{effective_end.isoformat()}_por_nodo.xlsx"
    )
    used_sheets: Dict[str, int] = {}
    meta_rows = []
    escritos_csv = 0

    print(f"Escribiendo Excel: {xlsx_path.name}", flush=True)
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        # hoja índice
        for n in nodos:
            nid = n["nodeId"]
            rows = by_node.get(nid) or []
            df = pd.DataFrame(rows)
            if df.empty:
                df = pd.DataFrame(columns=["fecha", "totalM3", "expenses", "efficiency"])
            else:
                df.insert(0, "nodeName", n["nodeName"])
                df.insert(0, "nodeId", nid)
                df.insert(0, "companyName", n["companyName"])
                df.insert(0, "companyId", n["companyId"])

            sheet = sanitize_sheet_name(f"{nid}_{n['nodeName']}", used_sheets)
            df.to_excel(writer, sheet_name=sheet, index=False)

            csv_name = sanitize_filename(f"{nid}_{n['nodeName']}") + ".csv"
            csv_path = csv_dir / csv_name
            df.to_csv(csv_path, index=False, encoding="utf-8")
            escritos_csv += 1

            fechas = [r["fecha"] for r in rows]
            nz = [r for r in rows if (r.get("totalM3") or 0) > 0]
            meta_rows.append(
                {
                    "companyId": n["companyId"],
                    "companyName": n["companyName"],
                    "nodeId": nid,
                    "nodeName": n["nodeName"],
                    "sheet": sheet,
                    "filas": len(rows),
                    "dias_con_consumo": len(nz),
                    "primera_fecha": fechas[0] if fechas else "",
                    "ultima_fecha": fechas[-1] if fechas else "",
                    "suma_totalM3": round(sum(r.get("totalM3") or 0 for r in rows), 3),
                }
            )

        pd.DataFrame(meta_rows).to_excel(writer, sheet_name="00_indice", index=False)

    # reordenar: openpyxl pone 00_indice al final; mover al inicio
    try:
        from openpyxl import load_workbook

        wb = load_workbook(xlsx_path)
        if "00_indice" in wb.sheetnames:
            idx = wb.sheetnames.index("00_indice")
            wb.move_sheet(wb["00_indice"], offset=-idx)
            wb.save(xlsx_path)
    except Exception as exc:
        print(f"[WARN] No se pudo reordenar hojas: {exc}", flush=True)

    resumen = out_dir / "resumen_export.json"
    import json

    resumen.write_text(
        json.dumps(
            {
                "desde": desde.isoformat(),
                "hasta_pedido": hasta.isoformat(),
                "hasta_efectivo": effective_end.isoformat(),
                "nodos": len(nodos),
                "csv_escritos": escritos_csv,
                "excel": str(xlsx_path),
                "fuente": "measures/dates (diario totalM3 vía BFF)",
                "nota": (
                    "CSV horario de la app requiere WES_X_AUTH_TOKEN; "
                    "este export usa totales diarios disponibles sin login."
                ),
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    print(f"[OK] Excel: {xlsx_path}")
    print(f"[OK] CSV por nodo: {csv_dir} ({escritos_csv} archivos)")
    print(f"[OK] Rango efectivo: {desde} → {effective_end}")
    print(f"[OK] Resumen: {resumen}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
