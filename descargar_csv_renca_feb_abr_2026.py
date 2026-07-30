from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path
from typing import Dict, List

import pandas as pd
import requests

from descargar_csv_medidas_wes import _fusionar_csv_bloques, _get_csv_bloques_utc_dia_chile
from generar_reporte_word import _requests_session

ENTITY_BASE = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
COMPANY_ID = "000017"  # Renca
DESDE = date(2026, 2, 1)
HASTA = date(2026, 4, 30)


def obtener_nodos_renca() -> List[Dict[str, str]]:
    r = requests.get(f"{ENTITY_BASE}/companies/{COMPANY_ID}", timeout=30)
    r.raise_for_status()
    nodes = r.json().get("nodes") or []
    out: List[Dict[str, str]] = []
    for n in nodes:
        nid = str(n.get("nodeId", "")).strip()
        name = str(n.get("name", "")).strip()
        if nid and name:
            out.append({"nodeId": nid, "nodeName": name})
    out.sort(key=lambda x: x["nodeId"])
    return out


def iter_days(d0: date, d1: date):
    d = d0
    while d <= d1:
        yield d
        d += timedelta(days=1)


def sanitize_sheet_name(name: str) -> str:
    bad = "[]:*?/\\"
    out = "".join("_" if c in bad else c for c in name).strip()
    if not out:
        out = "sheet"
    return out[:31]


def main() -> int:
    root = Path(__file__).resolve().parent
    out_dir = root / "reports" / "Renca" / "csv_puntos_feb_abr_2026"
    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx_path = out_dir / "renca_feb_abr_2026_por_nodo.xlsx"

    nodos = obtener_nodos_renca()
    if not nodos:
        print("[ERROR] No se encontraron nodos para Renca.")
        return 1

    sess = _requests_session()
    escritos = 0
    with pd.ExcelWriter(xlsx_path) as writer:
        for i, n in enumerate(nodos, start=1):
            nid = n["nodeId"]
            nname = n["nodeName"]
            print(f"[{i}/{len(nodos)}] {nid} - {nname}", flush=True)
            folder = out_dir / f"{nid}_{nname.replace(' ', '_')}"
            folder.mkdir(parents=True, exist_ok=True)

            dfs: List[pd.DataFrame] = []
            for d in iter_days(DESDE, HASTA):
                try:
                    bloques = _get_csv_bloques_utc_dia_chile(nid, d, session=sess)
                    texto = _fusionar_csv_bloques(bloques)
                    day_file = folder / f"{d.isoformat()}.csv"
                    day_file.write_text(texto, encoding="utf-8")
                    escritos += 1

                    df_day = pd.read_csv(day_file)
                    df_day.insert(0, "dia_chile", d.isoformat())
                    dfs.append(df_day)
                except Exception:
                    continue

            if dfs:
                df_node = pd.concat(dfs, ignore_index=True)
            else:
                df_node = pd.DataFrame(columns=["dia_chile", "TIME", "VALUE"])
            sheet = sanitize_sheet_name(f"{nid}_{nname}")
            df_node.to_excel(writer, sheet_name=sheet, index=False)

    print(f"[OK] Excel por hojas: {xlsx_path}")
    print(f"[OK] CSV diarios escritos: {escritos}")
    print(f"[OK] Carpeta base: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
