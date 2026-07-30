"""
Reconstruye ``checkpoint_consolidado_m3_mensual_<year>.json`` desde archivos::

  <dir>/<node_id>/<AAAA-MM>_dates.measures.csv

generados por ``descargar_csv_mensual_puente_alto.py``. La suma mensual es la misma
que el script de consolidado **sin** reconciliar con ``totalM3`` JSON (solo serie CSV).

Uso::
  python reconstruir_checkpoint_desde_csv_mensual_pa.py --year 2025

Luego::
  python generar_consolidado_m3_mensual_puente_alto.py --solo-exportar-excel --year 2025
"""
from __future__ import annotations

import argparse
import calendar
import json
from datetime import date
from pathlib import Path
from typing import Dict

from generar_reporte_word import _chile_hours_from_dates_measures_csv_text
from reporte_puente_alto_lxm import obtener_nodos_puente_alto

ROOT = Path(__file__).resolve().parent
CSV_DEFAULT = ROOT / "reports" / "proyeccion ahorre puente 2025" / "csv_mensual_por_nodo"
OUT_DEFAULT = ROOT / "reports" / "proyeccion ahorre puente 2025"


def _consumo_mes_desde_texto_csv(text: str, year: int, month: int) -> tuple[float, int]:
    last_d = calendar.monthrange(year, month)[1]
    total = 0.0
    dias_con = 0
    for day in range(1, last_d + 1):
        d = date(year, month, day)
        horas = _chile_hours_from_dates_measures_csv_text(text, d)
        s = sum(float(horas.get(h, 0.0)) for h in range(24))
        if s > 1e-9:
            dias_con += 1
        total += s
    return round(total, 4), dias_con


def main() -> int:
    ap = argparse.ArgumentParser(description="Checkpoint desde CSV mensual guardado (Puente Alto).")
    ap.add_argument("--year", type=int, default=2025)
    ap.add_argument("--mes-desde", type=int, default=1)
    ap.add_argument("--mes-fin", type=int, default=12)
    ap.add_argument("--csv-dir", type=Path, default=CSV_DEFAULT)
    ap.add_argument("--out-dir", type=Path, default=OUT_DEFAULT)
    args = ap.parse_args()

    year = args.year
    m0 = max(1, min(12, args.mes_desde))
    m1 = max(1, min(12, args.mes_fin))
    if m0 > m1:
        print("[ERROR] mes-desde > mes-fin")
        return 1

    base = Path(args.csv_dir).expanduser().resolve()
    out_dir = Path(args.out_dir).expanduser().resolve()
    if not base.is_dir():
        print(f"[ERROR] No existe carpeta CSV: {base}")
        return 1

    nodos = obtener_nodos_puente_alto()
    nodos.sort(key=lambda x: x["nodeId"])
    checkpoint: Dict[str, Dict] = {}

    for n in nodos:
        nid = n["nodeId"]
        nombre = n["nodeName"]
        dir_n = base / nid
        if not dir_n.is_dir():
            print(f"[WARN] Sin carpeta {dir_n}; se omite {nid}")
            continue

        ck: Dict = {"node_id": nid, "colegio": nombre}
        total_anio = 0.0

        for mes in range(m0, m1 + 1):
            fn = dir_n / f"{year}-{mes:02d}_dates.measures.csv"
            mk = f"{year}-{mes:02d}"
            if not fn.is_file():
                ck[mk] = 0.0
                ck[f"{mk}_dias_con_dato"] = 0
                ck[f"{mk}_metodo"] = "csv_local_ausente"
                continue
            text = fn.read_text(encoding="utf-8", errors="replace")
            if not text.strip():
                ck[mk] = 0.0
                ck[f"{mk}_dias_con_dato"] = 0
                ck[f"{mk}_metodo"] = "csv_local_vacio"
                continue
            m3, dias_con = _consumo_mes_desde_texto_csv(text, year, mes)
            ck[mk] = m3
            ck[f"{mk}_dias_con_dato"] = dias_con
            ck[f"{mk}_metodo"] = "csv_local_guardado"
            total_anio += m3

        ck["total_anio_m3"] = round(total_anio, 4)
        ck["completo"] = True
        ck["_meta_year"] = year
        ck["_meta_mes_desde"] = m0
        ck["_meta_mes_fin"] = m1
        checkpoint[nid] = ck

    ck_path = out_dir / f"checkpoint_consolidado_m3_mensual_{year}.json"
    out_dir.mkdir(parents=True, exist_ok=True)
    ck_path.write_text(json.dumps(checkpoint, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OK] Checkpoint ({len(checkpoint)} nodos): {ck_path}")
    print("[INFO] Totales desde CSV guardado (sin totalM3 JSON). Para alinear con app, ejecute descarga API completa.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
