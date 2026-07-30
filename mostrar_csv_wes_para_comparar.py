"""
Muestra las URLs exactas del CSV (y una muestra del texto) que usa la logica del informe/Excel.

No necesitas inspeccionar la app: ejecuta este script y compara los numeros con lo que ves en la app.

Ejemplos (PowerShell):
  python mostrar_csv_wes_para_comparar.py
  python mostrar_csv_wes_para_comparar.py --node-id 000017-08 --dia 2026-03-25
"""
from __future__ import annotations

import argparse
from datetime import date, datetime, timezone

from generar_reporte_word import (
    _dt_to_chile,
    _requests_session,
    _utc_calendar_dates_for_chile_day,
    acl_node_base_url,
)


def main() -> int:
    p = argparse.ArgumentParser(
        description="Imprime URLs GET del CSV WES y muestra TIME,VALUE (misma fusion que el Excel)."
    )
    p.add_argument("--node-id", default="000017-08", help="ID del nodo WES")
    p.add_argument(
        "--dia",
        type=lambda s: date.fromisoformat(s),
        default=date(2026, 3, 25),
        help="Dia civil Chile (AAAA-MM-DD)",
    )
    args = p.parse_args()

    node_id = args.node_id
    dia: date = args.dia
    url_base = f"{acl_node_base_url()}/nodes/{node_id}/dates.measures.csv"

    print("=" * 72)
    print("1) Que copiar si alguien te pide 'la URL del CSV'")
    print("   Es un GET a este path con start y end en formato ddMMyyyy (mismo dia en ambos).")
    print()
    uds = _utc_calendar_dates_for_chile_day(dia)
    print(f"   Para el dia Chile {dia} el informe pide UNO o DOS archivos (dias UTC que cubren ese dia):")
    for ud in uds:
        ddmmyyyy = ud.strftime("%d%m%Y")
        full = f"{url_base}?start={ddmmyyyy}&end={ddmmyyyy}"
        print(f"   - {full}")
    print()
    print("2) Puedes pegar esa URL en el navegador (Chrome/Edge): deberias ver texto TIME,VALUE")
    print()
    print("3) Valores por hora CHILE al fusionar solo el CSV de la API (serie cruda TIME,VALUE):")
    print("-" * 72)

    acc: dict[int, float] = {}
    sess = _requests_session()
    for ud in uds:
        ddmmyyyy = ud.strftime("%d%m%Y")
        r = sess.get(url_base, params=[("start", ddmmyyyy), ("end", ddmmyyyy)], timeout=60)
        r.raise_for_status()
        for line in r.text.strip().split("\n")[1:]:
            if not line.strip():
                continue
            parts = line.split(",", 1)
            if len(parts) < 2:
                continue
            try:
                ts = parts[0].strip().replace("Z", "+00:00")
                v = float(parts[1].strip().replace(" ", "").replace(",", "."))
                dt = datetime.fromisoformat(ts)
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=timezone.utc)
                ch = _dt_to_chile(dt)
                if ch.date() != dia:
                    continue
                acc[int(ch.hour)] = v
            except (ValueError, TypeError, IndexError):
                continue

    for h in range(24):
        v = acc.get(h, 0.0)
        bar = "#" * min(40, int(v * 20 + 0.5)) if v > 0 else ""
        print(f"   {h:02d}:00  {v:8.3f} m3/h  {bar}")

    print("-" * 72)
    print(
        "4) Serie que usa el Excel / informes Word (get_hourly_measures_for_day): "
        "si el JSON tiene totalM3 pero measures vacio, es totalM3/24 en todas las horas "
        "(igual criterio que los totales diarios del JSON). Si no, sigue el CSV."
    )
    print("-" * 72)
    from generar_reporte_word import get_hourly_measures_for_day

    t = datetime.combine(dia, datetime.min.time())
    gh = get_hourly_measures_for_day(node_id, t) or []
    gdict = {int(h): float(v) for h, v in gh}
    for h in range(24):
        v = gdict.get(h, 0.0)
        bar = "#" * min(40, int(v * 20 + 0.5)) if v > 0 else ""
        print(f"   {h:02d}:00  {v:8.3f} m3/h  {bar}")
    print("-" * 72)
    print("Si necesitas la tabla del apartado 3 en el Excel:  set WES_HOURLY_SIN_MEASURES=csv")
    print("Si en la app ves otros numeros para las mismas horas:")
    print("  - Comprueba que sea el mismo nodo y el mismo dia civil Chile.")
    print("  - Prueba el mismo host que la app:  set WES_API_BASE_URL=https://...")
    print("  - O guarda el CSV de la app como  AAAA-MM-DD.csv  y:  set WES_MEDIDAS_CSV_DIR=ruta\\carpeta")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
