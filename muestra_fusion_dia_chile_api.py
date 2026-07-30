"""
Muestra paso a paso cómo un día civil Chile se construye desde dates.measures.csv.
Ejecutar: python muestra_fusion_dia_chile_api.py
"""
from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timezone

from generar_reporte_word import (
    BASE_URL,
    _dt_to_chile,
    _requests_session,
    _utc_calendar_dates_for_chile_day,
    get_hourly_measures_for_day,
)

# Cambiar aquí el día y nodo de ejemplo (mismo default que auditoría ICCO)
NODE_ID = "000017-08"
DIA_CHILE = date(2026, 3, 24)


def main() -> None:
    print("=" * 78)
    print(f"  Dia civil CHILE: {DIA_CHILE:%d-%m-%Y}  |  nodo: {NODE_ID}")
    print("=" * 78)
    print()
    print("PASO 1 - La API entrega CSV por DIA UTC (no por dia Chile).")
    print("        Para este día Chile, se consultan estos calendarios UTC")
    print("        (cada uno: GET con start=end=ddmmyyyy):")
    uds = _utc_calendar_dates_for_chile_day(DIA_CHILE)
    for u in uds:
        print(f"          - {u:%d-%m-%Y} UTC  |  parametro start=end={u.strftime('%d%m%Y')}")
    print()
    print("PASO 2 - Cada CSV trae 24 filas: T00Z a T23Z de ese dia UTC, columna VALUE = m3/h.")
    print()

    sess = _requests_session()
    url = f"{BASE_URL}/nodes/{NODE_ID}/dates.measures.csv"
    filas: list[tuple[str, float, object, str, date]] = []

    for ud in uds:
        ds = ud.strftime("%d%m%Y")
        r = sess.get(url, params=[("start", ds), ("end", ds)], timeout=30)
        r.raise_for_status()
        for line in r.text.strip().split("\n")[1:]:
            if not line.strip():
                continue
            parts = line.split(",", 1)
            if len(parts) < 2:
                continue
            ts = parts[0].strip()
            v = float(parts[1].strip().replace(" ", "").replace(",", "."))
            dt = datetime.fromisoformat(ts.replace("Z", "+00:00"))
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            ch = _dt_to_chile(dt)
            inc = "si" if ch.date() == DIA_CHILE else "no"
            filas.append((ts, v, ch, inc, ud))

    print("PASO 3 - Cada TIME se pasa a hora Chile. Solo filas con fecha Chile = dia pedido entran en el dia.")
    print()
    hdr = f"{'TIME (UTC)':<30} {'m3/h':>8}   {'Hora Chile':<28} {'24-03 Chile?':>14}"
    print(hdr)
    print("-" * len(hdr))
    for ts, v, ch, inc, _ud in filas:
        print(f"{ts:<30} {v:8.2f}   {ch.strftime('%Y-%m-%d %H:%M %z'):<28} {inc:>14}")

    acc: dict[int, float] = defaultdict(float)
    for _ts, v, ch, inc, _ud in filas:
        if ch.date() != DIA_CHILE:
            continue
        hi = int(ch.hour)
        if 0 <= hi < 24:
            acc[hi] += v

    print()
    print("PASO 4 - Se agrupa por hora Chile 0-23 (si hay dos filas en la misma hora Chile, se suman).")
    print()
    print(f"{'Hora Chile':<14} {'m3/h':>10}")
    print("-" * 26)
    for h in range(24):
        print(f"{h:02d}:00-{(h+1)%24:02d}:00   {acc.get(h, 0.0):10.2f}")

    s = sum(acc.get(h, 0.0) for h in range(24))
    print("-" * 26)
    print(f"{'Suma 24 h (~ m3 dia)':<14} {s:10.2f}")
    print()

    target = datetime.combine(DIA_CHILE, datetime.min.time())
    api = dict(get_hourly_measures_for_day(NODE_ID, target) or [])
    coinciden = all(abs(acc.get(h, 0.0) - float(api.get(h, 0.0))) < 1e-5 for h in range(24))
    print("Comprobacion: coincide con get_hourly_measures_for_day:", "si" if coinciden else "NO - revisar")
    print()
    print(
        "Lectura rapida: el punto 04:00 Chile del grafico usa filas cuyo instante UTC, "
        "al pasar a Santiago, cae entre las 04:00 y 04:59 del 24-03 (no la fila T04Z sola en UTC)."
    )


if __name__ == "__main__":
    main()
