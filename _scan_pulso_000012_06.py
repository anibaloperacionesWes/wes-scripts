"""Busca pulsos solitarios 02:00-04:00 Chile en 000012-06 (dic 2025 - may 2026)."""
from __future__ import annotations

from datetime import date, datetime, timedelta, timezone
from typing import Dict, List, Optional, Tuple

import requests

from control_nocturno import (
    BASE_URL,
    _chile_hours_from_csv_text,
    _dt_to_chile,
    _utc_calendar_dates_for_chile_day,
    _value_by_time_last_row,
)
from generar_reporte_word import _value_by_time_sum_duplicate_rows

NODE = "000012-06"
HORAS_VENTANA = (2, 3, 4)  # 02:00, 03:00, 04:00 Chile


def fetch_chile_hours(node_id: str, dia: date) -> Dict[int, float]:
    acc: Dict[int, float] = {}
    url = f"{BASE_URL}/nodes/{node_id}/dates.measures.csv"
    for ud in _utc_calendar_dates_for_chile_day(dia):
        r = requests.get(
            url,
            params=[("start", ud.strftime("%d%m%Y")), ("end", ud.strftime("%d%m%Y"))],
            timeout=60,
        )
        r.raise_for_status()
        acc.update(_chile_hours_from_csv_text(r.text, dia))
    return {h: acc.get(h, 0.0) for h in range(24)}


def fetch_chile_hours_sum_dup(node_id: str, dia: date) -> Dict[int, float]:
    """Modo antiguo (suma duplicados) + hora UTC mal mapeada como Chile."""
    acc: Dict[int, float] = {}
    url = f"{BASE_URL}/nodes/{node_id}/dates.measures.csv"
    dd = dia.strftime("%d%m%Y")
    r = requests.get(url, params=[("start", dd), ("end", dd)], timeout=60)
    r.raise_for_status()
    for t_raw, val in sorted(_value_by_time_sum_duplicate_rows(r.text).items()):
        if "T" not in t_raw:
            continue
        try:
            tpart = t_raw.split("T", 1)[1].replace("Z", "").replace("z", "")
            if "." in tpart:
                tpart = tpart.split(".", 1)[0]
            hour = int(tpart.split(":", 1)[0])
            if 0 <= hour < 24:
                acc[hour] = float(val)
        except Exception:
            continue
    return {h: acc.get(h, 0.0) for h in range(24)}


def raw_rows_2_4_chile(node_id: str, dia: date) -> List[Tuple[str, float, int]]:
    """Filas CSV que caen en horas Chile 2,3,4."""
    out: List[Tuple[str, float, int]] = []
    url = f"{BASE_URL}/nodes/{node_id}/dates.measures.csv"
    for ud in _utc_calendar_dates_for_chile_day(dia):
        r = requests.get(
            url,
            params=[("start", ud.strftime("%d%m%Y")), ("end", ud.strftime("%d%m%Y"))],
            timeout=60,
        )
        r.raise_for_status()
        for t_raw, val in _value_by_time_last_row(r.text).items():
            if "T" not in t_raw:
                continue
            try:
                ts_norm = t_raw.strip().replace("Z", "+00:00")
                dt = datetime.fromisoformat(ts_norm)
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=timezone.utc)
                ch = _dt_to_chile(dt)
                if ch.date() == dia and ch.hour in HORAS_VENTANA:
                    out.append((t_raw, val, ch.hour))
            except (ValueError, TypeError):
                continue
    return sorted(out, key=lambda x: x[0])


def detectar_pulso_solitario(
    horas: Dict[int, float],
    *,
    min_pico: float = 2.0,
    min_ratio_vecinos: float = 1.8,
    max_vecino_pico: float = 1.2,
) -> Optional[Tuple[int, float, str]]:
    """
    Pulso solitario: una hora en {2,3,4} claramente mayor que las otras dos
    y mayor que horas adyacentes 1 y 5 (si hay dato).
    """
    vals = {h: horas.get(h, 0.0) for h in HORAS_VENTANA}
    if max(vals.values()) < min_pico:
        return None

    h_max = max(vals, key=lambda h: vals[h])
    v_max = vals[h_max]
    otros = [vals[h] for h in HORAS_VENTANA if h != h_max]
    if not otros or v_max < min_ratio_vecinos * max(otros):
        return None
    if max(otros) > max_vecino_pico and v_max < 2.5:
        return None

    v_1 = horas.get(1, 0.0)
    v_5 = horas.get(5, 0.0)
    if v_max < 1.5 * max(v_1, v_5, 0.01) and max(v_1, v_5) > 0.5:
        return None

    detalle = f"h2={vals[2]:.2f} h3={vals[3]:.2f} h4={vals[4]:.2f} | h1={v_1:.2f} h5={v_5:.2f}"
    return h_max, v_max, detalle


def main() -> None:
    inicio = date(2025, 12, 1)
    fin = date(2026, 5, 31)
    hoy = date.today()
    if fin > hoy:
        fin = hoy

    pulsos: List[Tuple[date, int, float, str]] = []
    top_h3: List[Tuple[date, float, str]] = []

    d = inicio
    while d <= fin:
        try:
            h = fetch_chile_hours(NODE, d)
            det = detectar_pulso_solitario(h)
            if det:
                hi, vmax, txt = det
                pulsos.append((d, hi, vmax, txt))
            v3 = h.get(3, 0.0)
            if v3 >= 1.61:
                top_h3.append((d, v3, f"h2={h[2]:.2f} h3={v3:.2f} h4={h[4]:.2f}"))
        except Exception as ex:
            print(f"  [skip] {d}: {ex}")
        d += timedelta(days=1)

    print(f"Nodo {NODE} | ventana Chile 02:00-04:59 (horas 2,3,4)")
    print(f"Periodo: {inicio} a {fin}")
    print()
    print("=== Pulsos solitarios (pico en UNA hora 2-4, mucho mayor que las otras) ===")
    if not pulsos:
        print("  (ninguno con criterio estricto)")
    for dia, hi, vmax, txt in sorted(pulsos, key=lambda x: -x[2]):
        wd = ["Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Dom"][dia.weekday()]
        print(f"  {dia.strftime('%d-%m-%Y')} ({wd})  hora {hi:02d}:00  pico={vmax:.2f}  |  {txt}")

    print()
    print("=== Top 20 por hora 03:00 Chile (metodo actual reporte) ===")
    for dia, v3, txt in sorted(top_h3, key=lambda x: -x[1])[:20]:
        wd = ["Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Dom"][dia.weekday()]
        print(f"  {dia.strftime('%d-%m-%Y')} ({wd})  {v3:.2f}  |  {txt}")

    # Verificar fechas que el usuario no ve en app: muestra CSV crudo
    print()
    print("=== Verificacion CSV crudo (fechas top abril) ===")
    for dia in [date(2026, 4, 12), date(2026, 4, 13), date(2026, 4, 9)]:
        rows = raw_rows_2_4_chile(NODE, dia)
        print(f"\n  {dia}:")
        for t, v, ch in rows:
            print(f"    TIME={t}  VALUE={v}  -> hora Chile {ch:02d}:00")

    # Buscar pulso con criterio mas laxo: cualquier hora 2-4 > 3 y las otras dos < 1.5
    print()
    print("=== Picos fuertes aislados (h2-4: una hora >3.0 y las otras dos <1.5) ===")
    d = inicio
    while d <= fin:
        try:
            h = fetch_chile_hours(NODE, d)
            v2, v3, v4 = h[2], h[3], h[4]
            for hi, vmax, otros in [(2, v2, [v3, v4]), (3, v3, [v2, v4]), (4, v4, [v2, v3])]:
                if vmax >= 3.0 and max(otros) < 1.5:
                    wd = ["Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Dom"][d.weekday()]
                    print(
                        f"  {d.strftime('%d-%m-%Y')} ({wd})  hora {hi:02d}:00={vmax:.2f}  "
                        f"resto 2-4: {otros[0]:.2f}, {otros[1]:.2f}"
                    )
        except Exception:
            pass
        d += timedelta(days=1)


if __name__ == "__main__":
    main()
