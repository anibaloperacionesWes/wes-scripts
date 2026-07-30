import sys

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from generar_reporte_word import (
    acl_node_base_url,
    fetch_json,
    flatten_measures,
    format_currency_chilean,
    format_number_chilean,
    normalize_measures_payload,
    parse_date,
    summarize_consumption,
)

NODOS = [
    ("000029-07", "Sala de Bomba Principal"),
    ("000029-08", "Sala de Bomba Sexto Piso"),
    ("000029-09", "Medidor Principal Sanitaria"),
    ("000029-10", "Sala de Bomba N°2"),
]

FACTURA_M3 = 6696.0
FACTURA_CLP = 18_538_860.0
PRECIO = FACTURA_CLP / FACTURA_M3


def periodo(start_s, end_s, ndias, label):
    start = parse_date(start_s)
    end = parse_date(end_s, end_of_day=True)
    print("=" * 70)
    print(f"{label} | {start_s} - {end_s} ({ndias} dias)")
    print("=" * 70)
    sanit = None
    bombas = 0.0
    for nid, name in NODOS:
        raw = fetch_json(
            f"{acl_node_base_url()}/nodes/measures/dates",
            params=[
                ("id", nid),
                ("start", start.strftime("%d%m%Y")),
                ("end", end.strftime("%d%m%Y")),
            ],
        )
        payload = normalize_measures_payload(raw, nid)
        measures = flatten_measures(payload)
        total = summarize_consumption(measures)["total"]
        prom = total / ndias
        proy = prom * 30.0
        clp = proy * PRECIO
        tipo = "cuenta" if nid.endswith("-09") else "bomba"
        if tipo == "cuenta":
            sanit = (total, prom, proy, clp)
        else:
            bombas += proy
        print(
            f"  {nid} {name}: periodo={total:.1f} m³ | "
            f"prom/día={prom:.2f} | proy30={proy:.1f} m³ | "
            f"{format_currency_chilean(clp)}"
        )
    if sanit:
        pct = bombas / sanit[2] * 100 if sanit[2] else 0
        gap_m3 = sanit[2] - bombas
        gap_clp = gap_m3 * PRECIO
        gap_pct = 100 - pct
        print(
            f"\n  Proyección cuenta (Sanitaria WES): {sanit[2]:.1f} m³ / "
            f"{format_currency_chilean(sanit[3])}"
        )
        print(f"  Factura julio (referencia histórica): {FACTURA_M3:.0f} m³ / "
              f"{format_currency_chilean(FACTURA_CLP)}")
        print(f"  Salas proyectadas: {bombas:.1f} m³ ({pct:.1f}% de la proyección Sanitaria)")
        print(
            f"  NO MONITOREADO (gap): {gap_m3:.1f} m³ ({gap_pct:.1f}%) / "
            f"{format_currency_chilean(gap_clp)}"
        )
    print()


if __name__ == "__main__":
    periodo("23/07/2026", "28/07/2026", 6, "INCLUYE HOY (día incompleto)")
    periodo("23/07/2026", "27/07/2026", 5, "SOLO DÍAS COMPLETOS")
