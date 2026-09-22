"""
Etapa N°5 (000027-03) desde instalación del sensor de pulso (13/11/2025):
consumo real vs caudal imposible en tubería DN90, y fecha de arranque del error.

Hidráulica: DN90 PE100 SDR17 (OD 90 mm, e≈5,4 mm → ID≈79,2 mm).
Q (m³/h) = v (m/s) · π · (ID/2)² · 3600.
"""

from __future__ import annotations

import csv
import json
import math
from collections import defaultdict
from datetime import date, datetime, time, timedelta
from pathlib import Path
from statistics import median
from zoneinfo import ZoneInfo

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from matplotlib.backends.backend_pdf import PdfPages

from generar_reporte_word import format_number_chilean
from generar_tabla_zapallar_matriz_vs_inferior import (
    HEADER_FILL,
    ROW_ALERTA_FILL,
    TOTAL_FILL,
    WES_BLUE,
    _fmt_m3,
    _set_cell_text,
    _set_landscape,
    load_hourly_range,
)
from generar_tabla_zapallar_wes_dia_completo import (
    expected_hours_chile,
    index_chile_hours,
)

CL = ZoneInfo("America/Santiago")
NODO_E5 = "000027-03"
NODO_INF = "000027-02"
INI_SENSOR = date(2025, 11, 13)
HOY = date(2026, 9, 22)

# DN90 PE100 SDR17
ID_M = 0.0792
AREA_M2 = math.pi * (ID_M / 2.0) ** 2


def q_ms(v: float) -> float:
    return v * AREA_M2 * 3600.0


# 2,0 m/s = tope de diseño habitual en red potable.
LIM_DISENO = q_ms(2.0)  # ≈ 35,5 m³/h
# 2,5 m/s = ya agresivo.
LIM_AGRESIVO = q_ms(2.5)  # ≈ 44,4 m³/h
# 3,0 m/s = no es consumo real en DN90 de loteo.
LIM_ERROR = q_ms(3.0)  # ≈ 53,3 m³/h

# Mesetas típicas del sensor inductivo (m³/h).
PLATEAU_96 = (80.0, 110.0)
PLATEAU_200 = (170.0, 230.0)

MES_ES = {
    1: "Ene",
    2: "Feb",
    3: "Mar",
    4: "Abr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Ago",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dic",
}


def v_from_q(q: float) -> float:
    if AREA_M2 <= 0:
        return 0.0
    return q / AREA_M2 / 3600.0


def clasificar(q: float) -> str:
    if PLATEAU_96[0] <= q <= PLATEAU_96[1]:
        return "meseta_96"
    if PLATEAU_200[0] <= q <= PLATEAU_200[1]:
        return "meseta_200"
    if q >= LIM_ERROR:
        return "error_dn90"
    if q >= LIM_AGRESIVO:
        return "sospechoso"
    return "ok"


def iter_chile_hours(start_d: date, end_d: date):
    d = start_d
    while d <= end_d:
        for h in sorted(expected_hours_chile(d)):
            yield d, h
        d += timedelta(days=1)


def stats(vals: list[float]) -> dict:
    if not vals:
        return {"n": 0, "min": 0, "p50": 0, "p90": 0, "p99": 0, "max": 0, "mean": 0}
    s = sorted(vals)
    n = len(s)

    def pct(p: float) -> float:
        if n == 1:
            return s[0]
        i = min(n - 1, max(0, int(round((p / 100.0) * (n - 1)))))
        return s[i]

    return {
        "n": n,
        "min": round(s[0], 3),
        "p50": round(pct(50), 3),
        "p90": round(pct(90), 3),
        "p99": round(pct(99), 3),
        "max": round(s[-1], 3),
        "mean": round(sum(s) / n, 3),
    }


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out = Path(f"reports/Fundo_Zapallar/ANALISIS/etapa5_pulso_dn90_{stamp}")
    out.mkdir(parents=True, exist_ok=True)

    print("[INFO] Descargando Etapa N°5 e Inferior...", flush=True)
    e5_rows = load_hourly_range(NODO_E5, INI_SENSOR, HOY)
    inf_rows = load_hourly_range(NODO_INF, INI_SENSOR, HOY)
    idx_e5 = index_chile_hours(e5_rows)
    idx_inf = index_chile_hours(inf_rows)

    hours: list[dict] = []
    first_error_hour: dict | None = None
    first_plateau: dict | None = None
    first_sustained: dict | None = None
    consec_bad = 0

    for d, h in iter_chile_hours(INI_SENSOR, HOY):
        key = (d, h)
        presente = key in idx_e5
        q = float(idx_e5.get(key, 0.0)) if presente else None
        kind = "hueco" if not presente else clasificar(q or 0.0)
        rec = {
            "date": d.isoformat(),
            "hour": h,
            "m3h": None if q is None else round(q, 4),
            "v_ms": None if q is None else round(v_from_q(q), 3),
            "clase": kind,
            "inf_m3h": round(float(idx_inf[key]), 4) if key in idx_inf else None,
        }
        hours.append(rec)
        if presente and kind in {"meseta_96", "meseta_200", "error_dn90"}:
            if first_error_hour is None:
                first_error_hour = rec
            if kind.startswith("meseta") and first_plateau is None:
                first_plateau = rec
            consec_bad += 1
            if consec_bad >= 3 and first_sustained is None:
                first_sustained = rec
        else:
            if presente:
                consec_bad = 0

    # Primera jornada con ≥ 6 h imposibles.
    by_day: dict[date, list[dict]] = defaultdict(list)
    for rec in hours:
        by_day[date.fromisoformat(rec["date"])].append(rec)

    first_bad_day: date | None = None
    daily: list[dict] = []
    for d in sorted(by_day):
        recs = by_day[d]
        presentes = [r for r in recs if r["clase"] != "hueco"]
        ok = [r for r in presentes if r["clase"] == "ok"]
        bad = [r for r in presentes if r["clase"] in {"meseta_96", "meseta_200", "error_dn90"}]
        susp = [r for r in presentes if r["clase"] == "sospechoso"]
        tot = sum(r["m3h"] or 0.0 for r in presentes)
        tot_ok = sum(r["m3h"] or 0.0 for r in ok)
        tot_bad = sum(r["m3h"] or 0.0 for r in bad)
        inf = sum(r["inf_m3h"] or 0.0 for r in recs if r["inf_m3h"] is not None)
        max_h = max((r["m3h"] or 0.0) for r in presentes) if presentes else 0.0
        n96 = sum(1 for r in presentes if r["clase"] == "meseta_96")
        n200 = sum(1 for r in presentes if r["clase"] == "meseta_200")
        nerr = sum(1 for r in presentes if r["clase"] == "error_dn90")
        if first_bad_day is None and (n96 + n200 + nerr) >= 6:
            first_bad_day = d
        daily.append(
            {
                "date": d.isoformat(),
                "horas": len(presentes),
                "m3": round(tot, 2),
                "m3_ok": round(tot_ok, 2),
                "m3_error": round(tot_bad, 2),
                "max_m3h": round(max_h, 2),
                "v_max": round(v_from_q(max_h), 2),
                "h_ok": len(ok),
                "h_sospechoso": len(susp),
                "h_error": nerr,
                "h_96": n96,
                "h_200": n200,
                "inf_m3": round(inf, 2),
                "e5_gt_inf": bool(presentes and tot > inf > 0 and tot > inf * 1.05),
            }
        )

    # Ventana sana: 13/11 hasta el día previo al primer error (hora o día sostenido).
    start_err = None
    if first_sustained:
        start_err = date.fromisoformat(first_sustained["date"])
    elif first_error_hour:
        start_err = date.fromisoformat(first_error_hour["date"])
    elif first_bad_day:
        start_err = first_bad_day

    sana_fin = (start_err - timedelta(days=1)) if start_err else HOY
    if sana_fin < INI_SENSOR:
        sana_fin = INI_SENSOR

    horas_sanas = [
        r["m3h"]
        for r in hours
        if r["clase"] == "ok"
        and r["m3h"] is not None
        and date.fromisoformat(r["date"]) <= sana_fin
    ]
    horas_sanas_pos = [v for v in horas_sanas if v > 0]
    dias_sanos = [
        d
        for d in daily
        if date.fromisoformat(d["date"]) <= sana_fin and d["h_96"] + d["h_200"] + d["h_error"] == 0
    ]
    m3_dia_sano = [d["m3"] for d in dias_sanos]
    m3_dia_sano_ok = [d["m3_ok"] for d in dias_sanos]

    # Mensual
    mensual: dict[tuple[int, int], dict] = {}
    for drow in daily:
        d = date.fromisoformat(drow["date"])
        k = (d.year, d.month)
        slot = mensual.setdefault(
            k,
            {
                "m3": 0.0,
                "m3_ok": 0.0,
                "m3_error": 0.0,
                "h_ok": 0,
                "h_error": 0,
                "h_96": 0,
                "h_200": 0,
                "h_sospechoso": 0,
                "dias": 0,
                "dias_error": 0,
                "max_m3h": 0.0,
                "inf_m3": 0.0,
            },
        )
        slot["m3"] += drow["m3"]
        slot["m3_ok"] += drow["m3_ok"]
        slot["m3_error"] += drow["m3_error"]
        slot["h_ok"] += drow["h_ok"]
        slot["h_error"] += drow["h_error"]
        slot["h_96"] += drow["h_96"]
        slot["h_200"] += drow["h_200"]
        slot["h_sospechoso"] += drow["h_sospechoso"]
        slot["dias"] += 1
        if drow["h_96"] + drow["h_200"] + drow["h_error"] > 0:
            slot["dias_error"] += 1
        slot["max_m3h"] = max(slot["max_m3h"], drow["max_m3h"])
        slot["inf_m3"] += drow["inf_m3"]

    meses = []
    for (y, m), sl in sorted(mensual.items()):
        meses.append(
            {
                "mes": f"{MES_ES[m]}-{str(y)[2:]}",
                "year": y,
                "month": m,
                **{k: (round(v, 2) if isinstance(v, float) else v) for k, v in sl.items()},
                "v_max": round(v_from_q(sl["max_m3h"]), 2),
            }
        )

    st_h = stats([v for v in horas_sanas if v is not None])
    st_h_pos = stats(horas_sanas_pos)
    st_dia = stats(m3_dia_sano)

    hallazgo = {
        "nodo": NODO_E5,
        "punto": "Etapa N°5",
        "sensor_desde": INI_SENSOR.isoformat(),
        "hasta": HOY.isoformat(),
        "dn90": {
            "id_mm": round(ID_M * 1000, 1),
            "q_1ms": round(q_ms(1.0), 1),
            "q_15ms": round(q_ms(1.5), 1),
            "q_2ms": round(LIM_DISENO, 1),
            "q_25ms": round(LIM_AGRESIVO, 1),
            "q_3ms": round(LIM_ERROR, 1),
            "q_96_vel": round(v_from_q(96.0), 2),
            "q_200_vel": round(v_from_q(200.0), 2),
        },
        "primer_hora_imposible": first_error_hour,
        "primera_meseta": first_plateau,
        "primer_tramo_3h": first_sustained,
        "primer_dia_6h_error": first_bad_day.isoformat() if first_bad_day else None,
        "ventana_sana": {
            "desde": INI_SENSOR.isoformat(),
            "hasta": sana_fin.isoformat(),
            "dias_sin_error": len(dias_sanos),
            "horario_m3h": st_h,
            "horario_m3h_solo_positivo": st_h_pos,
            "diario_m3": st_dia,
            "diario_mediana": round(median(m3_dia_sano), 2) if m3_dia_sano else 0,
            "diario_promedio": round(sum(m3_dia_sano) / len(m3_dia_sano), 2) if m3_dia_sano else 0,
        },
        "totales": {
            "m3_reportado": round(sum(d["m3"] for d in daily), 1),
            "m3_horas_ok": round(sum(d["m3_ok"] for d in daily), 1),
            "m3_horas_error": round(sum(d["m3_error"] for d in daily), 1),
            "horas_ok": sum(d["h_ok"] for d in daily),
            "horas_96": sum(d["h_96"] for d in daily),
            "horas_200": sum(d["h_200"] for d in daily),
            "horas_error_otros": sum(d["h_error"] for d in daily),
            "dias_e5_gt_inf": sum(1 for d in daily if d["e5_gt_inf"]),
        },
        "meses": meses,
    }

    (out / "hallazgo_etapa5_pulso_dn90.json").write_text(
        json.dumps(hallazgo, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    with (out / "diario_etapa5.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(daily[0].keys()))
        w.writeheader()
        w.writerows(daily)

    print(json.dumps({k: hallazgo[k] for k in hallazgo if k != "meses"}, indent=2, ensure_ascii=False))
    print("\n=== MENSUAL ===")
    for m in meses:
        print(
            f"{m['mes']:8} reportado={m['m3']:9.1f}  ok={m['m3_ok']:8.1f}  "
            f"error={m['m3_error']:9.1f}  h96={m['h_96']:4} h200={m['h_200']:4} "
            f"max={m['max_m3h']:.1f} m³/h ({m['v_max']} m/s)"
        )

    # --- gráficos ---
    fechas = [date.fromisoformat(d["date"]) for d in daily]
    fig, ax = plt.subplots(figsize=(14.2, 5.2))
    ax.bar(
        fechas,
        [d["m3_ok"] for d in daily],
        color="#70AD47",
        width=1.0,
        label="Horas ≤ 2,5 m/s (plausible DN90)",
    )
    ax.bar(
        fechas,
        [d["m3_error"] for d in daily],
        bottom=[d["m3_ok"] for d in daily],
        color="#C00000",
        width=1.0,
        label="Horas error (meseta 96/200 o > 3 m/s)",
    )
    if start_err:
        ax.axvline(start_err, color="#1F4E79", ls="--", lw=1.4, label=f"Arranque error {start_err.strftime('%d/%m/%Y')}")
    ax.set_ylabel("m³/día")
    ax.set_title(
        "Etapa N°5 desde sensor de pulso (13/11/2025): m³/día real vs error",
        fontweight="bold",
        color=WES_BLUE,
    )
    ax.legend(frameon=False, fontsize=8)
    ax.grid(axis="y", ls="--", alpha=0.35)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    chart_dia = out / "chart_diario.png"
    fig.savefig(chart_dia, dpi=150)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(14.2, 4.6))
    ax.plot(fechas, [d["max_m3h"] for d in daily], color=WES_BLUE, lw=1.1, label="Máximo horario del día")
    ax.axhline(LIM_DISENO, color="#70AD47", ls="--", lw=1.1, label=f"2,0 m/s ≈ {_fmt_m3(LIM_DISENO, 0)} m³/h")
    ax.axhline(LIM_ERROR, color="#ED7D31", ls="--", lw=1.1, label=f"3,0 m/s ≈ {_fmt_m3(LIM_ERROR, 0)} m³/h")
    ax.axhline(96, color="#C00000", ls=":", lw=1.2, label=f"Meseta ~96 m³/h (≈ {v_from_q(96):.1f} m/s)")
    ax.axhline(200, color="#7030A0", ls=":", lw=1.2, label=f"Meseta ~200 m³/h (≈ {v_from_q(200):.1f} m/s)")
    ax.set_ylabel("m³/h")
    ax.set_title("Caudal horario máximo diario vs capacidad DN90", fontweight="bold", color=WES_BLUE)
    ax.legend(frameon=False, fontsize=7.5, ncol=2)
    ax.set_ylim(0, max(220, max(d["max_m3h"] for d in daily) * 1.05))
    ax.grid(axis="y", ls="--", alpha=0.35)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    chart_max = out / "chart_max_horario.png"
    fig.savefig(chart_max, dpi=150)
    plt.close(fig)

    labels = [m["mes"] for m in meses]
    x = np.arange(len(labels))
    fig, ax = plt.subplots(figsize=(13.5, 4.5))
    ax.bar(x - 0.2, [m["m3_ok"] for m in meses], 0.4, color="#70AD47", label="Horas plausibles DN90")
    ax.bar(x + 0.2, [m["m3_error"] for m in meses], 0.4, color="#C00000", label="Horas error")
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.set_ylabel("m³")
    ax.set_title("Etapa N°5 mensual: volumen en horas reales vs horas de error", fontweight="bold", color=WES_BLUE)
    ax.legend(frameon=False, fontsize=8)
    ax.grid(axis="y", ls="--", alpha=0.35)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    chart_mes = out / "chart_mensual.png"
    fig.savefig(chart_mes, dpi=150)
    plt.close(fig)

    pdf_path = out / "Etapa5_sensor_pulso_DN90.pdf"
    docx_path = out / "Etapa5_sensor_pulso_DN90.docx"
    _pdf(hallazgo, daily, chart_dia, chart_max, chart_mes, pdf_path)
    _word(hallazgo, chart_dia, chart_max, chart_mes, docx_path)
    print(f"[OK] {pdf_path}")
    print(f"[OK] {docx_path}")
    print(f"[OK] {out}")


def _pdf(h: dict, daily: list[dict], c1: Path, c2: Path, c3: Path, out: Path) -> None:
    vs = h["ventana_sana"]
    dn = h["dn90"]
    tot = h["totales"]
    fe = h["primer_hora_imposible"]
    fs = h["primer_tramo_3h"]
    fd = h["primer_dia_6h_error"]

    def fhora(rec: dict | None) -> str:
        if not rec:
            return "no aparece en la serie"
        d = date.fromisoformat(rec["date"])
        return f"{d.strftime('%d/%m/%Y')} {rec['hour']:02d}:00  →  {_fmt_m3(rec['m3h'], 1)} m³/h  ({rec['v_ms']} m/s)  [{rec['clase']}]"

    with PdfPages(out) as pdf:
        fig = plt.figure(figsize=(16.4, 10.6))
        fig.suptitle(
            "Fundo Zapallar — Etapa N°5 · sensor de pulso desde 13/11/2025 · tubería DN90",
            fontsize=14.2,
            fontweight="bold",
            color=WES_BLUE,
            x=0.02,
            ha="left",
            y=0.985,
        )
        lines = [
            (
                f"DN90 PE100 SDR17 (ID 79,2 mm): 1,0 m/s = {dn['q_1ms']} m³/h · "
                f"2,0 m/s = {dn['q_2ms']} m³/h (diseño) · 3,0 m/s = {dn['q_3ms']} m³/h (tope). "
                f"96 m³/h = {dn['q_96_vel']} m/s · 200 m³/h = {dn['q_200_vel']} m/s → no es agua en esa cañería."
            ),
            (
                f"Consumo real (días sanos {date.fromisoformat(vs['desde']).strftime('%d/%m/%Y')}–"
                f"{date.fromisoformat(vs['hasta']).strftime('%d/%m/%Y')}, {vs['dias_sin_error']} días): "
                f"mediana {_fmt_m3(vs['diario_mediana'], 1)} m³/día · promedio {_fmt_m3(vs['diario_promedio'], 1)} m³/día · "
                f"P90 horario {_fmt_m3(vs['horario_m3h']['p90'], 2)} m³/h · máx {_fmt_m3(vs['horario_m3h']['max'], 1)} m³/h."
            ),
            f"Primer error puntual: {fhora(fe)}   |   Error sostenido (≥3 h): {fhora(fs)}   |   ≥6 h en un día: {fd or '—'}",
            (
                f"13/11/25–22/09/26: reportado {_fmt_m3(tot['m3_reportado'], 0)} m³ · "
                f"horas OK {_fmt_m3(tot['m3_horas_ok'], 0)} m³ · horas error {_fmt_m3(tot['m3_horas_error'], 0)} m³ · "
                f"h@96={tot['horas_96']} · h@200={tot['horas_200']} · días Etapa 5 > Inferior={tot['dias_e5_gt_inf']}."
            ),
        ]
        y = 0.955
        for line in lines:
            fig.text(0.02, y, line, fontsize=8.0, va="top", color="#222222")
            y -= 0.022
        headers = [
            "Mes",
            "m³ reportado",
            "m³ horas OK",
            "m³ horas error",
            "h 96",
            "h 200",
            "Máx m³/h",
            "v máx m/s",
            "Días con error",
        ]
        body = []
        for m in h["meses"]:
            body.append(
                [
                    m["mes"],
                    _fmt_m3(m["m3"], 1),
                    _fmt_m3(m["m3_ok"], 1),
                    _fmt_m3(m["m3_error"], 1),
                    str(m["h_96"]),
                    str(m["h_200"]),
                    _fmt_m3(m["max_m3h"], 1),
                    format_number_chilean(m["v_max"], 2),
                    f"{m['dias_error']}/{m['dias']}",
                ]
            )
        tot_row = [
            "TOTAL",
            _fmt_m3(tot["m3_reportado"], 1),
            _fmt_m3(tot["m3_horas_ok"], 1),
            _fmt_m3(tot["m3_horas_error"], 1),
            str(tot["horas_96"]),
            str(tot["horas_200"]),
            "",
            "",
            "",
        ]
        cell = [headers] + body + [tot_row]
        ax = fig.add_axes([0.02, 0.48, 0.96, 0.32])
        ax.axis("off")
        table = ax.table(cellText=cell, loc="center", cellLoc="center")
        table.auto_set_font_size(False)
        table.set_fontsize(7.4)
        table.scale(1, 1.38)
        for c in range(len(headers)):
            hd = table[0, c]
            hd.set_facecolor(WES_BLUE)
            hd.get_text().set_color("white")
            hd.get_text().set_fontweight("bold")
        for r, m in enumerate(h["meses"], start=1):
            bad = m["h_96"] + m["h_200"] + m["h_error"] > 0
            for c in range(len(headers)):
                ci = table[r, c]
                if c in (3, 4, 5) and bad:
                    ci.set_facecolor("#F4C7C3")
                elif c in (2,) and m["m3_ok"] > 0:
                    ci.set_facecolor("#E2EFDA")
        last = len(cell) - 1
        for c in range(len(headers)):
            table[last, c].set_facecolor("#D6E3F0")
            table[last, c].get_text().set_fontweight("bold")

        ax2 = fig.add_axes([0.04, 0.04, 0.92, 0.42])
        ax2.imshow(plt.imread(str(c1)))
        ax2.axis("off")
        fig.text(
            0.02,
            0.012,
            "Hora OK = caudal < 2,5 m/s en DN90 (≈44 m³/h). Error = meseta 80–110 o 170–230 m³/h, o cualquier hora ≥ 3 m/s. "
            "Acta 05/08/2025 N°2014: peso pulso Superior 1 / Etapa 5 = 10 L/pulso. "
            "96 m³/h ≈ 2,67 pulsos/s; 200 m³/h = 50 Hz/9 con ese factor → ruido eléctrico del inductivo, no caudal.",
            fontsize=7.2,
            color="#444444",
        )
        pdf.savefig(fig)
        plt.close(fig)

        fig = plt.figure(figsize=(16.4, 10.6))
        fig.suptitle(
            "Etapa N°5 — caudal máximo diario y volumen mensual (real vs error)",
            fontsize=14,
            fontweight="bold",
            color=WES_BLUE,
            x=0.02,
            ha="left",
            y=0.98,
        )
        ax = fig.add_axes([0.04, 0.50, 0.92, 0.44])
        ax.imshow(plt.imread(str(c2)))
        ax.axis("off")
        axb = fig.add_axes([0.04, 0.04, 0.92, 0.42])
        axb.imshow(plt.imread(str(c3)))
        axb.axis("off")
        pdf.savefig(fig)
        plt.close(fig)


def _word(h: dict, c1: Path, c2: Path, c3: Path, out: Path) -> None:
    vs = h["ventana_sana"]
    dn = h["dn90"]
    tot = h["totales"]
    doc = Document()
    _set_landscape(doc)
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("Fundo Zapallar — Etapa N°5: consumo real en DN90 y arranque del error")
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = RGBColor(31, 78, 121)

    p = doc.add_paragraph()
    p.add_run(
        f"Sensor de pulso (inductivo sobre Sensus, cámara MAP) desde el 13/11/2025. "
        f"Tubería DN90 PE100 SDR17, ID 79,2 mm. "
        f"2,0 m/s = {dn['q_2ms']} m³/h (diseño). 3,0 m/s = {dn['q_3ms']} m³/h (tope). "
        f"96 m³/h = {dn['q_96_vel']} m/s y 200 m³/h = {dn['q_200_vel']} m/s: no pasan por esa cañería."
    ).font.size = Pt(10)

    fe = h["primer_hora_imposible"]
    fs = h["primer_tramo_3h"]
    doc.add_paragraph().add_run("Hallazgo").bold = True
    bullets = [
        f"Consumo real (sin mesetas) {date.fromisoformat(vs['desde']).strftime('%d/%m/%Y')}–"
        f"{date.fromisoformat(vs['hasta']).strftime('%d/%m/%Y')}: "
        f"mediana {format_number_chilean(vs['diario_mediana'], 1)} m³/día, "
        f"promedio {format_number_chilean(vs['diario_promedio'], 1)} m³/día. "
        f"Horario P50 {format_number_chilean(vs['horario_m3h']['p50'], 2)} m³/h, "
        f"P90 {format_number_chilean(vs['horario_m3h']['p90'], 2)} m³/h "
        f"(~{format_number_chilean(v_from_q(vs['horario_m3h']['p90']), 2)} m/s).",
        (
            "Primer caudal imposible: "
            + (
                f"{date.fromisoformat(fe['date']).strftime('%d/%m/%Y')} {fe['hour']:02d}:00 "
                f"= {format_number_chilean(fe['m3h'], 1)} m³/h ({fe['v_ms']} m/s)."
                if fe
                else "no aparece."
            )
        ),
        (
            "Error sostenido (≥ 3 h seguidas): "
            + (
                f"{date.fromisoformat(fs['date']).strftime('%d/%m/%Y')} {fs['hour']:02d}:00."
                if fs
                else "no aparece."
            )
        ),
        f"Primer día con ≥ 6 h imposibles: {h['primer_dia_6h_error'] or '—'}.",
        f"Del 13/11/2025 al 22/09/2026 el punto reportó {format_number_chilean(tot['m3_reportado'], 0)} m³; "
        f"de esos, {format_number_chilean(tot['m3_horas_error'], 0)} m³ están en horas de error "
        f"({tot['horas_96']} h en ~96 m³/h y {tot['horas_200']} h en ~200 m³/h). "
        f"Etapa 5 superó al Estanque Inferior en {tot['dias_e5_gt_inf']} días: imposible hidráulicamente.",
    ]
    for b in bullets:
        doc.add_paragraph(b, style="List Bullet")

    headers = [
        "Mes",
        "m³ reportado",
        "m³ horas OK",
        "m³ horas error",
        "h ~96",
        "h ~200",
        "Máx m³/h",
        "v máx",
        "Días error",
    ]
    table = doc.add_table(rows=1 + len(h["meses"]) + 1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = 1  # center
    for i, name in enumerate(headers):
        _set_cell_text(table.rows[0].cells[i], name, bold=True, size=8, color=RGBColor(255, 255, 255), fill=HEADER_FILL)
    for r_i, m in enumerate(h["meses"], start=1):
        vals = [
            m["mes"],
            _fmt_m3(m["m3"], 1),
            _fmt_m3(m["m3_ok"], 1),
            _fmt_m3(m["m3_error"], 1),
            str(m["h_96"]),
            str(m["h_200"]),
            _fmt_m3(m["max_m3h"], 1),
            f"{format_number_chilean(m['v_max'], 2)} m/s",
            f"{m['dias_error']}/{m['dias']}",
        ]
        bad = m["h_96"] + m["h_200"] > 0
        for c, val in enumerate(vals):
            fill = ROW_ALERTA_FILL if bad and c in (3, 4, 5) else ("E2EFDA" if c == 2 else None)
            _set_cell_text(table.rows[r_i].cells[c], val, size=8, fill=fill)
    last = table.rows[-1]
    tots = [
        "TOTAL",
        _fmt_m3(tot["m3_reportado"], 1),
        _fmt_m3(tot["m3_horas_ok"], 1),
        _fmt_m3(tot["m3_horas_error"], 1),
        str(tot["horas_96"]),
        str(tot["horas_200"]),
        "",
        "",
        "",
    ]
    for c, val in enumerate(tots):
        _set_cell_text(last.cells[c], val, bold=True, size=8, fill=TOTAL_FILL)

    for img in (c1, c2, c3):
        doc.add_picture(str(img), width=Inches(9.6))

    note = doc.add_paragraph()
    note.add_run(
        "Criterio: hora OK si v < 2,5 m/s (≈44 m³/h). Error si meseta 80–110 o 170–230 m³/h, o v ≥ 3 m/s. "
        "Peso pulso del acta 05/08/2025 N°2014 = 10 L/pulso. Las mesetas 96 y 200 m³/h calzan con acople "
        "del sensor inductivo a 50 Hz, no con un caudal de DN90."
    ).font.size = Pt(9)
    doc.save(out)


if __name__ == "__main__":
    main()
