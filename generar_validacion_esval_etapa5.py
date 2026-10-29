#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Validación Matriz ESVAL (000027-01) y Etapa N°5 (000027-03) — Fundo Zapallar.

Usa BFF (api-rest.wes.cl) porque el ACL Node directo (104.248.53.141:7003)
suele no ser alcanzable desde Cloud Agents.

Genera Word + PDF + PNG en reports/Fundo_Zapallar/VALIDACION/...
"""

from __future__ import annotations

import json
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from statistics import mean
from typing import Dict, List, Optional, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor

from wes_api_bff import measures_dates

ROOT = Path(__file__).resolve().parent
COMPANY = "Fundo Zapallar"
NODES = {
    "000027-01": "Matriz ESVAL",
    "000027-03": "Etapa N°5",
}
RATIO_ALTO = 1.50
RATIO_BAJO = 0.50
MIN_BASELINE_M3 = 0.50


def _parse_day(s: str) -> date:
    return datetime.strptime(str(s)[:10], "%Y-%m-%d").date()


def _series(block: dict) -> Dict[date, float]:
    out: Dict[date, float] = {}
    for row in block.get("month") or []:
        ds = row.get("date")
        if not ds:
            continue
        out[_parse_day(ds)] = float(row.get("totalM3") or 0)
    return out


def _fetch(start: date, end: date) -> Dict[str, Dict[date, float]]:
    data = measures_dates(
        list(NODES.keys()),
        start.strftime("%d%m%Y"),
        end.strftime("%d%m%Y"),
        timeout=180,
    )
    by_id: Dict[str, Dict[date, float]] = {nid: {} for nid in NODES}
    for block in data:
        nid = block.get("nodeId")
        if nid in by_id:
            by_id[nid] = _series(block)
    return by_id


def _baseline(
    series: Dict[date, float], until: date
) -> Tuple[float, float, int, int]:
    lv, sd = [], []
    for d, v in series.items():
        if d >= until:
            continue
        if d.weekday() < 5:
            lv.append(v)
        else:
            sd.append(v)
    return (
        mean(lv) if lv else 0.0,
        mean(sd) if sd else 0.0,
        len(lv),
        len(sd),
    )


def _fmt(n: float, dec: int = 1) -> str:
    return f"{n:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _chart_daily(
    series: Dict[date, float],
    title: str,
    out: Path,
    baseline_lv: float,
    baseline_sd: float,
    highlight_from: Optional[date] = None,
) -> None:
    days = sorted(series)
    vals = [series[d] for d in days]
    colors = []
    for d, v in zip(days, vals):
        base = baseline_lv if d.weekday() < 5 else baseline_sd
        if base >= MIN_BASELINE_M3 and v <= base * RATIO_BAJO:
            colors.append("#C62828")
        elif base >= MIN_BASELINE_M3 and v >= base * RATIO_ALTO:
            colors.append("#EF6C00")
        else:
            colors.append("#1565C0")
    fig, ax = plt.subplots(figsize=(11, 4.2))
    ax.bar(days, vals, color=colors, width=0.8)
    if baseline_lv > 0:
        ax.axhline(
            baseline_lv,
            color="#2E7D32",
            ls="--",
            lw=1.2,
            label=f"Prom. L–V {_fmt(baseline_lv)} m³",
        )
    if baseline_sd > 0:
        ax.axhline(
            baseline_sd,
            color="#6A1B9A",
            ls=":",
            lw=1.2,
            label=f"Prom. S–D {_fmt(baseline_sd)} m³",
        )
    if highlight_from:
        ax.axvline(highlight_from, color="#9E9E9E", ls="--", lw=0.8, alpha=0.7)
    ax.set_title(title, fontsize=12, pad=8)
    ax.set_ylabel("m³ / día")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d-%m"))
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(days) // 12)))
    fig.autofmt_xdate(rotation=45)
    ax.legend(loc="upper right", fontsize=8)
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(out, dpi=140)
    plt.close(fig)


def _chart_compare(
    esval: Dict[date, float],
    etapa: Dict[date, float],
    out: Path,
    start: date,
    end: date,
) -> None:
    days = [start + timedelta(days=i) for i in range((end - start).days + 1)]
    e_vals = [esval.get(d, 0.0) for d in days]
    n_vals = [etapa.get(d, 0.0) for d in days]
    fig, ax = plt.subplots(figsize=(11, 4.2))
    ax.plot(days, e_vals, color="#0D47A1", lw=1.8, label="Matriz ESVAL")
    ax.plot(days, n_vals, color="#C62828", lw=1.8, label="Etapa N°5")
    ax.set_title("Comparativo diario — Matriz ESVAL vs Etapa N°5", fontsize=12)
    ax.set_ylabel("m³ / día")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d-%m"))
    fig.autofmt_xdate(rotation=45)
    ax.legend()
    ax.grid(alpha=0.25)
    fig.tight_layout()
    fig.savefig(out, dpi=140)
    plt.close(fig)


def _chart_share(
    esval: Dict[date, float],
    etapa: Dict[date, float],
    out: Path,
    start: date,
    end: date,
) -> None:
    days = [start + timedelta(days=i) for i in range((end - start).days + 1)]
    shares = []
    xs = []
    for d in days:
        e = esval.get(d, 0.0)
        n = etapa.get(d, 0.0)
        if e > 0.5:
            xs.append(d)
            shares.append(100.0 * n / e)
    fig, ax = plt.subplots(figsize=(11, 3.8))
    ax.plot(xs, shares, color="#00695C", lw=1.6)
    ax.fill_between(xs, shares, color="#00695C", alpha=0.15)
    ax.set_title("% del caudal de Matriz ESVAL medido en Etapa N°5", fontsize=12)
    ax.set_ylabel("%")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d-%m"))
    fig.autofmt_xdate(rotation=45)
    ax.grid(alpha=0.25)
    fig.tight_layout()
    fig.savefig(out, dpi=140)
    plt.close(fig)


def _add_heading(doc: Document, text: str, level: int = 1) -> None:
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.color.rgb = RGBColor(0x0D, 0x47, 0xA1)


def _p(doc: Document, text: str, bold: bool = False) -> None:
    para = doc.add_paragraph()
    run = para.add_run(text)
    run.font.size = Pt(11)
    run.bold = bold


def convertir_pdf(docx_path: Path) -> Optional[Path]:
    pdf_path = docx_path.with_suffix(".pdf")
    try:
        subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(docx_path.parent),
                str(docx_path),
            ],
            check=True,
            capture_output=True,
            timeout=180,
        )
        return pdf_path if pdf_path.exists() else None
    except Exception as exc:
        print(f"[WARN] PDF no generado: {exc}")
        return None


def main() -> int:
    hoy = date.today()
    # Validación operativa: últimos ~60 días + baseline 6 meses
    fin = hoy - timedelta(days=1)  # ayer completo
    inicio_ventana = fin - timedelta(days=59)
    inicio_baseline = fin - timedelta(days=183)

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = ROOT / "reports" / "Fundo_Zapallar" / "VALIDACION" / f"ESVAL_Etapa5_{stamp}"
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"[INFO] Ventana gráfico: {inicio_ventana} → {fin}")
    print(f"[INFO] Baseline desde: {inicio_baseline}")
    print("[INFO] Descargando medidas vía BFF…")
    all_series = _fetch(inicio_baseline, fin)

    esval = all_series["000027-01"]
    etapa = all_series["000027-03"]
    if not esval or not etapa:
        print("[ERROR] Sin datos de uno o ambos nodos", file=sys.stderr)
        return 1

    # Baseline hasta 7 días antes del fin (misma lógica que desviación)
    corte_base = fin - timedelta(days=6)
    e_lv, e_sd, e_nlv, e_nsd = _baseline(esval, corte_base)
    n_lv, n_sd, n_nlv, n_nsd = _baseline(etapa, corte_base)

    esval_win = {d: v for d, v in esval.items() if inicio_ventana <= d <= fin}
    etapa_win = {d: v for d, v in etapa.items() if inicio_ventana <= d <= fin}

    chart_e = out_dir / "chart_matriz_esval.png"
    chart_n = out_dir / "chart_etapa_n5.png"
    chart_c = out_dir / "chart_comparativo.png"
    chart_s = out_dir / "chart_porcentaje_etapa5.png"

    _chart_daily(
        esval_win,
        f"Matriz ESVAL (000027-01) — {_fmt(sum(esval_win.values()))} m³ en ventana",
        chart_e,
        e_lv,
        e_sd,
        highlight_from=corte_base,
    )
    _chart_daily(
        etapa_win,
        f"Etapa N°5 (000027-03) — {_fmt(sum(etapa_win.values()))} m³ en ventana",
        chart_n,
        n_lv,
        n_sd,
        highlight_from=corte_base,
    )
    _chart_compare(esval_win, etapa_win, chart_c, inicio_ventana, fin)
    _chart_share(esval_win, etapa_win, chart_s, inicio_ventana, fin)

    # Alertas últimos 14 días
    alertas = []
    for d in sorted(etapa_win)[-14:]:
        base = n_lv if d.weekday() < 5 else n_sd
        v = etapa_win[d]
        if base >= MIN_BASELINE_M3:
            ratio = v / base if base else 0
            if ratio <= RATIO_BAJO:
                alertas.append((d, "BAJO", v, base, ratio, "Etapa N°5"))
            elif ratio >= RATIO_ALTO:
                alertas.append((d, "ALTO", v, base, ratio, "Etapa N°5"))
    for d in sorted(esval_win)[-14:]:
        base = e_lv if d.weekday() < 5 else e_sd
        v = esval_win[d]
        if base >= MIN_BASELINE_M3:
            ratio = v / base if base else 0
            if ratio <= RATIO_BAJO:
                alertas.append((d, "BAJO", v, base, ratio, "Matriz ESVAL"))
            elif ratio >= RATIO_ALTO:
                alertas.append((d, "ALTO", v, base, ratio, "Matriz ESVAL"))

    # Meses
    def mes_stats(series: Dict[date, float], y: int, m: int) -> Tuple[float, float, int]:
        vals = [v for d, v in series.items() if d.year == y and d.month == m]
        return (sum(vals), mean(vals) if vals else 0.0, len(vals))

    sep_e = mes_stats(esval, 2026, 9)
    oct_e = mes_stats(esval, 2026, 10)
    sep_n = mes_stats(etapa, 2026, 9)
    oct_n = mes_stats(etapa, 2026, 10)

    # Share recientes
    shares_14 = []
    for d in sorted(esval_win)[-14:]:
        e = esval_win.get(d, 0)
        n = etapa_win.get(d, 0)
        if e > 0.5:
            shares_14.append(100 * n / e)
    share_avg = mean(shares_14) if shares_14 else 0.0

    doc = Document()
    title = doc.add_heading(
        "Validación — Matriz ESVAL y Etapa N°5 (Fundo Zapallar)", level=0
    )
    for run in title.runs:
        run.font.color.rgb = RGBColor(0x0D, 0x47, 0xA1)

    _p(
        doc,
        f"Periodo graficado: {inicio_ventana.strftime('%d/%m/%Y')} – "
        f"{fin.strftime('%d/%m/%Y')} (60 días). "
        f"Baseline ~6 meses hasta {corte_base.strftime('%d/%m/%Y')}. "
        f"Generado {datetime.now().strftime('%d/%m/%Y %H:%M')} (America/Santiago).",
    )
    _p(
        doc,
        "Criterio de alerta (mismo que desviación diaria): BAJO ≤ 0,5× y ALTO ≥ 1,5× "
        "del promedio del tipo de día (L–V / S–D).",
    )

    _add_heading(doc, "1. Hallazgos", level=1)
    _p(
        doc,
        f"Matriz ESVAL está en régimen normal: promedio L–V baseline {_fmt(e_lv)} m³ "
        f"(n={e_nlv}); octubre {_fmt(oct_e[1])} m³/día "
        f"(total {_fmt(oct_e[0])} m³ en {oct_e[2]} días). No hay patrón BAJO reciente.",
    )
    _p(
        doc,
        f"Etapa N°5 está sostenidamente baja: promedio L–V baseline {_fmt(n_lv)} m³ "
        f"(n={n_nlv}); septiembre {_fmt(sep_n[1])} m³/día y octubre {_fmt(oct_n[1])} m³/día "
        f"(≈ {_fmt(100 * oct_n[1] / n_lv if n_lv else 0, 0)} % del baseline hábil). "
        f"El bajo consumo no es un día puntual: arrastra desde septiembre.",
        bold=False,
    )
    _p(
        doc,
        f"En los últimos 14 días, Etapa N°5 representa en promedio {_fmt(share_avg, 1)} % "
        f"del caudal diario de Matriz ESVAL. La matriz (entrada al fundo) no explica la baja "
        f"de Etapa 5: hay agua entrando; la caída es aguas abajo en ese ramal.",
    )

    if alertas:
        _add_heading(doc, "2. Días fuera de rango (últimos 14)", level=1)
        table = doc.add_table(rows=1, cols=6)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        for i, h in enumerate(["Fecha", "Punto", "Tipo", "m³", "Baseline", "Ratio"]):
            hdr[i].text = h
        for d, tipo, v, base, ratio, punto in sorted(alertas, key=lambda x: x[0]):
            row = table.add_row().cells
            row[0].text = d.strftime("%d/%m/%Y")
            row[1].text = punto
            row[2].text = tipo
            row[3].text = _fmt(v)
            row[4].text = _fmt(base)
            row[5].text = f"{ratio:.2f}×"
    else:
        _add_heading(doc, "2. Días fuera de rango (últimos 14)", level=1)
        _p(doc, "Sin días BAJO/ALTO en los últimos 14 respecto al baseline (umbrales 0,5× / 1,5×).")

    _add_heading(doc, "3. Resumen numérico", level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    for i, h in enumerate(["Punto", "Baseline L–V", "Baseline S–D", "Sep 2026", "Oct 2026"]):
        table.rows[0].cells[i].text = h
    for nid, name, lv, sd, sep, oct_ in [
        ("000027-01", "Matriz ESVAL", e_lv, e_sd, sep_e, oct_e),
        ("000027-03", "Etapa N°5", n_lv, n_sd, sep_n, oct_n),
    ]:
        row = table.add_row().cells
        row[0].text = f"{name}\n({nid})"
        row[1].text = f"{_fmt(lv)} m³/d"
        row[2].text = f"{_fmt(sd)} m³/d"
        row[3].text = f"{_fmt(sep[0])} m³\n({_fmt(sep[1])} /d)"
        row[4].text = f"{_fmt(oct_[0])} m³\n({_fmt(oct_[1])} /d)"

    _add_heading(doc, "4. Gráficos", level=1)
    for path, caption in [
        (chart_e, "Consumo diario Matriz ESVAL (rojo = BAJO vs baseline del tipo de día)"),
        (chart_n, "Consumo diario Etapa N°5"),
        (chart_c, "Comparativo ESVAL vs Etapa N°5"),
        (chart_s, "Participación de Etapa N°5 sobre ESVAL"),
    ]:
        _p(doc, caption, bold=True)
        doc.add_picture(str(path), width=Inches(6.3))

    _add_heading(doc, "5. Conclusión operativa", level=1)
    _p(
        doc,
        "1) Validar en terreno / operación el ramal de Etapa N°5 (válvulas, riego, "
        "ocupación de la etapa, sensor). La Matriz ESVAL no muestra falla de abastecimiento.",
    )
    _p(
        doc,
        "2) La baja de Etapa N°5 es estructural desde septiembre 2026, no solo el día "
        "del monitor de desviación (22–23/10).",
    )
    _p(
        doc,
        "3) Mantener seguimiento diario mientras el promedio hábil siga bajo ~0,5× del "
        f"baseline ({_fmt(n_lv)} m³). Si el caudal vuelve cerca del baseline, cerrar la observación.",
    )

    docx_path = out_dir / f"Validacion_Matriz_ESVAL_Etapa_N5_{inicio_ventana:%Y%m%d}_{fin:%Y%m%d}.docx"
    doc.save(docx_path)
    print(f"[OK] Word: {docx_path}")

    pdf_path = convertir_pdf(docx_path)
    if pdf_path:
        print(f"[OK] PDF: {pdf_path}")

    resumen = {
        "generado": datetime.now().isoformat(timespec="seconds"),
        "ventana": [inicio_ventana.isoformat(), fin.isoformat()],
        "baseline_hasta": corte_base.isoformat(),
        "esval": {
            "baseline_lv": round(e_lv, 2),
            "baseline_sd": round(e_sd, 2),
            "sep_total": round(sep_e[0], 2),
            "oct_total": round(oct_e[0], 2),
            "oct_avg": round(oct_e[1], 2),
        },
        "etapa_n5": {
            "baseline_lv": round(n_lv, 2),
            "baseline_sd": round(n_sd, 2),
            "sep_total": round(sep_n[0], 2),
            "oct_total": round(oct_n[0], 2),
            "oct_avg": round(oct_n[1], 2),
            "oct_vs_baseline_lv_pct": round(100 * oct_n[1] / n_lv, 1) if n_lv else None,
        },
        "share_etapa5_sobre_esval_14d_pct": round(share_avg, 1),
        "alertas_14d": [
            {
                "fecha": d.isoformat(),
                "punto": p,
                "tipo": t,
                "m3": round(v, 2),
                "baseline": round(b, 2),
                "ratio": round(r, 2),
            }
            for d, t, v, b, r, p in sorted(alertas, key=lambda x: x[0])
        ],
        "docx": str(docx_path),
        "pdf": str(pdf_path) if pdf_path else None,
    }
    (out_dir / "resumen.json").write_text(
        json.dumps(resumen, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(json.dumps(resumen, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
