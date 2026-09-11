"""
Gestión hídrica — Clínica Bupa Antofagasta (formato Fundo Zapallar).

Genera:
  - One_Pager_Gestion_Hidrica_Bupa_Antofagasta_....pdf
  - Informe_Mensual_Bupa_Antofagasta_....pdf

Periodo por defecto: desde el inicio del monitoreo (23/07/2026) hasta hoy.
Gráfica comparativa de meses: solo desde junio 2026 (sin meses previos).
Sin referencias a riego (no aplica en clínica).

Uso:
  python generar_informes_gestion_hidrica_bupa_antofagasta.py
  python generar_informes_gestion_hidrica_bupa_antofagasta.py --start 23/07/2026 --end 11/09/2026
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import List

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
    except Exception:
        pass

import generar_informes_gestion_hidrica_lote_agosto2026 as lote
from informe_gestion_hidrica_pdf import (
    Accion,
    Hallazgo,
    InformeSpec,
    build_chart_6_meses,
    render_mensual,
    render_one_pager,
)

DEFAULT_START = "23/07/2026"
FOLDER = "Bupa_Antofagasta"
# Comparativo mensual: solo desde junio 2026 en adelante.
MES_INICIO_GRAFICO = (2026, 6)

_MESES = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}


def _fecha_larga(d: datetime) -> str:
    return f"{d.day} de {_MESES[d.month]} de {d.year}"


def _sin_riego(texto: str) -> str:
    """Bupa no tiene riego: limpia redacción genérica del lote Zapallar."""
    reemplazos = [
        ("horarios de riego, llenado o uso nocturno", "horarios de llenado o uso nocturno"),
        ("horarios de uso, riego o llenado", "horarios de uso o llenado"),
        ("fuga, riego continuo o falla", "fuga o falla"),
        ("fuga, riego continuo o error", "fuga o error"),
        ("riego, llenado o uso", "llenado o uso"),
        ("riego o llenado", "llenado"),
        ("riego continuo o ", ""),
        ("riego, ", ""),
        (" riego", ""),
        ("Riego", "Uso"),
        ("riego", "uso"),
    ]
    out = texto
    for a, b in reemplazos:
        out = out.replace(a, b)
    return out


def _limpiar_hallazgos(hallazgos: List[Hallazgo]) -> List[Hallazgo]:
    return [
        Hallazgo(
            prioridad=h.prioridad,
            titulo=_sin_riego(h.titulo),
            detalle=_sin_riego(h.detalle),
            lectura=_sin_riego(h.lectura),
        )
        for h in hallazgos
    ]


def _limpiar_acciones(acciones: List[Accion]) -> List[Accion]:
    cleaned: List[Accion] = []
    for a in acciones:
        cleaned.append(
            Accion(
                accion=_sin_riego(a.accion),
                plazo=a.plazo,
                objetivo=_sin_riego(a.objetivo),
                responsable=a.responsable,
            )
        )
    # Asegurar acciones típicas de clínica sin riego.
    if not any("nocturno" in a.accion.lower() or "llenado" in a.accion.lower() for a in cleaned):
        cleaned.insert(
            0,
            Accion(
                "Confirmar horarios de llenado o uso nocturno.",
                "7 días",
                "Separar consumo programado de posibles pérdidas.",
                "Administración / operación",
            ),
        )
    return cleaned[:3]


def _meses_antes_de_junio(end_dt: datetime) -> List[tuple]:
    """Meses a excluir del comparativo para dejar solo desde junio 2026."""
    y, m = MES_INICIO_GRAFICO
    excl: List[tuple] = []
    # Excluye desde ~1 año atrás hasta el mes anterior a junio 2026.
    cy, cm = end_dt.year, end_dt.month
    for _ in range(18):
        if (cy, cm) < (y, m):
            excl.append((cy, cm))
        cm -= 1
        if cm == 0:
            cm = 12
            cy -= 1
    return excl


def _cfg(start: str, end: str) -> dict:
    start_dt = datetime.strptime(start, "%d/%m/%Y")
    end_dt = datetime.strptime(end, "%d/%m/%Y")
    periodo_corto = (
        f"Clínica Bupa Antofagasta · {_fecha_larga(start_dt)} al {_fecha_larga(end_dt)}"
    )
    return {
        "key": "bupa_antofagasta",
        "company_id": "000029",
        "folder": FOLDER,
        "cliente": "Bupa Antofagasta",
        "sitio": "Clínica Bupa Antofagasta",
        "sujeto": "la clínica",
        "node_ids": [
            "000029-09",  # Matriz / medidor principal primero
            "000029-07",
            "000029-08",
            "000029-10",
        ],
        "start": start,
        "end": end,
        "periodo_corto": periodo_corto,
        "usar_kpi_ultimo_mes_6m": False,
        # Solo meses desde junio 2026 en la gráfica comparativa.
        "excluir_meses_6m": _meses_antes_de_junio(end_dt),
        "nota_agosto": (
            f"El periodo se informa del {start_dt.strftime('%d/%m/%Y')} "
            f"al {end_dt.strftime('%d/%m/%Y')} (desde el inicio del monitoreo). "
            "No se extrapola el consumo."
        ),
        "panorama_nota": (
            f"El consumo se considera del {start_dt.strftime('%d/%m/%Y')} "
            f"al {end_dt.strftime('%d/%m/%Y')}. No se extrapola el resto del mes."
        ),
        "apply_exclusions": False,
        "matriz_id": "000029-09",
        "matriz_name": "Medidor Principal Sanitaria",
        "additive": False,
        "nocturnal_explain": None,
        "kpi_label": "Consumo de entrada",
        "short_names": {
            "000029-09": "Medidor Principal Sanitaria",
            "000029-07": "Sala de Bomba Principal",
            "000029-08": "Sala de Bomba Sexto Piso",
            "000029-10": "Sala de Bomba N°2",
        },
        "leyenda": (
            "Medidor Principal Sanitaria (entrada real)",
            "Salas de bomba (aguas abajo / control interno)",
        ),
        "chart_nota": (
            "el Medidor Principal Sanitaria representa la entrada real a la clínica. "
            "Las salas de bomba miden tramos o impulsiones internas; por lo tanto, "
            "sus consumos no se suman al total de entrada."
        ),
        "nocturno_nota": (
            "Los volúmenes de las salas de bomba se expresan respecto de la entrada "
            "principal, pero no son aditivos porque corresponden a impulsiones internas."
        ),
    }


def _ajustar_spec_clinica(spec: InformeSpec, data: dict, label_periodo: str, cfg: dict) -> InformeSpec:
    pct = float(data["kpi"]["pct_nocturno"])
    motivo = (
        f"consumo nocturno del {pct:.0f} % sobre la entrada (Medidor Principal Sanitaria) "
        "que requiere seguimiento y validación frente a la operación habitual de la clínica "
        "(salas de bomba e impulsiones internas)."
    )
    hallazgos = _limpiar_hallazgos(list(spec.hallazgos))
    acciones = _limpiar_acciones(list(spec.acciones))

    # Rehacer gráfica 6m solo con meses desde junio (por si el cache trae más).
    charts = Path("reports") / FOLDER / "GESTION_HIDRICA" / "_charts"
    charts.mkdir(parents=True, exist_ok=True)
    labels: List[str] = []
    vals: List[float] = []
    for item in data.get("serie_6_meses") or []:
        lab = str(item["label"])
        parsed = lote._parse_lab_mes(lab)
        if parsed and parsed < MES_INICIO_GRAFICO:
            continue
        labels.append(lab.replace("*", "").split()[0].capitalize())
        vals.append(float(item["m3"]))
    chart_6m = spec.chart_6m
    if labels:
        chart_6m = build_chart_6_meses(charts / f"{cfg['key']}_6m.png", labels, vals)

    # Limpiar "riego" también en conclusiones / lectura ejecutiva.
    conclusion = [
        [(_sin_riego(t), b) for t, b in para]
        for para in (spec.conclusion or [])
    ]
    lectura = [
        [(_sin_riego(t), b) for t, b in para]
        for para in (spec.lectura_ejecutiva or [])
    ]

    kwargs = dict(
        footer=f"Informe de gestión hídrica - {cfg['cliente']} | {label_periodo}",
        titulo_onepager="Resumen ejecutivo de gestión hídrica",
        titulo_mensual="Informe mensual de gestión hídrica",
        hallazgos=hallazgos,
        acciones=acciones,
        conclusion=conclusion,
        lectura_ejecutiva=lectura,
        chart_6m=chart_6m,
    )
    if pct >= 18:
        kwargs["clasificacion"] = "EN OBSERVACIÓN"
        kwargs["motivo"] = motivo
    return replace(spec, **kwargs)


def main() -> int:
    ap = argparse.ArgumentParser(description="Gestión hídrica Bupa Antofagasta (formato Zapallar)")
    ap.add_argument("--start", default=DEFAULT_START, help="dd/mm/aaaa")
    ap.add_argument(
        "--end",
        default=datetime.now().strftime("%d/%m/%Y"),
        help="dd/mm/aaaa (default: hoy)",
    )
    args = ap.parse_args()

    cfg = _cfg(args.start, args.end)
    start_dt = datetime.strptime(args.start, "%d/%m/%Y")
    end_dt = datetime.strptime(args.end, "%d/%m/%Y")
    periodo_tag = f"{start_dt:%Y%m%d}_{end_dt:%Y%m%d}"
    label_periodo = f"{start_dt:%d/%m/%Y}–{end_dt:%d/%m/%Y}"

    lote.CACHE_DIR = Path("/tmp/gh_bupa_antofagasta")
    cache_file = lote.CACHE_DIR / f"{cfg['key']}.json"
    if cache_file.exists():
        cache_file.unlink()

    print("=" * 70)
    print("GESTIÓN HÍDRICA — Bupa Antofagasta (formato Fundo Zapallar)")
    print(f"Periodo: {args.start} → {args.end}")
    print("Gráfica mensual: desde junio 2026")
    print("Sin referencias a riego")
    print(f"Nodos: {', '.join(cfg['node_ids'])}")
    print("=" * 70)

    data = lote.fetch_cliente(cfg)
    # Forzar comparativo solo desde junio 2026 (evita meses anteriores al inicio).
    data["serie_6_meses"] = [
        item
        for item in (data.get("serie_6_meses") or [])
        if (lote._parse_lab_mes(str(item["label"])) or (9999, 12)) >= MES_INICIO_GRAFICO
    ]
    print(
        "[INFO] Meses en comparativo:",
        ", ".join(i["label"] for i in data.get("serie_6_meses") or []) or "(sin datos)",
        flush=True,
    )
    spec = lote.build_spec(cfg, data, visitas=[])
    spec = _ajustar_spec_clinica(spec, data, label_periodo, cfg)

    out_dir = Path("reports") / FOLDER / "GESTION_HIDRICA"
    charts = out_dir / "_charts"
    charts.mkdir(parents=True, exist_ok=True)

    one = out_dir / f"One_Pager_Gestion_Hidrica_Bupa_Antofagasta_{periodo_tag}.pdf"
    monthly = out_dir / f"Informe_Mensual_Bupa_Antofagasta_{periodo_tag}.pdf"

    render_one_pager(spec, one)
    render_mensual(spec, monthly, charts)
    print(f"[OK] {one}")
    print(f"[OK] {monthly}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
