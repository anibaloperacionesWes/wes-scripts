"""
Informe de gestión hídrica CORMUP — vacaciones septiembre 2026
(formato Zapallar / igual a Informe_Mensual_CORMUP_Agosto_2026.pdf).

Periodo informado: 14–20/09/2026 (semana con control especial por vacaciones).
Referencia comparativa: 07–13/09/2026 (semana sin control).

Cohorte: 10 colegios con control (excluye Tobalaba por pulso y los 3 sin válvula).

Uso:
  python generar_informe_gestion_hidrica_cormup_vacaciones_sept2026.py
"""

from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import generar_informes_gestion_hidrica_lote_agosto2026 as motor
from generar_comparativo_cormup_vacaciones_sept2026 import (
    COLEGIOS,
    CON_CONTROL_FIN,
    CON_CONTROL_INI,
    SIN_CONTROL,
    SIN_CONTROL_FIN,
    SIN_CONTROL_INI,
    TOBALABA,
    evaluar_colegios,
    evaluar_fuera_comparativo,
    precio_referencia_clp,
)
from generar_informes_gestion_hidrica_lote_agosto2026 import (
    build_spec,
    fetch_cliente,
)
from informe_gestion_hidrica_pdf import render_mensual, render_one_pager
from visitas_tecnicas_formulario import cargar_visitas_periodo, visitas_de_cliente

ROOT = Path(__file__).resolve().parent
CACHE = Path("/tmp/gh_cormup_vacaciones_sept2026")
OUT_DIR = ROOT / "reports" / "CORMUP" / "GESTION_HIDRICA"

SHORT_NAMES = {
    "000008-01": "Hermida F.",
    "000008-03": "C. Fernández",
    "000008-05": "Santa María",
    "000008-06": "Arrieta C.",
    "000008-07": "Erasmo Escala",
    "000008-09": "J.B. Pastene",
    "000008-10": "M. Huici",
    "000008-11": "Valle Hermoso",
    "000008-12": "Unión Árabe",
    "000008-14": "Juan Pablo II",
}


def _fmt(v: float, d: int = 1) -> str:
    s = f"{v:.{d}f}"
    a, _, b = s.partition(".")
    neg = a.startswith("-")
    if neg:
        a = a[1:]
    g = ""
    for i, ch in enumerate(reversed(a)):
        if i and i % 3 == 0:
            g = "." + g
        g = ch + g
    out = ("-" if neg else "") + g
    if d <= 0:
        return out
    b = b.rstrip("0") or "0"
    return out + "," + b


def _fmt_clp(v: float) -> str:
    return f"${_fmt(v, 0)}"


def _wow() -> Dict[str, Any]:
    """Comparativo 7–13 vs 14–20 para hallazgos (misma cohorte del informe)."""
    precio = precio_referencia_clp()
    filas = evaluar_colegios(max_workers=2)
    fuera = evaluar_fuera_comparativo(max_workers=2)
    tot_sin = sum(f.m3_sin for f in filas)
    tot_con = sum(f.m3_con for f in filas)
    ahorro = tot_sin - tot_con
    pct = (100.0 * ahorro / tot_sin) if tot_sin > 1e-9 else 0.0
    tob = next(f for f in fuera if f.node_id == TOBALABA[0])
    solo = [f for f in fuera if f.node_id != TOBALABA[0]]
    return {
        "precio": precio,
        "tot_sin": tot_sin,
        "tot_con": tot_con,
        "ahorro": ahorro,
        "pct": pct,
        "clp_ahorro": ahorro * precio,
        "clp_sin": tot_sin * precio,
        "clp_con": tot_con * precio,
        "tobalaba": tob,
        "solo_monitoreo": solo,
        "filas": filas,
    }


def _cfg(wow: Dict[str, Any]) -> dict:
    tob = wow["tobalaba"]
    solo_txt = ", ".join(f.nombre for f in wow["solo_monitoreo"])
    return {
        "key": "cormup_vac_14_20_sep2026",
        "company_id": "000008",
        "folder": "CORMUP",
        "cliente": "CORMUP",
        "sitio": "CORMUP Peñalolén",
        "sujeto": "la corporación",
        "verbo_registro": "registró",
        "node_ids": [nid for nid, _ in COLEGIOS],
        "start": "14/09/2026",
        "end": "20/09/2026",
        "periodo_corto": (
            "CORMUP Peñalolén · 14 al 20 de septiembre de 2026 "
            "(semana con control especial por vacaciones)"
        ),
        "apply_exclusions": False,
        "matriz_id": None,
        "matriz_name": "",
        "additive": True,
        "nocturnal_explain": None,
        "kpi_label": "Consumo total",
        "workers": 2,
        "short_names": dict(SHORT_NAMES),
        "leyenda": None,
        "skip_serie_6m": True,
        "chart_nota": (
            "Cada barra es un establecimiento de la cohorte con control. "
            "Tobalaba (pulso) y los tres colegios sin válvula no entran en este total."
        ),
        "nocturno_nota": (
            "El nocturno se suma entre establecimientos porque son recintos distintos. "
            "En CORMUP la ventana es UTC 00:00–07:00, igual que en la app."
        ),
        "ventana_nocturna": (
            "En los colegios CORMUP el nocturno se toma del CSV horario, marcas UTC 00:00 "
            "a 07:00 (misma ventana que la app). "
        ),
        "nota_agosto": (
            "Periodo informado: 14–20/09/2026 (7 días con control especial por vacaciones). "
            "No se extrapola. La semana previa 7–13/09 sirve de referencia sin control "
            f"({_fmt(wow['tot_sin'])} m³ / {_fmt_clp(wow['clp_sin'])}). "
            "Tobalaba queda fuera por falla de pulso; "
            f"{solo_txt} no tienen válvula de control."
        ),
        "panorama_nota": (
            "Semana con control especial por vacaciones de alumnos. "
            "Cohorte: 10 colegios con CPA/WES. Excluidos del total: Tobalaba "
            "(pulso — seguimiento pendiente) y Eduardo de la Barra, Alicura y Likankura "
            "(solo monitoreo, sin válvula)."
        ),
        "hallazgo_dato": {
            "prioridad": "INFORMATIVA",
            "titulo": (
                f"Ahorro vs semana sin control: {_fmt(wow['ahorro'])} m³ "
                f"({_fmt(wow['pct'])} %)"
            ),
            "dato": (
                f"7–13/09 (sin control): {_fmt(wow['tot_sin'])} m³ "
                f"({_fmt_clp(wow['clp_sin'])}). "
                f"14–20/09 (con control vacaciones): {_fmt(wow['tot_con'])} m³ "
                f"({_fmt_clp(wow['clp_con'])}). "
                f"Ahorro estimado: {_fmt_clp(wow['clp_ahorro'])} a "
                f"{_fmt_clp(wow['precio'])}/m³."
            ),
            "lectura": (
                "El resultado incorpora habilitaciones por obras (C. Fernández y J.B. Pastene "
                "09:00–18:00 toda la semana), patinaje (Arrieta lun–mar 17:30–21:00; "
                "Valle Hermoso lun y mié 17:30–21:00) y revisiones puntuales "
                "(Arrieta mar 15; Valle Hermoso mié 16)."
            ),
        },
        "hallazgos_extra": [
            {
                "prioridad": "ATENCIÓN",
                "titulo": "Tobalaba — seguimiento pendiente (pulso)",
                "dato": (
                    f"Excluido del total. Serie referencial no validada: "
                    f"7–13/09 {_fmt(tob.m3_sin)} m³; 14–20/09 {_fmt(tob.m3_con)} m³."
                ),
                "lectura": (
                    "Debe revisarse el pulso/telemetría para reincorporarlo al control "
                    "y a la cuantificación de ahorro. Mientras no se normalice, no entra "
                    "en rankings de cumplimiento."
                ),
            },
            {
                "prioridad": "INFORMATIVA",
                "titulo": "Colegios sin control hidráulico",
                "dato": (
                    "Eduardo de la Barra, Alicura y Likankura solo tienen monitoreo "
                    "(sin válvula WES). No formaron parte del programa de vacaciones "
                    "ni del cálculo de ahorro."
                    + "".join(
                        f" {f.nombre}: 7–13={_fmt(f.m3_sin)} m³, "
                        f"14–20={_fmt(f.m3_con)} m³."
                        for f in wow["solo_monitoreo"]
                    )
                ),
                "lectura": (
                    "Cualquier reducción en estos establecimientos requiere intervención "
                    "en terreno o ampliación de infraestructura de control."
                ),
            },
        ],
    }


def _inyectar_hallazgos_extra(spec, extras: Sequence[dict]) -> None:
    from informe_gestion_hidrica_pdf import Hallazgo

    extras_h = [
        Hallazgo(
            prioridad=str(h["prioridad"]),
            titulo=str(h["titulo"]),
            detalle=str(h["dato"]),
            lectura=str(h["lectura"]),
        )
        for h in extras
    ]
    # Insert after the first (hallazgo_dato / ahorro), keep the rest.
    if spec.hallazgos:
        spec.hallazgos = [spec.hallazgos[0]] + extras_h + list(spec.hallazgos[1:])
    else:
        spec.hallazgos = extras_h


def _patch_spec_copy(spec, wow: Dict[str, Any]) -> None:
    spec.footer = "Informe de gestión hídrica - CORMUP | Vacaciones 14-20 septiembre 2026"
    spec.titulo_mensual = "Informe de gestión hídrica — vacaciones"
    spec.titulo_onepager = "Resumen ejecutivo de gestión hídrica — vacaciones"
    # Refuerzo de conclusión con ahorro económico.
    ahorro_txt = (
        f"Respecto de la semana sin control (7–13/09), el ahorro volumétrico fue "
        f"{_fmt(wow['ahorro'])} m³ ({_fmt(wow['pct'])} %), equivalente a "
        f"{_fmt_clp(wow['clp_ahorro'])} a tarifa referencial de {_fmt_clp(wow['precio'])}/m³."
    )
    if spec.conclusion and isinstance(spec.conclusion, list):
        spec.conclusion = list(spec.conclusion) + [[(ahorro_txt, False)]]


def main() -> int:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
            sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
        except Exception:
            pass

    print("=" * 72)
    print("CORMUP — gestión hídrica vacaciones (formato Informe Mensual Agosto)")
    print("=" * 72)

    CACHE.mkdir(parents=True, exist_ok=True)
    motor.CACHE_DIR = CACHE

    print("\n[INFO] Comparativo 7–13 vs 14–20 (cohorte con control)…")
    wow = _wow()
    print(
        f"  sin={wow['tot_sin']:.1f} m³ | con={wow['tot_con']:.1f} m³ | "
        f"ahorro={wow['ahorro']:.1f} m³ ({wow['pct']:.1f} %) | "
        f"{wow['clp_ahorro']:.0f} CLP @ {wow['precio']:.0f}/m³"
    )

    cfg = _cfg(wow)
    extras = list(cfg.pop("hallazgos_extra"))

    print("\n[INFO] Fetch periodo 14–20/09…")
    data = fetch_cliente(cfg)

    start_dt = datetime(2026, 9, 14)
    end_dt = datetime(2026, 9, 20, 23, 59, 59)
    try:
        todas = cargar_visitas_periodo(start_dt, end_dt)
        visitas = visitas_de_cliente(todas, cfg)
        print(f"[INFO] Visitas técnicas en el periodo: {len(visitas)}")
    except Exception as exc:
        visitas = []
        print(f"[ADVERTENCIA] Visitas no disponibles: {exc}")

    spec = build_spec(cfg, data, visitas)
    _inyectar_hallazgos_extra(spec, extras)
    _patch_spec_copy(spec, wow)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    charts = OUT_DIR / "_charts_vacaciones_sep2026"
    charts.mkdir(parents=True, exist_ok=True)

    one = OUT_DIR / "One_Pager_Gestion_Hidrica_CORMUP_Vacaciones_14_20_Sep_2026.pdf"
    monthly = OUT_DIR / "Informe_Gestion_Hidrica_CORMUP_Vacaciones_14_20_Sep_2026.pdf"
    render_one_pager(spec, one)
    render_mensual(spec, monthly, charts)

    meta = {
        "one_pager": str(one),
        "informe": str(monthly),
        "periodo": "14/09/2026–20/09/2026",
        "referencia_sin_control": "07/09/2026–13/09/2026",
        "precio_clp_m3": wow["precio"],
        "total_sin_m3": wow["tot_sin"],
        "total_con_m3": wow["tot_con"],
        "ahorro_m3": wow["ahorro"],
        "ahorro_clp": wow["clp_ahorro"],
        "nodos": [nid for nid, _ in COLEGIOS],
    }
    (OUT_DIR / "meta_vacaciones_sep2026.json").write_text(
        json.dumps(meta, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"\n[OK] One-pager: {one}")
    print(f"[OK] Informe:   {monthly}")
    print(json.dumps(meta, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
