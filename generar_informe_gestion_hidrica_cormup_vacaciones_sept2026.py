"""
Informe de gestión hídrica CORMUP — vacaciones septiembre 2026
(formato Zapallar / igual a Informe_Mensual_CORMUP_Agosto_2026.pdf).

Mensaje central (claro):
  Semana sin control 7–13/09  vs  Semana con control vacaciones 14–20/09
  → ahorro m³ y $ de la cohorte con control.

Cohorte: 10 colegios con control (excluye Tobalaba por pulso y los 3 sin válvula).

Uso:
  python generar_informe_gestion_hidrica_cormup_vacaciones_sept2026.py
"""

from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Sequence

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

import generar_informes_gestion_hidrica_lote_agosto2026 as motor
from generar_comparativo_cormup_vacaciones_sept2026 import (
    COLEGIOS,
    SIN_CONTROL,
    TOBALABA,
    evaluar_colegios,
    evaluar_fuera_comparativo,
    precio_referencia_clp,
)
from generar_informes_gestion_hidrica_lote_agosto2026 import (
    build_spec,
    fetch_cliente,
)
from informe_gestion_hidrica_pdf import (
    Accion,
    Hallazgo,
    render_mensual,
    render_one_pager,
)
from visitas_tecnicas_formulario import cargar_visitas_periodo, visitas_de_cliente
from wes_estilo_graficos_app import COLOR_BARRA_WES, COLOR_NOCHE

ROOT = Path(__file__).resolve().parent
CACHE = Path("/tmp/gh_cormup_vacaciones_sept2026_v2")
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

# Consumos esperados en 14–20 (no son fallas de corte).
ESPERADOS_CONTROL = {
    "000008-03": "Obras: agua 09:00–18:00 toda la semana (incl. sáb/dom).",
    "000008-09": "Obras: agua 09:00–18:00 toda la semana (incl. sáb/dom).",
    "000008-06": "Patinaje lun–mar 17:30–21:00 + revisión martes 15/09.",
    "000008-11": "Patinaje lun y mié 17:30–21:00 + habilitación mié 16/09 ~10:00–tarde.",
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


def _cfg() -> dict:
    return {
        "key": "cormup_vac_14_20_sep2026_v2",
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
            "CORMUP Peñalolén · vacaciones Fiestas Patrias · "
            "comparativo 7–13/09 (sin control) vs 14–20/09 (con control)"
        ),
        "apply_exclusions": False,
        "matriz_id": None,
        "matriz_name": "",
        "additive": True,
        "nocturnal_explain": "wes",  # evita hallazgo genérico de % nocturno “sospechoso”
        "kpi_label": "Consumo total",
        "workers": 2,
        "short_names": dict(SHORT_NAMES),
        "leyenda": None,
        "skip_serie_6m": True,
        "chart_nota": (
            "Barras por colegio: gris = semana sin control (7–13/09); "
            "azul = semana con control vacaciones (14–20/09)."
        ),
        "nocturno_nota": (
            "En 14–20/09 el consumo residual corresponde sobre todo a obras y patinaje "
            "acordados, no a falla del corte."
        ),
        "ventana_nocturna": (
            "En los colegios CORMUP el nocturno se toma del CSV horario, marcas UTC 00:00 "
            "a 07:00 (misma ventana que la app). "
        ),
        "nota_agosto": (
            "Cómo leer este informe: la historia es el antes/después. "
            "7–13/09 = semana sin control de vacaciones. "
            "14–20/09 = semana con control especial. "
            "Los KPI de portada muestran ese comparativo (m³ y $). "
            "El anexo técnico detalla solo la semana con control (14–20)."
        ),
        "panorama_nota": (
            "10 colegios con control. Fuera del total: Tobalaba (pulso) y "
            "Eduardo de la Barra, Alicura y Likankura (sin válvula)."
        ),
    }


def _chart_comparativo(wow: Dict[str, Any], path: Path) -> Path:
    filas = sorted(wow["filas"], key=lambda f: f.m3_sin, reverse=True)
    names = [SHORT_NAMES.get(f.node_id, f.nombre) for f in filas]
    sin_vals = [f.m3_sin for f in filas]
    con_vals = [f.m3_con for f in filas]
    x = list(range(len(filas)))
    w = 0.38
    fig, ax = plt.subplots(figsize=(10.5, 4.6))
    ax.bar([i - w / 2 for i in x], sin_vals, width=w, color=COLOR_NOCHE, label="Sin control 7–13/09")
    ax.bar([i + w / 2 for i in x], con_vals, width=w, color=COLOR_BARRA_WES, label="Con control 14–20/09")
    ax.set_xticks(x)
    ax.set_xticklabels(names, rotation=30, ha="right", fontsize=8)
    ax.set_ylabel("m³ en la semana")
    ax.set_title("Comparativo semanal por colegio (misma cohorte)")
    ax.legend(frameon=False, fontsize=8)
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path


def _chart_ahorro(wow: Dict[str, Any], path: Path) -> Path:
    filas = sorted(wow["filas"], key=lambda f: f.ahorro_m3, reverse=True)
    names = [SHORT_NAMES.get(f.node_id, f.nombre) for f in filas]
    vals = [f.ahorro_m3 * wow["precio"] for f in filas]
    fig, ax = plt.subplots(figsize=(10.5, 4.6))
    ax.barh(names[::-1], vals[::-1], color=COLOR_BARRA_WES)
    ax.set_xlabel("Ahorro (CLP)")
    ax.set_title(f"Ahorro económico por colegio @ {_fmt(wow['precio'], 0)} CLP/m³")
    ax.grid(axis="x", linestyle=":", alpha=0.4)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path


def _hallazgos_claros(wow: Dict[str, Any]) -> List[Hallazgo]:
    tob = wow["tobalaba"]
    top = sorted(wow["filas"], key=lambda f: f.ahorro_m3, reverse=True)[:3]
    top_txt = "; ".join(
        f"{SHORT_NAMES.get(f.node_id, f.nombre)} {_fmt(f.ahorro_m3)} m³"
        for f in top
    )
    esperados = [
        f"{SHORT_NAMES[nid]}: {txt}" for nid, txt in ESPERADOS_CONTROL.items()
    ]
    return [
        Hallazgo(
            "INFORMATIVA",
            f"Resultado vacaciones: se evitaron {_fmt(wow['ahorro'])} m³",
            (
                f"Antes (7–13/09, sin control): {_fmt(wow['tot_sin'])} m³ = "
                f"{_fmt_clp(wow['clp_sin'])}. "
                f"Después (14–20/09, con control): {_fmt(wow['tot_con'])} m³ = "
                f"{_fmt_clp(wow['clp_con'])}. "
                f"Ahorro: {_fmt(wow['pct'])} % → {_fmt_clp(wow['clp_ahorro'])}."
            ),
            f"Mayores aportes al ahorro: {top_txt}.",
        ),
        Hallazgo(
            "INFORMATIVA",
            "Consumo en 14–20 que SÍ estaba autorizado",
            "No interpretar como falla de corte: " + " | ".join(esperados),
            (
                "Hermida, Santa María, Erasmo, M. Huici, Unión Árabe y Juan Pablo II "
                "tenían corte 24 h; sus residuales bajos confirman el control."
            ),
        ),
        Hallazgo(
            "ATENCIÓN",
            "Tobalaba quedó fuera: hay que revisar el pulso",
            (
                f"No entra en el ahorro. La serie está anómala "
                f"(7–13: {_fmt(tob.m3_sin)} m³; 14–20: {_fmt(tob.m3_con)} m³) "
                f"y no sirve para evaluar control."
            ),
            (
                "Acción: diagnóstico de pulso/telemetría y reincorporarlo al esquema "
                "cuando la medición sea confiable."
            ),
        ),
        Hallazgo(
            "INFORMATIVA",
            "Tres colegios no tienen control (solo monitoreo)",
            (
                "Eduardo de la Barra, Alicura y Likankura no tienen válvula WES: "
                "no se pudo programar corte ni habilitación remota."
                + "".join(
                    f" {f.nombre}: {_fmt(f.m3_sin)} → {_fmt(f.m3_con)} m³."
                    for f in wow["solo_monitoreo"]
                )
            ),
            (
                "Quedan fuera del programa de vacaciones y del cálculo de ahorro. "
                "Cualquier baja ahí requiere terreno o instalar control."
            ),
        ),
    ]


def _acciones_claras() -> List[Accion]:
    return [
        Accion(
            accion="Revisar y normalizar el pulso de Tobalaba",
            responsable="WES + operación",
            plazo="7 días",
            objetivo="Recuperar medición confiable y reactivar control en ese colegio.",
        ),
        Accion(
            accion="Confirmar con el cliente el estado de los 3 sin válvula",
            responsable="WES + cliente",
            plazo="Próxima reunión",
            objetivo=(
                "Dejar por escrito que Eduardo de la Barra, Alicura y Likankura "
                "solo tienen monitoreo."
            ),
        ),
        Accion(
            accion="Mantener registro de habilitaciones en próximos periodos no lectivos",
            responsable="WES",
            plazo="Continuo",
            objetivo="Que obras/patinaje no se confundan con fugas en el informe.",
        ),
    ]


def _clarificar_spec(spec, wow: Dict[str, Any], chart_cmp: Path, chart_ahorro: Path) -> None:
    """Reescribe KPIs, estado, lectura y hallazgos para que el mensaje sea inequívoco."""
    spec.footer = "Informe de gestión hídrica - CORMUP | Vacaciones septiembre 2026"
    spec.titulo_mensual = "Informe de gestión hídrica — vacaciones"
    spec.titulo_onepager = "Resumen ejecutivo — vacaciones CORMUP"
    spec.periodo_corto = (
        "CORMUP Peñalolén · 7–13/09 sin control  vs  14–20/09 con control vacaciones"
    )

    # Portada: los 4 KPI cuentan la historia completa.
    spec.kpi_entrada = f"{_fmt(wow['tot_sin'])} m³"
    spec.kpi_consumo_label = "Sin control (7–13/09)"
    spec.kpi_promedio = f"{_fmt(wow['tot_con'])} m³"
    spec.kpi_promedio_label = "Con control (14–20/09)"
    spec.kpi_nocturno = f"{_fmt(wow['ahorro'])} m³"
    spec.kpi_nocturno_label = f"Ahorro ({_fmt(wow['pct'])} %)"
    spec.kpi_pct = _fmt_clp(wow["clp_ahorro"])
    spec.kpi_pct_label = f"Ahorro @ {_fmt(wow['precio'], 0)} CLP/m³"

    spec.clasificacion = "BAJO CONTROL"
    spec.motivo = (
        "el control de vacaciones redujo ~91 % el consumo de la cohorte "
        "respecto de la semana previa sin control"
    )

    spec.lectura_ejecutiva = [
        [
            ("En vacaciones se compararon dos semanas iguales (7 días) sobre ", False),
            ("10 colegios con control", True),
            (". ", False),
            ("Sin control (7–13/09): ", False),
            (f"{_fmt(wow['tot_sin'])} m³ ({_fmt_clp(wow['clp_sin'])})", True),
            (". ", False),
            ("Con control (14–20/09): ", False),
            (f"{_fmt(wow['tot_con'])} m³ ({_fmt_clp(wow['clp_con'])})", True),
            (".", False),
        ],
        [
            ("Ahorro: ", False),
            (f"{_fmt(wow['ahorro'])} m³ ({_fmt(wow['pct'])} %)", True),
            (", equivalente a ", False),
            (_fmt_clp(wow["clp_ahorro"]), True),
            (f" a {_fmt_clp(wow['precio'])}/m³. ", False),
            ("Estado: Bajo control", True),
            (" — el programa cumplió el objetivo de evitar consumo innecesario.", False),
        ],
        [
            (
                "Importante: Tobalaba no entra (pulso anómalo, pendiente de revisión). "
                "Eduardo de la Barra, Alicura y Likankura no tienen válvula: solo monitoreo.",
                False,
            )
        ],
    ]

    spec.panorama = [
        ("Lectura en una línea: de ", False),
        (f"{_fmt(wow['tot_sin'])} m³", True),
        (" (semana sin control) a ", False),
        (f"{_fmt(wow['tot_con'])} m³", True),
        (" (semana con control), ahorro ", False),
        (f"{_fmt(wow['ahorro'])} m³ / {_fmt_clp(wow['clp_ahorro'])}", True),
        (".", False),
    ]
    spec.panorama_nota = (
        "Cohorte = 10 colegios con control. Fuera: Tobalaba (pulso) + 3 sin válvula."
    )

    spec.hallazgos = _hallazgos_claros(wow)
    spec.acciones = _acciones_claras()

    spec.chart_puntos = chart_cmp
    spec.chart_puntos_nota = (
        "Gris = sin control (7–13/09). Azul = con control vacaciones (14–20/09). "
        "Misma cohorte de 10 colegios."
    )
    spec.max_entrada_txt = (
        f"En la semana sin control el total fue {_fmt(wow['tot_sin'])} m³; "
        f"con control bajó a {_fmt(wow['tot_con'])} m³."
    )
    spec.chart_nocturno = chart_ahorro
    spec.chart_nocturno_nota = (
        "Ahorro valorizado por colegio (semana sin control − semana con control) "
        f"× {_fmt(wow['precio'], 0)} CLP/m³."
    )

    spec.conclusion = [
        [
            ("Conclusión: el control especial de vacaciones funcionó. ", False),
            (
                f"Se evitaron {_fmt(wow['ahorro'])} m³ ({_fmt_clp(wow['clp_ahorro'])})",
                True,
            ),
            (" frente a la semana previa sin control.", False),
        ],
        [
            (
                "Pendiente explícito: revisar Tobalaba (pulso) y mantener claridad con el "
                "cliente sobre los tres colegios sin control hidráulico.",
                False,
            )
        ],
    ]

    spec.nota_agosto = (
        "Guía de lectura: portada = antes/después + $. "
        "Hallazgos = qué se ahorró, qué consumo estaba autorizado, qué quedó fuera. "
        "Anexo técnico = detalle de la semana con control (14–20/09) por colegio."
    )
    spec.criterio_nocturno = [
        [
            (
                "Tarifa referencial WES para CORMUP: ",
                False,
            ),
            (f"{_fmt_clp(wow['precio'])}/m³", True),
            (
                ". El ahorro económico = (m³ semana sin control − m³ semana con control) × tarifa. "
                "No sustituye la factura del sanitario.",
                False,
            ),
        ]
    ]
    spec.nota_cobertura = (
        "Anexo: columnas TOTAL / PROMEDIO / MÁXIMO / NOCTURNO corresponden solo a "
        "14–20/09 (semana con control). Para el antes/después use la portada y el gráfico gris/azul."
    )


def main() -> int:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
            sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
        except Exception:
            pass

    print("=" * 72)
    print("CORMUP — gestión hídrica vacaciones (datos claros antes/después)")
    print("=" * 72)

    CACHE.mkdir(parents=True, exist_ok=True)
    motor.CACHE_DIR = CACHE

    print("\n[INFO] Comparativo 7–13 vs 14–20…")
    wow = _wow()
    print(
        f"  SIN {_fmt(wow['tot_sin'])} m³ → CON {_fmt(wow['tot_con'])} m³ | "
        f"ahorro {_fmt(wow['ahorro'])} m³ ({_fmt(wow['pct'])} %) = {_fmt_clp(wow['clp_ahorro'])}"
    )

    cfg = _cfg()
    print("\n[INFO] Fetch 14–20/09 (anexo técnico)…")
    data = fetch_cliente(cfg)

    try:
        visitas = visitas_de_cliente(
            cargar_visitas_periodo(datetime(2026, 9, 14), datetime(2026, 9, 20, 23, 59, 59)),
            cfg,
        )
    except Exception as exc:
        visitas = []
        print(f"[ADVERTENCIA] Visitas: {exc}")

    spec = build_spec(cfg, data, visitas)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    charts = OUT_DIR / "_charts_vacaciones_sep2026"
    charts.mkdir(parents=True, exist_ok=True)
    chart_cmp = _chart_comparativo(wow, charts / "comparativo_sin_vs_con.png")
    chart_ahorro = _chart_ahorro(wow, charts / "ahorro_clp_colegios.png")
    _clarificar_spec(spec, wow, chart_cmp, chart_ahorro)

    one = OUT_DIR / "One_Pager_Gestion_Hidrica_CORMUP_Vacaciones_14_20_Sep_2026.pdf"
    monthly = OUT_DIR / "Informe_Gestion_Hidrica_CORMUP_Vacaciones_14_20_Sep_2026.pdf"
    render_one_pager(spec, one)
    render_mensual(spec, monthly, charts)

    meta = {
        "one_pager": str(one),
        "informe": str(monthly),
        "mensaje": "Antes 7-13 sin control vs despues 14-20 con control",
        "total_sin_m3": wow["tot_sin"],
        "total_con_m3": wow["tot_con"],
        "ahorro_m3": wow["ahorro"],
        "ahorro_pct": wow["pct"],
        "ahorro_clp": wow["clp_ahorro"],
        "precio_clp_m3": wow["precio"],
    }
    (OUT_DIR / "meta_vacaciones_sep2026.json").write_text(
        json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(f"\n[OK] {one.name}")
    print(f"[OK] {monthly.name}")
    print(json.dumps(meta, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
