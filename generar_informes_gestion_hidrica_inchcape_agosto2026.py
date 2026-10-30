"""
One-pager y informe mensual de gestión hídrica — Inchcape Quilicura, agosto 2026.

Formato Zapallar, con texto blanco en el cuadro de hallazgos y
PRIORIDAD (SEGUIMIENTO) en una sola línea.
"""

from __future__ import annotations

import json
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Optional

from generar_datos_gestion_hidrica_inchcape import OUT as DATOS_JSON, main as fetch_datos
from informe_gestion_hidrica_pdf import (
    Accion,
    Hallazgo,
    InformeSpec,
    PuntoIndicador,
    SerieDiaria,
    _fecha_es,
    _fmt,
    _fmt_clp,
    build_chart_6_meses,
    build_chart_nocturno,
    build_chart_puntos,
    render_mensual,
    render_one_pager,
    resolve_logo,
)

COMPANY_FOLDER = Path("reports/Inchcape/GESTION_HIDRICA")
CHART_DIR = COMPANY_FOLDER / "_charts"
PERIODO_DIAS = 28
MATRIZ_NAME = "Matriz Principal"


def _load_datos() -> dict:
    if not DATOS_JSON.is_file():
        fetch_datos()
    return json.loads(DATOS_JSON.read_text(encoding="utf-8"))


def _daily_full(nodo: dict) -> SerieDiaria:
    by = {d["date"]: float(d["m3"]) for d in nodo.get("daily") or []}
    fechas = []
    valores: List[Optional[float]] = []
    cur = datetime(2026, 8, 1)
    end = datetime(2026, 8, 28)
    while cur <= end:
        key = cur.strftime("%Y-%m-%d")
        fechas.append(cur)
        valores.append(by[key] if key in by else None)
        cur += timedelta(days=1)
    return SerieDiaria(
        nombre=nodo["short_name"],
        fechas=fechas,
        valores=valores,
        lectura="",
        es_matriz=nodo["node_id"] == "000012-06",
    )


def build_spec(data: dict) -> InformeSpec:
    kpi = data["kpi"]
    nodos = data["nodos"]
    matriz = next(n for n in nodos if n["node_id"] == "000012-06")
    camarines = next(n for n in nodos if n["node_id"] == "000012-11")
    entrada = float(kpi["entrada"])
    nocturno = float(kpi["nocturno"])
    pct = float(kpi["pct_nocturno"])
    promedio = entrada / PERIODO_DIAS
    max_m3 = float(kpi["max_m3"])
    max_fecha = kpi["max_fecha"]
    costo = float(kpi["costo_nocturno"])
    pct_txt = f"{_fmt(round(pct), 0)} %"
    entrada_txt = f"{_fmt(entrada, 1)} m³"
    noct_txt = f"{_fmt(nocturno, 1)} m³"
    prom_txt = f"{_fmt(promedio, 1)} m³/día"

    hallazgos = [
        Hallazgo(
            prioridad="INFORMATIVA",
            titulo="10 % del consumo nocturno es operación WES",
            detalle=f"{_fmt(nocturno, 1)} m³ entre 00:00 y 06:59.",
            lectura=(
                "El caudal de madrugada de la Matriz Principal se explica por los ciclos "
                "de control y regulación del sistema WES, no por una pérdida de la sucursal. "
                "Se mantiene como referencia operativa."
            ),
        ),
        Hallazgo(
            prioridad="SEGUIMIENTO",
            titulo="Evento en Camarines",
            detalle=f"{_fmt(camarines['max_m3'], 1)} m³ el {_fecha_es(camarines['max_fecha'])}.",
            lectura=(
                "El punto interno registró un peak claramente por sobre su promedio. "
                "Corresponde confirmar si fue una limpieza o uso puntual. Si se repite "
                "sin explicación, elevar el estado a Requiere atención."
            ),
        ),
        Hallazgo(
            prioridad="INFORMATIVA",
            titulo="Matriz Principal con cobertura desde el 5 de agosto",
            detalle="La serie de entrada tiene 24 días con datos.",
            lectura=(
                "No hay registros de la matriz entre el 1 y el 4 de agosto; el 5 de agosto "
                "aparece un día parcial. No se extrapola. Los puntos internos sí cubren el mes."
            ),
        ),
    ]

    acciones = [
        Accion(
            accion="Mantener el 10 % nocturno como línea base WES.",
            responsable="WES",
            plazo="Próximo informe",
            objetivo="Detectar desviaciones sobre la operación habitual.",
        ),
        Accion(
            accion="Revisar el evento de Camarines del 9 de agosto.",
            responsable="Operación + WES",
            plazo="7 días",
            objetivo="Validar causa y descartar fuga o dato anómalo.",
        ),
        Accion(
            accion="Completar cobertura de Matriz Principal y Camarines.",
            responsable="WES",
            plazo="7 días",
            objetivo="Asegurar serie continua para el próximo periodo.",
        ),
    ]

    clasificacion = "BAJO CONTROL"
    motivo = (
        "el 10 % nocturno de la Matriz Principal corresponde al funcionamiento del "
        "sistema WES y no se identifica pérdida en la sucursal."
    )

    panorama = [
        ("La sucursal de Quilicura registró ", False),
        (entrada_txt, True),
        (" de entrada, con un promedio de ", False),
        (f"{_fmt(promedio, 1)} m³ diarios", True),
        (" sobre los 28 días del periodo. El mayor consumo de la matriz se produjo el ", False),
        (_fecha_es(max_fecha), True),
        (", con ", False),
        (f"{_fmt(max_m3, 1)} m³", True),
        (".", False),
    ]

    conclusion = [
        [
            ("El estado se clasifica como ", False),
            ("“Bajo control”", True),
            (". El ", False),
            (f"{_fmt(round(pct), 0)} %", True),
            (" de participación nocturna de la Matriz Principal se explica por el "
             "funcionamiento del sistema WES (ciclos de control y regulación), no por un "
             "consumo anómalo de la sucursal. Los demás medidores son puntos internos de "
             "esa matriz y ", False),
            ("no se suman", True),
            (" al total.", False),
        ],
        [
            ("El peak de Camarines (", False),
            (f"{_fmt(camarines['max_m3'], 1)} m³ el {_fecha_es(camarines['max_fecha'])}", True),
            (") debe validarse operativamente. La matriz no registra los días 1 al 4 de "
             "agosto; esa ausencia no se extrapola. Si el nocturno se desvía de esta "
             "referencia o aparecen eventos sin explicación, el estado avanzará a ", False),
            ("“Requiere atención”", True),
            (".", False),
        ],
    ]

    lectura = [
        [
            ("Durante el periodo analizado la Matriz Principal de Inchcape Quilicura "
             "registró ", False),
            (entrada_txt, True),
            (". El consumo entre las 00:00 y las 06:59 alcanzó ", False),
            (noct_txt, True),
            (", equivalente al ", False),
            (f"{_fmt(round(pct), 0)} %", True),
            (" del volumen de entrada y a un costo referencial de ", False),
            (_fmt_clp(costo), True),
            (".", False),
        ],
        [
            ("Ese caudal de madrugada corresponde al funcionamiento del sistema WES "
             "y no se interpreta como pérdida. Por ello el estado se clasifica como ", False),
            ("Bajo control", True),
            (". Los puntos internos (Dercomaq, Lav. Máquinas, Casino, Proderco, "
             "Camarines y Edificio JCB) miden derivaciones de la misma matriz: ", False),
            ("no se suman al total de la sucursal", True),
            (".", False),
        ],
    ]

    indicadores = []
    for n in sorted(nodos, key=lambda x: -float(x["total"])):
        max_dt = datetime.strptime(n["max_fecha"], "%Y-%m-%d") if n.get("max_fecha") else None
        indicadores.append(
            PuntoIndicador(
                nombre=n["short_name"],
                total=float(n["total"]),
                promedio=float(n["total"]) / PERIODO_DIAS,
                max_m3=float(n["max_m3"]),
                max_fecha=max_dt.strftime("%d/%m") if max_dt else "—",
                nocturno=float(n["nocturno_m3"]),
                cobertura=int(n["nocturno_cobertura"]),
                es_matriz=n["node_id"] == "000012-06",
            )
        )

    serie_matriz = _daily_full(matriz)
    serie_matriz.lectura = (
        f"La entrada comienza el 5 de agosto (sin datos el 1 al 4). El máximo fue de "
        f"{_fmt(max_m3, 1)} m³ el {_fecha_es(max_fecha)}; debe contrastarse con la "
        "operación de ese día. El 5 de agosto registra un día parcial."
    )
    serie_cam = _daily_full(camarines)
    serie_cam.lectura = (
        f"El 9 de agosto se registraron {_fmt(camarines['max_m3'], 1)} m³, muy por sobre "
        "el promedio del punto, con otro alza el 11 de agosto. Hacia el final del mes "
        "la serie se interrumpe: corresponde validar causa operacional y cobertura."
    )
    casino = next(n for n in nodos if n["node_id"] == "000012-09")
    serie_casino = _daily_full(casino)
    serie_casino.lectura = (
        "El Casino muestra un patrón de días hábiles con fin de semana en cero, "
        "coherente con operación de sucursal y no con una pérdida continua."
    )

    labels_6 = []
    vals_6 = []
    for item in data["serie_6_meses"]:
        lab = item["label"].split()[0].capitalize()
        if item["label"].endswith("*"):
            lab += "*"
        labels_6.append(lab)
        vals_6.append(float(item["m3"]))

    names = [n["short_name"] for n in nodos]
    totals = [float(n["total"]) for n in nodos]
    nocts = [float(n["nocturno_m3"]) for n in nodos]

    CHART_DIR.mkdir(parents=True, exist_ok=True)
    chart_6m = build_chart_6_meses(CHART_DIR / "inchcape_6m.png", labels_6, vals_6)
    chart_pts = build_chart_puntos(CHART_DIR / "inchcape_puntos.png", names, totals, MATRIZ_NAME)
    chart_noc = build_chart_nocturno(CHART_DIR / "inchcape_nocturno.png", names, nocts, MATRIZ_NAME)

    return InformeSpec(
        cliente="Inchcape",
        sitio="Inchcape Quilicura",
        periodo_corto="Inchcape Quilicura · 1 al 28 de agosto de 2026",
        footer="Informe mensual - Inchcape | Agosto 2026",
        titulo_onepager="Resumen ejecutivo de gestión hídrica",
        titulo_mensual="Informe mensual de gestión hídrica",
        clasificacion=clasificacion,
        motivo=motivo,
        kpi_entrada=entrada_txt,
        kpi_promedio=prom_txt,
        kpi_nocturno=noct_txt,
        kpi_pct=pct_txt,
        panorama=panorama,
        panorama_nota=(
            "Agosto comprende 28 días y no se extrapola ni se compara directamente con meses completos."
        ),
        hallazgos=hallazgos,
        acciones=acciones,
        conclusion=conclusion,
        lectura_ejecutiva=lectura,
        nota_agosto=(
            "* Agosto comprende 28 días; por ello no se compara directamente con meses completos. "
            "No se extrapola el consumo. La Matriz Principal tiene 24 días con datos."
        ),
        chart_6m=chart_6m,
        chart_puntos=chart_pts,
        chart_puntos_nota=(
            "la Matriz Principal representa el consumo real de la sucursal. Los demás "
            "medidores están aguas abajo y miden derivaciones del mismo caudal; por lo tanto, "
            "sus consumos no se suman."
        ),
        max_entrada_txt=(
            f"El mayor consumo diario de la entrada ocurrió el {_fecha_es(max_fecha)}, "
            f"con {_fmt(max_m3, 1)} m³."
        ),
        chart_nocturno=chart_noc,
        chart_nocturno_nota=(
            "Los porcentajes y volúmenes de los medidores interiores se expresan respecto "
            "de la Matriz Principal, pero no son aditivos porque corresponden a derivaciones "
            "internas de la sucursal."
        ),
        indicadores=indicadores,
        criterio_nocturno=[
            [
                (
                    "Se considera nocturno el volumen medido entre las 00:00 y las 06:59, "
                    "hora de Chile. Los valores corresponden únicamente a días con datos y no "
                    "se proyectan. El costo nocturno de la entrada principal se estima en ",
                    False,
                ),
                (_fmt_clp(costo), True),
                (", utilizando la tarifa referencial de ", False),
                (f"{_fmt_clp(float(data['price_per_m3']))}/m³", True),
                (".", False),
            ]
        ],
        nota_cobertura=(
            "La cobertura nocturna indica cuántos días cuentan con registros en esa franja. "
            "En la Matriz Principal faltan el 1 al 4 de agosto; en Camarines la serie se "
            "interrumpe hacia el cierre del mes. Los días sin datos no se interpolan."
        ),
        series_diarias=[serie_matriz, serie_cam, serie_casino],
        logo_path=resolve_logo(),
    )


def main() -> None:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
            sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
        except Exception:
            pass
    print("[INFO] Generando informes de gestión hídrica Inchcape agosto 2026", flush=True)
    data = _load_datos()
    spec = build_spec(data)
    one = COMPANY_FOLDER / "One_Pager_Gestion_Hidrica_Inchcape_Agosto_2026.pdf"
    monthly = COMPANY_FOLDER / "Informe_Mensual_Inchcape_Agosto_2026.pdf"
    render_one_pager(spec, one)
    render_mensual(spec, monthly, CHART_DIR)
    print(f"[OK] {one}", flush=True)
    print(f"[OK] {monthly}", flush=True)


if __name__ == "__main__":
    main()
