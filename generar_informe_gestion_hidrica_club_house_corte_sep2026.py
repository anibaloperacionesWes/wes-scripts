"""
Informe de gestión hídrica — Club House CDUC (000021-01).

Corte total programado: 17/09/2026 15:00 → 29/09/2026 06:00 (Chile).
Mismo formato que CORMUP vacaciones (Zapallar / gestión hídrica).

Uso:
  python generar_informe_gestion_hidrica_club_house_corte_sep2026.py
"""

from __future__ import annotations

import json
import sys
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Tuple
from zoneinfo import ZoneInfo

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

from control_nocturno import obtener_datos_horarios_dia
from informe_gestion_hidrica_pdf import (
    Accion,
    Hallazgo,
    InformeSpec,
    PuntoIndicador,
    render_mensual,
    render_one_pager,
    resolve_logo,
)
from wes_estilo_graficos_app import COLOR_BARRA_FACT, COLOR_BARRA_WES, COLOR_NOCHE

CHILE = ZoneInfo("America/Santiago")
ROOT = Path(__file__).resolve().parent
NODE_ID = "000021-01"
NOMBRE = "Club House CDUC"
CLIENTE = "CDUC"
OUT_DIR = ROOT / "reports" / "CDUC" / "GESTION_HIDRICA"
TARIFA_CLP_M3 = 1400.0  # misma tarifa referencial del informe CORMUP vacaciones

CORTE_INICIO = datetime(2026, 9, 17, 15, 0)  # programado
CORTE_FIN_PROG = datetime(2026, 9, 29, 6, 0)  # reposición indicada
# Primer caudal tras el cero (telemetría).
REPOSICION_OBS = datetime(2026, 9, 20, 8, 0)

SERIE_DESDE = datetime(2026, 9, 1, 0, 0)
NOCTURNO_HORAS = range(0, 7)


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


def _hora_cierre_informe() -> datetime:
    ahora = datetime.now(CHILE).replace(tzinfo=None, second=0, microsecond=0)
    cierre = ahora.replace(minute=0)
    if ahora.minute == 0:
        cierre -= timedelta(hours=1)
    if cierre > CORTE_FIN_PROG:
        return CORTE_FIN_PROG
    return cierre


def _fetch_hourly(desde: datetime, hasta: datetime) -> Dict[datetime, Dict[int, float]]:
    out: Dict[datetime, Dict[int, float]] = {}
    d = desde.replace(hour=0, minute=0, second=0, microsecond=0)
    last = hasta.replace(hour=0, minute=0, second=0, microsecond=0)
    while d <= last:
        out[d] = obtener_datos_horarios_dia(NODE_ID, d)
        d += timedelta(days=1)
    return out


def _m3_hora(serie: Dict[datetime, Dict[int, float]], ts: datetime) -> float:
    dia = ts.replace(hour=0, minute=0, second=0, microsecond=0)
    return float(serie.get(dia, {}).get(ts.hour, 0.0) or 0.0)


def _suma(serie: Dict[datetime, Dict[int, float]], t0: datetime, t1: datetime) -> Tuple[float, int]:
    """Suma m³ en [t0, t1) por hora Chile."""
    tot = 0.0
    n = 0
    t = t0.replace(minute=0, second=0, microsecond=0)
    fin = t1.replace(minute=0, second=0, microsecond=0)
    while t < fin:
        tot += _m3_hora(serie, t)
        n += 1
        t += timedelta(hours=1)
    return tot, n


def _diario(serie: Dict[datetime, Dict[int, float]], dia: datetime) -> float:
    return sum(float(serie.get(dia, {}).get(h, 0.0) or 0.0) for h in range(24))


def _nocturno(serie: Dict[datetime, Dict[int, float]], t0: datetime, t1: datetime) -> float:
    tot = 0.0
    t = t0.replace(minute=0, second=0, microsecond=0)
    fin = t1.replace(minute=0, second=0, microsecond=0)
    while t < fin:
        if t.hour in NOCTURNO_HORAS:
            tot += _m3_hora(serie, t)
        t += timedelta(hours=1)
    return tot


@dataclass
class Metricas:
    cierre: datetime
    serie: Dict[datetime, Dict[int, float]] = field(repr=False)
    m3_cero: float = 0.0
    h_cero: int = 0
    m3_base: float = 0.0
    h_base: int = 0
    m3_desde_repo: float = 0.0
    h_desde_repo: int = 0
    m3_ventana_prog_hasta_cierre: float = 0.0
    h_ventana_prog_hasta_cierre: int = 0
    m3_17_antes: float = 0.0
    dias: List[datetime] = field(default_factory=list)

    @property
    def ahorro_m3(self) -> float:
        return self.m3_base - self.m3_cero

    @property
    def ahorro_pct(self) -> float:
        if self.m3_base <= 1e-9:
            return 0.0
        return 100.0 * self.ahorro_m3 / self.m3_base

    @property
    def clp_ahorro(self) -> float:
        return self.ahorro_m3 * TARIFA_CLP_M3

    @property
    def clp_base(self) -> float:
        return self.m3_base * TARIFA_CLP_M3

    @property
    def clp_cero(self) -> float:
        return self.m3_cero * TARIFA_CLP_M3

    @property
    def clp_desde_repo(self) -> float:
        return self.m3_desde_repo * TARIFA_CLP_M3


def calcular(serie: Dict[datetime, Dict[int, float]], cierre: datetime) -> Metricas:
    m = Metricas(cierre=cierre, serie=serie)
    m.m3_cero, m.h_cero = _suma(serie, CORTE_INICIO, REPOSICION_OBS)
    base_fin = CORTE_INICIO
    base_ini = base_fin - timedelta(hours=m.h_cero)
    m.m3_base, m.h_base = _suma(serie, base_ini, base_fin)
    repo_fin = min(cierre, CORTE_FIN_PROG)
    if repo_fin > REPOSICION_OBS:
        m.m3_desde_repo, m.h_desde_repo = _suma(serie, REPOSICION_OBS, repo_fin)
    m.m3_ventana_prog_hasta_cierre, m.h_ventana_prog_hasta_cierre = _suma(
        serie, CORTE_INICIO, min(cierre, CORTE_FIN_PROG)
    )
    m.m3_17_antes = sum(_m3_hora(serie, datetime(2026, 9, 17, h)) for h in range(15))
    d = SERIE_DESDE
    last = cierre.replace(hour=0, minute=0, second=0, microsecond=0)
    while d <= last:
        m.dias.append(d)
        d += timedelta(days=1)
    return m


def _chart_diario(m: Metricas, path: Path) -> Path:
    labels = [d.strftime("%d/%m") for d in m.dias]
    vals = [_diario(m.serie, d) for d in m.dias]
    colors = []
    for d in m.dias:
        if d.date() < CORTE_INICIO.date():
            colors.append(COLOR_NOCHE)
        elif d.date() == CORTE_INICIO.date():
            colors.append("#8E9AA0")
        elif d.date() < REPOSICION_OBS.date():
            colors.append(COLOR_BARRA_WES)
        else:
            colors.append(COLOR_BARRA_FACT)
    fig, ax = plt.subplots(figsize=(10.5, 4.4))
    ax.bar(labels, vals, color=colors, width=0.72)
    ax.set_ylabel("m³ / día")
    ax.set_title("Club House CDUC — consumo diario (1–21/09/2026)")
    ax.legend(
        handles=[
            Patch(color=COLOR_NOCHE, label="Antes del corte total"),
            Patch(color="#8E9AA0", label="17/09 (corte 15:00)"),
            Patch(color=COLOR_BARRA_WES, label="Corte efectivo (18–19/09 = 0 m³)"),
            Patch(color=COLOR_BARRA_FACT, label="Después de la reposición (20/09 ~08:00)"),
        ],
        frameon=False,
        fontsize=7.5,
        loc="upper left",
    )
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    ymax = max(vals) if vals else 1.0
    ax.set_ylim(0, ymax * 1.18)
    for label, val, d in zip(labels, vals, m.dias):
        if val >= 0.05:
            ax.text(
                label,
                val + ymax * 0.015,
                _fmt(val, 1),
                ha="center",
                va="bottom",
                fontsize=6.5,
                color="#20313D",
            )
        elif d.date() in (datetime(2026, 9, 18).date(), datetime(2026, 9, 19).date()):
            ax.text(label, ymax * 0.02, "0", ha="center", va="bottom", fontsize=6.5, color="#087EAE")
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path


def _chart_antes_despues(m: Metricas, path: Path) -> Path:
    fig, ax = plt.subplots(figsize=(10.5, 4.2))
    cats = [
        f"Sin corte\n({m.h_base} h previas)",
        f"Con corte total\n({m.h_cero} h a cero)",
        f"Tras reposición\n({m.h_desde_repo} h al cierre)",
    ]
    vals = [m.m3_base, m.m3_cero, m.m3_desde_repo]
    cols = [COLOR_NOCHE, COLOR_BARRA_WES, COLOR_BARRA_FACT]
    bars = ax.bar(cats, vals, color=cols, width=0.55)
    ax.set_ylabel("m³")
    ax.set_title(
        f"Antes / durante / después del corte total  @ {_fmt(TARIFA_CLP_M3, 0)} CLP/m³"
    )
    ax.grid(axis="y", linestyle=":", alpha=0.4)
    ymax = max(vals + [1.0])
    ax.set_ylim(0, ymax * 1.28)
    for bar, val in zip(bars, vals):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            val + ymax * 0.04,
            f"{_fmt(val, 1)} m³\n{_fmt_clp(val * TARIFA_CLP_M3)}",
            ha="center",
            va="bottom",
            fontsize=8,
            color="#20313D",
        )
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path


def _max_dia_en(m: Metricas, t0: datetime, t1: datetime) -> Tuple[float, str]:
    best = 0.0
    best_d = t0
    d = t0.replace(hour=0, minute=0, second=0, microsecond=0)
    last = (t1 - timedelta(seconds=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    while d <= last:
        v = _diario(m.serie, d)
        # En el 17 solo cuenta desde las 15:00; en el 20 solo hasta las 08:00.
        if d.date() == CORTE_INICIO.date():
            v = sum(_m3_hora(m.serie, datetime(d.year, d.month, d.day, h)) for h in range(15, 24))
        elif d.date() == REPOSICION_OBS.date():
            v = sum(_m3_hora(m.serie, datetime(d.year, d.month, d.day, h)) for h in range(REPOSICION_OBS.hour))
        if v > best:
            best = v
            best_d = d
        d += timedelta(days=1)
    return best, f"{best_d.day:02d}/{best_d.month:02d}"


def _build_spec(m: Metricas, chart_diario: Path, chart_cmp: Path) -> InformeSpec:
    cierre_txt = m.cierre.strftime("%d/%m/%Y %H:%M")
    base_ini = CORTE_INICIO - timedelta(hours=m.h_base)
    max_m3, max_f = _max_dia_en(m, CORTE_INICIO, REPOSICION_OBS)
    noc = _nocturno(m.serie, CORTE_INICIO, REPOSICION_OBS)
    cobertura = 4  # 17 (tarde), 18, 19 y 20 hasta 08:00, todos con serie horaria

    hallazgos = [
        Hallazgo(
            "INFORMATIVA",
            f"El corte total dejó el Club House en 0,0 m³ durante {m.h_cero} h",
            (
                f"Programado el jueves 17/09 a las 15:00. Desde esa hora hasta las 07:00 del "
                f"domingo 20/09 el medidor {NODE_ID} registró {_fmt(m.m3_cero, 1)} m³ "
                f"({m.h_cero} horas). El 18 y el 19 de septiembre (Fiestas Patrias) fueron "
                f"cero las 24 h."
            ),
            (
                f"En las {m.h_base} horas inmediatamente anteriores "
                f"({base_ini.strftime('%d/%m %H:%M')} a {CORTE_INICIO.strftime('%d/%m %H:%M')}) "
                f"el mismo punto consumió {_fmt(m.m3_base, 1)} m³. Ahorro del tramo a cero: "
                f"{_fmt(m.ahorro_m3, 1)} m³ ({_fmt(m.ahorro_pct, 1)} %) = {_fmt_clp(m.clp_ahorro)}."
            ),
        ),
        Hallazgo(
            "ATENCIÓN",
            "La reposición se observa el 20/09 ~08:00, no el 29/09 06:00",
            (
                "Lo indicado para este informe: corte total desde el 17/09 15:00 y suministro "
                "de vuelta el 29/09 a las 06:00. La telemetría muestra el primer caudal el "
                f"domingo 20/09 a las 08:00 ({_fmt(_m3_hora(m.serie, REPOSICION_OBS), 2)} m³/h) "
                f"y {_fmt(m.m3_desde_repo, 1)} m³ entre esa hora y el cierre del informe "
                f"({cierre_txt})."
            ),
            (
                "Conviene confirmar con operación si la fecha de reapertura era el 20/09 o el "
                "29/09. Si el corte debía mantenerse hasta el 29/09 06:00, el consumo del 20 y "
                "21/09 está dentro de la ventana programada."
            ),
        ),
        Hallazgo(
            "INFORMATIVA",
            "No hay consumo nocturno (00:00–06:59) en lo que va de septiembre",
            (
                "Entre el 1 y el 21/09 el Club House no registra m³ en madrugada. El residual "
                "del recinto es diurno. Eso respalda que el cero del 17–20/09 es corte de "
                "válvula, no un medidor caído: el 17/09 de 07:00 a 13:00 aún hubo "
                f"{_fmt(m.m3_17_antes, 1)} m³ y después de las 15:00 quedó en cero."
            ),
            "La red entra en cero cuando se programa el corte; no aparece fuga nocturna de fondo.",
        ),
        Hallazgo(
            "INFORMATIVA",
            "Tras la reapertura el recinto volvió a consumir en horario diurno",
            (
                f"20/09: {_fmt(_diario(m.serie, datetime(2026, 9, 20)), 1)} m³ (desde las 08:00). "
                f"21/09 (parcial al cierre): {_fmt(_diario(m.serie, datetime(2026, 9, 21)), 1)} m³. "
                f"Valorizado al cierre: {_fmt_clp(m.clp_desde_repo)}."
            ),
            "Patrón coherente con uso del Club House, no con un pulso anómalo.",
        ),
    ]

    acciones = [
        Accion(
            accion="Confirmar la fecha/hora real de reposición programada (20/09 vs 29/09 06:00)",
            responsable="WES + CDUC",
            plazo="48 h",
            objetivo="Dejar por escrito el calendario de corte total y evitar lecturas cruzadas.",
        ),
        Accion(
            accion=(
                "Si el corte debía durar hasta el 29/09, revisar quién habilitó el "
                "suministro el 20/09 ~08:00"
            ),
            responsable="WES + operación Club House",
            plazo="Esta semana",
            objetivo="Determinar si fue reapertura acordada o habilitación anticipada.",
        ),
        Accion(
            accion="Registrar en bitácora los próximos cortes totales (inicio, fin y excepciones)",
            responsable="WES",
            plazo="Continuo",
            objetivo="Que una reapertura no se confunda con fuga ni con pérdida de control.",
        ),
    ]

    clasificacion = "EN OBSERVACIÓN"
    motivo = (
        "el corte total funcionó (0 m³ por 65 h), pero la reposición telemedida "
        "es el 20/09 y no el 29/09 06:00 indicado"
    )

    spec = InformeSpec(
        cliente=CLIENTE,
        sitio=NOMBRE,
        periodo_corto=(
            "Club House CDUC · corte total 17/09 15:00 → reposición indicada 29/09 06:00 · "
            f"telemetría al {cierre_txt}"
        ),
        footer="Informe de gestión hídrica - CDUC Club House | Corte total septiembre 2026",
        titulo_onepager="Resumen ejecutivo — corte total Club House",
        titulo_mensual="Informe de gestión hídrica — corte total",
        clasificacion=clasificacion,
        motivo=motivo,
        kpi_entrada=f"{_fmt(m.m3_base)} m³",
        kpi_consumo_label=f"Sin corte ({m.h_base} h previas)",
        kpi_promedio=f"{_fmt(m.m3_cero)} m³",
        kpi_promedio_label=f"Con corte ({m.h_cero} h a cero)",
        kpi_nocturno=f"{_fmt(m.ahorro_m3)} m³",
        kpi_nocturno_label=f"Ahorro ({_fmt(m.ahorro_pct)} %)",
        kpi_pct=_fmt_clp(m.clp_ahorro),
        kpi_pct_label=f"Ahorro @ {_fmt(TARIFA_CLP_M3, 0)} CLP/m³",
        panorama=[
            ("Lectura en una línea: de ", False),
            (f"{_fmt(m.m3_base)} m³", True),
            (f" en las {m.h_base} h previas al corte, a ", False),
            (f"{_fmt(m.m3_cero)} m³", True),
            (f" durante {m.h_cero} h de corte total. Ahorro ", False),
            (f"{_fmt(m.ahorro_m3)} m³ / {_fmt_clp(m.clp_ahorro)}", True),
            (". La reposición se ve el 20/09 ~08:00 (indicada: 29/09 06:00).", False),
        ],
        panorama_nota=(
            f"Punto {NODE_ID}. Tarifa referencial {_fmt_clp(TARIFA_CLP_M3)}/m³ "
            "(misma del informe CORMUP vacaciones). No sustituye la factura del sanitario."
        ),
        hallazgos=hallazgos,
        acciones=acciones,
        conclusion=[
            [
                ("Conclusión: el corte total del Club House se ejecutó. ", False),
                (
                    f"Se evitaron {_fmt(m.ahorro_m3)} m³ ({_fmt_clp(m.clp_ahorro)})",
                    True,
                ),
                (
                    " frente a las horas equivalentes inmediatamente anteriores, con la red en "
                    "cero el 18 y 19/09.",
                    False,
                ),
            ],
            [
                (
                    "Pendiente: confirmar si la reposición programada era el 20/09 o el "
                    f"29/09 06:00. Telemetría: caudal desde el 20/09 08:00 "
                    f"({_fmt(m.m3_desde_repo)} m³ al cierre).",
                    False,
                )
            ],
        ],
        lectura_ejecutiva=[
            [
                ("Se programó corte total en ", False),
                ("Club House CDUC (000021-01)", True),
                (" el jueves ", False),
                ("17/09/2026 a las 15:00", True),
                (", con reposición indicada el ", False),
                ("29/09/2026 a las 06:00", True),
                (". ", False),
                ("El 17/09 de 07:00 a 13:00 aún hubo consumo ", False),
                (f"({_fmt(m.m3_17_antes)} m³)", True),
                ("; desde las 15:00 el medidor quedó en cero.", False),
            ],
            [
                ("Ventana efectiva a cero: ", False),
                ("17/09 15:00 – 20/09 08:00", True),
                (f" → {_fmt(m.m3_cero)} m³ en {m.h_cero} h. ", False),
                ("Periodo equivalente previo: ", False),
                (f"{_fmt(m.m3_base)} m³", True),
                (f" ({_fmt_clp(m.clp_base)}). Ahorro: ", False),
                (f"{_fmt(m.ahorro_pct)} % → {_fmt_clp(m.clp_ahorro)}", True),
                (".", False),
            ],
            [
                ("Estado: En observación", True),
                (
                    " — el control hidráulico del corte funcionó; hay que cuadrar la fecha "
                    "de reposición (20/09 telemedida vs 29/09 indicada).",
                    False,
                ),
            ],
        ],
        nota_agosto=(
            "Cómo leer este informe: la portada es el antes/después del tramo en que el "
            "medidor estuvo en cero (17/09 15:00 a 20/09 08:00), valorizado a "
            f"{_fmt(TARIFA_CLP_M3, 0)} CLP/m³. El gráfico diario muestra también la "
            "reposición del 20/09. El anexo técnico detalla solo la ventana a cero."
        ),
        chart_6m=None,
        chart_puntos=chart_diario,
        chart_puntos_nota=(
            "Rojo = días previos al corte. Gris = 17/09 (corte a las 15:00). "
            "Azul = 18–19/09 en cero. Naranja = desde el 20/09, cuando volvió el caudal."
        ),
        max_entrada_txt=(
            f"En las {m.h_base} h previas el Club House registró {_fmt(m.m3_base)} m³; "
            f"en las {m.h_cero} h de corte efectivo, {_fmt(m.m3_cero)} m³."
        ),
        chart_nocturno=chart_cmp,
        chart_nocturno_nota=(
            "Tres bloques de igual lectura operativa: horas equivalentes antes del corte, "
            "horas en cero, y horas desde la reposición observada hasta el cierre del informe."
        ),
        indicadores=[
            PuntoIndicador(
                nombre=NOMBRE,
                total=m.m3_cero,
                promedio=(m.m3_cero / (m.h_cero / 24.0)) if m.h_cero else 0.0,
                max_m3=max_m3,
                max_fecha=max_f,
                nocturno=noc,
                cobertura=cobertura,
                es_matriz=True,
            )
        ],
        criterio_nocturno=[
            [
                ("Tarifa referencial WES (igual a CORMUP vacaciones): ", False),
                (f"{_fmt_clp(TARIFA_CLP_M3)}/m³", True),
                (
                    ". El ahorro económico = (m³ del tramo equivalente previo − m³ del tramo "
                    "en cero) × tarifa. No sustituye la factura del sanitario.",
                    False,
                ),
            ],
            [
                (
                    "Nocturno Club House: horas Chile 00:00–06:59 (criterio operativo WES "
                    "fuera de colegios CORMUP). En esta serie el nocturno es 0,0 m³.",
                    False,
                )
            ],
        ],
        nota_cobertura=(
            "Anexo: TOTAL / PROMEDIO / MÁXIMO / NOCTURNO corresponden solo a la ventana "
            "en cero (17/09 15:00–20/09 08:00). Para el antes/después y la reposición "
            "use la portada y el gráfico diario."
        ),
        series_diarias=[],
        logo_path=resolve_logo(),
        visitas=[],
    )
    return spec


def main() -> int:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
            sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
        except Exception:
            pass

    cierre = _hora_cierre_informe()
    print("=" * 72)
    print("CDUC Club House — gestión hídrica corte total (17/09 15:00 → 29/09 06:00)")
    print(f"Cierre de datos: {cierre:%Y-%m-%d %H:%M} Chile")
    print("=" * 72)

    print("\n[INFO] Descargando serie horaria 000021-01…")
    serie = _fetch_hourly(SERIE_DESDE, cierre)
    m = calcular(serie, cierre)
    print(
        f"  CERO {m.h_cero} h = {_fmt(m.m3_cero)} m³ | "
        f"BASE {m.h_base} h = {_fmt(m.m3_base)} m³ | "
        f"ahorro {_fmt(m.ahorro_m3)} m³ ({_fmt(m.ahorro_pct)} %) = {_fmt_clp(m.clp_ahorro)}"
    )
    print(
        f"  Reposición observada 20/09 08:00 → cierre: "
        f"{_fmt(m.m3_desde_repo)} m³ en {m.h_desde_repo} h"
    )

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    charts = OUT_DIR / "_charts_club_house_corte_sep2026"
    charts.mkdir(parents=True, exist_ok=True)
    chart_diario = _chart_diario(m, charts / "diario_sep2026.png")
    chart_cmp = _chart_antes_despues(m, charts / "antes_durante_despues.png")
    spec = _build_spec(m, chart_diario, chart_cmp)

    one = OUT_DIR / "One_Pager_Gestion_Hidrica_Club_House_Corte_Total_Sep_2026.pdf"
    monthly = OUT_DIR / "Informe_Gestion_Hidrica_Club_House_Corte_Total_Sep_2026.pdf"
    render_one_pager(spec, one)
    render_mensual(spec, monthly, charts)

    meta = {
        "one_pager": str(one),
        "informe": str(monthly),
        "node_id": NODE_ID,
        "nombre": NOMBRE,
        "corte_programado_inicio": CORTE_INICIO.isoformat(sep=" "),
        "reposicion_indicada": CORTE_FIN_PROG.isoformat(sep=" "),
        "reposicion_observada": REPOSICION_OBS.isoformat(sep=" "),
        "cierre_datos": cierre.isoformat(sep=" "),
        "horas_cero": m.h_cero,
        "m3_cero": m.m3_cero,
        "horas_base": m.h_base,
        "m3_base": m.m3_base,
        "ahorro_m3": m.ahorro_m3,
        "ahorro_pct": m.ahorro_pct,
        "ahorro_clp": m.clp_ahorro,
        "m3_desde_reposicion": m.m3_desde_repo,
        "precio_clp_m3": TARIFA_CLP_M3,
        "clasificacion": spec.clasificacion,
    }
    (OUT_DIR / "meta_club_house_corte_sep2026.json").write_text(
        json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(f"\n[OK] {one.name}")
    print(f"[OK] {monthly.name}")
    print(json.dumps(meta, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
