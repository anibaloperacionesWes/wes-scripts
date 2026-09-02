"""
One-pager semanal de gestión hídrica (formato Zapallar).

El informe mensual se mantiene a fin de mes. Este one-pager cubre una semana
completa (lunes a domingo), la compara con la semana previa y marca qué hay
que atacar antes del cierre.

Por defecto: Fundo Zapallar, última semana completa.

Uso:
  python generar_informes_gestion_hidrica_semanal.py
  python generar_informes_gestion_hidrica_semanal.py --cliente zapallar
  python generar_informes_gestion_hidrica_semanal.py --hasta 30/08/2026
"""

from __future__ import annotations

import argparse
import copy
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from generar_informes_gestion_hidrica_lote_agosto2026 import (
    CLIENTES,
    _clasificar,
    _fmt,
    _pretty_cls,
    fetch_cliente,
)
from informe_gestion_hidrica_pdf import (
    Accion,
    Hallazgo,
    InformeSpec,
    _fecha_es,
    render_one_pager,
    resolve_logo,
)
from visitas_tecnicas_formulario import cargar_visitas_periodo, visitas_de_cliente

CACHE_PREFIX = "sem"


def _lunes(dt: datetime) -> datetime:
    d = datetime(dt.year, dt.month, dt.day)
    return d - timedelta(days=d.weekday())


def _semana_completa(hasta: Optional[datetime] = None) -> Tuple[datetime, datetime]:
    """Última semana lunes–domingo ya cerrada (o la que contiene `hasta` si cae en domingo)."""
    ref = hasta or (datetime.now() - timedelta(days=1))
    ref = datetime(ref.year, ref.month, ref.day)
    lunes = _lunes(ref)
    domingo = lunes + timedelta(days=6)
    if domingo > ref:
        lunes = lunes - timedelta(days=7)
        domingo = lunes + timedelta(days=6)
    return lunes, domingo


def _semana_previa(inicio: datetime) -> Tuple[datetime, datetime]:
    prev_fin = inicio - timedelta(days=1)
    prev_ini = prev_fin - timedelta(days=6)
    return prev_ini, prev_fin


def _rango_es(start: datetime, end: datetime) -> str:
    meses = (
        "enero",
        "febrero",
        "marzo",
        "abril",
        "mayo",
        "junio",
        "julio",
        "agosto",
        "septiembre",
        "octubre",
        "noviembre",
        "diciembre",
    )
    if start.month == end.month and start.year == end.year:
        return f"{start.day} al {end.day} de {meses[start.month - 1]} de {start.year}"
    return (
        f"{start.day} de {meses[start.month - 1]} al {end.day} de "
        f"{meses[end.month - 1]} de {end.year}"
    )


def _cfg_semana(base: dict, start: datetime, end: datetime) -> dict:
    cfg = copy.deepcopy(base)
    cfg["start"] = start.strftime("%d/%m/%Y")
    cfg["end"] = end.strftime("%d/%m/%Y")
    cfg["key"] = f"{base['key']}_{CACHE_PREFIX}_{start.strftime('%Y%m%d')}"
    cfg["skip_serie_6m"] = True
    cfg.pop("excluir_meses_6m", None)
    cfg.pop("usar_kpi_ultimo_mes_6m", None)
    cfg.pop("hallazgo_dato", None)
    return cfg


def _wow(actual: float, previa: float) -> Optional[float]:
    if previa <= 0:
        return None
    return (actual - previa) / previa * 100.0


def _hallazgos_semanal(
    cfg: dict,
    data: dict,
    data_prev: dict,
) -> Tuple[List[Hallazgo], bool]:
    kpi = data["kpi"]
    prev = data_prev["kpi"]
    nodos = data["nodos"]
    entrada = float(kpi["entrada"])
    pct = float(kpi["pct_nocturno"])
    nocturno = float(kpi["nocturno"])
    prev_ent = float(prev["entrada"])
    prev_pct = float(prev["pct_nocturno"])
    wow = _wow(entrada, prev_ent)
    wow_n = _wow(nocturno, float(prev["nocturno"]))
    hall: List[Hallazgo] = []
    evento_fuerte = False

    if wow is not None and wow >= 25:
        hall.append(
            Hallazgo(
                "ATENCIÓN",
                f"La semana subió {_fmt(wow, 0)} % vs la previa",
                f"{_fmt(entrada, 1)} m³ frente a {_fmt(prev_ent, 1)} m³ la semana anterior.",
                "Revisar causa esta semana (fuga, riego continuo o dato anómalo) "
                "para no arrastrarlo al cierre de mes.",
            )
        )
    elif wow is not None and wow >= 10:
        hall.append(
            Hallazgo(
                "SEGUIMIENTO",
                f"Alza de {_fmt(wow, 0)} % vs la semana previa",
                f"{_fmt(entrada, 1)} m³ esta semana; {_fmt(prev_ent, 1)} m³ la anterior.",
                "Confirmar si corresponde a operación. Si se sostiene, actuar antes de fin de mes.",
            )
        )
    elif wow is not None and wow <= -15:
        hall.append(
            Hallazgo(
                "INFORMATIVA",
                f"Bajó {_fmt(abs(wow), 0)} % vs la semana previa",
                f"{_fmt(entrada, 1)} m³ esta semana; {_fmt(prev_ent, 1)} m³ la anterior.",
                "Queda como referencia. Si el nocturno no baja con el total, revisar pérdidas.",
            )
        )
    else:
        hall.append(
            Hallazgo(
                "INFORMATIVA",
                "Consumo en línea con la semana previa",
                f"{_fmt(entrada, 1)} m³ esta semana"
                + (f" ({_fmt(wow, 0)} % vs la anterior)." if wow is not None else "."),
                "Seguir el nocturno y los picos diarios hasta el cierre de mes.",
            )
        )

    delta_pct = pct - prev_pct
    if abs(delta_pct) >= 5 or pct >= 18:
        prio = "ATENCIÓN" if pct >= 35 else ("SEGUIMIENTO" if pct >= 18 or delta_pct >= 5 else "INFORMATIVA")
        sentido = "subió" if delta_pct >= 0 else "bajó"
        hall.append(
            Hallazgo(
                prio,
                f"{_fmt(round(pct), 0)} % nocturno esta semana",
                f"{_fmt(nocturno, 1)} m³ entre 00:00 y 06:59. "
                f"La semana previa fue {_fmt(round(prev_pct), 0)} % "
                f"({sentido} {_fmt(abs(delta_pct), 0)} puntos).",
                "Separar riego, llenado o bombas de una posible pérdida. "
                "Si el alza se sostiene, elevar el estado antes del cierre.",
            )
        )

    avg = entrada / max(int(data.get("periodo_dias") or 7), 1)
    mx = float(kpi.get("max_m3") or 0)
    mx_f = kpi.get("max_fecha")
    if avg > 0 and mx >= max(2.5 * avg, 8.0) and mx_f:
        evento_fuerte = mx >= 4.0 * avg and mx >= 12
        hall.append(
            Hallazgo(
                "SEGUIMIENTO",
                f"Pico de {_fmt(mx, 1)} m³ el {_fecha_es(mx_f)}",
                f"El promedio de la semana fue {_fmt(avg, 1)} m³/día.",
                "Confirmar si fue una maniobra. Si se repite, atacar esta semana.",
            )
        )

    if not cfg.get("additive") and cfg.get("matriz_name"):
        if len(hall) < 3:
            hall.append(
                Hallazgo(
                    "INFORMATIVA",
                    "La referencia es la matriz, no la suma interna",
                    f"La {cfg['matriz_name']} es el consumo real de {cfg['sujeto']}.",
                    "Los demás medidores se leen como control interno.",
                )
            )
    elif len(nodos) >= 2 and len(hall) < 3:
        top = max(nodos, key=lambda n: float(n["total"]))
        pct_top = float(top["total"]) / entrada * 100.0 if entrada else 0.0
        hall.append(
            Hallazgo(
                "INFORMATIVA",
                f"{top['short_name']} concentra el {_fmt(round(pct_top), 0)} % de la semana",
                f"{_fmt(top['total'], 1)} m³ en estos días.",
                "Es el punto a vigilar si el total se dispara antes de fin de mes.",
            )
        )

    _ = wow_n
    return hall[:3], evento_fuerte


def _clasificar_semanal(
    cfg: dict,
    pct: float,
    evento_fuerte: bool,
    wow: Optional[float],
) -> Tuple[str, str]:
    cls, motivo = _clasificar(cfg, pct, evento_fuerte, salto=None)
    if wow is not None and wow >= 50:
        return (
            "REQUIERE ATENCIÓN",
            f"el consumo de la semana subió {_fmt(wow, 0)} % respecto de la semana previa.",
        )
    if wow is not None and wow >= 25 and cls == "BAJO CONTROL":
        return (
            "EN OBSERVACIÓN",
            f"el consumo de la semana subió {_fmt(wow, 0)} % respecto de la semana previa.",
        )
    return cls, motivo


def _acciones_semanal(hallazgos: Sequence[Hallazgo], cfg: dict) -> List[Accion]:
    acts: List[Accion] = []
    for h in hallazgos:
        t = h.titulo.lower()
        if h.prioridad == "ATENCIÓN" or "subió" in t:
            acts.append(
                Accion(
                    "Diagnosticar el alza de esta semana.",
                    "Esta semana",
                    "Atacar la causa antes del cierre de mes.",
                    "Operación + WES",
                )
            )
        elif "nocturno" in t:
            acts.append(
                Accion(
                    "Confirmar riego, llenado o bombas en horario nocturno.",
                    "Esta semana",
                    "Separar operación de una posible pérdida antes de fin de mes.",
                    "Administración / operación",
                )
            )
        elif "pico" in t or "alza" in t:
            acts.append(
                Accion(
                    "Revisar el pico diario de esta semana.",
                    "Esta semana",
                    "Validar causa y descartar fuga antes de fin de mes.",
                    "Operación + WES",
                )
            )
    acts.append(
        Accion(
            "Mantener seguimiento hasta el informe de cierre.",
            "Antes de fin de mes",
            "Detectar un nuevo alza a tiempo y no esperar el mensual.",
            "WES + cliente",
        )
    )
    seen = set()
    out: List[Accion] = []
    for a in acts:
        if a.accion in seen:
            continue
        seen.add(a.accion)
        out.append(a)
        if len(out) == 3:
            break
    return out


def build_spec_semanal(
    cfg: dict,
    data: dict,
    data_prev: dict,
    visitas: Sequence[Any],
    start: datetime,
    end: datetime,
    prev_start: datetime,
    prev_end: datetime,
) -> InformeSpec:
    kpi = data["kpi"]
    prev = data_prev["kpi"]
    entrada = float(kpi["entrada"])
    nocturno = float(kpi["nocturno"])
    pct = float(kpi["pct_nocturno"])
    promedio = float(kpi["promedio"])
    prev_ent = float(prev["entrada"])
    wow = _wow(entrada, prev_ent)
    hallazgos, evento_fuerte = _hallazgos_semanal(cfg, data, data_prev)
    clasificacion, motivo = _clasificar_semanal(cfg, pct, evento_fuerte, wow)
    acciones = _acciones_semanal(hallazgos, cfg)
    dias = int(data.get("periodo_dias") or 7)
    fecha_max = _fecha_es(kpi["max_fecha"]) if kpi.get("max_fecha") else "—"
    max_txt = f"{_fmt(float(kpi['max_m3']), 1)} m³" if kpi.get("max_fecha") else "—"
    if wow is None:
        vs = "sin semana previa comparable"
        vs_run = [("sin semana previa comparable", False)]
    elif wow >= 0:
        vs = f"subió {_fmt(wow, 0)} % vs la semana previa"
        vs_run = [("subió ", False), (f"{_fmt(wow, 0)} %", True), (" vs la semana previa", False)]
    else:
        vs = f"bajó {_fmt(abs(wow), 0)} % vs la semana previa"
        vs_run = [("bajó ", False), (f"{_fmt(abs(wow), 0)} %", True), (" vs la semana previa", False)]

    panorama = [
        (f"En {cfg['sitio']} esta semana se registraron ", False),
        (f"{_fmt(entrada, 1)} m³", True),
        (", con un promedio de ", False),
        (f"{_fmt(promedio, 1)} m³ diarios", True),
        (f" en {dias} días. El total ", False),
        *vs_run,
        (f" ({_fmt(prev_ent, 1)} m³ del {_rango_es(prev_start, prev_end)}). ", False),
        ("El mayor consumo diario fue el ", False),
        (fecha_max, True),
        (", con ", False),
        (max_txt, True),
        (".", False),
    ]
    lectura = [
        [
            (f"Seguimiento semanal: {cfg['sujeto']} registró ", False),
            (f"{_fmt(entrada, 1)} m³", True),
            (". El nocturno alcanzó ", False),
            (f"{_fmt(nocturno, 1)} m³", True),
            (" (", False),
            (f"{_fmt(round(pct), 0)} %", True),
            (f"). El total {vs}.", False),
        ],
        [
            ("El estado de la semana se clasifica como ", False),
            (_pretty_cls(clasificacion), True),
            (". ", False),
            (motivo[0].upper() + motivo[1:] if motivo else "", False),
        ],
    ]
    conclusion = [
        [
            ("El estado de la semana es ", False),
            (f"“{_pretty_cls(clasificacion)}”", True),
            (". ", False),
            (motivo[0].upper() + motivo[1:] if motivo else "", False),
        ],
        [
            (
                "Este one-pager no reemplaza el informe de fin de mes: sirve para atacar "
                "alzas, picos o nocturno anómalo ahora. Si el patrón se sostiene, "
                "la clasificación de cierre avanzará a ",
                False,
            ),
            (
                "“Crítico”"
                if clasificacion == "REQUIERE ATENCIÓN"
                else "“Requiere atención”",
                True,
            ),
            (".", False),
        ],
    ]
    return InformeSpec(
        cliente=cfg["cliente"],
        sitio=cfg["sitio"],
        periodo_corto=f"{cfg['sitio']} · {_rango_es(start, end)}",
        footer=f"Seguimiento semanal - {cfg['cliente']} | {_rango_es(start, end)}",
        titulo_onepager="Seguimiento semanal de gestión hídrica",
        titulo_mensual="Seguimiento semanal de gestión hídrica",
        clasificacion=clasificacion,
        motivo=motivo,
        kpi_entrada=f"{_fmt(entrada, 1)} m³",
        kpi_promedio=f"{_fmt(promedio, 1)} m³/día",
        kpi_nocturno=f"{_fmt(nocturno, 1)} m³",
        kpi_pct=f"{_fmt(round(pct), 0)} %",
        panorama=panorama,
        panorama_nota=(
            "One-pager de seguimiento semanal (lunes a domingo), comparado con la semana "
            "previa. No se extrapola. El informe de cierre se emite a fin de mes."
            + (
                f" En la semana se registró {len(visitas)} visita(s) técnica(s)."
                if visitas
                else ""
            )
        ),
        hallazgos=hallazgos,
        acciones=acciones,
        conclusion=conclusion,
        lectura_ejecutiva=lectura,
        nota_agosto="",
        kpi_consumo_label="Consumo de la semana",
        logo_path=resolve_logo(),
    )


def generar_semanal(
    base: dict,
    start: datetime,
    end: datetime,
) -> Path:
    prev_start, prev_end = _semana_previa(start)
    cfg = _cfg_semana(base, start, end)
    cfg_prev = _cfg_semana(base, prev_start, prev_end)
    print(
        f"[INFO] Semana {start.strftime('%d/%m')}–{end.strftime('%d/%m/%Y')} "
        f"(previa {prev_start.strftime('%d/%m')}–{prev_end.strftime('%d/%m')})",
        flush=True,
    )
    data = fetch_cliente(cfg)
    data_prev = fetch_cliente(cfg_prev)
    try:
        todas = cargar_visitas_periodo(start, end)
        visitas = visitas_de_cliente(todas, base)
    except Exception as e:
        print(f"[ADVERTENCIA] Visitas: {e}", flush=True)
        visitas = []
    spec = build_spec_semanal(
        cfg, data, data_prev, visitas, start, end, prev_start, prev_end
    )
    out_dir = Path("reports") / base["folder"] / "GESTION_HIDRICA" / "SEMANAL"
    slug = base["cliente"].replace(" ", "_").replace("Á", "A").replace("á", "a")
    out = out_dir / (
        f"One_Pager_Semanal_{slug}_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.pdf"
    )
    render_one_pager(spec, out)
    print(
        f"[OK] {out}  {spec.clasificacion}  "
        f"{spec.kpi_entrada}  noct {spec.kpi_pct}",
        flush=True,
    )
    return out


def _parse_hasta(raw: Optional[str]) -> Optional[datetime]:
    if not raw:
        return None
    return datetime.strptime(raw.strip(), "%d/%m/%Y")


def main() -> int:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
            sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
        except Exception:
            pass
    parser = argparse.ArgumentParser()
    parser.add_argument("--cliente", default="zapallar", help="key del lote (default: zapallar)")
    parser.add_argument(
        "--hasta",
        default=None,
        help="Último día (dd/mm/YYYY). Por defecto: última semana lunes–domingo cerrada.",
    )
    args = parser.parse_args()
    base = next((c for c in CLIENTES if c["key"] == args.cliente), None)
    if base is None:
        print(f"[ERROR] Cliente no está en el lote: {args.cliente}", file=sys.stderr)
        return 1
    start, end = _semana_completa(_parse_hasta(args.hasta))
    generar_semanal(base, start, end)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
