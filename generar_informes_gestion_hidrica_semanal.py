"""
One-pager semanal de gestión hídrica (formato Zapallar).

El informe de cierre se mantiene a fin de mes. Cada lunes se emite el par
semanal (one-pager + informe extendido) de lunes a domingo, comparado con
la semana previa.

Por defecto: Fundo Zapallar, última semana completa.
Con --todos: los 15 clientes del lote de fin de mes (8 + colegios + COPEC +
CDUC + Fleming).

Uso:
  python generar_informes_gestion_hidrica_semanal.py
  python generar_informes_gestion_hidrica_semanal.py --cliente zapallar
  python generar_informes_gestion_hidrica_semanal.py --todos --hasta 20/09/2026 --subir-drive
"""

from __future__ import annotations

import argparse
import copy
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from generar_informes_gestion_hidrica_lote_agosto2026 import (
    CACHE_DIR,
    CLIENTES,
    _clasificar,
    _fmt,
    _parrafo_visitas,
    _pretty_cls,
    _series_diarias,
    _visitas_spec,
    fetch_cliente,
)
from informe_gestion_hidrica_pdf import (
    Accion,
    Hallazgo,
    InformeSpec,
    PuntoIndicador,
    _fecha_es,
    _fmt_clp,
    build_chart_6_meses,
    build_chart_nocturno,
    build_chart_puntos,
    render_mensual,
    render_one_pager,
    resolve_logo,
)
from visitas_tecnicas_formulario import cargar_visitas_periodo, visitas_de_cliente

CACHE_PREFIX = "sem"


def clientes_fin_de_mes() -> List[dict]:
    """Lote de fin de mes: 8 comerciales + colegios + COPEC + CDUC + Fleming."""
    from generar_informes_gestion_hidrica_cduc_agosto2026 import CLIENTES as CLIENTES_CDUC
    from generar_informes_gestion_hidrica_colegios_agosto2026 import (
        CLIENTES as CLIENTES_COLEGIOS,
    )
    from generar_informes_gestion_hidrica_copec_agosto2026 import CLIENTES as CLIENTES_COPEC
    from generar_informes_gestion_hidrica_fleming_agosto2026 import (
        CLIENTES as CLIENTES_FLEMING,
    )

    out: List[dict] = []
    seen = set()
    for grupo in (
        CLIENTES,
        CLIENTES_COLEGIOS,
        CLIENTES_COPEC,
        CLIENTES_CDUC,
        CLIENTES_FLEMING,
    ):
        for cfg in grupo:
            key = cfg["key"]
            if key in seen:
                continue
            seen.add(key)
            out.append(cfg)
    return out


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


def _cfg_semana(
    base: dict,
    start: datetime,
    end: datetime,
    *,
    skip_serie_6m: bool = True,
) -> dict:
    cfg = copy.deepcopy(base)
    cfg["start"] = start.strftime("%d/%m/%Y")
    cfg["end"] = end.strftime("%d/%m/%Y")
    cfg["key"] = f"{base['key']}_{CACHE_PREFIX}_{start.strftime('%Y%m%d')}"
    cfg["skip_serie_6m"] = skip_serie_6m
    if skip_serie_6m:
        cfg.pop("excluir_meses_6m", None)
    cfg.pop("usar_kpi_ultimo_mes_6m", None)
    cfg.pop("hallazgo_dato", None)
    cfg.pop("nota_agosto", None)
    cfg.pop("panorama_nota", None)
    cfg.pop("periodo_corto", None)
    return cfg


def _drop_cache_sin_6m(cfg: dict) -> None:
    """La corrida anterior cacheó la semana sin serie de 6 meses."""
    path = CACHE_DIR / f"{cfg['key']}.json"
    if not path.is_file():
        return
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        path.unlink(missing_ok=True)
        return
    if not payload.get("serie_6_meses"):
        path.unlink(missing_ok=True)


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
                    f"La {cfg['matriz_name']} es el consumo real.",
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
                "Este informe semanal no reemplaza el de fin de mes: sirve para atacar "
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
        titulo_mensual="Informe semanal de gestión hídrica",
        clasificacion=clasificacion,
        motivo=motivo,
        kpi_entrada=f"{_fmt(entrada, 1)} m³",
        kpi_promedio=f"{_fmt(promedio, 1)} m³/día",
        kpi_nocturno=f"{_fmt(nocturno, 1)} m³",
        kpi_pct=f"{_fmt(round(pct), 0)} %",
        panorama=panorama,
        panorama_nota=(
            "Seguimiento semanal (lunes a domingo), comparado con la semana "
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


def _anexar_extendido(
    spec: InformeSpec,
    cfg: dict,
    data: dict,
    visitas: Sequence[Any],
    start: datetime,
    end: datetime,
) -> None:
    """Completa gráficos, tabla por punto y anexos del informe extendido."""
    nodos = data["nodos"]
    dias = int(data.get("periodo_dias") or 7)
    kpi = data["kpi"]
    extra = _parrafo_visitas(visitas)
    if extra:
        spec.lectura_ejecutiva.append(extra)
    spec.visitas = _visitas_spec(visitas)
    spec.nota_agosto = (
        f"Este informe cubre la semana {_rango_es(start, end)}. "
        "La gráfica de 6 meses es contexto del año; no se extrapola ni se cierra el mes."
    )
    spec.chart_puntos_nota = cfg.get("chart_nota") or (
        "cada barra es el consumo del punto en la semana."
    )
    spec.chart_nocturno_nota = cfg.get("nocturno_nota") or (
        "El consumo nocturno corresponde a 00:00–06:59, hora de Chile."
    )
    fecha_max = _fecha_es(kpi["max_fecha"]) if kpi.get("max_fecha") else "—"
    max_txt = f"{_fmt(float(kpi['max_m3']), 1)} m³" if kpi.get("max_fecha") else "—"
    spec.max_entrada_txt = (
        f"El mayor consumo diario de la referencia ocurrió el {fecha_max}, con {max_txt}."
    )
    indicadores = []
    orden = sorted(nodos, key=lambda n: -float(n["total"]))
    if cfg.get("matriz_id") and not cfg["additive"]:
        matriz = next(n for n in nodos if n["node_id"] == cfg["matriz_id"])
        resto = [n for n in orden if n["node_id"] != cfg["matriz_id"]]
        orden = [matriz] + resto
    for n in orden:
        max_dt = datetime.strptime(n["max_fecha"], "%Y-%m-%d") if n.get("max_fecha") else None
        indicadores.append(
            PuntoIndicador(
                nombre=n["short_name"],
                total=float(n["total"]),
                promedio=float(n["total"]) / dias if dias else 0.0,
                max_m3=float(n["max_m3"]),
                max_fecha=max_dt.strftime("%d/%m") if max_dt else "—",
                nocturno=float(n["nocturno_m3"]),
                cobertura=int(n["nocturno_cobertura"]),
                es_matriz=n["node_id"] == cfg.get("matriz_id"),
            )
        )
    spec.indicadores = indicadores
    costo = float(kpi.get("costo_nocturno") or 0)
    spec.criterio_nocturno = [
        [
            (
                (
                    cfg.get("ventana_nocturna")
                    or (
                        "Se considera nocturno el volumen medido entre las 00:00 y las 06:59, "
                        "hora de Chile. "
                    )
                )
                + "Los valores corresponden únicamente a días con datos y no "
                "se proyectan. El costo nocturno de la referencia se estima en ",
                False,
            ),
            (_fmt_clp(costo), True),
            (", con tarifa referencial de ", False),
            (f"{_fmt_clp(float(data.get('price_per_m3') or 0))}/m³", True),
            (".", False),
        ]
    ]
    spec.nota_cobertura = (
        "La cobertura nocturna indica cuántos días de la semana cuentan con registros "
        "en esa franja. Los días sin datos no se interpolan."
    )
    spec.series_diarias = _series_diarias(cfg, nodos, spec.hallazgos, data)

    out_dir = Path("reports") / cfg["folder"] / "GESTION_HIDRICA" / "SEMANAL"
    charts = out_dir / "_charts"
    charts.mkdir(parents=True, exist_ok=True)
    highlight = cfg.get("matriz_name") or ""
    if not cfg["additive"]:
        names = [n["short_name"] for n in nodos]
        totals = [float(n["total"]) for n in nodos]
        nocts = [float(n["nocturno_m3"]) for n in nodos]
    else:
        by_tot = sorted(nodos, key=lambda x: -float(x["total"]))
        names = [n["short_name"] for n in by_tot]
        totals = [float(n["total"]) for n in by_tot]
        nocts = [float(n["nocturno_m3"]) for n in by_tot]
    labels_6, vals_6 = [], []
    for item in data.get("serie_6_meses") or []:
        lab = str(item.get("label") or "").replace("*", "").split()[0].capitalize()
        labels_6.append(lab)
        vals_6.append(float(item.get("m3") or 0))
    if labels_6:
        spec.chart_6m = build_chart_6_meses(charts / f"{cfg['key']}_6m.png", labels_6, vals_6)
    spec.chart_puntos = build_chart_puntos(
        charts / f"{cfg['key']}_puntos.png",
        names,
        totals,
        highlight,
        additive=bool(cfg.get("additive")),
    )
    spec.chart_nocturno = build_chart_nocturno(
        charts / f"{cfg['key']}_nocturno.png",
        names,
        nocts,
        highlight,
        cfg.get("leyenda"),
    )


def _subir_drive(pdf: Path, folder: str) -> str:
    try:
        from wes_google_drive import credenciales_configuradas, subir_a_drive
    except Exception as e:
        print(f"[ADVERTENCIA] Drive no disponible: {e}", flush=True)
        return ""
    if not credenciales_configuradas():
        print("[ADVERTENCIA] Sin credenciales Drive; se omite la subida.", flush=True)
        return ""
    sub = f"{folder}/GESTION_HIDRICA/SEMANAL"
    info = subir_a_drive(pdf, subcarpeta=sub)
    link = info.get("web_view_link") or ""
    print(f"[OK] Drive ({sub}): {link}", flush=True)
    return link


def generar_semanal(
    base: dict,
    start: datetime,
    end: datetime,
    *,
    subir_drive: bool = False,
) -> dict:
    prev_start, prev_end = _semana_previa(start)
    cfg = _cfg_semana(base, start, end, skip_serie_6m=False)
    cfg_prev = _cfg_semana(base, prev_start, prev_end, skip_serie_6m=True)
    _drop_cache_sin_6m(cfg)
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
    stamp = f"{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}"
    one = out_dir / f"One_Pager_Semanal_{slug}_{stamp}.pdf"
    ext = out_dir / f"Informe_Semanal_{slug}_{stamp}.pdf"
    render_one_pager(spec, one)
    print(
        f"[OK] {one.name}  {spec.clasificacion}  "
        f"{spec.kpi_entrada}  noct {spec.kpi_pct}",
        flush=True,
    )
    _anexar_extendido(spec, cfg, data, visitas, start, end)
    render_mensual(spec, ext, out_dir / "_charts")
    print(f"[OK] {ext.name}", flush=True)
    drive_one = _subir_drive(one, base["folder"]) if subir_drive else ""
    drive_ext = _subir_drive(ext, base["folder"]) if subir_drive else ""
    return {
        "one_pager": one,
        "extendido": ext,
        "drive_one": drive_one,
        "drive_ext": drive_ext,
        "clasificacion": spec.clasificacion,
        "kpi": spec.kpi_entrada,
        "noct": spec.kpi_pct,
    }


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
    parser.add_argument(
        "--cliente",
        default=None,
        help="key del lote (default: zapallar). Varios separados por coma.",
    )
    parser.add_argument(
        "--todos",
        action="store_true",
        help="Genera one-pager e informe extendido de los 15 clientes de fin de mes.",
    )
    parser.add_argument(
        "--hasta",
        default=None,
        help="Último día (dd/mm/YYYY). Por defecto: última semana lunes–domingo cerrada.",
    )
    parser.add_argument(
        "--subir-drive",
        action="store_true",
        help="Sube cada PDF a Drive en <cliente>/GESTION_HIDRICA/SEMANAL.",
    )
    parser.add_argument(
        "--sin-drive",
        action="store_true",
        help="No subir a Drive (prioridad sobre --subir-drive).",
    )
    args = parser.parse_args()
    catalogo = clientes_fin_de_mes()
    if args.todos:
        seleccion = catalogo
    elif args.cliente:
        keys = [k.strip().lower() for k in args.cliente.split(",") if k.strip()]
        by_key = {c["key"]: c for c in catalogo}
        missing = [k for k in keys if k not in by_key]
        if missing:
            print(f"[ERROR] Cliente(s) no están en el lote: {', '.join(missing)}", file=sys.stderr)
            print("Disponibles:", ", ".join(c["key"] for c in catalogo), file=sys.stderr)
            return 1
        seleccion = [by_key[k] for k in keys]
    else:
        seleccion = [next(c for c in catalogo if c["key"] == "zapallar")]
    start, end = _semana_completa(_parse_hasta(args.hasta))
    subir = bool(args.subir_drive) and not args.sin_drive
    print(
        f"INFORMES SEMANALES (one-pager + extendido) · {_rango_es(start, end)} · {len(seleccion)} cliente(s)\n",
        flush=True,
    )
    ok: List[dict] = []
    errors: List[str] = []
    for base in seleccion:
        try:
            out = generar_semanal(base, start, end, subir_drive=subir)
            ok.append(
                {
                    "key": base["key"],
                    "cliente": base["cliente"],
                    "one_pager": str(out["one_pager"]),
                    "extendido": str(out["extendido"]),
                    "drive_one": out["drive_one"],
                    "drive_ext": out["drive_ext"],
                    "clasificacion": out["clasificacion"],
                    "kpi": out["kpi"],
                    "noct": out["noct"],
                }
            )
        except Exception as e:
            errors.append(f"{base['cliente']}: {e}")
            print(f"[ERROR] {base['cliente']}: {e}", flush=True)
            import traceback

            traceback.print_exc()
    resumen = Path("reports") / "CONSOLIDADO" / "SEMANAL" / (
        f"One_Pagers_Semanal_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.json"
    )
    resumen.parent.mkdir(parents=True, exist_ok=True)
    resumen.write_text(
        json.dumps(
            {
                "periodo": _rango_es(start, end),
                "start": start.strftime("%Y-%m-%d"),
                "end": end.strftime("%Y-%m-%d"),
                "ok": ok,
                "errors": errors,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"\n[INFO] Completados: {len(ok)}/{len(seleccion)}", flush=True)
    print(f"[INFO] Resumen: {resumen}", flush=True)
    if errors:
        print("[INFO] Fallidos:", flush=True)
        for e in errors:
            print("  -", e, flush=True)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
