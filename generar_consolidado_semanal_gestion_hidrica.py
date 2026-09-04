"""
Consolidado semanal de gestión hídrica: puntos a revisar.

Se corre los lunes. Cubre la semana lunes–domingo anterior, la compara con
la previa y lista solo los puntos que hay que atacar antes de fin de mes.

No reemplaza el informe de cierre. Destinatarios: Juan y Aníbal.

Uso:
  python generar_consolidado_semanal_gestion_hidrica.py
  python generar_consolidado_semanal_gestion_hidrica.py --hasta 31/08/2026
"""

from __future__ import annotations

import argparse
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests

from generar_informes_gestion_hidrica_cduc_agosto2026 import CLIENTES as CLIENTES_CDUC
from generar_informes_gestion_hidrica_colegios_agosto2026 import CLIENTES as CLIENTES_COLEGIOS
from generar_informes_gestion_hidrica_copec_agosto2026 import CLIENTES as CLIENTES_COPEC
from generar_informes_gestion_hidrica_fleming_agosto2026 import CLIENTES as CLIENTES_FLEMING
from generar_informes_gestion_hidrica_lote_agosto2026 import CLIENTES as CLIENTES_LOTE
from generar_informes_gestion_hidrica_lote_agosto2026 import (
    CACHE_DIR,
    _fmt,
    _iter_days,
    fetch_cliente,
)
from generar_informes_gestion_hidrica_semanal import (
    _cfg_semana,
    _rango_es,
    _semana_completa,
    _semana_previa,
    _wow,
)
from generar_reporte_word import acl_node_base_url
from informe_gestion_hidrica_pdf import build_chart_datos_perdidos, render_consolidado_semanal
from puntos_control_hidrico import estado_control, nota_red_cliente
from reporte_puntos_en_cero import (
    HORAS_UMBRAL_CONEXION_APP,
    _CHILE_TZ,
    obtener_estado_conexion_nodo,
)
from control_nocturno import _dt_local_from_csv_time, _value_by_time_last_row

PRIO_ORDEN = {"ATENCIÓN": 0, "AVISO": 1, "SEGUIMIENTO": 2}

# Puntos que OPERAN sin control (WES actúa). El resto de alzas es aviso al cliente.
SIN_CONTROL_REAL = {"000002-01", "000002-03", "000021-03"}

NOTAS_PUNTO: Dict[str, Dict[str, str]] = {
    "000002-01": {
        "tipo": "SIN CONTROL",
        "prio": "ATENCIÓN",
        "lectura": (
            "Está SIN CONTROL: el CPA está instalado pero no opera. "
            "Hay que activarlo y programarlo (peak 22:00–03:00)."
        ),
        "revisar": "Activar y programar el CPA",
    },
    "000021-03": {
        "tipo": "SIN CONTROL",
        "prio": "SEGUIMIENTO",
        "lectura": "Está SIN CONTROL. Punto bien. Visita técnica hoy.",
        "revisar": "Visita hoy",
    },
    "000022-00": {
        "tipo": "MONITOREO",
        "prio": "SEGUIMIENTO",
        "lectura": "Problema de monitoreo (sensor de pulso). Cambiar a ultrasonido.",
        "revisar": "Cambiar monitoreo: pulso → ultrasonido",
    },
    "000009-06": {
        "tipo": "AVISO",
        "prio": "SEGUIMIENTO",
        "lectura": "Matriz Principal no tiene CPA. Aviso al cliente.",
        "revisar": "Aviso al cliente: sin CPA",
    },
}

# Lectura conocida de horas perdidas. El resto de huecos queda sin causa.
NOTAS_HORAS: Dict[str, Dict[str, Any]] = {
    "000008-08": {
        "desconectado": True,
        "lectura": (
            "Visita 4/09: recinto desenergizado. "
            "Sin energía el WES no guarda datos ni puede monitorear."
        ),
    },
    "000022-00": {
        "lectura": (
            "Sensor de pulso: los huecos son de monitoreo. Cambiar a ultrasonido."
        ),
    },
}

LECTURA_SIN_CONEXION = (
    "Sin conexión. En terreno suele ser recinto desenergizado: "
    "sin energía el WES no guarda datos ni puede monitorear."
)


def _aplicar_lectura_operativa(filas: List[dict]) -> List[dict]:
    """Lo Valledor y Tupper = sin control. El resto de alzas = aviso al cliente."""
    out: List[dict] = []
    for f in filas:
        f = dict(f)
        nid = str(f.get("node_id") or "")
        nota = NOTAS_PUNTO.get(nid)
        if nota:
            f["tipo"] = nota["tipo"]
            f["control"] = nota["tipo"] if nota["tipo"] != "AVISO" else "AVISO CLIENTE"
            if nota["tipo"] == "AVISO":
                f["control"] = "AVISO CLIENTE"
            f["prio"] = nota.get("prio", f["prio"])
            f["lectura"] = nota["lectura"]
            f["revisar"] = nota["revisar"]
            f["tiene_control"] = nota["tipo"] not in ("SIN CONTROL", "AVISO")
            f["orden"] = (PRIO_ORDEN.get(f["prio"], 9), -float(f.get("total") or 0))
            out.append(f)
            continue
        if f.get("prio") == "ATENCIÓN" and nid not in SIN_CONTROL_REAL:
            wow = f.get("wow") or "—"
            f["tipo"] = "AVISO"
            f["prio"] = "AVISO"
            f["control"] = "AVISO CLIENTE"
            f["tiene_control"] = False
            f["lectura"] = (
                f"Aviso al cliente: consumo {wow} vs semana previa "
                f"({f.get('m3')} m³ vs {f.get('prev_m3')} m³). "
                "Informar al recinto para que confirme uso o evento."
            )
            extra = []
            if "pico" in (f.get("revisar") or "").lower():
                # keep peak date if the auto text had it
                for part in (f.get("revisar") or "").split(";"):
                    if "pico" in part.lower():
                        extra.append(part.strip())
            f["revisar"] = "Avisar al cliente el alza" + (
                f"; {extra[0]}" if extra else ""
            )
            f["orden"] = (PRIO_ORDEN["AVISO"], -float(f.get("total") or 0))
        else:
            f.setdefault("tipo", "SIN CONTROL" if nid in SIN_CONTROL_REAL else f.get("control"))
        out.append(f)
    return out


def _fecha_corta(iso: Optional[str]) -> str:
    if not iso:
        return ""
    try:
        return datetime.strptime(str(iso)[:10], "%Y-%m-%d").strftime("%d/%m")
    except ValueError:
        return ""


def clientes_seguimiento() -> List[dict]:
    out: List[dict] = []
    seen = set()
    for grupo in (
        CLIENTES_LOTE,
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


def _nodo_por_id(data: dict, node_id: str) -> Optional[dict]:
    for n in data.get("nodos") or []:
        if n.get("node_id") == node_id:
            return n
    return None


def _fila(
    cfg: dict,
    nodo: dict,
    *,
    prio: str,
    motivos: List[str],
    total: float,
    pct: float,
    wow: Optional[float],
    prev_t: float = 0.0,
    lectura: str = "",
) -> Dict[str, Any]:
    wow_txt = "—"
    if wow is not None:
        wow_txt = f"+{_fmt(wow, 0)} %" if wow >= 0 else f"−{_fmt(abs(wow), 0)} %"
    etiqueta, detalle, tiene = estado_control(cfg, nodo.get("node_id"))
    return {
        "prio": prio,
        "cliente": cfg["cliente"],
        "punto": nodo.get("short_name") or nodo.get("node_id"),
        "node_id": nodo.get("node_id"),
        "m3": _fmt(total, 1),
        "prev_m3": _fmt(prev_t, 1) if prev_t else "—",
        "wow": wow_txt,
        "noct": f"{_fmt(round(pct), 0)} %",
        "control": etiqueta,
        "control_detalle": detalle,
        "tiene_control": tiene,
        "lectura": lectura,
        "revisar": "; ".join(motivos),
        "orden": (PRIO_ORDEN.get(prio, 9), -total),
        "total": total,
    }


def _evaluar_punto(cfg: dict, nodo: dict, prev: Optional[dict]) -> Optional[Dict[str, Any]]:
    total = float(nodo.get("total") or 0)
    noct = float(nodo.get("nocturno_m3") or 0)
    pct = (noct / total * 100.0) if total else 0.0
    prev_t = float(prev.get("total") or 0) if prev else 0.0
    prev_dias = int(prev.get("dias") or 0) if prev else 0
    dias = int(nodo.get("dias") or 0)
    wow = _wow(total, prev_t) if prev_dias >= 5 and dias >= 5 and prev_t >= 20 else None
    avg = total / max(dias, 1)
    mx = float(nodo.get("max_m3") or 0)
    es_matriz = nodo.get("node_id") == cfg.get("matriz_id")
    additive = bool(cfg.get("additive"))
    explain = cfg.get("nocturnal_explain")
    motivos: List[str] = []
    prio: Optional[str] = None

    if total < 20 or dias < 5:
        return None
    # Encadenados (Zapallar / Inchcape): solo la matriz. COPEC sí informa interiores.
    if (
        not additive
        and cfg.get("matriz_id")
        and not es_matriz
        and explain in ("bombas_estanques", "wes")
    ):
        return None

    etiqueta, detalle, tiene = estado_control(cfg, nodo.get("node_id"))
    red = nota_red_cliente(cfg, nodo.get("node_id"), tiene)

    nombre = (nodo.get("short_name") or "").lower()
    infra_noct = any(k in nombre for k in ("estanque", "pozo"))
    noct_ok = additive or es_matriz or not cfg.get("matriz_id")
    if noct_ok and explain not in ("wes", "mercado") and not infra_noct:
        if round(pct) >= 35:
            if tiene:
                motivos.append("Verificar que el corte deje la madrugada cerca de 0")
            else:
                motivos.append("Revisar caudal 22–06 h; evaluar instalar o activar control")
            prio = "ATENCIÓN"
        elif round(pct) >= 25 and explain not in ("bombas_estanques", "piscina") and total >= 50:
            if tiene:
                motivos.append("Verificar programación y umbral del corte")
            else:
                motivos.append("Revisar caudal 22–06 h")
            prio = prio or "SEGUIMIENTO"

    if wow is not None and wow >= 25 and total >= 50:
        if tiene:
            motivos.append("Revisar si el CPA/WES cortó; si cortó, el extra es diurno")
            prio = "ATENCIÓN"
        elif wow >= 40 or (wow >= 25 and total >= 80):
            motivos.append("Confirmar uso vs fuga (no hay corte que lo contenga)")
            prio = "ATENCIÓN" if wow >= 40 else (prio or "SEGUIMIENTO")

    if avg > 0 and mx >= max(3.0 * avg, 20.0) and (wow is None or wow >= 0):
        when = _fecha_corta(str(nodo.get("max_fecha") or ""))
        extra = f" el {when}" if when else ""
        motivos.append(f"Revisar día pico ({_fmt(mx, 1)} m³{extra})")
        prio = prio or "SEGUIMIENTO"

    if not motivos or not prio:
        return None

    if wow is not None and wow >= 25:
        if tiene:
            lectura = (
                f"Alza de {_fmt(wow, 0)} % CON CONTROL activo. "
                "El equipo no evitó el aumento: revisar si el corte operó o si el extra es diurno."
            )
        else:
            lectura = (
                f"Alza de {_fmt(wow, 0)} % SIN CONTROL. El aumento corre sin corte automático."
            )
            if red:
                lectura += " " + red
    elif tiene and round(pct) >= 25:
        lectura = (
            f"Nocturno {_fmt(round(pct), 0)} % CON CONTROL. "
            "El corte debería dejar la madrugada cerca de 0."
        )
    elif round(pct) >= 25:
        lectura = f"Nocturno {_fmt(round(pct), 0)} % SIN CONTROL."
        if red:
            lectura += " " + red
    else:
        lectura = f"{etiqueta}: {detalle}."

    return _fila(
        cfg,
        nodo,
        prio=prio,
        motivos=motivos,
        total=total,
        pct=pct,
        wow=wow,
        prev_t=prev_t,
        lectura=lectura,
    )


def evaluar_cliente(cfg: dict, data: dict, data_prev: dict) -> Tuple[List[dict], bool]:
    filas: List[dict] = []
    nodos = list(data.get("nodos") or [])
    for nodo in nodos:
        prev = _nodo_por_id(data_prev, nodo["node_id"])
        fila = _evaluar_punto(cfg, nodo, prev)
        if fila:
            filas.append(fila)
    if cfg.get("cpa_estado") == "instalado_pendiente" and nodos:
        top = max(nodos, key=lambda n: float(n.get("total") or 0))
        top_id = top.get("node_id")
        cpa_txt = "Activar y programar el CPA (peak 22:00–03:00)"
        lectura_cpa = (
            "Está SIN CONTROL: el CPA está instalado pero no opera. "
            "Hay que activarlo y programarlo (peak 22:00–03:00)."
        )
        if not any(f.get("node_id") == top_id for f in filas):
            total = float(top.get("total") or 0)
            noct = float(top.get("nocturno_m3") or 0)
            pct = (noct / total * 100.0) if total else 0.0
            prev = _nodo_por_id(data_prev, top_id)
            prev_t = float(prev.get("total") or 0) if prev else 0.0
            wow = _wow(total, prev_t)
            filas.append(
                _fila(
                    cfg,
                    top,
                    prio="ATENCIÓN",
                    motivos=[cpa_txt],
                    total=total,
                    pct=pct,
                    wow=wow,
                    prev_t=prev_t,
                    lectura=lectura_cpa,
                )
            )
        else:
            for f in filas:
                if f.get("node_id") == top_id and "CPA" not in f["revisar"]:
                    f["revisar"] = cpa_txt + "; " + f["revisar"]
                    f["prio"] = "ATENCIÓN"
                    f["orden"] = (0, f["orden"][1])
                    f["lectura"] = lectura_cpa
    return filas, bool(filas)


MAX_HUECOS_CHART = 20
_MES_ABREV = (
    "ene",
    "feb",
    "mar",
    "abr",
    "may",
    "jun",
    "jul",
    "ago",
    "sep",
    "oct",
    "nov",
    "dic",
)


def _label_semana_iso(start: datetime, end: datetime) -> str:
    _iso_y, iso_w, _ = start.isocalendar()
    return (
        f"Semana {iso_w} — {start.day:02d} {_MES_ABREV[start.month - 1]} – "
        f"{end.day:02d} {_MES_ABREV[end.month - 1]} {end.year}"
    )


def _label_horas_punto(nombre: str, node_id: str, max_name: int = 28) -> str:
    name = (nombre or node_id or "—").strip()
    if len(name) > max_name:
        name = name[: max_name - 1] + "..."
    return f"{name} ({node_id})"


def _csv_horas_dia(node_id: str, dia: datetime) -> int:
    cache = CACHE_DIR / "csv_hours" / f"{node_id}_{dia.strftime('%Y%m%d')}.txt"
    if cache.is_file():
        try:
            return int(cache.read_text(encoding="utf-8").strip() or "0")
        except ValueError:
            pass
    url = f"{acl_node_base_url()}/nodes/{node_id}/dates.measures.csv"
    ds = dia.strftime("%d%m%Y")
    n = 0
    try:
        r = requests.get(url, params=[("start", ds), ("end", ds)], timeout=25)
        if r.status_code == 200 and r.text.strip():
            hours = set()
            for ts in _value_by_time_last_row(r.text):
                dt = _dt_local_from_csv_time(ts)
                if dt is not None and dt.date() == dia.date():
                    hours.add(int(dt.hour))
            n = len(hours)
    except Exception:
        n = 0
    cache.parent.mkdir(parents=True, exist_ok=True)
    cache.write_text(str(n), encoding="utf-8")
    return n


def _horas_perdidas_nodo(nodo: dict, start: datetime, end: datetime) -> int:
    have = {d["date"] for d in nodo.get("daily") or []}
    esperadas = 0
    presentes = 0
    for cur in _iter_days(start, end):
        esperadas += 24
        key = cur.strftime("%Y-%m-%d")
        if key not in have:
            continue
        presentes += _csv_horas_dia(str(nodo.get("node_id")), cur)
    return max(0, esperadas - presentes)


def _desconectado_ahora(node_id: str) -> bool:
    st = obtener_estado_conexion_nodo(node_id)
    last = st.get("lastUpdate")
    if last is None:
        return True
    now = datetime.now(_CHILE_TZ)
    if getattr(last, "tzinfo", None) is None:
        last = last.replace(tzinfo=_CHILE_TZ)
    return (now - last).total_seconds() >= HORAS_UMBRAL_CONEXION_APP * 3600


def _filas_horas_perdidas(nodos: List[dict], start: datetime, end: datetime) -> List[dict]:
    lost: Dict[str, int] = {}

    def _one(nodo: dict) -> Tuple[str, int]:
        nid = str(nodo.get("node_id"))
        return nid, _horas_perdidas_nodo(nodo, start, end)

    with ThreadPoolExecutor(max_workers=16) as ex:
        futs = {ex.submit(_one, n): n for n in nodos}
        for fut in as_completed(futs):
            nid, hrs = fut.result()
            lost[nid] = hrs

    cand = [n for n in nodos if lost.get(str(n.get("node_id")), 0) >= 1]
    estado: Dict[str, bool] = {}
    with ThreadPoolExecutor(max_workers=16) as ex:
        futs = {
            ex.submit(_desconectado_ahora, str(n.get("node_id"))): str(n.get("node_id"))
            for n in cand
        }
        for fut in as_completed(futs):
            nid = futs[fut]
            try:
                estado[nid] = bool(fut.result())
            except Exception:
                estado[nid] = True

    out: List[dict] = []
    for n in cand:
        nid = str(n.get("node_id"))
        nombre = n.get("node_name") or n.get("short_name") or nid
        out.append(
            {
                "node_id": nid,
                "punto": str(nombre),
                "horas": lost[nid],
                "desconectado": estado.get(nid, True),
                "label": _label_horas_punto(str(nombre), nid),
                "lectura": "—",
            }
        )
    out.sort(key=lambda r: (-float(r["horas"]), r["label"]))
    return out


def _aplicar_lectura_horas(filas: List[dict]) -> List[dict]:
    """Completa la lectura solo si hay causa conocida, o si está sin conexión."""
    out: List[dict] = []
    for f in filas:
        f = dict(f)
        nid = str(f.get("node_id") or "")
        nota = NOTAS_HORAS.get(nid)
        if nota:
            if "desconectado" in nota:
                f["desconectado"] = bool(nota["desconectado"])
            f["lectura"] = nota["lectura"]
        elif f.get("desconectado"):
            f["lectura"] = LECTURA_SIN_CONEXION
        else:
            f["lectura"] = "—"
        out.append(f)
    out.sort(key=lambda r: (-float(r["horas"]), r.get("label") or ""))
    return out


def generar_consolidado(start: datetime, end: datetime) -> Tuple[Path, List[dict], List[str]]:
    prev_start, prev_end = _semana_previa(start)
    print(
        f"[INFO] Consolidado {start.strftime('%d/%m')}–{end.strftime('%d/%m/%Y')} "
        f"(previa {prev_start.strftime('%d/%m')}–{prev_end.strftime('%d/%m')})",
        flush=True,
    )
    revisables: List[dict] = []
    sin_alerta: List[str] = []
    nodos_flota: List[dict] = []
    for base in clientes_seguimiento():
        cfg = _cfg_semana(base, start, end)
        cfg_prev = _cfg_semana(base, prev_start, prev_end)
        try:
            data = fetch_cliente(cfg)
            data_prev = fetch_cliente(cfg_prev)
        except Exception as e:
            print(f"[ADVERTENCIA] {base['cliente']}: {e}", flush=True)
            continue
        nodos_flota.extend(data.get("nodos") or [])
        filas, hay = evaluar_cliente(base, data, data_prev)
        if hay:
            revisables.extend(filas)
            print(f"  {base['cliente']}: {len(filas)} punto(s) a revisar", flush=True)
        else:
            sin_alerta.append(base["cliente"])
            print(f"  {base['cliente']}: sin alerta", flush=True)

    revisables.sort(key=lambda r: r["orden"])
    revisables = _aplicar_lectura_operativa(revisables)
    sin_control = [r for r in revisables if r.get("tipo") == "SIN CONTROL"]
    avisos = [
        r for r in revisables if r.get("tipo") == "AVISO" and r.get("prio") == "AVISO"
    ]
    seguimiento = [
        r
        for r in revisables
        if r.get("tipo") == "MONITOREO"
        or (r.get("tipo") == "AVISO" and r.get("prio") == "SEGUIMIENTO")
    ]
    print("[INFO] Horas perdidas de la semana...", flush=True)
    filas_all = _aplicar_lectura_horas(_filas_horas_perdidas(nodos_flota, start, end))
    n_nodos = len(nodos_flota)
    n_dias = (end.date() - start.date()).days + 1
    horas_esperadas = n_nodos * n_dias * 24
    horas_flota = sum(int(round(float(f["horas"]))) for f in filas_all)
    n_desc = sum(1 for f in filas_all if f.get("desconectado"))
    pct = (100.0 * horas_flota / horas_esperadas) if horas_esperadas else 0.0
    pct_txt = f"{pct:.1f}".replace(".", ",")
    titulo_sem = _label_semana_iso(start, end)
    filas_horas = filas_all[:MAX_HUECOS_CHART]
    nota_perdidos = (
        f"{len(filas_all)} puntos · {horas_flota} h de flota "
        f"({pct_txt}% de las esperadas) · {n_desc} de ellos están desconectados ahora."
    )
    resumen = (
        "De los 5, solo Lo Valledor está realmente sin control (CPA no opera). "
        "Raimundo Tupper también está sin control (visita hoy, punto bien). "
        "El resto es aviso al cliente, no falla de control."
    )
    if filas_horas:
        resumen += f" {len(filas_horas)} punto(s) con horas perdidas esta semana."
    chart_path: Optional[Path] = None
    periodo = f"Semana {_rango_es(start, end)}  ·  vs {_rango_es(prev_start, prev_end)}"
    footer = f"Consolidado semanal | {_rango_es(start, end)}"
    out_dir = Path("reports") / "CONSOLIDADO" / "SEMANAL"
    out = out_dir / (
        f"Consolidado_Semanal_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.pdf"
    )
    if filas_horas:
        chart_path = out_dir / "_charts" / "datos_perdidos.png"
        build_chart_datos_perdidos(chart_path, filas_horas, titulo=titulo_sem)
        print(f"[INFO] Datos perdidos: {nota_perdidos}", flush=True)
    render_consolidado_semanal(
        out,
        periodo=periodo,
        footer=footer,
        sin_control=sin_control,
        avisos=avisos,
        seguimiento=seguimiento,
        sin_alerta=sin_alerta,
        resumen=resumen,
        chart_perdidos=chart_path,
        nota_perdidos=nota_perdidos,
        n_perdidos=len(filas_horas),
        titulo_perdidos=titulo_sem,
        filas_perdidos=filas_horas,
    )
    print(f"[OK] {out}", flush=True)
    return out, revisables, sin_alerta


def main() -> int:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
            sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
        except Exception:
            pass
    parser = argparse.ArgumentParser()
    parser.add_argument("--hasta", default=None, help="Último día dd/mm/YYYY")
    args = parser.parse_args()
    hasta = datetime.strptime(args.hasta, "%d/%m/%Y") if args.hasta else None
    start, end = _semana_completa(hasta)
    generar_consolidado(start, end)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
