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
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from generar_informes_gestion_hidrica_cduc_agosto2026 import CLIENTES as CLIENTES_CDUC
from generar_informes_gestion_hidrica_colegios_agosto2026 import CLIENTES as CLIENTES_COLEGIOS
from generar_informes_gestion_hidrica_copec_agosto2026 import CLIENTES as CLIENTES_COPEC
from generar_informes_gestion_hidrica_fleming_agosto2026 import CLIENTES as CLIENTES_FLEMING
from generar_informes_gestion_hidrica_lote_agosto2026 import CLIENTES as CLIENTES_LOTE
from generar_informes_gestion_hidrica_lote_agosto2026 import _fmt, fetch_cliente
from generar_informes_gestion_hidrica_semanal import (
    _cfg_semana,
    _rango_es,
    _semana_completa,
    _semana_previa,
    _wow,
)
from informe_gestion_hidrica_pdf import render_consolidado_semanal

PRIO_ORDEN = {"ATENCIÓN": 0, "SEGUIMIENTO": 1}


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


def _evaluar_punto(cfg: dict, nodo: dict, prev: Optional[dict]) -> Optional[Dict[str, Any]]:
    total = float(nodo.get("total") or 0)
    noct = float(nodo.get("nocturno_m3") or 0)
    pct = (noct / total * 100.0) if total else 0.0
    prev_t = float(prev.get("total") or 0) if prev else 0.0
    wow = _wow(total, prev_t)
    dias = max(int(nodo.get("dias") or 7), 1)
    avg = total / dias
    mx = float(nodo.get("max_m3") or 0)
    es_matriz = nodo.get("node_id") == cfg.get("matriz_id")
    additive = bool(cfg.get("additive"))
    explain = cfg.get("nocturnal_explain")
    motivos: List[str] = []
    prio: Optional[str] = None

    if cfg.get("cpa_estado") == "instalado_pendiente" and (es_matriz or additive):
        motivos.append("CPA instalado, falta activar")
        prio = "ATENCIÓN"

    noct_ok = True
    if not additive and not es_matriz:
        noct_ok = False

    if noct_ok:
        if pct >= 35:
            motivos.append(f"nocturno {_fmt(round(pct), 0)} %")
            prio = "ATENCIÓN"
        elif pct >= 18:
            if explain in ("wes", "mercado"):
                pass
            elif explain == "bombas_estanques":
                motivos.append(f"nocturno {_fmt(round(pct), 0)} % (bombas/estanques)")
                prio = prio or "SEGUIMIENTO"
            else:
                motivos.append(f"nocturno {_fmt(round(pct), 0)} %")
                prio = prio or "SEGUIMIENTO"

    if wow is not None and wow >= 25:
        motivos.append(f"subió {_fmt(wow, 0)} % vs semana previa")
        prio = "ATENCIÓN"
    elif wow is not None and wow >= 15:
        motivos.append(f"subió {_fmt(wow, 0)} % vs semana previa")
        prio = prio or "SEGUIMIENTO"

    if avg > 0 and mx >= max(2.5 * avg, 8.0):
        motivos.append(f"pico {_fmt(mx, 1)} m³")
        prio = prio or "SEGUIMIENTO"

    if not motivos or not prio:
        return None

    wow_txt = "—"
    if wow is not None:
        wow_txt = f"+{_fmt(wow, 0)} %" if wow >= 0 else f"−{_fmt(abs(wow), 0)} %"
    return {
        "prio": prio,
        "cliente": cfg["cliente"],
        "punto": nodo.get("short_name") or nodo.get("node_id"),
        "m3": _fmt(total, 1),
        "wow": wow_txt,
        "noct": f"{_fmt(round(pct), 0)} %",
        "revisar": "; ".join(motivos),
        "orden": (PRIO_ORDEN.get(prio, 9), -total),
    }


def evaluar_cliente(cfg: dict, data: dict, data_prev: dict) -> Tuple[List[dict], bool]:
    filas: List[dict] = []
    for nodo in data.get("nodos") or []:
        prev = _nodo_por_id(data_prev, nodo["node_id"])
        fila = _evaluar_punto(cfg, nodo, prev)
        if fila:
            filas.append(fila)
    return filas, bool(filas)


def generar_consolidado(start: datetime, end: datetime) -> Tuple[Path, List[dict], List[str]]:
    prev_start, prev_end = _semana_previa(start)
    print(
        f"[INFO] Consolidado {start.strftime('%d/%m')}–{end.strftime('%d/%m/%Y')} "
        f"(previa {prev_start.strftime('%d/%m')}–{prev_end.strftime('%d/%m')})",
        flush=True,
    )
    revisables: List[dict] = []
    sin_alerta: List[str] = []
    for base in clientes_seguimiento():
        cfg = _cfg_semana(base, start, end)
        cfg_prev = _cfg_semana(base, prev_start, prev_end)
        try:
            data = fetch_cliente(cfg)
            data_prev = fetch_cliente(cfg_prev)
        except Exception as e:
            print(f"[ADVERTENCIA] {base['cliente']}: {e}", flush=True)
            continue
        filas, hay = evaluar_cliente(base, data, data_prev)
        if hay:
            revisables.extend(filas)
            print(f"  {base['cliente']}: {len(filas)} punto(s) a revisar", flush=True)
        else:
            sin_alerta.append(base["cliente"])
            print(f"  {base['cliente']}: sin alerta", flush=True)

    revisables.sort(key=lambda r: r["orden"])
    table = [
        [r["cliente"], r["punto"], r["m3"], r["wow"], r["noct"], r["revisar"]]
        for r in revisables
    ]
    n_cli = len({r["cliente"] for r in revisables})
    resumen = (
        f"{len(revisables)} punto(s) a revisar en {n_cli} cliente(s). "
        f"{len(sin_alerta)} cliente(s) sin alerta esta semana."
    )
    periodo = f"Semana {_rango_es(start, end)}  ·  vs {_rango_es(prev_start, prev_end)}"
    footer = f"Consolidado semanal | {_rango_es(start, end)}"
    out_dir = Path("reports") / "CONSOLIDADO" / "SEMANAL"
    out = out_dir / (
        f"Consolidado_Semanal_{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}.pdf"
    )
    render_consolidado_semanal(
        out,
        periodo=periodo,
        footer=footer,
        filas=table,
        sin_alerta=sin_alerta,
        resumen=resumen,
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
