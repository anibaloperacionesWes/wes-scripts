"""
Informes de gestión hídrica — octubre 2026 (modelo Fundo Zapallar).

One pager + informe mensual para:
  CDUC, AGUNSA Lampa, AGUNSA Intermodal, Inchcape, Colegios Providencia,
  CORMUP, La Reina, Club Providencia, Nido de Águilas, Fundo Zapallar,
  UDD y Mercado Lo Valledor.

  python generar_informes_gestion_hidrica_lote_octubre2026.py zapallar udd valledor

La API interna no es alcanzable desde el agente: las medidas diarias salen
por el BFF público. El CSV horario exige sesión y no se inventa un nocturno.
"""

from __future__ import annotations

import copy
import json
import sys
import threading
from datetime import datetime
from pathlib import Path
from urllib.parse import urlencode, urlparse, urlunparse

import requests

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from wes_api_bff import bff_request  # noqa: E402

_cache_lock = threading.Lock()
_cache: dict[str, tuple[int, str, bytes]] = {}
_ORIG = requests.Session.request


def _full_url(url: str, params) -> str:
    if not params:
        return url
    if isinstance(params, dict):
        q = urlencode(params, doseq=True)
    else:
        q = urlencode(list(params), doseq=True)
    parts = urlparse(url)
    query = parts.query + ("&" if parts.query and q else "") + q
    return urlunparse(parts._replace(query=query))


def _to_public(url: str) -> str:
    return (
        str(url)
        .replace("http://104.248.53.141:7001", "https://api-rest.wes.cl")
        .replace("http://104.248.53.141:7003", "https://api-rest.wes.cl")
    )


def _bff_request(self, method, url, **kwargs):
    if (method or "GET").upper() != "GET" or "104.248.53.141" not in str(url):
        return _ORIG(self, method, url, **kwargs)
    full = _to_public(_full_url(str(url), kwargs.get("params")))
    if "dates.measures.csv" in full or "realtime.measures.csv" in full:
        resp = requests.Response()
        resp.status_code = 401
        resp.url = full
        resp._content = b'{"e":"unauthorized"}'
        resp.headers["Content-Type"] = "application/json"
        resp.encoding = "utf-8"
        return resp
    with _cache_lock:
        hit = _cache.get(full)
    if hit is None:
        data = bff_request(full, timeout=120)
        if isinstance(data, (dict, list)):
            body = json.dumps(data).encode("utf-8")
            ctype = "application/json"
        elif isinstance(data, str):
            body = data.encode("utf-8")
            ctype = "text/plain"
        else:
            body = b""
            ctype = "application/json"
        hit = (200, ctype, body)
        with _cache_lock:
            _cache[full] = hit
    status, ctype, body = hit
    resp = requests.Response()
    resp.status_code = status
    resp.url = full
    resp._content = body
    resp.headers["Content-Type"] = ctype
    resp.encoding = "utf-8"
    return resp


requests.Session.request = _bff_request  # type: ignore

import generar_reporte_word as grw  # noqa: E402


def _sin_serie_horaria(node_id, start_date, end_date, company_id=None):
    dias = (end_date.date() - start_date.date()).days + 1
    return {
        "dias_con_consumo_nocturno": 0,
        "dias_sin_consumo_nocturno": 0,
        "dias_sin_datos_horarios": max(dias, 0),
        "dias_con_datos_horarios": 0,
        "consumo_nocturno_total": 0.0,
        "consumo_diurno_efectivo": 0.0,
    }


grw.calculate_nocturnal_metrics = _sin_serie_horaria

import generar_informes_gestion_hidrica_lote_septiembre2026 as lote  # noqa: E402
from generar_informes_gestion_hidrica_cduc_agosto2026 import CLIENTES as CLIENTES_CDUC  # noqa: E402
from generar_informes_gestion_hidrica_colegios_agosto2026 import (  # noqa: E402
    CLIENTES as CLIENTES_COLEGIOS,
)

lote.calculate_nocturnal_metrics = _sin_serie_horaria
lote.START = "01/10/2026"
lote.END = "31/10/2026"
lote.START_DT = datetime(2026, 10, 1)
lote.END_DT = datetime(2026, 10, 31, 23, 59, 59)
lote.PERIODO_DIAS = 31
lote.CACHE_DIR = Path("/tmp/gh_octubre31")

WANTED = {
    "cduc",
    "lampa",
    "intermodal",
    "inchcape",
    "providencia",
    "cormup",
    "reina",
    "club",
    "nido",
    "zapallar",
    "udd",
    "valledor",
}


def _preparar(cfg: dict) -> dict:
    cfg = copy.deepcopy(cfg)
    cfg["start"] = "01/10/2026"
    cfg["end"] = "31/10/2026"
    cfg.pop("hallazgo_dato", None)
    cfg.pop("nota_agosto", None)
    cfg.pop("panorama_nota", None)
    cfg.pop("periodo_corto", None)
    cfg.pop("node_desde", None)
    if cfg["key"] == "providencia":
        cfg["cliente"] = "Corporación Providencia"
        cfg["sitio"] = "Corporación Providencia"
        cfg["sujeto"] = "los colegios"
        cfg["verbo_registro"] = "registraron"
    if cfg["key"] == "cormup":
        cfg["sujeto"] = "los colegios"
        cfg["verbo_registro"] = "registraron"
    if cfg["key"] == "reina":
        cfg["excluir_meses_6m"] = [(2026, 7)]
        cfg["usar_kpi_ultimo_mes_6m"] = True
        cfg["nota_agosto"] = (
            "Octubre comprende 31 días. Julio no entra en el comparativo de 6 meses "
            "por la falla del sensor de pulso. No se extrapola."
        )
        cfg["panorama_nota"] = (
            "Octubre se informa completo, del 1 al 31. Julio queda fuera del "
            "comparativo por la falla del sensor de pulso. No se extrapola."
        )
    if cfg["key"] == "valledor":
        cfg["sitio"] = "Mercado Lo Valledor"
        cfg["panorama_nota"] = (
            "Octubre se evalúa del 1 al 31 y no se extrapola. "
            "El equipo CPA está instalado y falta activar y programar."
        )
    if cfg["key"] == "zapallar":
        cfg["usar_kpi_ultimo_mes_6m"] = True
        cfg["nota_agosto"] = (
            "Octubre comprende 31 días. Matriz ESVAL y Etapa N.º 5 se informan "
            "el mes completo. No se extrapola."
        )
    return cfg


def main() -> None:
    pedidos = [a.strip().lower() for a in sys.argv[1:] if a.strip()]
    objetivo = set(pedidos) if pedidos else set(WANTED)
    desconocidos = objetivo - WANTED
    if desconocidos:
        raise SystemExit(f"Clientes no reconocidos: {sorted(desconocidos)}")
    catalogo = []
    seen = set()
    for grupo in (lote.CLIENTES, CLIENTES_CDUC, CLIENTES_COLEGIOS):
        for cfg in grupo:
            if cfg["key"] in objetivo and cfg["key"] not in seen:
                seen.add(cfg["key"])
                catalogo.append(_preparar(cfg))
    faltan = objetivo - seen
    if faltan:
        raise SystemExit(f"Faltan clientes en el catálogo: {sorted(faltan)}")
    orden = [
        "cduc",
        "lampa",
        "intermodal",
        "inchcape",
        "providencia",
        "cormup",
        "reina",
        "club",
        "nido",
        "zapallar",
        "udd",
        "valledor",
    ]
    by = {c["key"]: c for c in catalogo}
    lote.run_lote(
        [by[k] for k in orden if k in by],
        "GESTIÓN HÍDRICA — octubre 2026 — modelo Fundo Zapallar",
    )


if __name__ == "__main__":
    main()
