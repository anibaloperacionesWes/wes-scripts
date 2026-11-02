"""
Informes de gestión hídrica — octubre 2026 (modelo Fundo Zapallar).

One pager + informe mensual para:
  CDUC, AGUNSA Lampa, AGUNSA Intermodal, Inchcape, Colegios Providencia,
  CORMUP, La Reina, Club Providencia, Nido de Águilas, Fundo Zapallar,
  UDD y Mercado Lo Valledor.

  python generar_informes_gestion_hidrica_lote_octubre2026.py zapallar udd valledor

La API interna no es alcanzable desde el agente. Los totales diarios salen
por el BFF público. El Excel horario (dates.measures.csv) se baja con la
sesión de la app: variables WES_APP_USER y WES_APP_PASSWORD. No van en el código.
"""

from __future__ import annotations

import base64
import copy
import json
import os
import sys
import threading
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path
from urllib.parse import urlencode, urlparse, urlunparse

import requests

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from wes_api_bff import bff_request  # noqa: E402

_cache_lock = threading.Lock()
_token_lock = threading.Lock()
_cache: dict[str, tuple[int, str, bytes]] = {}
_ORIG = requests.Session.request
_TOKEN = ""


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


def _app_session() -> str:
    """Token de la app (POST login vía BFF). No se imprime ni se guarda en el repo."""
    global _TOKEN
    with _token_lock:
        if _TOKEN:
            return _TOKEN
        user = os.environ.get("WES_APP_USER", "").strip()
        password = os.environ.get("WES_APP_PASSWORD", "").strip()
        if not user or not password:
            raise SystemExit("Faltan WES_APP_USER y WES_APP_PASSWORD para el Excel horario.")
        login_url = "https://api-rest.wes.cl/wes/api/wes-mc-api/v1/login"
        payload = {"m": "POST", "u": login_url, "b": {"username": user, "password": password}}
        raw = base64.b64encode(json.dumps(payload, separators=(",", ":")).encode()).decode()
        req = urllib.request.Request(
            "https://api-rest.wes.cl/q",
            data=raw.encode(),
            headers={
                "Content-Type": "text/plain",
                "Origin": "https://dashboard.wes.cl",
                "Referer": "https://dashboard.wes.cl/",
                "Accept": "text/plain, */*",
                "User-Agent": "Mozilla/5.0 (compatible; WES-Scripts/1.0)",
            },
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            token = resp.read().decode("utf-8", errors="replace").strip().strip('"')
        if token.count(".") < 2:
            raise SystemExit("La app no entregó sesión para bajar el Excel horario.")
        _TOKEN = token
        return _TOKEN


def _csv_autenticado(url: str) -> bytes:
    token = _app_session()
    payload = {"m": "GET", "u": url, "b": None}
    raw = base64.b64encode(json.dumps(payload, separators=(",", ":")).encode()).decode()
    req = urllib.request.Request(
        "https://api-rest.wes.cl/q",
        data=raw.encode(),
        headers={
            "Content-Type": "text/plain",
            "authorization": token,
            "Origin": "https://dashboard.wes.cl",
            "Referer": "https://dashboard.wes.cl/",
            "Accept": "text/csv, text/plain, */*",
            "User-Agent": "Mozilla/5.0 (compatible; WES-Scripts/1.0)",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=90) as resp:
            return resp.read()
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            global _TOKEN
            with _token_lock:
                _TOKEN = ""
        raise


def _bff_request(self, method, url, **kwargs):
    if (method or "GET").upper() != "GET" or "104.248.53.141" not in str(url):
        return _ORIG(self, method, url, **kwargs)
    full = _to_public(_full_url(str(url), kwargs.get("params")))
    with _cache_lock:
        hit = _cache.get(full)
    if hit is None:
        if "dates.measures.csv" in full or "realtime.measures.csv" in full:
            body = _csv_autenticado(full)
            ctype = "text/csv"
            hit = (200, ctype, body)
        else:
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

import generar_informes_gestion_hidrica_lote_septiembre2026 as lote  # noqa: E402
from generar_informes_gestion_hidrica_cduc_agosto2026 import CLIENTES as CLIENTES_CDUC  # noqa: E402
from generar_informes_gestion_hidrica_colegios_agosto2026 import (  # noqa: E402
    CLIENTES as CLIENTES_COLEGIOS,
)

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
    if lote.CACHE_DIR.exists():
        for viejo in lote.CACHE_DIR.glob("*.json"):
            viejo.unlink()
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
