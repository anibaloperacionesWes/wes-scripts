"""
Cliente HTTP hacia la API WES vía BFF de dashboard.wes.cl.

La app Angular no llama directo a /wes/api/* desde el navegador en todos los
casos: arma un payload opaco y consulta https://api-rest.wes.cl/q?d=...

Los endpoints diarios (measures/dates, companies, alerts) responden sin token.
Algunos CSV horarios exigen x-auth-token (login wes-mc-api).
"""

from __future__ import annotations

import base64
import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Dict, List, Optional

API_REST = os.environ.get("WES_API_REST", "https://api-rest.wes.cl").rstrip("/")
BFF_URL = f"{API_REST}/q"
ACL_NODE = f"{API_REST}/wes/api/acl-node/v1"
ACL_ENTITIES = f"{API_REST}/wes/api/acl-entities/v1"
MC_API = f"{API_REST}/wes/api/wes-mc-api/v1"

_DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (compatible; WES-Scripts/1.0; +https://dashboard.wes.cl)"
    ),
    "Accept": "application/json, text/plain, */*",
    "Origin": "https://dashboard.wes.cl",
    "Referer": "https://dashboard.wes.cl/",
}


def opaque_url(target_url: str, method: str = "GET", token: str = "") -> str:
    """Misma codificación que la app: btoa(unescape(encodeURIComponent(JSON)))."""
    payload = json.dumps(
        {"m": method, "u": target_url, "t": token or ""},
        separators=(",", ":"),
        ensure_ascii=False,
    )
    encoded = base64.b64encode(payload.encode("utf-8")).decode("ascii")
    return f"{BFF_URL}?d={urllib.parse.quote(encoded)}"


def bff_request(
    target_url: str,
    *,
    method: str = "GET",
    token: str = "",
    timeout: float = 90.0,
    retries: int = 3,
) -> Any:
    """GET/POST vía BFF. Devuelve JSON parseado o texto crudo."""
    url = opaque_url(target_url, method=method, token=token)
    last_err: Optional[Exception] = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers=dict(_DEFAULT_HEADERS), method="GET")
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read()
                text = raw.decode("utf-8", errors="replace")
                if not text:
                    return None
                try:
                    return json.loads(text)
                except json.JSONDecodeError:
                    return text
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace") if e.fp else ""
            last_err = RuntimeError(f"HTTP {e.code} {target_url}: {body[:200]}")
            if e.code in (429, 500, 502, 503, 504) and attempt + 1 < retries:
                time.sleep(1.5 * (attempt + 1))
                continue
            raise last_err from e
        except Exception as e:
            last_err = e
            if attempt + 1 < retries:
                time.sleep(1.5 * (attempt + 1))
                continue
            raise
    if last_err:
        raise last_err
    return None


def listar_companias() -> List[Dict[str, Any]]:
    data = bff_request(f"{ACL_ENTITIES}/configuration/companies")
    if not isinstance(data, list):
        raise RuntimeError(f"Respuesta inesperada companies: {type(data)}")
    return data


def auth_token() -> str:
    """Token de sesión dashboard (localStorage x-auth-token). Env ``WES_X_AUTH_TOKEN``."""
    return (
        os.environ.get("WES_X_AUTH_TOKEN", "").strip()
        or os.environ.get("WES_AUTH_TOKEN", "").strip()
    )


def measures_dates(
    node_ids: List[str],
    start_ddmmyyyy: str,
    end_ddmmyyyy: str,
    *,
    timeout: float = 120.0,
    token: str = "",
) -> List[Dict[str, Any]]:
    """Consumo diario por nodo (campo month[].totalM3)."""
    if not node_ids:
        return []
    qs = "&".join(f"id={urllib.parse.quote(nid)}" for nid in node_ids)
    url = (
        f"{ACL_NODE}/nodes/measures/dates?{qs}"
        f"&start={start_ddmmyyyy}&end={end_ddmmyyyy}"
    )
    data = bff_request(url, timeout=timeout, token=token or auth_token())
    if data is None:
        return []
    if isinstance(data, dict) and data.get("e"):
        raise RuntimeError(f"BFF error: {data}")
    if not isinstance(data, list):
        raise RuntimeError(f"Respuesta inesperada measures/dates: {type(data)}")
    return data


def dates_measures_csv(
    node_ids: List[str],
    start_ddmmyyyy: str,
    end_ddmmyyyy: str,
    *,
    timeout: float = 180.0,
    token: Optional[str] = None,
) -> str:
    """CSV horario de la app: GET nodes/dates.measures.csv (requiere x-auth-token)."""
    if not node_ids:
        return "TIME,VALUE\n"
    tok = auth_token() if token is None else token
    if not tok:
        raise RuntimeError(
            "CSV horario requiere WES_X_AUTH_TOKEN (x-auth-token del dashboard)."
        )
    qs = "&".join(f"id={urllib.parse.quote(nid)}" for nid in node_ids)
    url = (
        f"{ACL_NODE}/nodes/dates.measures.csv?{qs}"
        f"&start={start_ddmmyyyy}&end={end_ddmmyyyy}"
    )
    data = bff_request(url, timeout=timeout, token=tok)
    if data is None:
        return "TIME,VALUE\n"
    if isinstance(data, dict) and data.get("e"):
        raise RuntimeError(f"BFF error CSV: {data}")
    if not isinstance(data, str):
        raise RuntimeError(f"Respuesta CSV inesperada: {type(data)}")
    return data


def listar_nodos_companias(
    *,
    exclude_company_ids: Optional[set] = None,
) -> List[Dict[str, str]]:
    """Lista nodos por compañía vía BFF. Cada ítem: companyId, companyName, nodeId, nodeName."""
    excluded = exclude_company_ids or set()
    out: List[Dict[str, str]] = []
    for c in listar_companias():
        cid = str(c.get("companyId", "")).strip()
        cname = str(c.get("name", "")).strip()
        if not cid or cid in excluded:
            continue
        raw = bff_request(f"{ACL_ENTITIES}/companies/{cid}")
        nodes = raw.get("nodes") if isinstance(raw, dict) else None
        if not nodes:
            continue
        for n in nodes:
            nid = str(n.get("nodeId", "")).strip()
            nname = str(n.get("name", "")).strip()
            if nid:
                out.append(
                    {
                        "companyId": cid,
                        "companyName": cname,
                        "nodeId": nid,
                        "nodeName": nname or nid,
                    }
                )
    out.sort(key=lambda x: (x["companyId"], x["nodeId"]))
    return out
