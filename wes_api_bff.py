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


def measures_dates(
    node_ids: List[str],
    start_ddmmyyyy: str,
    end_ddmmyyyy: str,
    *,
    timeout: float = 120.0,
) -> List[Dict[str, Any]]:
    """Consumo diario por nodo (campo month[].totalM3)."""
    if not node_ids:
        return []
    qs = "&".join(f"id={urllib.parse.quote(nid)}" for nid in node_ids)
    url = (
        f"{ACL_NODE}/nodes/measures/dates?{qs}"
        f"&start={start_ddmmyyyy}&end={end_ddmmyyyy}"
    )
    data = bff_request(url, timeout=timeout)
    if data is None:
        return []
    if isinstance(data, dict) and data.get("e"):
        raise RuntimeError(f"BFF error: {data}")
    if not isinstance(data, list):
        raise RuntimeError(f"Respuesta inesperada measures/dates: {type(data)}")
    return data
