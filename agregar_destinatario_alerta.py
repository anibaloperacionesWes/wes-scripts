# -*- coding: utf-8 -*-
"""
Agrega un correo como destinatario de alerta de filtración (umbral diario).

La app WES hace PUT sobre la alerta. Si el nodo nunca tuvo alerta, esa llamada
responde HTTP 404 ("error 4"). Este script:

  - POST  si no existe alerta  → la crea con el correo
  - PUT   si ya existe         → agrega el correo sin borrar destinatarios previos

Uso:
  python agregar_destinatario_alerta.py
  python agregar_destinatario_alerta.py --company-id 000025 --email anibal.aoperaciones@wes.cl
"""

from __future__ import annotations

import argparse
import sys
from typing import Any

import requests

ENTITY_BASE = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
EMAIL_DEFAULT = "anibal.aoperaciones@wes.cl"
COMPANY_DEFAULT = "000025"

# Puntos fuera de operación / no se les carga alerta de umbral.
PA_FUERA = frozenset(
    {
        "000025-02",
        "000025-03",
        "000025-05",
        "000025-06",
        "000025-11",
        "000025-14",
        "000025-15",
        "000025-16",
        "000025-25",
        "000025-26",
    }
)


def _threshold_ok(raw: Any) -> bool:
    if raw is None:
        return False
    s = str(raw).strip().replace(",", ".")
    if not s:
        return False
    try:
        return float(s) > 0
    except ValueError:
        return False


def _emails_alerta(company_id: str, node_id: str) -> tuple[int, list[str]]:
    url = (
        f"{ENTITY_BASE}/companies/{company_id}/node/{node_id}"
        f"/alert/FILTRATION/information"
    )
    r = requests.get(url, timeout=30)
    if r.status_code == 404:
        return 404, []
    r.raise_for_status()
    data = r.json() or {}
    emails = [
        str(x.get("email") or "").strip().lower()
        for x in (data.get("receiverList") or [])
        if str(x.get("email") or "").strip()
    ]
    return 200, emails


def _agregar(company_id: str, node_id: str, email: str, *, crear: bool) -> tuple[int, str]:
    url = f"{ENTITY_BASE}/configuration/companies/{company_id}/alert/nodes/{node_id}"
    body = {"alertType": "FILTRATION", "notifyTo": [email]}
    fn = requests.post if crear else requests.put
    r = fn(url, json=body, timeout=30)
    return r.status_code, (r.text or "")[:300]


def listar_nodos(company_id: str) -> list[dict]:
    r = requests.get(f"{ENTITY_BASE}/companies/{company_id}", timeout=60)
    r.raise_for_status()
    rows = []
    for node in r.json().get("nodes") or []:
        nid = str(node.get("nodeId") or "").strip()
        if nid:
            rows.append(node)
    rows.sort(key=lambda n: n.get("nodeId") or "")
    return rows


def procesar(company_id: str, email: str) -> int:
    email_l = email.strip().lower()
    nodos = listar_nodos(company_id)
    ok = skip = fail = 0
    print(f"Compañía {company_id}  destinatario {email}")
    print("-" * 72)
    for n in nodos:
        nid = str(n.get("nodeId") or "").strip()
        name = str(n.get("name") or nid)
        th = (n.get("configuration") or {}).get("threshold")
        if nid in PA_FUERA:
            print(f"  SKIP  {nid:<12} fuera de operación  {name}")
            skip += 1
            continue
        if not _threshold_ok(th):
            print(f"  SKIP  {nid:<12} umbral={th!r} (no operativo)  {name}")
            skip += 1
            continue
        st, emails = _emails_alerta(company_id, nid)
        if st == 200 and email_l in emails:
            print(f"  OK    {nid:<12} ya estaba  umbral={th}  {name}")
            ok += 1
            continue
        crear = st == 404
        code, body = _agregar(company_id, nid, email, crear=crear)
        accion = "POST crea" if crear else "PUT agrega"
        if code == 200:
            print(f"  OK    {nid:<12} {accion}  umbral={th}  {name}")
            ok += 1
        else:
            print(f"  FAIL  {nid:<12} {accion} HTTP {code}  {name}")
            print(f"        {body}")
            fail += 1
    print("-" * 72)
    print(f"Listos: {ok}   saltados: {skip}   fallidos: {fail}")
    return 1 if fail else 0


def main() -> int:
    if sys.platform == "win32":
        for s in (sys.stdout, sys.stderr):
            try:
                s.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
            except Exception:
                pass
    p = argparse.ArgumentParser(description="Agregar destinatario a alerta de umbral")
    p.add_argument("--company-id", default=COMPANY_DEFAULT)
    p.add_argument("--email", default=EMAIL_DEFAULT)
    args = p.parse_args()
    return procesar(args.company_id.strip(), args.email.strip())


if __name__ == "__main__":
    raise SystemExit(main())
