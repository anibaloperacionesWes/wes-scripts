# -*- coding: utf-8 -*-
"""
Elimina un usuario WES y lo crea de nuevo con otro nombre, mismos
username / companyId / allowedNodes.

La API no tiene PUT de nombre. Recrear cambia el userId y deja la clave
para setearla en dashboard.wes.cl (no usar POST /users/changePassword).

Uso:
  python recrear_usuario.py --email an_ambiental_pae@linkes.cl --nombre Sergio --apellido Fuenzalida
"""

from __future__ import annotations

import argparse
import json
import sys

import requests

ENTITY_BASE = "http://104.248.53.141:7001/wes/api/acl-entities/v1"


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"Accept": "application/json", "Content-Type": "application/json"})
    return s


def _sin_password(user: dict) -> dict:
    out = dict(user or {})
    out.pop("password", None)
    return out


def _get_por_email(session: requests.Session, email: str) -> dict | None:
    r = session.get(f"{ENTITY_BASE}/users", params={"email": email}, timeout=25)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    data = r.json()
    if not isinstance(data, dict):
        raise RuntimeError(f"Respuesta inesperada GET /users: {data!r}")
    return data


def recrear(email: str, nombre: str, apellido: str) -> dict:
    session = _session()
    actual = _get_por_email(session, email)
    if not actual:
        raise SystemExit(f"No existe el usuario {email}")

    user_id = str(actual.get("userId") or "").strip()
    company_id = str(actual.get("companyId") or "").strip()
    username = str(actual.get("username") or email).strip()
    nodes = list(actual.get("allowedNodes") or [])
    if not user_id or not company_id or not username:
        raise SystemExit(f"Faltan campos en el usuario: {json.dumps(_sin_password(actual), ensure_ascii=False)}")

    print("[INFO] Usuario actual:")
    print(json.dumps(_sin_password(actual), indent=2, ensure_ascii=False))
    print(f"[INFO] Eliminando userId={user_id} ...")
    d = session.delete(f"{ENTITY_BASE}/configuration/users/{user_id}", timeout=25)
    print(f"[INFO] DELETE {d.status_code} {d.text[:200]!r}")
    d.raise_for_status()

    if _get_por_email(session, email):
        raise SystemExit("El usuario sigue existiendo después del DELETE")

    body = {
        "username": username,
        "name": nombre,
        "lastName": apellido,
        "companyId": company_id,
        "allowedNodes": nodes,
    }
    print("[INFO] Creando usuario con:")
    print(json.dumps(body, indent=2, ensure_ascii=False))
    c = session.post(
        f"{ENTITY_BASE}/configuration/users",
        json=body,
        timeout=25,
        headers={"Accept": "*/*", "Content-Type": "application/json"},
    )
    print(f"[INFO] POST {c.status_code} {c.text[:300]!r}")
    if c.status_code >= 400:
        raise SystemExit(
            "Falló el POST de creación. El usuario quedó eliminado; "
            "reintentar este script con los mismos datos."
        )

    nuevo = _get_por_email(session, email)
    if not nuevo:
        raise SystemExit("Se creó pero GET /users?email= no lo encuentra")
    print("[OK] Usuario recreado:")
    print(json.dumps(_sin_password(nuevo), indent=2, ensure_ascii=False))
    print("[AVISO] La clave hay que setearla en https://dashboard.wes.cl")
    return nuevo


def main() -> int:
    parser = argparse.ArgumentParser(description="Eliminar y recrear un usuario WES cambiando el nombre")
    parser.add_argument("--email", required=True)
    parser.add_argument("--nombre", required=True)
    parser.add_argument("--apellido", required=True)
    args = parser.parse_args()
    recrear(args.email.strip(), args.nombre.strip(), args.apellido.strip())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
