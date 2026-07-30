"""
Consulta qué usuarios tienen nodos 000025-* en allowedNodes vía API acl-entities.

GET {ENTITY}/users?email=... devuelve allowedNodes por usuario. No existe en esta API
un listado público de "todos los usuarios de la empresa 000025"; por eso hay que indicar
qué correos revisar (--emails o --contactos-pa).

Salida: CSV con columnas ID_nodo, Nombre_punto, usuarios (emails con acceso a ese nodo).
"""

from __future__ import annotations

import argparse
import csv
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Set

import requests

ENTITY_BASE = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
COMPANY_PA = "000025"
PREFIX_PA = "000025-"


def _fetch_user_by_email(email: str) -> dict | None:
    r = requests.get(
        f"{ENTITY_BASE}/users",
        params={"email": email.strip()},
        timeout=30,
    )
    if r.status_code == 404:
        return None
    r.raise_for_status()
    data = r.json()
    if isinstance(data, dict):
        data.pop("password", None)
    return data


def _nodos_pa_desde_empresa() -> Dict[str, str]:
    r = requests.get(f"{ENTITY_BASE}/companies/{COMPANY_PA}", timeout=60)
    r.raise_for_status()
    out: Dict[str, str] = {}
    for n in r.json().get("nodes") or []:
        nid = str(n.get("nodeId", "")).strip()
        if nid.startswith(PREFIX_PA):
            out[nid] = str(n.get("name", "") or "").strip()
    return dict(sorted(out.items()))


def _emails_desde_contactos_pa() -> List[str]:
    from lista_contactos_reportes import CONTACTOS_REPORTES

    emails: Set[str] = set()
    for _k, c in CONTACTOS_REPORTES.items():
        emp = c.get("empresas_interes") or []
        if not isinstance(emp, list):
            continue
        if "Parque Arauco" not in emp:
            continue
        e = (c.get("email") or "").strip()
        if e:
            emails.add(e)
    return sorted(emails)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Mapeo nodo PA -> usuarios con allowedNodes (consulta por email)"
    )
    parser.add_argument(
        "--emails",
        type=Path,
        help="Archivo de texto: un correo por línea (además de los de --contactos-pa si se usa)",
    )
    parser.add_argument(
        "--contactos-pa",
        action="store_true",
        help="Incluir correos de lista_contactos_reportes con interés Parque Arauco",
    )
    parser.add_argument(
        "--salida",
        type=Path,
        default=Path("reports")
        / "Parque_Arauco"
        / "Reportes_agregados_reunion_abril_2026"
        / "usuarios_habilitados_por_nodo_pa.csv",
        help="CSV de salida",
    )
    args = parser.parse_args()

    emails: Set[str] = set()
    if args.contactos_pa:
        emails.update(_emails_desde_contactos_pa())
    if args.emails and args.emails.exists():
        for line in args.emails.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                emails.add(line)

    if not emails:
        print(
            "[ERROR] Indique correos: --contactos-pa y/o --emails archivo.txt\n"
            "Ejemplo: python consultar_usuarios_habilitados_pa.py --contactos-pa"
        )
        return 1

    nodos_nombre = _nodos_pa_desde_empresa()
    # node_id -> set of emails that have this node in allowedNodes
    por_nodo: Dict[str, Set[str]] = defaultdict(set)

    for email in sorted(emails):
        try:
            u = _fetch_user_by_email(email)
        except requests.RequestException as e:
            print(f"[WARN] {email}: {e}")
            continue
        if not u:
            print(f"[INFO] Sin usuario API: {email}")
            continue
        allowed = u.get("allowedNodes") or []
        if not isinstance(allowed, list):
            continue
        nombre = f"{u.get('name', '')} {u.get('lastName', '')}".strip()
        print(f"[OK] {email} ({nombre}) — {len(allowed)} nodos en allowedNodes")
        for nid in allowed:
            if isinstance(nid, str) and nid.startswith(PREFIX_PA):
                por_nodo[nid].add(email)

    args.salida.parent.mkdir(parents=True, exist_ok=True)
    with open(args.salida, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["ID_nodo", "Nombre_punto", "Usuarios_con_acceso"])
        for nid in sorted(nodos_nombre.keys()):
            users = sorted(por_nodo.get(nid, set()))
            w.writerow([nid, nodos_nombre.get(nid, ""), ", ".join(users) if users else ""])

    sin_listar = [nid for nid in nodos_nombre if nid not in por_nodo]
    print(f"\n[OK] CSV: {args.salida.resolve()}")
    print(
        f"[INFO] Nodos sin ningún usuario en la muestra consultada: {len(sin_listar)} "
        f"(corre la consulta con más emails o revisa backoffice WES para el listado completo)."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
