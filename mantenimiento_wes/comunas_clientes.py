# -*- coding: utf-8 -*-
"""Comuna de cada cliente (y de cada sitio cuando el cliente abarca más de una)."""

from __future__ import annotations

import json
import unicodedata
from pathlib import Path
from typing import Any, Dict, Optional

ROOT = Path(__file__).resolve().parent
CATALOGO = ROOT / "catalogos" / "comunas_clientes.json"


def cargar_comunas(path: Optional[Path] = None) -> Dict[str, Any]:
    raw = (path or CATALOGO).read_text(encoding="utf-8")
    data = json.loads(raw)
    for key in ("por_cliente", "alias", "por_sitio", "reglas", "varias_comunas"):
        data.setdefault(key, {} if key != "reglas" and key != "varias_comunas" else [])
    return data


def _norm(texto: str) -> str:
    s = unicodedata.normalize("NFD", str(texto or ""))
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return " ".join(s.upper().split())


def comuna_de(cliente: str, maquina: str = "", catalogo: Optional[Dict[str, Any]] = None) -> str:
    """Devuelve la comuna conocida, o cadena vacía si no hay dato."""
    data = catalogo if catalogo is not None else cargar_comunas()
    cli = str(cliente or "").strip()
    maq = str(maquina or "").strip()
    if not cli:
        return ""
    alias = str((data.get("alias") or {}).get(cli) or cli)
    sitios = (data.get("por_sitio") or {}).get(cli) or (data.get("por_sitio") or {}).get(alias) or {}
    if maq and maq in sitios:
        return str(sitios[maq])
    maq_n = _norm(maq)
    for regla in data.get("reglas") or []:
        dueño = str(regla.get("cliente") or "")
        if dueño not in (cli, alias) or not maq_n:
            continue
        if _norm(regla.get("contiene")) in maq_n:
            return str(regla.get("comuna") or "")
    varias = set(data.get("varias_comunas") or [])
    if not maq and (cli in varias or alias in varias):
        return ""
    por = data.get("por_cliente") or {}
    return str(por.get(cli) or por.get(alias) or "")


def completar_comuna(data: Dict[str, Any]) -> str:
    """Rellena data['comuna'] solo si el técnico la dejó vacía. Respeta una corrección manual."""
    actual = str((data or {}).get("comuna") or "").strip()
    if actual:
        return actual
    comuna = comuna_de(str(data.get("cliente") or ""), str(data.get("maquina") or ""))
    if comuna:
        data["comuna"] = comuna
    return comuna
