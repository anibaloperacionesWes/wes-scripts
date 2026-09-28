#!/usr/bin/env python3
"""Configura alertas WES (filtración o fuga) de un cliente desde la nube.

La API de entidades no usa usuario ni contraseña. Sin ``--confirmar`` el
comando solo muestra el plan. Con ``--confirmar`` escribe en el servidor.

Ejemplos:
  python configurar_alerta.py listar-empresas
  python configurar_alerta.py listar-puntos --empresa "Parque Arauco"
  python configurar_alerta.py ver --empresa "Parque Arauco" --punto "000025-01"
  python configurar_alerta.py crear --empresa "Parque Arauco" --punto "000025-01" \\
      --tipo FILTRATION --correo tecnico@cliente.cl --umbral 30 --confirmar
"""

from __future__ import annotations

import argparse
import sys
import unicodedata
from typing import Any, Iterable

import requests

ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
TIPOS_ALERTA = ("FILTRATION", "LEAK")
TIMEOUT = 25


class AlertaError(Exception):
    """Error de uso o de la API al configurar una alerta."""


def normalizar(texto: str) -> str:
    base = unicodedata.normalize("NFD", texto or "")
    sin_acentos = "".join(ch for ch in base if unicodedata.category(ch) != "Mn")
    return " ".join(sin_acentos.casefold().split())


def parse_umbral(valor: str) -> str:
    texto = (valor or "").strip().replace(",", ".")
    if not texto:
        raise AlertaError("El umbral está vacío.")
    try:
        numero = float(texto)
    except ValueError as exc:
        raise AlertaError(f"Umbral inválido: {valor!r}. Usa un número, por ejemplo 30 o 2,5.") from exc
    if numero < 0:
        raise AlertaError("El umbral no puede ser negativo.")
    if numero.is_integer():
        return str(int(numero))
    return format(numero, "f").rstrip("0").rstrip(".")


def buscar_empresa(empresas: list[dict], consulta: str) -> dict:
    texto = (consulta or "").strip()
    if not texto:
        raise AlertaError("Falta el nombre o el ID de la empresa.")
    clave = normalizar(texto)
    por_id = [e for e in empresas if str(e.get("companyId", "")).strip() == texto]
    if len(por_id) == 1:
        return por_id[0]
    exactas = [e for e in empresas if normalizar(str(e.get("name", ""))) == clave]
    if len(exactas) == 1:
        return exactas[0]
    if len(exactas) > 1:
        raise AlertaError(_mensaje_ambiguo("empresa", texto, exactas))
    parciales = [e for e in empresas if clave in normalizar(str(e.get("name", "")))]
    if len(parciales) == 1:
        return parciales[0]
    if not parciales:
        raise AlertaError(f"No encontré la empresa {texto!r}. Usa listar-empresas para ver los nombres.")
    raise AlertaError(_mensaje_ambiguo("empresa", texto, parciales))


def buscar_puntos(nodos: list[dict], consulta: str) -> list[dict]:
    texto = (consulta or "").strip()
    if not texto:
        raise AlertaError("Falta el nombre o el ID del punto.")
    clave = normalizar(texto)
    por_id = [n for n in nodos if str(n.get("nodeId", "")).strip() == texto]
    if len(por_id) == 1:
        return por_id
    if len(por_id) > 1:
        raise AlertaError(_mensaje_ambiguo("punto", texto, por_id, id_key="nodeId"))
    exactas = [n for n in nodos if normalizar(str(n.get("name", ""))) == clave]
    if len(exactas) == 1:
        return exactas
    if len(exactas) > 1:
        raise AlertaError(_mensaje_ambiguo("punto", texto, exactas, id_key="nodeId"))
    parciales = [n for n in nodos if clave in normalizar(str(n.get("name", "")))]
    if len(parciales) == 1:
        return parciales
    if not parciales:
        raise AlertaError(f"No encontré el punto {texto!r} en esa empresa.")
    raise AlertaError(_mensaje_ambiguo("punto", texto, parciales, id_key="nodeId"))


def _mensaje_ambiguo(etiqueta: str, consulta: str, items: list[dict], id_key: str = "companyId") -> str:
    lineas = [f"Hay varias coincidencias de {etiqueta} para {consulta!r}:"]
    for item in items:
        nombre = item.get("name") or ""
        lineas.append(f"  - {item.get(id_key)}  {nombre}")
    lineas.append("Indica el ID exacto.")
    return "\n".join(lineas)


def _nodos(empresa: dict) -> list[dict]:
    nodos = empresa.get("nodes") or []
    return [n for n in nodos if isinstance(n, dict)]


def umbral_de(nodo: dict) -> str | None:
    config = nodo.get("configuration") or {}
    valor = config.get("threshold")
    if valor is None or str(valor).strip() == "":
        return None
    return str(valor)


class ClienteAlertas:
    def __init__(self, base_url: str = ENTITY_BASE_URL, session: requests.Session | None = None):
        self.base_url = base_url.rstrip("/")
        self.session = session or requests.Session()
        self.session.headers.update({"Accept": "application/json"})

    def listar_empresas(self) -> list[dict]:
        url = f"{self.base_url}/configuration/companies"
        respuesta = self.session.get(url, timeout=TIMEOUT)
        self._exigir_ok(respuesta, "No pude listar las empresas")
        data = respuesta.json()
        if not isinstance(data, list):
            raise AlertaError("La API no devolvió una lista de empresas.")
        return data

    def obtener_alerta(self, company_id: str, node_id: str, tipo: str) -> dict | None:
        tipo = _tipo(tipo)
        url = f"{self.base_url}/companies/{company_id}/node/{node_id}/alert/{tipo}/information"
        respuesta = self.session.get(url, timeout=TIMEOUT)
        if respuesta.status_code == 404:
            return None
        self._exigir_ok(respuesta, f"No pude leer la alerta {tipo} de {node_id}")
        data = respuesta.json()
        if not isinstance(data, dict):
            raise AlertaError(f"Respuesta inesperada al leer la alerta de {node_id}.")
        return data

    def crear_alerta(self, company_id: str, node_id: str, tipo: str, correos: list[str]) -> dict:
        return self._escribir_alerta("post", company_id, node_id, tipo, correos, "crear la alerta")

    def agregar_correos(self, company_id: str, node_id: str, tipo: str, correos: list[str]) -> dict:
        return self._escribir_alerta("put", company_id, node_id, tipo, correos, "agregar correos")

    def quitar_correos(self, company_id: str, node_id: str, tipo: str, correos: list[str]) -> dict:
        return self._escribir_alerta("delete", company_id, node_id, tipo, correos, "quitar correos")

    def definir_umbral(self, company_id: str, node_id: str, umbral: str) -> Any:
        url = f"{self.base_url}/configuration/companies/{company_id}/nodes/{node_id}/threshold"
        respuesta = self.session.put(
            url,
            json={"threshold": umbral},
            headers={"Content-Type": "application/json"},
            timeout=TIMEOUT,
        )
        self._exigir_ok(respuesta, f"No pude guardar el umbral de {node_id}")
        if not respuesta.content:
            return {"threshold": umbral}
        try:
            return respuesta.json()
        except ValueError:
            return {"threshold": umbral, "raw": respuesta.text}

    def _escribir_alerta(
        self,
        metodo: str,
        company_id: str,
        node_id: str,
        tipo: str,
        correos: list[str],
        accion: str,
    ) -> dict:
        tipo = _tipo(tipo)
        correos = _correos(correos)
        url = f"{self.base_url}/configuration/companies/{company_id}/alert/nodes/{node_id}"
        cuerpo = {"alertType": tipo, "notifyTo": correos}
        respuesta = self.session.request(
            metodo.upper(),
            url,
            json=cuerpo,
            headers={"Content-Type": "application/json"},
            timeout=TIMEOUT,
        )
        self._exigir_ok(respuesta, f"No pude {accion} en {node_id}")
        if not respuesta.content:
            return {"alertType": tipo, "notifyTo": correos}
        try:
            data = respuesta.json()
        except ValueError:
            return {"alertType": tipo, "notifyTo": correos, "raw": respuesta.text}
        return data if isinstance(data, dict) else {"alertType": tipo, "notifyTo": correos}

    @staticmethod
    def _exigir_ok(respuesta: requests.Response, mensaje: str) -> None:
        if 200 <= respuesta.status_code < 300:
            return
        detalle = respuesta.text.strip().replace("\n", " ")
        if len(detalle) > 400:
            detalle = detalle[:400] + "..."
        raise AlertaError(f"{mensaje} (HTTP {respuesta.status_code}). {detalle}")


def _tipo(tipo: str) -> str:
    valor = (tipo or "FILTRATION").strip().upper()
    alias = {
        "FILTRACION": "FILTRATION",
        "FILTRACIÓN": "FILTRATION",
        "FUGA": "LEAK",
        "LEAK": "LEAK",
        "FILTRATION": "FILTRATION",
    }
    if valor not in alias:
        raise AlertaError("El tipo debe ser FILTRATION (filtración / consumo nocturno) o LEAK (fuga).")
    return alias[valor]


def _correos(correos: Iterable[str]) -> list[str]:
    limpios: list[str] = []
    vistos: set[str] = set()
    for correo in correos:
        for parte in str(correo).replace(";", ",").split(","):
            email = parte.strip()
            if not email:
                continue
            if "@" not in email or email.startswith("@") or email.endswith("@"):
                raise AlertaError(f"Correo inválido: {email}")
            clave = email.casefold()
            if clave in vistos:
                continue
            vistos.add(clave)
            limpios.append(email)
    if not limpios:
        raise AlertaError("Indica al menos un correo con --correo.")
    return limpios


def _correos_actuales(alerta: dict | None) -> list[str]:
    if not alerta:
        return []
    salida = []
    for receptor in alerta.get("receiverList") or []:
        if isinstance(receptor, dict) and receptor.get("email"):
            salida.append(str(receptor["email"]))
        elif isinstance(receptor, str):
            salida.append(receptor)
    return salida


def _imprimir_alerta(tipo: str, alerta: dict | None, umbral: str | None) -> None:
    print(f"  Tipo: {tipo}")
    print(f"  Umbral del punto: {umbral if umbral is not None else 'sin umbral'}")
    if alerta is None:
        print("  Alerta: no configurada")
        return
    print(f"  Lugar: {alerta.get('venueName') or '—'}")
    correos = _correos_actuales(alerta)
    if correos:
        print("  Avisa a:")
        for correo in correos:
            print(f"    - {correo}")
    else:
        print("  Avisa a: nadie")


def _resolver(cliente: ClienteAlertas, empresa_q: str, punto_q: str | None, todos: bool) -> tuple[dict, list[dict]]:
    empresa = buscar_empresa(cliente.listar_empresas(), empresa_q)
    nodos = _nodos(empresa)
    if todos:
        if not nodos:
            raise AlertaError(f"{empresa.get('name')} no tiene puntos.")
        return empresa, nodos
    if not punto_q:
        raise AlertaError("Indica --punto o --todos-los-puntos.")
    return empresa, buscar_puntos(nodos, punto_q)


def cmd_listar_empresas(cliente: ClienteAlertas, _args: argparse.Namespace) -> int:
    empresas = cliente.listar_empresas()
    print(f"{len(empresas)} empresas")
    for empresa in sorted(empresas, key=lambda e: str(e.get("companyId", ""))):
        print(f"{empresa.get('companyId')}  {empresa.get('name')}  ({len(_nodos(empresa))} puntos)")
    return 0


def cmd_listar_puntos(cliente: ClienteAlertas, args: argparse.Namespace) -> int:
    empresa = buscar_empresa(cliente.listar_empresas(), args.empresa)
    nodos = _nodos(empresa)
    print(f"{empresa.get('companyId')}  {empresa.get('name')}  ({len(nodos)} puntos)")
    for nodo in nodos:
        umbral = umbral_de(nodo)
        extra = f"  umbral {umbral}" if umbral is not None else ""
        print(f"  {nodo.get('nodeId')}  {nodo.get('name')}{extra}")
    return 0


def cmd_ver(cliente: ClienteAlertas, args: argparse.Namespace) -> int:
    empresa, puntos = _resolver(cliente, args.empresa, args.punto, args.todos_los_puntos)
    tipos = [_tipo(args.tipo)] if args.tipo else list(TIPOS_ALERTA)
    print(f"{empresa.get('companyId')}  {empresa.get('name')}")
    for nodo in puntos:
        print(f"\n{nodo.get('nodeId')}  {nodo.get('name')}")
        for tipo in tipos:
            alerta = cliente.obtener_alerta(str(empresa["companyId"]), str(nodo["nodeId"]), tipo)
            _imprimir_alerta(tipo, alerta, umbral_de(nodo))
    return 0


def cmd_crear(cliente: ClienteAlertas, args: argparse.Namespace) -> int:
    tipo = _tipo(args.tipo)
    correos = _correos(args.correo or [])
    umbral = parse_umbral(args.umbral) if args.umbral else None
    empresa, puntos = _resolver(cliente, args.empresa, args.punto, args.todos_los_puntos)
    company_id = str(empresa["companyId"])
    print(f"Empresa: {empresa.get('name')} ({company_id})")
    print(f"Tipo: {tipo}")
    print("Correos: " + ", ".join(correos))
    if umbral is not None:
        print(f"Umbral: {umbral}")
    print(f"Puntos: {len(puntos)}")

    planes = []
    for nodo in puntos:
        node_id = str(nodo["nodeId"])
        actual = cliente.obtener_alerta(company_id, node_id, tipo)
        ya = {c.casefold() for c in _correos_actuales(actual)}
        nuevos = [c for c in correos if c.casefold() not in ya]
        planes.append((nodo, actual, nuevos))
        estado = "ya existe" if actual else "se crea"
        print(f"  - {node_id} {nodo.get('name')}: {estado}" + (f", suma {len(nuevos)} correo(s)" if actual else ""))

    if not args.confirmar:
        print("\nPlan listo. Vuelve a ejecutar con --confirmar para aplicarlo en WES.")
        return 0

    for nodo, actual, nuevos in planes:
        node_id = str(nodo["nodeId"])
        if actual is None:
            cliente.crear_alerta(company_id, node_id, tipo, correos)
            print(f"Creada {tipo} en {node_id}.")
        elif nuevos:
            cliente.agregar_correos(company_id, node_id, tipo, nuevos)
            print(f"Sumé correos en {node_id}: {', '.join(nuevos)}")
        else:
            print(f"{node_id} ya avisaba a esos correos.")
        if umbral is not None and umbral_de(nodo) != umbral:
            cliente.definir_umbral(company_id, node_id, umbral)
            print(f"Umbral de {node_id}: {umbral}")
    print("Listo.")
    return 0


def cmd_correos(cliente: ClienteAlertas, args: argparse.Namespace, quitar: bool) -> int:
    tipo = _tipo(args.tipo)
    correos = _correos(args.correo or [])
    empresa, puntos = _resolver(cliente, args.empresa, args.punto, args.todos_los_puntos)
    company_id = str(empresa["companyId"])
    verbo = "Quitar" if quitar else "Agregar"
    print(f"{verbo} en {empresa.get('name')} ({company_id}), tipo {tipo}: {', '.join(correos)}")
    if not args.confirmar:
        for nodo in puntos:
            print(f"  - {nodo.get('nodeId')}  {nodo.get('name')}")
        print("\nPlan listo. Vuelve a ejecutar con --confirmar para aplicarlo en WES.")
        return 0
    for nodo in puntos:
        node_id = str(nodo["nodeId"])
        if quitar:
            cliente.quitar_correos(company_id, node_id, tipo, correos)
        else:
            cliente.agregar_correos(company_id, node_id, tipo, correos)
        print(f"Actualicé {node_id}.")
    return 0


def cmd_umbral(cliente: ClienteAlertas, args: argparse.Namespace) -> int:
    umbral = parse_umbral(args.umbral)
    empresa, puntos = _resolver(cliente, args.empresa, args.punto, args.todos_los_puntos)
    company_id = str(empresa["companyId"])
    print(f"Umbral {umbral} en {empresa.get('name')} ({company_id})")
    for nodo in puntos:
        print(f"  - {nodo.get('nodeId')}  {nodo.get('name')}  actual: {umbral_de(nodo) or 'sin umbral'}")
    if not args.confirmar:
        print("\nPlan listo. Vuelve a ejecutar con --confirmar para aplicarlo en WES.")
        return 0
    for nodo in puntos:
        node_id = str(nodo["nodeId"])
        cliente.definir_umbral(company_id, node_id, umbral)
        print(f"Umbral de {node_id}: {umbral}")
    return 0


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Configura alertas WES de un cliente.")
    parser.add_argument("--api", default=ENTITY_BASE_URL, help="Base de acl-entities.")
    sub = parser.add_subparsers(dest="comando", required=True)

    sub.add_parser("listar-empresas", help="Lista empresas y cuántos puntos tiene cada una.")

    puntos = sub.add_parser("listar-puntos", help="Lista los puntos de una empresa y su umbral.")
    puntos.add_argument("--empresa", required=True, help="Nombre o companyId.")

    ver = sub.add_parser("ver", help="Muestra alertas y umbral de uno o todos los puntos.")
    _args_objetivo(ver)
    ver.add_argument("--tipo", help="FILTRATION o LEAK. Si se omite, muestra ambos.")

    crear = sub.add_parser("crear", help="Crea la alerta o suma correos si ya existe.")
    _args_objetivo(crear)
    crear.add_argument("--tipo", default="FILTRATION", help="FILTRATION (por defecto) o LEAK.")
    crear.add_argument("--correo", action="append", required=True, help="Se puede repetir. Acepta lista separada por coma.")
    crear.add_argument("--umbral", help="Umbral del punto. Ej: 30 o 2,5.")
    crear.add_argument("--confirmar", action="store_true", help="Aplica el cambio en WES.")

    agregar = sub.add_parser("agregar-correos", help="Suma destinatarios a una alerta existente.")
    _args_correos(agregar)
    quitar = sub.add_parser("quitar-correos", help="Saca destinatarios de una alerta existente.")
    _args_correos(quitar)

    umbral = sub.add_parser("umbral", help="Define el umbral de alerta del punto.")
    _args_objetivo(umbral)
    umbral.add_argument("--umbral", required=True)
    umbral.add_argument("--confirmar", action="store_true")
    return parser


def _args_objetivo(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--empresa", required=True, help="Nombre o companyId.")
    grupo = parser.add_mutually_exclusive_group(required=True)
    grupo.add_argument("--punto", help="Nombre o nodeId.")
    grupo.add_argument("--todos-los-puntos", action="store_true", help="Aplica a todos los puntos de la empresa.")


def _args_correos(parser: argparse.ArgumentParser) -> None:
    _args_objetivo(parser)
    parser.add_argument("--tipo", default="FILTRATION")
    parser.add_argument("--correo", action="append", required=True)
    parser.add_argument("--confirmar", action="store_true")


def main(argv: list[str] | None = None) -> int:
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    args = _parser().parse_args(argv)
    cliente = ClienteAlertas(base_url=args.api)
    try:
        if args.comando == "listar-empresas":
            return cmd_listar_empresas(cliente, args)
        if args.comando == "listar-puntos":
            return cmd_listar_puntos(cliente, args)
        if args.comando == "ver":
            return cmd_ver(cliente, args)
        if args.comando == "crear":
            return cmd_crear(cliente, args)
        if args.comando == "agregar-correos":
            return cmd_correos(cliente, args, quitar=False)
        if args.comando == "quitar-correos":
            return cmd_correos(cliente, args, quitar=True)
        if args.comando == "umbral":
            return cmd_umbral(cliente, args)
    except AlertaError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except requests.RequestException as exc:
        print(f"Error de red: {exc}", file=sys.stderr)
        return 1
    print(f"Comando no implementado: {args.comando}", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
