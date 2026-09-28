#!/usr/bin/env python3
"""Genera alertas propias del agente WES, una revisión por cliente.

No llama a las alertas de la plataforma. Lee el consumo horario y aplica las
reglas de ``alertas_agente.json``.

Reglas:
- consumo_fuera_de_horario: alguna hora de la ventana supera el umbral del agente.
- dia_sin_consumo: el día tiene lecturas y el total no pasa del mínimo.
- sin_datos: no hay serie horaria usable para ese día.

Ejemplos:
  python generar_alertas_agente.py
  python generar_alertas_agente.py --empresa "Club Providencia" --fecha 2026-09-27
  python generar_alertas_agente.py --sembrar
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable
from zoneinfo import ZoneInfo

import requests

from exclusiones_reportes import filter_node_ids, is_company_excluded

ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"
NODE_BASE_URL = "http://104.248.53.141:7003/wes/api/acl-node/v1"
# La API deja el nombre de UDD vacío; en los reportes del equipo es UDD.
NOMBRE_SI_API_VACIO = {"000026": "UDD"}
ZONA = ZoneInfo("America/Santiago")
ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "alertas_agente.json"
SALIDA_DIR = ROOT / "reports" / "alertas_agente"

REGLAS = ("consumo_fuera_de_horario", "dia_sin_consumo", "sin_datos")
HORAS_MINIMAS_LECTURA = 6


class AlertaAgenteError(Exception):
    """Error de configuración o de lectura al armar las alertas del agente."""


def normalizar(texto: str) -> str:
    base = unicodedata.normalize("NFD", texto or "")
    sin_acentos = "".join(ch for ch in base if unicodedata.category(ch) != "Mn")
    return " ".join(sin_acentos.casefold().split())


def config_por_defecto() -> dict[str, Any]:
    return {
        "descripcion": (
            "Alertas propias del agente WES. No usan las alertas de filtración o fuga de la plataforma."
        ),
        "umbral_m3h": 0.5,
        "ventana_horas": [0, 1, 2, 3, 4, 5, 6],
        "minimo_dia_m3": 0.01,
        "horas_minimas_para_dia": 18,
        "reglas": {regla: True for regla in REGLAS},
        "clientes": {},
    }


def cargar_config(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise AlertaAgenteError(f"No está el archivo de reglas: {path}")
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise AlertaAgenteError("alertas_agente.json tiene que ser un objeto.")
    base = config_por_defecto()
    base.update({k: v for k, v in data.items() if k != "reglas" and k != "clientes"})
    reglas = dict(base["reglas"])
    reglas.update(data.get("reglas") or {})
    base["reglas"] = {regla: bool(reglas.get(regla, True)) for regla in REGLAS}
    clientes = data.get("clientes") or {}
    if not isinstance(clientes, dict):
        raise AlertaAgenteError("El bloque clientes tiene que ser un objeto por companyId.")
    base["clientes"] = clientes
    base["ventana_horas"] = _ventana(base.get("ventana_horas"))
    base["umbral_m3h"] = float(base["umbral_m3h"])
    base["minimo_dia_m3"] = float(base["minimo_dia_m3"])
    base["horas_minimas_para_dia"] = int(base["horas_minimas_para_dia"])
    return base


def _ventana(valor: Any) -> list[int]:
    if not isinstance(valor, list) or not valor:
        raise AlertaAgenteError("ventana_horas tiene que listar horas entre 0 y 23.")
    horas: list[int] = []
    for item in valor:
        hora = int(item)
        if hora < 0 or hora > 23:
            raise AlertaAgenteError(f"Hora fuera de rango: {hora}")
        if hora not in horas:
            horas.append(hora)
    return horas


def guardar_config(path: Path, config: dict[str, Any]) -> None:
    path.write_text(json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


@dataclass(frozen=True)
class ReglaPunto:
    umbral_m3h: float
    ventana_horas: list[int]
    minimo_dia_m3: float
    horas_minimas_para_dia: int
    reglas: dict[str, bool]


def regla_para(config: dict[str, Any], company_id: str, node_id: str) -> ReglaPunto:
    cliente = (config.get("clientes") or {}).get(company_id) or {}
    punto = ((cliente.get("puntos") or {}).get(node_id) or {}) if isinstance(cliente, dict) else {}
    ventana = punto.get("ventana_horas") or cliente.get("ventana_horas") or config["ventana_horas"]
    umbral = punto.get("umbral_m3h", cliente.get("umbral_m3h", config["umbral_m3h"]))
    if umbral is None:
        umbral = config["umbral_m3h"]
    reglas = dict(config["reglas"])
    reglas.update(cliente.get("reglas") or {})
    reglas.update(punto.get("reglas") or {})
    return ReglaPunto(
        umbral_m3h=float(umbral),
        ventana_horas=_ventana(ventana),
        minimo_dia_m3=float(config["minimo_dia_m3"]),
        horas_minimas_para_dia=int(config["horas_minimas_para_dia"]),
        reglas={regla: bool(reglas.get(regla, False)) for regla in REGLAS},
    )


@dataclass(frozen=True)
class Hallazgo:
    fecha: str
    company_id: str
    cliente: str
    node_id: str
    punto: str
    regla: str
    detalle: str
    valor_m3h: float | None
    umbral_m3h: float | None


def evaluar_serie(horas: dict[int, float] | None, regla: ReglaPunto) -> list[tuple[str, str, float | None, float | None]]:
    """Devuelve (regla, detalle, valor, umbral) sin datos de cliente."""
    if not horas or len(horas) < HORAS_MINIMAS_LECTURA:
        if regla.reglas.get("sin_datos"):
            cantidad = 0 if not horas else len(horas)
            return [("sin_datos", f"Lecturas horarias: {cantidad}.", None, None)]
        return []

    hallazgos: list[tuple[str, str, float | None, float | None]] = []
    if regla.reglas.get("consumo_fuera_de_horario"):
        excesos = [
            (hora, float(horas[hora]))
            for hora in regla.ventana_horas
            if hora in horas and float(horas[hora]) > regla.umbral_m3h
        ]
        if excesos:
            detalle = "; ".join(f"{hora:02d}:00 ({valor:.3f} m³/h)" for hora, valor in excesos)
            hallazgos.append(
                (
                    "consumo_fuera_de_horario",
                    detalle,
                    max(valor for _, valor in excesos),
                    regla.umbral_m3h,
                )
            )

    if regla.reglas.get("dia_sin_consumo") and len(horas) >= regla.horas_minimas_para_dia:
        total = sum(float(v) for v in horas.values())
        if total <= regla.minimo_dia_m3:
            hallazgos.append(
                (
                    "dia_sin_consumo",
                    f"Total del día {total:.3f} m³ con {len(horas)} horas leídas.",
                    total,
                    regla.minimo_dia_m3,
                )
            )
    return hallazgos


def evaluar_punto(
    *,
    horas: dict[int, float] | None,
    regla: ReglaPunto,
    fecha: date,
    company_id: str,
    cliente: str,
    node_id: str,
    punto: str,
) -> list[Hallazgo]:
    salida = []
    for nombre, detalle, valor, umbral in evaluar_serie(horas, regla):
        salida.append(
            Hallazgo(
                fecha=fecha.isoformat(),
                company_id=company_id,
                cliente=cliente,
                node_id=node_id,
                punto=punto,
                regla=nombre,
                detalle=detalle,
                valor_m3h=valor,
                umbral_m3h=umbral,
            )
        )
    return salida


def listar_empresas(session: requests.Session | None = None) -> list[dict]:
    sesion = session or requests.Session()
    respuesta = sesion.get(f"{ENTITY_BASE_URL}/configuration/companies", timeout=30)
    if respuesta.status_code != 200:
        raise AlertaAgenteError(f"No pude listar empresas (HTTP {respuesta.status_code}).")
    data = respuesta.json()
    if not isinstance(data, list):
        raise AlertaAgenteError("La API no devolvió la lista de empresas.")
    return data


def puntos_de_empresa(empresa: dict) -> list[dict]:
    company_id = str(empresa.get("companyId") or "")
    nombre = str(empresa.get("name") or company_id)
    if is_company_excluded(company_id, nombre):
        return []
    nodos = []
    for nodo in empresa.get("nodes") or []:
        node_id = str(nodo.get("nodeId") or "").strip()
        if not node_id:
            continue
        nodos.append({"nodeId": node_id, "name": str(nodo.get("name") or node_id)})
    permitidos = set(filter_node_ids([n["nodeId"] for n in nodos], company_id, nombre))
    return [n for n in nodos if n["nodeId"] in permitidos]


def cliente_activo(config: dict[str, Any], company_id: str) -> bool:
    bloque = (config.get("clientes") or {}).get(company_id)
    if not isinstance(bloque, dict):
        return True
    return bool(bloque.get("activa", True))


def buscar_empresa(empresas: list[dict], consulta: str) -> dict:
    texto = consulta.strip()
    clave = normalizar(texto)
    por_id = [e for e in empresas if str(e.get("companyId")) == texto]
    if len(por_id) == 1:
        return por_id[0]
    exactas = [e for e in empresas if normalizar(str(e.get("name") or "")) == clave]
    if len(exactas) == 1:
        return exactas[0]
    parciales = [e for e in empresas if clave and clave in normalizar(str(e.get("name") or ""))]
    candidatos = exactas or parciales
    if len(candidatos) == 1:
        return candidatos[0]
    if not candidatos:
        raise AlertaAgenteError(f"No encontré la empresa {texto!r}.")
    lineas = [f"Hay varias empresas para {texto!r}:"]
    for empresa in candidatos:
        lineas.append(f"  - {empresa.get('companyId')}  {empresa.get('name')}")
    raise AlertaAgenteError("\n".join(lineas))


def sembrar_clientes(config: dict[str, Any], empresas: list[dict]) -> dict[str, Any]:
    """Agrega cada cliente operativo sin borrar umbrales ya editados."""
    clientes = dict(config.get("clientes") or {})
    for empresa in sorted(empresas, key=lambda e: str(e.get("companyId"))):
        company_id = str(empresa.get("companyId") or "")
        if not puntos_de_empresa(empresa):
            continue
        nombre = str(empresa.get("name") or "").strip() or NOMBRE_SI_API_VACIO.get(company_id, company_id)
        actual = clientes.get(company_id)
        if not isinstance(actual, dict):
            clientes[company_id] = {"nombre": nombre, "activa": True}
            continue
        actual = dict(actual)
        actual["nombre"] = nombre
        actual.setdefault("activa", True)
        clientes[company_id] = actual
    config = dict(config)
    config["clientes"] = dict(sorted(clientes.items()))
    return config


def horas_desde_csv(csv_content: str, dia: date) -> dict[int, float]:
    """Último m³/h por hora, con la marca TIME tal como la grafica la app."""
    ultimo: dict[str, float] = {}
    for linea in csv_content.strip().splitlines()[1:]:
        if not linea.strip():
            continue
        partes = linea.split(",", 1)
        if len(partes) < 2:
            continue
        try:
            ultimo[partes[0].strip()] = float(partes[1].strip().replace(" ", "").replace(",", "."))
        except ValueError:
            continue
    acc: dict[int, float] = {}
    prefijo = dia.isoformat()
    for marca in sorted(ultimo):
        if not marca.startswith(prefijo) or "T" not in marca:
            continue
        try:
            hora = int(marca[11:13])
        except ValueError:
            continue
        if 0 <= hora < 24:
            acc[hora] = ultimo[marca]
    return acc


def _dias_csv(dia: date) -> list[date]:
    inicio = datetime.combine(dia, datetime.min.time()).replace(tzinfo=ZONA)
    fin = inicio + timedelta(days=1) - timedelta(microseconds=1)
    cursor = inicio.astimezone(timezone.utc).date()
    ultimo = fin.astimezone(timezone.utc).date()
    dias: list[date] = []
    while cursor <= ultimo:
        dias.append(cursor)
        cursor += timedelta(days=1)
    return dias


def lector_horario_wes(node_id: str, dia: date) -> dict[int, float] | None:
    """Horas de la app para ese día. None si no hay lecturas."""
    acc: dict[int, float] = {}
    url = f"{NODE_BASE_URL}/nodes/{node_id}/dates.measures.csv"
    try:
        for dia_csv in _dias_csv(dia):
            respuesta = requests.get(
                url,
                params=[("start", dia_csv.strftime("%d%m%Y")), ("end", dia_csv.strftime("%d%m%Y"))],
                timeout=60,
            )
            if respuesta.status_code != 200 or not respuesta.text.strip():
                continue
            acc.update(horas_desde_csv(respuesta.text, dia))
    except requests.RequestException:
        return None
    if not acc:
        return None
    return {hora: float(acc.get(hora, 0.0)) for hora in range(24)}


def generar(
    config: dict[str, Any],
    empresas: list[dict],
    dia: date,
    lector: Callable[[str, date], dict[int, float] | None],
    company_id: str | None = None,
) -> list[Hallazgo]:
    hallazgos: list[Hallazgo] = []
    for empresa in empresas:
        cid = str(empresa.get("companyId") or "")
        if company_id and cid != company_id:
            continue
        if not cliente_activo(config, cid):
            continue
        puntos = puntos_de_empresa(empresa)
        if not puntos:
            continue
        nombre = str(empresa.get("name") or "").strip() or NOMBRE_SI_API_VACIO.get(cid, cid)
        bloque = (config.get("clientes") or {}).get(cid) or {}
        if isinstance(bloque, dict) and bloque.get("nombre"):
            nombre = str(bloque["nombre"])
        for punto in puntos:
            node_id = punto["nodeId"]
            regla = regla_para(config, cid, node_id)
            horas = lector(node_id, dia)
            hallazgos.extend(
                evaluar_punto(
                    horas=horas,
                    regla=regla,
                    fecha=dia,
                    company_id=cid,
                    cliente=nombre,
                    node_id=node_id,
                    punto=punto["name"],
                )
            )
    return hallazgos


def escribir_salida(hallazgos: list[Hallazgo], dia: date, destino: Path) -> Path:
    carpeta = destino / dia.isoformat()
    carpeta.mkdir(parents=True, exist_ok=True)
    por_cliente: dict[str, list[Hallazgo]] = {}
    for hallazgo in hallazgos:
        por_cliente.setdefault(hallazgo.company_id, []).append(hallazgo)

    campos = [
        "fecha",
        "cliente",
        "companyId",
        "punto",
        "nodeId",
        "regla",
        "detalle",
        "valor_m3h",
        "umbral_m3h",
    ]
    for company_id, filas in sorted(por_cliente.items()):
        ruta = carpeta / f"{company_id}.csv"
        with ruta.open("w", encoding="utf-8", newline="") as archivo:
            writer = csv.DictWriter(archivo, fieldnames=campos)
            writer.writeheader()
            for fila in filas:
                writer.writerow(
                    {
                        "fecha": fila.fecha,
                        "cliente": fila.cliente,
                        "companyId": fila.company_id,
                        "punto": fila.punto,
                        "nodeId": fila.node_id,
                        "regla": fila.regla,
                        "detalle": fila.detalle,
                        "valor_m3h": "" if fila.valor_m3h is None else f"{fila.valor_m3h:.3f}",
                        "umbral_m3h": "" if fila.umbral_m3h is None else f"{fila.umbral_m3h:.3f}",
                    }
                )

    resumen = carpeta / "resumen.txt"
    lineas = [f"Alertas del agente WES — {dia.isoformat()}", f"Hallazgos: {len(hallazgos)}", ""]
    if not hallazgos:
        lineas.append("Ningún cliente tuvo hallazgos con las reglas actuales.")
    else:
        for company_id, filas in sorted(por_cliente.items(), key=lambda item: item[1][0].cliente):
            lineas.append(f"{filas[0].cliente} ({company_id}) — {len(filas)}")
            for fila in filas:
                valor = "" if fila.valor_m3h is None else f" [{fila.valor_m3h:.3f}]"
                lineas.append(f"  {fila.node_id} {fila.punto} — {fila.regla}{valor}: {fila.detalle}")
            lineas.append("")
    resumen.write_text("\n".join(lineas).rstrip() + "\n", encoding="utf-8")
    return carpeta


def _ayer_chile() -> date:
    return datetime.now(ZONA).date() - timedelta(days=1)


def main(argv: list[str] | None = None) -> int:
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Genera alertas propias del agente WES por cliente.")
    parser.add_argument("--config", type=Path, default=CONFIG_PATH)
    parser.add_argument("--salida", type=Path, default=SALIDA_DIR)
    parser.add_argument("--fecha", help="Día a revisar, AAAA-MM-DD. Por defecto, ayer en Chile.")
    parser.add_argument("--empresa", help="Nombre o companyId. Si se omite, revisa todos los clientes activos.")
    parser.add_argument("--sembrar", action="store_true", help="Actualiza el catálogo de clientes en el JSON.")
    args = parser.parse_args(argv)

    try:
        config = cargar_config(args.config) if args.config.is_file() else config_por_defecto()
        empresas = listar_empresas()
        if args.sembrar:
            config = sembrar_clientes(config, empresas)
            guardar_config(args.config, config)
            print(f"Catálogo actualizado: {len(config['clientes'])} clientes en {args.config}")
            if not args.fecha and not args.empresa:
                return 0
        elegida = None
        if args.empresa:
            elegida = str(buscar_empresa(empresas, args.empresa).get("companyId"))
        dia = date.fromisoformat(args.fecha) if args.fecha else _ayer_chile()
        hallazgos = generar(config, empresas, dia, lector_horario_wes, elegida)
        carpeta = escribir_salida(hallazgos, dia, args.salida)
    except AlertaAgenteError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except (requests.RequestException, OSError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Día {dia.isoformat()}: {len(hallazgos)} alerta(s).")
    print(f"Salida: {carpeta}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
