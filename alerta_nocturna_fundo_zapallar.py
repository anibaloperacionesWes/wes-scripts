"""
Alertas de Fundo Zapallar con tres niveles: seguimiento, preventiva y crítica.

No usa el promedio de noches. La ventana nocturna es 00:00 a 06:00 (horas 00 a 06).
La crítica queda para consumos excepcionalmente altos. Si un seguimiento se repite
al día siguiente, pasa a preventiva.

Reglas nocturnas:
- Etapa N°5: seguimiento >1 a 5 m³; preventiva >5 a 10 m³; crítica >10 m³ o >4 m³/h.
- Etapas 1-4: seguimiento >40 a 45 m³ o >12 a 15 m³/h; preventiva >45 a 60 m³ o
  >15 a 20 m³/h; crítica >60 m³ o >20 m³/h.
- Matriz ESVAL: seguimiento >55 a 65 m³ o >25 a 30 m³/h; preventiva >65 a 80 m³ o
  >30 a 35 m³/h; crítica >80 m³ o >35 m³/h.

Balances diarios (valor absoluto):
- Matriz ESVAL − Estanque Inferior: seguimiento >21 a 25 m³; preventiva >25 a 42 m³
  o >21 m³ durante 2 días; crítica >42 m³.
- Estanque Inferior − (Etapa N°5 + Etapas 1-4): seguimiento >30 a 36 m³; preventiva
  >36 a 60 m³ o >30 m³ durante 2 días; crítica >60 m³.

El aviso a Aníbal y Juan indica el punto, la regla, el valor, el umbral y el período.
"""

from __future__ import annotations

import argparse
import base64
import csv
import json
import os
import smtplib
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, make_msgid
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
from zoneinfo import ZoneInfo

import requests

from config_correos_equipo import obtener_email, obtener_nombre

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

CHILE_TZ = ZoneInfo("America/Santiago")
BASE_URL = "http://104.248.53.141:7003/wes/api/acl-node/v1"
# La app entra por api-rest.wes.cl. El puerto 7003 de la IP pública suele no responder desde este agente.
APP_API = "https://api-rest.wes.cl/wes/api/acl-node/v1"
APP_LOGIN = "https://api-rest.wes.cl/wes/api/wes-mc-api/v1/login"
APP_BFF = "https://api-rest.wes.cl/q"
_TOKEN_APP: Optional[str] = None
_TOKEN_LOCK = threading.Lock()
HORAS_NOCHE: Tuple[int, ...] = (0, 1, 2, 3, 4, 5, 6)
HORA_CIERRE_VENTANA = 7

NODO_MATRIZ = "000027-01"
NODO_ESTANQUE = "000027-02"
NODO_ETAPA5 = "000027-03"
NODO_ETAPAS_14 = "000027-04"
NODOS = (NODO_MATRIZ, NODO_ESTANQUE, NODO_ETAPA5, NODO_ETAPAS_14)

NOMBRES = {
    NODO_MATRIZ: "Matriz ESVAL",
    NODO_ESTANQUE: "Estanque Inferior",
    NODO_ETAPA5: "Etapa N°5",
    NODO_ETAPAS_14: "Etapas 1-4",
}

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

ROOT = Path(__file__).resolve().parent
ESTADO_PATH = ROOT / "reports" / "Fundo_Zapallar" / "alerta_nocturna" / "enviados.json"
SALIDA_DIR = ROOT / "reports" / "Fundo_Zapallar" / "alerta_nocturna"


def ultima_noche_completa(ahora: datetime) -> date:
    local = ahora.astimezone(CHILE_TZ)
    if local.hour >= HORA_CIERRE_VENTANA:
        return local.date()
    return local.date() - timedelta(days=1)


def consumo_ventana(horas: Dict[int, float], ventana: Sequence[int] = HORAS_NOCHE) -> float:
    return float(sum(float(horas.get(h, 0.0)) for h in ventana))


def maximo_hora(horas: Dict[int, float], ventana: Sequence[int] = HORAS_NOCHE) -> float:
    return float(max(float(horas.get(h, 0.0)) for h in ventana))


def racha_flujo(horas: Dict[int, float], ventana: Sequence[int] = HORAS_NOCHE) -> int:
    mejor = 0
    actual = 0
    for hora in ventana:
        if float(horas.get(hora, 0.0)) > 0:
            actual += 1
            mejor = max(mejor, actual)
        else:
            actual = 0
    return mejor


def consumo_dia(horas: Dict[int, float]) -> float:
    return float(sum(float(horas.get(h, 0.0)) for h in range(24)))


def aplicar_persistencia(nivel: str, previo: Optional[str]) -> str:
    """Un seguimiento que se repite al día siguiente pasa a preventiva."""
    if nivel == "seguimiento" and previo in ("seguimiento", "preventiva", "critica"):
        return "preventiva"
    return nivel


def _nivel_banda(
    total: float,
    max_h: float,
    seguimiento_total: float,
    preventiva_total: float,
    critica_total: float,
    seguimiento_hora: Optional[float],
    preventiva_hora: Optional[float],
    critica_hora: Optional[float],
    previo: Optional[str] = None,
) -> str:
    """Los topes de cada banda son inclusivos: >a hasta b queda en esa banda."""
    if total > critica_total or (critica_hora is not None and max_h > critica_hora):
        return "critica"
    if total > preventiva_total or (preventiva_hora is not None and max_h > preventiva_hora):
        return "preventiva"
    en_seguimiento = total > seguimiento_total or (
        seguimiento_hora is not None and max_h > seguimiento_hora
    )
    if en_seguimiento:
        return aplicar_persistencia("seguimiento", previo)
    return "normal"


def nivel_etapa5(total: float, max_h: float, previo: Optional[str] = None) -> str:
    return _nivel_banda(total, max_h, 1, 5, 10, None, None, 4, previo)


def nivel_etapas_14(total: float, max_h: float, previo: Optional[str] = None) -> str:
    return _nivel_banda(total, max_h, 40, 45, 60, 12, 15, 20, previo)


def nivel_matriz(total: float, max_h: float, previo: Optional[str] = None) -> str:
    return _nivel_banda(total, max_h, 55, 65, 80, 25, 30, 35, previo)


def nivel_balance(
    delta: float,
    limite_normal: float,
    limite_seguimiento: float,
    limite_preventiva: float,
    previo_abs: Optional[float] = None,
) -> str:
    """
    |∆| ≤ limite_normal es normal.
    Hasta limite_seguimiento es seguimiento.
    Hasta limite_preventiva, o sobre limite_normal durante 2 días, es preventiva.
    Sobre limite_preventiva es crítica.
    """
    absoluto = abs(float(delta))
    if absoluto > limite_preventiva:
        return "critica"
    if absoluto > limite_seguimiento:
        return "preventiva"
    if absoluto > limite_normal:
        if previo_abs is not None and previo_abs > limite_normal:
            return "preventiva"
        return "seguimiento"
    return "normal"


def _fmt_m3(valor: float) -> str:
    texto = f"{valor:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def _smtp_password() -> str:
    p = (
        os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
        or os.environ.get("WES_SMTP_PASSWORD", "").strip()
        or os.environ.get("SMTP_PASSWORD", "").strip()
    )
    if p:
        return p.replace(" ", "")
    f = ROOT / "gmail_oauth" / "app_password.txt"
    if f.is_file():
        for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                return line.replace(" ", "")
    return "vxbynfpoehbweelj"


def _parse_horas_csv(csv_content: str, dia: date) -> Optional[Dict[int, float]]:
    prefijo = dia.isoformat()
    por_hora: Dict[int, float] = {}
    vio_horario = False
    for line in csv_content.strip().splitlines()[1:]:
        if not line.strip():
            continue
        parts = line.split(",", 1)
        if len(parts) < 2:
            continue
        time_str = parts[0].strip()
        if "T" not in time_str or not time_str.startswith(prefijo):
            continue
        try:
            hora = int(time_str[11:13])
            valor = float(parts[1].strip().replace(" ", "").replace(",", "."))
        except (ValueError, IndexError):
            continue
        if 0 <= hora < 24:
            vio_horario = True
            por_hora[hora] = valor
    if not vio_horario:
        return None
    return por_hora


def _password_app() -> str:
    p = os.environ.get("WES_APP_PASSWORD", "").strip()
    if p:
        return p
    f = ROOT / "gmail_oauth" / "wes_app_password.txt"
    if f.is_file():
        for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                return line
    return ""


def _token_app() -> Optional[str]:
    """Token de la app (dashboard.wes.cl). None si no hay contraseña configurada."""
    global _TOKEN_APP
    with _TOKEN_LOCK:
        if _TOKEN_APP:
            return _TOKEN_APP
        password = _password_app()
        if not password:
            return None
        usuario = os.environ.get("WES_APP_USER", "").strip() or "agente.ia@wes.cl"
        sobre = json.dumps(
            {"m": "POST", "u": APP_LOGIN, "b": {"username": usuario, "password": password}},
            separators=(",", ":"),
        )
        resp = requests.post(
            APP_BFF,
            data=base64.b64encode(sobre.encode()).decode(),
            headers={"Content-Type": "text/plain"},
            timeout=30,
        )
        resp.raise_for_status()
        token = resp.text.strip().strip('"')
        if token.count(".") != 2:
            raise requests.RequestException("La app no devolvió un token de sesión")
        _TOKEN_APP = token
        return _TOKEN_APP


def _descargar_dia_app(node_id: str, dia: date) -> Optional[Dict[int, float]]:
    token = _token_app()
    if not token:
        return None
    fecha = dia.strftime("%d%m%Y")
    url = f"{APP_API}/nodes/{node_id}/dates.measures.csv?start={fecha}&end={fecha}"
    sobre = json.dumps({"m": "GET", "u": url, "b": None}, separators=(",", ":"))
    resp = requests.post(
        APP_BFF,
        data=base64.b64encode(sobre.encode()).decode(),
        headers={"Content-Type": "text/plain", "authorization": token},
        timeout=60,
    )
    if resp.status_code in (401, 403):
        global _TOKEN_APP
        with _TOKEN_LOCK:
            _TOKEN_APP = None
        resp.raise_for_status()
    resp.raise_for_status()
    return _parse_horas_csv(resp.text, dia)


def _descargar_dia(node_id: str, dia: date, intentos: int = 3) -> Optional[Dict[int, float]]:
    fecha = dia.strftime("%d%m%Y")
    ultimo_error = ""
    if _password_app():
        for intento in range(intentos):
            try:
                return _descargar_dia_app(node_id, dia)
            except requests.RequestException as exc:
                ultimo_error = str(exc)
                time.sleep(0.4 * (intento + 1))
    url = f"{BASE_URL}/nodes/{node_id}/dates.measures.csv"
    for intento in range(intentos):
        try:
            resp = requests.get(url, params=[("start", fecha), ("end", fecha)], timeout=20)
            resp.raise_for_status()
            return _parse_horas_csv(resp.text, dia)
        except requests.RequestException as exc:
            ultimo_error = str(exc)
            time.sleep(0.4 * (intento + 1))
    print(f"[WARN] Sin datos {node_id} {dia.isoformat()}: {ultimo_error}", file=sys.stderr)
    return None


def _cargar_estado() -> dict:
    if not ESTADO_PATH.is_file():
        return {"enviados": [], "evaluaciones": {}}
    try:
        data = json.loads(ESTADO_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"enviados": [], "evaluaciones": {}}
    enviados = data.get("enviados")
    evaluaciones = data.get("evaluaciones")
    return {
        "enviados": [str(x) for x in enviados] if isinstance(enviados, list) else [],
        "evaluaciones": evaluaciones if isinstance(evaluaciones, dict) else {},
    }


def _guardar_estado(estado: dict) -> None:
    ESTADO_PATH.parent.mkdir(parents=True, exist_ok=True)
    ESTADO_PATH.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")


def _clave_envio(dia: date, regla: str) -> str:
    return f"{dia.isoformat()}|{regla}"


def _previo_nivel(estado: dict, dia: date, regla: str) -> Optional[str]:
    bloque = (estado.get("evaluaciones") or {}).get((dia - timedelta(days=1)).isoformat()) or {}
    item = bloque.get(regla) or {}
    nivel = item.get("nivel")
    return str(nivel) if nivel else None


def _previo_abs(estado: dict, dia: date, regla: str) -> Optional[float]:
    bloque = (estado.get("evaluaciones") or {}).get((dia - timedelta(days=1)).isoformat()) or {}
    item = bloque.get(regla) or {}
    if "valor" not in item:
        return None
    try:
        return abs(float(item["valor"]))
    except (TypeError, ValueError):
        return None


def _registrar(estado: dict, dia: date, regla: str, nivel: str, valor: float) -> None:
    evaluaciones = estado.setdefault("evaluaciones", {})
    bloque = evaluaciones.setdefault(dia.isoformat(), {})
    bloque[regla] = {"nivel": nivel, "valor": round(float(valor), 4)}


def _alerta(
    dia: date,
    periodo: str,
    regla: str,
    punto: str,
    nivel: str,
    valor: float,
    umbral: str,
    detalle: str,
    accion: str,
) -> dict:
    return {
        "fecha": dia.isoformat(),
        "periodo": periodo,
        "regla": regla,
        "punto": punto,
        "nivel": nivel,
        "valor": valor,
        "umbral": umbral,
        "detalle": detalle,
        "accion": accion,
        "notificar": nivel in ("seguimiento", "preventiva", "critica"),
    }


def evaluar_noche(horas_por_nodo: Dict[str, Optional[Dict[int, float]]], noche: date, estado: dict) -> List[dict]:
    filas: List[dict] = []
    periodo = f"{noche.strftime('%d-%m-%Y')} 00:00 a 06:00"

    def horas_de(node_id: str) -> Optional[Dict[int, float]]:
        return horas_por_nodo.get(node_id)

    especificaciones = (
        (
            "etapa5-noche",
            NODO_ETAPA5,
            "Etapa N°5 — noche",
            lambda total, max_h, _racha, previo: nivel_etapa5(total, max_h, previo),
            "Seguimiento: >1 a 5 m³/noche. Preventiva: >5 a 10 m³/noche. Crítica: >10 m³/noche o >4 m³/h.",
            "Revisar riego abierto o pérdida.",
        ),
        (
            "etapas14-noche",
            NODO_ETAPAS_14,
            "Etapas 1-4 — noche",
            lambda total, max_h, _racha, previo: nivel_etapas_14(total, max_h, previo),
            "Seguimiento: >40 a 45 m³/noche o >12 a 15 m³/h. Preventiva: >45 a 60 m³/noche o >15 a 20 m³/h. Crítica: >60 m³/noche o >20 m³/h.",
            "Contrastar con el programa de riego.",
        ),
        (
            "matriz-noche",
            NODO_MATRIZ,
            "Matriz ESVAL — noche",
            lambda total, max_h, _racha, previo: nivel_matriz(total, max_h, previo),
            "Seguimiento: >55 a 65 m³/noche o >25 a 30 m³/h. Preventiva: >65 a 80 m³/noche o >30 a 35 m³/h. Crítica: >80 m³/noche o >35 m³/h.",
            "Revisar carga de estanques y consumos simultáneos.",
        ),
    )
    for regla, node_id, punto, clasificar, umbral, accion in especificaciones:
        horas = horas_de(node_id)
        if horas is None:
            filas.append(
                _alerta(noche, periodo, regla, punto, "sin_dato", 0.0, umbral, "Sin lectura de la noche.", accion)
            )
            continue
        total = consumo_ventana(horas)
        max_h = maximo_hora(horas)
        racha = racha_flujo(horas)
        previo = _previo_nivel(estado, noche, regla)
        nivel = clasificar(total, max_h, racha, previo)
        detalle = (
            f"Total {_fmt_m3(total)} m³; máximo {_fmt_m3(max_h)} m³/h; "
            f"racha con flujo {racha} h. "
            + "; ".join(f"{h:02d}:00={_fmt_m3(float(horas.get(h, 0.0)))}" for h in HORAS_NOCHE)
        )
        if nivel == "preventiva" and clasificar(total, max_h, racha, None) == "seguimiento":
            detalle += ". La alerta de seguimiento se repite, por eso el nivel pasa a preventiva."
        _registrar(estado, noche, regla, nivel, total)
        filas.append(_alerta(noche, periodo, regla, punto, nivel, total, umbral, detalle, accion))
    return filas


def evaluar_balances(
    horas_por_nodo_dia: Dict[str, Optional[Dict[int, float]]],
    dia: date,
    estado: dict,
) -> List[dict]:
    filas: List[dict] = []
    periodo = f"{dia.strftime('%d-%m-%Y')} día completo"

    def total(node_id: str) -> Optional[float]:
        horas = horas_por_nodo_dia.get(node_id)
        if horas is None:
            return None
        return consumo_dia(horas)

    matriz = total(NODO_MATRIZ)
    estanque = total(NODO_ESTANQUE)
    etapa5 = total(NODO_ETAPA5)
    etapas14 = total(NODO_ETAPAS_14)

    if matriz is None or estanque is None:
        filas.append(
            _alerta(
                dia,
                periodo,
                "balance-matriz-estanque",
                "Matriz ESVAL − Estanque Inferior",
                "sin_dato",
                0.0,
                "Seguimiento >21 a 25 m³/día. Preventiva >25 a 42 m³/día o >21 m³ por 2 días. Crítica >42 m³/día.",
                "Sin lectura diaria de Matriz o Estanque Inferior.",
                "Revisar el saldo no explicado por el estanque (capacidad 21 m³).",
            )
        )
    else:
        delta = matriz - estanque
        regla = "balance-matriz-estanque"
        nivel = nivel_balance(delta, 21, 25, 42, _previo_abs(estado, dia, regla))
        _registrar(estado, dia, regla, nivel, delta)
        detalle_balance = f"Matriz {_fmt_m3(matriz)} m³; Estanque Inferior {_fmt_m3(estanque)} m³; ∆ {_fmt_m3(delta)} m³."
        if nivel == "preventiva" and nivel_balance(delta, 21, 25, 42) == "seguimiento":
            detalle_balance += " El saldo supera 21 m³ por segundo día, por eso pasa a preventiva."
        filas.append(
            _alerta(
                dia,
                periodo,
                regla,
                "Matriz ESVAL − Estanque Inferior (Estanque del Medio)",
                nivel,
                delta,
                "Seguimiento >21 a 25 m³/día. Preventiva >25 a 42 m³/día o >21 m³ por 2 días. Crítica >42 m³/día.",
                detalle_balance,
                "Revisar llenado o descarga del estanque. Sobre 21 m³ hay saldo no explicado por almacenamiento.",
            )
        )

    if estanque is None or etapa5 is None or etapas14 is None:
        filas.append(
            _alerta(
                dia,
                periodo,
                "balance-estanque-etapas",
                "Estanque Inferior − etapas",
                "sin_dato",
                0.0,
                "Seguimiento >30 a 36 m³/día. Preventiva >36 a 60 m³/día o >30 m³ por 2 días. Crítica >60 m³/día.",
                "Sin lectura diaria del estanque o de las etapas.",
                "Revisar el saldo frente a los estanques de etapa (15 m³ cada uno, 30 m³ en conjunto).",
            )
        )
    else:
        salida_etapas = etapa5 + etapas14
        delta = estanque - salida_etapas
        regla = "balance-estanque-etapas"
        nivel = nivel_balance(delta, 30, 36, 60, _previo_abs(estado, dia, regla))
        _registrar(estado, dia, regla, nivel, delta)
        detalle_balance = (
            f"Estanque {_fmt_m3(estanque)} m³; Etapa N°5 {_fmt_m3(etapa5)} m³; "
            f"Etapas 1-4 {_fmt_m3(etapas14)} m³; ∆ {_fmt_m3(delta)} m³."
        )
        if nivel == "preventiva" and nivel_balance(delta, 30, 36, 60) == "seguimiento":
            detalle_balance += " El saldo supera 30 m³ por segundo día, por eso pasa a preventiva."
        filas.append(
            _alerta(
                dia,
                periodo,
                regla,
                "Estanque Inferior − (Etapa N°5 + Etapas 1-4)",
                nivel,
                delta,
                "Seguimiento >30 a 36 m³/día. Preventiva >36 a 60 m³/día o >30 m³ por 2 días. Crítica >60 m³/día.",
                detalle_balance,
                "Sobre 30 m³ debe revisarse el balance de los dos estanques de etapa.",
            )
        )
    return filas


def _descargar_mapa(fechas: Sequence[date], workers: int = 8) -> Dict[Tuple[str, date], Optional[Dict[int, float]]]:
    tareas = [(node_id, dia) for dia in fechas for node_id in NODOS]
    serie: Dict[Tuple[str, date], Optional[Dict[int, float]]] = {}
    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futuros = {pool.submit(_descargar_dia, node_id, dia): (node_id, dia) for node_id, dia in tareas}
        for futuro in as_completed(futuros):
            node_id, dia = futuros[futuro]
            serie[(node_id, dia)] = futuro.result()
    return serie


def _por_nodo(serie: Dict[Tuple[str, date], Optional[Dict[int, float]]], dia: date) -> Dict[str, Optional[Dict[int, float]]]:
    return {node_id: serie.get((node_id, dia)) for node_id in NODOS}


def _guardar_csv(filas: Sequence[dict], noche: date) -> Path:
    SALIDA_DIR.mkdir(parents=True, exist_ok=True)
    path = SALIDA_DIR / f"alerta_fundo_zapallar_{noche.strftime('%Y%m%d')}.csv"
    campos = ["fecha", "periodo", "regla", "punto", "nivel", "valor", "umbral", "detalle", "accion", "notificar"]
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=campos)
        writer.writeheader()
        for fila in filas:
            writer.writerow({k: fila.get(k, "") for k in campos})
    return path


def _cuerpo_correo(filas: Sequence[dict]) -> str:
    lineas = [
        "Alerta Fundo Zapallar",
        "",
        "Tres niveles: seguimiento, preventiva y crítica. Ventana 00:00 a 06:00 y balance diario.",
        "La crítica se reserva para consumos excepcionalmente altos.",
        "Si un seguimiento se repite al día siguiente, pasa a preventiva.",
        "Seguimiento o preventiva: contrastar con riegos, llenados u operaciones extraordinarias antes de escalar al fundo.",
        "Crítica: aviso prioritario.",
        "",
    ]
    for fila in filas:
        lineas.extend(
            [
                f"- {fila['nivel'].upper()}: {fila['punto']}",
                f"  Regla: {fila['regla']}",
                f"  Período: {fila['periodo']}",
                f"  Valor observado: {_fmt_m3(float(fila['valor']))} m³",
                f"  Umbral: {fila['umbral']}",
                f"  Detalle: {fila['detalle']}",
                f"  Acción: {fila['accion']}",
                "",
            ]
        )
    return "\n".join(lineas).rstrip() + "\n"


def enviar_correo(filas: Sequence[dict]) -> None:
    if not filas:
        return
    destinatarios = [obtener_email("anibal"), obtener_email("juan")]
    nombres = f"{obtener_nombre('anibal')} y {obtener_nombre('juan')}"
    if any(f["nivel"] == "critica" for f in filas):
        nivel = "crítica"
    elif any(f["nivel"] == "preventiva" for f in filas):
        nivel = "preventiva"
    else:
        nivel = "de seguimiento"
    fechas = sorted({f["fecha"] for f in filas})
    fecha_txt = fechas[0] if len(fechas) == 1 else f"{fechas[0]} a {fechas[-1]}"
    puntos = ", ".join(f["punto"] for f in filas)
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = f"Alerta {nivel} Fundo Zapallar {fecha_txt}"
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="wes.cl")
    msg.attach(MIMEText(f"Hola {nombres},\n\n{_cuerpo_correo(filas)}", "plain", "utf-8"))
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO, timeout=60) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(SMTP_USUARIO, _smtp_password())
        server.send_message(msg, to_addrs=destinatarios)
    print(f"[OK] Correo enviado a {', '.join(destinatarios)}: {puntos}")


def _imprimir(filas: Sequence[dict]) -> None:
    print("")
    print(f"{'Regla':<28} {'Nivel':<12} {'Valor':>10}")
    for fila in filas:
        print(f"{fila['regla']:<28} {fila['nivel']:<12} {fila['valor']:>10.2f}")


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Alertas Fundo Zapallar del informe de seguimiento.")
    parser.add_argument("--fecha", help="Noche a evaluar, YYYY-MM-DD. Por defecto, la última cerrada.")
    parser.add_argument("--dry-run", action="store_true", help="Evalúa e imprime, sin enviar correo.")
    parser.add_argument("--forzar-reenvio", action="store_true")
    parser.add_argument("--workers", type=int, default=8)
    args = parser.parse_args(argv)

    noche = date.fromisoformat(args.fecha) if args.fecha else ultima_noche_completa(datetime.now(CHILE_TZ))
    dia_balance = noche - timedelta(days=1)
    fechas = [dia_balance - timedelta(days=1), dia_balance, noche]
    print(
        f"[INFO] Noche {noche.isoformat()} 00:00-06:00; balance diario {dia_balance.isoformat()}"
    )
    serie = _descargar_mapa(fechas, workers=args.workers)
    estado = _cargar_estado()

    # El día anterior al balance solo alimenta la persistencia de 2 días.
    evaluar_balances(_por_nodo(serie, dia_balance - timedelta(days=1)), dia_balance - timedelta(days=1), estado)
    filas_noche_previa = evaluar_noche(_por_nodo(serie, dia_balance), dia_balance, estado)
    filas = evaluar_noche(_por_nodo(serie, noche), noche, estado)
    filas.extend(evaluar_balances(_por_nodo(serie, dia_balance), dia_balance, estado))
    # La noche previa queda registrada para la persistencia, pero no se reavisa.
    _ = filas_noche_previa

    csv_path = _guardar_csv(filas, noche)
    _imprimir(filas)
    print(f"[INFO] Detalle: {csv_path}")

    avisos = [f for f in filas if f["notificar"]]
    if not avisos:
        print("[INFO] Ninguna regla de seguimiento, preventiva o crítica. No se envía correo.")
        _guardar_estado(estado)
        return 0

    pendientes = []
    for fila in avisos:
        clave = _clave_envio(date.fromisoformat(fila["fecha"]), fila["regla"])
        if args.forzar_reenvio or clave not in estado["enviados"]:
            pendientes.append(fila)
        else:
            print(f"[INFO] Ya se avisó {fila['regla']} para {fila['fecha']}.")

    if not pendientes:
        print("[INFO] Esas alertas ya fueron enviadas.")
        _guardar_estado(estado)
        return 0
    if args.dry_run:
        print(f"[DRY-RUN] Se enviaría correo por {len(pendientes)} regla(s).")
        _guardar_estado(estado)
        return 0

    enviar_correo(pendientes)
    for fila in pendientes:
        estado["enviados"].append(_clave_envio(date.fromisoformat(fila["fecha"]), fila["regla"]))
    _guardar_estado(estado)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
