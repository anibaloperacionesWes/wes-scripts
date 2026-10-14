"""
Alertas de Fundo Zapallar según el informe de seguimiento (6 de octubre de 2026).

No usa el promedio de noches. La ventana nocturna es 00:00 a 06:00 (horas 00 a 06).
Una condición crítica prevalece sobre la preventiva. Si una preventiva se repite al
día siguiente, pasa a crítica.

Reglas nocturnas:
- Etapa N°5: preventiva >1 y ≤5 m³/noche o flujo >0 por 2 horas; crítica >5 m³/noche,
  >2 m³/h o flujo >0 por 4 horas.
- Etapas 1-4: preventiva >40 y ≤45 m³/noche o >12 y ≤15 m³/h; crítica >45 m³/noche
  o >15 m³/h.
- Matriz ESVAL: preventiva >55 y ≤65 m³/noche o >25 y ≤30 m³/h; crítica >65 m³/noche
  o >30 m³/h.

Balances diarios (valor absoluto):
- Matriz ESVAL − Estanque Inferior (Estanque del Medio del informe): preventiva
  >21 y ≤25 m³; crítica >25 m³ o >21 m³ durante 2 días.
- Estanque Inferior − (Etapa N°5 + Etapas 1-4): preventiva >30 y ≤36 m³; crítica
  >36 m³ o >30 m³ durante 2 días.

El aviso a Aníbal y Juan indica el punto, la regla, el valor, el umbral y el período.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import smtplib
import sys
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
    """Una preventiva que se repite al día siguiente pasa a crítica."""
    if nivel == "preventiva" and previo in ("preventiva", "critica"):
        return "critica"
    return nivel


def nivel_etapa5(total: float, max_h: float, racha: int, previo: Optional[str] = None) -> str:
    if total > 5 or max_h > 2 or racha >= 4:
        return "critica"
    preventiva = (total > 1 and total <= 5) or racha >= 2
    if preventiva:
        return aplicar_persistencia("preventiva", previo)
    return "normal"


def nivel_etapas_14(total: float, max_h: float, previo: Optional[str] = None) -> str:
    if total > 45 or max_h > 15:
        return "critica"
    preventiva = (total > 40 and total <= 45) or (max_h > 12 and max_h <= 15)
    if preventiva:
        return aplicar_persistencia("preventiva", previo)
    return "normal"


def nivel_matriz(total: float, max_h: float, previo: Optional[str] = None) -> str:
    if total > 65 or max_h > 30:
        return "critica"
    preventiva = (total > 55 and total <= 65) or (max_h > 25 and max_h <= 30)
    if preventiva:
        return aplicar_persistencia("preventiva", previo)
    return "normal"


def nivel_balance(
    delta: float,
    limite_normal: float,
    limite_preventivo: float,
    previo_abs: Optional[float] = None,
) -> str:
    """
    |∆| ≤ limite_normal es normal.
    Hasta limite_preventivo es preventiva.
    Sobre limite_preventivo, o sobre limite_normal durante 2 días, es crítica.
    """
    absoluto = abs(float(delta))
    if absoluto > limite_preventivo or (
        previo_abs is not None and absoluto > limite_normal and previo_abs > limite_normal
    ):
        return "critica"
    if absoluto > limite_normal:
        return "preventiva"
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


def _descargar_dia(node_id: str, dia: date, intentos: int = 3) -> Optional[Dict[int, float]]:
    url = f"{BASE_URL}/nodes/{node_id}/dates.measures.csv"
    fecha = dia.strftime("%d%m%Y")
    ultimo_error = ""
    for intento in range(intentos):
        try:
            resp = requests.get(url, params=[("start", fecha), ("end", fecha)], timeout=45)
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
        "notificar": nivel in ("preventiva", "critica"),
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
            lambda total, max_h, racha, previo: nivel_etapa5(total, max_h, racha, previo),
            "Preventiva: >1 y ≤5 m³/noche, o flujo >0 durante 2 horas. Crítica: >5 m³/noche, >2 m³/h o flujo >0 durante 4 horas.",
            "Revisar riego abierto o pérdida.",
        ),
        (
            "etapas14-noche",
            NODO_ETAPAS_14,
            "Etapas 1-4 — noche",
            lambda total, max_h, _racha, previo: nivel_etapas_14(total, max_h, previo),
            "Preventiva: >40 y ≤45 m³/noche, o >12 y ≤15 m³/h. Crítica: >45 m³/noche o >15 m³/h.",
            "Contrastar con el programa de riego.",
        ),
        (
            "matriz-noche",
            NODO_MATRIZ,
            "Matriz ESVAL — noche",
            lambda total, max_h, _racha, previo: nivel_matriz(total, max_h, previo),
            "Preventiva: >55 y ≤65 m³/noche, o >25 y ≤30 m³/h. Crítica: >65 m³/noche o >30 m³/h.",
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
        if previo and nivel == "critica" and clasificar(total, max_h, racha, None) == "preventiva":
            detalle += ". La preventiva del día anterior se mantiene, por eso el nivel pasa a crítico."
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
                "Preventiva >21 y ≤25 m³/día. Crítica >25 m³/día o >21 m³ por 2 días.",
                "Sin lectura diaria de Matriz o Estanque Inferior.",
                "Revisar el saldo no explicado por el estanque (capacidad 21 m³).",
            )
        )
    else:
        delta = matriz - estanque
        regla = "balance-matriz-estanque"
        nivel = nivel_balance(delta, 21, 25, _previo_abs(estado, dia, regla))
        _registrar(estado, dia, regla, nivel, delta)
        filas.append(
            _alerta(
                dia,
                periodo,
                regla,
                "Matriz ESVAL − Estanque Inferior (Estanque del Medio)",
                nivel,
                delta,
                "Normal |∆| ≤21 m³/día. Preventiva >21 y ≤25. Crítica >25, o >21 durante 2 días.",
                f"Matriz {_fmt_m3(matriz)} m³; Estanque Inferior {_fmt_m3(estanque)} m³; ∆ {_fmt_m3(delta)} m³.",
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
                "Preventiva >30 y ≤36 m³/día. Crítica >36 m³/día o >30 m³ por 2 días.",
                "Sin lectura diaria del estanque o de las etapas.",
                "Revisar el saldo frente a los estanques de etapa (15 m³ cada uno, 30 m³ en conjunto).",
            )
        )
    else:
        salida_etapas = etapa5 + etapas14
        delta = estanque - salida_etapas
        regla = "balance-estanque-etapas"
        nivel = nivel_balance(delta, 30, 36, _previo_abs(estado, dia, regla))
        _registrar(estado, dia, regla, nivel, delta)
        filas.append(
            _alerta(
                dia,
                periodo,
                regla,
                "Estanque Inferior − (Etapa N°5 + Etapas 1-4)",
                nivel,
                delta,
                "Normal |∆| ≤30 m³/día. Preventiva >30 y ≤36. Crítica >36, o >30 durante 2 días.",
                (
                    f"Estanque {_fmt_m3(estanque)} m³; Etapa N°5 {_fmt_m3(etapa5)} m³; "
                    f"Etapas 1-4 {_fmt_m3(etapas14)} m³; ∆ {_fmt_m3(delta)} m³."
                ),
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
        "Reglas del informe de seguimiento: umbrales fijos de noche (00:00 a 06:00) y balance diario.",
        "Una condición crítica prevalece sobre la preventiva.",
        "Antes de escalar al fundo, contrastar el evento con riegos, llenados u operaciones extraordinarias.",
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
    hay_critica = any(f["nivel"] == "critica" for f in filas)
    nivel = "crítica" if hay_critica else "preventiva"
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
        print("[INFO] Ninguna regla preventiva o crítica. No se envía correo.")
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
