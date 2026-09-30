"""
Alerta nocturna de Fundo Zapallar.

Cada noche completa (00:00 a 06:00, horas 00–06 inclusive) se compara con el
promedio de las 90 noches anteriores, solo en los puntos de Fundo Zapallar.

Si el consumo de la noche supera ese promedio en un 25%
(umbral = promedio × 1,25), se envía un correo a Aníbal y Juan indicando
qué punto se pasó.

La noche evaluada no entra en el promedio. La ventana se considera cerrada
a las 07:00 hora Chile, cuando termina la hora 06.

Uso:
  python alerta_nocturna_fundo_zapallar.py
  python alerta_nocturna_fundo_zapallar.py --dry-run
  python alerta_nocturna_fundo_zapallar.py --fecha 2026-09-30
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
from exclusiones_reportes import FUNDO_ZAPALLAR_NODE_IDS

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

CHILE_TZ = ZoneInfo("America/Santiago")
BASE_URL = "http://104.248.53.141:7003/wes/api/acl-node/v1"

# Horas 00, 01, 02, 03, 04, 05 y 06: consumo entre las 00:00 y las 07:00.
HORAS_VENTANA: Tuple[int, ...] = (0, 1, 2, 3, 4, 5, 6)
NOCHES_BASE = 90
# La noche alerta cuando supera al promedio en un 25%.
FACTOR_UMBRAL = 1.25
HORA_CIERRE_VENTANA = 7

NOMBRES_PUNTOS = {
    "000027-01": "Matriz ESVAL",
    "000027-02": "Estanque Inferior",
    "000027-03": "Etapa N°5",
    "000027-04": "Etapa N°1 al 4",
    "000027-06": "Etapa N°1",
    "000027-07": "Etapa N°2",
    "000027-08": "Etapa N°3",
    "000027-09": "Riego Llenado de Estanque ESVAL",
}

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

ROOT = Path(__file__).resolve().parent
ESTADO_PATH = ROOT / "reports" / "Fundo_Zapallar" / "alerta_nocturna" / "enviados.json"
SALIDA_DIR = ROOT / "reports" / "Fundo_Zapallar" / "alerta_nocturna"


def ultima_noche_completa(ahora: datetime) -> date:
    """Día civil Chile cuya ventana 00:00–06:00 ya terminó."""
    local = ahora.astimezone(CHILE_TZ)
    if local.hour >= HORA_CIERRE_VENTANA:
        return local.date()
    return local.date() - timedelta(days=1)


def noches_base(noche: date, cantidad: int = NOCHES_BASE) -> List[date]:
    """Las ``cantidad`` noches anteriores a ``noche``, de la más antigua a la más nueva."""
    if cantidad < 1:
        raise ValueError("cantidad debe ser >= 1")
    inicio = noche - timedelta(days=cantidad)
    return [inicio + timedelta(days=i) for i in range(cantidad)]


def consumo_ventana(horas: Dict[int, float]) -> float:
    return float(sum(float(horas.get(h, 0.0)) for h in HORAS_VENTANA))


def promedio_noches(consumos: Sequence[float]) -> Optional[float]:
    if not consumos:
        return None
    return float(sum(consumos)) / float(len(consumos))


def supera_umbral(consumo: float, promedio: float, factor: float = FACTOR_UMBRAL) -> bool:
    """True si la noche supera al promedio en el porcentaje del factor (1,25 = +25%)."""
    if consumo < 0 or promedio < 0 or factor < 0:
        raise ValueError("consumo, promedio y factor no pueden ser negativos")
    return consumo > (promedio * factor)


def exceso_porcentaje(consumo: float, promedio: float) -> Optional[float]:
    if promedio <= 0:
        return None
    return (consumo / promedio - 1.0) * 100.0


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
    # Mismo fallback que el resto de los envíos automáticos del repo.
    return "vxbynfpoehbweelj"


def _fmt_m3(valor: float) -> str:
    texto = f"{valor:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt_pct(valor: Optional[float]) -> str:
    if valor is None:
        return "s/promedio"
    texto = f"{valor:,.1f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".") + " %"


def _parse_horas_csv(csv_content: str, dia: date) -> Optional[Dict[int, float]]:
    """
    Horas 0–23 del día civil.

    El CSV marca la hora como ``2026-09-30T03:00:00.000Z`` y la app la grafica
    como las 03:00 de ese día, sin convertir UTC→Chile. Si la respuesta trae
    totales diarios (sin hora), no hay serie nocturna.
    """
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
            resp = requests.get(
                url,
                params=[("start", fecha), ("end", fecha)],
                timeout=45,
            )
            resp.raise_for_status()
            return _parse_horas_csv(resp.text, dia)
        except requests.RequestException as exc:
            ultimo_error = str(exc)
            time.sleep(0.4 * (intento + 1))
    print(f"[WARN] Sin datos {node_id} {dia.isoformat()}: {ultimo_error}", file=sys.stderr)
    return None


def _cargar_estado() -> Dict[str, List[str]]:
    if not ESTADO_PATH.is_file():
        return {"enviados": []}
    try:
        data = json.loads(ESTADO_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"enviados": []}
    enviados = data.get("enviados")
    if not isinstance(enviados, list):
        return {"enviados": []}
    return {"enviados": [str(x) for x in enviados]}


def _guardar_estado(estado: Dict[str, List[str]]) -> None:
    ESTADO_PATH.parent.mkdir(parents=True, exist_ok=True)
    ESTADO_PATH.write_text(
        json.dumps(estado, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _clave_envio(noche: date, node_id: str) -> str:
    return f"{noche.isoformat()}|{node_id}"


def evaluar_puntos(
    noche: date,
    node_ids: Sequence[str],
    noches: int = NOCHES_BASE,
    factor: float = FACTOR_UMBRAL,
    workers: int = 12,
) -> List[dict]:
    """Descarga la ventana nocturna y arma el resultado por punto."""
    base = noches_base(noche, noches)
    dias = list(base) + [noche]
    tareas = [(node_id, dia) for node_id in node_ids for dia in dias]
    serie: Dict[Tuple[str, date], Optional[Dict[int, float]]] = {}

    print(
        f"[INFO] Fundo Zapallar: {len(node_ids)} puntos, noche {noche.isoformat()}, "
        f"base {base[0].isoformat()} a {base[-1].isoformat()} ({len(tareas)} consultas)"
    )
    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futuros = {
            pool.submit(_descargar_dia, node_id, dia): (node_id, dia) for node_id, dia in tareas
        }
        listos = 0
        for futuro in as_completed(futuros):
            node_id, dia = futuros[futuro]
            serie[(node_id, dia)] = futuro.result()
            listos += 1
            if listos % 100 == 0 or listos == len(tareas):
                print(f"[INFO] Consultas listas: {listos}/{len(tareas)}")

    filas: List[dict] = []
    for node_id in node_ids:
        consumos_base: List[float] = []
        noches_sin_dato = 0
        for dia in base:
            horas = serie.get((node_id, dia))
            if horas is None:
                noches_sin_dato += 1
                continue
            consumos_base.append(consumo_ventana(horas))

        horas_noche = serie.get((node_id, noche))
        promedio = promedio_noches(consumos_base)
        if horas_noche is None or promedio is None:
            filas.append(
                {
                    "node_id": node_id,
                    "nombre": NOMBRES_PUNTOS.get(node_id, node_id),
                    "noche": noche.isoformat(),
                    "consumo_m3": None,
                    "promedio_m3": promedio,
                    "umbral_m3": None if promedio is None else promedio * factor,
                    "exceso_pct": None,
                    "alerta": False,
                    "noches_con_dato": len(consumos_base),
                    "noches_sin_dato": noches_sin_dato,
                    "detalle_horario": "",
                    "motivo": "sin dato de la noche" if horas_noche is None else "sin base",
                }
            )
            continue

        consumo = consumo_ventana(horas_noche)
        umbral = promedio * factor
        detalle = "; ".join(
            f"{h:02d}:00={_fmt_m3(float(horas_noche.get(h, 0.0)))}" for h in HORAS_VENTANA
        )
        filas.append(
            {
                "node_id": node_id,
                "nombre": NOMBRES_PUNTOS.get(node_id, node_id),
                "noche": noche.isoformat(),
                "consumo_m3": consumo,
                "promedio_m3": promedio,
                "umbral_m3": umbral,
                "exceso_pct": exceso_porcentaje(consumo, promedio),
                "alerta": supera_umbral(consumo, promedio, factor),
                "noches_con_dato": len(consumos_base),
                "noches_sin_dato": noches_sin_dato,
                "detalle_horario": detalle,
                "motivo": "",
            }
        )
    return filas


def _guardar_csv(filas: Sequence[dict], noche: date) -> Path:
    SALIDA_DIR.mkdir(parents=True, exist_ok=True)
    path = SALIDA_DIR / f"alerta_nocturna_fundo_zapallar_{noche.strftime('%Y%m%d')}.csv"
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=[
                "noche",
                "node_id",
                "nombre",
                "consumo_m3",
                "promedio_m3",
                "umbral_m3",
                "exceso_pct",
                "alerta",
                "noches_con_dato",
                "noches_sin_dato",
                "detalle_horario",
                "motivo",
            ],
        )
        writer.writeheader()
        for fila in filas:
            writer.writerow(fila)
    return path


def _cuerpo_correo(filas_alerta: Sequence[dict], noche: date, factor: float) -> str:
    pct = int(round((factor - 1.0) * 100))
    lineas = [
        "Alerta nocturna — Fundo Zapallar",
        "",
        f"Noche evaluada: {noche.strftime('%d-%m-%Y')} (00:00 a 06:00, hora Chile).",
        f"Promedio: últimas {NOCHES_BASE} noches anteriores, misma ventana.",
        f"Criterio: se avisa si el consumo de la noche supera ese promedio en un {pct} % "
        f"(umbral = promedio × {factor:.2f}).",
        "",
        "Puntos sobre el umbral:",
    ]
    for fila in filas_alerta:
        exceso = _fmt_pct(fila["exceso_pct"])
        lineas.extend(
            [
                "",
                f"- {fila['nombre']} ({fila['node_id']})",
                f"  Consumo de la noche: {_fmt_m3(fila['consumo_m3'])} m³",
                f"  Promedio {fila['noches_con_dato']} noches: {_fmt_m3(fila['promedio_m3'])} m³",
                f"  Umbral: {_fmt_m3(fila['umbral_m3'])} m³",
                f"  Exceso sobre el promedio: {exceso}",
                f"  Horario: {fila['detalle_horario']}",
            ]
        )
    lineas.extend(
        [
            "",
            "Solo se incluyen puntos de Fundo Zapallar.",
            "Este aviso lo genera el control nocturno automático de WES.",
        ]
    )
    return "\n".join(lineas)


def enviar_correo(filas_alerta: Sequence[dict], noche: date, factor: float) -> None:
    if not filas_alerta:
        return
    destinatarios = [obtener_email("anibal"), obtener_email("juan")]
    nombres = f"{obtener_nombre('anibal')} y {obtener_nombre('juan')}"
    puntos = ", ".join(f"{f['nombre']} ({f['node_id']})" for f in filas_alerta)
    if len(filas_alerta) == 1:
        asunto = (
            f"Alerta nocturna Fundo Zapallar {noche.strftime('%d-%m-%Y')} — "
            f"{filas_alerta[0]['nombre']}"
        )
    else:
        asunto = (
            f"Alerta nocturna Fundo Zapallar {noche.strftime('%d-%m-%Y')} — "
            f"{len(filas_alerta)} puntos"
        )

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = asunto
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="wes.cl")
    cuerpo = _cuerpo_correo(filas_alerta, noche, factor)
    cuerpo = f"Hola {nombres},\n\n{cuerpo}\n"
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO, timeout=60) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(SMTP_USUARIO, _smtp_password())
        server.send_message(msg, to_addrs=destinatarios)
    print(f"[OK] Correo enviado a {', '.join(destinatarios)}: {puntos}")


def _imprimir_resumen(filas: Sequence[dict]) -> None:
    print("")
    print(f"{'Punto':<36} {'Noche m3':>10} {'Promedio':>10} {'Umbral':>10} {'Alerta':>8}")
    for fila in filas:
        consumo = "s/d" if fila["consumo_m3"] is None else f"{fila['consumo_m3']:.2f}"
        promedio = "s/d" if fila["promedio_m3"] is None else f"{fila['promedio_m3']:.2f}"
        umbral = "s/d" if fila["umbral_m3"] is None else f"{fila['umbral_m3']:.2f}"
        print(
            f"{fila['nombre'][:36]:<36} {consumo:>10} {promedio:>10} {umbral:>10} "
            f"{'SI' if fila['alerta'] else 'no':>8}"
        )


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Alerta nocturna Fundo Zapallar (00:00–06:00).")
    parser.add_argument("--fecha", help="Noche a evaluar, YYYY-MM-DD. Por defecto, la última cerrada.")
    parser.add_argument("--noches", type=int, default=NOCHES_BASE, help="Noches del promedio (default 90).")
    parser.add_argument("--factor", type=float, default=FACTOR_UMBRAL, help="Umbral = promedio × factor.")
    parser.add_argument("--dry-run", action="store_true", help="Evalúa e imprime, sin enviar correo.")
    parser.add_argument("--forzar-reenvio", action="store_true", help="Vuelve a enviar aunque ya se avisó esa noche.")
    parser.add_argument("--workers", type=int, default=12)
    args = parser.parse_args(argv)

    if args.fecha:
        noche = date.fromisoformat(args.fecha)
    else:
        noche = ultima_noche_completa(datetime.now(CHILE_TZ))

    node_ids = list(FUNDO_ZAPALLAR_NODE_IDS)
    filas = evaluar_puntos(
        noche,
        node_ids,
        noches=args.noches,
        factor=args.factor,
        workers=args.workers,
    )
    csv_path = _guardar_csv(filas, noche)
    _imprimir_resumen(filas)
    print(f"[INFO] Detalle: {csv_path}")

    alertas = [f for f in filas if f["alerta"]]
    if not alertas:
        print("[INFO] Ningún punto de Fundo Zapallar supera el umbral. No se envía correo.")
        return 0

    estado = _cargar_estado()
    pendientes = []
    for fila in alertas:
        clave = _clave_envio(noche, fila["node_id"])
        if args.forzar_reenvio or clave not in estado["enviados"]:
            pendientes.append(fila)
        else:
            print(f"[INFO] Ya se avisó {fila['nombre']} ({fila['node_id']}) para {noche.isoformat()}.")

    if not pendientes:
        print("[INFO] Esas alertas ya fueron enviadas.")
        return 0
    if args.dry_run:
        print(f"[DRY-RUN] Se enviaría correo por {len(pendientes)} punto(s).")
        return 0

    enviar_correo(pendientes, noche, args.factor)
    for fila in pendientes:
        estado["enviados"].append(_clave_envio(noche, fila["node_id"]))
    _guardar_estado(estado)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
