#!/usr/bin/env python3
"""Primera alerta del agente: caudal máximo de Etapa N°5, Fundo Zapallar.

Revisa el punto 000027-03 a las 08:00, 16:00 y 23:00 (hora Chile).
Solo escribe correo a Aníbal y Juan si alguna hora de esa revisión supera 60 m³/h.
Ese tope es el caudal que entrega una tubería de 3 pulgadas.

  python revisar_alerta_etapa5_zapallar.py
  python revisar_alerta_etapa5_zapallar.py --programado
  python revisar_alerta_etapa5_zapallar.py --revision 16:00 --sin-correo
"""

from __future__ import annotations

import argparse
import json
import smtplib
import sys
from datetime import date, datetime, timedelta
from email.mime.text import MIMEText
from pathlib import Path
from zoneinfo import ZoneInfo

import requests

from generar_alertas_agente import CONFIG_PATH, NODE_BASE_URL, ZONA, horas_desde_csv

ALERTA_ID = "fundo-zapallar-etapa-5-maximo"
REVISIONES = ("08:00", "16:00", "23:00")
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
ROOT = Path(__file__).resolve().parent
ESTADO_PATH = ROOT / "reports" / "alertas_agente" / "etapa5_envios.json"


class RevisionError(Exception):
    """La revisión no pudo leer el punto o enviar el correo."""


def cargar_alerta(path: Path = CONFIG_PATH) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    for alerta in data.get("alertas") or []:
        if alerta.get("id") == ALERTA_ID and alerta.get("activa", True):
            return alerta
    raise RevisionError(f"No está activa la alerta {ALERTA_ID} en {path}.")


def horas_de_la_revision(dia: date, revision: str) -> list[tuple[date, int]]:
    """Horas ya cerradas que corresponde mirar en esa revisión."""
    if revision == "08:00":
        ayer = dia - timedelta(days=1)
        return [(ayer, 23), *[(dia, hora) for hora in range(0, 8)]]
    if revision == "16:00":
        return [(dia, hora) for hora in range(8, 16)]
    if revision == "23:00":
        return [(dia, hora) for hora in range(16, 23)]
    raise RevisionError(f"Revisión desconocida: {revision}")


def revision_para(ahora: datetime, programado: bool) -> tuple[date, str] | None:
    momento = ahora.astimezone(ZONA)
    if programado:
        revision = {8: "08:00", 16: "16:00", 23: "23:00"}.get(momento.hour)
        if revision is None:
            return None
        return momento.date(), revision
    if momento.hour >= 23:
        return momento.date(), "23:00"
    if momento.hour >= 16:
        return momento.date(), "16:00"
    if momento.hour >= 8:
        return momento.date(), "08:00"
    return momento.date() - timedelta(days=1), "23:00"


def excesos(
    lecturas: dict[tuple[date, int], float | None],
    pedidas: list[tuple[date, int]],
    umbral: float,
) -> list[tuple[date, int, float]]:
    encontrados = []
    for clave in pedidas:
        valor = lecturas.get(clave)
        if valor is not None and float(valor) > umbral:
            encontrados.append((clave[0], clave[1], float(valor)))
    return encontrados


def _fmt_m3(valor: float) -> str:
    return f"{valor:.3f}".replace(".", ",")


def armar_correo(
    alerta: dict,
    dia: date,
    revision: str,
    filas: list[tuple[date, int, float]],
) -> tuple[str, str]:
    umbral = float(alerta["umbral_m3h"])
    asunto = (
        f"Alerta Fundo Zapallar — Etapa N°5 superó {_fmt_m3(umbral)} m³/h "
        f"({revision})"
    )
    lineas = [
        "Estimados Aníbal y Juan,",
        "",
        (
            f"En la revisión de las {revision} del {dia.strftime('%d-%m-%Y')}, "
            f"{alerta['punto']} del {alerta['cliente']} ({alerta['nodeId']}) "
            f"superó {_fmt_m3(umbral)} m³/h."
        ),
        alerta.get("motivo") or "Ese es el caudal máximo de la tubería de 3 pulgadas.",
        "",
        "Horas sobre el umbral:",
    ]
    for fecha, hora, valor in filas:
        lineas.append(f"- {fecha.strftime('%d-%m-%Y')} {hora:02d}:00 — {_fmt_m3(valor)} m³/h")
    lineas.extend(["", "Saludos,", "Agente WES"])
    return asunto, "\n".join(lineas) + "\n"


def leer_horas(node_id: str, dia: date) -> dict[int, float]:
    url = f"{NODE_BASE_URL}/nodes/{node_id}/dates.measures.csv"
    fecha = dia.strftime("%d%m%Y")
    try:
        respuesta = requests.get(url, params=[("start", fecha), ("end", fecha)], timeout=60)
    except requests.RequestException as exc:
        raise RevisionError(f"No pude leer {node_id} el {dia.isoformat()}: {exc}") from exc
    if respuesta.status_code != 200:
        raise RevisionError(f"No pude leer {node_id} el {dia.isoformat()} (HTTP {respuesta.status_code}).")
    return horas_desde_csv(respuesta.text or "", dia)


def lecturas_de_la_revision(node_id: str, pedidas: list[tuple[date, int]]) -> dict[tuple[date, int], float | None]:
    por_dia: dict[date, dict[int, float]] = {}
    salida: dict[tuple[date, int], float | None] = {}
    for dia, _hora in pedidas:
        if dia not in por_dia:
            por_dia[dia] = leer_horas(node_id, dia)
    for dia, hora in pedidas:
        serie = por_dia[dia]
        salida[(dia, hora)] = serie.get(hora) if hora in serie else None
    return salida


def _smtp_password() -> str:
    import os

    password = (
        os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
        or os.environ.get("WES_SMTP_PASSWORD", "").strip()
    )
    if password:
        return password.replace(" ", "")
    archivo = ROOT / "gmail_oauth" / "app_password.txt"
    if archivo.is_file():
        for linea in archivo.read_text(encoding="utf-8", errors="replace").splitlines():
            linea = linea.strip()
            if linea and not linea.startswith("#"):
                return linea.replace(" ", "")
    from enviar_control_nocturno_pdf import SMTP_PASSWORD

    return (SMTP_PASSWORD or "").replace(" ", "").strip()


def enviar_correo(destinatarios: list[str], asunto: str, cuerpo: str) -> None:
    password = _smtp_password()
    if not password:
        raise RevisionError("Falta la contraseña SMTP para avisar a Aníbal y Juan.")
    mensaje = MIMEText(cuerpo, "plain", "utf-8")
    mensaje["From"] = SMTP_USUARIO
    mensaje["To"] = ", ".join(destinatarios)
    mensaje["Subject"] = asunto
    try:
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO, timeout=30) as server:
            server.starttls()
            server.login(SMTP_USUARIO, password)
            server.sendmail(SMTP_USUARIO, destinatarios, mensaje.as_string())
    except (OSError, smtplib.SMTPException) as exc:
        raise RevisionError(f"No pude enviar el correo: {exc}") from exc


def _ya_avisado(dia: date, revision: str) -> bool:
    if not ESTADO_PATH.is_file():
        return False
    try:
        data = json.loads(ESTADO_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False
    return f"{dia.isoformat()} {revision}" in set(data.get("avisos") or [])


def _marcar_avisado(dia: date, revision: str) -> None:
    avisos: list[str] = []
    if ESTADO_PATH.is_file():
        try:
            avisos = list(json.loads(ESTADO_PATH.read_text(encoding="utf-8")).get("avisos") or [])
        except (OSError, json.JSONDecodeError):
            avisos = []
    clave = f"{dia.isoformat()} {revision}"
    if clave not in avisos:
        avisos.append(clave)
    ESTADO_PATH.parent.mkdir(parents=True, exist_ok=True)
    ESTADO_PATH.write_text(json.dumps({"avisos": avisos[-60:]}, indent=2) + "\n", encoding="utf-8")


def revisar(
    alerta: dict,
    dia: date,
    revision: str,
    *,
    enviar: bool,
    lecturas: dict[tuple[date, int], float | None] | None = None,
) -> list[tuple[date, int, float]]:
    pedidas = horas_de_la_revision(dia, revision)
    if lecturas is None:
        lecturas = lecturas_de_la_revision(str(alerta["nodeId"]), pedidas)
    filas = excesos(lecturas, pedidas, float(alerta["umbral_m3h"]))
    presentes = [float(lecturas[clave]) for clave in pedidas if lecturas.get(clave) is not None]
    if presentes:
        mayor = max(presentes)
        print(f"Máximo de la revisión {revision}: {_fmt_m3(mayor)} m³/h (umbral {_fmt_m3(float(alerta['umbral_m3h']))}).")
    else:
        print(f"Revisión {revision}: sin lecturas en las horas pedidas.")
    if not filas:
        print(f"{revision} del {dia.isoformat()}: Etapa N°5 no superó {alerta['umbral_m3h']} m³/h. Sin correo.")
        return []
    asunto, cuerpo = armar_correo(alerta, dia, revision, filas)
    print(asunto)
    print(cuerpo)
    if not enviar:
        print("Sin envío (--sin-correo).")
        return filas
    if _ya_avisado(dia, revision):
        print("Ese aviso ya se envió. No se repite el correo.")
        return filas
    destinatarios = list(alerta["correos"])
    enviar_correo(destinatarios, asunto, cuerpo)
    _marcar_avisado(dia, revision)
    print("Correo enviado a " + ", ".join(destinatarios))
    return filas


def main(argv: list[str] | None = None) -> int:
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Revisa el caudal máximo de Etapa N°5.")
    parser.add_argument("--programado", action="store_true", help="Solo corre a las 08:00, 16:00 o 23:00.")
    parser.add_argument("--revision", choices=REVISIONES)
    parser.add_argument("--fecha", help="Día de la revisión, AAAA-MM-DD.")
    parser.add_argument("--sin-correo", action="store_true")
    args = parser.parse_args(argv)
    ahora = datetime.now(ZONA)
    try:
        alerta = cargar_alerta()
        if args.revision:
            dia = date.fromisoformat(args.fecha) if args.fecha else ahora.date()
            revision = args.revision
        else:
            elegida = revision_para(ahora, args.programado)
            if elegida is None:
                print(f"{ahora.strftime('%H:%M')} no es horario de revisión.")
                return 0
            dia, revision = elegida
            if args.fecha:
                dia = date.fromisoformat(args.fecha)
        revisar(alerta, dia, revision, enviar=not args.sin_correo)
    except RevisionError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
