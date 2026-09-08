#!/usr/bin/env python3
"""
Revisión de correo agente.ia@wes.cl con aprobación humana (Aníbal).

Flujo (lun–vie 08:00–18:00 Chile, salvo --forzar):
  1) Procesa aprobaciones: si Aníbal responde DALE NOMÁS a un [BORRADOR ...],
     envía la respuesta al destinatario original.
  2) Revisa UNSEEN del inbox, arma borradores y los envía SOLO a Aníbal.
     Nunca responde al cliente sin esa aprobación.

Uso (Automation Cursor / local):
  python revisar_correo_wes_borradores.py
  python revisar_correo_wes_borradores.py --forzar
  python revisar_correo_wes_borradores.py --solo-aprobaciones
  python revisar_correo_wes_borradores.py --solo-borradores

Secretos Cloud Agents:
  WES_GMAIL_APP_PASSWORD  (contraseña de aplicación de agente.ia@wes.cl)
"""

from __future__ import annotations

import argparse
import email
import imaplib
import json
import os
import re
import smtplib
import sys
import uuid
from datetime import datetime, timedelta, timezone
from email.header import decode_header, make_header
from email.message import EmailMessage
from email.utils import formatdate, make_msgid, parseaddr, parsedate_to_datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    from zoneinfo import ZoneInfo

    _CHILE_TZ = ZoneInfo("America/Santiago")
except Exception:  # pragma: no cover
    _CHILE_TZ = timezone(timedelta(hours=-4))

ROOT = Path(__file__).resolve().parent
STATE_DIR = ROOT / "logs" / "borradores_correo"
STATE_DIR.mkdir(parents=True, exist_ok=True)
STATE_INDEX = STATE_DIR / "index.json"

SMTP_USUARIO = os.environ.get("WES_SMTP_USUARIO", "agente.ia@wes.cl").strip()
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
IMAP_SERVIDOR = "imap.gmail.com"
IMAP_PUERTO = 993

ANIBAL_EMAIL = os.environ.get(
    "WES_APROBADOR_EMAIL", "anibal.aoperaciones@wes.cl"
).strip().lower()

HORA_INI = int(os.environ.get("WES_CORREO_HORA_INI", "8"))
HORA_FIN = int(os.environ.get("WES_CORREO_HORA_FIN", "18"))  # exclusivo
MAX_UNSEEN = int(os.environ.get("WES_CORREO_MAX_UNSEEN", "15"))

# Remitentes que no generan borrador de respuesta al cliente.
IGNORAR_REMITENTES = {
    SMTP_USUARIO.lower(),
    ANIBAL_EMAIL,
    "mailer-daemon@googlemail.com",
    "noreply@wes.cl",
}
IGNORAR_ASUNTO_PREFIX = (
    "[borrador ",
    "undeliverable",
    "delivery status notification",
)

PALABRAS_REPORTE = (
    "reporte",
    "informe",
    "agregado",
    "consumo",
    "puntos en cero",
    "fuga",
    "alerta",
    "word",
    "pdf",
    "excel",
)


def _password() -> str:
    pwd = (
        os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
        or os.environ.get("WES_SMTP_PASSWORD", "").strip()
    )
    if pwd:
        return pwd.replace(" ", "")
    app_file = ROOT / "gmail_oauth" / "app_password.txt"
    if app_file.is_file():
        for line in app_file.read_text(encoding="utf-8", errors="replace").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                return line.replace(" ", "")
    return ""


def _decode_mime(value: Optional[str]) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def _ahora_chile() -> datetime:
    return datetime.now(_CHILE_TZ)


def en_ventana_laboral(ahora: Optional[datetime] = None) -> bool:
    ahora = ahora or _ahora_chile()
    if ahora.weekday() >= 5:  # sáb/dom
        return False
    return HORA_INI <= ahora.hour < HORA_FIN


def _load_index() -> Dict[str, Any]:
    if not STATE_INDEX.is_file():
        return {"drafts": {}, "seen_message_ids": []}
    try:
        return json.loads(STATE_INDEX.read_text(encoding="utf-8"))
    except Exception:
        return {"drafts": {}, "seen_message_ids": []}


def _save_index(data: Dict[str, Any]) -> None:
    STATE_INDEX.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _extraer_cuerpo(msg: email.message.Message) -> str:
    textos: List[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition") or "")
            if "attachment" in disp.lower():
                continue
            if ctype == "text/plain":
                payload = part.get_payload(decode=True) or b""
                charset = part.get_content_charset() or "utf-8"
                textos.append(payload.decode(charset, errors="replace"))
            elif ctype == "text/html" and not textos:
                payload = part.get_payload(decode=True) or b""
                charset = part.get_content_charset() or "utf-8"
                html = payload.decode(charset, errors="replace")
                textos.append(re.sub(r"<[^>]+>", " ", html))
    else:
        payload = msg.get_payload(decode=True) or b""
        charset = msg.get_content_charset() or "utf-8"
        textos.append(payload.decode(charset, errors="replace"))
    cuerpo = "\n".join(textos).strip()
    cuerpo = re.sub(r"\n{3,}", "\n\n", cuerpo)
    return cuerpo[:8000]


def _imap_connect() -> imaplib.IMAP4_SSL:
    pwd = _password()
    if not pwd:
        raise RuntimeError(
            "Falta WES_GMAIL_APP_PASSWORD (Cloud Agents → Secrets) "
            "o gmail_oauth/app_password.txt"
        )
    mail = imaplib.IMAP4_SSL(IMAP_SERVIDOR, IMAP_PUERTO)
    mail.login(SMTP_USUARIO, pwd)
    mail.select("INBOX")
    return mail


def _smtp_send(msg: EmailMessage) -> None:
    pwd = _password()
    if not pwd:
        raise RuntimeError("Falta contraseña SMTP / WES_GMAIL_APP_PASSWORD")
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO, timeout=60) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pwd)
        server.send_message(msg)


def _clasificar(asunto: str, cuerpo: str) -> str:
    blob = f"{asunto}\n{cuerpo}".lower()
    if any(p in blob for p in PALABRAS_REPORTE):
        return "posible_solicitud_reporte"
    if "?" in blob or any(
        w in blob for w in ("consulta", "puedes", "podrías", "necesito", "urgente")
    ):
        return "consulta"
    return "otro"


def _sugerir_respuesta(
    tipo: str, remitente_nombre: str, asunto: str, cuerpo: str
) -> str:
    saludo = f"Hola {remitente_nombre}," if remitente_nombre else "Hola,"
    if tipo == "posible_solicitud_reporte":
        return (
            f"{saludo}\n\n"
            "Recibimos tu correo. Estamos preparando la información / reporte "
            "solicitado y te lo enviaremos a la brevedad.\n\n"
            "Quedo atento.\n\n"
            "Saludos,\nEquipo WES"
        )
    if tipo == "consulta":
        return (
            f"{saludo}\n\n"
            "Gracias por escribirnos. Revisamos tu consulta y te respondemos "
            "con el detalle correspondiente.\n\n"
            f"(Referencia asunto: {asunto})\n\n"
            "Saludos,\nEquipo WES"
        )
    return (
        f"{saludo}\n\n"
        "Gracias por tu correo. Lo estamos revisando y te contactamos "
        "con una respuesta.\n\n"
        "Saludos,\nEquipo WES"
    )


def _parse_msg(raw: bytes) -> Dict[str, Any]:
    msg = email.message_from_bytes(raw)
    from_name, from_email = parseaddr(msg.get("From", ""))
    from_email = (from_email or "").lower()
    asunto = _decode_mime(msg.get("Subject"))
    message_id = (msg.get("Message-ID") or "").strip()
    in_reply_to = (msg.get("In-Reply-To") or "").strip()
    references = (msg.get("References") or "").strip()
    date_hdr = msg.get("Date")
    try:
        fecha = parsedate_to_datetime(date_hdr).isoformat() if date_hdr else ""
    except Exception:
        fecha = ""
    cuerpo = _extraer_cuerpo(msg)
    return {
        "from_name": _decode_mime(from_name) or from_email,
        "from_email": from_email,
        "asunto": asunto,
        "message_id": message_id,
        "in_reply_to": in_reply_to,
        "references": references,
        "fecha": fecha,
        "cuerpo": cuerpo,
    }


def _uid_search(mail: imaplib.IMAP4_SSL, criteria: str) -> List[bytes]:
    status, data = mail.uid("search", None, criteria)
    if status != "OK" or not data or not data[0]:
        return []
    return data[0].split()


def _fetch_uid(mail: imaplib.IMAP4_SSL, uid: bytes) -> Optional[Dict[str, Any]]:
    typ, msg_data = mail.uid("fetch", uid, "(RFC822)")
    if typ != "OK" or not msg_data or not msg_data[0]:
        return None
    raw = msg_data[0]
    if isinstance(raw, tuple):
        raw = raw[1]
    if not isinstance(raw, (bytes, bytearray)):
        return None
    parsed = _parse_msg(bytes(raw))
    parsed["imap_uid"] = uid.decode() if isinstance(uid, bytes) else str(uid)
    return parsed


def _marcar_leido(mail: imaplib.IMAP4_SSL, uid: bytes) -> None:
    try:
        mail.uid("store", uid, "+FLAGS", r"(\Seen)")
    except Exception as exc:
        print(f"  [ADVERTENCIA] No se pudo marcar leído UID {uid!r}: {exc}")


def _es_aprobacion(parsed: Dict[str, Any]) -> Optional[str]:
    """Si el correo es de Aníbal con DALE NOMÁS, retorna draft_id."""
    if parsed["from_email"] != ANIBAL_EMAIL:
        return None
    asunto = parsed["asunto"] or ""
    cuerpo = parsed["cuerpo"] or ""
    blob = f"{asunto}\n{cuerpo}"
    if re.search(r"\bCANCELAR\b", blob, re.IGNORECASE):
        return None
    if not re.search(r"\bDALE\s*NOM[AÁ]S\b", blob, re.IGNORECASE):
        return None
    m = re.search(r"\[BORRADOR\s+([A-Za-z0-9\-]+)\]", asunto, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r"\bDALE\s*NOM[AÁ]S\s+([A-Za-z0-9\-]{6,})\b", blob, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r"\bID[:\s]+([A-Za-z0-9\-]{6,})\b", blob, re.IGNORECASE)
    if m:
        return m.group(1)
    return None


def _respuesta_aprobada_editada(cuerpo: str) -> Optional[str]:
    """
    Si Aníbal escribe texto después de DALE NOMÁS (antes de citas), úsalo
    como respuesta final.
    """
    m = re.search(
        r"DALE\s*NOM[AÁ]S(?:\s+[A-Za-z0-9\-]+)?\s*[:\-]?\s*(.*)",
        cuerpo,
        re.IGNORECASE | re.DOTALL,
    )
    if not m:
        return None
    resto = m.group(1).strip()
    # Cortar citas típicas
    for sep in ("\n>", "\nEl ", "\nFrom:", "\nDe:", "-----Original Message-----"):
        if sep in resto:
            resto = resto.split(sep, 1)[0].strip()
    if len(resto) < 20:
        return None
    return resto


def _parse_bloque_meta(cuerpo: str) -> Dict[str, str]:
    """Lee el bloque META embebido en el mail [BORRADOR] (sobrevive entre VMs)."""
    m = re.search(
        r"===== META BORRADOR \(NO BORRAR\) =====\s*(.*?)\s*===== FIN META =====",
        cuerpo,
        re.IGNORECASE | re.DOTALL,
    )
    if not m:
        return {}
    meta: Dict[str, str] = {}
    for line in m.group(1).splitlines():
        if ":" not in line:
            continue
        k, v = line.split(":", 1)
        meta[k.strip().lower()] = v.strip()
    return meta


def _recuperar_borrador_desde_hilo(
    mail: imaplib.IMAP4_SSL, draft_id: str, aprobacion: Dict[str, Any]
) -> Optional[Dict[str, Any]]:
    """Si el index local no tiene el draft (Cloud Agent efímero), lo arma desde el aviso."""

    def _draft_from_parsed(parsed: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        meta = _parse_bloque_meta(parsed["cuerpo"])
        if not meta.get("to"):
            return None
        sugerida = ""
        m = re.search(
            r"========== RESPUESTA SUGERIDA ==========\s*(.*?)\s*========== CÓMO APROBAR",
            parsed["cuerpo"],
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            sugerida = m.group(1).strip()
        asunto_resp = meta.get("asunto_respuesta") or f"Re: {meta.get('asunto', '')}"
        return {
            "id": draft_id,
            "status": "pendiente",
            "to_email": meta.get("to", ""),
            "asunto_original": meta.get("asunto", ""),
            "asunto_respuesta": asunto_resp,
            "message_id": meta.get("message_id_original", ""),
            "references": meta.get("references", ""),
            "respuesta_sugerida": sugerida,
        }

    # 1) META citado en la reply de Aníbal (lo más fiable en Cloud)
    draft = _draft_from_parsed(aprobacion)
    if draft:
        return draft

    # 2) Buscar el aviso en Sent (Gmail)
    try:
        status, _ = mail.select('"[Gmail]/Sent Mail"')
        if status == "OK":
            uids = _uid_search(mail, f'(SUBJECT "[BORRADOR {draft_id}]")')
            uids = sorted(uids, key=lambda u: int(u), reverse=True)[:5]
            for uid in uids:
                parsed = _fetch_uid(mail, uid)
                if not parsed:
                    continue
                draft = _draft_from_parsed(parsed)
                if draft:
                    mail.select("INBOX")
                    return draft
    except Exception as exc:
        print(f"  [ADVERTENCIA] No se pudo leer Sent Mail: {exc}")
    finally:
        try:
            mail.select("INBOX")
        except Exception:
            pass

    return None


def _ya_enviado(mail: imaplib.IMAP4_SSL, draft_id: str) -> bool:
    # Confirmaciones van a Aníbal; también pueden estar en Sent
    if _uid_search(mail, f'(SUBJECT "[ENVIADO] Borrador {draft_id}")'):
        return True
    try:
        status, _ = mail.select('"[Gmail]/Sent Mail"')
        if status == "OK":
            found = bool(
                _uid_search(mail, f'(SUBJECT "[ENVIADO] Borrador {draft_id}")')
            )
            mail.select("INBOX")
            return found
    except Exception:
        try:
            mail.select("INBOX")
        except Exception:
            pass
    return False


def procesar_aprobaciones(mail: imaplib.IMAP4_SSL, index: Dict[str, Any]) -> int:
    enviados = 0
    uids = _uid_search(mail, "UNSEEN")
    uids_extra = _uid_search(
        mail, f'(FROM "{ANIBAL_EMAIL}" SUBJECT "[BORRADOR")'
    )
    seen = set(uids)
    for u in uids_extra:
        if u not in seen:
            uids.append(u)
            seen.add(u)
    uids = sorted(uids, key=lambda u: int(u), reverse=True)[:30]

    print(f"[INFO] Revisando aprobaciones en {len(uids)} correo(s)...")
    for uid in uids:
        parsed = _fetch_uid(mail, uid)
        if not parsed:
            continue
        draft_id = _es_aprobacion(parsed)
        if not draft_id:
            continue

        if _ya_enviado(mail, draft_id):
            print(f"  [INFO] Borrador {draft_id} ya tiene [ENVIADO]; se ignora.")
            _marcar_leido(mail, uid)
            continue

        draft = index.get("drafts", {}).get(draft_id)
        if not draft:
            draft = _recuperar_borrador_desde_hilo(mail, draft_id, parsed)
        if not draft:
            print(
                f"  [ADVERTENCIA] Aprobación {draft_id} sin datos de borrador "
                "(¿se borró el aviso [BORRADOR]?)."
            )
            continue
        if draft.get("status") == "enviado":
            print(f"  [INFO] Borrador {draft_id} ya enviado; se ignora.")
            _marcar_leido(mail, uid)
            continue

        respuesta = _respuesta_aprobada_editada(parsed["cuerpo"]) or draft.get(
            "respuesta_sugerida", ""
        )
        to_email = draft.get("to_email")
        if not to_email or not respuesta:
            print(f"  [ERROR] Borrador {draft_id} incompleto (to/respuesta)")
            continue

        out = EmailMessage()
        out["From"] = SMTP_USUARIO
        out["To"] = to_email
        out["Subject"] = draft.get("asunto_respuesta") or f"Re: {draft.get('asunto_original', '')}"
        out["Date"] = formatdate(localtime=True)
        out["Message-ID"] = make_msgid(domain="wes.cl")
        if draft.get("message_id"):
            out["In-Reply-To"] = draft["message_id"]
            refs = (draft.get("references") or "").strip()
            out["References"] = f"{refs} {draft['message_id']}".strip()
        out.set_content(respuesta)

        try:
            _smtp_send(out)
        except Exception as exc:
            print(f"  [ERROR] No se pudo enviar respuesta {draft_id}: {exc}")
            continue

        draft["status"] = "enviado"
        draft["enviado_at"] = _ahora_chile().isoformat()
        draft["aprobado_por_message_id"] = parsed.get("message_id")
        index.setdefault("drafts", {})[draft_id] = draft
        _save_index(index)
        _marcar_leido(mail, uid)
        enviados += 1
        print(f"  [OK] Enviado al cliente ({to_email}) borrador {draft_id}")

        conf = EmailMessage()
        conf["From"] = SMTP_USUARIO
        conf["To"] = ANIBAL_EMAIL
        conf["Subject"] = f"[ENVIADO] Borrador {draft_id} → {to_email}"
        conf["Date"] = formatdate(localtime=True)
        conf.set_content(
            f"Se envió la respuesta aprobada.\n\n"
            f"ID: {draft_id}\n"
            f"Para: {to_email}\n"
            f"Asunto: {out['Subject']}\n\n"
            f"--- Respuesta ---\n{respuesta}\n"
        )
        try:
            _smtp_send(conf)
        except Exception as exc:
            print(f"  [ADVERTENCIA] Confirmación a Aníbal falló: {exc}")

    return enviados


def _debe_ignorar(parsed: Dict[str, Any], index: Dict[str, Any]) -> bool:
    email_from = parsed["from_email"]
    if not email_from:
        return True
    if email_from in IGNORAR_REMITENTES:
        return True
    # Equipo interno: no generar borrador automático (evita ruido).
    if email_from.endswith("@wes.cl"):
        return True
    asunto_l = (parsed["asunto"] or "").lower().strip()
    if any(asunto_l.startswith(p) for p in IGNORAR_ASUNTO_PREFIX):
        return True
    mid = parsed.get("message_id") or ""
    if mid and mid in index.get("seen_message_ids", []):
        return True
    for d in index.get("drafts", {}).values():
        if d.get("message_id") == mid and d.get("status") in {"pendiente", "enviado"}:
            return True
    return False


def crear_borradores(mail: imaplib.IMAP4_SSL, index: Dict[str, Any]) -> int:
    uids = _uid_search(mail, "UNSEEN")
    uids = sorted(uids, key=lambda u: int(u), reverse=True)[:MAX_UNSEEN]
    print(f"[INFO] UNSEEN encontrados: {len(uids)}")
    creados = 0

    for uid in uids:
        parsed = _fetch_uid(mail, uid)
        if not parsed:
            continue

        # Aprobaciones se manejan en otra función; no marcar aquí si es Aníbal+DALE
        if _es_aprobacion(parsed):
            continue

        if _debe_ignorar(parsed, index):
            # Correos internos/ruido: marcar leído para no ciclar
            if parsed["from_email"] in IGNORAR_REMITENTES or (
                parsed["asunto"] or ""
            ).lower().startswith("[borrador "):
                _marcar_leido(mail, uid)
            print(
                f"  [SKIP] {parsed['from_email']} | {parsed['asunto'][:60]}"
            )
            continue

        tipo = _clasificar(parsed["asunto"], parsed["cuerpo"])
        draft_id = uuid.uuid4().hex[:10]
        respuesta = _sugerir_respuesta(
            tipo, parsed["from_name"].split()[0] if parsed["from_name"] else "",
            parsed["asunto"], parsed["cuerpo"],
        )
        asunto_resp = parsed["asunto"]
        if not asunto_resp.lower().startswith("re:"):
            asunto_resp = f"Re: {asunto_resp}"

        draft = {
            "id": draft_id,
            "status": "pendiente",
            "tipo": tipo,
            "created_at": _ahora_chile().isoformat(),
            "to_email": parsed["from_email"],
            "to_name": parsed["from_name"],
            "asunto_original": parsed["asunto"],
            "asunto_respuesta": asunto_resp,
            "message_id": parsed["message_id"],
            "references": parsed["references"],
            "imap_uid": parsed["imap_uid"],
            "cuerpo_original": parsed["cuerpo"][:4000],
            "respuesta_sugerida": respuesta,
        }
        index.setdefault("drafts", {})[draft_id] = draft
        if parsed["message_id"]:
            seen = index.setdefault("seen_message_ids", [])
            seen.append(parsed["message_id"])
            index["seen_message_ids"] = seen[-500:]

        # Guardar JSON individual
        (STATE_DIR / f"{draft_id}.json").write_text(
            json.dumps(draft, ensure_ascii=False, indent=2), encoding="utf-8"
        )

        # Mail a Aníbal
        aviso = EmailMessage()
        aviso["From"] = SMTP_USUARIO
        aviso["To"] = ANIBAL_EMAIL
        aviso["Subject"] = f"[BORRADOR {draft_id}] {parsed['asunto']}"
        aviso["Date"] = formatdate(localtime=True)
        aviso["Message-ID"] = make_msgid(domain="wes.cl")
        meta_block = (
            "===== META BORRADOR (NO BORRAR) =====\n"
            f"id: {draft_id}\n"
            f"to: {parsed['from_email']}\n"
            f"asunto: {parsed['asunto']}\n"
            f"asunto_respuesta: {asunto_resp}\n"
            f"message_id_original: {parsed['message_id']}\n"
            f"references: {parsed['references']}\n"
            f"tipo: {tipo}\n"
            "===== FIN META =====\n"
        )
        cuerpo_aviso = (
            f"Hola Aníbal,\n\n"
            f"Hay un correo en {SMTP_USUARIO} que necesita respuesta.\n\n"
            f"{meta_block}\n"
            f"ID borrador: {draft_id}\n"
            f"Tipo: {tipo}\n"
            f"De: {parsed['from_name']} <{parsed['from_email']}>\n"
            f"Asunto: {parsed['asunto']}\n"
            f"Fecha: {parsed['fecha']}\n\n"
            f"========== CORREO ORIGINAL ==========\n"
            f"{parsed['cuerpo'][:3500]}\n\n"
            f"========== RESPUESTA SUGERIDA ==========\n"
            f"{respuesta}\n\n"
            f"========== CÓMO APROBAR ==========\n"
            f"1) Revisá el borrador (y el informe si aplica).\n"
            f"2) Respondé ESTE correo con la primera línea:\n"
            f"   DALE NOMÁS {draft_id}\n"
            f"3) Opcional: debajo de esa línea pegá una respuesta editada;\n"
            f"   si no, se usa la sugerida.\n"
            f"4) Si no va: no respondás DALE (o escribí CANCELAR).\n\n"
            f"Hasta que apruebes, el cliente NO recibe nada.\n"
        )
        aviso.set_content(cuerpo_aviso)

        try:
            _smtp_send(aviso)
            print(
                f"  [OK] Borrador {draft_id} → Aníbal | {parsed['from_email']} | {tipo}"
            )
            creados += 1
            _marcar_leido(mail, uid)
        except Exception as exc:
            print(f"  [ERROR] No se pudo avisar borrador {draft_id}: {exc}")
            # No marcar leído para reintentar
            index["drafts"].pop(draft_id, None)

        _save_index(index)

    return creados


def main(argv: Optional[List[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    ap = argparse.ArgumentParser(description="Correo WES 8-18 con aprobación Aníbal")
    ap.add_argument("--forzar", action="store_true", help="Ignorar ventana lun–vie 8–18")
    ap.add_argument("--solo-aprobaciones", action="store_true")
    ap.add_argument("--solo-borradores", action="store_true")
    args = ap.parse_args(argv)

    ahora = _ahora_chile()
    print("=" * 70)
    print("REVISIÓN CORREO WES → BORRADORES A ANÍBAL")
    print("=" * 70)
    print(f"Hora Chile: {ahora:%Y-%m-%d %H:%M (%A)}")
    print(f"Cuenta: {SMTP_USUARIO}")
    print(f"Aprobador: {ANIBAL_EMAIL}")
    print()

    if not args.forzar and not en_ventana_laboral(ahora):
        print(
            f"[INFO] Fuera de ventana laboral "
            f"(lun–vie {HORA_INI:02d}:00–{HORA_FIN:02d}:00 Chile). Nada que hacer."
        )
        return 0

    index = _load_index()
    mail = None
    try:
        mail = _imap_connect()
        enviados = 0
        creados = 0
        if not args.solo_borradores:
            enviados = procesar_aprobaciones(mail, index)
        if not args.solo_aprobaciones:
            creados = crear_borradores(mail, index)
        print()
        print(f"[RESUMEN] borradores_nuevos={creados} enviados_aprobados={enviados}")
        print(f"[RESUMEN] estado={STATE_INDEX}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}")
        import traceback

        traceback.print_exc()
        return 1
    finally:
        if mail is not None:
            try:
                mail.logout()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
