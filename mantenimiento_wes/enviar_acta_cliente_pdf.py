# -*- coding: utf-8 -*-
"""Envía el PDF de acta al cliente solicitando acusar recibo."""

from __future__ import annotations

import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

SMTP_USUARIO = os.environ.get("WES_SMTP_USUARIO", "agente.ia@wes.cl").strip()
SMTP_SERVIDOR = os.environ.get("WES_SMTP_SERVIDOR", "smtp.gmail.com").strip()
SMTP_PUERTO = int(os.environ.get("WES_SMTP_PUERTO", "587"))


def _smtp_password() -> str:
    return (
        os.environ.get("SMTP_PASSWORD", "").strip()
        or os.environ.get("WES_SMTP_PASSWORD", "").strip()
        or os.environ.get("GMAIL_APP_PASSWORD", "").strip()
    )


def _split_emails(raw: str) -> List[str]:
    if not raw:
        return []
    parts = []
    for chunk in str(raw).replace(";", ",").split(","):
        e = chunk.strip()
        if e and "@" in e:
            parts.append(e)
    return parts


def construir_cuerpo(data: Dict[str, Any]) -> Tuple[str, str, str]:
    cliente = data.get("cliente") or "Cliente"
    maquina = data.get("maquina") or "sitio"
    tecnico = data.get("tecnico") or "equipo WES"
    fecha = data.get("fecha") or ""
    recibido = data.get("recibido_por") or ""

    subject = f"WES · Acta de visita técnica — {cliente} / {maquina} · Acusar recibo"
    text = f"""Estimados/as {cliente},

Adjuntamos el acta en PDF de la visita técnica realizada por WES en {maquina}
(fecha {fecha}, técnico {tecnico}).

Solicitamos por favor ACUSAR RECIBO de esta acta respondiendo este correo
con la frase «Acuso recibo» (puede indicar nombre y cargo).

Quien recibió en terreno: {recibido or '—'}.

Quedamos atentos.
— Sociedad Tecnológica WES SpA
www.wes.cl · agente.ia@wes.cl
"""
    html = f"""
    <div style="font-family:Segoe UI,Arial,sans-serif;color:#14202b;line-height:1.45">
      <p>Estimados/as <b>{cliente}</b>,</p>
      <p>Adjuntamos el <b>acta en PDF</b> de la visita técnica realizada por WES en
      <b>{maquina}</b> (fecha {fecha}, técnico {tecnico}).</p>
      <p style="background:#e7f0f8;border-left:4px solid #1f4e79;padding:12px 14px">
        Solicitamos por favor <b>ACUSAR RECIBO</b> de esta acta respondiendo este correo
        con la frase «Acuso recibo» (puede indicar nombre y cargo).
      </p>
      <p>Quien recibió en terreno: {recibido or '—'}.</p>
      <p>Quedamos atentos.<br/>— Sociedad Tecnológica WES SpA<br/>
      <a href="https://www.wes.cl">www.wes.cl</a> · agente.ia@wes.cl</p>
    </div>
    """
    return subject, text, html


def enviar_acta_pdf_cliente(
    pdf_path: Path,
    data: Dict[str, Any],
    *,
    dry_run: bool = False,
) -> Dict[str, Any]:
    to_list = _split_emails(str(data.get("email_cliente") or ""))
    if not to_list:
        return {"ok": False, "skip": "Falta email_cliente"}

    cc_list = _split_emails(str(data.get("email_cc") or ""))
    # CC interno por defecto si no viene
    if not cc_list:
        cc_list = ["anibal.aoperaciones@wes.cl"]

    pw = _smtp_password()
    if dry_run:
        return {
            "ok": True,
            "dry_run": True,
            "to": to_list,
            "cc": cc_list,
            "pdf": str(pdf_path),
        }
    if not pw:
        return {
            "ok": False,
            "skip": "Falta SMTP_PASSWORD / WES_SMTP_PASSWORD en el entorno",
            "to": to_list,
        }

    subject, text, html = construir_cuerpo(data)
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(text, "plain", "utf-8"))
    alt.attach(MIMEText(html, "html", "utf-8"))
    msg.attach(alt)

    with Path(pdf_path).open("rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=Path(pdf_path).name)
        msg.attach(part)

    destinatarios = to_list + cc_list
    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.sendmail(SMTP_USUARIO, destinatarios, msg.as_string())

    return {"ok": True, "to": to_list, "cc": cc_list, "pdf": str(pdf_path)}
