# -*- coding: utf-8 -*-
"""
Servidor local GRATUITO para firmar en el teléfono (misma WiFi que el PC).

Uso:
  python servir_firma.py

En el teléfono abrí la URL que imprime (ej. http://192.168.x.x:8765).
Las firmas se guardan en: mantenimiento wes/firmas/
"""

from __future__ import annotations

import json
import re
import socket
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

ROOT = Path(__file__).resolve().parent
HTML = ROOT / "firma_visita.html"
FIRMAS = ROOT / "firmas"
PORT = 8765


def _lan_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except OSError:
        return "127.0.0.1"


def _safe_name(name: str) -> str:
    name = unquote(name or "firma.png")
    name = Path(name).name
    name = re.sub(r"[^\w.\-]+", "_", name, flags=re.UNICODE)
    return name[:120] or f"firma_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt: str, *args) -> None:
        print(f"[firma] {self.address_string()} - {fmt % args}")

    def _send(self, code: int, body: bytes, content_type: str) -> None:
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:  # noqa: N802
        if self.path in ("/", "/firma", "/firma_visita.html"):
            if not HTML.is_file():
                self._send(404, b"Falta firma_visita.html", "text/plain; charset=utf-8")
                return
            self._send(200, HTML.read_bytes(), "text/html; charset=utf-8")
            return
        if self.path == "/health":
            self._send(200, b'{"ok":true}', "application/json")
            return
        self._send(404, b"Not found", "text/plain; charset=utf-8")

    def do_POST(self) -> None:  # noqa: N802
        if self.path != "/api/guardar_firma":
            self._send(404, b"Not found", "text/plain; charset=utf-8")
            return
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length)
        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            self._send(400, b'{"error":"multipart required"}', "application/json")
            return

        boundary = None
        for part in content_type.split(";"):
            part = part.strip()
            if part.startswith("boundary="):
                boundary = part.split("=", 1)[1].strip().strip('"')
        if not boundary:
            self._send(400, b'{"error":"no boundary"}', "application/json")
            return

        fields, files = _parse_multipart(raw, boundary.encode("ascii", "ignore"))
        blob = files.get("file")
        if not blob:
            self._send(400, b'{"error":"file missing"}', "application/json")
            return

        FIRMAS.mkdir(parents=True, exist_ok=True)
        filename = _safe_name(blob["filename"])
        if not filename.lower().endswith(".png"):
            filename += ".png"
        dest = FIRMAS / filename
        if dest.exists():
            stem = dest.stem
            dest = FIRMAS / f"{stem}_{datetime.now().strftime('%H%M%S')}.png"
        dest.write_bytes(blob["data"])

        meta = {
            "saved_as": dest.name,
            "path": str(dest),
            "cliente": fields.get("cliente", ""),
            "maquina": fields.get("maquina", ""),
            "fecha": fields.get("fecha", ""),
            "nombre": fields.get("nombre", ""),
            "cargo": fields.get("cargo", ""),
            "ot": fields.get("ot", ""),
            "saved_at": datetime.now().isoformat(timespec="seconds"),
        }
        (FIRMAS / (dest.stem + ".json")).write_text(
            json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        self._send(
            200,
            json.dumps(meta, ensure_ascii=False).encode("utf-8"),
            "application/json; charset=utf-8",
        )


def _parse_multipart(raw: bytes, boundary: bytes):
    fields: dict[str, str] = {}
    files: dict[str, dict] = {}
    parts = raw.split(b"--" + boundary)
    for part in parts:
        if not part or part in (b"--\r\n", b"--", b"\r\n"):
            continue
        if part.startswith(b"\r\n"):
            part = part[2:]
        if part.endswith(b"\r\n"):
            part = part[:-2]
        if part.endswith(b"--"):
            part = part[:-2]
        header_blob, _, body = part.partition(b"\r\n\r\n")
        headers = header_blob.decode("utf-8", "ignore")
        if body.endswith(b"\r\n"):
            body = body[:-2]
        name_m = re.search(r'name="([^"]+)"', headers)
        if not name_m:
            continue
        name = name_m.group(1)
        file_m = re.search(r'filename="([^"]*)"', headers)
        if file_m:
            files[name] = {"filename": file_m.group(1), "data": body}
        else:
            fields[name] = body.decode("utf-8", "ignore")
    return fields, files


def main() -> int:
    import sys

    if sys.platform == "win32":
        for stream in (sys.stdout, sys.stderr):
            try:
                stream.reconfigure(encoding="utf-8", line_buffering=True)
            except Exception:
                pass

    if not HTML.is_file():
        print(f"[ERROR] Falta {HTML}")
        return 1
    FIRMAS.mkdir(parents=True, exist_ok=True)
    ip = _lan_ip()
    server = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    print("=" * 60)
    print("Firma WES - servidor local gratuito")
    print(f"En el PC:     http://127.0.0.1:{PORT}")
    print(f"En el telefono (misma WiFi): http://{ip}:{PORT}")
    print(f"Firmas -> {FIRMAS}")
    print("Ctrl+C para detener")
    print("=" * 60)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nDetenido.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
