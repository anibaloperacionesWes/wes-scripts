"""
Busca direcciones @linkes.cl y @parauco.com en archivos de texto (proyecto y carpetas extra).

La API WES solo tiene GET /users?email= y GET /users/{userId}; no hay listado por dominio.
Ver ALTERNATIVAS_LISTADO_USUARIOS_DOMINIO.txt (dos alternativas: directorio Google/Microsoft
y export desde equipo WES).

Uso:
  python buscar_emails_linkes_parauco.py
  python buscar_emails_linkes_parauco.py --extra-dir "C:\\ruta\\export_correos"
  python buscar_emails_linkes_parauco.py --verificar-api
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, Iterable, List, Set, Tuple

import requests

ENTITY_BASE = "http://104.248.53.141:7001/wes/api/acl-entities/v1"

DOMINIOS = ("@linkes.cl", "@parauco.com")
# Regex: usuario local + dominio
EMAIL_RE = re.compile(
    r"[a-zA-Z0-9._%+-]+@(?:linkes\.cl|parauco\.com)\b",
    re.IGNORECASE,
)

EXTENSIONES = {".log", ".txt", ".csv", ".json", ".md", ".py", ".eml", ".html"}
# Logs de monitoreo pueden superar 5 MB; leer por tramos si hace falta.
MAX_BYTES_DEFAULT = 5_000_000
MAX_BYTES_LOG = 80_000_000


def _iter_text_files(raiz: Path) -> Iterable[Path]:
    for p in raiz.rglob("*"):
        if not p.is_file():
            continue
        suf = p.suffix.lower()
        if suf not in EXTENSIONES:
            continue
        try:
            size = p.stat().st_size
        except OSError:
            continue
        lim = MAX_BYTES_LOG if suf == ".log" else MAX_BYTES_DEFAULT
        if size > lim:
            continue
        yield p


def extraer_emails_en_texto(texto: str) -> Set[str]:
    return {m.group(0).lower() for m in EMAIL_RE.finditer(texto)}


def _escanear_raiz(raiz: Path) -> Tuple[Set[str], Dict[str, Set[str]]]:
    todos: Set[str] = set()
    por_archivo: Dict[str, Set[str]] = {}
    try:
        raiz_res = raiz.resolve()
    except OSError:
        return todos, por_archivo
    for fp in _iter_text_files(raiz):
        try:
            raw = fp.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        found = extraer_emails_en_texto(raw)
        if not found:
            continue
        try:
            rel = str(fp.relative_to(raiz_res))
        except ValueError:
            rel = str(fp)
        clave = f"{raiz_res.name}/{rel}"
        for e in found:
            todos.add(e)
            por_archivo.setdefault(clave, set()).add(e)
    return todos, por_archivo


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extrae @linkes.cl y @parauco.com del workspace y opcionalmente verifica en API"
    )
    parser.add_argument(
        "--raiz",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="Carpeta raíz principal a escanear (default: carpeta de wes-scripts)",
    )
    parser.add_argument(
        "--extra-dir",
        type=Path,
        action="append",
        default=[],
        metavar="DIR",
        help="Carpeta adicional a escanear (puede repetirse). Ej.: export de buzón o OneDrive",
    )
    parser.add_argument(
        "--salida",
        type=Path,
        default=Path("reports")
        / "Parque_Arauco"
        / "Reportes_agregados_reunion_abril_2026"
        / "emails_linkes_parauco_encontrados.txt",
        help="Archivo de salida con un correo por línea",
    )
    parser.add_argument(
        "--verificar-api",
        action="store_true",
        help="Para cada correo, GET /users?email= e informa si existe usuario en WES",
    )
    args = parser.parse_args()

    raices: List[Path] = [args.raiz]
    for d in args.extra_dir:
        if d and d.exists():
            raices.append(d)
        elif d:
            print(f"[WARN] No existe carpeta extra (se omite): {d}")

    todos: Set[str] = set()
    por_archivo: Dict[str, Set[str]] = {}
    for raiz in raices:
        t, p = _escanear_raiz(raiz)
        todos |= t
        for k, v in p.items():
            por_archivo.setdefault(k, set()).update(v)

    args.salida.parent.mkdir(parents=True, exist_ok=True)
    with open(args.salida, "w", encoding="utf-8") as f:
        f.write("# Dominios: @linkes.cl y @parauco.com\n")
        f.write(f"# Carpetas escaneadas: {', '.join(str(r.resolve()) for r in raices)}\n\n")
        for e in sorted(todos):
            f.write(e + "\n")

    print(f"[OK] {len(todos)} correo(s) único(s) -> {args.salida.resolve()}\n")

    if por_archivo:
        print("Por archivo (muestra):")
        for rel in sorted(por_archivo.keys())[:25]:
            print(f"  {rel}: {', '.join(sorted(por_archivo[rel]))}")
        if len(por_archivo) > 25:
            print(f"  ... y {len(por_archivo) - 25} archivo(s) más con coincidencias.")

    if args.verificar_api:
        print("\nAPI /users?email= :")
        for email in sorted(todos):
            try:
                r = requests.get(
                    f"{ENTITY_BASE}/users",
                    params={"email": email},
                    timeout=20,
                )
            except requests.RequestException as ex:
                print(f"  {email}  ERROR {ex}")
                continue
            if r.status_code == 200:
                u = r.json()
                uid = u.get("userId", "?")
                name = f"{u.get('name', '')} {u.get('lastName', '')}".strip()
                n_nodes = len(u.get("allowedNodes") or [])
                print(f"  {email}  OK  userId={uid}  {name}  allowedNodes={n_nodes}")
            elif r.status_code == 404:
                print(f"  {email}  (sin usuario en API)")
            else:
                print(f"  {email}  HTTP {r.status_code}")

    if not todos:
        print(
            "\n[INFO] No se encontraron correos. Amplía EXTENSIONES o revisa otra carpeta con --raiz."
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
