"""
Respaldo diario de la cuenta WES y control de horas repetidas.

Cada pasada baja el detalle hora por hora de los puntos de la cuenta,
marca si una hora quedó guardada más de una vez y actualiza el Excel
de respaldo en reports/Respaldo_WES/.

Uso:
  python respaldo_diario_wes.py --ayer
  python respaldo_diario_wes.py --desde 2025-01-01 --hasta 2026-10-22 --subir

Credenciales (en este orden):
  1. WES_APP_USER y WES_APP_PASSWORD
  2. WES_APP_LOGIN_FILE (JSON con username y password)
  3. /cursor/stores/self/wes_app_login.json

La hora del día es la que muestra la app (la hora del TIME, sin pasar a Chile).
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import socket
import sqlite3
import time
import urllib.error
import urllib.request
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

CHILE = ZoneInfo("America/Santiago")
API = "https://api-rest.wes.cl"
BFF = f"{API}/q"
LOGIN_URL = f"{API}/wes/api/wes-mc-api/v1/login"
NODE_API = f"{API}/wes/api/acl-node/v1"
ENTITIES = f"{API}/wes/api/acl-entities/v1"

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "reports" / "Respaldo_WES"
DB_PATH = OUT_DIR / "respaldo_wes.sqlite"
XLSX_PATH = OUT_DIR / "Respaldo_general_WES.xlsx"
LOGIN_CANDIDATES = [
    Path(os.environ.get("WES_APP_LOGIN_FILE", "")),
    Path("/cursor/stores/self/wes_app_login.json"),
]

NAVY = "1B3A4B"
TEAL = "0F6E6E"
RED = "9C2F2F"
WHITE = "FFFFFF"
ZEBRA = "F4F7F8"
AMBER = "FFF4E5"

THIN = Border(
    left=Side(style="thin", color="D5DDE2"),
    right=Side(style="thin", color="D5DDE2"),
    top=Side(style="thin", color="D5DDE2"),
    bottom=Side(style="thin", color="D5DDE2"),
)


def chile_today() -> date:
    return datetime.now(CHILE).date()


def credenciales() -> tuple[str, str]:
    user = os.environ.get("WES_APP_USER", "").strip()
    password = os.environ.get("WES_APP_PASSWORD", "").strip()
    if user and password:
        return user, password
    for path in LOGIN_CANDIDATES:
        if path and path.is_file():
            data = json.loads(path.read_text(encoding="utf-8"))
            user = str(data.get("username") or "").strip()
            password = str(data.get("password") or "").strip()
            if user and password:
                return user, password
    raise SystemExit(
        "Faltan las credenciales de la app WES. Definí WES_APP_USER y "
        "WES_APP_PASSWORD en Cursor → Cloud Agents → Secrets."
    )


class Wes:
    def __init__(self, username: str, password: str) -> None:
        self.username = username
        self.password = password
        self.token = ""
        self.login()

    def login(self) -> None:
        code, body = self._raw(
            "POST",
            LOGIN_URL,
            {"username": self.username, "password": self.password},
            token=None,
        )
        if code != 200:
            raise SystemExit(f"No se pudo entrar a la app WES (HTTP {code}).")
        self.token = body.decode("utf-8", "replace").strip().strip('"')
        if self.token.count(".") != 2:
            raise SystemExit("La app no devolvió una sesión válida.")

    def _raw(self, method: str, url: str, body, token, timeout: int = 60):
        payload = {"m": method, "u": url, "b": body}
        raw = base64.b64encode(json.dumps(payload).encode()).decode()
        headers = {"Content-Type": "text/plain", "User-Agent": "wes-respaldo"}
        if token:
            headers["authorization"] = token
        req = urllib.request.Request(BFF, data=raw.encode(), method="POST", headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.status, resp.read()
        except urllib.error.HTTPError as exc:
            return exc.code, exc.read()
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            return 0, str(exc).encode()

    def get(self, url: str, timeout: int = 35, retries: int = 3) -> tuple[int, bytes]:
        last = (0, b"")
        for attempt in range(retries):
            code, body = self._raw("GET", url, None, self.token, timeout=timeout)
            if code == 401 and attempt == 0:
                self.login()
                continue
            if code == 200:
                return code, body
            last = (code, body)
            time.sleep(0.4 * (attempt + 1))
        return last

    def json(self, url: str):
        code, body = self.get(url)
        if code != 200:
            raise RuntimeError(f"HTTP {code} {url}")
        return json.loads(body.decode("utf-8", "replace"))


def user_id(token: str) -> str:
    mid = token.split(".")[1]
    mid += "=" * ((4 - len(mid) % 4) % 4)
    claims = json.loads(base64.urlsafe_b64decode(mid))
    return str(claims["sub"])


def cargar_nodos(api: Wes) -> list[dict]:
    uid = user_id(api.token)
    info = api.json(f"{API}/wes/api/wes-mc-api/v1/users/{uid}/information")
    nodes = info.get("allowedNodes") or []
    empresas: dict[str, str] = {}
    salida = []
    for node in nodes:
        nid = str(node.get("nodeId") or "").strip()
        if not nid:
            continue
        pref = nid.split("-", 1)[0]
        if pref not in empresas:
            try:
                comp = api.json(f"{ENTITIES}/companies/{pref}")
                empresas[pref] = str(comp.get("name") or "").strip()
            except Exception:
                empresas[pref] = ""
        if pref == "000026" and not empresas[pref]:
            empresas[pref] = "UDD"
        salida.append(
            {
                "node_id": nid,
                "name": str(node.get("name") or "").strip(),
                "company_id": pref,
                "company_name": empresas[pref] or pref,
            }
        )
    salida.sort(key=lambda n: (n["company_id"], _nro(n["node_id"]), n["node_id"]))
    return salida


def _nro(node_id: str) -> int:
    suf = node_id.partition("-")[2]
    return int(suf) if suf.isdigit() else 10**6


def abrir_db() -> sqlite3.Connection:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA synchronous=NORMAL")
    con.executescript(
        """
        CREATE TABLE IF NOT EXISTS nodos (
            node_id TEXT PRIMARY KEY,
            name TEXT,
            company_id TEXT,
            company_name TEXT
        );
        CREATE TABLE IF NOT EXISTS dias (
            node_id TEXT,
            day TEXT,
            rows INTEGER,
            unique_h INTEGER,
            ident INTEGER,
            differ INTEGER,
            max_copias INTEGER,
            sum_all REAL,
            sum_real REAL,
            is_dup INTEGER,
            error TEXT,
            PRIMARY KEY (node_id, day)
        );
        CREATE TABLE IF NOT EXISTS horas (
            node_id TEXT,
            day TEXT,
            time_iso TEXT,
            hour TEXT,
            value_one REAL,
            copies INTEGER,
            readings TEXT,
            sum_values REAL,
            identical INTEGER,
            PRIMARY KEY (node_id, time_iso)
        );
        CREATE TABLE IF NOT EXISTS pasadas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cuando TEXT,
            desde TEXT,
            hasta TEXT,
            nodos INTEGER,
            dias_nuevos INTEGER,
            dias_dup INTEGER
        );
        """
    )
    return con


def analizar_csv(text: str) -> tuple[dict, list[dict]]:
    groups: dict[str, list[float]] = defaultdict(list)
    rows = 0
    for line in text.splitlines()[1:]:
        line = line.strip()
        if not line or "," not in line:
            continue
        time_str, value_str = line.split(",", 1)
        try:
            value = float(value_str.strip().replace(" ", "").replace(",", "."))
        except ValueError:
            continue
        groups[time_str.strip()].append(value)
        rows += 1

    ident = differ = 0
    max_copias = 0
    sum_all = 0.0
    sum_real = 0.0
    horas = []
    for time_iso in sorted(groups):
        vals = groups[time_iso]
        max_copias = max(max_copias, len(vals))
        suma = float(sum(vals))
        sum_all += suma
        identical = len(vals) > 1 and all(abs(v - vals[0]) < 1e-6 for v in vals)
        if len(vals) > 1 and identical:
            ident += 1
            uno = vals[-1]
            sum_real += uno
        elif len(vals) > 1:
            differ += 1
            uno = vals[-1]
            sum_real += suma
        else:
            uno = vals[0]
            sum_real += uno
        hour = time_iso[11:16] if len(time_iso) >= 16 else ""
        horas.append(
            {
                "time_iso": time_iso,
                "hour": hour,
                "value_one": uno,
                "copies": len(vals),
                "readings": " | ".join(f"{v:.2f}" for v in vals),
                "sum_values": suma,
                "identical": 1 if identical else 0,
            }
        )
    unique = len(groups)
    is_dup = unique > 0 and ident > 0
    dia = {
        "rows": rows,
        "unique_h": unique,
        "ident": ident,
        "differ": differ,
        "max_copias": max_copias,
        "sum_all": round(sum_all, 4),
        "sum_real": round(sum_real, 4),
        "is_dup": 1 if is_dup else 0,
        "error": "",
    }
    return dia, horas


def ya_guardado(con: sqlite3.Connection, node_id: str, day: str) -> bool:
    row = con.execute(
        "SELECT error FROM dias WHERE node_id=? AND day=?",
        (node_id, day),
    ).fetchone()
    return row is not None and not row[0]


def guardar(con: sqlite3.Connection, node_id: str, day: str, dia: dict, horas: list[dict]) -> None:
    con.execute("DELETE FROM horas WHERE node_id=? AND day=?", (node_id, day))
    con.execute(
        """
        INSERT INTO dias (node_id, day, rows, unique_h, ident, differ, max_copias,
                          sum_all, sum_real, is_dup, error)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(node_id, day) DO UPDATE SET
            rows=excluded.rows, unique_h=excluded.unique_h, ident=excluded.ident,
            differ=excluded.differ, max_copias=excluded.max_copias,
            sum_all=excluded.sum_all, sum_real=excluded.sum_real,
            is_dup=excluded.is_dup, error=excluded.error
        """,
        (
            node_id,
            day,
            dia["rows"],
            dia["unique_h"],
            dia["ident"],
            dia["differ"],
            dia["max_copias"],
            dia["sum_all"],
            dia["sum_real"],
            dia["is_dup"],
            dia["error"],
        ),
    )
    con.executemany(
        """
        INSERT INTO horas (node_id, day, time_iso, hour, value_one, copies,
                           readings, sum_values, identical)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(node_id, time_iso) DO UPDATE SET
            day=excluded.day, hour=excluded.hour, value_one=excluded.value_one,
            copies=excluded.copies, readings=excluded.readings,
            sum_values=excluded.sum_values, identical=excluded.identical
        """,
        [
            (
                node_id,
                day,
                h["time_iso"],
                h["hour"],
                h["value_one"],
                h["copies"],
                h["readings"],
                h["sum_values"],
                h["identical"],
            )
            for h in horas
        ],
    )


def bajar_dia(api: Wes, node_id: str, day: date) -> tuple[dict, list[dict]]:
    stamp = day.strftime("%d%m%Y")
    url = (
        f"{NODE_API}/nodes/{node_id}/dates.measures.csv"
        f"?start={stamp}&end={stamp}"
    )
    code, body = api.get(url)
    if code != 200:
        return {
            "rows": 0,
            "unique_h": 0,
            "ident": 0,
            "differ": 0,
            "max_copias": 0,
            "sum_all": 0.0,
            "sum_real": 0.0,
            "is_dup": 0,
            "error": f"HTTP {code}",
        }, []
    return analizar_csv(body.decode("utf-8", "replace"))


def descargar(api: Wes, nodos: list[dict], desde: date, hasta: date, workers: int) -> int:
    con = abrir_db()
    con.executemany(
        """
        INSERT INTO nodos (node_id, name, company_id, company_name)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(node_id) DO UPDATE SET
            name=excluded.name, company_id=excluded.company_id,
            company_name=excluded.company_name
        """,
        [(n["node_id"], n["name"], n["company_id"], n["company_name"]) for n in nodos],
    )
    con.commit()

    jobs = []
    day = desde
    while day <= hasta:
        iso = day.isoformat()
        for nodo in nodos:
            if not ya_guardado(con, nodo["node_id"], iso):
                jobs.append((nodo["node_id"], day))
        day += timedelta(days=1)
    print(f"Por bajar: {len(jobs)} lecturas ({desde} a {hasta}, {len(nodos)} puntos).", flush=True)
    if not jobs:
        con.close()
        return 0

    nuevos = 0
    errores = 0
    t0 = time.time()
    lote = []

    def trabajo(job):
        nid, dia = job
        rec, horas = bajar_dia(api, nid, dia)
        return nid, dia.isoformat(), rec, horas

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futuros = [pool.submit(trabajo, job) for job in jobs]
        for i, fut in enumerate(as_completed(futuros), 1):
            nid, iso, rec, horas = fut.result()
            lote.append((nid, iso, rec, horas))
            if rec["error"]:
                errores += 1
            else:
                nuevos += 1
            if len(lote) >= 150:
                for item in lote:
                    guardar(con, *item)
                con.commit()
                lote.clear()
            if i % 400 == 0 or i == len(jobs):
                print(
                    f"  {i}/{len(jobs)}  ok={nuevos} err={errores}  {time.time()-t0:.0f}s",
                    flush=True,
                )
    for item in lote:
        guardar(con, *item)
    con.commit()
    con.close()
    return nuevos


def _fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def _font(bold: bool = False, color: str = "1B3A4B", size: int = 11) -> Font:
    return Font(name="Calibri", bold=bold, color=color, size=size)


def _header(ws, row: int, headers: list[str], fill: str) -> None:
    for col, text in enumerate(headers, 1):
        cell = ws.cell(row, col, text)
        cell.fill = _fill(fill)
        cell.font = _font(True, WHITE, 11)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN
    ws.row_dimensions[row].height = 30
    ws.freeze_panes = f"A{row + 1}"


def _widths(ws, widths: list[int]) -> None:
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def exportar_excel(con: sqlite3.Connection) -> list[Path]:
    """Arma el Excel general. Las horas van en el mismo libro, partidas por año."""
    nodos = {
        r[0]: r
        for r in con.execute(
            "SELECT node_id, name, company_id, company_name FROM nodos"
        )
    }
    dias = list(
        con.execute(
            """
            SELECT d.node_id, d.day, d.unique_h, d.ident, d.differ, d.max_copias,
                   d.sum_all, d.sum_real, d.is_dup, d.error
            FROM dias d
            WHERE d.error IS NULL OR d.error = ''
            ORDER BY substr(d.node_id, 1, 6),
                     CAST(substr(d.node_id, 8) AS INTEGER),
                     d.node_id, d.day
            """
        )
    )
    n_horas = con.execute("SELECT COUNT(*) FROM horas").fetchone()[0]
    dup_dias = sum(1 for d in dias if d[8])
    cuando = datetime.now(CHILE).strftime("%d-%m-%Y %H:%M")

    wb = Workbook()
    portada = wb.active
    portada.title = "Respaldo"
    portada.sheet_properties.tabColor = NAVY
    portada.page_setup.orientation = "landscape"
    portada.page_setup.fitToPage = True
    portada.page_setup.fitToWidth = 1
    portada.page_setup.fitToHeight = 1
    portada.page_setup.paperSize = portada.PAPERSIZE_A4
    portada.sheet_view.showGridLines = False
    portada.sheet_view.view = "pageBreakPreview"
    portada.oddHeader.left.text = "WES  ·  Respaldo de la cuenta"
    portada.oddFooter.left.text = "Confidencial"
    portada.oddFooter.right.text = "Página &P de &N"

    portada["A1"] = "Respaldo general de la cuenta WES"
    portada["A1"].font = _font(True, NAVY, 22)
    portada.merge_cells("A1:B1")
    portada["A2"] = (
        "Copia de las horas de cada punto. Se actualiza todos los días a las 00:01, "
        "hora de Chile, con el día que acaba de cerrar."
    )
    portada["A2"].font = _font(False, "3D4C54", 12)
    portada.merge_cells("A2:B2")
    portada.row_dimensions[2].height = 32
    bloques = [
        ("Última pasada", cuando + " hora Chile"),
        ("Puntos", f"{len(nodos)}"),
        ("Días guardados", f"{len(dias)}"),
        ("Horas guardadas", f"{n_horas}"),
        ("Días con hora repetida", f"{dup_dias}"),
    ]
    for i, (etiqueta, valor) in enumerate(bloques):
        fila = 4 + i
        portada.cell(fila, 1, etiqueta).font = _font(True, TEAL, 12)
        portada.cell(fila, 2, valor).font = _font(False, NAVY, 12)
    portada["A10"] = (
        "En Diario está el total de cada punto y cada día. "
        "m³ en la app suma todas las lecturas de la hora. "
        "m³ una copia cuenta esa hora una sola vez cuando las lecturas son iguales. "
        "Repetido = Sí cuando alguna hora quedó guardada más de una vez con el mismo valor. "
        "La hoja Duplicados lista solo esos días. "
        "Cada hora queda en Respaldo_horas_AAAA.xlsx, en la misma carpeta."
    )
    portada["A10"].alignment = Alignment(wrap_text=True, vertical="top")
    portada["A10"].font = _font(False, "3D4C54", 11)
    portada.merge_cells("A10:B12")
    portada.row_dimensions[10].height = 22
    portada.row_dimensions[11].height = 22
    portada.row_dimensions[12].height = 22
    _widths(portada, [32, 78])
    portada.sheet_view.zoomScale = 120
    portada.print_title_rows = "1:1"
    portada.page_setup.horizontalCentered = True
    portada.sheet_properties.pageSetUpPr.fitToPage = True

    diario = wb.create_sheet("Diario")
    diario.sheet_properties.tabColor = TEAL
    headers_d = [
        "Empresa",
        "ID nodo",
        "Nombre",
        "Fecha",
        "Horas",
        "Horas repetidas",
        "Máx. copias",
        "m³ en la app",
        "m³ una copia",
        "Diferencia",
        "Repetido",
    ]
    _header(diario, 1, headers_d, TEAL)
    dup_sheet = wb.create_sheet("Duplicados")
    dup_sheet.sheet_properties.tabColor = RED
    _header(dup_sheet, 1, headers_d, RED)

    def pintar_fila(ws, fila, valores, dup: bool, zebra: bool) -> None:
        for col, valor in enumerate(valores, 1):
            cell = ws.cell(fila, col, valor)
            cell.font = _font(False, RED if dup else "1B3A4B", 11)
            cell.border = THIN
            cell.alignment = Alignment(vertical="center", horizontal="center" if col >= 4 else "left")
            if dup:
                cell.fill = _fill("FDECEC")
            elif zebra:
                cell.fill = _fill(ZEBRA)
            if col in (8, 9, 10) and isinstance(valor, float):
                cell.number_format = '#,##0.00'
            if col == 4 and isinstance(valor, datetime):
                cell.number_format = "DD-MM-YYYY"

    n_dup_rows = 1
    for i, rec in enumerate(dias):
        nid, day, unique_h, ident, differ, max_c, sum_all, sum_real, is_dup, _err = rec
        meta = nodos.get(nid, (nid, "", "", ""))
        diff = round((sum_all or 0) - (sum_real or 0), 4)
        try:
            fecha = datetime.strptime(day, "%Y-%m-%d")
        except ValueError:
            fecha = day
        valores = [
            meta[3],
            nid,
            meta[1],
            fecha,
            unique_h,
            ident,
            max_c,
            float(sum_all or 0),
            float(sum_real or 0),
            diff,
            "Sí" if is_dup else "No",
        ]
        pintar_fila(diario, i + 2, valores, bool(is_dup), i % 2 == 1)
        if is_dup:
            n_dup_rows += 1
            pintar_fila(dup_sheet, n_dup_rows, valores, True, False)

    if len(dias) == 0:
        diario.cell(2, 1, "Todavía no hay días guardados.")
    if n_dup_rows == 1:
        dup_sheet.cell(2, 1, "Ningún día tiene la hora repetida.")
        dup_sheet.cell(2, 1).font = _font(True, TEAL, 12)

    for ws in (diario, dup_sheet):
        last = ws.max_row
        ws.auto_filter.ref = f"A1:K{last}"
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:K{max(last, 2)}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.horizontalCentered = True
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = "1:1"
        ws.page_setup.scale = 70
        ws.oddHeader.left.text = "WES  ·  Respaldo"
        ws.oddFooter.right.text = "Página &P de &N"
        ws.sheet_view.view = "pageBreakPreview"
        ws.sheet_view.zoomScale = 110
        _widths(ws, [22, 14, 38, 14, 10, 16, 14, 16, 16, 14, 12])

    archivos = [XLSX_PATH]
    hora_headers = [
        "Fecha",
        "Hora",
        "Empresa",
        "ID nodo",
        "Nombre",
        "m³ de una copia",
        "Copias",
        "Lecturas",
        "m³ sumando copias",
    ]
    years = [
        r[0]
        for r in con.execute("SELECT DISTINCT substr(day, 1, 4) FROM horas ORDER BY 1")
    ]
    if not years:
        ws = wb.create_sheet("Horas")
        _header(ws, 1, hora_headers, NAVY)
        ws.cell(2, 1, "Todavía no hay horas guardadas.")
    else:
        portada["A14"] = "Horas de cada año"
        portada["A14"].font = _font(True, TEAL, 12)
        portada["A15"] = "Respaldo_horas_AAAA.xlsx"
        portada["A15"].font = _font(False, NAVY, 12)
        portada["B15"] = "Misma carpeta. Una fila por hora de cada punto. Si el año no cabe en una hoja, se parte en ene-jun y jul-dic."
        portada["B15"].font = _font(False, "3D4C54", 11)
        portada["B15"].alignment = Alignment(wrap_text=True, vertical="center")
        portada.row_dimensions[15].height = 32
        for year in years:
            archivos.extend(exportar_horas_anio(con, nodos, year, hora_headers))

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(XLSX_PATH)
    return archivos


def exportar_horas_anio(con: sqlite3.Connection, nodos: dict, year: str, headers: list[str]) -> list[Path]:
    """Escribe las horas de un año. Parte el archivo si un tramo supera el límite de Excel."""
    import xlsxwriter

    def escribir(path: Path, inicio: str, fin: str) -> int:
        libro = xlsxwriter.Workbook(str(path), {"constant_memory": True, "strings_to_urls": False})
        ws = libro.add_worksheet("Horas")
        head = libro.add_format({"bold": True, "font_color": "FFFFFF", "bg_color": NAVY, "font_name": "Calibri", "align": "center"})
        normal = libro.add_format({"font_name": "Calibri", "font_color": "1B3A4B", "num_format": "dd-mm-yyyy"})
        num = libro.add_format({"font_name": "Calibri", "font_color": "1B3A4B", "num_format": "#,##0.00"})
        texto = libro.add_format({"font_name": "Calibri", "font_color": "1B3A4B"})
        rojo = libro.add_format({"font_name": "Calibri", "font_color": RED, "bg_color": "FDECEC"})
        rojo_num = libro.add_format({"font_name": "Calibri", "font_color": RED, "bg_color": "FDECEC", "num_format": "#,##0.00"})
        rojo_fecha = libro.add_format({"font_name": "Calibri", "font_color": RED, "bg_color": "FDECEC", "num_format": "dd-mm-yyyy"})
        for col, text in enumerate(headers):
            ws.write(0, col, text, head)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, 0, len(headers) - 1)
        anchos = [14, 10, 22, 14, 38, 18, 12, 28, 20]
        for i, w in enumerate(anchos):
            ws.set_column(i, i, w)
        n = 0
        sql = """
            SELECT node_id, day, hour, value_one, copies, readings, sum_values, identical
            FROM horas
            WHERE day >= ? AND day <= ?
            ORDER BY day, node_id, time_iso
        """
        for nid, day, hour, value_one, copies, readings, sum_values, identical in con.execute(sql, (inicio, fin)):
            n += 1
            if n > 1_000_000:
                break
            meta = nodos.get(nid, (nid, "", "", ""))
            repetida = bool(identical) and (copies or 0) > 1
            f_fecha = rojo_fecha if repetida else normal
            f_txt = rojo if repetida else texto
            f_num = rojo_num if repetida else num
            try:
                fecha = datetime.strptime(day, "%Y-%m-%d")
            except ValueError:
                fecha = day
            ws.write_datetime(n, 0, fecha, f_fecha)
            ws.write(n, 1, hour or "", f_txt)
            ws.write(n, 2, meta[3], f_txt)
            ws.write(n, 3, nid, f_txt)
            ws.write(n, 4, meta[1], f_txt)
            ws.write_number(n, 5, float(value_one or 0), f_num)
            ws.write_number(n, 6, int(copies or 0), f_txt)
            ws.write(n, 7, readings if copies and copies > 1 else "", f_txt)
            ws.write_number(n, 8, float(sum_values or 0), f_num)
        ws.autofilter(0, 0, max(n, 1), len(headers) - 1)
        libro.close()
        return n

    total = con.execute(
        "SELECT COUNT(*) FROM horas WHERE day >= ? AND day <= ?",
        (f"{year}-01-01", f"{year}-12-31"),
    ).fetchone()[0]
    if total <= 1_000_000:
        path = OUT_DIR / f"Respaldo_horas_{year}.xlsx"
        escribir(path, f"{year}-01-01", f"{year}-12-31")
        return [path]
    path_a = OUT_DIR / f"Respaldo_horas_{year}_ene-jun.xlsx"
    path_b = OUT_DIR / f"Respaldo_horas_{year}_jul-dic.xlsx"
    escribir(path_a, f"{year}-01-01", f"{year}-06-30")
    escribir(path_b, f"{year}-07-01", f"{year}-12-31")
    return [path_a, path_b]


def subir(paths: list[Path]) -> list[str]:
    from wes_google_drive import subir_a_drive

    links = []
    for path in paths:
        info = subir_a_drive(path, subcarpeta="Respaldo_WES")
        links.append(info.get("web_view_link") or "")
        print(f"Drive: {info.get('name')} {info.get('web_view_link')}", flush=True)
    return links


def resumen(con: sqlite3.Connection, desde: date, hasta: date) -> dict:
    row = con.execute(
        """
        SELECT COUNT(*),
               SUM(CASE WHEN is_dup=1 THEN 1 ELSE 0 END),
               SUM(CASE WHEN error IS NOT NULL AND error != '' THEN 1 ELSE 0 END)
        FROM dias
        WHERE day >= ? AND day <= ?
        """,
        (desde.isoformat(), hasta.isoformat()),
    ).fetchone()
    dups = list(
        con.execute(
            """
            SELECT d.day, n.company_name, d.node_id, n.name, d.ident, d.unique_h,
                   d.sum_all, d.sum_real
            FROM dias d
            JOIN nodos n ON n.node_id = d.node_id
            WHERE d.day >= ? AND d.day <= ? AND d.is_dup = 1
            ORDER BY d.day, d.node_id
            """,
            (desde.isoformat(), hasta.isoformat()),
        )
    )
    return {
        "dias": row[0] or 0,
        "dup": row[1] or 0,
        "errores": row[2] or 0,
        "detalle": dups,
    }


def parse_dia(texto: str) -> date:
    return datetime.strptime(texto, "%Y-%m-%d").date()


def main() -> int:
    parser = argparse.ArgumentParser(description="Respaldo diario WES y control de horas repetidas")
    parser.add_argument("--desde", help="AAAA-MM-DD")
    parser.add_argument("--hasta", help="AAAA-MM-DD")
    parser.add_argument("--ayer", action="store_true", help="Solo el día que acaba de cerrar en Chile")
    parser.add_argument("--workers", type=int, default=6)
    parser.add_argument("--subir", action="store_true", help="Subir el Excel a Google Drive")
    parser.add_argument("--solo-excel", action="store_true", help="Rearmar el Excel con lo ya guardado")
    args = parser.parse_args()

    socket.setdefaulttimeout(40)
    hoy = chile_today()
    if args.ayer:
        desde = hasta = hoy - timedelta(days=1)
    else:
        hasta = parse_dia(args.hasta) if args.hasta else hoy - timedelta(days=1)
        desde = parse_dia(args.desde) if args.desde else hasta
    if desde > hasta:
        raise SystemExit("--desde no puede ser posterior a --hasta")

    con = abrir_db()
    nuevos = 0
    if not args.solo_excel:
        user, password = credenciales()
        api = Wes(user, password)
        nodos = cargar_nodos(api)
        print(f"Cuenta: {len(nodos)} puntos. Tramo {desde} a {hasta}.", flush=True)
        nuevos = descargar(api, nodos, desde, hasta, args.workers)
        con.close()
        con = abrir_db()

    paths = exportar_excel(con)
    info = resumen(con, desde, hasta)
    con.execute(
        "INSERT INTO pasadas (cuando, desde, hasta, nodos, dias_nuevos, dias_dup) VALUES (?, ?, ?, ?, ?, ?)",
        (
            datetime.now(CHILE).isoformat(timespec="seconds"),
            desde.isoformat(),
            hasta.isoformat(),
            con.execute("SELECT COUNT(*) FROM nodos").fetchone()[0],
            nuevos,
            info["dup"],
        ),
    )
    con.commit()
    print(
        f"Listo. Días en el tramo: {info['dias']}. "
        f"Con hora repetida: {info['dup']}. Errores: {info['errores']}.",
        flush=True,
    )
    for rec in info["detalle"][:40]:
        print(
            f"  REPETIDO {rec[0]} {rec[2]} {rec[3]}  horas {rec[4]}/{rec[5]}  "
            f"app {rec[6]}  una copia {rec[7]}",
            flush=True,
        )
    print(f"Excel: {XLSX_PATH}", flush=True)
    if args.subir:
        subir(paths)
    con.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
