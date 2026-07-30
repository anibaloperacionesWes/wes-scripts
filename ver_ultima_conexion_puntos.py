"""
Última medida recibida por nodo (proxy de última conexión / envío de datos).

Consulta dates.measures.csv (API acl-node, puerto 7003) y toma la marca TIME más reciente
de los últimos N días. Si el servidor de medidas está caído, lo reporta.

Uso:
  python ver_ultima_conexion_puntos.py
  python ver_ultima_conexion_puntos.py --solo-control-nocturno
  python ver_ultima_conexion_puntos.py --dias 14 --csv salida.csv
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests

from control_nocturno import cargar_targets_desde_excel, default_excel_path
from generar_reporte_word import NODE_NAMES, acl_node_base_url

ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"

try:
    from zoneinfo import ZoneInfo

    CHILE_TZ = ZoneInfo("America/Santiago")
except Exception:
    CHILE_TZ = timezone(timedelta(hours=-4))


def _api_medidas_viva() -> Tuple[bool, str]:
    url = f"{acl_node_base_url()}/nodes/000008-01"
    try:
        r = requests.get(url, timeout=8)
        return True, f"HTTP {r.status_code}"
    except requests.RequestException as e:
        return False, str(e)


def _parse_time_utc(s: str) -> Optional[datetime]:
    try:
        return datetime.fromisoformat(s.strip().replace("Z", "+00:00"))
    except (ValueError, TypeError):
        return None


def _ultima_marca_csv(node_id: str, dias: int) -> Tuple[Optional[datetime], str]:
    """Retorna (última marca UTC en CSV, estado)."""
    hoy = datetime.now(timezone.utc).date()
    url = f"{acl_node_base_url()}/nodes/{node_id}/dates.measures.csv"
    ultima: Optional[datetime] = None
    errores = 0
    for d in range(dias):
        dia = hoy - timedelta(days=d)
        ds = dia.strftime("%d%m%Y")
        try:
            r = requests.get(url, params=[("start", ds), ("end", ds)], timeout=25)
            if r.status_code != 200:
                errores += 1
                continue
            for line in r.text.strip().split("\n")[1:]:
                if not line.strip() or "," not in line:
                    continue
                t = line.split(",", 1)[0].strip()
                dt = _parse_time_utc(t)
                if dt and (ultima is None or dt > ultima):
                    ultima = dt
        except requests.RequestException:
            errores += 1
    if ultima:
        return ultima, "OK"
    if errores >= dias:
        return None, "sin respuesta API"
    return None, "sin datos en ventana"


def _listar_nodos_todos() -> List[Dict[str, str]]:
    from reporte_puntos_en_cero import obtener_todos_los_nodos

    return obtener_todos_los_nodos()


def _listar_nodos_control_nocturno() -> List[Dict[str, str]]:
    targets = cargar_targets_desde_excel(default_excel_path())
    out: List[Dict[str, str]] = []
    for nid, meta in targets.items():
        out.append(
            {
                "nodeId": nid,
                "nodeName": str(meta.get("nodeName") or NODE_NAMES.get(nid, nid)),
                "companyName": str(meta.get("cliente", "")),
            }
        )
    return out


def _fmt_chile(dt: Optional[datetime]) -> str:
    if dt is None:
        return "—"
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    ch = dt.astimezone(CHILE_TZ)
    return ch.strftime("%Y-%m-%d %H:%M:%S")


def _horas_desde(dt: Optional[datetime]) -> Optional[float]:
    if dt is None:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return (datetime.now(timezone.utc) - dt).total_seconds() / 3600.0


def main() -> int:
    parser = argparse.ArgumentParser(description="Última medida por nodo (API WES)")
    parser.add_argument(
        "--solo-control-nocturno",
        action="store_true",
        help="Solo nodos del Excel de control nocturno",
    )
    parser.add_argument("--dias", type=int, default=7, help="Días hacia atrás a revisar")
    parser.add_argument("--workers", type=int, default=12)
    parser.add_argument("--csv", default="", help="Ruta CSV de salida opcional")
    args = parser.parse_args()

    viva, det = _api_medidas_viva()
    print("=" * 72)
    print("ESTADO API MEDIDAS (acl-node, puerto 7003)")
    print("=" * 72)
    print(f"URL base: {acl_node_base_url()}")
    if viva:
        print(f"Estado: EN LÍNEA ({det})")
    else:
        print(f"Estado: CAÍDA o inaccesible")
        print(f"Detalle: {det}")
        print()
        print(
            "Si la app WES no carga gráficos ni consumos, es coherente: la app usa este servicio.\n"
            "La API de configuración (puerto 7001) puede seguir respondiendo sin mostrar medidas."
        )
        # entities check
        try:
            r = requests.get(f"{ENTITY_BASE_URL}/configuration/companies", timeout=8)
            print(f"API configuración (7001): HTTP {r.status_code}")
        except requests.RequestException as e:
            print(f"API configuración (7001): error — {e}")
        return 2

    nodos = (
        _listar_nodos_control_nocturno()
        if args.solo_control_nocturno
        else _listar_nodos_todos()
    )
    print(f"\nNodos a revisar: {len(nodos)} (últimos {args.dias} días)\n")

    filas: List[Dict[str, Any]] = []

    def _one(n: Dict[str, str]) -> Dict[str, Any]:
        nid = n["nodeId"]
        ultima, estado = _ultima_marca_csv(nid, args.dias)
        horas = _horas_desde(ultima)
        return {
            "nodeId": nid,
            "nodeName": n.get("nodeName", nid),
            "companyName": n.get("companyName", ""),
            "ultima_utc": ultima.isoformat() if ultima else "",
            "ultima_chile": _fmt_chile(ultima),
            "hace_horas": round(horas, 1) if horas is not None else "",
            "estado": estado,
        }

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
        futs = {ex.submit(_one, n): n for n in nodos}
        done = 0
        for fut in as_completed(futs):
            done += 1
            if done % 25 == 0 or done == len(nodos):
                print(f"  Progreso: {done}/{len(nodos)}", flush=True)
            filas.append(fut.result())

    filas.sort(key=lambda r: (r.get("ultima_utc") or "", r["nodeId"]))

    sin_datos = [f for f in filas if not f.get("ultima_utc")]
    con_datos = [f for f in filas if f.get("ultima_utc")]

    print("\n" + "=" * 72)
    print("RESUMEN")
    print("=" * 72)
    print(f"Con medida en ventana: {len(con_datos)}")
    print(f"Sin medida / sin respuesta: {len(sin_datos)}")

    if con_datos:
        mas_viejo = min(con_datos, key=lambda f: f["ultima_utc"])
        mas_nuevo = max(con_datos, key=lambda f: f["ultima_utc"])
        print(
            f"Última medida más reciente (cualquier punto): {mas_nuevo['nodeName']} ({mas_nuevo['nodeId']}) "
            f"→ {mas_nuevo['ultima_chile']} Chile"
        )
        print(
            f"Última medida más antigua en ventana: {mas_viejo['nodeName']} ({mas_viejo['nodeId']}) "
            f"→ {mas_viejo['ultima_chile']} Chile"
        )

    print("\nPuntos sin datos en los últimos días (primeros 20):")
    for f in sin_datos[:20]:
        print(f"  {f['nodeId']} | {f['nodeName'][:40]} | {f['estado']}")
    if len(sin_datos) > 20:
        print(f"  ... y {len(sin_datos) - 20} más")

    if args.csv.strip():
        out = Path(args.csv)
        out.parent.mkdir(parents=True, exist_ok=True)
        with out.open("w", encoding="utf-8", newline="") as fh:
            w = csv.DictWriter(
                fh,
                fieldnames=[
                    "nodeId",
                    "nodeName",
                    "companyName",
                    "ultima_chile",
                    "hace_horas",
                    "estado",
                    "ultima_utc",
                ],
            )
            w.writeheader()
            for f in sorted(filas, key=lambda x: x.get("ultima_utc") or ""):
                w.writerow(f)
        print(f"\nCSV guardado: {out.resolve()}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
