"""
Alerta diaria matutina — Fundo Zapallar Etapa N°5 (000027-03).

Revisa caudales horarios recientes y marca lecturas sobre el techo hidráulico
de tubería DN90 (~60 m³/h), típico de error de memoria/placa.

Uso:
  python alerta_diaria_etapa5_zapallar.py
  python alerta_diaria_etapa5_zapallar.py --dias 2
  python alerta_diaria_etapa5_zapallar.py --json
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict, dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import List, Optional, Tuple
from zoneinfo import ZoneInfo

import requests

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "reports" / "Fundo_Zapallar" / "Alertas_Diarias"

try:
    from generar_reporte_word import acl_node_base_url, CHILE_TZ
except Exception:
    def acl_node_base_url() -> str:
        return "http://104.248.53.141:7003/wes/api/acl-node/v1"

    CHILE_TZ = ZoneInfo("America/Santiago")

COMPANY = "Fundo Zapallar"
COMPANY_ID = "000027"
NODE_ID = "000027-03"
NODE_NAME = "Etapa N°5"
CAUDAL_MAX_DN90_M3H = 60.0
CAUDAL_AVISO_M3H = 45.0  # elevado pero bajo el techo duro


@dataclass
class Pico:
    fecha_hora_chile: str
    m3h: float
    nivel: str  # ALERTA | AVISO


@dataclass
class ResultadoAlerta:
    ok: bool
    node_id: str
    node_name: str
    umbral_alerta_m3h: float
    umbral_aviso_m3h: float
    desde: str
    hasta: str
    horas_revisadas: int
    max_m3h: Optional[float]
    picos_alerta: List[dict]
    picos_aviso: List[dict]
    estado: str  # OK | AVISO | ALERTA | SIN_DATOS | ERROR
    mensaje: str
    generado: str


def _parse_csv_horario(text: str) -> List[Tuple[datetime, float]]:
    """Parsea CSV TIME,VALUE (UTC) → lista (dt Chile, m³/h)."""
    out: List[Tuple[datetime, float]] = []
    lines = [ln for ln in text.strip().splitlines() if ln.strip()]
    if not lines:
        return out
    start = 1 if ("TIME" in lines[0].upper() or "FECHA" in lines[0].upper()) else 0
    for line in lines[start:]:
        parts = line.split(",")
        if len(parts) < 2:
            continue
        try:
            raw = parts[0].strip().replace("Z", "+00:00")
            dt = datetime.fromisoformat(raw)
            if dt.tzinfo is None:
                from datetime import timezone

                dt = dt.replace(tzinfo=timezone.utc)
            dt_cl = dt.astimezone(CHILE_TZ)
            val = float(parts[1].strip())
            out.append((dt_cl, val))
        except (ValueError, TypeError):
            continue
    return out


def obtener_horarios(node_id: str, dia: date) -> List[Tuple[datetime, float]]:
    base = acl_node_base_url()
    ds = dia.strftime("%d%m%Y")
    url = f"{base}/nodes/{node_id}/dates.measures.csv"
    r = requests.get(url, params=[("start", ds), ("end", ds)], timeout=45)
    r.raise_for_status()
    return _parse_csv_horario(r.text)


def revisar(dias: int = 2) -> ResultadoAlerta:
    ahora = datetime.now(CHILE_TZ)
    hasta = ahora.date()
    desde = hasta - timedelta(days=max(dias - 1, 0))
    generado = ahora.strftime("%Y-%m-%d %H:%M:%S %Z")

    series: List[Tuple[datetime, float]] = []
    try:
        d = desde
        while d <= hasta:
            series.extend(obtener_horarios(NODE_ID, d))
            d += timedelta(days=1)
    except Exception as exc:
        return ResultadoAlerta(
            ok=False,
            node_id=NODE_ID,
            node_name=NODE_NAME,
            umbral_alerta_m3h=CAUDAL_MAX_DN90_M3H,
            umbral_aviso_m3h=CAUDAL_AVISO_M3H,
            desde=desde.isoformat(),
            hasta=hasta.isoformat(),
            horas_revisadas=0,
            max_m3h=None,
            picos_alerta=[],
            picos_aviso=[],
            estado="ERROR",
            mensaje=f"Error consultando API WES: {exc}",
            generado=generado,
        )

    if not series:
        return ResultadoAlerta(
            ok=True,
            node_id=NODE_ID,
            node_name=NODE_NAME,
            umbral_alerta_m3h=CAUDAL_MAX_DN90_M3H,
            umbral_aviso_m3h=CAUDAL_AVISO_M3H,
            desde=desde.isoformat(),
            hasta=hasta.isoformat(),
            horas_revisadas=0,
            max_m3h=None,
            picos_alerta=[],
            picos_aviso=[],
            estado="SIN_DATOS",
            mensaje="Sin datos horarios en el rango revisado.",
            generado=generado,
        )

    alertas: List[Pico] = []
    avisos: List[Pico] = []
    max_v = max(v for _, v in series)
    for dt, v in series:
        stamp = dt.strftime("%Y-%m-%d %H:%M")
        if v >= CAUDAL_MAX_DN90_M3H:
            alertas.append(Pico(stamp, round(v, 2), "ALERTA"))
        elif v >= CAUDAL_AVISO_M3H:
            avisos.append(Pico(stamp, round(v, 2), "AVISO"))

    if alertas:
        estado = "ALERTA"
        mensaje = (
            f"ALERTA: {len(alertas)} hora(s) ≥ {CAUDAL_MAX_DN90_M3H:.0f} m³/h "
            f"(techo DN90). Posible recurrencia de error de memoria/placa. "
            f"Máximo: {max_v:.1f} m³/h."
        )
    elif avisos:
        estado = "AVISO"
        mensaje = (
            f"AVISO: {len(avisos)} hora(s) ≥ {CAUDAL_AVISO_M3H:.0f} m³/h "
            f"(elevado para DN90). Máximo: {max_v:.1f} m³/h."
        )
    else:
        estado = "OK"
        mensaje = (
            f"OK: sin caudales anómalos. Máximo del período: {max_v:.1f} m³/h "
            f"(umbral alerta {CAUDAL_MAX_DN90_M3H:.0f} m³/h)."
        )

    return ResultadoAlerta(
        ok=True,
        node_id=NODE_ID,
        node_name=NODE_NAME,
        umbral_alerta_m3h=CAUDAL_MAX_DN90_M3H,
        umbral_aviso_m3h=CAUDAL_AVISO_M3H,
        desde=desde.isoformat(),
        hasta=hasta.isoformat(),
        horas_revisadas=len(series),
        max_m3h=round(max_v, 2),
        picos_alerta=[asdict(p) for p in alertas],
        picos_aviso=[asdict(p) for p in avisos],
        estado=estado,
        mensaje=mensaje,
        generado=generado,
    )


def guardar(resultado: ResultadoAlerta) -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(CHILE_TZ).strftime("%Y%m%d_%H%M")
    path = OUT_DIR / f"alerta_etapa5_{stamp}.json"
    path.write_text(json.dumps(asdict(resultado), ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def main() -> int:
    parser = argparse.ArgumentParser(description="Alerta diaria Etapa N°5 Zapallar")
    parser.add_argument("--dias", type=int, default=2, help="Días hacia atrás a revisar (default 2)")
    parser.add_argument("--json", action="store_true", help="Salida solo JSON")
    parser.add_argument("--no-save", action="store_true", help="No guardar JSON en reports/")
    args = parser.parse_args()

    res = revisar(dias=args.dias)
    path = None if args.no_save else guardar(res)

    if args.json:
        payload = asdict(res)
        if path:
            payload["archivo"] = str(path)
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print("=" * 60)
        print(f"Alerta diaria — {COMPANY} / {NODE_NAME} ({NODE_ID})")
        print(f"Rango: {res.desde} → {res.hasta}  |  horas: {res.horas_revisadas}")
        print(f"Estado: {res.estado}")
        print(res.mensaje)
        if res.picos_alerta:
            print("\nPicos ALERTA:")
            for p in res.picos_alerta:
                print(f"  {p['fecha_hora_chile']}  {p['m3h']} m³/h")
        if res.picos_aviso:
            print("\nPicos AVISO:")
            for p in res.picos_aviso:
                print(f"  {p['fecha_hora_chile']}  {p['m3h']} m³/h")
        if path:
            print(f"\nJSON: {path}")
        print("=" * 60)

    return 0 if res.estado in ("OK", "AVISO", "SIN_DATOS") else 2


if __name__ == "__main__":
    raise SystemExit(main())
