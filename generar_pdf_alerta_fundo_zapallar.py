"""
PDF de una página por noche de Fundo Zapallar.

Azul = seguimiento, amarillo = preventiva, rojo = crítica, gris = sin alerta.
La página diaria usa las filas que ya calculó alerta_nocturna_fundo_zapallar.py.
"""

from __future__ import annotations

import argparse
import csv
import re
from datetime import date, datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import Rectangle

NAVY = "#1B3A4B"
BLUE = "#2E7AC8"
RED = "#C0392B"
AMBER = "#D68910"
INK = "#1F2A30"
MUTED = "#5C6B73"
OK = "#8A97A0"
PAPER = "#F7F9FB"
LINE = "#E1E8ED"

SALIDA_DIR = Path("reports/Fundo_Zapallar/alerta_nocturna")

_HORAS = re.compile(r"(\d{2}):00=([0-9]+,[0-9]+|[0-9]+)")
_NOTA = re.compile(
    r"(La alerta de seguimiento se repite[^.]*\.|El saldo supera[^.]*\.)"
)


def _fmt(valor: float) -> str:
    texto = f"{float(valor):,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def _color(nivel: str) -> str:
    return {
        "seguimiento": BLUE,
        "preventiva": AMBER,
        "critica": RED,
        "cero": NAVY,
        "sin_dato": OK,
    }.get(nivel, OK)


def _etiqueta(nivel: str) -> str:
    return {
        "normal": "Sin alerta",
        "seguimiento": "Seguimiento",
        "preventiva": "Preventiva",
        "critica": "Crítica",
        "cero": "En cero",
        "sin_dato": "Sin dato",
    }.get(nivel, nivel)


def _nivel_maximo(filas: list[dict]) -> str:
    orden = {"critica": 4, "cero": 3, "preventiva": 2, "seguimiento": 1}
    mejor = "normal"
    for fila in filas:
        if orden.get(fila["nivel"], 0) > orden.get(mejor, 0):
            mejor = fila["nivel"]
    return mejor


def _titulo_nivel(nivel: str) -> str:
    return {
        "critica": "Alerta crítica",
        "cero": "Punto en cero",
        "preventiva": "Alerta preventiva",
        "seguimiento": "Alerta de seguimiento",
    }.get(nivel, "Sin alerta")


def _horas(detalle: str) -> list[tuple[str, float]]:
    return [(hora, float(valor.replace(",", "."))) for hora, valor in _HORAS.findall(detalle)]


def _nota(detalle: str) -> str:
    hallazgo = _NOTA.search(detalle)
    return hallazgo.group(1) if hallazgo else ""


def _base(detalle: str) -> str:
    return _NOTA.sub("", detalle).strip(" .")


def _fecha_corta(iso: str) -> str:
    return datetime.strptime(iso, "%Y-%m-%d").strftime("%d-%m-%Y")


def _encabezado(fig: plt.Figure, derecha: str, bajada: str) -> None:
    ax = fig.add_axes([0, 0.948, 1, 0.052])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.add_patch(Rectangle((0, 0), 1, 1, color=NAVY, transform=ax.transAxes))
    ax.text(0.045, 0.56, "WES", color="white", fontsize=13, fontweight="bold", va="center")
    ax.text(0.115, 0.56, "Seguimiento técnico", color="#D6E4EE", fontsize=8, va="center")
    ax.text(0.955, 0.66, derecha, color="white", fontsize=10, fontweight="bold", va="center", ha="right")
    ax.text(0.955, 0.28, bajada, color="#D6E4EE", fontsize=7.5, va="center", ha="right")


def _pie(fig: plt.Figure) -> None:
    fig.text(
        0.055,
        0.018,
        "WES SpA  ·  Seguimiento o preventiva: contrastar con riegos, llenados u operaciones extraordinarias antes de escalar al fundo. "
        "Crítica: aviso prioritario.\n"
        "Azul seguimiento  ·  amarillo preventiva  ·  rojo crítica  ·  gris sin alerta. "
        "Si un seguimiento se repite al día siguiente, pasa a preventiva.",
        fontsize=6.1,
        color=MUTED,
        va="bottom",
    )


def _tarjeta(fig: plt.Figure, caja: tuple[float, float, float, float], fila: dict) -> None:
    x0, y0, _ancho, _alto = caja
    ax = fig.add_axes([x0, y0, _ancho, _alto])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    color = _color(fila["nivel"])
    ax.add_patch(Rectangle((0, 0), 1, 1, facecolor=PAPER, edgecolor=LINE, lw=0.6))
    ax.add_patch(Rectangle((0, 0), 0.012, 1, color=color, transform=ax.transAxes))
    ax.text(0.04, 0.86, fila["punto"], fontsize=9, color=NAVY, fontweight="bold", va="center")
    ax.text(0.97, 0.86, _etiqueta(fila["nivel"]), fontsize=8.5, color=color, fontweight="bold", va="center", ha="right")
    ax.text(0.04, 0.66, f"{_fmt(float(fila['valor']))} m³", fontsize=13, color=color, fontweight="bold", va="center")
    ax.text(0.97, 0.66, fila["periodo"], fontsize=7, color=MUTED, va="center", ha="right")

    horas = _horas(fila["detalle"])
    if horas:
        techo = max(valor for _, valor in horas) or 1.0
        for i, (hora, valor) in enumerate(horas):
            x = 0.04 + i * 0.052
            alto_barra = 0.14 * (valor / techo) if valor else 0.01
            ax.add_patch(Rectangle((x, 0.40), 0.036, alto_barra, color=color if valor else "#D5DDE3"))
            ax.text(x + 0.018, 0.36, hora, fontsize=5.3, color=MUTED, ha="center", va="top")
        ax.text(
            0.44,
            0.46,
            _base(fila["detalle"]).split(". 00:")[0].split(" 00:")[0],
            fontsize=6.3,
            color=INK,
            va="center",
        )
    else:
        ax.text(0.04, 0.44, _base(fila["detalle"]), fontsize=7.2, color=INK, va="center")

    nota = _nota(fila["detalle"])
    ax.text(0.04, 0.20, fila["umbral"], fontsize=5.5, color=MUTED, va="center")
    if nota:
        ax.text(0.04, 0.07, nota, fontsize=6.2, color=color, va="center")
    else:
        ax.text(0.04, 0.07, fila["accion"], fontsize=6.3, color=INK, va="center")


def pagina_dia(filas: list[dict]) -> plt.Figure:
    fig = plt.figure(figsize=(8.27, 11.69), facecolor="white")
    noches = [f for f in filas if "00:00" in f["periodo"]]
    balances = [f for f in filas if "día completo" in f["periodo"]]
    noche = noches[0]["fecha"] if noches else filas[0]["fecha"]
    nivel = _nivel_maximo(filas)
    _encabezado(fig, "Fundo Zapallar", _fecha_corta(noche))
    fig.text(0.055, 0.918, _titulo_nivel(nivel), fontsize=16, color=_color(nivel) if nivel != "normal" else NAVY, fontweight="bold")
    fig.text(
        0.055,
        0.894,
        "Tres niveles: seguimiento, preventiva y crítica. Ventana 00:00 a 06:00. "
        "La crítica queda para consumos excepcionalmente altos.",
        fontsize=7.4,
        color=INK,
    )
    orden = noches + balances
    alto = 0.148
    gap = 0.012
    y = 0.868
    for fila in orden:
        y -= alto
        _tarjeta(fig, (0.055, y, 0.89, alto - 0.004), fila)
        y -= gap
    _pie(fig)
    return fig


def _celda(ax, x: float, y: float, fila: dict | None) -> None:
    if fila is None:
        ax.text(x, y, "—", fontsize=7, color=MUTED, va="center")
        return
    color = _color(fila["nivel"])
    ax.plot([x - 0.012], [y], marker="o", color=color, markersize=4.2, transform=ax.transAxes)
    ax.text(x, y + 0.012, f"{_fmt(float(fila['valor']))} m³", fontsize=7.2, color=INK, va="center", fontweight="bold")
    ax.text(x, y - 0.014, _etiqueta(fila["nivel"]), fontsize=6.2, color=color, va="center")


def pagina_resumen(dias: list[tuple[date, list[dict]]]) -> plt.Figure:
    fig = plt.figure(figsize=(8.27, 11.69), facecolor="white")
    inicio = dias[0][0].strftime("%d-%m-%Y")
    fin = dias[-1][0].strftime("%d-%m-%Y")
    _encabezado(fig, "Fundo Zapallar", f"{inicio}  ·  {fin}")
    fig.text(0.055, 0.918, "Alertas nocturnas y balances diarios", fontsize=15, color=NAVY, fontweight="bold")
    fig.text(
        0.055,
        0.892,
        "Cada fila es una noche cerrada. El balance corresponde al día calendario anterior.\n"
        "Estanque del Medio = Estanque Inferior (000027-02).",
        fontsize=7.4,
        color=INK,
        va="top",
    )

    ax = fig.add_axes([0.04, 0.16, 0.92, 0.69])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.add_patch(Rectangle((0, 0), 1, 1, facecolor=PAPER, edgecolor=LINE, lw=0.6))
    columnas = [
        (0.04, "Noche"),
        (0.18, "Etapa N°5"),
        (0.35, "Etapas 1-4"),
        (0.52, "Matriz ESVAL"),
        (0.69, "Matriz − Estanque"),
        (0.86, "Estanque − etapas"),
    ]
    ax.add_patch(Rectangle((0, 0.955), 1, 0.045, color=NAVY))
    for x, titulo in columnas:
        ax.text(x, 0.977, titulo, fontsize=6.6, color="white", fontweight="bold", va="center")

    n = len(dias)
    alto_fila = 0.90 / max(n, 1)
    for i, (noche, filas) in enumerate(dias):
        y = 0.93 - (i + 0.5) * alto_fila
        if i % 2 == 0:
            ax.add_patch(Rectangle((0, y - alto_fila / 2), 1, alto_fila, color="#EEF3F6", lw=0))
        por_regla = {fila["regla"]: fila for fila in filas}
        ax.text(0.04, y + 0.012, noche.strftime("%d-%m"), fontsize=8, color=NAVY, fontweight="bold", va="center")
        balances = [f for f in filas if "día completo" in f["periodo"]]
        if balances:
            ax.text(0.04, y - 0.014, f"bal. {_fecha_corta(balances[0]['fecha'])[0:5]}", fontsize=5.4, color=MUTED, va="center")
        _celda(ax, 0.20, y, por_regla.get("etapa5-noche"))
        _celda(ax, 0.37, y, por_regla.get("etapas14-noche"))
        _celda(ax, 0.54, y, por_regla.get("matriz-noche"))
        _celda(ax, 0.71, y, por_regla.get("balance-matriz-estanque"))
        _celda(ax, 0.88, y, por_regla.get("balance-estanque-etapas"))

    ceros = [
        noche.strftime("%d-%m")
        for noche, filas in dias
        if any(f["regla"] == "etapas14-noche" and float(f["valor"]) == 0 and f["nivel"] == "normal" for f in filas)
    ]
    if ceros:
        fig.text(
            0.055,
            0.118,
            f"Etapas 1-4 quedó en 0,00 m³ en la noche del {ceros[0]} al {ceros[-1]}.\n"
            "Si el día completo también está en cero, el balance del estanque sube aunque el medidor haya respondido.",
            fontsize=7,
            color=INK,
            va="top",
        )
    _pie(fig)
    return fig


def leer_csv(path: Path) -> list[dict]:
    with path.open(encoding="utf-8", newline="") as fh:
        return list(csv.DictReader(fh))


def csv_del_dia(noche: date, carpeta: Path = SALIDA_DIR) -> Path:
    return carpeta / f"alerta_fundo_zapallar_{noche.strftime('%Y%m%d')}.csv"


def guardar_pdf_dia(filas: list[dict], noche: date, carpeta: Path = SALIDA_DIR) -> Path:
    carpeta.mkdir(parents=True, exist_ok=True)
    destino = carpeta / f"Alerta_Fundo_Zapallar_{noche.strftime('%Y%m%d')}.pdf"
    fig = pagina_dia(filas)
    fig.savefig(destino, format="pdf")
    plt.close(fig)
    return destino


def guardar_pdf_periodo(inicio: date, fin: date, carpeta: Path = SALIDA_DIR) -> Path:
    dias: list[tuple[date, list[dict]]] = []
    cursor = inicio
    while cursor <= fin:
        path = csv_del_dia(cursor, carpeta)
        if path.is_file():
            dias.append((cursor, leer_csv(path)))
        cursor = date.fromordinal(cursor.toordinal() + 1)
    if not dias:
        raise FileNotFoundError(f"No hay CSV entre {inicio.isoformat()} y {fin.isoformat()}")
    carpeta.mkdir(parents=True, exist_ok=True)
    destino = carpeta / (
        f"Alerta_Fundo_Zapallar_{inicio.strftime('%Y%m%d')}_{fin.strftime('%Y%m%d')}.pdf"
    )
    with PdfPages(destino) as pdf:
        resumen = pagina_resumen(dias)
        pdf.savefig(resumen)
        plt.close(resumen)
        for _noche, filas in dias:
            fig = pagina_dia(filas)
            pdf.savefig(fig)
            plt.close(fig)
    return destino


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="PDF de las alertas de Fundo Zapallar.")
    parser.add_argument("--fecha", help="Una noche, YYYY-MM-DD.")
    parser.add_argument("--desde", help="Inicio del período, YYYY-MM-DD.")
    parser.add_argument("--hasta", help="Fin del período, YYYY-MM-DD.")
    args = parser.parse_args(argv)
    if args.desde and args.hasta:
        path = guardar_pdf_periodo(date.fromisoformat(args.desde), date.fromisoformat(args.hasta))
    else:
        if args.fecha:
            noche = date.fromisoformat(args.fecha)
        else:
            candidatos = sorted(SALIDA_DIR.glob("alerta_fundo_zapallar_*.csv"))
            if not candidatos:
                raise SystemExit("No hay CSV de alerta para armar el PDF.")
            noche = datetime.strptime(candidatos[-1].stem[-8:], "%Y%m%d").date()
        path = guardar_pdf_dia(leer_csv(csv_del_dia(noche)), noche)
    print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
