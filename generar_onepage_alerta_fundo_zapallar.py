"""
Una página formal de las alertas de Fundo Zapallar, 23-09-2026 al 14-10-2026.

Tres niveles: preventiva (azul), seguimiento (amarillo) y crítica (rojo).
"""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

OUT = Path("reports/Fundo_Zapallar/alerta_nocturna/Alerta_Fundo_Zapallar_20260923_20261014.pdf")

NAVY = "#1B3A4B"
BLUE = "#2E7AC8"
RED = "#C0392B"
AMBER = "#D68910"
INK = "#1F2A30"
MUTED = "#5C6B73"
GRID = "#E6ECF0"
OK = "#D5DDE3"

# Total nocturno de Etapa N°5 (m³). None = noche sin alerta.
ETAPA5 = {
    date(2026, 9, 23): 3.20,
    date(2026, 9, 24): 4.20,
    date(2026, 9, 25): 3.10,
    date(2026, 9, 26): 3.30,
    date(2026, 9, 27): 3.10,
    date(2026, 9, 28): 3.10,
    date(2026, 9, 29): None,
    date(2026, 9, 30): 9.00,
    date(2026, 10, 1): 2.80,
    date(2026, 10, 2): 3.10,
    date(2026, 10, 3): 5.00,
    date(2026, 10, 4): 2.90,
    date(2026, 10, 5): 4.80,
    date(2026, 10, 6): 7.50,
    date(2026, 10, 7): 10.30,
    date(2026, 10, 8): 6.20,
    date(2026, 10, 9): 6.00,
    date(2026, 10, 10): 5.40,
    date(2026, 10, 11): 5.80,
    date(2026, 10, 12): 5.60,
    date(2026, 10, 13): 8.50,
    date(2026, 10, 14): 5.30,
}


def _dias() -> list[date]:
    d = date(2026, 9, 23)
    out = []
    while d <= date(2026, 10, 14):
        out.append(d)
        d += timedelta(days=1)
    return out


def _fmt(valor: float) -> str:
    return f"{valor:.2f}".replace(".", ",")


def _color_etapa5(valor: float | None) -> str:
    """Azul = preventiva (>1 a 5). Amarillo = seguimiento (>5 a 10). Rojo = crítica (>10)."""
    if valor is None:
        return OK
    if valor > 10:
        return RED
    if valor > 5:
        return AMBER
    return BLUE


def _color_nivel(nivel: str) -> str:
    if nivel.startswith("Crítica"):
        return RED
    if nivel.startswith("Seguimiento"):
        return AMBER
    return BLUE


def construir() -> Path:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    dias = _dias()
    fig = plt.figure(figsize=(8.27, 11.69), facecolor="white")

    encabezado = fig.add_axes([0, 0.935, 1, 0.065])
    encabezado.set_xlim(0, 1)
    encabezado.set_ylim(0, 1)
    encabezado.axis("off")
    encabezado.add_patch(plt.Rectangle((0, 0), 1, 1, color=NAVY, transform=encabezado.transAxes))
    encabezado.text(0.045, 0.62, "WES", color="white", fontsize=15, fontweight="bold", va="center")
    encabezado.text(0.12, 0.64, "Seguimiento técnico", color="#D6E4EE", fontsize=8, va="center")
    encabezado.text(
        0.955,
        0.64,
        "Fundo Zapallar",
        color="white",
        fontsize=11,
        fontweight="bold",
        va="center",
        ha="right",
    )
    encabezado.text(
        0.955,
        0.30,
        "23-09-2026  ·  14-10-2026",
        color="#D6E4EE",
        fontsize=8,
        va="center",
        ha="right",
    )

    fig.text(0.055, 0.908, "Alertas nocturnas y balances diarios", fontsize=15, color=NAVY, fontweight="bold")
    fig.text(
        0.055,
        0.888,
        "Tres niveles de aviso: preventiva, seguimiento y crítica. Ventana 00:00 a 06:00. La categoría crítica se reserva para\n"
        "consumos excepcionalmente altos. Si una preventiva se repite al día siguiente, pasa a seguimiento.\n"
        "Estanque del Medio = Estanque Inferior (000027-02).",
        fontsize=7.6,
        color=INK,
        va="top",
        linespacing=1.35,
    )

    # Tarjetas
    tarjetas = [
        ("49", "alertas en el período"),
        ("21 / 22", "noches con seguimiento · Etapa N°5"),
        ("10,30 m³", "máximo nocturno Etapa N°5, 07-10"),
        ("70,82 m³", "máximo nocturno Matriz, 30-09"),
    ]
    acentos = [BLUE, AMBER, RED, "#C5D0D6"]
    for i, (grande, etiqueta) in enumerate(tarjetas):
        ax = fig.add_axes([0.055 + i * 0.232, 0.798, 0.215, 0.055])
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")
        ax.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor="#F4F7F9", edgecolor="#D5DEE5", lw=0.6))
        ax.add_patch(plt.Rectangle((0, 0), 0.025, 1, color=acentos[i]))
        ax.text(0.10, 0.62, grande, fontsize=12, color=NAVY, fontweight="bold", va="center")
        ax.text(0.10, 0.26, etiqueta, fontsize=6.4, color=MUTED, va="center")

    ax = fig.add_axes([0.055, 0.568, 0.62, 0.188])
    xs = list(range(len(dias)))
    valores = [ETAPA5[d] for d in dias]
    colores = [_color_etapa5(valor) for valor in valores]
    altos = [0 if v is None else v for v in valores]
    ax.bar(xs, altos, color=colores, width=0.78, zorder=3)
    for x, valor in zip(xs, valores):
        if valor is None:
            continue
        # El rótulo vertical de 9,00 cruzaría la línea de crítica (10 m³).
        if valor <= 10 and valor + 1.35 > 10:
            ax.text(
                x,
                valor - 0.12,
                _fmt(valor),
                ha="center",
                va="top",
                rotation=90,
                fontsize=5.7,
                color="white",
                clip_on=False,
                zorder=4,
            )
        else:
            ax.text(
                x,
                valor + 0.18,
                _fmt(valor),
                ha="center",
                va="bottom",
                rotation=90,
                fontsize=5.7,
                color=INK,
                clip_on=False,
                zorder=4,
            )
    ax.axhline(1, color=BLUE, lw=0.8, ls=":", zorder=2)
    ax.axhline(5, color=AMBER, lw=0.8, ls="--", zorder=2)
    ax.axhline(10, color=RED, lw=0.9, ls="-", zorder=2)
    ax.set_xlim(-0.6, len(dias) - 0.4)
    ax.set_ylim(0, 13.6)
    ax.set_ylabel("m³ por noche", fontsize=8, color=INK)
    ax.set_title(
        "Etapa N°5 — m³ entre las 00:00 y las 06:00",
        loc="left",
        fontsize=9,
        color=NAVY,
        pad=8,
    )
    ax.set_xticks(xs)
    ax.set_xticklabels([d.strftime("%d-%m") for d in dias], rotation=90, fontsize=5.8, color=MUTED)
    ax.tick_params(axis="x", pad=1)
    ax.tick_params(axis="y", labelsize=7, colors=MUTED)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#D5DEE5")
    ax.spines["bottom"].set_color("#D5DEE5")
    ax.yaxis.grid(True, color=GRID, zorder=0)
    ax.set_axisbelow(True)
    ax.legend(
        handles=[
            Patch(facecolor=BLUE, label="Preventiva: >1 a 5 m³"),
            Patch(facecolor=AMBER, label="Seguimiento: >5 a 10 m³"),
            Patch(facecolor=RED, label="Crítica: >10 m³ o >4 m³/h"),
            Patch(facecolor=OK, label="29-09 sin alerta"),
        ],
        frameon=False,
        fontsize=6.4,
        loc="center left",
        bbox_to_anchor=(1.03, 0.62),
        borderaxespad=0,
        handlelength=1.1,
        handleheight=0.8,
        labelspacing=0.55,
    )

    fig.text(0.055, 0.512, "Otras reglas que se cumplieron", fontsize=11, color=NAVY, fontweight="bold")

    bloques = [
        (
            "Etapas 1-4  ·  noche",
            "Preventiva >40 a 45 m³ o >12 a 15 m³/h. Seguimiento >45 a 60 m³\n"
            "o >15 a 20 m³/h. Crítica >60 m³ o >20 m³/h.",
            [
                ("23-09", "40,36 m³", "Preventiva"),
                ("27-09", "39,13 m³ · máx. 12,24", "Preventiva"),
                ("29-09", "37,02 m³ · máx. 13,08", "Preventiva"),
                ("30-09", "44,72 m³", "Seguimiento, se repite"),
                ("02-10", "47,05 m³", "Seguimiento"),
                ("03-10", "42,78 m³", "Seguimiento, se repite"),
                ("04-10", "33,37 m³ · máx. 12,49", "Seguimiento, se repite"),
                ("05-10", "41,62 m³", "Seguimiento, se repite"),
            ],
        ),
        (
            "Matriz ESVAL  ·  noche",
            "Preventiva >55 a 65 m³ o >25 a 30 m³/h. Seguimiento >65 a 80 m³\n"
            "o >30 a 35 m³/h. Crítica >80 m³ o >35 m³/h.",
            [
                ("28-09", "58,20 m³", "Preventiva"),
                ("30-09", "70,82 m³", "Seguimiento"),
                ("01-10", "60,83 m³", "Seguimiento, se repite"),
                ("02-10", "59,25 m³", "Seguimiento, se repite"),
                ("03-10", "70,33 m³", "Seguimiento"),
                ("05-10", "57,30 m³", "Preventiva"),
                ("07-10", "58,27 m³", "Preventiva"),
            ],
        ),
        (
            "Matriz − Estanque Inferior  ·  día",
            "Preventiva >21 a 25 m³. Seguimiento >25 a 42 m³ o >21 m³ durante\n"
            "2 días. Crítica >42 m³.",
            [
                ("23-09", "23,02 m³", "Seguimiento, viene de 28,66 el 22-09"),
                ("01-10", "36,43 m³", "Seguimiento"),
                ("03-10", "25,96 m³", "Seguimiento"),
                ("05-10", "23,31 m³", "Preventiva"),
                ("07-10", "29,47 m³", "Seguimiento"),
                ("13-10", "23,54 m³", "Preventiva"),
            ],
        ),
        (
            "Estanque − (Etapa N°5 + Etapas 1-4)  ·  día",
            "Preventiva >30 a 36 m³. Seguimiento >36 a 60 m³ o >30 m³ durante\n"
            "2 días. Crítica >60 m³.",
            [
                ("23-09", "31,92 m³", "Preventiva"),
                ("24-09", "41,12 m³", "Seguimiento"),
                ("26-09", "35,35 m³", "Preventiva"),
                ("09-10", "40,08 m³", "Seguimiento"),
                ("10-10", "40,28 m³", "Seguimiento"),
                ("11-10", "32,88 m³", "Seguimiento, se repite"),
                ("12-10", "40,60 m³", "Seguimiento"),
            ],
        ),
    ]

    cajas = [
        (0.055, 0.272, 0.44, 0.210),
        (0.515, 0.272, 0.43, 0.210),
        (0.055, 0.046, 0.44, 0.208),
        (0.515, 0.046, 0.43, 0.208),
    ]
    for (titulo, criterio, filas), (x0, y0, ancho, alto) in zip(bloques, cajas):
        axb = fig.add_axes([x0, y0, ancho, alto])
        axb.set_xlim(0, 1)
        axb.set_ylim(0, 1)
        axb.axis("off")
        axb.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor="#F8FAFB", edgecolor="#E1E8ED", lw=0.6))
        axb.text(0.04, 0.92, titulo, fontsize=8, color=NAVY, fontweight="bold", va="center")
        axb.text(0.04, 0.80, criterio, fontsize=5.5, color=MUTED, va="top", linespacing=1.25)
        paso = 0.62 / max(len(filas), 1)
        for i, (fecha, valor, nivel) in enumerate(filas):
            yy = 0.64 - i * paso
            color = _color_nivel(nivel)
            axb.plot([0.045], [yy], marker="o", color=color, markersize=4.5, transform=axb.transAxes)
            axb.text(0.09, yy, f"{fecha}    {valor}", fontsize=6.5, color=INK, va="center")
            axb.text(0.97, yy, nivel, fontsize=6.0, color=color, va="center", ha="right")

    fig.text(
        0.055,
        0.012,
        "WES SpA  ·  Documento de una página para revisión interna. Una alerta preventiva o de seguimiento requiere contraste con riegos, llenados u operaciones\n"
        "extraordinarias antes de escalar al fundo. Una alerta crítica requiere aviso prioritario. El 14-10 el día calendario aún no estaba cerrado; el balance no se incluye.",
        fontsize=6.2,
        color=MUTED,
        va="bottom",
    )

    fig.savefig(OUT, format="pdf")
    fig.savefig(OUT.with_suffix(".png"), dpi=140)
    plt.close(fig)
    return OUT


if __name__ == "__main__":
    path = construir()
    print(path)
