"""
Gabinetes WES: qué nodos comparten el mismo gabinete (1 placa o hasta 4 placas).

La API no informa gabinete. Este registro es la fuente para el Excel de equipos
vigentes. Cada entrada es un gabinete con 1–5 nodos (una placa por nodo).
Solo `confianza: "confirmado"` se vuelca al Excel; el resto queda para que se complete a mano.

Para corregir: editar GABINETES y volver a generar
  python exportar_listado_equipos_puntos_en_cero.py
"""

from __future__ import annotations

from typing import Dict, List, Optional

# confianza: alta = mismo recinto y mismo armario evidente; media = mismo sitio, confirmar placas.
GABINETES: List[Dict[str, object]] = [
    # --- BUPA Antofagasta: 4 nodos / 4 placas en un gabinete ---
    {
        "id": "BUPA-ANTOF-01",
        "nombre": "BUPA Antofagasta — Clínica (4 placas)",
        "nodos": ["000029-07", "000029-08", "000029-09", "000029-10"],
        "confianza": "alta",
        "notas": "Un gabinete con 4 placas (Sala Bomba Principal, 6° piso, Sanitaria, Bomba N°2).",
    },
    # --- Parque Arauco ---
    {
        "id": "PA-EST-01",
        "nombre": "Parque Arauco Estación — gabinete 4 placas",
        "nodos": ["000025-01", "000025-04", "000025-07", "000025-19"],
        "confianza": "media",
        "notas": "Estanque Norte, Baños Públicos, Pizza Hut y Estanque Sur. Confirmar en terreno.",
    },
    {
        "id": "PA-MAM-01",
        "nombre": "Parque Arauco Maipú — gabinete 4 placas",
        "nodos": ["000025-08", "000025-10", "000025-32", "000025-33"],
        "confianza": "alta",
        "notas": "Placa Bancaria, Ripley, Pasillo Técnico y ARROW. Falabella (sala bombas) va aparte.",
    },
    {
        "id": "PA-AEB-01",
        "nombre": "Parque Arauco El Bosque — Anillo + Matriz A.A",
        "nodos": ["000025-12", "000025-30"],
        "confianza": "media",
        "notas": "Anillo Plaza y Matriz A.A (reemplazo de 000025-11). Confirmar si es el mismo gabinete.",
    },
    {
        "id": "PA-MAQ-01",
        "nombre": "Parque Arauco Quilicura — Matriz + Baños",
        "nodos": ["000025-13", "000025-34"],
        "confianza": "media",
        "notas": "Matriz Principal y Alimentación Baños.",
    },
    {
        "id": "PA-CUR-01",
        "nombre": "Parque Arauco Curauma — Anillo Norte/Sur",
        "nodos": ["000025-37", "000025-38"],
        "confianza": "alta",
        "notas": "CUR Anillo Sur y CUR Anillo Norte.",
    },
    {
        "id": "PA-PAK-ANDEN",
        "nombre": "Parque Arauco Kennedy — Andén 3-4",
        "nodos": ["000025-20", "000025-21", "000025-29"],
        "confianza": "alta",
        "notas": "Tres impulsiones Andén 3-4 (Matriz, Locales Gast., Restaurante). Confirmar si hay 4ª placa.",
    },
    {
        "id": "PA-PAK-PILETA",
        "nombre": "Parque Arauco Kennedy — Piletas",
        "nodos": ["000025-23", "000025-24"],
        "confianza": "alta",
        "notas": "Llenado Pileta y Llenado Pileta Cascada.",
    },
    # --- COPEC Costanera (confirmados en terreno) ---
    {
        "id": "COPEC-LAV-01",
        "nombre": "COPEC Costanera — Lavados",
        "nodos": ["000009-03", "000009-04", "000009-09", "000009-10"],
        "confianza": "confirmado",
        "notas": "Un gabinete con los cuatro lavados: automático N/S y autoservicio N/S.",
    },
    {
        "id": "COPEC-MATRIZ-01",
        "nombre": "COPEC Costanera — Matriz / riego / estanque",
        "nodos": ["000009-02", "000009-05", "000009-06"],
        "confianza": "confirmado",
        "notas": "Un gabinete: Estanque reutilización, Riego y Matriz principal.",
    },
    # --- DERCO Quilicura (7 nodos → 4+3 placas) ---
    {
        "id": "DERCO-Q-01",
        "nombre": "DERCO Quilicura — gabinete 4 placas",
        "nodos": ["000012-06", "000012-07", "000012-08", "000012-09"],
        "confianza": "media",
        "notas": "Matriz, Dercomaq, Lavado de máquinas y Casino. Confirmar armado real.",
    },
    {
        "id": "DERCO-Q-02",
        "nombre": "DERCO Quilicura — gabinete 3 placas",
        "nodos": ["000012-10", "000012-11", "000012-12"],
        "confianza": "media",
        "notas": "Proderco, Camarines y Edificio JCB.",
    },
    # --- Fundo Zapallar ---
    {
        "id": "ZAP-ETAPAS",
        "nombre": "Fundo Zapallar — etapas (4 placas)",
        "nodos": ["000027-04", "000027-06", "000027-07", "000027-08"],
        "confianza": "alta",
        "notas": "Etapa 1 al 4, Etapa 1, Etapa 2 y Etapa 3.",
    },
    {
        "id": "ZAP-MATRIZ",
        "nombre": "Fundo Zapallar — matriz / estanque (4 placas)",
        "nodos": ["000027-01", "000027-02", "000027-03", "000027-09"],
        "confianza": "media",
        "notas": "Matriz ESVAL, Estanque Inferior, Etapa 5 y Riego llenado ESVAL.",
    },
    # --- AGUNSA Lampa ---
    {
        "id": "AGUNSA-LAMPA",
        "nombre": "AGUNSA Lampa — módulos",
        "nodos": ["000020-02", "000020-03", "000020-04"],
        "confianza": "media",
        "notas": "Módulo D, ABC y E. Depósito e Intermodal San Antonio van aparte.",
    },
    # --- Nido estanques ---
    {
        "id": "NIDO-EST",
        "nombre": "Nido de Águilas — estanques B y C",
        "nodos": ["000007-01", "000007-07"],
        "confianza": "media",
        "notas": "Estanque B y Estanque C. Confirmar si comparten gabinete.",
    },
    # --- Renca Cumbre ---
    {
        "id": "RENCA-CUMBRE",
        "nombre": "Renca — Cumbre de Cóndores",
        "nodos": ["000017-07", "000017-08"],
        "confianza": "media",
        "notas": "Poniente y Oriente. Confirmar si es un gabinete o dos.",
    },
    # --- Club Providencia ---
    {
        "id": "CLUB-PROV",
        "nombre": "Club Providencia — Fitness / Piscina",
        "nodos": ["000031-01", "000031-02"],
        "confianza": "media",
        "notas": "Matriz Fitness y Matriz Piscina.",
    },
]


def _indice() -> Dict[str, Dict[str, object]]:
    out: Dict[str, Dict[str, object]] = {}
    for g in GABINETES:
        nodos = list(g.get("nodos") or [])
        for nid in nodos:
            if nid in out:
                raise ValueError(f"Nodo {nid} está en más de un gabinete")
            out[nid] = g
    return out


_INDICE = _indice()


def tipo_gabinete(n_placas: int) -> str:
    if n_placas <= 1:
        return "1 placa"
    return f"{n_placas} placas"


def relleno_confirmado(node_id: str, nombres: Dict[str, str]) -> Dict[str, str]:
    """Gabinete / tipo / nodos para el Excel, solo si está confirmado."""
    g = _INDICE.get(node_id)
    if not g or str(g.get("confianza") or "") != "confirmado":
        return {}
    ids = [str(x) for x in (g.get("nodos") or [])]
    n = len(ids)
    partes = []
    for nid in ids:
        nom = (nombres.get(nid) or "").strip()
        partes.append(f"{nid} {nom}".strip() if nom else nid)
    return {
        "gabinete": str(g.get("nombre") or ""),
        "tipo": tipo_gabinete(n),
        "nodos_gabinete": "; ".join(partes),
    }


def info_gabinete(
    node_id: str,
    *,
    company_name: str = "",
    node_name: str = "",
) -> Dict[str, str]:
    """Datos de gabinete para una fila del Excel."""
    g = _INDICE.get(node_id)
    if g:
        nodos = [str(x) for x in (g.get("nodos") or [])]
        n = len(nodos)
        otros = [x for x in nodos if x != node_id]
        return {
            "gabinete_id": str(g.get("id") or ""),
            "gabinete": str(g.get("nombre") or ""),
            "nodos_en_gabinete": str(n),
            "placas": str(n),
            "tipo": tipo_gabinete(n),
            "otros_nodos": ", ".join(otros),
            "confianza": str(g.get("confianza") or ""),
            "notas_gabinete": str(g.get("notas") or ""),
        }
    # Un nodo, una placa (colegio, mall con punto suelto, etc.)
    etiqueta = node_name or node_id
    cliente = company_name or ""
    nombre = f"{cliente} — {etiqueta}".strip(" —") if cliente else etiqueta
    return {
        "gabinete_id": f"SOLO-{node_id}",
        "gabinete": nombre + " (1 placa)",
        "nodos_en_gabinete": "1",
        "placas": "1",
        "tipo": tipo_gabinete(1),
        "otros_nodos": "",
        "confianza": "1 placa",
        "notas_gabinete": "Sin agrupación registrada: un nodo / una placa.",
    }


def filas_resumen_gabinetes(
    nodos: List[Dict[str, str]],
) -> List[Dict[str, str]]:
    """Una fila por gabinete (solo los que tocan el universo vigente)."""
    vigentes = {n["nodeId"]: n for n in nodos if n.get("nodeId")}
    vistos = set()
    filas: List[Dict[str, str]] = []

    # Primero gabinetes registrados que tengan al menos un vigente
    for g in GABINETES:
        ids = [str(x) for x in (g.get("nodos") or [])]
        presentes = [i for i in ids if i in vigentes]
        if not presentes:
            continue
        vistos.update(presentes)
        nombres = [vigentes[i].get("nodeName") or i for i in presentes]
        cliente = vigentes[presentes[0]].get("companyName") or ""
        n = len(presentes)
        filas.append(
            {
                "gabinete_id": str(g.get("id") or ""),
                "gabinete": str(g.get("nombre") or ""),
                "cliente": cliente,
                "nodos": str(n),
                "placas": str(n),
                "tipo": tipo_gabinete(n),
                "ids": ", ".join(presentes),
                "nombres": "; ".join(nombres),
                "confianza": str(g.get("confianza") or ""),
                "notas": str(g.get("notas") or ""),
            }
        )

    # Gabinetes de 1 placa no registrados
    for nid, n in sorted(vigentes.items(), key=lambda x: (x[1].get("companyName") or "", x[0])):
        if nid in vistos:
            continue
        info = info_gabinete(
            nid,
            company_name=n.get("companyName") or "",
            node_name=n.get("nodeName") or "",
        )
        filas.append(
            {
                "gabinete_id": info["gabinete_id"],
                "gabinete": info["gabinete"],
                "cliente": n.get("companyName") or "",
                "nodos": "1",
                "placas": "1",
                "tipo": tipo_gabinete(1),
                "ids": nid,
                "nombres": n.get("nodeName") or "",
                "confianza": info["confianza"],
                "notas": info["notas_gabinete"],
            }
        )
    return filas
