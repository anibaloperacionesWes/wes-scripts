"""
Genera SOLO la presentación PPT para Providencia, filtrando por nodeId específicos.

Uso (PowerShell):
  python generar_ppt_solo_providencia.py

Notas:
- Obtiene los datos desde la API (obtener_datos_agregados) y genera la PPT con crear_ppt_analisis (lógica centralizada en Maipú).
- Permite normalizar typos comunes en nodeId (ej: 0000006-04 -> 000006-04).
"""

import sys
from pathlib import Path
from datetime import datetime
import importlib.util

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from generar_reporte_word import get_company_name


def _normalizar_node_id(node_id: str) -> str:
    n = (node_id or "").strip()
    # Typo frecuente reportado por el usuario
    n = n.replace("0000006-", "000006-")
    return n


def main():
    COMPANY_ID = "000006"  # Providencia
    START_DATE = "27/12/2025"
    END_DATE = "01/12/2026"

    # Solo estos nodos (según solicitud del usuario)
    NODE_IDS = ["000006-01", "000006-02", "0000006-04", "000006-05"]
    node_ids = []
    for x in NODE_IDS:
        nx = _normalizar_node_id(x)
        if nx and nx not in node_ids:
            node_ids.append(nx)

    print("=" * 70)
    print("  GENERACIÓN SOLO PPT - PROVIDENCIA (NODOS ESPECÍFICOS)")
    print(f"  Company ID: {COMPANY_ID}")
    print(f"  Nodos: {', '.join(node_ids)}")
    print(f"  Período: {START_DATE} - {END_DATE}")
    print("=" * 70)
    print()

    # Importar funciones de PPT desde el script de Maipú (centralizado)
    maipu_script = Path(__file__).parent / "generar_reportes_y_ppt_mall_maipu.py"
    spec = importlib.util.spec_from_file_location("maipu_ppt", maipu_script)
    maipu_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(maipu_module)

    obtener_datos_agregados = maipu_module.obtener_datos_agregados
    guardar_datos_json = maipu_module.guardar_datos_json
    crear_ppt_analisis = maipu_module.crear_ppt_analisis
    convertir_ppt_a_pdf = maipu_module.convertir_ppt_a_pdf

    # Directorio de salida
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    # Para Providencia la estructura estándar es: reports\Providencia\ABREGADO
    output_dir_base = Path("reports") / "Providencia" / "ABREGADO"
    output_dir_base.mkdir(parents=True, exist_ok=True)
    ppt_dir = output_dir_base / f"AGREGADO_{timestamp}"
    ppt_dir.mkdir(exist_ok=True)

    # Obtener datos y generar PPT
    print("[1/2] Obteniendo datos agregados desde API...")
    datos = obtener_datos_agregados(node_ids, START_DATE, END_DATE)
    # Forzar uso de API en PPT (especialmente para gráficas horario máx/mín y alertas)
    datos["_usar_api"] = True
    print(f"[OK] Datos obtenidos: {len(datos.get('all_measures', []))} medidas")

    # Guardar JSON por trazabilidad / reuso
    json_path = ppt_dir / "datos_agregados.json"
    guardar_datos_json(datos, json_path)
    print(f"[OK] Datos guardados en JSON: {json_path}")
    print()

    print("[2/2] Generando presentación PPT...")
    ppt_path = ppt_dir / "Agregado PPT.pptx"
    crear_ppt_analisis(datos, ppt_path, mall_name="Providencia")
    print(f"[OK] PPT generada: {ppt_path}")

    # PDF (best-effort)
    pdf_path = convertir_ppt_a_pdf(ppt_path)
    if pdf_path:
        print(f"[OK] PDF generado: {pdf_path}")


if __name__ == "__main__":
    main()

