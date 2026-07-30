"""
Script para gestionar incidentes usando las APIs de WES.
Basado en PROMPT_INCIDENTES.md

Uso:
    python gestionar_incidentes.py --node-id 000025-20 --start-date 01112025 --end-date 07122025
"""

import argparse
import sys
from pathlib import Path
from typing import Optional, Union, Sequence
from datetime import datetime, timezone
import requests
import json

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# URLs base de las APIs
BASE_URL = "http://104.248.53.141:7003/wes/api/acl-node/v1"
ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"


def fetch_json(url: str, params: Optional[Union[dict, Sequence[tuple]]] = None) -> Union[dict, list]:
    """
    Realiza una petición GET a la API y retorna el JSON parseado.
    
    Args:
        url: URL completa del endpoint
        params: Parámetros de consulta (dict o lista de tuplas)
    
    Returns:
        dict o list con la respuesta JSON
    
    Raises:
        requests.RequestException: Si la petición falla
    """
    try:
        response = requests.get(url, params=params, timeout=30)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        print(f"[ERROR] Error al consultar {url}: {e}")
        raise


def obtener_informacion_nodo(node_id: str) -> Optional[dict]:
    """Obtiene información completa de un nodo."""
    try:
        return fetch_json(f"{BASE_URL}/nodes/{node_id}")
    except Exception as e:
        print(f"[ERROR] No se pudo obtener información del nodo {node_id}: {e}")
        return None


def obtener_medidas_consumo(node_id: str, start_date: str, end_date: str) -> Optional[dict]:
    """Obtiene medidas de consumo para un periodo."""
    try:
        return fetch_json(
            f"{BASE_URL}/nodes/measures/dates",
            params=[("id", node_id), ("start", start_date), ("end", end_date)],
        )
    except Exception as e:
        print(f"[ERROR] No se pudieron obtener medidas de consumo: {e}")
        return None


def obtener_alertas(node_id: str, start_date: str, end_date: str) -> list:
    """Obtiene alertas de fuga con estrategia de fallback."""
    # Intentar 1: /nodes/myalert/alerts con ddMMyyyy
    try:
        alerts = fetch_json(
            f"{BASE_URL}/nodes/myalert/alerts",
            params=[("id", node_id), ("start", start_date), ("end", end_date)],
        )
        if isinstance(alerts, list):
            return alerts
    except Exception:
        pass
    
    # Intentar 2: /nodes/leak/alerts con ddMMyyyy
    try:
        alerts = fetch_json(
            f"{BASE_URL}/nodes/leak/alerts",
            params=[("id", node_id), ("start", start_date), ("end", end_date)],
        )
        if isinstance(alerts, list):
            return alerts
    except Exception:
        pass
    
    # Intentar 3: /nodes/leak/alerts con ISO
    try:
        start_dt = datetime.strptime(start_date, "%d%m%Y")
        end_dt = datetime.strptime(end_date, "%d%m%Y")
        start_iso = start_dt.isoformat() + "Z"
        end_iso = end_dt.isoformat() + "Z"
        
        alerts = fetch_json(
            f"{BASE_URL}/nodes/leak/alerts",
            params=[("id", node_id), ("start", start_iso), ("end", end_iso)],
        )
        if isinstance(alerts, list):
            return alerts
    except Exception:
        pass
    
    return []


def obtener_datos_horarios(node_id: str, date_str: str) -> list:
    """Obtiene datos horarios para un día específico en formato CSV."""
    try:
        url = f"{BASE_URL}/nodes/{node_id}/dates.measures.csv"
        params = [("start", date_str), ("end", date_str)]
        
        response = requests.get(url, params=params, timeout=30)
        response.raise_for_status()
        
        # Parsear CSV
        hourly_data = []
        lines = response.text.strip().split('\n')
        
        for line in lines[1:]:  # Saltar encabezado
            if not line.strip():
                continue
            
            parts = line.split(',')
            if len(parts) >= 2:
                try:
                    time_str = parts[0].strip()
                    value_str = parts[1].strip()
                    
                    # Extraer hora del formato ISO
                    if 'T' in time_str:
                        hour_part = time_str.split('T')[1]
                        hour = int(hour_part.split(':')[0])
                    else:
                        dt = datetime.fromisoformat(time_str.replace('Z', '+00:00'))
                        hour = dt.hour
                    
                    value = float(value_str)
                    hourly_data.append((hour, value))
                except (ValueError, TypeError, IndexError):
                    continue
        
        return sorted(hourly_data, key=lambda x: x[0])
    except Exception as e:
        print(f"[ERROR] No se pudieron obtener datos horarios: {e}")
        return []


def obtener_informacion_empresa(company_id: str) -> Optional[dict]:
    """Obtiene información de una empresa."""
    try:
        return fetch_json(f"{ENTITY_BASE_URL}/companies/{company_id}")
    except Exception as e:
        print(f"[ERROR] No se pudo obtener información de la empresa {company_id}: {e}")
        return None


def obtener_precio_agua(node_id: str) -> Optional[float]:
    """Obtiene el precio del agua para un nodo."""
    try:
        # Extraer company_id del node_id (parte antes del guion)
        company_id = node_id.split("-")[0]
        
        # Intentar obtener desde /nodes/{nodeIdBase}
        data = fetch_json(f"{ENTITY_BASE_URL}/nodes/{company_id}")
        
        # Buscar el precio en diferentes campos
        # (lógica similar a get_water_price_per_m3 en generar_reporte_word.py)
        # Por ahora retornar None, se puede implementar la lógica completa
        return None
    except Exception as e:
        print(f"[ERROR] No se pudo obtener precio del agua: {e}")
        return None


def analizar_incidente(node_id: str, start_date: str, end_date: str) -> dict:
    """
    Analiza un incidente completo para un nodo en un periodo.
    Retorna un diccionario con toda la información recopilada.
    """
    print("=" * 60)
    print(f"ANÁLISIS DE INCIDENTE - Nodo: {node_id}")
    print("=" * 60)
    print(f"Periodo: {start_date} - {end_date}")
    print()
    
    resultado = {
        "node_id": node_id,
        "start_date": start_date,
        "end_date": end_date,
        "node_info": None,
        "company_info": None,
        "measures": None,
        "alerts": [],
        "hourly_data": {},
        "price_per_m3": None,
    }
    
    # 1. Obtener información del nodo
    print("[1/6] Obteniendo información del nodo...")
    node_info = obtener_informacion_nodo(node_id)
    if node_info:
        resultado["node_info"] = node_info
        print(f"  ✓ Nodo: {node_info.get('name', 'N/A')}")
        company_id = node_info.get("companyId", "").split("-")[0] if node_info.get("companyId") else None
        
        # 2. Obtener información de la empresa
        if company_id:
            print(f"[2/6] Obteniendo información de la empresa {company_id}...")
            company_info = obtener_informacion_empresa(company_id)
            if company_info:
                resultado["company_info"] = company_info
                print(f"  ✓ Empresa: {company_info.get('name', 'N/A')}")
    else:
        print("  ✗ No se pudo obtener información del nodo")
    
    # 3. Obtener medidas de consumo
    print("[3/6] Obteniendo medidas de consumo...")
    measures = obtener_medidas_consumo(node_id, start_date, end_date)
    if measures:
        resultado["measures"] = measures
        total_days = len(measures.get("month", []))
        print(f"  ✓ {total_days} días de datos obtenidos")
    else:
        print("  ✗ No se pudieron obtener medidas de consumo")
    
    # 4. Obtener alertas
    print("[4/6] Obteniendo alertas de fuga...")
    alerts = obtener_alertas(node_id, start_date, end_date)
    resultado["alerts"] = alerts
    print(f"  ✓ {len(alerts)} alertas encontradas")
    
    # 5. Obtener datos horarios para días con alertas
    if alerts:
        print("[5/6] Obteniendo datos horarios para días con alertas...")
        alert_dates = set()
        for alert in alerts:
            creation = alert.get("creationDate", "")
            if creation:
                try:
                    dt = datetime.fromisoformat(creation.replace("Z", "+00:00"))
                    date_str = dt.strftime("%d%m%Y")
                    alert_dates.add(date_str)
                except (ValueError, TypeError):
                    continue
        
        for date_str in alert_dates:
            hourly = obtener_datos_horarios(node_id, date_str)
            if hourly:
                resultado["hourly_data"][date_str] = hourly
                print(f"  ✓ Datos horarios obtenidos para {date_str}")
    
    # 6. Obtener precio del agua
    print("[6/6] Obteniendo precio del agua...")
    price = obtener_precio_agua(node_id)
    if price:
        resultado["price_per_m3"] = price
        print(f"  ✓ Precio: ${price:,.0f} CLP/m³")
    else:
        print("  ⚠ Precio no disponible")
    
    print()
    print("=" * 60)
    print("ANÁLISIS COMPLETADO")
    print("=" * 60)
    
    return resultado


def crear_carpeta_incidentes(node_id: str, node_name: Optional[str] = None) -> Path:
    """
    Crea la estructura de carpetas para reportes de incidentes.
    
    Estructura: reports/incidentes/[NOMBRE_REPORTE]_[FECHA_CREACION]/
    
    Args:
        node_id: ID del nodo
        node_name: Nombre del nodo (opcional, se usa node_id si no se proporciona)
    
    Returns:
        Path a la carpeta creada
    """
    # Directorio base: C:\Users\joseo\Desktop\wes-scripts\reports\incidentes
    base_dir = Path.home() / "Desktop" / "wes-scripts" / "reports" / "incidentes"
    
    # Crear nombre del reporte
    if node_name:
        # Limpiar nombre para usarlo como nombre de carpeta
        safe_name = "".join(c for c in node_name if c.isalnum() or c in (" ", "-", "_")).strip()
        safe_name = safe_name.replace(" ", "_")
        report_name = f"{safe_name}_{node_id}"
    else:
        report_name = f"NODO_{node_id}"
    
    # Fecha de creación en formato YYYYMMDD_HHMM
    creation_date = datetime.now(timezone.utc)
    folder_name = f"{report_name}_{creation_date.strftime('%Y%m%d_%H%M')}"
    
    # Crear la carpeta completa
    output_dir = base_dir / folder_name
    output_dir.mkdir(parents=True, exist_ok=True)
    
    return output_dir


def main():
    parser = argparse.ArgumentParser(
        description="Gestionar y analizar incidentes en el sistema WES",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  # Analizar incidente para un nodo (se guarda automáticamente en reports/incidentes/)
  python gestionar_incidentes.py --node-id 000025-20 --start-date 01112025 --end-date 07122025
  
  # Analizar incidente con nombre personalizado
  python gestionar_incidentes.py --node-id 000025-20 --start-date 01112025 --end-date 07122025 --nombre "Fuga crítica"
        """
    )
    
    parser.add_argument(
        "--node-id",
        required=True,
        help="ID del nodo (ej: 000025-20)"
    )
    
    parser.add_argument(
        "--start-date",
        required=True,
        help="Fecha inicio en formato ddMMyyyy (ej: 01112025)"
    )
    
    parser.add_argument(
        "--end-date",
        required=True,
        help="Fecha fin en formato ddMMyyyy (ej: 07122025)"
    )
    
    parser.add_argument(
        "--nombre",
        help="Nombre personalizado para el reporte (opcional)"
    )
    
    parser.add_argument(
        "--no-guardar",
        action="store_true",
        help="No guardar los resultados en archivo (solo mostrar en consola)"
    )
    
    args = parser.parse_args()
    
    # Validar formato de fechas
    try:
        datetime.strptime(args.start_date, "%d%m%Y")
        datetime.strptime(args.end_date, "%d%m%Y")
    except ValueError:
        print("[ERROR] Las fechas deben estar en formato ddMMyyyy (ej: 01112025)")
        sys.exit(1)
    
    # Analizar incidente
    resultado = analizar_incidente(args.node_id, args.start_date, args.end_date)
    
    # Crear carpeta de incidentes y guardar resultados
    if not args.no_guardar:
        # Obtener nombre del nodo si está disponible
        node_name = None
        if resultado["node_info"]:
            node_name = resultado["node_info"].get("name")
        
        # Usar nombre personalizado si se proporcionó
        if args.nombre:
            node_name = args.nombre
        
        # Crear carpeta
        output_dir = crear_carpeta_incidentes(args.node_id, node_name)
        
        # Guardar resultados en JSON
        json_path = output_dir / "resultado_analisis.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(resultado, f, indent=2, ensure_ascii=False, default=str)
        
        print(f"\n[INFO] Carpeta de incidente creada: {output_dir}")
        print(f"[INFO] Resultados guardados en: {json_path}")
    
    # Resumen
    print("\n" + "=" * 60)
    print("## RESUMEN DEL INCIDENTE")
    print("=" * 60)
    print(f"- Nodo: {args.node_id}")
    if resultado["node_info"]:
        print(f"- Nombre: {resultado['node_info'].get('name', 'N/A')}")
    if resultado["company_info"]:
        print(f"- Empresa: {resultado['company_info'].get('name', 'N/A')}")
    if resultado["measures"]:
        total_days = len(resultado["measures"].get("month", []))
        print(f"- Días con datos: {total_days}")
    print(f"- Alertas encontradas: {len(resultado['alerts'])}")
    if resultado["hourly_data"]:
        print(f"- Días con datos horarios: {len(resultado['hourly_data'])}")
    print("=" * 60)


if __name__ == "__main__":
    main()

