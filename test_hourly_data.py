"""Script de prueba para obtener datos horarios de un día específico."""

import requests
from datetime import datetime

BASE_URL = "http://104.248.53.141:7003/wes/api/acl-node/v1"

def test_hourly_data_csv(node_id: str, date_str: str):
    """Prueba obtener datos horarios usando el endpoint CSV."""
    print(f"\n{'='*60}")
    print(f"Probando endpoint CSV para nodo {node_id}, fecha {date_str}")
    print(f"{'='*60}")
    
    # Según la imagen, el endpoint correcto es /nodes/{id}/dates.measures.csv
    # con parámetros start y end (en inglés) en formato DDMMYYYY
    url = f"{BASE_URL}/nodes/{node_id}/dates.measures.csv"
    
    params = [
        ("start", date_str),
        ("end", date_str),
    ]
        
    try:
        print(f"\nIntentando: {url}")
        print(f"Parámetros: start={date_str}, end={date_str}")
        
        response = requests.get(url, params=params, timeout=30)
        print(f"Status Code: {response.status_code}")
        
        if response.status_code == 200:
            csv_content = response.text
            print(f"[OK] Exito! Respuesta recibida ({len(csv_content)} caracteres)")
            print(f"\nPrimeros 1000 caracteres de la respuesta:")
            print("-" * 60)
            print(csv_content[:1000])
            print("-" * 60)
            
            # Intentar parsear el CSV
            lines = csv_content.strip().split('\n')
            print(f"\nTotal de lineas: {len(lines)}")
            if len(lines) > 0:
                print(f"Primera linea (encabezado?): {lines[0]}")
            if len(lines) > 1:
                print(f"Segunda linea (primer dato?): {lines[1]}")
            
            # Parsear datos horarios
            # Formato: TIME,VALUE donde TIME es ISO 8601 (ej: 2025-12-05T00:00:00.000Z)
            hourly_data = []
            for i, line in enumerate(lines[1:], start=2):  # Saltar encabezado
                if not line.strip():
                    continue
                parts = line.split(',')
                if len(parts) >= 2:
                    try:
                        time_str = parts[0].strip()
                        value_str = parts[1].strip()
                        
                        # Extraer la hora del formato ISO: 2025-12-05T00:00:00.000Z
                        # Buscar el patrón THH: donde HH es la hora
                        if 'T' in time_str:
                            # Formato: 2025-12-05T00:00:00.000Z
                            hour_part = time_str.split('T')[1]  # Obtener "00:00:00.000Z"
                            hour = int(hour_part.split(':')[0])  # Extraer "00"
                        else:
                            # Si no tiene formato ISO, intentar parsear como fecha
                            from datetime import datetime
                            dt = datetime.fromisoformat(time_str.replace('Z', '+00:00'))
                            hour = dt.hour
                        
                        value = float(value_str)
                        hourly_data.append((hour, value))
                    except (ValueError, TypeError, IndexError) as e:
                        print(f"  Error parseando linea {i}: {e}")
                        continue
            
            if hourly_data:
                hourly_data.sort(key=lambda x: x[0])
                print(f"\n[OK] Datos horarios parseados: {len(hourly_data)} horas")
                print("\nDatos por hora:")
                for h, v in hourly_data:
                    print(f"  Hora {h:02d}:00 = {v:.2f} m3/hr")
                return hourly_data
            else:
                print("\n[ADVERTENCIA] No se pudieron parsear datos horarios del CSV")
            
        else:
            print(f"[ERROR] Status {response.status_code}")
            print(f"Respuesta: {response.text[:500]}")
            
    except requests.RequestException as e:
        print(f"[ERROR] Error de conexion: {e}")
    except Exception as e:
        print(f"[ERROR] Error inesperado: {e}")
        import traceback
        traceback.print_exc()
    
    return None


def test_hourly_data_json(node_id: str, date_str: str):
    """Prueba obtener datos horarios usando el endpoint JSON."""
    print(f"\n{'='*60}")
    print(f"Probando endpoint JSON para nodo {node_id}, fecha {date_str}")
    print(f"{'='*60}")
    
    url = f"{BASE_URL}/nodes/measures/dates"
    params = [
        ("id", node_id),
        ("start", date_str),
        ("end", date_str),
    ]
    
    try:
        print(f"\nIntentando: {url}")
        print(f"Parámetros: id={node_id}, start={date_str}, end={date_str}")
        
        response = requests.get(url, params=params, timeout=30)
        print(f"Status Code: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print(f"[OK] Exito! Respuesta JSON recibida")
            print(f"\nEstructura de la respuesta:")
            print(f"  Tipo: {type(data)}")
            
            if isinstance(data, list):
                print(f"  Es una lista con {len(data)} elementos")
                if len(data) > 0:
                    print(f"  Primer elemento: {list(data[0].keys()) if isinstance(data[0], dict) else data[0]}")
            elif isinstance(data, dict):
                print(f"  Es un diccionario con claves: {list(data.keys())}")
            
            # Buscar el día específico
            target_date_str = "2025-12-04"
            if isinstance(data, list):
                for item in data:
                    if isinstance(item, dict) and item.get("nodeId") == node_id:
                        data = item
                        break
            
            if isinstance(data, dict):
                month_data = data.get("month", [])
                print(f"\nDatos en 'month': {len(month_data)} dias")
                
                for node_measure in month_data:
                    date_str_measure = node_measure.get("date", "")
                    if target_date_str in date_str_measure or date_str_measure.startswith(target_date_str[:10]):
                        print(f"\n[OK] Dia encontrado: {date_str_measure}")
                        print(f"  Total diario (totalM3): {node_measure.get('totalM3', 'N/A')}")
                        
                        measures_list = node_measure.get("measures", [])
                        if measures_list:
                            print(f"  [OK] Medidas horarias encontradas: {len(measures_list)} horas")
                            hourly_data = []
                            for measure in measures_list:
                                hour_str = measure.get("hour", "")
                                measurement = measure.get("measurement", "0")
                                try:
                                    hour = int(hour_str) if hour_str else 0
                                    value = float(measurement)
                                    hourly_data.append((hour, value))
                                except (ValueError, TypeError):
                                    continue
                            
                            if hourly_data:
                                hourly_data.sort(key=lambda x: x[0])
                                print(f"\nDatos por hora:")
                                for h, v in hourly_data:
                                    print(f"  Hora {h:02d}:00 = {v:.2f} m3/hr")
                                return hourly_data
                            else:
                                print("  [ADVERTENCIA] No se pudieron parsear las medidas horarias")
                        else:
                            print("  [ADVERTENCIA] No hay campo 'measures' o esta vacio")
                            print(f"  Campos disponibles: {list(node_measure.keys())}")
            
            # Mostrar respuesta completa para debugging
            import json
            print(f"\nRespuesta completa (primeros 2000 caracteres):")
            print("-" * 60)
            print(json.dumps(data, indent=2, ensure_ascii=False)[:2000])
            print("-" * 60)
            
        else:
            print(f"[ERROR] Status {response.status_code}: {response.text[:200]}")
            
    except requests.RequestException as e:
        print(f"[ERROR] Error de conexion: {e}")
    except Exception as e:
        print(f"[ERROR] Error inesperado: {e}")
        import traceback
        traceback.print_exc()
    
    return None


if __name__ == "__main__":
    node_id = "000025-17"
    # Probar con el formato exacto de la imagen (05122025 = 05-12-2025)
    date_str = "05122025"  # Formato DDMMYYYY según la imagen
    
    print(f"\n{'#'*60}")
    print(f"PRUEBA DE OBTENCIÓN DE DATOS HORARIOS")
    print(f"Nodo: {node_id}")
    print(f"Fecha: 05-12-2025 (formato API: {date_str}) - Segun imagen")
    print(f"{'#'*60}")
    
    # También probar con rango de fechas (04-12 a 05-12)
    print(f"\nProbando tambien con rango de fechas (04-12 a 05-12)...")
    date_start = "04122025"
    date_end = "05122025"
    
    # Probar endpoint CSV primero
    csv_data = test_hourly_data_csv(node_id, date_str)
    
    # Probar también con rango de fechas usando el endpoint correcto
    if not csv_data:
        print(f"\nProbando con rango de fechas: start={date_start}, end={date_end}")
        url_range = f"{BASE_URL}/nodes/{node_id}/dates.measures.csv"
        params_range = [
            ("start", date_start),
            ("end", date_end),
        ]
        try:
            response_range = requests.get(url_range, params=params_range, timeout=30)
            print(f"Status Code (rango): {response_range.status_code}")
            if response_range.status_code == 200:
                csv_content_range = response_range.text
                print(f"[OK] Exito con rango! Respuesta recibida ({len(csv_content_range)} caracteres)")
                print(f"\nPrimeros 500 caracteres:")
                print("-" * 60)
                print(csv_content_range[:500])
                print("-" * 60)
                csv_data = True  # Marcar como exitoso
            else:
                print(f"[ERROR] Status (rango): {response_range.status_code}")
                print(f"Respuesta: {response_range.text[:500]}")
        except Exception as e:
            print(f"[ERROR] Error con rango: {e}")
    
    # Si CSV no funciona, probar JSON
    if not csv_data:
        json_data = test_hourly_data_json(node_id, date_str)
        if json_data:
            print(f"\n[OK] Datos obtenidos exitosamente desde JSON")
        else:
            print(f"\n[ERROR] No se pudieron obtener datos horarios desde ningun endpoint")
    else:
        print(f"\n[OK] Datos obtenidos exitosamente desde CSV")

