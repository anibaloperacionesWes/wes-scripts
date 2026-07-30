# Prompt para Gestión de Incidentes - WES API

Este prompt proporciona acceso completo a todas las APIs del sistema WES para la gestión y análisis de incidentes relacionados con consumo de agua, fugas y alertas.

## 📁 Estructura de Carpetas para Reportes

**IMPORTANTE:** Todos los reportes de incidentes deben guardarse en:

```
C:\Users\joseo\Desktop\wes-scripts\reports\incidentes\[NOMBRE_REPORTE]_[FECHA_CREACION]\
```

Donde:
- `[NOMBRE_REPORTE]`: Nombre descriptivo del incidente o nodo analizado
- `[FECHA_CREACION]`: Fecha y hora de creación en formato `YYYYMMDD_HHMM`

**Ejemplo:**
```
C:\Users\joseo\Desktop\wes-scripts\reports\incidentes\NODO_000025-20_20251205_1430\
```

La carpeta `incidentes` se crea automáticamente si no existe.

## 🔍 Reporte de Incidentes del Día Anterior

El script `reporte_incidentes_dia_anterior.py` genera un reporte Word que analiza **todos los puntos** del sistema y reporta:

1. **Puntos con consumo cero**: Puntos que registraron consumo cero durante todo el día anterior
2. **Puntos sin respuesta**: Puntos que no respondieron a la consulta de datos (error de API o sin conexión)
3. **Puntos con datos incompletos**: Puntos que tienen menos de 24 horas de datos, indicando qué horas faltan

### Uso del Script

```powershell
python reporte_incidentes_dia_anterior.py
```

El script:
- Obtiene automáticamente todos los nodos del sistema (iterando por empresas 000000-000100)
- Para cada nodo, consulta los datos horarios del día anterior usando el endpoint `/nodes/{id}/dates.measures.csv`
- Analiza los datos y clasifica cada punto
- Genera un documento Word con tablas detalladas de cada categoría
- Guarda el reporte en `reports/incidentes/REPORTE_DIA_ANTERIOR_[FECHA]/`

### Estructura del Reporte Word

El reporte incluye:
- **Título**: "REPORTE DE INCIDENTES - CONSUMO DÍA ANTERIOR"
- **Fecha analizada**: Día anterior a la generación del reporte
- **Resumen ejecutivo**: Totales de cada categoría
- **Tabla de puntos con consumo cero**: Nodo ID, Nombre, Empresa
- **Tabla de puntos sin respuesta**: Nodo ID, Nombre, Empresa
- **Tabla de puntos con datos incompletos**: Nodo ID, Nombre, Empresa, Horas faltantes

## 📡 APIs Disponibles

### 1. ACL Node API
**Base URL:** `http://104.248.53.141:7003/wes/api/acl-node/v1`

Endpoints disponibles:
- `GET /nodes/measures/dates` - Obtener medidas de consumo por rango de fechas
- `GET /nodes/{id}/dates.measures.csv` - Obtener medidas horarias en formato CSV
- `GET /nodes/myalert/alerts` - Obtener alertas de fuga (formato ddMMyyyy)
- `GET /nodes/leak/alerts` - Obtener alertas de fuga (formato ISO)
- `GET /nodes/{id}` - Obtener información de un nodo específico
- `GET /nodes?id=...` - Buscar nodos por ID

### 2. ACL Entities API
**Base URL:** `http://104.248.53.141:7001/wes/api/acl-entities/v1`

Endpoints disponibles:
- `GET /companies/{companyId}` - Obtener información de una empresa
- `GET /companies/{companyId}/nodes` - Obtener todos los nodos de una empresa
- `GET /nodes/{nodeIdBase}` - Obtener información de un nodo (sin guion)
- `POST /configuration/users` - Crear usuarios en el sistema

## 🔧 Funciones de Acceso a APIs

### Función Base para Llamadas HTTP

```python
import requests
from typing import Optional, Union, Sequence

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
        print(f"Error al consultar {url}: {e}")
        raise
```

## 📋 Endpoints Detallados

### 1. Obtener Medidas de Consumo

**Endpoint:** `GET /nodes/measures/dates`

**Parámetros:**
- `id` (string): ID del nodo (ej: "000025-20")
- `start` (string): Fecha inicio en formato `ddMMyyyy` (ej: "01112025")
- `end` (string): Fecha fin en formato `ddMMyyyy` (ej: "07122025")

**Ejemplo:**
```python
BASE_URL = "http://104.248.53.141:7003/wes/api/acl-node/v1"

data = fetch_json(
    f"{BASE_URL}/nodes/measures/dates",
    params=[
        ("id", "000025-20"),
        ("start", "01112025"),
        ("end", "07122025"),
    ],
)
```

**Respuesta esperada:**
```json
{
    "month": [
        {
            "date": "2025-11-01",
            "totalM3": 123.45,
            "measures": [
                {
                    "hour": "00",
                    "measurement": "5.12"
                },
                ...
            ]
        },
        ...
    ]
}
```

### 2. Obtener Medidas Horarias (CSV)

**Endpoint:** `GET /nodes/{id}/dates.measures.csv`

**Parámetros:**
- `start` (string): Fecha inicio en formato `ddMMyyyy`
- `end` (string): Fecha fin en formato `ddMMyyyy`

**Ejemplo:**
```python
import requests

node_id = "000025-20"
date_str = "04122025"  # 04-12-2025

url = f"{BASE_URL}/nodes/{node_id}/dates.measures.csv"
params = [
    ("start", date_str),
    ("end", date_str),
]

response = requests.get(url, params=params, timeout=30)
csv_content = response.text

# Parsear CSV
# Formato: TIME,VALUE
# TIME: ISO 8601 (ej: 2025-12-05T00:00:00.000Z)
# VALUE: valor numérico
```

### 3. Obtener Alertas de Fuga

**Endpoint Principal:** `GET /nodes/myalert/alerts`

**Parámetros:**
- `id` (string): ID del nodo
- `start` (string): Fecha inicio en formato `ddMMyyyy`
- `end` (string): Fecha fin en formato `ddMMyyyy`

**Ejemplo:**
```python
alerts = fetch_json(
    f"{BASE_URL}/nodes/myalert/alerts",
    params=[
        ("id", "000025-20"),
        ("start", "01112025"),
        ("end", "07122025"),
    ],
)
```

**Endpoint Alternativo:** `GET /nodes/leak/alerts`

**Nota:** Este endpoint puede requerir formato ISO (`YYYY-MM-DDTHH:MM:SSZ`) o `ddMMyyyy` según la versión.

**Respuesta esperada:**
```json
[
    {
        "creationDate": "2025-11-15T10:30:00.000Z",
        "measure": "2.5",
        "nodeId": "000025-20",
        ...
    },
    ...
]
```

### 4. Obtener Información de Nodo

**Endpoint:** `GET /nodes/{id}`

**Ejemplo:**
```python
node_data = fetch_json(f"{BASE_URL}/nodes/000025-20")
```

### 5. Obtener Información de Empresa

**Endpoint:** `GET /companies/{companyId}`

**Ejemplo:**
```python
ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"

company_data = fetch_json(f"{ENTITY_BASE_URL}/companies/000025")
```

**Respuesta incluye:**
- Información de la empresa
- Lista de nodos asociados (`nodes` array)

### 6. Obtener Nodos de una Empresa

**Endpoint:** `GET /companies/{companyId}/nodes`

**Ejemplo:**
```python
nodes = fetch_json(f"{ENTITY_BASE_URL}/companies/000025/nodes")
```

### 7. Obtener Precio del Agua

**Endpoint:** `GET /nodes/{nodeIdBase}` (sin guion en el ID)

**Ejemplo:**
```python
# Para nodo "000025-20", usar "000025" como nodeIdBase
node_id_base = "000025"
entity_data = fetch_json(f"{ENTITY_BASE_URL}/nodes/{node_id_base}")
```

**Nota:** El precio del agua puede estar en diferentes campos:
- `amount` (puede haber múltiples, usar el más reciente según fecha)
- `nodeKpi.expenses` (calcular: expenses / consumo total)

## 🎯 Casos de Uso para Incidentes

### Caso 1: Analizar Incidente de Fuga

```python
# 1. Obtener alertas del periodo
alerts = fetch_json(
    f"{BASE_URL}/nodes/myalert/alerts",
    params=[
        ("id", node_id),
        ("start", start_date),
        ("end", end_date),
    ],
)

# 2. Obtener medidas de consumo del mismo periodo
measures = fetch_json(
    f"{BASE_URL}/nodes/measures/dates",
    params=[
        ("id", node_id),
        ("start", start_date),
        ("end", end_date),
    ],
)

# 3. Obtener datos horarios del día con mayor alerta
# (usar el endpoint CSV para datos detallados)
```

### Caso 2: Comparar Múltiples Puntos

```python
# 1. Obtener todos los nodos de una empresa
company_id = "000025"
nodes = fetch_json(f"{ENTITY_BASE_URL}/companies/{company_id}/nodes")

# 2. Para cada nodo, obtener medidas y alertas
for node in nodes:
    node_id = node.get("nodeId")
    # Obtener medidas y alertas...
```

### Caso 3: Obtener Información Completa de un Incidente

```python
# 1. Información del nodo
node_info = fetch_json(f"{BASE_URL}/nodes/{node_id}")

# 2. Información de la empresa
company_id = node_info.get("companyId")
company_info = fetch_json(f"{ENTITY_BASE_URL}/companies/{company_id}")

# 3. Medidas de consumo
measures = fetch_json(...)

# 4. Alertas
alerts = fetch_json(...)

# 5. Precio del agua
price_data = fetch_json(f"{ENTITY_BASE_URL}/nodes/{company_id}")
```

## 📝 Formatos de Fecha

### Formato ddMMyyyy (Recomendado)
- Ejemplo: `"01112025"` = 01 de noviembre 2025
- Usado en: `/nodes/measures/dates`, `/nodes/myalert/alerts`

### Formato ISO 8601
- Ejemplo: `"2025-11-01T00:00:00.000Z"`
- Usado en: `/nodes/leak/alerts` (alternativo)

### Conversión entre formatos

```python
from datetime import datetime

# ddMMyyyy a datetime
date_str = "01112025"
dt = datetime.strptime(date_str, "%d%m%Y")

# datetime a ddMMyyyy
formatted = dt.strftime("%d%m%Y")

# datetime a ISO
iso_str = dt.isoformat() + "Z"
```

## 🔍 Manejo de Errores

### Estrategia de Fallback

```python
def get_alerts_with_fallback(node_id: str, start_date: str, end_date: str):
    """
    Intenta obtener alertas con múltiples endpoints y formatos.
    """
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
        # Convertir fechas a ISO
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
```

## 📊 Estructura de Datos Común

### Medida de Consumo
```json
{
    "date": "2025-11-01",
    "totalM3": 123.45,
    "measures": [
        {
            "hour": "00",
            "measurement": "5.12"
        }
    ]
}
```

### Alerta de Fuga
```json
{
    "creationDate": "2025-11-15T10:30:00.000Z",
    "measure": "2.5",
    "nodeId": "000025-20",
    "stream": "..."
}
```

### Información de Nodo
```json
{
    "nodeId": "000025-20",
    "companyId": "000025",
    "name": "Nombre del Nodo",
    ...
}
```

### Información de Empresa
```json
{
    "companyId": "000025",
    "name": "Parque Arauco",
    "nodes": [
        {
            "nodeId": "000025-20",
            ...
        }
    ]
}
```

## 🚀 Ejemplo Completo: Análisis de Incidente

```python
import requests
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional, Union, Sequence
import json

BASE_URL = "http://104.248.53.141:7003/wes/api/acl-node/v1"
ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"

def fetch_json(url: str, params: Optional[Union[dict, Sequence[tuple]]] = None) -> Union[dict, list]:
    response = requests.get(url, params=params, timeout=30)
    response.raise_for_status()
    return response.json()

def crear_carpeta_incidentes(node_id: str, node_name: Optional[str] = None) -> Path:
    """
    Crea la estructura de carpetas para reportes de incidentes.
    Estructura: C:\Users\joseo\Desktop\wes-scripts\reports\incidentes\[NOMBRE]_[FECHA]/
    """
    base_dir = Path.home() / "Desktop" / "wes-scripts" / "reports" / "incidentes"
    
    if node_name:
        safe_name = "".join(c for c in node_name if c.isalnum() or c in (" ", "-", "_")).strip()
        safe_name = safe_name.replace(" ", "_")
        report_name = f"{safe_name}_{node_id}"
    else:
        report_name = f"NODO_{node_id}"
    
    creation_date = datetime.now(timezone.utc)
    folder_name = f"{report_name}_{creation_date.strftime('%Y%m%d_%H%M')}"
    
    output_dir = base_dir / folder_name
    output_dir.mkdir(parents=True, exist_ok=True)
    
    return output_dir

def analizar_incidente(node_id: str, start_date: str, end_date: str):
    """
    Analiza un incidente completo para un nodo en un periodo.
    """
    print(f"Analizando incidente para nodo {node_id}")
    print(f"Periodo: {start_date} - {end_date}")
    
    # 1. Obtener información del nodo
    try:
        node_info = fetch_json(f"{BASE_URL}/nodes/{node_id}")
        print(f"✓ Información del nodo obtenida")
    except Exception as e:
        print(f"✗ Error obteniendo información del nodo: {e}")
        node_info = None
    
    # 2. Obtener medidas de consumo
    try:
        measures = fetch_json(
            f"{BASE_URL}/nodes/measures/dates",
            params=[("id", node_id), ("start", start_date), ("end", end_date)],
        )
        print(f"✓ Medidas de consumo obtenidas")
    except Exception as e:
        print(f"✗ Error obteniendo medidas: {e}")
        measures = None
    
    # 3. Obtener alertas
    alerts = []
    try:
        alerts = fetch_json(
            f"{BASE_URL}/nodes/myalert/alerts",
            params=[("id", node_id), ("start", start_date), ("end", end_date)],
        )
        if isinstance(alerts, list):
            print(f"✓ {len(alerts)} alertas obtenidas")
    except Exception:
        try:
            alerts = fetch_json(
                f"{BASE_URL}/nodes/leak/alerts",
                params=[("id", node_id), ("start", start_date), ("end", end_date)],
            )
            if isinstance(alerts, list):
                print(f"✓ {len(alerts)} alertas obtenidas (endpoint alternativo)")
        except Exception as e:
            print(f"✗ Error obteniendo alertas: {e}")
    
    # 4. Obtener precio del agua
    price_data = None
    if node_info:
        company_id = node_info.get("companyId", "").split("-")[0]
        try:
            price_data = fetch_json(f"{ENTITY_BASE_URL}/nodes/{company_id}")
            print(f"✓ Información de precio obtenida")
        except Exception as e:
            print(f"✗ Error obteniendo precio: {e}")
    
    resultado = {
        "node_info": node_info,
        "measures": measures,
        "alerts": alerts,
        "price_data": price_data,
    }
    
    # 5. Crear carpeta y guardar resultados
    node_name = node_info.get("name") if node_info else None
    output_dir = crear_carpeta_incidentes(node_id, node_name)
    
    json_path = output_dir / "resultado_analisis.json"
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(resultado, f, indent=2, ensure_ascii=False, default=str)
    
    print(f"✓ Resultados guardados en: {output_dir}")
    
    return resultado

# Uso
resultado = analizar_incidente("000025-20", "01112025", "07122025")
```

## 📚 Referencias

- **Script Principal:** `generar_reporte_word.py` - Contiene todas las funciones de acceso a APIs
- **Documentación de Usuarios:** `PROMPT_CREAR_USUARIO.md`
- **Documentación de Reportes:** `PROMPT_REPORTE_ALERTAS.md`

## ✅ Checklist para Trabajar con Incidentes

### Para Análisis de Incidente Individual:
- [ ] Identificar el `node_id` del punto afectado
- [ ] Obtener `company_id` (puede extraerse del `node_id`)
- [ ] Definir rango de fechas en formato `ddMMyyyy`
- [ ] Obtener medidas de consumo del periodo
- [ ] Obtener alertas de fuga del periodo
- [ ] Obtener datos horarios si se necesita detalle
- [ ] Obtener información del nodo y empresa
- [ ] Obtener precio del agua para valorización
- [ ] Implementar manejo de errores con fallbacks
- [ ] Validar y normalizar datos recibidos
- [ ] **Crear carpeta en `reports/incidentes/` con nombre y fecha de creación**
- [ ] **Guardar todos los reportes y resultados en la carpeta creada**

### Para Reporte del Día Anterior:
- [ ] Obtener todos los nodos del sistema (iterar empresas 000000-000100)
- [ ] Para cada nodo, obtener datos horarios del día anterior usando `/nodes/{id}/dates.measures.csv`
- [ ] Clasificar puntos en:
  - [ ] Consumo cero (todos los valores = 0)
  - [ ] Sin respuesta (error de API o sin conexión)
  - [ ] Datos incompletos (menos de 24 horas, indicar horas faltantes)
- [ ] Generar documento Word con tablas para cada categoría
- [ ] Guardar en `reports/incidentes/REPORTE_DIA_ANTERIOR_[FECHA]/`

---

*Este prompt está diseñado para trabajar con incidentes relacionados con consumo de agua, fugas y alertas en el sistema WES*

