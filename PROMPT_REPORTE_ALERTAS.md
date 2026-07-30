# Prompt para Reportes de Consumo y Alertas

Usa este prompt cuando necesites generar un informe Word (.docx) que incluya gráficos de consumo y alertas de fuga a partir de la API WES ACL Node (`http://104.248.53.141:7003/wes/api/acl-node/v1`).

## Mapeo de IDs a Nombres de Nodos

El script incluye un diccionario que mapea los IDs de nodos a sus nombres descriptivos. Cuando se genera un reporte, el nombre del nodo se muestra en lugar del ID en el documento Word.

**Diccionario de nodos:** Ver `generar_reporte_word.py` - variable `NODE_NAMES` que contiene el mapeo completo de IDs a nombres.

```
Objetivo: Construir un informe Word con las métricas de consumo y las alertas de fuga de un nodo PIC.

Entradas obligatorias:
- CompanyId: [ID_EMPRESA]
- NodeId: [ID_PUNTO]
- Fecha inicio: [START_DATE]  (formato ddMMyyyy, ej. 01112025)
- Fecha término: [END_DATE]   (formato ddMMyyyy, ej. 26112025)

Pasos:
1. Medidas de consumo
   - Llama a GET /nodes/measures/dates con parámetros:
     id=[NodeId], start=[START_DATE], end=[END_DATE]  (todos en ddMMyyyy).
   - Si el endpoint específico GET /nodes/{id}/measures/dates estuviera disponible, úsalo como alternativa.
   - Normaliza los resultados: fecha, total m³, KPI relevantes si existen.

2. Alertas de fuga
   - Llama a GET /nodes/myalert/alerts con parámetros:
     id=[NodeId], start=[START_DATE], end=[END_DATE]  (ddMMyyyy).
   - Si falla, intenta con GET /nodes/leak/alerts usando formato ISO (YYYY-MM-DDTHH:MM:SSZ).

3. Visualizaciones
   - Genera gráfica de consumo diario (línea/área).
   - Genera gráfica de alertas (línea de tallos o barras con intensidad).

4. Documento Word
   - Portada con Empresa, Nombre del Punto (y NodeId entre paréntesis), rango y fecha de generación.
   - El nombre del punto se obtiene del diccionario NODE_NAMES usando el NodeId.
   - Resumen ejecutivo: consumo total, promedio diario, días pico, cantidad de alertas.
   - Sección “Consumo”: narrativa breve, gráfica, tabla con métricas (total, promedio, máximos/mínimos) y detalle diario (fecha, m³).
   - Sección “Alertas”: narrativa, gráfica, tabla con fecha y medida.
   - Conclusiones y recomendaciones.
   - Sección “Limitaciones” si alguna API falló (incluye el endpoint y el error HTTP).

5. Salida
   - Guarda el archivo como `Reporte_[CompanyId]_[NodeId]_[START]_[END].docx`.
   - Devuelve la ruta generada y un breve resumen de los datos obtenidos.

Notas:
- Usa librerías como requests, matplotlib y python-docx.
- Si algún endpoint responde 500 o 404, deja constancia en el informe pero continúa con el resto de la información disponible.
- Acepta fechas en ddMMyyyy de entrada y convértelos a ISO cuando un endpoint lo requiera.
```

Este prompt asegura que el reporte siempre incluya la información disponible y documente cualquier limitación del backend.

