### Prompt maestro - Auditoria ICCO (flujo completo)

Realiza una auditoria WES completa para el punto **{NODE_ID}** con el mismo estandar de "Auditoria ICCO".

#### Datos de entrada
- Nodo: `{NODE_ID}`
- Nombre del punto/establecimiento: `{NOMBRE_PUNTO}`
- Periodo con control: `{FECHA_INI_CON}` a `{FECHA_FIN_CON}`
- Periodo sin control (linea base): `{FECHA_INI_SIN}` a `{FECHA_FIN_SIN}`
- Jornada horaria: `{HORA_INICIO}` a `{HORA_FIN}` (Chile)
- Carpeta de salida: `{CARPETA_SALIDA}`
- Excluir nodos (opcional): `{LISTA_EXCLUSION}`

#### Instrucciones obligatorias (haz todo, en este orden)
1. Descargar los CSV horarios desde API WES para ambos periodos.
2. Generar y guardar los CSV intermedios en carpeta de trabajo.
3. Consolidar ambos periodos en Excel (mismo enfoque usado en ICCO).
4. Calcular totales, promedios diarios, diferencia neta y ahorro porcentual sin mostrar ahorros negativos.
5. Generar las mismas graficas del flujo ICCO:
   - Barras comparativas Con WES vs Sin WES.
   - Perfiles horarios 24h por dia homologo.
   - Perfil promedio 24h.
6. Generar reporte Word con el formato final ya definido:
   - Nombre base del archivo: **Auditoria_ICCO**.
   - Titulos y subtitulos con estilo ICCO.
   - Tablas con mismo color, bordes, fuente, tamano y alineaciones del master.
   - Fechas de periodo en formato compacto (ejemplo: "06 al 09 abril 2026").
7. Convertir el Word a PDF.
8. Entregar rutas finales de:
   - CSV descargados
   - Excel consolidado
   - Graficas PNG
   - DOCX final
   - PDF final

#### Criterios de calidad
- Mantener exactamente la estructura y estilo del informe ICCO master.
- Si falta algun dato de API, continuar y dejar advertencia clara en el informe.
- No sobrescribir archivos manuales salvo que se indique explicitamente.
- Confirmar al final un resumen ejecutivo de 5 lineas maximo con hallazgos clave.

---

### Ejemplo listo para usar

Realiza una auditoria WES completa para el punto **000017-08** con el mismo estandar de "Auditoria ICCO".
Periodo con control: 23-03-2026 a 26-03-2026.
Periodo sin control: 06-04-2026 a 09-04-2026.
Jornada: 00:00 a 24:00 (Chile).
Genera descarga de CSV, consolidado, graficas, Word y PDF finales.
