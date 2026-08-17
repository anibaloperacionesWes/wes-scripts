### Prompt maestro - Auditoria ICCO (flujo completo)

Realiza una auditoria WES completa para el punto **{NODE_ID}** con el mismo estandar de "Auditoria ICCO".

El master de formato es el Word de abril 2026, p. ej.
`reports/reporte de auditoria/auditoria_puntos_renca_abril_2026/Auditoria ICCO Renca 000017-08/Auditoria_Colegio_ICCO_Renca_000017-08.docx`.

#### Estructura fija del Word (no cambiar el orden)

1. **Portada** — título «Informe de Auditoría», referencia del archivo, establecimiento, marca WES, autor (Aníbal Aranda Alvarado).
2. **Índice** — Metodología / Registros de consumos / Resultados y Conclusiones.
3. **Metodología** — dos periodos de N días homólogos (Con WES vs Sin WES), rejilla horaria Chile 00:00–24:00, ahorro = max(0, promedio_sin − promedio_con).
4. **Registros de consumos**
   - Barras Σ Con WES vs Sin WES.
   - Cuadro 4 columnas: Condición | Periodo | Σ rejilla (m³/h) | Promedio diario (m³/h). Encabezado azul `#D9E1F2`, Calibri 10.
   - Perfiles 24 h área+líneas por día homólogo (Lun…Dom) + promedio 24 h. Azul Con WES / rojo-marrón Sin WES.
5. **Resultados y Conclusiones** — tabla de promedios diarios, ahorro m³/día y %, párrafo de volumen evitado (CLP orientativo 1.200 CLP/m³). No mostrar ahorros negativos.

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
   - Perfil promedio 24h (leyendas con las fechas reales del Excel, no marzo/abril fijos).
6. Generar reporte Word con el formato final ya definido:
   - Nombre base del archivo: **Auditoria_<establecimiento>_<nodo>**.
   - Titulos y subtitulos con estilo ICCO.
   - Tablas con mismo color, bordes, fuente, tamano y alineaciones del master.
   - Fechas de periodo en formato compacto (ejemplo: "17 al 23 agosto 2026").
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

#### Como correrlo (scripts)

Un punto:
```
python generar_auditoria_wes_cliente.py --node-id 000017-08 --nombre "Colegio ICCO Renca" \
  --con-desde 2026-08-10 --con-hasta 2026-08-16 \
  --sin-desde 2026-08-17 --sin-hasta 2026-08-23 \
  --base "reports/reporte de auditoria/auditoria_puntos_renca_agosto_2026"
```

Los 5 puntos de Renca (o 4, excluyendo el que sigue CON control):
```
python generar_auditorias_renca_lote.py --dry-run
python generar_auditorias_renca_lote.py --excluir 000017-XX
```

Ventana agosto 2026 (7+7 lunes–domingo): Con WES 10–16 ago; Sin WES 17–23 ago (el lunes 24 vuelve el control). No generar el informe completo hasta que cierre el domingo 23 (datos incompletos antes).

---

### Ejemplo listo para usar (abril 2026, 7+7)

Realiza una auditoria WES completa para el punto **000017-08** con el mismo estandar de "Auditoria ICCO".
Periodo con control: 13-04-2026 a 19-04-2026.
Periodo sin control: 06-04-2026 a 12-04-2026.
Jornada: 00:00 a 24:00 (Chile).
Genera descarga de CSV, consolidado, graficas, Word y PDF finales.

---

### Ejemplo Renca agosto 2026 (4 de 5 puntos sin control)

Mismos 5 puntos de abril: 000017-08 ICCO, 000017-04 Lo Velásquez, 000017-06 Piscina, 000017-05 Gimnasio, 000017-07 Cumbre de cóndores.
Periodo con control: 10-08-2026 a 16-08-2026.
Periodo sin control: 17-08-2026 a 23-08-2026.
Excluir el nodo que permanece CON control (`--excluir`).
Jornada: 00:00 a 24:00 (Chile).
Salida: `reports/reporte de auditoria/auditoria_puntos_renca_agosto_2026/`.
