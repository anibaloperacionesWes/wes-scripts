# Condición Particular: Fundo Zapallar

## 📋 Contexto Especial

**Fundo Zapallar** (Empresa ID: 000027) tiene una configuración especial en su sistema de distribución de agua que debe considerarse al generar reportes agregados.

## 🔄 Configuración del Sistema

### Punto Fuente de Agua
- **Nodo ID**: `000027-01`
- **Nombre**: Matriz ESVAL
- **Función**: Este punto es la **fuente principal de agua** que alimenta a todos los demás puntos del sistema.

### Puntos Alimentados (Consumidores)
Los siguientes puntos reciben agua de la Matriz ESVAL:

1. **Nodo ID**: `000027-02` - Estanque Inferior
2. **Nodo ID**: `000027-03` - Etapa N°5
3. **Nodo ID**: `000027-04` - Etapa N°1 al 4
4. **Nodo ID**: `000027-05` - Riego

## 📊 Análisis Requerido en Reportes Agregados

Cuando se genere un **reporte agregado** para Fundo Zapallar, se debe considerar:

### 1. Análisis de Balance Hídrico
- **Entrada de agua**: Consumo total de Matriz ESVAL (000027-01)
- **Salidas de agua**: Suma de consumo de todos los puntos alimentados
- **Balance**: Comparar entrada vs salidas para detectar pérdidas o fugas en el sistema

### 2. Distribución del Agua
- Calcular qué porcentaje del agua de ESVAL va a cada punto alimentado
- Identificar patrones de distribución
- Detectar desbalances o anomalías en la distribución

### 3. Eficiencia del Sistema
- Analizar si hay pérdidas significativas entre la fuente y los consumidores
- Identificar puntos con mayor consumo relativo
- Evaluar la eficiencia de la distribución

### 4. Visualizaciones Especiales
- Gráfico de flujo: ESVAL → Puntos alimentados
- Gráfico de balance: Entrada vs Salidas
- Gráfico de distribución porcentual por punto
- Comparación temporal de la distribución

## ⚠️ Importante

Esta condición particular **SOLO aplica para Fundo Zapallar**. Para otras empresas, usar el reporte agregado estándar.

## 📝 Notas Técnicas

- El nodo ESVAL (000027-01) debe tratarse como **fuente** en los cálculos
- Los demás nodos (000027-02 a 000027-05) son **consumidores**
- El análisis debe mostrar claramente la relación fuente-consumidores
- Las métricas deben reflejar esta jerarquía del sistema

---

*Última actualización: Diciembre 2025*














