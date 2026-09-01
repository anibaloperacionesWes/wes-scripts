# Condición Particular: Fundo Zapallar

## Contexto

**Fundo Zapallar** (Empresa ID: `000027`) no es un conjunto de puntos independientes: hay una jerarquía hidráulica. Los caudales aguas abajo **no se suman** al consumo del fundo. El consumo real (entrada / referencia de facturación) es la **Matriz ESVAL**.

## Circuito hidráulico (base de análisis)

```
Matriz ESVAL (000027-01)
        │
        ├──► Estanque Inferior (000027-02)
        │            │
        │            └──► Estanque Superior (sin medidor WES)
        │                         │
        │                         ├──► matriz Etapa N°5          → 000027-03
        │                         └──► matriz Etapa N°1 al 4     → 000027-04
        │                                      │
        │                                      ├── Etapa N°1     → 000027-06
        │                                      ├── Etapa N°2     → 000027-07
        │                                      └── Etapa N°3     → 000027-08
        │                                      (Etapa N°4 no tiene medidor propio)
        │
        └──► Riego / Riego Llenado ESVAL (000027-05 excluido; 000027-09 en reportes)
```

- ESVAL alimenta **estanque inferior** y **riego**.
- Del estanque inferior se **bombea/carga** el estanque superior.
- En la **salida del estanque superior** hay **dos matrices**:
  1. una reparte a **Etapa N°5**;
  2. otra alimenta **Etapas 1 al 4**.
- Etapas **1, 2 y 3** son **submediciones** de la matriz **Etapa N°1 al 4**. No hay medidor de Etapa N°4.

## Invariantes (si la topología y los medidores son correctos)

1. **Consumo del fundo** = Matriz ESVAL. No sumar estanques ni etapas.
2. `Etapa N°1 al 4  ≥  Etapa 1 + Etapa 2 + Etapa 3`  
   (la matriz madre no puede medir menos que la suma de sus ramales; Etapa 4 no está submedida).
3. En el tiempo, el volumen que sale del estanque inferior hacia el superior debe **poder explicar** `Etapa 5 + Etapa 1–4` (con holgura de almacenamiento en ambos estanques). Un ramal (p. ej. Etapa 3) **no puede** superar de forma sostenida ni a su matriz madre ni a la entrada ESVAL/estanque inferior.

Si un invariante se rompe, el primer sospechoso es **error de medición / factor de pulso / punto de instalación**, no un consumo real adicional.

## Puntos WES

| Nodo | Nombre | Rol |
|------|--------|-----|
| `000027-01` | Matriz ESVAL | Fuente / entrada al fundo |
| `000027-02` | Estanque Inferior | Carga desde ESVAL; alimenta estanque superior |
| `000027-03` | Etapa N°5 | Matriz desde estanque superior |
| `000027-04` | Etapa N°1 al 4 | Matriz desde estanque superior (madre de 1/2/3) |
| `000027-05` | Riego | Excluido de reportes (`EXCLUDED_NODE_IDS`) |
| `000027-06` | Etapa N°1 | Submedición de 000027-04 |
| `000027-07` | Etapa N°2 | Submedición de 000027-04 |
| `000027-08` | Etapa N°3 | Submedición de 000027-04 |
| `000027-09` | Riego Llenado de Estanque ESVAL | Riego asociado a llenado ESVAL |

`FUNDO_ZAPALLAR_NODE_IDS` en `exclusiones_reportes.py` lista los 8 puntos que entran a agregados (sin `000027-05`).

## Reportes agregados

- Total del fundo = consumo de **Matriz ESVAL**, no la suma de barras.
- Estanques y etapas son mediciones **aguas abajo** (doble conteo si se suman).
- Ante picos en un ramal (p. ej. Etapa 3), contrastar siempre contra ESVAL, estanque inferior y la matriz madre 1–4. Ver script `analizar_pico_zapallar_agosto.py`.

---

*Actualizado: septiembre 2026 — circuito ESVAL → inferior → superior → dos matrices (N5 y 1–4), con 1/2/3 como submediciones.*
