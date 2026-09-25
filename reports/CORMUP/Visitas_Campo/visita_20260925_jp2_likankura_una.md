# Visita de campo CORMUP — lecturas, medidor vs cuenta y ajuste ultrasonido

**Fecha registro:** 2026-09-25  
**Cliente / red:** CORMUP (000008)  
**Motivo:** Tomar lecturas en terreno, contrastar con cuenta / WES, y ajustar sensor de ultrasonido donde aplica.  
**Hipótesis a auditar:** Es posible que en el pasado lecturas de Aguas Andinas no se hayan tomado realmente y se facture con un consumo promedio estimado; conviene seguimiento con segunda lectura.

---

## 1. Juan Pablo II (Juan Pablo Segundo)

| Campo | Valor |
|-------|--------|
| Nodo WES | `000008-14` |
| Nombre | Juan Pablo II |
| Lectura en terreno | Tomada |
| Cuenta vs medidor real | **No corresponde** — la cuenta no es la que realmente corresponde al medidor |
| Lectura vs referencia | **Nada que ver** — lectura desalineada / no conversa |

**Estado:** Discrepancia de cuenta y de lectura. Requiere corrección de asociación cuenta–medidor y revisión.

---

## 2. Likankura (Liga Ancura)

| Campo | Valor |
|-------|--------|
| Nodo WES | `000008-13` |
| Nombre | Likankura (referido en visita como “Liga Ancura”) |
| Lectura en terreno | Tomada |
| Condición | **Misma condición que Juan Pablo II** — cuenta no corresponde / lectura no conversa |

**Estado:** Misma discrepancia de cuenta/lectura. Revisar asociación y datos.

---

## 3. Unión Nacional Árabe

| Campo | Valor |
|-------|--------|
| Nodo WES | `000008-12` |
| Nombre | Unión Nacional Árabe |
| Lectura en terreno | Tomada (primera de seguimiento) |
| Lectura vs cuenta | **Conversan** — lectura y cuenta sí coinciden entre sí |
| Lectura/cuenta vs WES | **No conversan** con la información de WES |

**Acción:** Dejar registrada esta lectura como **línea base** para una **segunda lectura** y auditar / hacer seguimiento del consumo real vs lo facturado / vs WES.

**Motivo del seguimiento:** Históricamente puede haber ocurrido que lecturas de Aguas Andinas no se tomen en terreno y se estime un consumo promedio; este punto permite verificarlo con dos lecturas reales.

---

## 4. Matilde Huici Navas (Matilde Huichinawa)

| Campo | Valor |
|-------|--------|
| Nodo WES | `000008-10` |
| Nombre | Matilde Huici Navas |
| Sensor | **Sí hay sensor de ultrasonido** |
| Problema encontrado | La pulsación del medidor **no cuadraba** contra el ultrasonido |
| Causa / criterio de ajuste | Al sacar el factor de escala en el tiempo, el factor quedaba **por sobre 1,5** |
| Acción en terreno | Se **aumentó el diámetro exterior** configurado de la tubería para que cuadrara (evitar escala > 1,5) |
| Estado actual | **Está cuadrando**; diferencia del orden de **~2 litros** |
| Lecturas | Se tomará / se tomó **lectura inicial y lectura final** para el ajuste |

**Estado:** Ajuste de ultrasonido aplicado (diámetro exterior ↑). Pulsación vs ultrasonido alineados con ~2 L de diferencia. Validar con lectura inicial y final.

---

## Pendientes / seguimiento

1. Corregir o validar cuenta–medidor en **Juan Pablo II** (`000008-14`) y **Likankura** (`000008-13`).
2. Programar **segunda lectura** en **Unión Nacional Árabe** (`000008-12`) para auditar consumo entre lecturas vs factura Aguas Andinas vs WES.
3. Completar / registrar **lectura inicial y final** en **Matilde Huici Navas** (`000008-10`) tras el ajuste de diámetro/ultrasonido.
4. Documentar números de cuenta, lectura (m³), diámetro configurado y foto/medidor cuando estén disponibles (esta nota no incluye aún los valores numéricos).

---

## Nota de transcripción

Dictado desde visita: “Juan Pablo Segundo”, “Liga Ancura”, “Unión Nacional Árabe”, “West”, “Matilde Huichinawa” → interpretados como Juan Pablo II, Likankura, Unión Nacional Árabe, WES y Matilde Huici Navas, según catálogo CORMUP del proyecto.
