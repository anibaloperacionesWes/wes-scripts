# Prompt — Formulario / acta de visita técnica WES

**Formulario web = acta.** Es el mismo documento oficial de visitas técnicas.
No armar un Word libre.

## Link fijo del formulario (terreno / técnicos)

https://script.google.com/a/macros/wes.cl/s/AKfycbx5DTddldWVkvK7QLxS85UbDDLQFG93_gJx3MYlL52gdPpq-FY3mRX9nxGQ7OOCbtxC/exec

Referencia en repo: `mantenimiento_wes/LINK_FIJO_FORMULARIO_WEB.txt`  
Doc del deploy: `mantenimiento_wes/FORMULARIO_VISITAS_TECNICAS.md`

## Dos formas de sacar el mismo acta

| Canal | Cuándo |
|-------|--------|
| **Formulario web** (link de arriba) | Técnico en terreno o en PC; firma en pantalla si es visita física |
| **Agente / CLI** `generar_acta_visita_tecnica.py` | Dictás los datos acá; se genera el mismo PDF oficial |

Ambos usan los mismos campos y el mismo layout PDF.

## Datos mínimos

| Campo | Ejemplo |
|-------|---------|
| Cliente | CORMUP / CDUC / UDD… |
| Máquina / sitio | ERASMO ESCALA |
| Comuna | Peñalolén |
| Fecha y hora | 2026-10-23 · 09:00 |
| Técnico WES | Anibal Aranda |
| Modalidad / motivo | Visita física · Soporte técnico a distancia |
| Tecnología | CPA y CIR · CIR · CPA · SAB · On/Off |
| Tipo mtto | Mtto Correctivo · Preventivo · Soporte remoto · Retiro CIR… |
| Tipo / falla específica | Ultrasonido · Programación de auditoría… |
| Solución / diagnóstico | Texto libre |
| Observaciones | Seguimiento, etc. |
| Recibido por + cargo | Nombre y cargo |
| Email cliente (si se envía) | Solo destinatarios del cliente |

Folio/OT: máximo del Registro de fallas (Sheet `1GlRn7QXWEre7ziau29ojR5lTl-bZ8T3mCT3cD93HZgM`, hoja Datos) **+ 1**.

## CLI (mismo PDF que el formulario)

```bash
python generar_acta_visita_tecnica.py --json visita.json
python generar_acta_visita_tecnica.py --json visita.json --subir-drive
python generar_acta_visita_tecnica.py --json visita.json --subir-drive --enviar-correo
```

Plantilla: `mantenimiento_wes/ejemplos/acta_ejemplo.json`

Drive: **Actas de Mantencion** → `{Cliente}/{año}/{mes}`  
Correo: solo cliente; CORMUP → `dcampos.henriquez@cormup.cl`. Pedir «Acuso recibo».

## Pedido típico al agente

```
Formulario/acta de visita:

- Cliente: CORMUP
- Sitio: ERASMO ESCALA
- Comuna: Peñalolén
- Fecha: 23/10/2026 09:00
- Técnico: Anibal Aranda
- Modalidad: Soporte remoto
- Folio: 2342
- Qué se hizo: …
- Observaciones: …
- Recibido por: Elizabeth González (Directora)
- Enviar a: dcampos.henriquez@cormup.cl
```
