# Formulario de visitas técnicas WES

Referencia del chat / trabajo del formulario web de actas de mantención.

## Link del formulario (oficial)

https://script.google.com/a/macros/wes.cl/s/AKfycbx5DTddldWVkvK7QLxS85UbDDLQFG93_gJx3MYlL52gdPpq-FY3mRX9nxGQ7OOCbtxC/exec

## Proyecto Apps Script

https://script.google.com/home/projects/1Ivk_JUyxqbR1B5BQyEEfNZJj8yL8OUs4jOpxpWd2reLwCrSsZ_nMlWLn/edit

## Cómo se actualiza

| Archivo | Cómo llega al /exec |
|---------|---------------------|
| **Formulario** (HTML) | Codigo lo lee desde Drive en cada carga. El agente sube el TXT y con refrescar alcanza. |
| **Codigo** (.gs) | Hay que **pegar** en Apps Script + **Nueva versión** del `/exec` (no nueva implementación). |

### TXT en Drive

- Formulario: https://drive.google.com/file/d/1UVCdra_Xsvozajx-32xnAOQh4Z6rNc2C/view?usp=drivesdk  
- Codigo: https://drive.google.com/file/d/111G4RIplGujgMnJk2_x9Pgc0bzw1WUt6/view?usp=drivesdk  
- Instrucciones: https://drive.google.com/file/d/1Lnbep5y8OkwtKJhYTx8JgKzs47jS12aP/view?usp=drivesdk  

Pegar siempre desde **Descargar → Bloc de notas**, no desde vista Google Docs.

## Funciones clave en Codigo

- `procesarVisita` — envía acta (PDF + Excel + correo opcional)
- `listarOTsPendientes` — panel OT abiertas / en_curso
- `obtenerVisitaPorFolio` — continuar / cerrar OT (mismo folio)
- `marcarOTCerrada` — cierre rápido sin PDF
- `updateSheetByFolio_` — al cerrar: `getRange(fila, 1, 1, nCols)` (fix 21T)
- `loadFormularioTemplate_` — HTML desde Drive (`FORMULARIO_HTML_DRIVE_ID`)
- `getWesApiVersion` — diagnóstico de deploy

## UX del formulario (resumen)

- **Modalidad:** Visita física · Soporte técnico a distancia  
- **Tecnología:** CPA y CIR · SAB · CPA · CIR · On/Off (sin EYES)  
- **Checklist:** botones ▸ Abrir solo si hay CIR/CPA/SAB/On-Off  
- **OT abiertas:** panel arriba para continuar/cerrar  
- **Correo opcional:** desmarcar = PDF interno sin mail al cliente  

## Registro Excel

https://docs.google.com/spreadsheets/d/1GlRn7QXWEre7ziau29ojR5lTl-bZ8T3mCT3cD93HZgM/edit  

Hoja `Datos`: Folio (A) = N OT (L); Estado visita (M) = abierta / en_curso / cerrada.

## Rama Git

`cursor/formulario-tecnico-vistoso-1384` · PR #9  

Archivos fuente: `mantenimiento_wes/apps_script_export/Formulario.html`, `Codigo.gs`
