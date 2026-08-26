# Formulario de visitas técnicas WES

Estado verificado (agosto 2026). Objetivo: acta web estable, deploy Drive, OT abiertas/cierre, modalidad, Excel.

## Link del formulario (oficial)

https://script.google.com/a/macros/wes.cl/s/AKfycbx5DTddldWVkvK7QLxS85UbDDLQFG93_gJx3MYlL52gdPpq-FY3mRX9nxGQ7OOCbtxC/exec

## Proyecto Apps Script

https://script.google.com/home/projects/1Ivk_JUyxqbR1B5BQyEEfNZJj8yL8OUs4jOpxpWd2reLwCrSsZ_nMlWLn/edit

## Deploy

| Archivo | Cómo llega al /exec |
|---------|---------------------|
| **Formulario** (HTML) | `loadFormularioTemplate_()` lee Drive en cada `doGet`. Subir TXT → refrescar form. Build HTML actual: **21S**. |
| **Codigo** (.gs) | Pegar en Apps Script + **Nueva versión** del `/exec`. API actual en repo/Drive: **21T**. |

### TXT Drive

- Formulario: https://drive.google.com/file/d/1UVCdra_Xsvozajx-32xnAOQh4Z6rNc2C/view?usp=drivesdk  
- Codigo: https://drive.google.com/file/d/111G4RIplGujgMnJk2_x9Pgc0bzw1WUt6/view?usp=drivesdk  
- Instrucciones: https://drive.google.com/file/d/1Lnbep5y8OkwtKJhYTx8JgKzs47jS12aP/view?usp=drivesdk  

Pegar desde **Descargar → Bloc de notas** (no Google Docs).

### Verificar Codigo desplegado

Ctrl+F en Apps Script → Codigo:

- `procesarVisita`
- `listarOTsPendientes`
- `getWesApiVersion` (debe devolver `version: 21T`)
- `getRange(found.row, 1, 1, row.length)` ← fix cierre OT

## Funciones clave

- `procesarVisita` — PDF + Excel + correo opcional; `folio_reusar` actualiza fila (no crea folio nuevo)
- `listarOTsPendientes` / `obtenerVisitaPorFolio` / `marcarOTCerrada`
- `updateSheetByFolio_` — **numRows=1** (bug 1775 filas corregido en 21T)
- `loadFormularioTemplate_` — HTML desde `FORMULARIO_HTML_DRIVE_ID`

## UX

- **Modalidad:** Visita física · Soporte técnico a distancia  
- **Tipo mtto:** incluye **Soporte remoto** (se auto-elige con modalidad a distancia)  
- **Remoto:** sin firma en pantalla (PDF indica «sin firma presencial»)  
- **Ortografía:** spellcheck es en solución, observaciones, cargos y obs. de checklist  
- **Tecnología:** CPA y CIR · SAB · CPA · CIR · On/Off (sin EYES)  
- **Checklist:** ▸ Abrir solo con CIR/CPA/SAB/On-Off  
- **Panel OT:** Continuar/cerrar o Marcar cerrada  
- **Correo:** desmarcar = PDF interno  

## Registro Excel

https://docs.google.com/spreadsheets/d/1GlRn7QXWEre7ziau29ojR5lTl-bZ8T3mCT3cD93HZgM/edit  

Hoja `Datos`: Folio (A) = OT (L); Estado (M) = abierta / en_curso / cerrada.

## Prueba exitosa conocida

Folio **2277** · RENCA / GIMNASIO MUNICIPAL · PDF interno OK (tras pegar Codigo).

## Rama

`cursor/formulario-tecnico-vistoso-1384` · PR #9  
Fuentes: `apps_script_export/Formulario.html`, `Codigo.gs`
