/**
 * Formulario permanente de acta de visita WES (Apps Script Web App).
 *
 * Después de Implementar → Aplicación web (acceso: Cualquiera),
 * el link /exec queda fijo para técnicos.
 */

var SHEET_REGISTRO_ID = '1GlRn7QXWEre7ziau29ojR5lTl-bZ8T3mCT3cD93HZgM';
var SHEET_DATOS = 'Datos';
var CARPETA_ACTAS = 'Actas_visita_PDF';
var CARPETA_TECNICOS = 'Tecnicos_WES_Formulario';
/** Carpeta Drive fija del proyecto (evita searchFolders / duplicados). */
var CARPETA_TECNICOS_ID = '1RCtWP1hK4fKzjgjyvzzSbttWJZiNhtKC';
var CC_DEFAULT = 'anibal.aoperaciones@wes.cl';
var FOLIO_INICIAL = 2250;

function doGet() {
  var tpl = HtmlService.createTemplateFromFile('Formulario');
  tpl.CATALOGOS_JSON = HtmlService.createHtmlOutputFromFile('catalogos').getContent();
  // No tocar SpreadsheetApp acá: si falla el permiso, la página queda en blanco tras el login.
  var folioShow = FOLIO_INICIAL;
  try {
    var props = PropertiesService.getScriptProperties();
    var n = Number(props.getProperty('NEXT_FOLIO') || FOLIO_INICIAL);
    if (!isNaN(n) && n >= FOLIO_INICIAL) folioShow = n;
  } catch (e) {}
  tpl.PROXIMO_FOLIO = String(folioShow);
  return tpl
    .evaluate()
    .setTitle('Acta de visita WES')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getCatalogos() {
  return JSON.parse(HtmlService.createHtmlOutputFromFile('catalogos').getContent());
}

function getProximoFolio() {
  return peekProximoFolio_();
}

/**
 * Llamado desde el HTML con google.script.run.procesarVisita(payload)
 */
function procesarVisita(data) {
  if (!data || !data.cliente || !data.maquina) {
    throw new Error('Cliente y máquina son obligatorios');
  }
  if (!data.email_cliente) {
    throw new Error('Correo del cliente es obligatorio');
  }
  if (!data.solucion) {
    throw new Error('Solución / diagnóstico es obligatorio');
  }
  if (!data.firma_png) {
    throw new Error('Firma obligatoria');
  }

  var folio = asignarFolio_();
  data = Object.assign({}, data, { folio: folio, ot: String(folio) });

  var stamp = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyyMMdd_HHmmss');
  var stem = sanitizar_(
    'folio_' + folio + '_' + (data.fecha || '') + '_' + data.cliente + '_' + data.maquina + '_' + stamp
  );

  var carpeta = asegurarCarpetaActas_();
  var firmaFile = guardarFirma_(carpeta, stem, data.firma_png);
  var pdfFile = generarYGuardarPdf_(carpeta, stem, data, firmaFile);
  var row = appendSheet_(data, pdfFile.getUrl());
  var emailInfo = enviarCorreo_(data, pdfFile);

  return {
    ok: true,
    folio: folio,
    excel_row: row,
    pdf_url: pdfFile.getUrl(),
    drive_link: pdfFile.getUrl(),
    firma_url: firmaFile.getUrl(),
    email_ok: emailInfo.ok,
    email_to: emailInfo.to,
    email_skip: emailInfo.skip || '',
    message: emailInfo.ok
      ? 'Folio ' + folio + ' · PDF generado y correo enviado a ' + emailInfo.to.join(', ')
      : 'Folio ' + folio + ' · PDF generado. Correo: ' + (emailInfo.skip || 'pendiente'),
  };
}

/** Mira el próximo folio sin consumirlo (para mostrar en pantalla). */
function peekProximoFolio_() {
  asegurarHeaderFolio_();
  var maxSheet = maxFolioEnSheet_();
  var props = PropertiesService.getScriptProperties();
  var nextProp = Number(props.getProperty('NEXT_FOLIO') || FOLIO_INICIAL);
  if (isNaN(nextProp) || nextProp < FOLIO_INICIAL) nextProp = FOLIO_INICIAL;
  return Math.max(FOLIO_INICIAL, maxSheet + 1, nextProp);
}

/** Asigna y reserva el próximo folio (thread-safe). */
function asignarFolio_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var folio = peekProximoFolio_();
    PropertiesService.getScriptProperties().setProperty('NEXT_FOLIO', String(folio + 1));
    return folio;
  } finally {
    lock.releaseLock();
  }
}

function maxFolioEnSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_REGISTRO_ID);
  var sh = ss.getSheetByName(SHEET_DATOS) || ss.getSheets()[0];
  var last = sh.getLastRow();
  if (last < 2) return FOLIO_INICIAL - 1;
  var vals = sh.getRange(2, 1, last, 1).getValues();
  var max = FOLIO_INICIAL - 1;
  for (var i = 0; i < vals.length; i++) {
    var n = Number(vals[i][0]);
    if (!isNaN(n) && n > max) max = n;
  }
  return max;
}

function asegurarHeaderFolio_() {
  var ss = SpreadsheetApp.openById(SHEET_REGISTRO_ID);
  var sh = ss.getSheetByName(SHEET_DATOS) || ss.getSheets()[0];
  if (String(sh.getRange(1, 1).getValue() || '').trim() === '') {
    sh.getRange(1, 1).setValue('Folio');
  }
}

function sanitizar_(name) {
  return String(name || 'archivo')
    .replace(/[^\w.\-áéíóúÁÉÍÓÚñÑ]+/g, '_')
    .substring(0, 120);
}

function asegurarCarpetaActas_() {
  var parent;
  try {
    parent = DriveApp.getFolderById(CARPETA_TECNICOS_ID);
  } catch (e) {
    parent = asegurarCarpetaPorNombre_(CARPETA_TECNICOS, null);
  }
  return asegurarCarpetaPorNombre_(CARPETA_ACTAS, parent.getId());
}

function asegurarCarpetaPorNombre_(nombre, parentId) {
  // No usar DriveApp.searchFolders con mimeType/trashed: lanza
  // "Exception: Argumento no válido: q". getFoldersByName es el API correcto.
  var found;
  if (parentId) {
    found = DriveApp.getFolderById(parentId).getFoldersByName(nombre);
  } else {
    found = DriveApp.getFoldersByName(nombre);
  }
  if (found.hasNext()) {
    return found.next();
  }
  if (parentId) {
    return DriveApp.getFolderById(parentId).createFolder(nombre);
  }
  return DriveApp.createFolder(nombre);
}

function guardarFirma_(carpeta, stem, dataUrl) {
  var parts = String(dataUrl).split(',');
  if (parts.length < 2) {
    throw new Error('Firma inválida');
  }
  var bytes = Utilities.base64Decode(parts[1]);
  var blob = Utilities.newBlob(bytes, 'image/png', stem + '_firma.png');
  return carpeta.createFile(blob);
}

function generarYGuardarPdf_(carpeta, stem, data, firmaFile) {
  var doc = DocumentApp.create('Acta_' + stem);
  var body = doc.getBody();
  body.clear();
  body.appendParagraph('WES · Acta de visita técnica').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Folio / OT: ' + (data.folio || data.ot || '—'));
  body.appendParagraph('Cliente: ' + (data.cliente || ''));
  body.appendParagraph('Máquina / sitio: ' + (data.maquina || ''));
  body.appendParagraph('Comuna: ' + (data.comuna || ''));
  body.appendParagraph('Fecha: ' + (data.fecha || '') + '  Hora: ' + (data.hora || ''));
  body.appendParagraph('Técnico: ' + (data.tecnico || ''));
  body.appendParagraph('Motivos: ' + joinList_(data.motivos));
  body.appendParagraph('Tecnologías: ' + joinList_(data.tecnologias));
  body.appendParagraph('Tipo mtto: ' + (data.tipo_mtto || ''));
  body.appendParagraph('Tipo falla: ' + (data.tipo_falla || ''));
  body.appendParagraph('Falla específica: ' + (data.falla_especifica || ''));
  body.appendParagraph('Solución / diagnóstico:').setBold(true);
  body.appendParagraph(data.solucion || '');
  body.appendParagraph('Observaciones: ' + (data.observaciones || ''));
  body.appendParagraph('Estado visita: ' + (data.estado_visita || ''));
  body.appendParagraph('Lectura medidor: ' + (data.lectura_medidor || ''));
  body.appendParagraph('Recibido por: ' + (data.recibido_por || '') + ' (' + (data.cargo || '') + ')');
  body.appendParagraph('Email cliente: ' + (data.email_cliente || ''));
  appendChecklist_(body, 'Checklist CIR', data.checklist_cir);
  appendChecklist_(body, 'Checklist CPA', data.checklist_cpa);
  appendChecklist_(body, 'Checklist SAB', data.checklist_sab);
  body.appendParagraph('Firma del receptor:').setBold(true);
  try {
    body.appendImage(firmaFile.getBlob()).setWidth(280);
  } catch (e) {
    body.appendParagraph('(No se pudo incrustar la firma en el PDF)');
  }
  doc.saveAndClose();

  var file = DriveApp.getFileById(doc.getId());
  var pdfBlob = file.getAs('application/pdf').setName(stem + '.pdf');
  var pdfFile = carpeta.createFile(pdfBlob);
  // el Doc intermedio se puede dejar en la misma carpeta o trash
  carpeta.addFile(file);
  file.setTrashed(true);
  return pdfFile;
}

function joinList_(arr) {
  if (!arr || !arr.length) return '—';
  return arr.join(', ');
}

function appendChecklist_(body, titulo, items) {
  body.appendParagraph(titulo).setBold(true);
  if (!items || !items.length) {
    body.appendParagraph('—');
    return;
  }
  for (var i = 0; i < items.length; i++) {
    var it = items[i] || {};
    body.appendParagraph(
      '- ' + (it.elemento || '') + ': ' + (it.estado || '') + (it.obs ? ' (' + it.obs + ')' : '')
    );
  }
}

function appendSheet_(data, pdfLink) {
  asegurarHeaderFolio_();
  var ss = SpreadsheetApp.openById(SHEET_REGISTRO_ID);
  var sh = ss.getSheetByName(SHEET_DATOS) || ss.getSheets()[0];
  var fecha = data.fecha || Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd');
  var anio = '';
  var mes = '';
  try {
    var d = new Date(fecha);
    if (!isNaN(d.getTime())) {
      anio = d.getFullYear();
      mes = d.getMonth() + 1;
    }
  } catch (e) {}

  // Col A = Folio; B.. = mismo orden histórico del Sheet
  var row = [
    data.folio || '',
    data.cliente || '',
    data.maquina || '',
    data.tecnico || '',
    fecha,
    data.tipo_mtto || '',
    data.tipo_falla || '',
    data.falla_especifica || '',
    data.solucion || '',
    data.observaciones || '',
    data.recibido_por ? 'Sí - ' + data.recibido_por : 'Sí - firmada',
    data.folio || data.ot || '', // N OT = mismo folio
    data.estado_visita || 'cerrada',
    'Formulario web permanente ' + Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd HH:mm'),
    data.email_cliente || '',
    data.recibido_por || '',
    data.cargo || '',
    data.comuna || '',
    pdfLink || '',
    anio,
    mes,
  ];
  sh.appendRow(row);
  return sh.getLastRow();
}

function enviarCorreo_(data, pdfFile) {
  var to = splitEmails_(data.email_cliente);
  if (!to.length) {
    return { ok: false, skip: 'Falta email_cliente', to: [] };
  }
  var cc = splitEmails_(data.email_cc);
  if (!cc.length) {
    cc = [CC_DEFAULT];
  }
  var subject =
    'WES · Acta folio ' +
    (data.folio || '') +
    ' — ' +
    (data.cliente || 'Cliente') +
    ' / ' +
    (data.maquina || 'sitio') +
    ' · Acusar recibo';
  var html =
    '<div style="font-family:Segoe UI,Arial,sans-serif;color:#14202b;line-height:1.45">' +
    '<p>Estimados/as <b>' +
    escapeHtml_(data.cliente || '') +
    '</b>,</p>' +
    '<p>Adjuntamos el <b>acta en PDF</b> (folio <b>' +
    escapeHtml_(data.folio || '') +
    '</b>) de la visita técnica realizada por WES en <b>' +
    escapeHtml_(data.maquina || '') +
    '</b> (fecha ' +
    escapeHtml_(data.fecha || '') +
    ', técnico ' +
    escapeHtml_(data.tecnico || '') +
    ').</p>' +
    '<p style="background:#e7f0f8;border-left:4px solid #1f4e79;padding:12px 14px">' +
    'Solicitamos por favor <b>ACUSAR RECIBO</b> de esta acta respondiendo este correo ' +
    'con la frase «Acuso recibo» (puede indicar nombre y cargo).</p>' +
    '<p>Quien recibió en terreno: ' +
    escapeHtml_(data.recibido_por || '—') +
    '.</p>' +
    '<p>Quedamos atentos.<br/>— Sociedad Tecnológica WES SpA<br/>www.wes.cl</p></div>';

  try {
    GmailApp.sendEmail(to.join(','), subject, 'Adjuntamos acta de visita WES. Acusar recibo.', {
      htmlBody: html,
      cc: cc.join(','),
      attachments: [pdfFile.getAs(MimeType.PDF)],
      name: 'Agente IA WES',
    });
    return { ok: true, to: to, cc: cc };
  } catch (e) {
    return { ok: false, skip: String(e), to: to, cc: cc };
  }
}

function splitEmails_(raw) {
  if (!raw) return [];
  return String(raw)
    .replace(/;/g, ',')
    .split(',')
    .map(function (s) {
      return s.trim();
    })
    .filter(function (s) {
      return s.indexOf('@') > 0;
    });
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
