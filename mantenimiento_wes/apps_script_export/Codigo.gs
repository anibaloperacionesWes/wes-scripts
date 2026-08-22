/**
 * Formulario permanente de acta de visita WES (Apps Script Web App).
 *
 * Después de Implementar → Aplicación web (acceso: Cualquiera),
 * el link /exec queda fijo para técnicos.
 */

var SHEET_REGISTRO_ID = '1GlRn7QXWEre7ziau29ojR5lTl-bZ8T3mCT3cD93HZgM';
var SHEET_DATOS = 'Datos';
/** Excel/Sheet de contactos To/CC (Cliente · punto). */
var SHEET_CONTACTOS_ID = '1Tpjm1eXRXKuKvxachtbYVr9503wICJdsYDTjkbm__o8';
var SHEET_CONTACTOS_NOMBRE = 'Contactos';
var CARPETA_TECNICOS = 'Tecnicos_WES_Formulario';
/**
 * G:\Mi unidad\Agente WES\wes-scripts\mantenimiento wes
 * (código / formulario / auxiliares — no las actas PDF del terreno)
 */
var CARPETA_MANTENIMIENTO_WES_ID = '150GFVtGFlPXb_7bQfe7AS4SClKEXLEuX';
/** Tecnicos_WES_Formulario (dentro de mantenimiento wes). */
var CARPETA_TECNICOS_ID = '1RCtWP1hK4fKzjgjyvzzSbttWJZiNhtKC';
/**
 * G:\Mi unidad\Actas de Mantencion
 * Actas PDF: {Cliente}/{Año}/{mes}/
 */
var CARPETA_ACTAS_HISTORICAS_ID = '1-gDG2ND4beTpiqJqUG7d3dsT6wiHbKeQ';
/** Logo WES en Drive (Tecnicos_WES_Formulario). */
var LOGO_WES_ID = '1t4XYXYibZu_dwLftjjMw7hCX9CcSc4tY';
/**
 * HTML del formulario en Drive (Formulario_PEGAR_EN_APPS_SCRIPT.txt).
 * Así no hace falta pegar el Formulario gigante en el editor:
 * el agente actualiza este archivo y el /exec lo lee en vivo.
 */
var FORMULARIO_HTML_DRIVE_ID = '1UVCdra_Xsvozajx-32xnAOQh4Z6rNc2C';
var FOLIO_INICIAL = 2250;

// Paleta acta PDF (alineada al formulario / reportes WES)
var PDF_BLUE = '#1F4E79';
var PDF_BLUE2 = '#2E75B6';
var PDF_LIGHT = '#E7F0F8';
var PDF_LINE = '#C5D5E6';
var PDF_INK = '#14202b';
var PDF_MUTED = '#445566';
var PDF_SOFT = '#F7FAFD';
var PDF_SIGN_BG = '#F8FBFE';

/**
 * Carga el HTML del form desde Drive; si falla, usa el archivo local "Formulario".
 */
function loadFormularioTemplate_() {
  try {
    var raw = DriveApp.getFileById(FORMULARIO_HTML_DRIVE_ID).getBlob().getDataAsString('UTF-8');
    if (raw && raw.indexOf('<html') >= 0) {
      return HtmlService.createTemplate(raw);
    }
  } catch (e) {
    // Sin acceso a Drive o archivo movido → fallback local
  }
  return HtmlService.createTemplateFromFile('Formulario');
}

function doGet() {
  var tpl = loadFormularioTemplate_();
  tpl.CATALOGOS_JSON = JSON.stringify(parseCatalogosEmbed_());
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
    .setTitle('Acta de visita WES · 21P')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function parseCatalogosEmbed_() {
  var raw = HtmlService.createHtmlOutputFromFile('catalogos').getContent() || '';
  var i = raw.indexOf('{');
  var j = raw.lastIndexOf('}');
  if (i < 0 || j < i) {
    throw new Error('catalogos.html no tiene JSON válido');
  }
  return JSON.parse(raw.substring(i, j + 1));
}

function getCatalogos() {
  try {
    return buildCatalogosVivos_();
  } catch (e) {
    // Respaldo embebido si falla permiso/Sheet
    return parseCatalogosEmbed_();
  }
}

/**
 * Puntos oficiales: CONTACTOS_ENVIOS_ACTAS → Clientes_catalogo
 * (NO Base1 del Registro: ahí quedan nombres viejos, ej. RENCA).
 */
function buildCatalogosVivos_() {
  var base = parseCatalogosEmbed_();
  var ss = SpreadsheetApp.openById(SHEET_CONTACTOS_ID);
  var sh = ss.getSheetByName('Clientes_catalogo');
  if (!sh) {
    // Si renombran/borran, buscar hoja con Cliente + Máquina
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (name === 'Contactos' || name === 'Instrucciones') continue;
      var h = sheets[i].getRange(1, 1, 1, 3).getValues()[0];
      var h0 = String(h[0] || '').toLowerCase();
      var h1 = String(h[1] || '').toLowerCase();
      if (h0.indexOf('cliente') >= 0 && (h1.indexOf('m') >= 0 || h1.indexOf('sitio') >= 0)) {
        sh = sheets[i];
        break;
      }
    }
  }
  if (!sh) throw new Error('No encontré Clientes_catalogo');

  var values = sh.getDataRange().getValues();
  var map = {};
  for (var r = 1; r < values.length; r++) {
    var cli = String(values[r][0] || '').trim();
    var maq = String(values[r][1] || '').trim();
    if (!cli || !maq) continue;
    if (!map[cli]) map[cli] = [];
    if (map[cli].indexOf(maq) === -1) map[cli].push(maq);
  }
  // ordenar máquinas
  Object.keys(map).forEach(function (k) {
    map[k].sort(function (a, b) {
      return a.localeCompare(b, 'es');
    });
  });
  base.clientes_maquinas = map;
  base.fuente_puntos = 'CONTACTOS_ENVIOS_ACTAS!' + sh.getName();
  // Contactos vivos (general = David Campos; CC = encargado del punto)
  try {
    base.contactos = leerContactosDict_();
  } catch (eContactos) {
    base.contactos = base.contactos || {};
  }

  // Fallas desde Registro Base3
  try {
    var ssR = SpreadsheetApp.openById(SHEET_REGISTRO_ID);
    var b3 = ssR.getSheetByName('Base3');
    if (b3) {
      var fv = b3.getDataRange().getValues();
      var fallas = {};
      for (var j = 1; j < fv.length; j++) {
        var tipo = String(fv[j][1] || '').trim();
        var esp = String(fv[j][2] || '').trim();
        if (!tipo || !esp) continue;
        if (!fallas[tipo]) fallas[tipo] = [];
        if (fallas[tipo].indexOf(esp) === -1) fallas[tipo].push(esp);
      }
      if (Object.keys(fallas).length) base.tipos_falla = fallas;
    }
  } catch (e2) {}

  return base;
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
  var enviarCorreo = data.enviar_correo_cliente;
  if (enviarCorreo === undefined || enviarCorreo === null || enviarCorreo === '') {
    enviarCorreo = true;
  }
  if (typeof enviarCorreo === 'string') {
    enviarCorreo = ['0', 'false', 'no', 'off'].indexOf(String(enviarCorreo).toLowerCase()) < 0;
  }
  data.enviar_correo_cliente = !!enviarCorreo;
  data.trabajo_interno = !data.enviar_correo_cliente;

  if (data.enviar_correo_cliente && !data.email_cliente) {
    throw new Error('Correo del cliente es obligatorio si vas a enviar el PDF');
  }
  if (!data.solucion) {
    throw new Error('Solución / diagnóstico es obligatorio');
  }
  if (!data.firma_png) {
    throw new Error('Firma obligatoria');
  }
  if (!data.recibido_por) {
    data.recibido_por = data.trabajo_interno ? 'WES interno / oficina' : '';
    if (!data.recibido_por) {
      throw new Error('Nombre quien recibe es obligatorio');
    }
  }

  var folioReusar = Number(data.folio_reusar || data.folio_existente || 0);
  var folio;
  var esCierre = false;
  if (!isNaN(folioReusar) && folioReusar >= FOLIO_INICIAL) {
    var filaExistente = encontrarFilaPorFolio_(folioReusar);
    if (!filaExistente) {
      throw new Error('No encontré la OT ' + folioReusar + ' en el Registro');
    }
    folio = folioReusar;
    esCierre = true;
  } else {
    folio = asignarFolio_();
  }
  data = Object.assign({}, data, { folio: folio, ot: String(folio) });

  var stamp = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyyMMdd_HHmmss');
  var stem = sanitizar_(
    'folio_' + folio + '_' + (data.fecha || '') + '_' + data.cliente + '_' + data.maquina + '_' + stamp
  );

  var carpeta = asegurarCarpetaActas_(data.cliente, data.fecha);
  // Firma solo en memoria → va dentro del PDF; no se guarda PNG suelto en Drive.
  var firmaBlob = firmaBlobDesdeDataUrl_(data.firma_png);
  var pdfFile = generarYGuardarPdf_(carpeta, stem, data, firmaBlob);
  var row = esCierre
    ? updateSheetByFolio_(folio, data, pdfFile.getUrl())
    : appendSheet_(data, pdfFile.getUrl());
  try {
    if (data.enviar_correo_cliente) {
      guardarContactosDesdeVisita_(data);
    }
  } catch (eContact) {
    // No bloquea el envío si falla el registro de contactos
  }
  var emailInfo = data.enviar_correo_cliente
    ? enviarCorreo_(data, pdfFile)
    : { ok: false, skip: 'Trabajo interno: envío al cliente desactivado', to: [] };

  return {
    ok: true,
    folio: folio,
    excel_row: row,
    pdf_url: pdfFile.getUrl(),
    drive_link: pdfFile.getUrl(),
    email_ok: emailInfo.ok,
    email_to: emailInfo.to,
    email_skip: emailInfo.skip || '',
    trabajo_interno: !!data.trabajo_interno,
    cierre_ot: esCierre,
    message: emailInfo.ok
      ? 'Folio ' + folio + ' · PDF generado y correo enviado a ' + emailInfo.to.join(', ')
      : data.trabajo_interno
        ? 'Folio ' + folio + ' · PDF interno (sin correo al cliente)'
        : 'Folio ' + folio + ' · PDF generado. Correo: ' + (emailInfo.skip || 'pendiente'),
  };
}

/**
 * OT con estado abierta o en_curso (para panel de cierre).
 */
function listarOTsPendientes() {
  var ss = SpreadsheetApp.openById(SHEET_REGISTRO_ID);
  var sh = ss.getSheetByName(SHEET_DATOS) || ss.getSheets()[0];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last, 13).getValues(); // A..M
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var folio = Number(vals[i][0]);
    var estado = String(vals[i][12] || '').trim().toLowerCase();
    if (isNaN(folio) || folio < FOLIO_INICIAL) continue;
    if (estado !== 'abierta' && estado !== 'en_curso') continue;
    out.push({
      folio: folio,
      cliente: String(vals[i][1] || ''),
      maquina: String(vals[i][2] || ''),
      tecnico: String(vals[i][3] || ''),
      fecha: formatFechaSheet_(vals[i][4]),
      tipo_mtto: String(vals[i][5] || ''),
      estado: estado,
      row: i + 2,
    });
  }
  // más recientes primero
  out.sort(function (a, b) {
    return b.folio - a.folio;
  });
  return out;
}

/** Datos de una OT para continuar / cerrar en el formulario. */
function obtenerVisitaPorFolio(folio) {
  folio = Number(folio);
  if (isNaN(folio) || folio < FOLIO_INICIAL) {
    throw new Error('Folio inválido');
  }
  var ss = SpreadsheetApp.openById(SHEET_REGISTRO_ID);
  var sh = ss.getSheetByName(SHEET_DATOS) || ss.getSheets()[0];
  var last = sh.getLastRow();
  if (last < 2) throw new Error('Sin datos en Registro');
  var vals = sh.getRange(2, 1, last, 19).getValues(); // A..S
  for (var i = 0; i < vals.length; i++) {
    if (Number(vals[i][0]) !== folio) continue;
    return {
      ok: true,
      folio: folio,
      cliente: String(vals[i][1] || ''),
      maquina: String(vals[i][2] || ''),
      tecnico: String(vals[i][3] || ''),
      fecha: formatFechaSheet_(vals[i][4]),
      tipo_mtto: String(vals[i][5] || ''),
      tipo_falla: String(vals[i][6] || ''),
      falla_especifica: String(vals[i][7] || ''),
      solucion: String(vals[i][8] || ''),
      observaciones: String(vals[i][9] || ''),
      estado_visita: String(vals[i][12] || 'en_curso'),
      email_cliente: String(vals[i][14] || ''),
      recibido_por: String(vals[i][15] || ''),
      cargo: String(vals[i][16] || ''),
      comuna: String(vals[i][17] || ''),
      pdf_url: String(vals[i][18] || ''),
      row: i + 2,
    };
  }
  throw new Error('No encontré folio ' + folio);
}

/** Cierre rápido: solo cambia Estado visita → cerrada (sin nuevo PDF). */
function marcarOTCerrada(folio) {
  folio = Number(folio);
  var found = encontrarFilaPorFolio_(folio);
  if (!found) throw new Error('No encontré OT ' + folio);
  var ss = SpreadsheetApp.openById(SHEET_REGISTRO_ID);
  var sh = ss.getSheetByName(SHEET_DATOS) || ss.getSheets()[0];
  sh.getRange(found.row, 13).setValue('cerrada'); // col M
  return { ok: true, folio: folio, row: found.row, estado: 'cerrada' };
}

function formatFechaSheet_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, 'America/Santiago', 'yyyy-MM-dd');
  }
  var s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  return s;
}

function encontrarFilaPorFolio_(folio) {
  folio = Number(folio);
  if (isNaN(folio)) return null;
  var ss = SpreadsheetApp.openById(SHEET_REGISTRO_ID);
  var sh = ss.getSheetByName(SHEET_DATOS) || ss.getSheets()[0];
  var last = sh.getLastRow();
  if (last < 2) return null;
  var vals = sh.getRange(2, 1, last, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (Number(vals[i][0]) === folio) return { row: i + 2, sheet: sh };
  }
  return null;
}

function updateSheetByFolio_(folio, data, pdfLink) {
  var found = encontrarFilaPorFolio_(folio);
  if (!found) throw new Error('No encontré fila para folio ' + folio);
  var sh = found.sheet;
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
  var row = [
    data.folio || folio,
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
    data.folio || data.ot || folio,
    data.estado_visita || 'cerrada',
    'Cierre web ' + Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd HH:mm'),
    data.email_cliente || '',
    data.recibido_por || '',
    data.cargo || '',
    data.comuna || '',
    pdfLink || '',
    anio,
    mes,
  ];
  sh.getRange(found.row, 1, found.row, row.length).setValues([row]);
  return found.row;
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

function asegurarCarpetaActas_(cliente, fecha) {
  // G:\Mi unidad\Actas de Mantencion\{Cliente}\{Año}\{mes}\
  var rootActas;
  try {
    rootActas = DriveApp.getFolderById(CARPETA_ACTAS_HISTORICAS_ID);
  } catch (e) {
    throw new Error('No se pudo abrir la carpeta Actas de Mantencion');
  }
  var cli = mapNombreCarpetaCliente_(cliente);
  var parts = anioMesDesdeFecha_(fecha);
  var fCli = asegurarCarpetaPorNombre_(cli, rootActas.getId());
  var fAnio = asegurarCarpetaPorNombre_(parts.anio, fCli.getId());
  return asegurarCarpetaPorNombre_(parts.mes, fAnio.getId());
}

/** Alinea nombres del formulario con carpetas históricas cuando difieren. */
function mapNombreCarpetaCliente_(cliente) {
  var c = String(cliente || 'SIN_CLIENTE').trim() || 'SIN_CLIENTE';
  var map = {
    'COR. PUENTE': 'CORP PUENTE ALTO',
    'GENCHI': 'GENDARMERIA',
    'LA FLORIDA': 'CORP LA FLORIDA',
    'LA REINA': 'CORP LA REINA',
    'PROVIDENCIA': 'CORP PROVIDENCIA',
    'LAS CONDES': 'COLEGIO LAS CONDES',
    'NIDO': 'NIDO DE AGUILAS',
    'BUPA ANTOFGASTA': 'BUPA',
    'HEGC': 'HOSPITAL EXEQUIEL GONZALEZ CORTES',
    'MADECCO': 'MADECO',
    'MAE': 'MADECO',
    'PAE': 'PARQUE ARAUCO',
    'PAK': 'PARQUE ARAUCO',
  };
  return map[c] || c;
}

function anioMesDesdeFecha_(fecha) {
  var meses = [
    'enero',
    'febrero',
    'marzo',
    'abril',
    'mayo',
    'junio',
    'julio',
    'agosto',
    'septiembre',
    'octubre',
    'noviembre',
    'diciembre',
  ];
  var d = null;
  if (fecha) {
    // yyyy-MM-dd o similar
    var m = String(fecha).match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) {
      d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    } else {
      d = new Date(fecha);
    }
  }
  if (!d || isNaN(d.getTime())) {
    d = new Date();
  }
  var tz = 'America/Santiago';
  var anio = Utilities.formatDate(d, tz, 'yyyy');
  var mesIdx = Number(Utilities.formatDate(d, tz, 'M')) - 1;
  var mes = meses[mesIdx] || meses[new Date().getMonth()];
  return { anio: anio, mes: mes };
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

/** Blob de firma en memoria (no se crea archivo en Drive). */
function firmaBlobDesdeDataUrl_(dataUrl) {
  var parts = String(dataUrl).split(',');
  if (parts.length < 2) {
    throw new Error('Firma inválida');
  }
  var bytes = Utilities.base64Decode(parts[1]);
  return Utilities.newBlob(bytes, 'image/png', 'firma.png');
}

function generarYGuardarPdf_(carpeta, stem, data, firmaBlob) {
  var doc = DocumentApp.create('Acta_' + stem);
  var body = doc.getBody();
  body.clear();
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(42);
  body.setMarginRight(42);

  appendPdfHeader_(body);
  appendPdfMetaTable_(body, data);
  appendPdfSection_(body, 'Solución y/o diagnóstico');
  appendPdfBodyText_(body, data.solucion || '—');
  if (data.observaciones) {
    appendPdfSection_(body, 'Observaciones');
    appendPdfBodyText_(body, data.observaciones);
  }

  appendPdfSection_(body, 'Checklist CIR — eléctrico / electrónico');
  appendChecklistTable_(body, data.checklist_cir);
  appendPdfSection_(body, 'Checklist CPA — hídrico / cámara');
  appendChecklistTable_(body, data.checklist_cpa);
  appendPdfSection_(body, 'Checklist SAB (si aplica)');
  appendChecklistTable_(body, data.checklist_sab);

  appendPdfSection_(body, 'Recepción del cliente · acuse de recibo');
  appendPdfBodyText_(
    body,
    'Se solicita al cliente acusar recibo de esta acta respondiendo el correo ' +
      'con la frase «Acuso recibo» o firmando digitalmente abajo. ' +
      'La constancia queda registrada en el sistema de análisis WES.'
  );
  appendPdfFirmaBlock_(body, data, firmaBlob);
  appendPdfFooter_(body);

  doc.saveAndClose();

  var file = DriveApp.getFileById(doc.getId());
  var pdfBlob = file.getAs('application/pdf').setName(stem + '.pdf');
  var pdfFile = carpeta.createFile(pdfBlob);
  try {
    carpeta.addFile(file);
  } catch (e) {}
  file.setTrashed(true);
  return pdfFile;
}

function appendPdfHeader_(body) {
  var header = body.appendTable([['', '']]);
  header.setBorderWidth(0);
  var c0 = header.getCell(0, 0);
  var c1 = header.getCell(0, 1);
  c0.setBackgroundColor(PDF_BLUE);
  c1.setBackgroundColor(PDF_BLUE);
  c0.setWidth(120);
  c1.setWidth(360);
  c0.clear();
  c1.clear();

  try {
    var logoBlob = DriveApp.getFileById(LOGO_WES_ID).getBlob();
    var img = c0.appendImage(logoBlob);
    img.setWidth(110);
    img.setHeight(24);
  } catch (e) {
    var brand = c0.appendParagraph('WES');
    brand.setBold(true).setFontSize(18).setForegroundColor('#FFFFFF');
    brand.setSpacingBefore(6).setSpacingAfter(6);
  }

  var title = c1.appendParagraph('ACTA DE VISITA TÉCNICA');
  title
    .setBold(true)
    .setFontSize(15)
    .setForegroundColor('#FFFFFF')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setSpacingBefore(4).setSpacingAfter(0);
  var sub = c1.appendParagraph('Sociedad Tecnológica WES SpA · www.wes.cl');
  sub
    .setFontSize(9)
    .setForegroundColor('#D6E6F5')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  sub.setSpacingBefore(0).setSpacingAfter(4);

  body.appendParagraph('').setSpacingAfter(2);
}

function appendPdfMetaTable_(body, data) {
  var rows = [
    ['Folio / OT', String(data.folio || data.ot || '—'), 'Estado', String(data.estado_visita || '—')],
    ['Cliente', String(data.cliente || '—'), 'Máquina / sitio', String(data.maquina || '—')],
    ['Comuna', String(data.comuna || '—'), 'Fecha', String(data.fecha || '—')],
    ['Hora', String(data.hora || '—'), 'Técnico WES', String(data.tecnico || '—')],
    ['Lectura medidor', String(data.lectura_medidor || '—'), 'Tipo mtto', String(data.tipo_mtto || '—')],
    ['Motivo/modalidad', joinList_(data.motivos), 'Tecnología', joinList_(data.tecnologias)],
    ['Tipo falla', String(data.tipo_falla || '—'), 'Falla específica', String(data.falla_especifica || '—')],
  ];
  var table = body.appendTable(rows.map(function (r) {
    return [r[0], r[1], r[2], r[3]];
  }));
  table.setBorderColor(PDF_LINE);
  table.setBorderWidth(0.5);
  for (var i = 0; i < rows.length; i++) {
    styleMetaLabelCell_(table.getCell(i, 0));
    styleMetaValueCell_(table.getCell(i, 1));
    styleMetaLabelCell_(table.getCell(i, 2));
    styleMetaValueCell_(table.getCell(i, 3));
  }
  try {
    table.setColumnWidth(0, 95);
    table.setColumnWidth(1, 145);
    table.setColumnWidth(2, 95);
    table.setColumnWidth(3, 145);
  } catch (e) {}
}

function styleMetaLabelCell_(cell) {
  cell.setBackgroundColor(PDF_LIGHT);
  cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(5).setPaddingRight(5);
  var val = cell.getText();
  cell.clear();
  var para = cell.appendParagraph(val);
  para.setBold(true).setFontSize(8).setForegroundColor(PDF_BLUE);
  para.setSpacingBefore(0).setSpacingAfter(0);
}

function styleMetaValueCell_(cell) {
  cell.setBackgroundColor('#FFFFFF');
  cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(5).setPaddingRight(5);
  var val = cell.getText();
  cell.clear();
  var para = cell.appendParagraph(val || '—');
  para.setBold(false).setFontSize(9).setForegroundColor(PDF_INK);
  para.setSpacingBefore(0).setSpacingAfter(0);
}

function appendPdfSection_(body, title) {
  var t = body.appendParagraph(title);
  t.setBold(true)
    .setFontSize(11)
    .setForegroundColor(PDF_BLUE)
    .setSpacingBefore(10)
    .setSpacingAfter(3);
}

function appendPdfBodyText_(body, text) {
  var p = body.appendParagraph(String(text || '—'));
  p.setFontSize(10).setForegroundColor(PDF_INK).setSpacingBefore(0).setSpacingAfter(4);
}

function appendChecklistTable_(body, items) {
  var data = [['Elemento', 'Estado', 'Obs. / medición']];
  var list = items || [];
  if (!list.length) {
    data.push(['—', '—', '—']);
  } else {
    for (var i = 0; i < list.length; i++) {
      var it = list[i] || {};
      data.push([
        String(it.elemento || ''),
        String(it.estado || ''),
        String(it.obs || '—'),
      ]);
    }
  }
  var table = body.appendTable(data);
  table.setBorderColor(PDF_LINE);
  table.setBorderWidth(0.5);
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < 3; c++) {
      var cell = table.getCell(r, c);
      cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(4).setPaddingRight(4);
      var val = cell.getText();
      cell.clear();
      var para = cell.appendParagraph(val);
      para.setSpacingBefore(0).setSpacingAfter(0).setFontSize(8.5);
      if (r === 0) {
        cell.setBackgroundColor(PDF_LIGHT);
        para.setBold(true).setForegroundColor(PDF_BLUE);
      } else {
        cell.setBackgroundColor(r % 2 === 0 ? PDF_SOFT : '#FFFFFF');
        para.setBold(false).setForegroundColor(PDF_INK);
      }
    }
  }
  try {
    table.setColumnWidth(0, 200);
    table.setColumnWidth(1, 80);
    table.setColumnWidth(2, 200);
  } catch (e) {}
}

function appendPdfFirmaBlock_(body, data, firmaBlob) {
  var table = body.appendTable([
    [
      'Recibido por: ' + (data.recibido_por || '—'),
      'Cargo: ' + (data.cargo || '—'),
    ],
    ['', ''],
  ]);
  table.setBorderColor(PDF_BLUE2);
  table.setBorderWidth(1);
  for (var r = 0; r < 2; r++) {
    for (var c = 0; c < 2; c++) {
      table.getCell(r, c).setBackgroundColor(PDF_SIGN_BG);
      table.getCell(r, c).setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(8).setPaddingRight(8);
    }
  }

  // Row 0 labels
  for (var c0 = 0; c0 < 2; c0++) {
    var cell0 = table.getCell(0, c0);
    var t0 = cell0.getText();
    cell0.clear();
    var p0 = cell0.appendParagraph(t0);
    p0.setBold(true).setFontSize(9).setForegroundColor(PDF_INK).setSpacingBefore(0).setSpacingAfter(0);
  }

  // Signature + meta
  var cFirma = table.getCell(1, 0);
  cFirma.clear();
  var lab = cFirma.appendParagraph('Firma del receptor');
  lab.setFontSize(8).setForegroundColor(PDF_MUTED).setSpacingBefore(0).setSpacingAfter(2);
  try {
    var img = cFirma.appendImage(firmaBlob);
    img.setWidth(220);
    img.setHeight(70);
  } catch (e) {
    cFirma.appendParagraph('(Firma no disponible)').setFontSize(8).setForegroundColor(PDF_MUTED);
  }

  var cMeta = table.getCell(1, 1);
  cMeta.clear();
  var gen = Utilities.formatDate(new Date(), 'America/Santiago', 'dd/MM/yyyy HH:mm');
  var lines = [
    'Correo cliente: ' + (data.email_cliente || '—'),
    'Generado: ' + gen,
    'Documento digital WES',
  ];
  for (var i = 0; i < lines.length; i++) {
    var p = cMeta.appendParagraph(lines[i]);
    p.setFontSize(8).setForegroundColor(PDF_MUTED).setSpacingBefore(0).setSpacingAfter(2);
  }

  var acuse = body.appendParagraph('POR FAVOR ACUSAR RECIBO DE ESTA ACTA POR CORREO');
  acuse
    .setBold(true)
    .setFontSize(10)
    .setForegroundColor(PDF_BLUE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setSpacingBefore(10)
    .setSpacingAfter(2);
}

function appendPdfFooter_(body) {
  var p = body.appendParagraph(
    'WES · Estrecho de Magallanes 1481, Renca · +569 7559 5695 / +569 8198 1426 · www.wes.cl'
  );
  p.setFontSize(8)
    .setForegroundColor(PDF_MUTED)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setSpacingBefore(4)
    .setSpacingAfter(0);
}

function joinList_(arr) {
  if (!arr || !arr.length) return '—';
  return arr.join(', ');
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
  if (data && data.enviar_correo_cliente === false) {
    return { ok: false, skip: 'Trabajo interno: envío al cliente desactivado', to: [] };
  }
  var to = splitEmails_(data.email_cliente);
  if (!to.length) {
    return { ok: false, skip: 'Falta email_cliente', to: [] };
  }
  var cc = splitEmails_(data.email_cc);
  var subject =
    'WES · Acta folio ' +
    (data.folio || '') +
    ' — ' +
    (data.cliente || 'Cliente') +
    ' / ' +
    (data.maquina || 'sitio') +
    ' · Acusar recibo';
  var html =
    '<div style="margin:0;padding:0;background:#eef3f8">' +
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#eef3f8;padding:20px 12px">' +
    '<tr><td align="center">' +
    '<table role="presentation" width="560" cellspacing="0" cellpadding="0" style="max-width:560px;background:#ffffff;border:1px solid #c5d5e6;border-radius:14px;overflow:hidden;font-family:Segoe UI,Arial,sans-serif;color:#14202b">' +
    '<tr><td style="background:#1F4E79;padding:18px 22px;color:#fff">' +
    '<div style="font-size:11px;letter-spacing:0.14em;text-transform:uppercase;opacity:0.9;font-weight:700">WES</div>' +
    '<div style="font-size:20px;font-weight:750;margin-top:4px">Acta de visita técnica</div>' +
    '<div style="font-size:13px;opacity:0.92;margin-top:4px">Folio ' +
    escapeHtml_(data.folio || '') +
    ' · ' +
    escapeHtml_(data.cliente || '') +
    ' / ' +
    escapeHtml_(data.maquina || '') +
    '</div></td></tr>' +
    '<tr><td style="padding:20px 22px;line-height:1.5;font-size:14px">' +
    '<p style="margin:0 0 12px">Estimados/as <b>' +
    escapeHtml_(data.cliente || '') +
    '</b>,</p>' +
    '<p style="margin:0 0 12px">Adjuntamos el <b>acta en PDF</b> de la visita técnica realizada por WES' +
    ' en <b>' +
    escapeHtml_(data.maquina || '') +
    '</b> (fecha ' +
    escapeHtml_(data.fecha || '') +
    ', técnico ' +
    escapeHtml_(data.tecnico || '') +
    ').</p>' +
    '<div style="background:#e7f0f8;border-left:4px solid #1f4e79;padding:12px 14px;margin:14px 0;border-radius:0 10px 10px 0">' +
    'Solicitamos por favor <b>ACUSAR RECIBO</b> de esta acta respondiendo este correo ' +
    'con la frase «Acuso recibo» (puede indicar nombre y cargo).</div>' +
    '<p style="margin:0 0 8px">Quien recibió en terreno: <b>' +
    escapeHtml_(data.recibido_por || '—') +
    '</b>' +
    (data.cargo ? ' (' + escapeHtml_(data.cargo) + ')' : '') +
    '.</p>' +
    '<p style="margin:16px 0 0;color:#445566;font-size:13px">Quedamos atentos.<br/>— Sociedad Tecnológica WES SpA<br/>www.wes.cl</p>' +
    '</td></tr>' +
    '<tr><td style="background:#f7fafd;padding:12px 22px;font-size:11px;color:#445566;border-top:1px solid #c5d5e6">' +
    'Estrecho de Magallanes 1481, Renca · +569 7559 5695 / +569 8198 1426' +
    '</td></tr>' +
    '</table></td></tr></table></div>';

  var mailOpts = {
    htmlBody: html,
    attachments: [pdfFile.getAs(MimeType.PDF)],
    name: 'Agente IA WES',
  };
  if (cc.length) {
    mailOpts.cc = cc.join(',');
  }
  try {
    GmailApp.sendEmail(to.join(','), subject, 'Adjuntamos acta de visita WES. Acusar recibo.', mailOpts);
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

/**
 * API para el formulario: contactos guardados (encargado general + por punto).
 * Fuente viva: hoja "Contactos" del Sheet de registro (se actualiza en cada envío).
 */
function getContactos() {
  return leerContactosDict_();
}

function getContactosPara(cliente, maquina) {
  var all = leerContactosDict_();
  var cli = String(cliente || '').trim();
  var maq = String(maquina || '').trim();
  var base = all[cli] || {};
  var punto = (base.puntos && base.puntos[maq]) || {};
  // Si el nombre del catálogo no coincide exacto con Contactos, buscar parecido
  if (maq && (!punto || !punto.email_cc) && base.puntos) {
    var hit = buscarPuntoParecido_(base.puntos, maq);
    if (hit) punto = hit;
  }
  // Generales del cliente (máquina vacía) SIEMPRE — ej. David Campos en CORMUP
  var generalesDetalle = (base.generales && base.generales.length)
    ? base.generales.slice()
    : [];
  if (!generalesDetalle.length && base.email_general) {
    var emailsFallback = splitEmails_(base.email_general || '');
    for (var gi = 0; gi < emailsFallback.length; gi++) {
      generalesDetalle.push({
        nombre: gi === 0 ? (base.nombre_general || '') : '',
        cargo: gi === 0 ? (base.cargo_general || '') : '',
        email: emailsFallback[gi],
      });
    }
  }
  var emailsGen = generalesDetalle.map(function (g) { return g.email; }).filter(Boolean);
  if (!emailsGen.length && base.emails_general) emailsGen = base.emails_general.slice();
  return {
    cliente: cli,
    maquina: maq,
    email_general: emailsGen.join(', '),
    emails_general: emailsGen,
    nombre_general: (generalesDetalle[0] && generalesDetalle[0].nombre) || base.nombre_general || '',
    cargo_general: (generalesDetalle[0] && generalesDetalle[0].cargo) || base.cargo_general || '',
    // Misma forma que David Campos: lista de fichas TO + ficha CC del punto
    generales: generalesDetalle,
    punto: {
      nombre: (punto && punto.nombre) || '',
      cargo: (punto && punto.cargo) || '',
      email: (punto && punto.email_cc) || '',
    },
    email_cc: (punto && punto.email_cc) || '',
    nombre_punto: (punto && punto.nombre) || '',
    cargo_punto: (punto && punto.cargo) || '',
  };
}

function normKey_(s) {
  return String(s || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\b(ESCUELA|COLEGIO|CENTRO|EDUCACIONAL|LICEO|C E|CE)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function buscarPuntoParecido_(puntos, maquina) {
  var target = normKey_(maquina);
  if (!target || !puntos) return null;
  var keys = Object.keys(puntos);
  var i;
  // 1) exacto normalizado
  for (i = 0; i < keys.length; i++) {
    if (normKey_(keys[i]) === target) return puntos[keys[i]];
  }
  // 2) contiene (ej. CENTRO EDUCACIONAL VALLE HERMOSO ↔ C.E. VALLE HERMOSO)
  var best = null;
  var bestLen = 0;
  for (i = 0; i < keys.length; i++) {
    var k = normKey_(keys[i]);
    if (!k) continue;
    if (k.indexOf(target) >= 0 || target.indexOf(k) >= 0) {
      var score = Math.min(k.length, target.length);
      if (score > bestLen) {
        bestLen = score;
        best = puntos[keys[i]];
      }
    }
  }
  return best;
}

/**
 * Ya NO reescribe el encargado GENERAL del cliente (máquina vacía).
 * Ese dato se cura en el Excel (ej. David Campos en CORMUP).
 * Solo actualiza el contacto del PUNTO si hay CC, sin duplicar ni borrar nombre/cargo.
 */
function guardarContactosDesdeVisita_(data) {
  var cli = String(data.cliente || '').trim();
  if (!cli) return;
  var maq = String(data.maquina || '').trim();
  if (!maq) return;
  var stamp = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd HH:mm');
  var emailsCc = splitEmails_(data.email_cc);
  if (!emailsCc.length) return;

  // Un email por fila; no crear si ya existe mismo cliente+máquina+email
  for (var i = 0; i < emailsCc.length; i++) {
    upsertContactoFila_({
      cliente: cli,
      maquina: maq,
      rol: 'CC',
      nombre: String(data.recibido_por || '').trim(),
      cargo: String(data.cargo || '').trim(),
      email: emailsCc[i],
      actualizado: stamp,
      noPisarNombre: true,
    });
  }
}

function asegurarHojaContactos_() {
  var ss = SpreadsheetApp.openById(SHEET_CONTACTOS_ID);
  var sh = ss.getSheetByName(SHEET_CONTACTOS_NOMBRE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_CONTACTOS_NOMBRE);
  }
  var headers = ['Cliente', 'Máquina', 'Rol', 'Nombre', 'Cargo', 'Email', 'Actualizado'];
  var row1 = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  var empty = !row1[0];
  if (empty) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function upsertContactoFila_(row) {
  var sh = asegurarHojaContactos_();
  var last = sh.getLastRow();
  var cli = String(row.cliente || '').trim();
  var maq = String(row.maquina || '').trim();
  var rol = String(row.rol || '').trim();
  var email = String(row.email || '').trim();
  var emailKey = email.toLowerCase();
  var nombre = String(row.nombre || '').trim();
  var cargo = String(row.cargo || '').trim();
  var stamp = row.actualizado || '';
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last, 7).getValues();
    for (var i = 0; i < vals.length; i++) {
      var rCli = String(vals[i][0] || '').trim();
      var rMaq = String(vals[i][1] || '').trim();
      var rEmail = String(vals[i][5] || '').trim().toLowerCase();
      // Match por cliente+máquina+email (evita duplicar CC vs punto)
      if (rCli === cli && rMaq === maq && rEmail && rEmail === emailKey) {
        var prevNombre = String(vals[i][3] || '').trim();
        var prevCargo = String(vals[i][4] || '').trim();
        var prevRol = String(vals[i][2] || '').trim();
        var newNombre = nombre || prevNombre;
        var newCargo = cargo || prevCargo;
        var newRol = prevRol || rol;
        if (row.noPisarNombre) {
          newNombre = prevNombre || nombre;
          newCargo = prevCargo || cargo;
        }
        sh.getRange(i + 2, 1, 1, 7).setValues([
          [cli, maq, newRol, newNombre, newCargo, email, stamp],
        ]);
        return;
      }
    }
  }
  sh.appendRow([cli, maq, rol, nombre, cargo, email, stamp]);
}

function leerContactosDict_() {
  var out = {};
  try {
    var sh = asegurarHojaContactos_();
    var last = sh.getLastRow();
    if (last < 2) return out;
    var vals = sh.getRange(2, 1, last, 7).getValues();
    for (var i = 0; i < vals.length; i++) {
      var cli = String(vals[i][0] || '').trim();
      if (!cli) continue;
      var maq = String(vals[i][1] || '').trim();
      var rol = String(vals[i][2] || '').trim().toLowerCase();
      var nombre = String(vals[i][3] || '').trim();
      var cargo = String(vals[i][4] || '').trim();
      var email = String(vals[i][5] || '').trim();
      if (!email || email.indexOf('@') < 0) continue;
      if (!out[cli]) {
        out[cli] = {
          email_general: '',
          emails_general: [],
          nombre_general: '',
          cargo_general: '',
          generales: [],
          puntos: {},
        };
      }
      // Máquina vacía = siempre va (gerente / corporación / SLEP) — ej. David Campos
      if (!maq && (rol === 'general' || rol === 'to' || rol === 'cc' || !rol)) {
        var gIdx = -1;
        for (var gi = 0; gi < out[cli].generales.length; gi++) {
          if (String(out[cli].generales[gi].email || '') === email) {
            gIdx = gi;
            break;
          }
        }
        if (gIdx < 0) {
          out[cli].emails_general.push(email);
          out[cli].generales.push({ nombre: nombre, cargo: cargo, email: email });
        } else {
          var gPrev = out[cli].generales[gIdx] || { nombre: '', cargo: '', email: email };
          if (nombre && !gPrev.nombre) gPrev.nombre = nombre;
          if (cargo && !gPrev.cargo) gPrev.cargo = cargo;
          out[cli].generales[gIdx] = gPrev;
          if (out[cli].emails_general.indexOf(email) < 0) out[cli].emails_general.push(email);
        }
        out[cli].email_general = out[cli].emails_general.join(', ');
        if (nombre && !out[cli].nombre_general) out[cli].nombre_general = nombre;
        if (cargo && !out[cli].cargo_general) out[cli].cargo_general = cargo;
      } else if (maq) {
        // CC del punto: mismo detalle que el general (nombre + cargo + email)
        if (!out[cli].puntos[maq]) {
          out[cli].puntos[maq] = { email_cc: email, nombre: nombre, cargo: cargo };
        } else {
          var cur = out[cli].puntos[maq];
          var emails = splitEmails_(cur.email_cc);
          if (emails.indexOf(email) < 0) emails.push(email);
          cur.email_cc = emails.join(', ');
          if (nombre && !cur.nombre) cur.nombre = nombre;
          if (cargo && !cur.cargo) cur.cargo = cargo;
        }
      }
    }
  } catch (e) {}
  return out;
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
