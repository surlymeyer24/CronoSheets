/**
 * Utilidades varias (Drive, compresión, etc.).
 */

var CARPETA_ORIGEN_ID = '1QEaIcR5eGXEqFOaprvwgXmlX9o9xeVK9';
var CARPETA_DESTINO_ID = '1AKNdPXcjM71304UOdl96_3I-kN9zi_PR';

/** Nombre de archivo a ignorar en la compresión (sin incluir en el zip). */
var ARCHIVO_IGNORAR_EN_ZIP = 'HORA A BUSCAR';

var MIME_GOOGLE_SHEET = 'application/vnd.google-apps.spreadsheet';
var MIME_EXCEL_ARCHIVO = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel'
];

function incluirEnZip_(mime) {
  return mime === MIME_GOOGLE_SHEET || MIME_EXCEL_ARCHIVO.indexOf(mime) !== -1;
}

var MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/**
 * Exporta una Hoja de Google a .xlsx usando la API de Drive (servicio avanzado).
 * Si falla, devuelve null y se puede usar CSV como respaldo.
 */
function exportarSheetAXlsx_(archivo) {
  try {
    var fileId = archivo.getId();
    var file = Drive.Files.get(fileId);
    var exportLinks = file.exportLinks;
    if (!exportLinks || !exportLinks[MIME_XLSX]) return null;
    var url = exportLinks[MIME_XLSX];
    var resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return null;
    var blob = resp.getBlob();
    var nombre = archivo.getName();
    if (!/\.xlsx$/i.test(nombre)) nombre = nombre + '.xlsx';
    return blob.setName(nombre);
  } catch (e) {
    return null;
  }
}

/** Archivo Excel o Hoja de Google → blob para el zip. Hojas de Google → .xlsx (API Drive) o .csv si falla. */
function blobParaZip_(archivo) {
  var mime = archivo.getMimeType();
  var nombre = archivo.getName();
  if (mime === MIME_GOOGLE_SHEET) {
    var blob = exportarSheetAXlsx_(archivo);
    if (blob) return blob;
    blob = archivo.getAs('text/csv');
    if (!/\.csv$/i.test(nombre)) nombre = nombre + '.csv';
    return blob.setName(nombre);
  }
  return archivo.getBlob().setName(nombre);
}

/**
 * Crea el zip con archivos Excel (.xlsx/.xls) y Hojas de Google (exportadas a .xlsx).
 * @returns {{ success: boolean, fileId?: string, fileName?: string, message?: string }}
 */
function crearZipCarpetaParaDescarga() {
  try {
    var carpetaOrigen = DriveApp.getFolderById(CARPETA_ORIGEN_ID);
    var carpetaDestino = DriveApp.getFolderById(CARPETA_DESTINO_ID);

    var archivos = carpetaOrigen.getFiles();
    var blobs = [];

    while (archivos.hasNext()) {
      var archivo = archivos.next();
      if (archivo.getName() === ARCHIVO_IGNORAR_EN_ZIP) continue;
      if (incluirEnZip_(archivo.getMimeType())) {
        blobs.push(blobParaZip_(archivo));
      }
    }

    if (blobs.length === 0) {
      return { success: false, message: 'No hay archivos Excel ni Hojas de Google en la carpeta.' };
    }

    var nombreZip = carpetaOrigen.getName() + '_' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm') + '.zip';
    var zip = Utilities.zip(blobs, nombreZip);
    var file = carpetaDestino.createFile(zip);

    return { success: true, fileId: file.getId(), fileName: file.getName() };
  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

/** Ejecutar desde el editor de Apps Script (mismo comportamiento que antes). */
function comprimirCarpeta() {
  var r = crearZipCarpetaParaDescarga();
  if (r.success) Logger.log('Zip creado: ' + r.fileName);
  else Logger.log('Error: ' + r.message);
}
