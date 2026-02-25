
function doPost(e) {
  const scriptProps = PropertiesService.getScriptProperties();
  const OWNER_TOKEN = scriptProps.getProperty('OWNER_TOKEN');
  
  try {
    const body = e.postData && e.postData.type === 'application/json'
      ? JSON.parse(e.postData.contents)
      : {};

    if (!body.token || body.token !== OWNER_TOKEN) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'invalid_token' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (!body.spreadsheetId) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'spreadsheetId_requerido' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    let result = {};

    switch (body.accion) {
      case 'eliminarHojas':
        result = eliminarHojasOwner(body.spreadsheetId, body.nombres);
        break;
      case 'desprotegerHoja':
        result = desprotegerHojaOwner(body.spreadsheetId, body.sheetName);
        break;
      case 'reprotegerResumen':
        result = reprotegerResumenOwner(body.spreadsheetId);
        break;
      case 'confirmarProteccion':
        result = confirmarProteccionOwner(body.spreadsheetId, body.inicio, body.cantidad);
        break;
      case 'ping':
        result = pingOwner();
        break;
      default:
        result = { status: 'error', message: 'accion_desconocida' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function eliminarHojasOwner(spreadsheetId, nombres) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const eliminadas = [];
  const fallidas = [];
  nombres = nombres || [];

  nombres.forEach(function(nombre) {
    try {
      const hoja = ss.getSheetByName(nombre);
      if (hoja) {
        hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET)
          .concat(hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE))
          .forEach(p => { try { p.remove(); } catch(e) {} });
        ss.deleteSheet(hoja);
        eliminadas.push(nombre);
      }
    } catch(e) {
      fallidas.push({ nombre: nombre, motivo: e.message });
    }
  });

  return { status: 'ok', eliminadas: eliminadas, fallidas: fallidas };
}

function desprotegerHojaOwner(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const hoja = ss.getSheetByName(sheetName);
  if (!hoja) return { status: 'error', message: 'sheet_not_found' };

  hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .concat(hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE))
    .forEach(p => { try { p.remove(); } catch(e) {} });

  SpreadsheetApp.flush();
  return { status: 'ok' };
}

function reprotegerResumenOwner(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const hojaResumen = ss.getSheets().find(h => h.getName().toLowerCase() === 'resumen');
  if (!hojaResumen) return { status: 'error', message: 'resumen_not_found' };

  hojaResumen.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .concat(hojaResumen.getProtections(SpreadsheetApp.ProtectionType.RANGE))
    .forEach(p => { try { p.remove(); } catch(e) {} });

  const p = hojaResumen.protect().setDescription('Bloqueo automático');
  p.addEditor('desarrollo.it@bacarsa.com.ar');
  p.addEditor('implementaciones.it@bacarsa.com.ar');
  if (p.canDomainEdit()) p.setDomainEdit(false);
  p.setUnprotectedRanges([
    hojaResumen.getRange('B:D'),
    hojaResumen.getRange('L:L')
  ]);

  SpreadsheetApp.flush();
  return { status: 'ok' };
}

function confirmarProteccionOwner(spreadsheetId, inicio, cantidad) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const hojas = ss.getSheets();
  const fin = Math.min(inicio + cantidad, hojas.length);
  let protegidas = 0;
  let omitidas = 0;

  for (let i = inicio; i < fin; i++) {
    const hoja = hojas[i];
    const nombre = hoja.getName();
    const protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET);

    if (protecciones.length > 0) {
      omitidas++;
      continue;
    }

    const p = hoja.protect().setDescription('Bloqueo automático');
    p.addEditor('desarrollo.it@bacarsa.com.ar');
    p.addEditor('implementaciones.it@bacarsa.com.ar');
    if (p.canDomainEdit()) p.setDomainEdit(false);

    if (nombre.toLowerCase() === 'resumen') {
      p.setUnprotectedRanges([
        hoja.getRange('B:D'),
        hoja.getRange('L:L')
      ]);
    } else {
      p.setUnprotectedRanges([
        hoja.getRange('C9:D40'),
        hoja.getRange('K9:L40'),
        hoja.getRange('M9:N40'),
        hoja.getRange(1, 15, hoja.getMaxRows(), hoja.getMaxColumns() - 14)
      ]);
    }
    protegidas++;
  }

  SpreadsheetApp.flush();
  const siguienteInicio = fin < hojas.length ? fin : -1;
  return { status: 'ok', protegidas: protegidas, omitidas: omitidas, totalHojas: hojas.length, siguienteInicio: siguienteInicio };
}
function generateAndSetOwnerToken() {
  const token = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('OWNER_TOKEN', token);
  Logger.log('OWNER_TOKEN: ' + token);
  return token;
}

// Helper para establecer el token (ejecutar como propietario una vez)
function setOwnerToken(token) {
  if (!token || typeof token !== 'string') throw new Error('setOwnerToken requiere un token no vacío');
  PropertiesService.getScriptProperties().setProperty('OWNER_TOKEN', String(token));
  Logger.log('OWNER_TOKEN seteado');
}
function testAcceso() {
  try {
    const ss = SpreadsheetApp.openById('1v2HScN-71Am_bISIkS5OARhT-1Zhdif_Gt14w1mQqpM');
    Logger.log(ss.getName());
  } catch(e) {
    Logger.log('Error: ' + e.toString());
  }
}
function verToken() {
  Logger.log(PropertiesService.getScriptProperties().getProperty('OWNER_TOKEN'));
}

function pingOwner() {
  return { status: 'ok', message: 'owner_bridge_ready', timestamp: new Date().toISOString() };
}
