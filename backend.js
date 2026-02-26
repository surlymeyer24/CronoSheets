function doGet() {
  return HtmlService.createHtmlOutput("OK");
}

function limpiarTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    Logger.log('Trigger: ' + t.getHandlerFunction() + ' | Evento: ' + t.getEventType());
    if (t.getHandlerFunction() === 'runAgregarGuardia') {
      ScriptApp.deleteTrigger(t);
      Logger.log('  -> ELIMINADO');
    }
  });
}
// ============================================
// MENÚ PERSONALIZADO
// ============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Gestión de Guardias')
    .addItem('📋 Abrir Panel', 'mostrarSidebar')
    .addToUi();
}

function mostrarSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Gestión de Guardias')
    .setWidth(350);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

// Para incluir archivos HTML
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function obtenerConfigOwnerBridge_() {
  const scriptProps = PropertiesService.getScriptProperties();
  const url = (scriptProps.getProperty('OWNER_WEBAPP_URL') || '').trim();
  const token = (scriptProps.getProperty('OWNER_TOKEN') || '').trim();
  return { url: url, token: token, habilitado: !!(url && token) };
}

function llamarOwnerBridge_(accion, payload) {
  const config = obtenerConfigOwnerBridge_();
  if (!config.habilitado) {
    return { status: 'disabled', message: 'owner_bridge_not_configured' };
  }

  const body = Object.assign({}, payload || {}, {
    accion: accion,
    token: config.token
  });

  try {
    const resp = UrlFetchApp.fetch(config.url, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify(body)
    });

    const statusCode = resp.getResponseCode();
    const text = resp.getContentText() || '{}';
    let data;

    try {
      data = JSON.parse(text);
    } catch (jsonError) {
      return { status: 'error', message: 'owner_bridge_invalid_json', statusCode: statusCode, raw: text };
    }

    if (statusCode < 200 || statusCode >= 300) {
      return { status: 'error', message: 'owner_bridge_http_' + statusCode, data: data };
    }

    return data;
  } catch (error) {
    return { status: 'error', message: 'owner_bridge_fetch_failed', detail: String(error) };
  }
}

function setOwnerBridgeConfig(url, token) {
  if (!url || !token) {
    throw new Error('setOwnerBridgeConfig requiere URL y token');
  }
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.setProperty('OWNER_WEBAPP_URL', String(url).trim());
  scriptProps.setProperty('OWNER_TOKEN', String(token).trim());
  const out = { success: true, message: 'Owner Bridge configurado' };
  Logger.log('setOwnerBridgeConfig: ' + JSON.stringify(out));
  return out;
}

function verOwnerBridgeConfig() {
  const c = obtenerConfigOwnerBridge_();
  const out = {
    habilitado: c.habilitado,
    url: c.url || '(vacío)',
    tokenConfigurado: !!c.token,
    tokenPreview: c.token ? (c.token.substring(0, 8) + '...') : '(vacío)'
  };
  Logger.log('verOwnerBridgeConfig: ' + JSON.stringify(out));
  return out;
}

function diagnosticoOwnerBridge() {
  const c = obtenerConfigOwnerBridge_();
  const faltantes = [];
  if (!c.url) faltantes.push('OWNER_WEBAPP_URL');
  if (!c.token) faltantes.push('OWNER_TOKEN');

  const out = {
    success: faltantes.length === 0,
    message: faltantes.length === 0
      ? 'Owner Bridge configurado'
      : 'Faltan propiedades: ' + faltantes.join(', '),
    propiedadesFaltantes: faltantes,
    config: {
      habilitado: c.habilitado,
      url: c.url || '(vacío)',
      tokenConfigurado: !!c.token,
      tokenPreview: c.token ? (c.token.substring(0, 8) + '...') : '(vacío)'
    }
  };

  Logger.log('diagnosticoOwnerBridge: ' + JSON.stringify(out));
  return out;
}


function testOwnerBridgeConexion() {
  const c = obtenerConfigOwnerBridge_();
  if (!c.habilitado) {
    const outNoConfig = {
      success: false,
      message: 'Owner Bridge no configurado. Ejecutá setOwnerBridgeConfig(url, token).'
    };
    Logger.log('testOwnerBridgeConexion: ' + JSON.stringify(outNoConfig));
    return outNoConfig;
  }

  const resp = llamarOwnerBridge_('ping', {});
  if (resp.status === 'ok') {
    const outOk = {
      success: true,
      message: 'Conexión OK con Owner Bridge',
      detail: resp
    };
    Logger.log('testOwnerBridgeConexion: ' + JSON.stringify(outOk));
    return outOk;
  }

  // Compatibilidad: owner viejo sin acción 'ping' devuelve accion_desconocida,
  // pero esto confirma que URL/token/conectividad están bien.
  if (resp && resp.message === 'accion_desconocida') {
    const outLegacy = {
      success: true,
      message: 'Conexión OK con Owner Bridge (owner desactualizado: falta acción ping, re-deploy recomendado).',
      detail: resp
    };
    Logger.log('testOwnerBridgeConexion: ' + JSON.stringify(outLegacy));
    return outLegacy;
  }

  const outError = {
    success: false,
    message: 'No se pudo conectar con Owner Bridge',
    detail: resp
  };
  Logger.log('testOwnerBridgeConexion: ' + JSON.stringify(outError));
  return outError;
}

function ejecutarDiagnosticoOwnerBridge() {
  const config = verOwnerBridgeConfig();
  const test = testOwnerBridgeConexion();

  const resumen = {
    config: config,
    test: test
  };

  Logger.log('ejecutarDiagnosticoOwnerBridge: ' + JSON.stringify(resumen));
  return resumen;
}
// ============================================
// AGREGAR GUARDIAS (BATCH)
// ============================================
function agregarGuardiasBatch(listaJSON) {
  const lista = JSON.parse(listaJSON);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaResumen = obtenerHojaResumen(ss);

  if (!hojaResumen) return { success: false, message: 'Error: No se encontró la hoja "Resumen"' };
  if (!lista || lista.length === 0) return { success: false, message: 'Error: Lista vacía' };

  // FASE 1: datos en RESUMEN (desproteger → insertar → ordenar → renumerar → reproteger)
  const fase1 = agregarEnResumen(hojaResumen, lista);

  // FASE 2: hojas individuales (crear → mover → proteger → hipervínculos)
  var fase2 = { creadas: [], fallidas: [] };
  if (fase1.agregados.length > 0) {
    fase2 = crearHojasNuevas(ss, hojaResumen, fase1.agregados);
  }

  // Armar mensaje
  var msg = '';
  if (fase1.agregados.length > 0) {
    msg += '✅ Agregados en Resumen (' + fase1.agregados.length + '):\n';
    msg += fase1.agregados.map(function(n) { return '  ✓ ' + n; }).join('\n');
  }
  if (fase2.creadas.length > 0) {
    msg += '\n\n📄 Hojas creadas (' + fase2.creadas.length + ')';
  }
  if (fase2.fallidas.length > 0) {
    msg += '\n\n⚠ Hojas no creadas (' + fase2.fallidas.length + '):\n';
    msg += fase2.fallidas.map(function(f) { return '  ✗ ' + f.nombre + ': ' + f.motivo; }).join('\n');
  }
  if (fase1.fallidos.length > 0) {
    if (msg) msg += '\n\n';
    msg += '❌ No se pudieron agregar (' + fase1.fallidos.length + '):\n';
    msg += fase1.fallidos.map(function(f) { return '  ✗ ' + f.nombre + ': ' + f.motivo; }).join('\n');
  }
  if (!msg) msg = 'No se procesó ningún guardia.';

  Logger.log('=== RESULTADO agregarGuardiasBatch ===');
  Logger.log(msg);

  return { success: fase1.agregados.length > 0, message: msg };
}

// --- FASE 1: actualizar datos en RESUMEN ---
function agregarEnResumen(hojaResumen, lista) {
  var agregados = [];
  var fallidos = [];

  desprotegerHoja(hojaResumen);

  try {
    var datos = hojaResumen.getDataRange().getValues();
    var legajosExistentes = new Set();
    for (var i = 4; i < datos.length; i++) {
      if (datos[i][3]) legajosExistentes.add(datos[i][3].toString());
    }

    var validos = [];
    lista.forEach(function(g) {
      if (legajosExistentes.has(g.legajo.toString())) {
        fallidos.push({ nombre: g.nombre, motivo: 'Legajo ' + g.legajo + ' ya existe' });
      } else {
        validos.push(g);
        legajosExistentes.add(g.legajo.toString());
      }
    });

    if (validos.length > 0) {
      var ultimaFila = hojaResumen.getLastRow();
      var filas = validos.map(function(g) {
        return ['', '', g.nombre, g.legajo, '128:00', '0:00', '0:00', '0:00', '95:00', '33:00', ''];
      });
      hojaResumen.getRange(ultimaFila + 1, 1, filas.length, 11).setValues(filas);
      validos.forEach(function(g) { agregados.push(g.nombre); });
    }

    if (agregados.length > 0) {
      var ultimaFila2 = hojaResumen.getLastRow();
      if (ultimaFila2 > 5) {
        hojaResumen.getRange(5, 1, ultimaFila2 - 4, hojaResumen.getLastColumn())
          .sort({ column: 3, ascending: true });
      }
    }

    renumerarResumen();
    aplicarEstilosResumen(hojaResumen);

  } catch (error) {
    Logger.log('Error Fase 1 (RESUMEN): ' + error.message);
  }

  try { reprotegerResumen(hojaResumen); } catch (e) {}

  // recalcular totales maestros tras modificar el contenido de Resumen
  try {
    actualizarSumasMaestras();
  } catch (err) {
    Logger.log('No se pudo actualizar totales maestros tras agregar en resumen: ' + err.message);
  }

  return { agregados: agregados, fallidos: fallidos };
}

// --- FASE 2: crear hojas individuales ---
function crearHojasNuevas(ss, hojaResumen, nombres) {
  var creadas = [];
  var fallidas = [];

  // 2a. Crear hojas desde modelo (con flush cada 50 y reintento)
  var hojaModelo = ss.getSheetByName('modelo');
  if (!hojaModelo) {
    return { creadas: [], fallidas: nombres.map(function(n) { return { nombre: n, motivo: 'No existe la hoja "modelo"' }; }) };
  }

  var pendientesReintento = [];

  nombres.forEach(function(nombre, i) {
    try {
      crearHojaDesdeModelo(ss, hojaModelo, nombre);
      creadas.push(nombre);
    } catch (e) {
      Logger.log('Error creando "' + nombre + '": ' + e.message);
      pendientesReintento.push(nombre);
    }

    if ((i + 1) % 50 === 0) {
      SpreadsheetApp.flush();
      Utilities.sleep(1000);
    }
  });

  if (pendientesReintento.length > 0) {
    Logger.log('Reintentando ' + pendientesReintento.length + ' hojas...');
    SpreadsheetApp.flush();
    Utilities.sleep(2000);

    pendientesReintento.forEach(function(nombre, i) {
      try {
        crearHojaDesdeModelo(ss, hojaModelo, nombre);
        creadas.push(nombre);
      } catch (e) {
        fallidas.push({ nombre: nombre, motivo: e.message });
      }

      if ((i + 1) % 50 === 0) {
        SpreadsheetApp.flush();
        Utilities.sleep(1000);
      }
    });
  }

  // 2b. Mover a posición alfabética
  if (creadas.length > 0) {
    try {
      moverHojasNuevas(creadas);
    } catch (e) {
      Logger.log('Error moviendo hojas: ' + e.message);
    }
  }

  // 2c. Proteger hojas nuevas (una sola pasada con referencias cacheadas)
  if (creadas.length > 0) {
    protegerHojasBatch(ss, creadas);
  }

  // 2d. Actualizar hipervínculos en RESUMEN
  actualizarHipervinculosResumen(hojaResumen);

  return { creadas: creadas, fallidas: fallidas };
}

function actualizarHipervinculosResumen(hojaResumen) {
  desprotegerHoja(hojaResumen);
  try {
    crearHipervinculosEnResumen();
  } catch (e) {
    Logger.log('Error hipervínculos: ' + e.message);
  }
  try { reprotegerResumen(hojaResumen); } catch (e) {}
}

function moverHojasNuevas(nombresNuevos) {
  if (nombresNuevos.length === 0) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var hojas = ss.getSheets().filter(function(h) { return !esHojaIgnorada(h.getName()); });
  hojas.sort(function(a, b) {
    return a.getName().localeCompare(b.getName(), 'es', { sensitivity: 'base' });
  });

  hojas.forEach(function(hoja, index) {
    ss.setActiveSheet(hoja);
    ss.moveActiveSheet(3 + index + 1);
  });
}

function crearHojaDesdeModelo(ss, hojaModelo, nombreGuardia) {
  if (ss.getSheetByName(nombreGuardia)) return;

  var nuevaHoja = hojaModelo.copyTo(ss);
  nuevaHoja.setName(nombreGuardia);
  nuevaHoja.getRange('B1').setValue(nombreGuardia);
}

function esHojaIgnorada(nombre) {
  const ignoradas = ['resumen', 'indice', 'modelo'];
  return ignoradas.includes(nombre.toLowerCase());
}

function ordenarHojasPorNombre() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const hojas = ss.getSheets().filter(
    h => !esHojaIgnorada(h.getName())
  );
  
  // Ordenar alfabéticamente
  hojas.sort((a, b) => {
    return a.getName().localeCompare(b.getName(), 'es', { sensitivity: 'base' });
  });
  
  // Reposicionar
  hojas.forEach((hoja, index) => {
    ss.setActiveSheet(hoja);
    ss.moveActiveSheet(index + 4); // después de Resumen, Indice, modelo
  });
}

function moverHojaAPosicionAlfabetica(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log('Moviendo hoja "' + nombreHoja + '" a su posición alfabética...');
  
  const hojas = ss.getSheets().filter(
    h => !esHojaIgnorada(h.getName())
  );
  
  // Ordenar alfabéticamente por nombre
  const nombresOrdenados = hojas
    .map(h => h.getName())
    .sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));
  
  Logger.log('Hojas ordenadas: ' + JSON.stringify(nombresOrdenados));
  
  // Encontrar la posición donde debería estar
  const posicionEnArray = nombresOrdenados.indexOf(nombreHoja);
  
  if (posicionEnArray === -1) {
    Logger.log('ERROR: No se encontró la hoja en la lista');
    return;
  }
  
  // La posición en el spreadsheet es: 3 hojas ignoradas + posición en array + 1
  const posicionFinal = 3 + posicionEnArray + 1;
  
  Logger.log('Posición objetivo: ' + posicionFinal);
  
  // Mover la hoja
  const hoja = ss.getSheetByName(nombreHoja);
  if (hoja) {
    ss.setActiveSheet(hoja);
    ss.moveActiveSheet(posicionFinal);
    Logger.log('Hoja movida exitosamente');
  }
}

function ordenarResumenAlfabeticamente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaResumen = obtenerHojaResumen(ss);
  
  if (!hojaResumen) return;
  
  const ultimaFila = hojaResumen.getLastRow();
  if (ultimaFila <= 4) return; // solo headers (ahora empiezan en fila 5)
  
  const rango = hojaResumen.getRange(5, 1, ultimaFila - 4, hojaResumen.getLastColumn());
  rango.sort([{column: 2, ascending: true}]); // ordenar por columna B (nombres)
  
  renumerarResumen();
}

function renumerarResumen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaResumen = obtenerHojaResumen(ss);
  
  if (!hojaResumen) return;
  
  const ultimaFila = hojaResumen.getLastRow();
  if (ultimaFila <= 4) return; // Si no hay datos después de fila 4
  
  Logger.log('Renumerando columna B desde fila 5 hasta ' + ultimaFila);
  
  // Renumerar columna B desde fila 5
  const datos = [];
  for (let i = 5; i <= ultimaFila; i++) {
    datos.push([i - 4]); // fila 5 = número 1
  }
  hojaResumen.getRange(5, 2, datos.length, 1).setValues(datos); // Columna B, desde fila 5
}

function aplicarEstilosResumen(hojaResumen) {
  if (!hojaResumen) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    hojaResumen = obtenerHojaResumen(ss);
  }
  if (!hojaResumen) return;

  var ultimaFila = hojaResumen.getLastRow();
  if (ultimaFila < 5) return;

  var filas = ultimaFila - 4;
  var columnas = 10; // B hasta K = 10 columnas
  var rango = hojaResumen.getRange(5, 2, filas, columnas);

  rango
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // Col B: Corbel 16, blanco sobre rojo, centrado
  var rangoB = hojaResumen.getRange(5, 2, filas, 1);
  rangoB.setFontFamily('Corbel').setFontSize(16).setFontColor('#ffffff')
    .setBackground('#D21312').setHorizontalAlignment('center');

  // Col C a D: Arial 13 negrita, alineado a la izquierda
  var rangoCD = hojaResumen.getRange(5, 3, filas, 2);
  rangoCD.setFontFamily('Arial').setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('left');

  // Col E: Arial 13 negrita, centrado
  var rangoE = hojaResumen.getRange(5, 5, filas, 1);
  rangoE.setFontFamily('Arial').setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Col F a K: Calibri 12, centrado, formato hh:mm
  var rangoFK = hojaResumen.getRange(5, 6, filas, 6);
  rangoFK.setFontFamily('Calibri').setFontSize(12)
    .setHorizontalAlignment('center')
    .setNumberFormat('[h]:mm');

  // Zebra solo en C-K (B tiene fondo rojo fijo)
  var coloresZebra = [];
  for (var i = 0; i < filas; i++) {
    var color = (i % 2 === 0) ? '#ffffff' : '#f3f3f3';
    coloresZebra.push(Array(9).fill(color));
  }
  hojaResumen.getRange(5, 3, filas, 9).setBackgrounds(coloresZebra);

  // Auto-ajustar ancho de columnas B(2) a K(11)
  for (var col = 2; col <= 11; col++) {
    hojaResumen.autoResizeColumn(col);
  }
}

function desprotegerHoja(hoja) {
  const ss = hoja.getParent();
  const nombreHoja = hoja.getName();
  const usuarioActual = Session.getEffectiveUser().getEmail();
  Logger.log('desprotegerHoja: Intentando desproteger "' + nombreHoja + '" como ' + usuarioActual);
  
  // FASE 1: Intentar vía Owner Bridge (más confiable para usuarios sin permisos)
  const config = obtenerConfigOwnerBridge_();
  if (config.habilitado) {
    Logger.log('  → Intentando vía Owner Bridge...');
    const viaOwner = llamarOwnerBridge_('desprotegerHoja', {
      spreadsheetId: ss.getId(),
      sheetName: nombreHoja
    });

    if (viaOwner.status === 'ok') {
      Logger.log('  ✓ Owner Bridge desprotegió exitosamente');
      SpreadsheetApp.flush();
      return;
    } else if (viaOwner.status !== 'disabled') {
      Logger.log('  ✗ Owner Bridge falló: ' + JSON.stringify(viaOwner));
    }
  }

  // FASE 2: Intentar localmente
  Logger.log('  → Intentando desprotección local...');
  const protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .concat(hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE));

  if (protecciones.length === 0) {
    Logger.log('  ✓ No hay protecciones para remover');
    return;
  }

  let removidas = 0;
  let errores = 0;
  protecciones.forEach((p, i) => {
    try {
      const editor = p.getEditors();
      const desc = p.getDescription();
      Logger.log('    Protección ' + (i+1) + ': descripción="' + desc + '", ' + editor.length + ' editores');
      
      // Si el usuario NO es editor, esta línea fallará
      p.remove();
      removidas++;
      Logger.log('      ✓ Removida');
    } catch (e) {
      errores++;
      Logger.log('      ✗ Error: ' + e.message);
    }
  });

  if (removidas > 0) SpreadsheetApp.flush();

  Logger.log('  → Resultado: ' + removidas + ' removidas, ' + errores + ' errores');
  
  if (errores > 0 && removidas === 0) {
    // No se pudo remover NADA: probablemente permisos insuficientes
    Logger.log('  ⚠ ERROR CRÍTICO: No se tienen permisos para desproteger. El usuario actual (' + usuarioActual + ') no es editor de la protección.');
  }
}

function reprotegerResumen(hojaResumen) {
  const ss = hojaResumen.getParent();
  const emailUsuario = Session.getEffectiveUser().getEmail();
  const emailAdmin = 'implementaciones.it@bacarsa.com.ar';
  const emailDes = 'desarrollo.it@bacarsa.com.ar';
  
  Logger.log('reprotegerResumen: Iniciando para usuario ' + emailUsuario);

  // FASE 1: Intentar vía Owner Bridge (más confiable)
  const config = obtenerConfigOwnerBridge_();
  if (config.habilitado) {
    Logger.log('  → Intentando vía Owner Bridge...');
    const viaOwner = llamarOwnerBridge_('reprotegerResumen', {
      spreadsheetId: ss.getId()
    });

    if (viaOwner.status === 'ok') {
      Logger.log('  ✓ Owner Bridge reprotegió exitosamente');
      SpreadsheetApp.flush();
      return;
    } else if (viaOwner.status !== 'disabled') {
      Logger.log('  ✗ Owner Bridge falló: ' + JSON.stringify(viaOwner));
    }
  }

  // FASE 2: Proteger localmente
  Logger.log('  → Protegiendo localmente...');
  try {
    // Primero quitar todas las protecciones viejas
    const protecciones = hojaResumen.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .concat(hojaResumen.getProtections(SpreadsheetApp.ProtectionType.RANGE));
    
    let removidas = 0;
    protecciones.forEach(p => {
      try { 
        p.remove();
        removidas++;
      } catch (e) {
        Logger.log('    No se pudo quitar protección vieja: ' + e.message);
      }
    });
    
    if (removidas > 0) {
      Logger.log('    ✓ ' + removidas + ' protecciones viejas removidas');
      SpreadsheetApp.flush();
    }

    // Ahora crear protección nueva
    const p = hojaResumen.protect().setDescription('Bloqueo automático');
    p.addEditor(emailUsuario);
    p.addEditor(emailAdmin);
    p.addEditor(emailDes);
    if (p.canDomainEdit()) p.setDomainEdit(false);

    p.setUnprotectedRanges([
      hojaResumen.getRange('B:D'),
      hojaResumen.getRange('L:L')
    ]);

    SpreadsheetApp.flush();
    Logger.log('  ✓ Protección nueva creada exitosamente');
  } catch (error) {
    Logger.log('  ✗ Error en reprotegerResumen: ' + error.message);
    throw error;
  }
}
function esHojaResumen(nombre) {
  return nombre.toLowerCase() === 'resumen';
}

function obtenerHojaResumen(ss) {
  return ss.getSheets().find(h => h.getName().toLowerCase() === 'resumen') || null;
}

function protegerHojasBatch(ss, nombres) {
  var email = Session.getEffectiveUser().getEmail();
  var adminEmail = 'implementaciones.it@bacarsa.com.ar';
  var emailDes = 'desarrollo.it@bacarsa.com.ar';

  Logger.log('protegerHojasBatch: Protegiendo ' + nombres.length + ' hoja(s) como ' + email);

  nombres.forEach(function(nombre, i) {
    try {
      var hoja = ss.getSheetByName(nombre);
      if (!hoja) {
        Logger.log('  Hoja "' + nombre + '" no encontrada');
        return;
      }

      Logger.log('  Protegiendo "' + nombre + '"...');
      var p = hoja.protect().setDescription('Bloqueo hoja');
      p.addEditor(email);
      p.addEditor(adminEmail);
      p.addEditor(emailDes);
      if (p.canDomainEdit()) p.setDomainEdit(false);

      p.setUnprotectedRanges([
        hoja.getRange('C9:D40'),
        hoja.getRange('K9:L40'),
        hoja.getRange('M9:N40'),
        hoja.getRange(1, 15, hoja.getMaxRows(), hoja.getMaxColumns() - 14)
      ]);
      Logger.log('    ✓ Protegida');
    } catch (e) {
      Logger.log('  ✗ Error protegiendo "' + nombre + '": ' + e.message);
    }

    if ((i + 1) % 50 === 0) SpreadsheetApp.flush();
  });
  
  Logger.log('protegerHojasBatch: Completado');
}

function aplicarProteccionAHoja(nombreHoja, esNueva) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return;

  var email = Session.getEffectiveUser().getEmail();
  var adminEmail = 'implementaciones.it@bacarsa.com.ar';
  var emailDes = 'desarrollo.it@bacarsa.com.ar';

  if (!esNueva) {
    hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .concat(hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE))
      .forEach(function(p) { try { p.remove(); } catch(e) {} });
  }

  var p = hoja.protect().setDescription('Bloqueo hoja');
  p.addEditor(email);
  p.addEditor(adminEmail);
  p.addEditor(emailDes);
  if (p.canDomainEdit()) p.setDomainEdit(false);

  p.setUnprotectedRanges([
    hoja.getRange('C9:D40'),
    hoja.getRange('K9:L40'),
    hoja.getRange('M9:N40'),
    hoja.getRange(1, 15, hoja.getMaxRows(), hoja.getMaxColumns() - 14)
  ]);
}

function aplicarProteccionTotal() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const emailUsuario = Session.getEffectiveUser().getEmail();
    const emailAdmin = 'implementaciones.it@bacarsa.com.ar';
    const emailDes = 'desarrollo.it@bacarsa.com.ar';
    const hojas = ss.getSheets();

    Logger.log('aplicarProteccionTotal iniciado por ' + emailUsuario);

    let protegidas = 0;
    let omitidas = 0;

    hojas.forEach(hoja => {
      const protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET);

      if (protecciones.length > 0) {
        // ya existe protección; por seguridad añado al usuario actual y admins como editores
        protecciones.forEach(function(p) {
          try {
            p.addEditor(emailUsuario);
            p.addEditor(emailAdmin);
            p.addEditor(emailDes);
          } catch (e) {
            Logger.log('  no se pudo agregar editor a ' + hoja.getName() + ': ' + e.message);
          }
        });
        omitidas++;
        Logger.log('  ' + hoja.getName() + ': ya protegida (se actualizaron editores)');
        return;
      }

      const nombre = hoja.getName();
      Logger.log('Protegiendo hoja: ' + nombre);
      const p = hoja.protect().setDescription('Bloqueo automático');

      p.addEditor(emailUsuario);
      p.addEditor(emailAdmin);
      p.addEditor(emailDes);

      if (p.canDomainEdit()) {
        p.setDomainEdit(false);
      }

      if (esHojaResumen(nombre)) {
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
      Logger.log('  ' + nombre + ': protegida ✓');
    });

    SpreadsheetApp.flush();

    // después de aplicar protección total, recalcular los totales maestros en Resumen
    let totalesOk = false;
    try {
      actualizarSumasMaestras();
      Logger.log('  totales maestros actualizados');
      totalesOk = true;
    } catch (err) {
      Logger.log('  no se pudo actualizar totales maestros: ' + err.message);
    }

    var msgOk = '✅ Proceso completado\n\n' +
               '✓ ' + protegidas + ' hoja(s) nuevas protegidas\n' +
               '✓ ' + omitidas + ' hoja(s) ya tenían protección (no se tocaron)\n';
    if (totalesOk) {
      msgOk += '✓ Totales maestros recalculados\n';
    } else {
      msgOk += '⚠ No se pudieron recalcular totales maestros (ver logs)\n';
    }
    msgOk += '✓ Listo para usar';
    Logger.log('=== RESULTADO aplicarProteccionTotal ===');
    Logger.log(msgOk);
    return { success: true, message: msgOk };

  } catch (error) {
    Logger.log('ERROR aplicarProteccionTotal: ' + error.message);
    return {
      success: false,
      message: 'Error al aplicar protección: ' + error.message
    };
  }
}

function repararProteccionResumen() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaResumen = obtenerHojaResumen(ss);
    if (!hojaResumen) return { success: false, message: 'No se encontró la hoja Resumen' };

    const viaOwner = llamarOwnerBridge_('reprotegerResumen', { spreadsheetId: ss.getId() });
    if (viaOwner.status === 'ok') {
      return {
        success: true,
        message: '✅ Protección de Resumen reparada correctamente vía Owner Bridge.'
      };
    }

    const emailUsuario = Session.getEffectiveUser().getEmail();
    const emailAdmin = 'implementaciones.it@bacarsa.com.ar';

    const protecciones = hojaResumen.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .concat(hojaResumen.getProtections(SpreadsheetApp.ProtectionType.RANGE));

    let removidas = 0;
    protecciones.forEach(p => {
      try { p.remove(); removidas++; } catch (e) {
        Logger.log('No se pudo quitar protección: ' + e.message);
      }
    });

    if (removidas === 0 && protecciones.length > 0) {
      return {
        success: false,
        message: 'No se pudieron quitar las protecciones de Resumen.\nEl PROPIETARIO de la hoja debe ejecutar esta función.'
      };
    }

    const p = hojaResumen.protect().setDescription('Bloqueo automático');
    p.addEditor(emailUsuario);
    p.addEditor(emailAdmin);
    if (p.canDomainEdit()) p.setDomainEdit(false);
    p.setUnprotectedRanges([
      hojaResumen.getRange('B:D'),
      hojaResumen.getRange('L:L')
    ]);

    SpreadsheetApp.flush();

    var msgOk = '✅ Protección de Resumen reparada correctamente.\n' +
               'Se quitaron ' + removidas + ' protección(es) vieja(s) y se creó una nueva.';
    Logger.log('=== RESULTADO repararProteccionResumen ===');
    Logger.log(msgOk);
    return { success: true, message: msgOk };

  } catch (error) {
    Logger.log('ERROR repararProteccionResumen: ' + error.message);
    return { success: false, message: 'Error: ' + error.message };
  }
}

function crearHipervinculosEnResumen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaResumen = obtenerHojaResumen(ss);
  if (!hojaResumen) return;

  const hojasValidas = ss.getSheets()
    .filter(h => !esHojaIgnorada(h.getName()))
    .map(h => ({ nombre: h.getName(), gridId: h.getSheetId() }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' }));

  const filaInicio = 5;
  const ultimaFila = hojaResumen.getLastRow();

  if (ultimaFila >= filaInicio) {
    hojaResumen.getRange(filaInicio, 11, ultimaFila - filaInicio + 1, 1).clearContent();
  }

  if (hojasValidas.length === 0) return;

  hojasValidas.forEach((h, i) => {
    const filaActual = filaInicio + i;
    const formula = '=HYPERLINK("#gid=' + h.gridId + '"; "VER")';
    hojaResumen.getRange('K' + filaActual).setFormula(formula);
  });
}

// ============================================
// ELIMINAR GUARDIA
// ============================================
function obtenerListaGuardias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaResumen = obtenerHojaResumen(ss);
  
  if (!hojaResumen) {
    return [];
  }
  
  const datos = hojaResumen.getDataRange().getValues();
  const guardias = [];
  
  Logger.log('=== OBTENER LISTA GUARDIAS ===');
  Logger.log('Total filas: ' + datos.length);
  
  // Columna C = nombre (índice 2), Columna D = legajo (índice 3)
  // Los datos empiezan en fila 5 (índice 4)
  for (let i = 4; i < datos.length; i++) { // empezar desde fila 5 (índice 4)
    if (datos[i][2] && datos[i][3]) { // tiene nombre y legajo
      guardias.push({
        nombre: datos[i][2],
        legajo: datos[i][3],
        fila: i + 1
      });
    }
  }
  
  Logger.log('Total guardias encontrados: ' + guardias.length);
  
  return guardias;
}

function eliminarGuardiasBatch(legajosJSON) {
  var legajos = JSON.parse(legajosJSON);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaResumen = obtenerHojaResumen(ss);

  if (!hojaResumen) return { success: false, message: 'Error: No se encontró la hoja "Resumen"' };
  if (!legajos || legajos.length === 0) return { success: false, message: 'Error: Lista vacía' };

  // FASE 1: eliminar filas de RESUMEN (desproteger → borrar filas → renumerar → reproteger)
  var fase1 = eliminarDeResumen(hojaResumen, legajos);

  // FASE 2: eliminar hojas individuales + actualizar hipervínculos
  var fase2 = { eliminadas: [], fallidas: [] };
  if (fase1.eliminados.length > 0) {
    fase2 = eliminarHojasGuardias(ss, hojaResumen, fase1.eliminados);
  }

  // Armar mensaje
  var msg = '';
  if (fase1.eliminados.length > 0) {
    msg += '✅ Eliminados (' + fase1.eliminados.length + '):\n';
    msg += fase1.eliminados.map(function(n) { return '  ✓ ' + n; }).join('\n');
  }
  if (fase1.fallidos.length > 0) {
    if (msg) msg += '\n\n';
    msg += '❌ No se pudieron eliminar (' + fase1.fallidos.length + '):\n';
    msg += fase1.fallidos.map(function(f) { return '  ✗ ' + f.nombre + ': ' + f.motivo; }).join('\n');
  }
  if (fase2.fallidas.length > 0) {
    if (msg) msg += '\n\n';
    msg += '⚠ Hojas no eliminadas (' + fase2.fallidas.length + '):\n';
    msg += fase2.fallidas.map(function(f) { return '  ✗ ' + f.nombre + ': ' + f.motivo; }).join('\n');
  }
  if (!msg) msg = 'No se procesó ningún guardia.';

  Logger.log('=== RESULTADO eliminarGuardiasBatch ===');
  Logger.log(msg);

  return { success: fase1.eliminados.length > 0, message: msg };
}

// --- FASE 1: eliminar filas de RESUMEN ---
function eliminarDeResumen(hojaResumen, legajos) {
  var eliminados = [];
  var fallidos = [];

  desprotegerHoja(hojaResumen);

  try {
    var datos = hojaResumen.getDataRange().getValues();
    var legajoSet = new Set(legajos.map(function(l) { return l.toString(); }));
    var aEliminar = [];

    for (var i = 4; i < datos.length; i++) {
      if (datos[i][3] && legajoSet.has(datos[i][3].toString())) {
        aEliminar.push({ fila: i + 1, nombre: datos[i][2], legajo: datos[i][3].toString() });
      }
    }

    aEliminar.sort(function(a, b) { return b.fila - a.fila; });
    aEliminar.forEach(function(g) {
      try {
        hojaResumen.deleteRow(g.fila);
        eliminados.push(g.nombre);
      } catch (e) {
        fallidos.push({ nombre: g.nombre, motivo: e.message });
      }
    });

    var legajosEncontrados = new Set(aEliminar.map(function(g) { return g.legajo; }));
    legajos.forEach(function(l) {
      if (!legajosEncontrados.has(l.toString())) {
        fallidos.push({ nombre: 'Legajo ' + l, motivo: 'No encontrado en Resumen' });
      }
    });

    renumerarResumen();
    aplicarEstilosResumen(hojaResumen);

  } catch (error) {
    Logger.log('Error Fase 1 (RESUMEN eliminar): ' + error.message);
  }

  try { reprotegerResumen(hojaResumen); } catch (e) {}

  // actualizar totales maestros luego de la eliminación
  try {
    actualizarSumasMaestras();
  } catch (err) {
    Logger.log('No se pudo actualizar totales maestros tras eliminar en resumen: ' + err.message);
  }

  return { eliminados: eliminados, fallidos: fallidos };
}

// --- FASE 2: eliminar hojas + actualizar hipervínculos ---
function eliminarHojasGuardias(ss, hojaResumen, nombres) {
  var eliminadas = [];
  var fallidas = [];

  const viaOwner = llamarOwnerBridge_('eliminarHojas', {
    spreadsheetId: ss.getId(),
    nombres: nombres
  });

  if (viaOwner.status === 'ok') {
    eliminadas = (viaOwner.eliminadas || []).slice();
    fallidas = (viaOwner.fallidas || []).slice();
    actualizarHipervinculosResumen(hojaResumen);
    return { eliminadas: eliminadas, fallidas: fallidas };
  }

  if (viaOwner.status !== 'disabled') {
    Logger.log('Owner Bridge eliminarHojas falló: ' + JSON.stringify(viaOwner));
  }

  nombres.forEach(function(nombre) {
    try {
      var hoja = ss.getSheetByName(nombre);
      if (hoja) {
        ss.deleteSheet(hoja);
        eliminadas.push(nombre);
      }
    } catch (e) {
      fallidas.push({ nombre: nombre, motivo: e.message });
    }
  });

  actualizarHipervinculosResumen(hojaResumen);

  return { eliminadas: eliminadas, fallidas: fallidas };
}

/**
 * Borra hojas que existen en el libro pero cuyo nombre no está en Resumen (columna C).
 * Útil para limpiar hojas huérfanas.
 */
function borrarHojasNoEnResumen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaResumen = obtenerHojaResumen(ss);
  if (!hojaResumen) return { success: false, message: 'No se encontró la hoja "Resumen".' };

  var datos = hojaResumen.getDataRange().getValues();
  var nombresEnResumen = new Set();
  for (var i = 4; i < datos.length; i++) {
    var n = datos[i][2];
    if (n && typeof n === 'string' && n.trim() !== '') nombresEnResumen.add(n.trim());
  }

  var hojas = ss.getSheets();
  var aBorrar = [];
  hojas.forEach(function(hoja) {
    var nombre = hoja.getName();
    if (esHojaIgnorada(nombre)) return;
    if (!nombresEnResumen.has(nombre)) aBorrar.push(hoja);
  });

  var eliminadas = [];
  var fallidas = [];
  aBorrar.forEach(function(hoja) {
    var nombre = hoja.getName();
    try {
      desprotegerHoja(hoja);
      ss.deleteSheet(hoja);
      eliminadas.push(nombre);
    } catch (e) {
      fallidas.push({ nombre: nombre, motivo: e.message });
    }
  });

  if (eliminadas.length > 0) actualizarHipervinculosResumen(hojaResumen);

  var msg = '';
  if (eliminadas.length > 0) {
    msg += 'Hojas borradas (no estaban en Resumen) (' + eliminadas.length + '):\n';
    msg += eliminadas.map(function(n) { return '  ✓ ' + n; }).join('\n');
  }
  if (fallidas.length > 0) {
    if (msg) msg += '\n\n';
    msg += 'No se pudieron borrar (' + fallidas.length + '):\n';
    msg += fallidas.map(function(f) { return '  ✗ ' + f.nombre + ': ' + f.motivo; }).join('\n');
  }
  if (!msg) msg = 'No había hojas para borrar. Todas las hojas coinciden con Resumen.';

  return { success: true, message: msg };
}
function configurarOwnerBridge() {
  setOwnerBridgeConfig(
    'https://script.google.com/macros/s/AKfycbw8wrBStBIo0m77ei5v6MZ1p62FAOog0ck_OSa075WqO8CHwAEawAb8YaQgEjbDe3c/exec',
    '0209a839-f3e3-4fe9-b8d5-26ec29952651'
  );
}