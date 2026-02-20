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

  } catch (error) {
    Logger.log('Error Fase 1 (RESUMEN): ' + error.message);
  }

  try { reprotegerResumen(hojaResumen); } catch (e) {}

  return { agregados: agregados, fallidos: fallidos };
}

// --- FASE 2: crear hojas individuales ---
function crearHojasNuevas(ss, hojaResumen, nombres) {
  var creadas = [];
  var fallidas = [];

  nombres.forEach(function(nombre) {
    try {
      crearHojaDesdeModelo(nombre);
      creadas.push(nombre);
    } catch (e) {
      fallidas.push({ nombre: nombre, motivo: e.message });
    }
  });

  if (creadas.length > 0) {
    moverHojasNuevas(creadas);

    creadas.forEach(function(nombre) {
      try {
        aplicarProteccionAHoja(nombre);
      } catch (e) {
        Logger.log('Error protegiendo "' + nombre + '": ' + e.message);
      }
    });
  }

  desprotegerHoja(hojaResumen);
  try { crearHipervinculosEnResumen(); } catch (e) {
    Logger.log('Error hipervínculos: ' + e.message);
  }
  try { reprotegerResumen(hojaResumen); } catch (e) {}

  return { creadas: creadas, fallidas: fallidas };
}

function moverHojasNuevas(nombresNuevos) {
  if (nombresNuevos.length === 0) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nuevosSet = new Set(nombresNuevos);

  nombresNuevos
    .slice()
    .sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }))
    .forEach(nombre => {
      const nombresOrdenados = ss.getSheets()
        .filter(h => !esHojaIgnorada(h.getName()))
        .map(h => h.getName())
        .sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));

      const idx = nombresOrdenados.indexOf(nombre);
      if (idx === -1) return;

      const hoja = ss.getSheetByName(nombre);
      if (hoja) {
        ss.setActiveSheet(hoja);
        ss.moveActiveSheet(3 + idx + 1);
      }
    });
}

function crearHojaDesdeModelo(nombreGuardia) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaModelo = ss.getSheetByName('modelo');
  
  if (!hojaModelo) {
    throw new Error('No existe la hoja "modelo"');
  }
  
  // Verificar si ya existe la hoja
  if (ss.getSheetByName(nombreGuardia)) {
    return; // Ya existe, no crear duplicado
  }
  
  // Copiar modelo
  const nuevaHoja = hojaModelo.copyTo(ss);
  nuevaHoja.setName(nombreGuardia);
  
  // Poner nombre en título (B1:N3)
  const rango = nuevaHoja.getRange('B1:N3');
  rango.breakApart();
  rango.merge();
  rango.setValue(nombreGuardia);
  rango.setHorizontalAlignment('center');
  rango.setVerticalAlignment('middle');
  rango.setFontSize(37);
  rango.setFontWeight('bold');
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

function desprotegerHoja(hoja) {
  const protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .concat(hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE));

  if (protecciones.length === 0) return;

  let removidas = 0;
  protecciones.forEach(p => {
    try {
      p.remove();
      removidas++;
    } catch (e) {
      Logger.log('No se pudo quitar protección de "' + hoja.getName() + '": ' + e.message);
    }
  });

  if (removidas > 0) SpreadsheetApp.flush();

  if (removidas < protecciones.length) {
    Logger.log('AVISO: quedaron ' + (protecciones.length - removidas) +
      ' protecciones sin quitar en "' + hoja.getName() + '"');
  }
}

function reprotegerResumen(hojaResumen) {
  try {
    const emailUsuario = Session.getEffectiveUser().getEmail();
    const emailAdmin = 'implementaciones.it@bacarsa.com.ar';

    const p = hojaResumen.protect().setDescription('Bloqueo automático');
    p.addEditor(emailUsuario);
    p.addEditor(emailAdmin);
    if (p.canDomainEdit()) p.setDomainEdit(false);

    p.setUnprotectedRanges([
      hojaResumen.getRange('B:D'),
      hojaResumen.getRange('L:L')
    ]);

    SpreadsheetApp.flush();
  } catch (error) {
    Logger.log('Error en reprotegerResumen: ' + error.message);
  }
}
function esHojaResumen(nombre) {
  return nombre.toLowerCase() === 'resumen';
}

function obtenerHojaResumen(ss) {
  return ss.getSheets().find(h => h.getName().toLowerCase() === 'resumen') || null;
}

function aplicarProteccionAHoja(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  
  if (!hoja) {
    Logger.log('ERROR: No se encontró la hoja "' + nombreHoja + '"');
    return;
  }
  
  Logger.log('Aplicando protección a la hoja "' + nombreHoja + '"...');
  
  const email = Session.getEffectiveUser().getEmail();
  const adminEmail = 'implementaciones.it@bacarsa.com.ar';
  
  hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .concat(hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE))
    .forEach(p => { try { p.remove(); } catch(e) {} });
  
  const p = hoja.protect().setDescription('Bloqueo hoja');
  p.addEditor(email);
  p.addEditor(adminEmail);
  if (p.canDomainEdit()) p.setDomainEdit(false);
  
  p.setUnprotectedRanges([
    hoja.getRange('C9:D40'),
    hoja.getRange('K9:L40'),
    hoja.getRange('M9:N40'),
    hoja.getRange(1, 15, hoja.getMaxRows(), hoja.getMaxColumns() - 14)
  ]);
  
  Logger.log('Protección aplicada exitosamente');
}

function aplicarProteccionTotal() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const emailUsuario = Session.getEffectiveUser().getEmail();
    const emailAdmin = 'implementaciones.it@bacarsa.com.ar';
    const hojas = ss.getSheets();

    let protegidas = 0;
    let omitidas = 0;

    hojas.forEach(hoja => {
      const protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET);

      if (protecciones.length > 0) {
        omitidas++;
        return;
      }

      const nombre = hoja.getName();
      Logger.log('Protegiendo hoja: ' + nombre);
      const p = hoja.protect().setDescription('Bloqueo automático');

      p.addEditor(emailUsuario);
      p.addEditor(emailAdmin);

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
    });

    SpreadsheetApp.flush();

    return {
      success: true,
      message: '✅ Proceso completado\n\n' +
               '✓ ' + protegidas + ' hoja(s) nuevas protegidas\n' +
               '✓ ' + omitidas + ' hoja(s) ya tenían protección (no se tocaron)\n' +
               '✓ Listo para usar'
    };

  } catch (error) {
    Logger.log('ERROR: ' + error.message);
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

    return {
      success: true,
      message: '✅ Protección de Resumen reparada correctamente.\n' +
               'Se quitaron ' + removidas + ' protección(es) vieja(s) y se creó una nueva.'
    };

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

  } catch (error) {
    Logger.log('Error Fase 1 (RESUMEN eliminar): ' + error.message);
  }

  try { reprotegerResumen(hojaResumen); } catch (e) {}

  return { eliminados: eliminados, fallidos: fallidos };
}

// --- FASE 2: eliminar hojas + actualizar hipervínculos ---
function eliminarHojasGuardias(ss, hojaResumen, nombres) {
  var eliminadas = [];
  var fallidas = [];

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

  desprotegerHoja(hojaResumen);
  try { crearHipervinculosEnResumen(); } catch (e) {
    Logger.log('Error hipervínculos: ' + e.message);
  }
  try { reprotegerResumen(hojaResumen); } catch (e) {}

  return { eliminadas: eliminadas, fallidas: fallidas };
}