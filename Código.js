/**
 * SISTEMA PROFESIONAL 
 * VERSIÓN ULTRA-FLASH
 */
function doGet() {
  return HtmlService.createHtmlOutput("OK");
}

function onEditActive2(e) {
  if (!e) return;

  const sh = e.range.getSheet();
  const name = sh.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (name.toLowerCase() === "resumen") {
    if (row >= 5 && col >= 4 && col <= 9) {

      const filas = e.range.getNumRows();
      const cols  = e.range.getNumColumns();

      Utilities.sleep(50);

      if (filas > 1 || cols > 1) {
        restaurarResumenCompleto();
      } 
      else {
        restaurarResumenFila(row);
      }
    }
    return;
  }
  if (row >= 9 && row <= 40) {
    if (name.toLowerCase() !== "resumen") {
      ejecutarSincronizacionTotal(sh, e.source);
  }
  }
}

function onChange(e) {
  if (!e) return;

  if (e.changeType !== "INSERT_GRID" && e.changeType !== "COPY_GRID") {
    return;
  }

  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties();

  const antes = JSON.parse(props.getProperty("HOJAS_CONOCIDAS") || "[]");
  const ahora = ss.getSheets().map(s => s.getName());

  const nuevas = ahora.filter(n => !antes.includes(n));

  nuevas.forEach(nombre => {
    const hoja = ss.getSheetByName(nombre);
    if (hoja) protegerHoja(hoja);
  });

  props.setProperty("HOJAS_CONOCIDAS", JSON.stringify(ahora));
}

function protegerHoja(hoja) {
  const email = Session.getEffectiveUser().getEmail();
  const adminEmail = "implementaciones.it@bacarsa.com.ar";

  hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .concat(hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE))
    .forEach(p => { try { p.remove(); } catch(e) {} });

  const p = hoja.protect().setDescription("Bloqueo hoja");
  p.addEditor(email);
  p.addEditor(adminEmail);
  if (p.canDomainEdit()) p.setDomainEdit(false);

  if (hoja.getName().toLowerCase() === "resumen") {
    p.setUnprotectedRanges([
      hoja.getRange("B:D"),
      hoja.getRange("L:L")
    ]);
  } else {
    p.setUnprotectedRanges([
      hoja.getRange("C9:D40"),
      hoja.getRange("K9:L40"),
      hoja.getRange("M9:N40"),
      hoja.getRange(1, 15, hoja.getMaxRows(), hoja.getMaxColumns() - 14)
    ]);
  }
}


function ejecutarSincronizacionTotal(sheet, ss) {
  const data = sheet.getRange("C9:N40").getValues();
  const resultados = [];

  for (let i = 0; i < data.length; i++) {
    const ent = data[i][0];
    const sal = data[i][1];
    const fr  = ["x", "si", "sí"].includes(data[i][10].toString().toLowerCase());
    const fe  = ["x", "si", "sí"].includes(data[i][11].toString().toLowerCase());

    if (ent && sal) {
      const c = calcularHorasRelampago(ent, sal);
      resultados.push([
        minToH(c.total), "",
        fr ? minToH(c.total) : "",
        fe ? minToH(c.total) : "",
        minToH(c.diurna),
        minToH(c.nocturna)
      ]);
    } else {
      resultados.push(["", "", "", "", "", ""]);
    }
  }

  sheet.getRange("E9:J39").setValues(resultados.slice(0, 31));

  const totales = actualizarTotalesFila41(sheet);
  actualizarFilaEnResumen(ss, sheet.getName(), totales);
}

function calcularHorasRelampago(ent, sal) {
  let inicio = hToMin(ent);
  let fin    = hToMin(sal);
  let total  = (fin <= inicio) ? (fin + 1440) - inicio : fin - inicio;

  function inter(a, b, x, y) {
    return Math.max(0, Math.min(b, y) - Math.max(a, x));
  }

  let noct = 0;

  if (fin <= inicio) {
    noct += inter(inicio, 1440, 1260, 1440);
    noct += inter(0, fin, 0, 360);
  } else {
    noct += inter(inicio, fin, 1260, 1440);
    noct += inter(inicio, fin, 0, 360);
  }

  return { total, nocturna: noct, diurna: total - noct };
}

function actualizarTotalesFila41(sheet) {
  const d = sheet.getRange("E9:J39").getValues();
  let sT = 0, s100 = 0, sP = 0, sD = 0, sN = 0;

  for (let i = 0; i < d.length; i++) {
    sT   += hToMin(d[i][0]);
    s100 += hToMin(d[i][2]);
    sP   += hToMin(d[i][3]);
    sD   += hToMin(d[i][4]);
    sN   += hToMin(d[i][5]);
  }

  sheet.getRange("E41").setFormula("=I41+J41");

  let exc = Math.max(0, (sT - 12000) - s100);
  const vals = [
    minToH(exc),
    minToH(s100),
    minToH(sP),
    minToH(sD),
    minToH(sN)
  ];

  sheet.getRange("F41:J41").setValues([vals]);
  return vals;
}

function obtenerHojaResumen(ss) {
  return ss.getSheets().find(h => h.getName().toLowerCase() === "resumen") || null;
}

function actualizarFilaEnResumen(ss, nombre, fila41) {
  const resu = obtenerHojaResumen(ss);
  if (!resu) return;
  const names = resu.getRange("C5:C250").getValues();

  const buscado = nombre.toString().trim().toLowerCase();

  for (let i = 0; i < names.length; i++) {
    if (names[i][0].toString().trim().toLowerCase() === buscado) {
      resu.getRange(i + 5, 6, 1, 5).setValues([fila41]); // Cambio: i + 5
      break;
    }
  }
}

function restaurarResumenFila(fila) {
  const ss = SpreadsheetApp.getActive();
  const resu = obtenerHojaResumen(ss);
  if (!resu) return;

  const nombre = resu.getRange(fila, 3).getValue();
  if (!nombre) return;

  const sh = ss.getSheetByName(nombre);
  if (!sh) return;

  const fila41 = sh.getRange("F41:J41").getValues();
  resu.getRange(fila, 6, 1, 5).setValues(fila41);
}

function restaurarResumenCompleto() {
  const ss = SpreadsheetApp.getActive();
  const resu = obtenerHojaResumen(ss);
  if (!resu) return;
  const nombres = resu.getRange("C5:C250").getValues();

  nombres.forEach((n, i) => {
    if (!n[0]) return;
    const sh = ss.getSheetByName(n[0]);
    if (!sh) return;

    const fila41 = sh.getRange("F41:J41").getValues();
    resu.getRange(i + 5, 6, 1, 5).setValues(fila41); // Cambio: i + 5
  });
}

function forzarReproteccionTotal() {
  const ss = SpreadsheetApp.getActive();
  const email = Session.getEffectiveUser().getEmail();
  const adminEmail = "implementaciones.it@bacarsa.com.ar";

  ss.getSheets().forEach(hoja => {
    hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .concat(hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE))
      .forEach(p => { try { p.remove(); } catch(e) {} });

    const p = hoja.protect().setDescription("Bloqueo hoja");
    p.addEditor(email);
    p.addEditor(adminEmail);
    if (p.canDomainEdit()) p.setDomainEdit(false);

    if (hoja.getName().toLowerCase() === "resumen") {
      p.setUnprotectedRanges([
        hoja.getRange("B:D"),
        hoja.getRange("L:L")
      ]);
    } else {
      p.setUnprotectedRanges([
        hoja.getRange("C9:D40"),
        hoja.getRange("K9:L40"),
        hoja.getRange("M9:N40"),
        hoja.getRange(1, 15, hoja.getMaxRows(), hoja.getMaxColumns() - 14)
      ]);
    }
  });
}
function hToMin(v) {
  if (!v) return 0;
  if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
  const p = v.toString().split(":");
  return p.length < 2 ? 0 : (+p[0] * 60 + (+p[1] || 0));
}

function minToH(m) {
  if (m <= 0) return "0:00";
  const h = Math.floor(m / 60);
  const min = Math.round(m % 60);
  return "'" + h + ":" + (min < 10 ? "0" + min : min);
}
function guardarEstadoHojas() {
  const ss = SpreadsheetApp.getActive();
  const nombres = ss.getSheets().map(s => s.getName());
  PropertiesService.getDocumentProperties()
    .setProperty("HOJAS_CONOCIDAS", JSON.stringify(nombres));
}