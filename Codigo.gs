/**
 * BASE INSPECCION MENSUAL WTP
 * - Hoja INSPECCIONES: un registro por inspección.
 * - Hoja FALTANTES: una fila por pieza/accesorio faltante.
 * - Hoja MASTER: lista de equipos (EQUIPO | N° BIEN | PLACA) que usa la app. Edítala cuando entren o salgan equipos.
 * - Los PDF se guardan en la carpeta de Drive "INSPECCION_MENSUAL_PDF".
 */
const CARPETA_PDF = 'INSPECCION_MENSUAL_PDF';
const CARPETA_FOTOS = 'INSPECCION_FOTOS';

const COLS = [
  ['ID', 'id'], ['FECHA REGISTRO', 'registro'], ['TRABAJO REALIZADO', 'trabajo'], ['ESTADO OT', 'estado'],
  ['EQUIPO', 'equipo'], ['N° BIEN', 'bien'], ['TIPO MANTTO', 'tipo'], ['FRENTE', 'frente'], ['UBICACION', 'ubicacion'],
  ['HOROMETRO/KM', 'horometro'], ['UNIDAD', 'unidad'], ['FECHA INICIO', 'fecha_inicio'], ['HORA INICIO', 'hora_inicio'],
  ['FECHA TERMINO', 'fecha_termino'], ['HORA FINAL', 'hora_final'], ['DURACION (H)', 'duracion_h'],
  ['DESCRIPCION', 'descripcion'], ['TRABAJADORES', 'trabajadores'], ['SUPERVISOR', 'supervisor'], ['MATERIALES', 'materiales'],
  ['OBSERVACIONES', 'observaciones'], ['N° OBSERVACIONES', 'n_observaciones'],
  ['GUIA DE REMISION', 'guia'], ['FALTANTES', 'faltantes_txt'], ['N° FALTANTES', 'n_faltantes'],
  ['ESTADO EQUIPO', 'estado_equipo'], ['N° FOTOS', 'n_fotos'], ['ARCHIVO', 'archivo'], ['PDF', 'pdf_link'],
  ['CARPETA FOTOS', 'carpeta_link']
];
const COLS_FALT = ['ID', 'FECHA', 'EQUIPO', 'N° BIEN', 'FRENTE', 'UBICACION', 'GUIA DE REMISION', 'PIEZA / ACCESORIO', 'CANTIDAD', 'COMENTARIO'];
const COLS_MASTER = ['EQUIPO', 'N° BIEN', 'PLACA'];

function doGet(e) {
  if (e && e.parameter && e.parameter.master) {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MASTER');
    if (!sh || sh.getLastRow() < 2) return json_({ ok: true, rows: [] });
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues()
      .filter(r => r[0] && r[1]);
    return json_({ ok: true, rows: rows });
  }
  return ContentService.createTextOutput('OK - Base Inspeccion Mensual activa');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = hoja_(ss, 'INSPECCIONES', COLS.map(c => c[0]));
    if (d.accion === 'foto') return guardarFoto_(sh, d);
    const shF = hoja_(ss, 'FALTANTES', COLS_FALT);

    if (d.pdf) {
      const blob = Utilities.newBlob(Utilities.base64Decode(d.pdf), 'application/pdf', d.archivo || (d.id + '.pdf'));
      d.pdf_link = carpeta_().createFile(blob).getUrl();
    }
    delete d.pdf;

    const fila = buscar_(sh, d.id);
    const colPdf = COLS.findIndex(c => c[1] === 'pdf_link') + 1;
    const colCarp = COLS.findIndex(c => c[1] === 'carpeta_link') + 1;
    if (fila && !d.pdf_link) d.pdf_link = sh.getRange(fila, colPdf).getValue();
    if (fila) d.carpeta_link = sh.getRange(fila, colCarp).getValue();
    if (!d.carpeta_link && d.carpeta) {
      const raiz = DriveApp.getFoldersByName(CARPETA_FOTOS);
      if (raiz.hasNext()) { const it = raiz.next().getFoldersByName(d.carpeta); if (it.hasNext()) d.carpeta_link = it.next().getUrl(); }
    }
    const valores = COLS.map(c => {
      const v = d[c[1]];
      if (c[1] === 'registro') return new Date(v || Date.now());
      return v === null || v === undefined ? '' : v;
    });
    if (fila) sh.getRange(fila, 1, 1, valores.length).setValues([valores]);
    else sh.appendRow(valores);

    // FALTANTES: reemplaza los de este registro
    const ids = shF.getLastRow() > 1 ? shF.getRange(2, 1, shF.getLastRow() - 1, 1).getValues() : [];
    for (let i = ids.length - 1; i >= 0; i--) if (ids[i][0] === d.id) shF.deleteRow(i + 2);
    (d.faltantes || []).forEach(f => shF.appendRow([
      d.id, d.fecha_inicio || '', d.equipo || '', d.bien || '', d.frente || '', d.ubicacion || '', d.guia || '',
      f.pieza || '', f.cantidad === '' ? '' : Number(f.cantidad), f.comentario || ''
    ]));

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function guardarFoto_(sh, d) {
  let raiz = DriveApp.getFoldersByName(CARPETA_FOTOS);
  raiz = raiz.hasNext() ? raiz.next() : DriveApp.createFolder(CARPETA_FOTOS);
  const it = raiz.getFoldersByName(d.carpeta);
  const carp = it.hasNext() ? it.next() : raiz.createFolder(d.carpeta);
  if (!carp.getFilesByName(d.nombre).hasNext()) {
    carp.createFile(Utilities.newBlob(Utilities.base64Decode(d.b64), 'image/jpeg', d.nombre));
  }
  const fila = buscar_(sh, d.id);
  if (fila) {
    const col = COLS.findIndex(c => c[1] === 'carpeta_link') + 1;
    const cel = sh.getRange(fila, col);
    if (!cel.getValue()) cel.setValue(carp.getUrl());
  }
  return json_({ ok: true });
}

function hoja_(ss, nombre, cab) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, cab.length).setValues([cab])
      .setFontWeight('bold').setBackground('#4b6f7f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    return sh;
  }
  // Si la hoja ya existía con menos columnas, actualiza el encabezado
  const actual = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  if (actual.join('|') !== cab.join('|')) {
    sh.getRange(1, 1, 1, cab.length).setValues([cab])
      .setFontWeight('bold').setBackground('#4b6f7f').setFontColor('#ffffff');
  }
  return sh;
}

function buscar_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}

function carpeta_() {
  const it = DriveApp.getFoldersByName(CARPETA_PDF);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_PDF);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Ejecuta esta función una vez: crea/actualiza las hojas y da permisos. */
function prepararHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hoja_(ss, 'INSPECCIONES', COLS.map(c => c[0]));
  hoja_(ss, 'FALTANTES', COLS_FALT);
  hoja_(ss, 'MASTER', COLS_MASTER);
  carpeta_();
  if (!DriveApp.getFoldersByName(CARPETA_FOTOS).hasNext()) DriveApp.createFolder(CARPETA_FOTOS);
}
