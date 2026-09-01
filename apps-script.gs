/**
 * Recibe los chequeos reproductivos de la tablet y guarda el histórico en esta hoja.
 * También sirve de puente para preparar un chequeo en el computador y recogerlo en la tablet.
 *
 * Pestañas que crea solo:
 *   Chequeos          → una fila por animal, de todos los chequeos (el histórico).
 *   Resumen           → una fila por chequeo, con los totales.
 *   Preparados        → listas enviadas desde el computador, esperando que la tablet las recoja.
 *   Preparados_Index  → control de esas listas.
 *   Papelera          → filas de "Chequeos" que se borraron desde la app (no se eliminan,
 *                       quedan aquí con la fecha de borrado, por si hay que recuperarlas).
 *
 * VERSIÓN 4 — agrega la acción 'borrar': mueve las filas de un chequeo de
 * "Chequeos" a "Papelera" (sin eliminarlas) y quita su fila de "Resumen".
 * Si ya tenías una versión anterior, después de pegar esto hay que hacer
 * Implementar → Administrar implementaciones → lápiz → Versión nueva → Implementar.
 */

function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);
    var libro = SpreadsheetApp.getActiveSpreadsheet();

    if (datos.prueba)                  return prueba_(libro);
    if (datos.accion === 'preparar')   return preparar_(libro, datos);
    if (datos.accion === 'pendientes') return pendientes_(libro);
    if (datos.accion === 'traer')      return traer_(libro, datos);
    if (datos.accion === 'recogido')   return recogido_(libro, datos);
    if (datos.accion === 'borrar')     return borrar_(libro, datos);
    return guardarChequeo_(libro, datos);

  } catch (err) {
    return salida_({ok: false, error: String(err)});
  }
}

function doGet() {
  return ContentService.createTextOutput('Servicio de chequeo reproductivo activo.');
}

/* ---------------------------------------------------------
   Guardar el chequeo terminado (lo que manda la tablet)
   --------------------------------------------------------- */
function guardarChequeo_(libro, datos) {
  var hoja = hoja_(libro, 'Chequeos', datos.columnas);

  // limpia filas de prueba que hayan quedado de versiones anteriores
  if (hoja.getLastRow() > 1) {
    var pa = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues();
    for (var k = pa.length - 1; k >= 0; k--) {
      if (String(pa[k][0]).indexOf('PRUEBA DE CONEXION') === 0) hoja.deleteRow(k + 2);
    }
  }

  var col = datos.columnas.indexOf('SESION') + 1;
  if (col > 0 && hoja.getLastRow() > 1) {
    var vals = hoja.getRange(2, col, hoja.getLastRow() - 1, 1).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (vals[i][0] === datos.sesionId) hoja.deleteRow(i + 2);
    }
  }

  if (datos.filas && datos.filas.length) {
    hoja.getRange(hoja.getLastRow() + 1, 1, datos.filas.length, datos.columnas.length)
        .setValues(datos.filas);
  }

  // ABORTOS va al final de los encabezados (no entre VACIAS y SIN REVISAR)
  // para no correr de columna los datos de los chequeos ya guardados.
  var res = hoja_(libro, 'Resumen',
    ['FECHA', 'SESION', 'ARCHIVO', 'REVISADO POR', 'ANIMALES', 'PRENADAS', 'VACIAS', 'SIN REVISAR', 'RECIBIDO', 'ABORTOS']);
  var iRes = datos.columnas.indexOf('RESULTADO');
  var cuenta = function (v) {
    return datos.filas.filter(function (f) { return f[iRes] === v; }).length;
  };
  if (res.getLastRow() > 1) {
    var vr = res.getRange(2, 2, res.getLastRow() - 1, 1).getValues();
    for (var j = vr.length - 1; j >= 0; j--) {
      if (vr[j][0] === datos.sesionId) res.deleteRow(j + 2);
    }
  }
  res.appendRow([
    datos.filas.length ? datos.filas[0][0] : '',
    datos.sesionId, datos.archivo || '', datos.operario || '',
    datos.filas.length, cuenta('PREÑADA'), cuenta('VACIA'), cuenta('SIN REVISAR'),
    new Date(), cuenta('ABORTO')
  ]);

  return salida_({ok: true, filas: datos.filas.length});
}

/* ---------------------------------------------------------
   Borrar un chequeo (lo pide la app desde Ajustes)
   Mueve sus filas de "Chequeos" a "Papelera" (no las borra) y
   quita su fila de "Resumen". Devuelve cuántas filas movió.
   --------------------------------------------------------- */
function borrar_(libro, datos) {
  var hoja = libro.getSheetByName('Chequeos');
  if (!hoja || hoja.getLastRow() < 2) {
    quitarDeResumen_(libro, datos.sesionId);
    return salida_({ok: true, movidas: 0});
  }

  var ancho = hoja.getLastColumn();
  var encabezados = hoja.getRange(1, 1, 1, ancho).getValues()[0];
  var colSesion = encabezados.indexOf('SESION') + 1;
  if (colSesion <= 0) return salida_({ok: false, error: 'la pestaña Chequeos no tiene columna SESION'});

  var vals = hoja.getRange(2, 1, hoja.getLastRow() - 1, ancho).getValues();
  var filas = [];
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][colSesion - 1]) === String(datos.sesionId)) filas.push(vals[i]);
  }

  if (filas.length) {
    var papelera = hoja_(libro, 'Papelera', encabezados.concat(['BORRADO EL']));
    var ahora = new Date();
    var filasPapelera = filas.map(function (f) { return f.concat([ahora]); });
    papelera.getRange(papelera.getLastRow() + 1, 1, filasPapelera.length, filasPapelera[0].length)
      .setValues(filasPapelera);

    // de abajo hacia arriba, para no desordenar los índices al borrar
    for (var j = vals.length - 1; j >= 0; j--) {
      if (String(vals[j][colSesion - 1]) === String(datos.sesionId)) hoja.deleteRow(j + 2);
    }
  }

  quitarDeResumen_(libro, datos.sesionId);
  return salida_({ok: true, movidas: filas.length});
}

function quitarDeResumen_(libro, sesionId) {
  var res = libro.getSheetByName('Resumen');
  if (!res || res.getLastRow() < 2) return;
  var vr = res.getRange(2, 2, res.getLastRow() - 1, 1).getValues();
  for (var j = vr.length - 1; j >= 0; j--) {
    if (String(vr[j][0]) === String(sesionId)) res.deleteRow(j + 2);
  }
}

/* ---------------------------------------------------------
   Preparar en el computador, recoger en la tablet
   --------------------------------------------------------- */
var COLS_PREP = ['SESION', 'ANIMAL', 'ESTADO', 'DEL', 'DUS', 'PARAMETRO', 'DATOS'];
var COLS_IDX  = ['SESION', 'FECHA', 'ARCHIVO', 'ANIMALES', 'ESTADO', 'META', 'CREADO'];

function preparar_(libro, datos) {
  var hp = hoja_(libro, 'Preparados', COLS_PREP);
  var hi = hoja_(libro, 'Preparados_Index', COLS_IDX);

  var filas = datos.animales.map(function (a) {
    return [datos.sesionId, a[0], a[1], a[2], a[3], a[4], a[5]];
  });
  if (filas.length) {
    hp.getRange(hp.getLastRow() + 1, 1, filas.length, COLS_PREP.length).setValues(filas);
  }
  hi.appendRow([
    datos.sesionId, datos.fecha, datos.archivo || '', filas.length,
    'pendiente', JSON.stringify(datos.meta || {}), new Date()
  ]);

  limpiarViejos_(libro);
  return salida_({ok: true, animales: filas.length});
}

function pendientes_(libro) {
  var hi = libro.getSheetByName('Preparados_Index');
  var lista = [];
  if (hi && hi.getLastRow() > 1) {
    var v = hi.getRange(2, 1, hi.getLastRow() - 1, COLS_IDX.length).getValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][4]) === 'pendiente') {
        lista.push({
          sesionId: String(v[i][0]),
          fecha: v[i][1] instanceof Date ? v[i][1].toISOString() : String(v[i][1]),
          archivo: String(v[i][2]),
          animales: Number(v[i][3]) || 0
        });
      }
    }
  }
  return salida_({ok: true, pendientes: lista});
}

function traer_(libro, datos) {
  var hi = libro.getSheetByName('Preparados_Index');
  var hp = libro.getSheetByName('Preparados');
  if (!hi || !hp) return salida_({ok: false, error: 'sin preparados'});

  var meta = {}, fecha = '', archivo = '';
  if (hi.getLastRow() > 1) {
    var v = hi.getRange(2, 1, hi.getLastRow() - 1, COLS_IDX.length).getValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][0]) === datos.sesionId) {
        fecha = v[i][1] instanceof Date ? v[i][1].toISOString() : String(v[i][1]);
        archivo = String(v[i][2]);
        try { meta = JSON.parse(v[i][5]); } catch (err) { meta = {}; }
      }
    }
  }

  var animales = [];
  if (hp.getLastRow() > 1) {
    var p = hp.getRange(2, 1, hp.getLastRow() - 1, COLS_PREP.length).getValues();
    for (var j = 0; j < p.length; j++) {
      if (String(p[j][0]) === datos.sesionId) {
        animales.push([p[j][1], p[j][2], p[j][3], p[j][4], p[j][5], p[j][6]]);
      }
    }
  }
  return salida_({ok: true, meta: meta, fecha: fecha, archivo: archivo, animales: animales});
}

function recogido_(libro, datos) {
  var hi = libro.getSheetByName('Preparados_Index');
  if (hi && hi.getLastRow() > 1) {
    var v = hi.getRange(2, 1, hi.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][0]) === datos.sesionId) hi.getRange(i + 2, 5).setValue('recogido');
    }
  }
  return salida_({ok: true});
}

/** Borra listas preparadas de más de 30 días para que la hoja no crezca sin control. */
function limpiarViejos_(libro) {
  var hi = libro.getSheetByName('Preparados_Index');
  var hp = libro.getSheetByName('Preparados');
  if (!hi || !hp || hi.getLastRow() < 2) return;
  var limite = new Date().getTime() - 30 * 24 * 60 * 60 * 1000;
  var v = hi.getRange(2, 1, hi.getLastRow() - 1, COLS_IDX.length).getValues();
  var viejas = {};
  for (var i = v.length - 1; i >= 0; i--) {
    var t = new Date(v[i][6]).getTime();
    if (t && t < limite) { viejas[String(v[i][0])] = true; hi.deleteRow(i + 2); }
  }
  if (!Object.keys(viejas).length || hp.getLastRow() < 2) return;
  var p = hp.getRange(2, 1, hp.getLastRow() - 1, 1).getValues();
  for (var j = p.length - 1; j >= 0; j--) {
    if (viejas[String(p[j][0])]) hp.deleteRow(j + 2);
  }
}

/* --------------------------------------------------------- */
function prueba_(libro) {
  var h = hoja_(libro, 'Prueba', ['PRUEBA', 'FECHA']);
  h.appendRow(['PRUEBA DE CONEXION', new Date()]);
  return salida_({ok: true, mensaje: 'prueba recibida'});
}

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function hoja_(libro, nombre, encabezados) {
  var h = libro.getSheetByName(nombre);
  if (!h) {
    h = libro.insertSheet(nombre);
    ponerEncabezados_(h, encabezados);
    return h;
  }
  // Si la fila 1 no tiene los encabezados correctos (por ejemplo porque la pestaña
  // la creó la prueba de conexión), los corrige sin tocar los datos de abajo.
  if (encabezados && encabezados.length) {
    var ancho = Math.max(h.getLastColumn(), encabezados.length);
    var fila1 = h.getRange(1, 1, 1, ancho).getValues()[0];
    var correcto = true;
    for (var i = 0; i < encabezados.length; i++) {
      if (String(fila1[i]).trim() !== String(encabezados[i]).trim()) { correcto = false; break; }
    }
    if (!correcto) ponerEncabezados_(h, encabezados);
  }
  return h;
}

function ponerEncabezados_(h, encabezados) {
  h.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold');
  h.setFrozenRows(1);
}
