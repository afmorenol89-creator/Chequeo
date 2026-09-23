/**
 * Recibe los chequeos reproductivos de la tablet y guarda el histórico en esta hoja.
 * También sirve de puente para preparar un chequeo en el computador y recogerlo en la tablet,
 * y da soporte a la herramienta "Lotes de ordeño".
 *
 * Pestañas que crea solo:
 *   Chequeos          → una fila por animal, de todos los chequeos (el histórico).
 *   Resumen           → una fila por chequeo, con los totales.
 *   Preparados        → listas enviadas desde el computador, esperando que la tablet las recoja.
 *   Preparados_Index  → control de esas listas.
 *   Papelera          → filas de "Chequeos" que se borraron desde la app (no se eliminan,
 *                       quedan aquí con la fecha de borrado, por si hay que recuperarlas).
 *   Animales_Ordeno   → un animal por fila; nunca se borran filas, un animal que sale queda
 *                       con estado = fuera.
 *   Movimientos       → registro que solo crece; el lote actual de un animal es su último
 *                       movimiento (por fecha_hora).
 *   Config_Lotes      → lista de lotes disponibles (se siembra con 1, 2, 3 al crearla).
 *   Trabajadores      → lista de nombres para el selector "¿quién eres?" (se crea vacía).
 *
 * VERSIÓN 6 — agrega accion=cargar_excel (aplica la revisión de la pantalla de carga)
 * y los campos origen/ingresado_por/fecha_entrada en lotes_estado. No cambia nada del
 * chequeo reproductivo.
 * Si ya tenías una versión anterior, después de pegar esto hay que hacer
 * Implementar → Administrar implementaciones → lápiz → Versión nueva → Implementar.
 */

function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);
    var libro = SpreadsheetApp.getActiveSpreadsheet();

    if (datos.prueba)                     return prueba_(libro);
    if (datos.accion === 'preparar')      return preparar_(libro, datos);
    if (datos.accion === 'pendientes')    return pendientes_(libro);
    if (datos.accion === 'traer')         return traer_(libro, datos);
    if (datos.accion === 'recogido')      return recogido_(libro, datos);
    if (datos.accion === 'borrar')        return borrar_(libro, datos);
    if (datos.accion === 'movimientos')   return movimientos_(libro, datos);
    if (datos.accion === 'animal_nuevo')      return animalNuevo_(libro, datos);
    if (datos.accion === 'animal_salida')     return animalSalida_(libro, datos);
    if (datos.accion === 'animal_reactivar')  return animalReactivar_(libro, datos);
    if (datos.accion === 'cargar_excel')      return cargarExcel_(libro, datos);
    return guardarChequeo_(libro, datos);

  } catch (err) {
    return salida_({ok: false, error: String(err)});
  }
}

function doGet(e) {
  try {
    var accion = e && e.parameter && e.parameter.accion;
    if (accion === 'lotes_estado') {
      return lotesEstado_(SpreadsheetApp.getActiveSpreadsheet());
    }
    return ContentService.createTextOutput('Servicio de chequeo reproductivo activo.');
  } catch (err) {
    return salida_({ok: false, error: String(err)});
  }
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

/* ---------------------------------------------------------
   Lotes de ordeño
   --------------------------------------------------------- */
var COLS_ANIMALES = ['numero', 'nombre', 'estado', 'origen', 'fecha_entrada',
                      'ingresado_por', 'fecha_salida', 'motivo_salida', 'sacado_por'];
var COLS_MOV       = ['id', 'fecha_hora', 'numero', 'lote_anterior', 'lote_nuevo', 'usuario', 'dispositivo'];
var COLS_LOTES     = ['lote'];
var COLS_TRABAJADORES = ['nombre'];

function hojaAnimales_(libro)     { return hoja_(libro, 'Animales_Ordeno', COLS_ANIMALES); }
function hojaMovimientos_(libro)  { return hoja_(libro, 'Movimientos', COLS_MOV); }
function hojaTrabajadores_(libro) { return hoja_(libro, 'Trabajadores', COLS_TRABAJADORES); }

function hojaConfigLotes_(libro) {
  var existia = !!libro.getSheetByName('Config_Lotes');
  var h = hoja_(libro, 'Config_Lotes', COLS_LOTES);
  if (!existia) {
    h.getRange(2, 1, 3, 1).setValues([[1], [2], [3]]);
  }
  return h;
}

/** Busca un animal por número. Devuelve {fila, valores} o null. */
function buscarAnimal_(h, numero) {
  if (h.getLastRow() < 2) return null;
  var vals = h.getRange(2, 1, h.getLastRow() - 1, COLS_ANIMALES.length).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(numero)) return {fila: i + 2, valores: vals[i]};
  }
  return null;
}

/**
 * GET ?accion=lotes_estado — todo lo que la app necesita para abrir sin más ida y vuelta:
 * animales activos con su lote actual, lotes disponibles, trabajadores y salidas recientes.
 */
function lotesEstado_(libro) {
  var hAnimales = hojaAnimales_(libro);
  var animales = [];
  var salidas = [];

  if (hAnimales.getLastRow() > 1) {
    var vals = hAnimales.getRange(2, 1, hAnimales.getLastRow() - 1, COLS_ANIMALES.length).getValues();
    var limite = new Date().getTime() - 30 * 24 * 60 * 60 * 1000;
    for (var i = 0; i < vals.length; i++) {
      var f = vals[i];
      var estado = String(f[2]);
      if (estado === 'activo') {
        animales.push({
          numero: f[0], nombre: f[1], origen: f[3],
          ingresado_por: f[5],
          fecha_entrada: f[4] instanceof Date ? f[4].toISOString() : String(f[4])
        });
      } else if (estado === 'fuera') {
        var t = f[6] instanceof Date ? f[6].getTime() : new Date(f[6]).getTime();
        if (t && t >= limite) {
          salidas.push({
            numero: f[0],
            nombre: f[1],
            fecha_salida: f[6] instanceof Date ? f[6].toISOString() : String(f[6]),
            motivo_salida: f[7],
            sacado_por: f[8]
          });
        }
      }
    }
  }

  // lote actual de cada animal = su movimiento con fecha_hora más reciente
  var loteMasReciente = {};
  var hMov = hojaMovimientos_(libro);
  if (hMov.getLastRow() > 1) {
    var mv = hMov.getRange(2, 1, hMov.getLastRow() - 1, COLS_MOV.length).getValues();
    for (var j = 0; j < mv.length; j++) {
      var numero = String(mv[j][2]);
      var fechaHora = mv[j][1] instanceof Date ? mv[j][1].getTime() : new Date(mv[j][1]).getTime();
      var actual = loteMasReciente[numero];
      if (!actual || fechaHora > actual.t) {
        loteMasReciente[numero] = {t: fechaHora, lote: mv[j][4]};
      }
    }
  }
  animales.forEach(function (a) {
    var m = loteMasReciente[String(a.numero)];
    a.lote = m ? m.lote : '';
  });

  var lotes = [];
  var hLotes = hojaConfigLotes_(libro);
  if (hLotes.getLastRow() > 1) {
    var lv = hLotes.getRange(2, 1, hLotes.getLastRow() - 1, 1).getValues();
    for (var k = 0; k < lv.length; k++) {
      if (lv[k][0] !== '') lotes.push(lv[k][0]);
    }
  }

  var trabajadores = [];
  var hTrab = hojaTrabajadores_(libro);
  if (hTrab.getLastRow() > 1) {
    var tv = hTrab.getRange(2, 1, hTrab.getLastRow() - 1, 1).getValues();
    for (var n = 0; n < tv.length; n++) {
      if (tv[n][0] !== '') trabajadores.push(tv[n][0]);
    }
  }

  return salida_({ok: true, animales: animales, lotes: lotes, trabajadores: trabajadores, salidas: salidas});
}

/**
 * POST accion=movimientos — recibe una lista de movimientos (asignar/mover de lote).
 * Ignora los que ya tengan ese id en la pestaña, así reintentar desde la tablet es seguro.
 */
function movimientos_(libro, datos) {
  var hMov = hojaMovimientos_(libro);
  var movs = datos.movimientos || [];
  if (!movs.length) return salida_({ok: true, guardados: 0});

  var existentes = {};
  if (hMov.getLastRow() > 1) {
    var idsExistentes = hMov.getRange(2, 1, hMov.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < idsExistentes.length; i++) existentes[String(idsExistentes[i][0])] = true;
  }

  var filas = [];
  for (var j = 0; j < movs.length; j++) {
    var m = movs[j];
    if (existentes[String(m.id)]) continue;
    filas.push([m.id, m.fecha_hora, m.numero, m.lote_anterior || '', m.lote_nuevo, m.usuario || '', m.dispositivo || '']);
    existentes[String(m.id)] = true;
  }

  if (filas.length) {
    hMov.getRange(hMov.getLastRow() + 1, 1, filas.length, COLS_MOV.length).setValues(filas);
  }
  return salida_({ok: true, guardados: filas.length});
}

/**
 * POST accion=animal_nuevo — crea un animal en Animales_Ordeno.
 * Si el número ya existe activo o fuera, no lo duplica: devuelve un error específico
 * para que la app avise o (si está fuera) ofrezca "Reactivar".
 */
function animalNuevo_(libro, datos) {
  var h = hojaAnimales_(libro);
  var encontrado = buscarAnimal_(h, datos.numero);
  if (encontrado) {
    var estado = String(encontrado.valores[2]);
    if (estado === 'activo') return salida_({ok: false, error: 'existeActivo'});
    if (estado === 'fuera')  return salida_({ok: false, error: 'existeFuera'});
  }

  h.appendRow([
    datos.numero, datos.nombre || '', 'activo', 'manual',
    new Date(), datos.usuario || '', '', '', ''
  ]);

  if (datos.lote) {
    var hMov = hojaMovimientos_(libro);
    hMov.appendRow([
      datos.id || Utilities.getUuid(), datos.fecha_hora || new Date(),
      datos.numero, '', datos.lote, datos.usuario || '', datos.dispositivo || ''
    ]);
  }

  return salida_({ok: true});
}

/**
 * POST accion=animal_salida — saca un animal de ordeño (secado, venta, muerte, otro).
 * Si ya estaba fuera, no vuelve a aplicar el cambio (reintentar es seguro).
 */
function animalSalida_(libro, datos) {
  var h = hojaAnimales_(libro);
  var encontrado = buscarAnimal_(h, datos.numero);
  if (!encontrado) return salida_({ok: false, error: 'noExiste'});
  if (String(encontrado.valores[2]) === 'fuera') return salida_({ok: true, ya: true});

  h.getRange(encontrado.fila, 3).setValue('fuera');
  h.getRange(encontrado.fila, 7).setValue(datos.fecha_hora ? new Date(datos.fecha_hora) : new Date());
  h.getRange(encontrado.fila, 8).setValue(datos.motivo || '');
  h.getRange(encontrado.fila, 9).setValue(datos.usuario || '');

  return salida_({ok: true});
}

/**
 * POST accion=animal_reactivar — vuelve a poner activo a un animal que había salido,
 * conservando la misma fila (y por lo tanto su historial de movimientos).
 * Si ya estaba activo, no hace nada (reintentar es seguro).
 */
function animalReactivar_(libro, datos) {
  var h = hojaAnimales_(libro);
  var encontrado = buscarAnimal_(h, datos.numero);
  if (!encontrado) return salida_({ok: false, error: 'noExiste'});
  if (String(encontrado.valores[2]) === 'activo') return salida_({ok: true, ya: true});

  h.getRange(encontrado.fila, 3).setValue('activo');
  return salida_({ok: true});
}

/**
 * POST accion=cargar_excel — aplica las decisiones ya tomadas en la pantalla de revisión.
 * "Mantener" y "confirmar que se queda" no cambian nada, así que solo llega lo que sí
 * cambia: nuevos (entran activos, sin lote) y salidas (con motivo, ya decidido en la app).
 * Reintentar es seguro: un número que ya existe no se duplica, y una salida ya aplicada no se repite.
 */
function cargarExcel_(libro, datos) {
  var h = hojaAnimales_(libro);
  var indice = {};
  if (h.getLastRow() > 1) {
    var vals = h.getRange(2, 1, h.getLastRow() - 1, COLS_ANIMALES.length).getValues();
    for (var i = 0; i < vals.length; i++) indice[String(vals[i][0])] = {fila: i + 2, valores: vals[i]};
  }

  var nuevos = datos.nuevos || [];
  var filasNuevas = [];
  var yaExistian = 0;
  for (var j = 0; j < nuevos.length; j++) {
    var n = nuevos[j];
    if (indice[String(n.numero)]) { yaExistian++; continue; }
    filasNuevas.push([n.numero, n.nombre || '', 'activo', 'excel', new Date(), 'Excel', '', '', '']);
  }
  if (filasNuevas.length) {
    h.getRange(h.getLastRow() + 1, 1, filasNuevas.length, COLS_ANIMALES.length).setValues(filasNuevas);
  }

  var salidas = datos.salidas || [];
  var sacados = 0;
  for (var k = 0; k < salidas.length; k++) {
    var s = salidas[k];
    var enc = indice[String(s.numero)];
    if (!enc || String(enc.valores[2]) === 'fuera') continue; // no existe o ya está fuera: no-op
    h.getRange(enc.fila, 3).setValue('fuera');
    h.getRange(enc.fila, 7).setValue(new Date());
    h.getRange(enc.fila, 8).setValue(s.motivo || '');
    h.getRange(enc.fila, 9).setValue(datos.usuario || '');
    sacados++;
  }

  return salida_({ok: true, creados: filasNuevas.length, yaExistian: yaExistian, sacados: sacados});
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
