"use strict";

/* =========================================================
   Envío/lectura contra el Apps Script (misma hoja que el chequeo)
   ========================================================= */

/** POST a la hoja. `cuerpo` va como texto plano para no disparar preflight de CORS. */
function pedir(urlSheet, cuerpo){
  if (!urlSheet) return Promise.reject(new Error('sin-url'));
  return fetch(urlSheet, {
    method: 'POST',
    headers: {'Content-Type': 'text/plain;charset=utf-8'},
    body: JSON.stringify(cuerpo)
  }).then(r => {
    if (!r.ok) throw new Error('http-' + r.status);
    return r.text();
  }).then(t => {
    let j;
    try { j = JSON.parse(t); } catch (e) { throw new Error('respuesta-rara'); }
    if (!j.ok) throw new Error(j.error || 'error');
    return j;
  });
}

/**
 * Pide una acción de "solo lectura", ej. pedirGet(url, 'lotes_estado').
 * Va por POST (igual que `pedir`), no por GET: los Web Apps de Apps Script
 * redirigen las peticiones GET a script.googleusercontent.com y ese salto
 * no siempre trae los encabezados CORS correctos al llamarlo desde otro
 * origen (como GitHub Pages) — con POST no pasa, así que evitamos GET del todo.
 */
function pedirGet(urlSheet, accion){
  return pedir(urlSheet, {accion});
}

/* =========================================================
   Cola de movimientos pendientes (para que la tablet funcione sin señal)
   ========================================================= */

/** Guarda un movimiento local (queda pendiente) y de una vez intenta subirlo si hay señal. */
function encolarMovimiento(urlSheet, mov){
  return dbGuardar('pendientes', mov).then(() => sincronizarPendientes(urlSheet));
}

/** Intenta subir todo lo pendiente. Ignora los que la hoja ya tenga (reintentar es seguro). */
function sincronizarPendientes(urlSheet){
  return dbTodo('pendientes').then(pendientes => {
    if (!pendientes.length) return {enviados: 0, quedan: 0};
    if (!navigator.onLine || !urlSheet) return {enviados: 0, quedan: pendientes.length};
    return pedir(urlSheet, {accion: 'movimientos', movimientos: pendientes}).then(() =>
      Promise.all(pendientes.map(p => dbBorrar('pendientes', p.id))).then(() => ({enviados: pendientes.length, quedan: 0}))
    ).catch(() => ({enviados: 0, quedan: pendientes.length}));
  });
}

/* =========================================================
   Cola de acciones individuales (nuevo animal, sacar, reactivar)
   Va aparte de la cola de movimientos: cada una se manda a su propia
   acción del Apps Script, no en lote.
   ========================================================= */
const ERRORES_CONFLICTO = ['existeActivo', 'existeFuera', 'noExiste'];

function encolarAccion(urlSheet, accion, datos){
  const item = {id: uuid(), accion, datos, creado: new Date().toISOString()};
  return dbGuardar('accionesPendientes', item).then(() => sincronizarAcciones(urlSheet));
}

/** Sube las acciones pendientes una por una. Un conflicto de negocio (ya existe, no existe)
 *  no tiene sentido reintentarlo solo, así que se quita de la cola y se reporta. */
function sincronizarAcciones(urlSheet){
  return dbTodo('accionesPendientes').then(items => {
    if (!items.length) return {enviados: 0, quedan: 0, conflictos: []};
    if (!navigator.onLine || !urlSheet) return {enviados: 0, quedan: items.length, conflictos: []};

    let enviados = 0;
    const conflictos = [];
    return items.reduce((cadena, item) => cadena.then(() => {
      const cuerpo = Object.assign({accion: item.accion}, item.datos);
      return pedir(urlSheet, cuerpo).then(() => {
        enviados++;
        return dbBorrar('accionesPendientes', item.id);
      }).catch(err => {
        const msg = err && err.message;
        if (ERRORES_CONFLICTO.indexOf(msg) >= 0){
          conflictos.push({item, error: msg});
          return dbBorrar('accionesPendientes', item.id);
        }
        // error de red u otro: se deja en la cola para reintentar después
      });
    }), Promise.resolve()).then(() =>
      dbTodo('accionesPendientes').then(quedan => ({enviados, quedan: quedan.length, conflictos}))
    );
  });
}

/** Intenta subir movimientos y acciones a la vez. Lo que usan los botones de "sincronizar". */
function sincronizarTodo(urlSheet){
  return Promise.all([sincronizarPendientes(urlSheet), sincronizarAcciones(urlSheet)]).then(([m, a]) => ({
    enviados: m.enviados + a.enviados,
    quedan: m.quedan + a.quedan,
    conflictos: a.conflictos
  }));
}

function contarPendientes(){
  return Promise.all([dbTodo('pendientes'), dbTodo('accionesPendientes')]).then(([p, a]) => p.length + a.length);
}
