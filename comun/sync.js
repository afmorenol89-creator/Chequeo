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

/** GET a la hoja, ej. pedirGet(url, 'lotes_estado'). */
function pedirGet(urlSheet, accion){
  if (!urlSheet) return Promise.reject(new Error('sin-url'));
  const sep = urlSheet.indexOf('?') >= 0 ? '&' : '?';
  return fetch(urlSheet + sep + 'accion=' + encodeURIComponent(accion)).then(r => {
    if (!r.ok) throw new Error('http-' + r.status);
    return r.text();
  }).then(t => {
    let j;
    try { j = JSON.parse(t); } catch (e) { throw new Error('respuesta-rara'); }
    if (!j.ok) throw new Error(j.error || 'error');
    return j;
  });
}
