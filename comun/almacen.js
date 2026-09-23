"use strict";

/* =========================================================
   Guardado local (IndexedDB) para las herramientas de la finca
   que no son el chequeo reproductivo (que tiene su propia base).
   ========================================================= */
const DB_NOMBRE_FINCA = 'lotes-ordeno';
const DB_VER_FINCA = 1;
let _dbFinca = null;

function abrirDB(){
  return new Promise((res, rej) => {
    if (_dbFinca) return res(_dbFinca);
    const req = indexedDB.open(DB_NOMBRE_FINCA, DB_VER_FINCA);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('config')) db.createObjectStore('config', {keyPath:'clave'});
      if (!db.objectStoreNames.contains('animales')) db.createObjectStore('animales', {keyPath:'numero'});
    };
    req.onsuccess = e => { _dbFinca = e.target.result; res(_dbFinca); };
    req.onerror = () => rej(req.error);
  });
}
function tx(store, modo){
  return abrirDB().then(db => db.transaction(store, modo).objectStore(store));
}
function dbGuardar(store, obj){
  return tx(store, 'readwrite').then(st => new Promise((res, rej) => {
    const r = st.put(obj); r.onsuccess = () => res(obj); r.onerror = () => rej(r.error);
  }));
}
function dbLeer(store, clave){
  return tx(store, 'readonly').then(st => new Promise((res, rej) => {
    const r = st.get(clave); r.onsuccess = () => res(r.result || null); r.onerror = () => rej(r.error);
  }));
}
function dbTodo(store){
  return tx(store, 'readonly').then(st => new Promise((res, rej) => {
    const r = st.getAll(); r.onsuccess = () => res(r.result || []); r.onerror = () => rej(r.error);
  }));
}

/** Config guardada por clave/valor: {clave:'urlSheet', valor:'...'} — mismo formato que usa el chequeo. */
function cargarConfig(){
  return dbTodo('config').then(filas => {
    const c = {urlSheet: '', usuario: '', dispositivo: ''};
    filas.forEach(f => { c[f.clave] = f.valor; });
    if (!c.dispositivo){
      c.dispositivo = (self.crypto && crypto.randomUUID)
        ? crypto.randomUUID()
        : ('disp_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8));
      dbGuardar('config', {clave: 'dispositivo', valor: c.dispositivo});
    }
    return c;
  });
}
function guardarConfigValor(clave, valor){
  return dbGuardar('config', {clave, valor});
}

/* =========================================================
   Utilidades compartidas (no son de almacenamiento, pero se
   necesitan en todas las pantallas y aquí se cargan primero).
   ========================================================= */
function esc(s){
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function norm(v){
  if (v === null || v === undefined) return '';
  return String(v).normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();
}
function fechaCorta(iso){
  const d = new Date(iso);
  if (isNaN(d)) return String(iso);
  return String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear();
}
let _toastT = null;
function toast(msg){
  const v = document.getElementById('_toast');
  if (v) v.remove();
  const el = document.createElement('div');
  el.className = 'toast'; el.id = '_toast'; el.textContent = msg;
  el.setAttribute('role', 'status');
  document.body.appendChild(el);
  clearTimeout(_toastT);
  _toastT = setTimeout(() => el.remove(), 2600);
}
