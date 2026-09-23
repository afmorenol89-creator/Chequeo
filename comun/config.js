"use strict";

/* =========================================================
   Configuración fija de la finca (la usan el chequeo y los lotes)
   ========================================================= */
const FINCA = {
  urlSheet: 'https://script.google.com/macros/s/AKfycbxjQIXRXmviqj-2UmbVOvKQ8LbeQmUaQBieOgme70koli900lWPUga6j3gq3eydYXa3iQ/exec'
};

/* =========================================================
   PIN de la finca: se pide una vez por dispositivo y viaja en cada solicitud.
   Va en localStorage (mismo origen) para que lo compartan las dos herramientas.
   ========================================================= */
const PIN_CLAVE = 'finca_pin';

function leerPin(){
  try { return localStorage.getItem(PIN_CLAVE) || ''; } catch (e) { return ''; }
}
function guardarPin(pin){
  try { localStorage.setItem(PIN_CLAVE, pin); } catch (e) { /* sin almacenamiento: se pedirá de nuevo */ }
}
function borrarPin(){
  try { localStorage.removeItem(PIN_CLAVE); } catch (e) { /* nada */ }
}

/** La hoja dijo que el PIN no sirve: se olvida y se avisa a la pantalla para que lo pida otra vez. */
function pinRechazado(){
  borrarPin();
  window.dispatchEvent(new Event('pin-invalido'));
}

/** Cuadro para escribir el PIN. Resuelve con el texto escrito. */
function pedirPinModal(aviso){
  return new Promise(resolve => {
    const fondo = document.createElement('div');
    fondo.style.cssText = 'position:fixed;inset:0;z-index:200;background:rgba(22,33,26,.6);display:flex;' +
      'align-items:center;justify-content:center;padding:20px';
    const caja = document.createElement('div');
    caja.style.cssText = 'background:#fff;border-radius:16px;padding:22px;width:100%;max-width:380px;' +
      'font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;color:#16211a';
    caja.innerHTML =
      '<div style="font-size:1.2rem;font-weight:800;margin-bottom:6px">PIN de la finca</div>' +
      '<div style="font-size:.9rem;color:#5c6b5e;margin-bottom:14px">Se pide una sola vez en este dispositivo.</div>' +
      (aviso ? '<div style="background:#fbe9e8;color:#a3231c;border-radius:10px;padding:10px 12px;font-size:.9rem;margin-bottom:12px">' +
        aviso.replace(/</g, '&lt;') + '</div>' : '') +
      '<input type="password" id="_pinInput" autocomplete="off" placeholder="Escribe el PIN" ' +
        'style="width:100%;min-height:54px;padding:12px 14px;border:1px solid #d3dace;border-radius:11px;font-size:1.1rem;margin-bottom:12px">' +
      '<button type="button" id="_pinBtn" style="width:100%;min-height:56px;border-radius:12px;background:#1f7a3d;color:#fff;' +
        'font-weight:700;font-size:1rem;border:none;cursor:pointer">Entrar</button>';
    fondo.appendChild(caja);
    document.body.appendChild(fondo);

    const input = caja.querySelector('#_pinInput');
    const enviar = () => {
      const v = input.value.trim();
      if (!v){ input.focus(); return; }
      fondo.remove();
      resolve(v);
    };
    caja.querySelector('#_pinBtn').addEventListener('click', enviar);
    input.addEventListener('keydown', e => { if (e.key === 'Enter') enviar(); });
    setTimeout(() => input.focus(), 50);
  });
}

let _pinEnCurso = null;

/** Devuelve el PIN guardado o, si falta, lo pide y lo guarda. Nunca abre dos cuadros a la vez. */
function asegurarPin(aviso){
  const guardado = leerPin();
  if (guardado) return Promise.resolve(guardado);
  if (_pinEnCurso) return _pinEnCurso;
  _pinEnCurso = pedirPinModal(aviso).then(pin => {
    guardarPin(pin);
    _pinEnCurso = null;
    return pin;
  });
  return _pinEnCurso;
}
