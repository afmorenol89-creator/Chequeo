"use strict";

/* =========================================================
   Selección de "¿quién eres?" — se guarda en el dispositivo
   ========================================================= */

/**
 * Pinta el selector dentro de `contenedor` y llama a onElegido(nombre)
 * cuando el trabajador toca su nombre o escribe uno nuevo.
 */
function mostrarSelectorUsuario(contenedor, trabajadores, onElegido){
  const lista = (trabajadores || []).filter(Boolean);

  let html = '<div class="card"><h3 style="margin-bottom:14px">¿Quién eres?</h3>';
  if (lista.length){
    html += '<div style="display:flex;flex-direction:column;gap:10px;margin-bottom:16px">' +
      lista.map(n => '<button type="button" class="btn btn-borde" data-nombre="' + esc(n) + '">' + esc(n) + '</button>').join('') +
      '</div>';
  } else {
    html += '<div class="aviso">Todavía no hay trabajadores en la lista. Escribe tu nombre abajo.</div>';
  }
  html +=
    '<div class="campo">' +
      '<label>Otro nombre</label>' +
      '<input type="text" id="usuarioOtro" placeholder="Escribe tu nombre">' +
    '</div>' +
    '<button type="button" class="btn btn-verde" id="usuarioOtroBtn">Continuar</button>' +
    '</div>';
  contenedor.innerHTML = html;

  contenedor.querySelectorAll('[data-nombre]').forEach(b => {
    b.addEventListener('click', () => onElegido(b.getAttribute('data-nombre')));
  });
  contenedor.querySelector('#usuarioOtroBtn').addEventListener('click', () => {
    const v = contenedor.querySelector('#usuarioOtro').value.trim();
    if (v) onElegido(v);
  });
}
