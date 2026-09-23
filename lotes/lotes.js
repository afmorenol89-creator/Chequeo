"use strict";

/* =========================================================
   Estado en memoria
   ========================================================= */
let config = {urlSheet: '', usuario: '', dispositivo: ''};
let vista = 'cargando';
let excelCargado = null;     // {archivo, animales:[{numero,nombre}]}
let estadoServidor = null;   // resultado de lotes_estado (animales activos, trabajadores, etc.)
let revisionActual = null;   // {nuevos, noVienen} calculado al comparar el Excel con la hoja
let decisionesSalida = {};   // numero (string) -> {tipo, texto}

let animalesPorNumero = {};   // numero (string) -> {numero, nombre, lote, origen, ingresado_por, fecha_entrada}
let lotesDisponibles = [];    // ej. [1,2,3]
let movimientosRecientes = []; // últimos movimientos de toda la finca (para el historial de cada animal)
let salidasRecientes = [];     // animales sacados en los últimos 30 días
let animalSeleccionado = null; // numero (string) abierto en la ficha
let rapidoLote = null;         // lote destino elegido en modo rápido
let rapidoHechos = {};         // numero (string) -> true, ya tocado en esta sesión de modo rápido
let nuevoLoteElegido = '';     // lote elegido en la pantalla de Nuevo animal

const COL_NUM = 'numero de animal'; // ya pasado por norm()
const COL_NOM = 'nombre';

/* =========================================================
   Navegación entre pantallas
   ========================================================= */
function ver(v){
  vista = v;
  ['quien', 'ajustes', 'inicio', 'cargar', 'revision', 'resultado', 'ficha', 'rapido', 'nuevo', 'salidas'].forEach(n => {
    const el = document.getElementById('v-' + n);
    if (el) el.classList.toggle('oculto', n !== v);
  });
  window.scrollTo(0, 0);
}

/** Vuelve a la lista principal usando lo que ya hay en memoria (sin ir a la hoja). */
function irAInicio(){
  pintarListaPrincipal();
  ver('inicio');
}

function actualizarEstadoSync(){
  const el = document.getElementById('net'), t = document.getElementById('netTxt');
  contarPendientes().then(n => {
    if (!navigator.onLine){
      el.classList.add('off');
      t.textContent = 'Sin conexión' + (n ? ' · ' + n + ' pendiente' + (n === 1 ? '' : 's') : '');
    } else if (n > 0){
      el.classList.add('off');
      t.textContent = n + ' cambio' + (n === 1 ? '' : 's') + ' pendiente' + (n === 1 ? '' : 's') + ' por subir';
    } else {
      el.classList.remove('off');
      t.textContent = 'Todo al día';
    }
  });
}
window.addEventListener('online', () => {
  actualizarEstadoSync();
  sincronizarTodo(config.urlSheet).then(actualizarEstadoSync);
});
window.addEventListener('offline', actualizarEstadoSync);

/* =========================================================
   Arranque: base local, hoja configurada, usuario del dispositivo
   ========================================================= */
function iniciar(){
  abrirDB().then(() => cargarConfig()).then(c => {
    config = c;
    config.urlSheet = FINCA.urlSheet; // la URL es fija, no se guarda por dispositivo
    actualizarEstadoSync();
    seguirInicio();
  });
}

/** El PIN cambió o no sirvió: se vuelve a pedir. */
window.addEventListener('pin-invalido', () => {
  accesoInicial('El PIN no es correcto. Escríbelo de nuevo.');
});

/**
 * Orden de arranque: PIN (una vez por dispositivo, con descarga de todos los datos de la hoja)
 * → "¿quién eres?" (una vez por dispositivo) → lista principal.
 */
function seguirInicio(){
  if (!leerPin()){ accesoInicial(); return; }
  if (!config.usuario){ mostrarSelectorUsuarioInicial(); return; }
  document.getElementById('nombreUsuario').textContent = config.usuario;
  ver('inicio');
  if (lotesDisponibles.length || Object.keys(animalesPorNumero).length) irAInicio();
  else refrescarYPintarLista();
}

/** Pide el PIN y, con él, descarga de la hoja todo lo que la app necesita (dispositivo nuevo). */
function accesoInicial(aviso){
  ver('quien');
  const cont = document.getElementById('contQuien');
  if (!navigator.onLine){
    cont.innerHTML = '<div class="card"><h3 style="margin-bottom:8px">Falta conexión</h3>' +
      '<div class="aviso">La primera vez, esta tablet necesita internet para descargar los datos de la finca.</div>' +
      '<button type="button" class="btn btn-verde" id="btnReintentarAcceso">Reintentar</button></div>';
    document.getElementById('btnReintentarAcceso').addEventListener('click', () => accesoInicial());
    return;
  }
  cont.innerHTML = '';
  asegurarPin(aviso).then(() => {
    cont.innerHTML = '<div class="aviso">Descargando los datos de la hoja…</div>';
    return pedirGet(config.urlSheet, 'lotes_estado');
  }).then(j => aplicarEstadoServidor(j)).then(() => seguirInicio())
    .catch(e => {
      if (e.message === 'pinInvalido') return; // el evento 'pin-invalido' ya lo vuelve a pedir
      cont.innerHTML = '<div class="card"><h3 style="margin-bottom:8px">No se pudo descargar</h3>' +
        '<div class="aviso malo">' + esc(e.message || 'Revisa la conexión.') + '</div>' +
        '<button type="button" class="btn btn-verde" id="btnReintentarAcceso">Reintentar</button></div>';
      document.getElementById('btnReintentarAcceso').addEventListener('click', () => accesoInicial());
    });
}

function mostrarSelectorUsuarioInicial(){
  ver('quien');
  const cont = document.getElementById('contQuien');
  dbLeer('config', 'cacheTrabajadores').then(c => {
    if (c) return c.valor || [];
    if (!navigator.onLine) return [];
    return pedirGet(config.urlSheet, 'lotes_estado').then(j => j.trabajadores || []).catch(() => []);
  }).then(lista => mostrarSelectorUsuario(cont, lista, elegirUsuario));
}

function elegirUsuario(nombre){
  config.usuario = nombre;
  guardarConfigValor('usuario', nombre).then(() => seguirInicio());
}

function cambiarUsuario(){
  config.usuario = '';
  guardarConfigValor('usuario', '').then(() => mostrarSelectorUsuarioInicial());
}

/* =========================================================
   Cargar Excel
   ========================================================= */
document.addEventListener('DOMContentLoaded', () => {
  const input = document.getElementById('archivoExcel');
  input.addEventListener('change', e => { if (e.target.files[0]) leerExcel(e.target.files[0]); });

  const drop = document.getElementById('drop');
  drop.addEventListener('dragover', e => e.preventDefault());
  drop.addEventListener('drop', e => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) leerExcel(e.dataTransfer.files[0]);
  });

  document.getElementById('buscador').addEventListener('input', pintarListaPrincipal);

  document.getElementById('net').addEventListener('click', () => {
    if (!navigator.onLine){ toast('Sin conexión'); return; }
    sincronizarTodo(config.urlSheet).then(r => {
      actualizarEstadoSync();
      if (r.enviados) toast(r.enviados + ' cambio' + (r.enviados === 1 ? '' : 's') + ' subido' + (r.enviados === 1 ? '' : 's'));
      if (r.conflictos && r.conflictos.length) toast(r.conflictos.length + ' acción(es) no se pudieron aplicar — revisa la lista.');
    });
  });

  iniciar();
});

function leerExcel(file){
  const info = document.getElementById('cargaInfo');
  info.innerHTML = '<div class="aviso">Leyendo el archivo…</div>';
  const fr = new FileReader();
  fr.onerror = () => {
    info.innerHTML = '<div class="aviso malo"><b>No se pudo leer el archivo</b>Inténtalo de nuevo.</div>';
  };
  fr.onload = ev => {
    try {
      const wb = XLSX.read(new Uint8Array(ev.target.result), {type: 'array'});
      const hoja = wb.Sheets[wb.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json(hoja, {header: 1, raw: true, defval: ''});
      const conDatos = filas.filter(f => f.some(c => c !== '' && c !== null && c !== undefined));
      if (conDatos.length < 2) throw {msg: 'El archivo está vacío.'};

      const heads = conDatos[0].map(h => (h === null || h === undefined) ? '' : String(h).trim());
      const iNum = heads.findIndex(h => norm(h) === COL_NUM);
      const iNom = heads.findIndex(h => norm(h) === COL_NOM);
      if (iNum < 0 || iNom < 0){
        throw {msg: 'El archivo debe tener las columnas "Número de animal" y "Nombre". Encontré: ' +
          (heads.filter(Boolean).join(', ') || '(sin encabezados)')};
      }

      const cuerpo = conDatos.slice(1);
      const vistos = {};
      const errores = [];
      const animales = [];

      cuerpo.forEach((f, i) => {
        const filaNum = i + 2; // fila 1 = encabezado
        const numRaw = f[iNum], nomRaw = f[iNom];
        const numVacio = numRaw === '' || numRaw === null || numRaw === undefined;
        const nomVacio = nomRaw === '' || nomRaw === null || nomRaw === undefined;
        if (numVacio && nomVacio) return; // fila en blanco

        let numero;
        if (typeof numRaw === 'number') {
          numero = numRaw;
        } else {
          const txt = String(numRaw).trim();
          numero = /^\d+$/.test(txt) ? parseInt(txt, 10) : NaN;
        }
        if (!Number.isInteger(numero)){
          errores.push('Fila ' + filaNum + ': "' + numRaw + '" no es un número entero.');
          return;
        }
        if (vistos[numero] !== undefined){
          errores.push('Fila ' + filaNum + ': el número ' + numero + ' está repetido (también en la fila ' + vistos[numero] + ').');
          return;
        }
        vistos[numero] = filaNum;
        const nombre = nomVacio ? '' : String(nomRaw).trim();
        animales.push({numero, nombre});
      });

      if (errores.length){
        info.innerHTML = '<div class="aviso malo"><b>No se cargó el archivo</b>' + errores.map(esc).join('<br>') + '</div>';
        return;
      }
      if (!animales.length) throw {msg: 'No encontré animales en el archivo.'};

      excelCargado = {archivo: file.name, animales};
      info.innerHTML = '<div class="aviso bien"><b>' + esc(file.name) + '</b>' + animales.length + ' animales en el archivo.</div>';
      prepararRevision();

    } catch (err){
      info.innerHTML = '<div class="aviso malo"><b>No se pudo leer el archivo</b>' +
        esc((err && err.msg) || 'Revisa que sea un Excel válido.') + '</div>';
    }
  };
  fr.readAsArrayBuffer(file);
}

/* =========================================================
   Revisión: compara el Excel contra lo que hay activo en la hoja
   ========================================================= */
function prepararRevision(){
  const info = document.getElementById('cargaInfo');
  if (!navigator.onLine){
    info.innerHTML += '<div class="aviso malo"><b>Sin conexión</b>Para revisar el Excel necesito comparar contra la hoja de Google. Conéctate e inténtalo de nuevo.</div>';
    return;
  }
  info.innerHTML += '<div class="aviso">Comparando contra la hoja…</div>';

  pedirGet(config.urlSheet, 'lotes_estado').then(j => {
    estadoServidor = j;
    decisionesSalida = {};
    pintarRevision();
    ver('revision');
  }).catch(e => {
    info.innerHTML += '<div class="aviso malo"><b>No se pudo comparar</b>' + esc(e.message || 'Revisa la conexión.') + '</div>';
  });
}

function pintarRevision(){
  const activos = estadoServidor.animales || [];
  const numerosExcel = {};
  excelCargado.animales.forEach(a => { numerosExcel[String(a.numero)] = true; });
  const numerosActivos = {};
  activos.forEach(a => { numerosActivos[String(a.numero)] = a; });

  const nuevos = excelCargado.animales.filter(a => !numerosActivos[String(a.numero)]);
  const sinCambio = excelCargado.animales.filter(a => numerosActivos[String(a.numero)]);
  const noVienen = activos.filter(a => !numerosExcel[String(a.numero)]);
  const noVienenExcel = noVienen.filter(a => a.origen !== 'manual');
  const noVienenManual = noVienen.filter(a => a.origen === 'manual');

  revisionActual = {nuevos, noVienen};

  document.getElementById('resumenRevision').innerHTML =
    '<b>' + esc(excelCargado.archivo) + '</b>' +
    '<div style="margin-top:8px;display:flex;flex-wrap:wrap;gap:14px;font-size:0.9rem;color:var(--muted)">' +
    '<span>🟢 ' + nuevos.length + ' nuevos</span>' +
    '<span>⚪ ' + sinCambio.length + ' sin cambio</span>' +
    '<span>🟠 ' + noVienen.length + ' no vienen en el Excel</span>' +
    '</div>';

  document.getElementById('grupoNuevos').innerHTML = nuevos.length ?
    '<div class="grupo-titulo">Nuevos en el Excel (entran activos, sin lote)</div>' +
    '<div class="lista-nombres">' + nuevos.map(a =>
      '<div class="fila-animal"><span class="num">' + esc(a.numero) + '</span><span class="nom">' + esc(a.nombre || '(sin nombre)') + '</span></div>'
    ).join('') + '</div>' : '';

  document.getElementById('grupoNoVienenExcel').innerHTML = pintarGrupoSalida(
    noVienenExcel, 'Activos que no vienen en el Excel',
    'Si no escoges un motivo, se mantienen activos.'
  );
  document.getElementById('grupoManual').innerHTML = pintarGrupoSalida(
    noVienenManual, 'Ingresados a mano que no están en el Excel',
    'Ingresados por un trabajador, no vienen en este archivo. Si no escoges un motivo, se mantienen activos.'
  );

  cablearMotivos();
}

function pintarGrupoSalida(lista, titulo, ayuda){
  if (!lista.length) return '';
  return '<div class="grupo-titulo">' + esc(titulo) + '</div>' +
    '<div class="aviso" style="margin-bottom:10px">' + esc(ayuda) + '</div>' +
    '<div class="lista-nombres">' + lista.map(a => {
      const meta = a.origen === 'manual'
        ? 'Ingresado por ' + esc(a.ingresado_por || '?') + (a.fecha_entrada ? ' el ' + esc(fechaCorta(a.fecha_entrada)) : '')
        : '';
      return '<div class="fila-animal" style="flex-direction:column;align-items:stretch;gap:8px">' +
        '<div style="display:flex;gap:12px;align-items:center">' +
          '<span class="num">' + esc(a.numero) + '</span>' +
          '<span class="nom">' + esc(a.nombre || '(sin nombre)') + '</span>' +
        '</div>' +
        (meta ? '<div class="meta">' + meta + '</div>' : '') +
        '<div class="motivos" data-numero="' + esc(a.numero) + '">' +
          '<button type="button" class="motivo-btn mantener" data-motivo="" aria-pressed="true">Mantener</button>' +
          '<button type="button" class="motivo-btn" data-motivo="Secado" aria-pressed="false">Secado</button>' +
          '<button type="button" class="motivo-btn" data-motivo="Venta" aria-pressed="false">Venta</button>' +
          '<button type="button" class="motivo-btn" data-motivo="Muerte" aria-pressed="false">Muerte</button>' +
          '<button type="button" class="motivo-btn" data-motivo="Otro" aria-pressed="false">Otro</button>' +
        '</div>' +
        '<input type="text" class="oculto" placeholder="¿Cuál?" data-otro-de="' + esc(a.numero) + '">' +
      '</div>';
    }).join('') + '</div>';
}

function cablearMotivos(){
  document.querySelectorAll('.motivos').forEach(grupo => {
    const numero = grupo.getAttribute('data-numero');
    grupo.querySelectorAll('.motivo-btn').forEach(btn => {
      btn.addEventListener('click', () => elegirMotivo(numero, btn.dataset.motivo, grupo));
    });
  });
  document.querySelectorAll('[data-otro-de]').forEach(inp => {
    inp.addEventListener('input', () => {
      const numero = inp.getAttribute('data-otro-de');
      if (decisionesSalida[numero]) decisionesSalida[numero].texto = inp.value.trim();
    });
  });
}

function elegirMotivo(numero, motivo, grupo){
  if (motivo === ''){
    delete decisionesSalida[numero];
  } else {
    const previo = (decisionesSalida[numero] && decisionesSalida[numero].texto) || '';
    decisionesSalida[numero] = {tipo: motivo, texto: previo};
  }
  grupo.querySelectorAll('.motivo-btn').forEach(b => {
    b.setAttribute('aria-pressed', b.dataset.motivo === motivo ? 'true' : 'false');
  });
  const otroInput = grupo.nextElementSibling;
  if (otroInput && otroInput.hasAttribute('data-otro-de')){
    otroInput.classList.toggle('oculto', motivo !== 'Otro');
  }
}

/* =========================================================
   Confirmar: envía nuevos + salidas decididas
   ========================================================= */
function confirmarCargaExcel(){
  const btn = document.getElementById('btnConfirmarRevision');
  btn.disabled = true; btn.textContent = 'Guardando…';

  const salidas = Object.keys(decisionesSalida).map(numero => {
    const d = decisionesSalida[numero];
    const motivo = d.tipo === 'Otro' ? ('Otro: ' + (d.texto || '')) : d.tipo;
    return {numero, motivo};
  });

  const cuerpo = {
    accion: 'cargar_excel',
    usuario: config.usuario,
    nuevos: revisionActual.nuevos.map(a => ({numero: a.numero, nombre: a.nombre})),
    salidas
  };

  pedir(config.urlSheet, cuerpo).then(j => {
    document.getElementById('resultadoTxt').innerHTML =
      '<b>Excel cargado</b>' +
      '<div style="margin-top:8px">' + j.creados + ' animales nuevos, ' + j.sacados + ' salidas' +
      (j.yaExistian ? ', ' + j.yaExistian + ' ya existían' : '') + '.</div>';
    excelCargado = null; estadoServidor = null; revisionActual = null; decisionesSalida = {};
    document.getElementById('cargaInfo').innerHTML = '';
    document.getElementById('archivoExcel').value = '';
    ver('resultado');
    refrescarYPintarLista();
  }).catch(e => {
    toast('No se pudo guardar: ' + (e.message || 'revisa la conexión'));
  }).then(() => {
    btn.disabled = false; btn.textContent = 'Guardar cambios';
  });
}

/* =========================================================
   Lista principal
   ========================================================= */

/** Trae el estado de la hoja (o el caché local si no hay señal) y pinta la lista. */
function refrescarYPintarLista(){
  const traer = navigator.onLine && config.urlSheet && leerPin()
    ? pedirGet(config.urlSheet, 'lotes_estado').then(aplicarEstadoServidor).catch(() => cargarCacheLocal())
    : cargarCacheLocal();

  return traer.then(aplicarPendientesLocales).then(pintarListaPrincipal);
}

/** Pone en memoria y guarda en el dispositivo todo lo que devolvió la hoja. */
function aplicarEstadoServidor(j){
  lotesDisponibles = j.lotes || [];
  movimientosRecientes = j.movimientos_recientes || [];
  salidasRecientes = j.salidas || [];
  animalesPorNumero = {};
  (j.animales || []).forEach(a => { animalesPorNumero[String(a.numero)] = a; });
  return guardarCacheServidor(j);
}

function guardarCacheServidor(j){
  return dbTodo('animales')
    .then(previos => Promise.all(previos.map(a => dbBorrar('animales', a.numero))))
    .then(() => Promise.all([
      guardarConfigValor('cacheLotes', j.lotes || []),
      guardarConfigValor('cacheMovimientos', j.movimientos_recientes || []),
      guardarConfigValor('cacheSalidas', j.salidas || []),
      guardarConfigValor('cacheTrabajadores', j.trabajadores || []),
      Promise.all((j.animales || []).map(a => dbGuardar('animales', a)))
    ]));
}

function cargarCacheLocal(){
  return Promise.all([
    dbTodo('animales'), dbLeer('config', 'cacheLotes'),
    dbLeer('config', 'cacheMovimientos'), dbLeer('config', 'cacheSalidas')
  ]).then(([animales, lotesC, movC, salC]) => {
      animalesPorNumero = {};
      animales.forEach(a => { animalesPorNumero[String(a.numero)] = a; });
      lotesDisponibles = (lotesC && lotesC.valor) || [];
      movimientosRecientes = (movC && movC.valor) || [];
      salidasRecientes = (salC && salC.valor) || [];
    });
}

/** Aplica sobre animalesPorNumero los movimientos que este dispositivo aún no ha subido. */
function aplicarPendientesLocales(){
  return dbTodo('pendientes').then(pendientes => {
    pendientes
      .slice()
      .sort((a, b) => new Date(a.fecha_hora) - new Date(b.fecha_hora))
      .forEach(p => {
        const a = animalesPorNumero[String(p.numero)];
        if (a) a.lote = p.lote_nuevo;
      });
  });
}

function pintarListaPrincipal(){
  const cont = document.getElementById('listaPrincipal');
  const filtro = norm(document.getElementById('buscador').value);
  const todos = Object.keys(animalesPorNumero).map(k => animalesPorNumero[k])
    .filter(a => !filtro || norm(a.numero).indexOf(filtro) >= 0 || norm(a.nombre).indexOf(filtro) >= 0)
    .sort((a, b) => Number(a.numero) - Number(b.numero));

  const banner = salidasRecientes.length
    ? '<button type="button" class="btn btn-borde" style="margin-bottom:16px" onclick="verSalidas()">🗂 ' +
        salidasRecientes.length + ' salida' + (salidasRecientes.length === 1 ? '' : 's') + ' reciente' +
        (salidasRecientes.length === 1 ? '' : 's') + '</button>'
    : '';

  if (!todos.length){
    cont.innerHTML = banner + '<div class="vacio"><span class="ic">🥛</span>' +
      (filtro ? 'No encontré animales con ese número o nombre.' : 'No hay animales activos todavía. Carga un Excel para empezar.') +
      '</div>';
    return;
  }

  let html = banner;
  const sinLote = todos.filter(a => !a.lote);
  if (sinLote.length){
    html += '<div class="grupo-titulo">Sin lote (' + sinLote.length + ')</div>' +
      '<div class="lista-nombres">' + sinLote.map(filaAnimalHtml).join('') + '</div>';
  }
  lotesDisponibles.forEach(lote => {
    const deEsteLote = todos.filter(a => String(a.lote) === String(lote));
    if (!deEsteLote.length && filtro) return; // con búsqueda activa, no mostrar lotes vacíos
    html += '<div class="grupo-titulo">Lote ' + esc(lote) + ' (' + deEsteLote.length + ')</div>';
    html += deEsteLote.length
      ? '<div class="lista-nombres">' + deEsteLote.map(filaAnimalHtml).join('') + '</div>'
      : '<div class="aviso">Sin animales en este lote.</div>';
  });

  cont.innerHTML = html;
  cont.querySelectorAll('[data-abrir]').forEach(el => {
    el.addEventListener('click', () => abrirFicha(el.getAttribute('data-abrir')));
  });
}

function filaAnimalHtml(a){
  return '<div class="fila-animal" data-abrir="' + esc(a.numero) + '" style="cursor:pointer">' +
    '<span class="num">' + esc(a.numero) + '</span>' +
    '<span class="nom">' + esc(a.nombre || '(sin nombre)') + '</span>' +
  '</div>';
}

/* =========================================================
   Ficha del animal
   ========================================================= */
function abrirFicha(numero){
  animalSeleccionado = String(numero);
  pintarFicha();
  ver('ficha');
}

function pintarFicha(){
  const a = animalesPorNumero[animalSeleccionado];
  if (!a){ irAInicio(); return; }
  const cont = document.getElementById('fichaCont');

  const historial = movimientosRecientes.filter(m => String(m.numero) === animalSeleccionado).slice(0, 6);

  let html = '<div class="card">' +
    '<div style="font-size:1.3rem;font-weight:800">' + esc(a.numero) + ' · ' + esc(a.nombre || '(sin nombre)') + '</div>' +
    '<div class="top-sub">' + (a.lote ? 'Lote actual: ' + esc(a.lote) : 'Sin lote') + '</div>' +
  '</div>';

  html += '<div class="grupo-titulo">Mover a</div>' +
    '<div class="motivos" style="margin-bottom:16px">' +
    lotesDisponibles.map(lote => {
      const esActual = String(a.lote) === String(lote);
      return '<button type="button" class="motivo-btn' + (esActual ? ' mantener' : '') + '" ' +
        'aria-pressed="' + (esActual ? 'true' : 'false') + '"' + (esActual ? ' disabled' : '') +
        ' data-lote="' + esc(lote) + '">Lote ' + esc(lote) + '</button>';
    }).join('') + '</div>';

  html += '<div class="grupo-titulo">Sacar de ordeño</div>' +
    '<button type="button" class="btn btn-borde" id="btnAbrirSacar">Sacar de ordeño</button>' +
    '<div id="sacarBloque" class="oculto">' +
      '<div class="motivos" style="margin-bottom:8px">' +
        '<button type="button" class="motivo-btn" data-motivo-directo="Secado">Secado</button>' +
        '<button type="button" class="motivo-btn" data-motivo-directo="Venta">Venta</button>' +
        '<button type="button" class="motivo-btn" data-motivo-directo="Muerte">Muerte</button>' +
        '<button type="button" class="motivo-btn" id="btnMotivoOtro">Otro</button>' +
      '</div>' +
      '<div id="otroBloque" class="oculto">' +
        '<input type="text" id="sacarOtroTexto" placeholder="¿Cuál?" style="margin-bottom:8px">' +
        '<button type="button" class="btn btn-ambar" id="btnConfirmarOtro">Sacar por este motivo</button>' +
      '</div>' +
    '</div>';

  html += '<div class="grupo-titulo">Historial</div>';
  html += historial.length
    ? '<div class="lista-nombres">' + historial.map(m =>
        '<div class="fila-animal" style="flex-direction:column;align-items:stretch;gap:4px">' +
          '<div>' + (m.lote_anterior ? 'Lote ' + esc(m.lote_anterior) + ' → ' : '') + 'Lote ' + esc(m.lote_nuevo) + '</div>' +
          '<div class="meta">' + esc(fechaCorta(m.fecha_hora)) + ' · ' + esc(m.usuario || '?') + '</div>' +
        '</div>'
      ).join('') + '</div>'
    : '<div class="aviso">Sin movimientos todavía.</div>';

  cont.innerHTML = html;
  cont.querySelectorAll('[data-lote]').forEach(btn => {
    btn.addEventListener('click', () => confirmarMoverLote(btn.getAttribute('data-lote')));
  });
  document.getElementById('btnAbrirSacar').addEventListener('click', () => {
    document.getElementById('sacarBloque').classList.toggle('oculto');
  });
  cont.querySelectorAll('[data-motivo-directo]').forEach(btn => {
    btn.addEventListener('click', () => ejecutarSacar(a, btn.getAttribute('data-motivo-directo')));
  });
  document.getElementById('btnMotivoOtro').addEventListener('click', () => {
    document.getElementById('otroBloque').classList.toggle('oculto');
  });
  document.getElementById('btnConfirmarOtro').addEventListener('click', () => {
    const texto = document.getElementById('sacarOtroTexto').value.trim();
    if (!texto){ toast('Escribe el motivo'); return; }
    ejecutarSacar(a, 'Otro: ' + texto);
  });
}

/** Saca al animal de ordeño (con confirmación) y lo quita de la lista de activos. */
function ejecutarSacar(a, motivo){
  if (!confirm('¿Sacar a ' + a.numero + ' ' + (a.nombre || '') + ' por ' + motivo.toLowerCase() + '?')) return;

  const datos = {
    numero: a.numero, motivo, usuario: config.usuario,
    dispositivo: config.dispositivo, fecha_hora: new Date().toISOString()
  };
  const aplicarLocal = () => {
    delete animalesPorNumero[animalSeleccionado];
    salidasRecientes.unshift({
      numero: a.numero, nombre: a.nombre, motivo_salida: motivo,
      fecha_salida: datos.fecha_hora, sacado_por: config.usuario
    });
    toast('Sacado de ordeño');
    irAInicio();
  };

  if (navigator.onLine && config.urlSheet){
    pedir(config.urlSheet, Object.assign({accion: 'animal_salida'}, datos)).then(aplicarLocal)
      .catch(e => toast('No se pudo sacar: ' + (e.message || 'revisa la conexión')));
  } else {
    aplicarLocal();
    encolarAccion(config.urlSheet, 'animal_salida', datos).then(actualizarEstadoSync);
  }
}

function confirmarMoverLote(loteNuevo){
  const a = animalesPorNumero[animalSeleccionado];
  if (!a) return;
  const desde = a.lote ? ('del Lote ' + a.lote) : 'sin lote';
  if (!confirm('¿Mover ' + a.numero + ' ' + (a.nombre || '') + ' ' + desde + ' al Lote ' + loteNuevo + '?')) return;

  registrarMovimiento(a, loteNuevo);
  pintarFicha();
}

/** Aplica el movimiento de forma optimista en memoria y lo encola para subir. */
function registrarMovimiento(a, loteNuevo){
  const mov = {
    id: uuid(), fecha_hora: new Date().toISOString(),
    numero: a.numero, lote_anterior: a.lote || '', lote_nuevo: loteNuevo,
    usuario: config.usuario, dispositivo: config.dispositivo
  };
  a.lote = loteNuevo;
  movimientosRecientes.unshift({
    numero: mov.numero, fecha_hora: mov.fecha_hora,
    lote_anterior: mov.lote_anterior, lote_nuevo: mov.lote_nuevo, usuario: mov.usuario
  });
  toast('Movido al Lote ' + loteNuevo);
  encolarMovimiento(config.urlSheet, mov).then(actualizarEstadoSync);
}

/* =========================================================
   Modo rápido: un lote destino, varios animales a golpe de toque
   ========================================================= */
function iniciarModoRapido(){
  if (!lotesDisponibles.length){ toast('No hay lotes configurados todavía.'); return; }
  rapidoLote = null;
  rapidoHechos = {};
  pintarRapidoElegirLote();
  ver('rapido');
}

function pintarRapidoElegirLote(){
  const cont = document.getElementById('rapidoCont');
  cont.innerHTML =
    '<div class="card"><b>Modo rápido</b><div class="top-sub">Elige primero el lote destino</div></div>' +
    '<div class="motivos" style="margin-bottom:16px">' +
    lotesDisponibles.map(lote => '<button type="button" class="motivo-btn" data-lote-rapido="' + esc(lote) + '">Lote ' + esc(lote) + '</button>').join('') +
    '</div>' +
    '<button type="button" class="btn btn-borde" onclick="irAInicio()">Cancelar</button>';
  cont.querySelectorAll('[data-lote-rapido]').forEach(b => {
    b.addEventListener('click', () => { rapidoLote = b.getAttribute('data-lote-rapido'); pintarRapidoTocar(); });
  });
}

function pintarRapidoTocar(){
  const cont = document.getElementById('rapidoCont');
  const todos = Object.keys(animalesPorNumero).map(k => animalesPorNumero[k]).sort((a, b) => Number(a.numero) - Number(b.numero));

  cont.innerHTML =
    '<div class="card"><b>Modo rápido · Lote ' + esc(rapidoLote) + '</b><div class="top-sub">Toca cada animal que entra a este lote</div></div>' +
    '<div class="lista-nombres">' + todos.map(a => {
      const hecho = !!rapidoHechos[String(a.numero)];
      return '<button type="button" class="fila-animal" style="width:100%;text-align:left' + (hecho ? ';opacity:.5' : '') + '" ' +
        'data-toca="' + esc(a.numero) + '"' + (hecho ? ' disabled' : '') + '>' +
        '<span class="num">' + esc(a.numero) + '</span>' +
        '<span class="nom">' + esc(a.nombre || '(sin nombre)') +
          (hecho ? ' ✓' : (a.lote ? ' · Lote ' + esc(a.lote) : '')) + '</span>' +
      '</button>';
    }).join('') + '</div>' +
    '<button type="button" class="btn btn-verde" style="margin-top:14px" onclick="irAInicio()">Terminar</button>';

  cont.querySelectorAll('[data-toca]').forEach(btn => {
    btn.addEventListener('click', () => tocarRapido(btn.getAttribute('data-toca')));
  });
}

function tocarRapido(numero){
  const a = animalesPorNumero[numero];
  if (!a || rapidoHechos[numero]) return;
  rapidoHechos[numero] = true;
  registrarMovimiento(a, rapidoLote);
  pintarRapidoTocar();
}

/* =========================================================
   Nuevo animal (disponible para todos)
   ========================================================= */
function abrirNuevoAnimal(){
  document.getElementById('nuevoNumero').value = '';
  document.getElementById('nuevoNombre').value = '';
  document.getElementById('nuevoAviso').innerHTML = '';
  document.getElementById('btnCrearAnimal').disabled = false;
  pintarNuevoLotes('');
  ver('nuevo');
}

function pintarNuevoLotes(seleccionado){
  nuevoLoteElegido = seleccionado;
  const cont = document.getElementById('nuevoLotes');
  cont.innerHTML = lotesDisponibles.map(lote => {
    const activo = String(nuevoLoteElegido) === String(lote);
    return '<button type="button" class="motivo-btn' + (activo ? ' mantener' : '') + '" ' +
      'aria-pressed="' + (activo ? 'true' : 'false') + '" data-lote-nuevo="' + esc(lote) + '">Lote ' + esc(lote) + '</button>';
  }).join('');
  cont.querySelectorAll('[data-lote-nuevo]').forEach(b => {
    b.addEventListener('click', () => {
      const lote = b.getAttribute('data-lote-nuevo');
      pintarNuevoLotes(String(nuevoLoteElegido) === String(lote) ? '' : lote);
    });
  });
}

function crearAnimalNuevo(){
  const numRaw = document.getElementById('nuevoNumero').value.trim();
  const nombre = document.getElementById('nuevoNombre').value.trim();
  const aviso = document.getElementById('nuevoAviso');
  aviso.innerHTML = '';

  if (!/^\d+$/.test(numRaw)){
    aviso.innerHTML = '<div class="aviso malo">Escribe un número entero.</div>';
    return;
  }
  const numero = parseInt(numRaw, 10);

  const existente = animalesPorNumero[String(numero)];
  if (existente){
    aviso.innerHTML = '<div class="aviso malo"><b>Ya existe activo</b>El número ' + numero + ' ya está activo (' +
      esc(existente.nombre || 'sin nombre') + ').</div>';
    return;
  }
  const enSalidas = salidasRecientes.find(s => String(s.numero) === String(numero));
  if (enSalidas){
    aviso.innerHTML = '<div class="aviso malo"><b>Ese número existe pero está fuera</b>' +
      esc(enSalidas.nombre || '') + ' salió por ' + esc(enSalidas.motivo_salida || '') + '.</div>' +
      '<button type="button" class="btn btn-verde" id="btnReactivarDesdeNuevo">Reactivar en vez de crear otro</button>';
    document.getElementById('btnReactivarDesdeNuevo').addEventListener('click', () => reactivarDesdeSalidas(numero));
    return;
  }

  const btn = document.getElementById('btnCrearAnimal');
  btn.disabled = true;
  const datos = {numero, nombre, usuario: config.usuario, dispositivo: config.dispositivo, lote: nuevoLoteElegido || ''};

  const aplicarLocal = () => {
    animalesPorNumero[String(numero)] = {numero, nombre, origen: 'manual', lote: datos.lote};
    toast('Animal creado');
    irAInicio();
  };

  if (navigator.onLine && config.urlSheet){
    pedir(config.urlSheet, Object.assign({accion: 'animal_nuevo'}, datos)).then(aplicarLocal).catch(e => {
      if (e.message === 'existeActivo'){
        aviso.innerHTML = '<div class="aviso malo"><b>Ya existe activo</b>Otra tablet ya lo creó.</div>';
      } else if (e.message === 'existeFuera'){
        aviso.innerHTML = '<div class="aviso malo"><b>Ese número existe pero está fuera</b>Reactívalo desde Salidas recientes.</div>';
      } else {
        aviso.innerHTML = '<div class="aviso malo">No se pudo crear: ' + esc(e.message || 'revisa la conexión') + '</div>';
      }
      btn.disabled = false;
    });
  } else {
    aplicarLocal();
    encolarAccion(config.urlSheet, 'animal_nuevo', datos).then(actualizarEstadoSync);
  }
}

/* =========================================================
   Salidas recientes (red de seguridad para corregir errores)
   ========================================================= */
function verSalidas(){
  pintarSalidas();
  ver('salidas');
}

function pintarSalidas(){
  const cont = document.getElementById('salidasCont');
  if (!salidasRecientes.length){
    cont.innerHTML = '<div class="vacio"><span class="ic">🗂</span>No hay salidas en los últimos 30 días.</div>';
    return;
  }
  cont.innerHTML = '<div class="lista-nombres">' + salidasRecientes.map(s =>
    '<div class="fila-animal" style="flex-direction:column;align-items:stretch;gap:8px">' +
      '<div style="display:flex;gap:12px;align-items:center">' +
        '<span class="num">' + esc(s.numero) + '</span>' +
        '<span class="nom">' + esc(s.nombre || '(sin nombre)') + '</span>' +
      '</div>' +
      '<div class="meta">' + esc(s.motivo_salida || '') + ' · ' + esc(fechaCorta(s.fecha_salida)) + ' · ' + esc(s.sacado_por || '?') + '</div>' +
      '<button type="button" class="btn btn-borde" style="margin-bottom:0" data-reactivar="' + esc(s.numero) + '">Reactivar</button>' +
    '</div>'
  ).join('') + '</div>';
  cont.querySelectorAll('[data-reactivar]').forEach(btn => {
    btn.addEventListener('click', () => reactivarDesdeSalidas(btn.getAttribute('data-reactivar')));
  });
}

function reactivarDesdeSalidas(numero){
  const s = salidasRecientes.find(x => String(x.numero) === String(numero));
  if (!s) return;
  if (!confirm('¿Reactivar a ' + numero + ' ' + (s.nombre || '') + '? Vuelve a estar activo, sin lote.')) return;

  const datos = {numero, usuario: config.usuario, dispositivo: config.dispositivo};
  const aplicarLocal = () => {
    salidasRecientes = salidasRecientes.filter(x => String(x.numero) !== String(numero));
    animalesPorNumero[String(numero)] = {numero: s.numero, nombre: s.nombre, origen: 'manual', lote: ''};
    toast('Reactivado');
    if (vista === 'salidas') pintarSalidas(); else irAInicio();
  };

  if (navigator.onLine && config.urlSheet){
    pedir(config.urlSheet, Object.assign({accion: 'animal_reactivar'}, datos)).then(aplicarLocal)
      .catch(e => toast('No se pudo reactivar: ' + (e.message || 'revisa la conexión')));
  } else {
    aplicarLocal();
    encolarAccion(config.urlSheet, 'animal_reactivar', datos).then(actualizarEstadoSync);
  }
}
