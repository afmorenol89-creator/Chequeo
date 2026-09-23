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

const COL_NUM = 'numero de animal'; // ya pasado por norm()
const COL_NOM = 'nombre';

/* =========================================================
   Navegación entre pantallas
   ========================================================= */
function ver(v){
  vista = v;
  ['quien', 'ajustes', 'inicio', 'cargar', 'revision', 'resultado'].forEach(n => {
    const el = document.getElementById('v-' + n);
    if (el) el.classList.toggle('oculto', n !== v);
  });
  window.scrollTo(0, 0);
}

function actualizarNet(){
  const el = document.getElementById('net'), t = document.getElementById('netTxt');
  if (navigator.onLine){ el.classList.remove('off'); t.textContent = 'Con internet'; }
  else { el.classList.add('off'); t.textContent = 'Sin conexión'; }
}
window.addEventListener('online', actualizarNet);
window.addEventListener('offline', actualizarNet);

/* =========================================================
   Arranque: base local, hoja configurada, usuario del dispositivo
   ========================================================= */
function iniciar(){
  actualizarNet();
  abrirDB().then(() => cargarConfig()).then(c => {
    config = c;
    seguirInicio();
  });
}

function seguirInicio(){
  if (!config.urlSheet){ ver('ajustes'); return; }
  if (!config.usuario){ mostrarSelectorUsuarioInicial(); return; }
  document.getElementById('nombreUsuario').textContent = config.usuario;
  ver('inicio');
}

function guardarUrl(){
  const v = document.getElementById('inputUrl').value.trim();
  if (!v) return;
  config.urlSheet = v;
  guardarConfigValor('urlSheet', v).then(() => {
    toast('Hoja guardada');
    seguirInicio();
  });
}

function mostrarSelectorUsuarioInicial(){
  ver('quien');
  const cont = document.getElementById('contQuien');
  if (!navigator.onLine){
    mostrarSelectorUsuario(cont, [], elegirUsuario);
    return;
  }
  cont.innerHTML = '<div class="aviso">Cargando trabajadores…</div>';
  pedirGet(config.urlSheet, 'lotes_estado').then(j => {
    mostrarSelectorUsuario(cont, j.trabajadores || [], elegirUsuario);
  }).catch(() => {
    mostrarSelectorUsuario(cont, [], elegirUsuario);
  });
}

function elegirUsuario(nombre){
  config.usuario = nombre;
  guardarConfigValor('usuario', nombre).then(() => {
    document.getElementById('nombreUsuario').textContent = nombre;
    ver('inicio');
  });
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
  }).catch(e => {
    toast('No se pudo guardar: ' + (e.message || 'revisa la conexión'));
  }).then(() => {
    btn.disabled = false; btn.textContent = 'Guardar cambios';
  });
}
