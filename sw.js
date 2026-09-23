/* Service worker: guarda la aplicación en la tablet para que abra sin internet. */
const CACHE = 'chequeo-v11';
const ARCHIVOS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png',
  './chequeo/index.html',
  './chequeo/xlsx.full.min.js',
  './comun/estilos.css',
  './comun/almacen.js',
  './comun/config.js',
  './comun/sync.js',
  './comun/usuario.js',
  './lotes/index.html',
  './lotes/lotes.js'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ARCHIVOS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                       // los envíos a Google no se cachean
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;        // nada externo

  e.respondWith(
    caches.match(req).then(hit => {
      if (hit) {
        // refresca en segundo plano, pero responde ya con lo guardado
        fetch(req).then(r => {
          if (r && r.ok) caches.open(CACHE).then(c => c.put(req, r.clone()));
        }).catch(() => {});
        return hit;
      }
      return fetch(req)
        .then(r => {
          if (r && r.ok) {
            const copia = r.clone();
            caches.open(CACHE).then(c => c.put(req, copia));
          }
          return r;
        })
        .catch(() => caches.match('./index.html'));
    })
  );
});
