/* Service worker de Â«Novedades ContymeÂ».
 * - Guarda la app (pÃ¡gina, manifiesto, Ã­conos) al instalarse: abre sin internet.
 * - PÃ¡gina: primero la red (para tomar la versiÃ³n nueva), si no hay red la guardada.
 * - Las fuentes vienen dentro de la app (carpeta fuentes/).
 * Al cambiar algo de la app, subir VERSION para que los telÃ©fonos tomen la nueva. */
const VERSION = 'novedades-v12';
const APP = [
  './',
  'index.html',
  'manifest.webmanifest',
  'iconos/logo-contyme-blanco-520.png',
  'iconos/icono-192.png',
  'iconos/icono-512.png',
  'iconos/icono-maskable-512.png',
  'iconos/apple-touch-icon.png',
  'iconos/favicon-32.png',
  'fuentes/fuentes.css',
  'fuentes/outfit.woff2',
  'fuentes/source-sans-3.woff2',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(APP)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((nombres) => Promise.all(nombres.filter((n) => n !== VERSION).map((n) => caches.delete(n))))
      .then(() => self.clients.claim()),
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req)
        .then((r) => { const copia = r.clone(); caches.open(VERSION).then((c) => c.put('index.html', copia)); return r; })
        .catch(() => caches.match('index.html')),
    );
    return;
  }

  if (url.origin === self.location.origin) {
    e.respondWith(caches.match(req).then((guardada) => guardada || fetch(req)));
  }
});
