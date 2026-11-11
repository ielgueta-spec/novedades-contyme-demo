/* Service worker de «Novedades Remuneraciones» (Contyme).
 * - Guarda la app (página, manifiesto, íconos, fuentes) al instalarse: abre sin internet.
 * - Página: primero la red, sin caché del navegador (para tomar siempre la versión nueva); sin red, la guardada.
 * - Cada mejora: subir VERSION aquí y VERSION_APP en index.html (mismo número). Los teléfonos instalados la detectan
 *   al abrir la app o al volver a ella, y se recargan solos cuando la persona no está a mitad de algo. */
const VERSION = 'novedades-v14';
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
  e.waitUntil(
    caches.open(VERSION)
      .then((c) => c.addAll(APP.map((u) => new Request(u, { cache: 'reload' }))))
      .then(() => self.skipWaiting()),
  );
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
      fetch(req.url, { cache: 'no-cache', credentials: 'same-origin' })
        .then((r) => {
          // Solo la app (./ o index.html) se guarda para abrir sin internet; otras páginas (ayuda.html) no la pisan.
          const esApp = /\/(index\.html)?$/.test(url.pathname);
          if (r.ok && esApp) { const copia = r.clone(); caches.open(VERSION).then((c) => c.put('index.html', copia)); }
          return r;
        })
        .catch(() => caches.match(req).then((g) => g || caches.match('index.html'))),
    );
    return;
  }

  if (url.origin === self.location.origin) {
    e.respondWith(caches.match(req).then((guardada) => guardada || fetch(req)));
  }
});
