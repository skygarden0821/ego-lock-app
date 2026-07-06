/* Xector1 EGO LOCK — service worker (offline-first) */
const CACHE = 'egolock-v8';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-180.png',
  './icon-maskable-512.png',
  './favicon-64.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => c.addAll(ASSETS))
      .then(() => self.skipWaiting())
      .catch(() => {})
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  e.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;
      return fetch(req)
        .then((resp) => {
          // runtime-cache same-origin GETs so the app keeps working offline
          try {
            if (resp && resp.status === 200 && new URL(req.url).origin === self.location.origin) {
              const copy = resp.clone();
              caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
            }
          } catch (_) {}
          return resp;
        })
        .catch(() => caches.match('./index.html'));
    })
  );
});
