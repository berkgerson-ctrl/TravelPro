// Rota — service worker
// Caches the app shell (this single HTML file + icons) so the PWA launches
// instantly from the home screen and the UI still loads with no signal.
// Data itself always goes straight to Google Sheets (see index.html) —
// this worker never caches API responses or third-party CDN scripts,
// only the app's own static assets.

const CACHE_NAME = 'rota-shell-v2';
const SHELL_FILES = [
  './index.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) =>
      // Cache each file independently so one missing/renamed asset can
      // never block the whole service worker from installing.
      Promise.allSettled(SHELL_FILES.map((url) => cache.add(url)))
    )
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Only ever manage our own same-origin static files. Everything else
  // (Google Apps Script, Chart.js/Leaflet/SortableJS CDNs, exchange-rate
  // and geocoding APIs, Google Fonts) is left completely untouched so it
  // always goes straight to the network with normal browser behavior.
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) {
    return;
  }

  event.respondWith(
    caches.match(req).then((cached) => {
      const network = fetch(req)
        .then((response) => {
          if (response && response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
          }
          return response;
        })
        .catch(() => cached);
      return cached || network;
    })
  );
});
