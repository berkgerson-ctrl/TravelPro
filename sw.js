// Rota — service worker
// Caches the app shell (this single HTML file + icons) so the PWA launches
// instantly from the home screen and the UI still loads with no signal.
// Data itself always goes straight to Google Sheets (see index.html) —
// this worker never caches API responses, only static assets.

const CACHE_NAME = 'rota-shell-v1';
const SHELL_FILES = [
  './',
  './index.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).catch(() => {})
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
  const url = event.request.url;

  // Never intercept calls to the Google Apps Script backend — those must
  // always hit the network live so data stays real-time and consistent.
  if (url.includes('script.google.com') || url.includes('script.googleusercontent.com')) {
    return;
  }

  // App shell: cache-first (instant load), falling back to network then
  // refreshing the cache for next time.
  event.respondWith(
    caches.match(event.request).then((cached) => {
      const network = fetch(event.request)
        .then((response) => {
          if (response && response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          }
          return response;
        })
        .catch(() => cached);
      return cached || network;
    })
  );
});
