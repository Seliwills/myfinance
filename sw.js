const CACHE_NAME = 'ledger-shell-v5';
const SHELL_FILES = [
  './',
  './index.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).catch(()=>{})
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

// Network-first for the app shell and page navigations: always try to get
// the freshest version first, and only fall back to the cached copy if the
// network genuinely fails (offline). This avoids ever serving a stale
// index.html after a deploy.
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  const isShellFile = SHELL_FILES.some((f) => url.pathname.endsWith(f.replace('./', '/')) || url.pathname === '/' );
  const isSameOrigin = url.origin === self.location.origin;

  if (isSameOrigin && (isShellFile || event.request.mode === 'navigate')) {
    event.respondWith(
      fetch(event.request).then((res) => {
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, res.clone()));
        return res;
      }).catch(() => caches.match(event.request))
    );
  }
  // else: let the browser handle it normally (network) — including calls
  // to the Apps Script Sheet, which should never be served from cache.
});
