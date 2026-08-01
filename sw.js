/**
 * THE LEDGER — Service Worker
 * Offline-first: network-first for navigation/API, falling back to cache.
 * Static-first (cache-first) for the app shell itself so it loads instantly.
 */
const CACHE_NAME = 'ledger-cache-v1';
const APP_SHELL = ['./', './index.html', './manifest.json'];

self.addEventListener('install', function(event){
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache){ return cache.addAll(APP_SHELL); }).then(function(){ return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function(event){
  event.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.filter(function(k){ return k !== CACHE_NAME; }).map(function(k){ return caches.delete(k); }));
    }).then(function(){ return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function(event){
  const req = event.request;
  if (req.method !== 'GET') return; // never cache POSTs (API writes)

  event.respondWith(
    fetch(req).then(function(res){
      const resClone = res.clone(); // clone synchronously, before any async step
      caches.open(CACHE_NAME).then(function(cache){ cache.put(req, resClone); });
      return res;
    }).catch(function(){
      return caches.match(req).then(function(cached){ return cached || caches.match('./index.html'); });
    })
  );
});
