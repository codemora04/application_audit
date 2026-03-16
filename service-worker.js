// service-worker.js
// ⚠️ IMPORTANT: bump this version number every time you update the app
//    so that phones discard the old cache and load the new version.
const CACHE_NAME = "audit-v2";

const APP_SHELL = [
  "./",
  "./login.html",
  "./accueil.html",
  "./baltimar.html",
  "./revey.html",
  "./visualisation.html",
  "./style.css",
  "./pwa.js",
  "./supabase.js",
  "./baltimar.js",
  "./revey.js",
  "./visualisation.js",
  "./manifest.json",
  "./icon.png",
  "./logo1.png",
  "./logo2.png",
];

// Pre-cache the app shell on install
self.addEventListener("install", (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
});

// Delete ALL old caches on activate → phone gets the new version immediately
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

// Network-first: always try the network (important for Supabase live data).
// Falls back to cache only when offline.
self.addEventListener("fetch", (event) => {
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // Update cache with the fresh response
        const clone = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
