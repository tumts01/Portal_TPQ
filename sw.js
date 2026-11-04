const CACHE_NAME = "portal-almabrur-v3";

self.addEventListener("install", (event) => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  if (event.request.mode === "navigate" || event.request.url.includes("manifest.json")) {
    event.respondWith(
      fetch(event.request, { cache: "no-store" }).catch(() => fetch(event.request))
    );
    return;
  }
});
