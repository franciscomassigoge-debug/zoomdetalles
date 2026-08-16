// Service worker de la app "Zoom Agricultura — Detalles Administración".
// Cachea el shell de la app para que sea instalable y funcione offline
// una vez cargada por primera vez. Los datos (Excel subido, Firestore)
// no se cachean acá.

const CACHE_NAME = "zoom-detalles-v2";
const APP_SHELL = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./store.js",
  "./firebase-config.js",
  "./manifest.json",
  "./assets/logo-pdf.jpg",
  "./assets/logo-app.jpg",
  "./assets/icon-192.png",
  "./assets/icon-512.png",
  "./assets/registro-zoom-base.xlsx"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // No interceptar llamadas a Firebase/Firestore: siempre red.
  if (url.hostname.includes("googleapis") || url.hostname.includes("firebaseio") || url.hostname.includes("gstatic")) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((response) => {
        // Cachear también librerías de CDN (SheetJS, jsPDF) la primera vez.
        if (event.request.method === "GET" && response.ok) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => cached);
    })
  );
});
