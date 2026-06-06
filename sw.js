const CACHE_NAME = 'gestorpro-v3';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './app_icon.png'
];

// Instalar el Service Worker y almacenar archivos en caché
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('[Service Worker] Precaching app shell v2...');
        return cache.addAll(ASSETS);
      })
      .then(() => self.skipWaiting())
  );
});

// Activar el Service Worker y limpiar cachés antiguas
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.map(key => {
          if (key !== CACHE_NAME) {
            console.log('[Service Worker] Removing old cache:', key);
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Interceptar peticiones y servir desde caché u obtener de red
self.addEventListener('fetch', event => {
  const url = event.request.url;

  // Excluir llamadas de API a Google Sheets de la caché para asegurar que los datos estén frescos
  if (url.includes('script.google.com') || url.includes('googleusercontent.com')) {
    return; // Bypass de caché para la API
  }

  // Excluir Google Fonts de la estrategia stale-while-revalidate (dejar que el navegador las cachee normalmente)
  if (url.includes('fonts.googleapis.com') || url.includes('fonts.gstatic.com')) {
    event.respondWith(
      caches.match(event.request).then(cached => {
        if (cached) return cached;
        return fetch(event.request).then(response => {
          if (response.status === 200) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          }
          return response;
        });
      })
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then(cachedResponse => {
      if (cachedResponse) {
        // Retornar recurso cacheado inmediatamente
        // Y paralelamente buscar en red para actualizar la caché en segundo plano (Stale-While-Revalidate)
        fetch(event.request).then(networkResponse => {
          if (networkResponse.status === 200) {
            caches.open(CACHE_NAME).then(cache => cache.put(event.request, networkResponse));
          }
        }).catch(err => console.log('[Service Worker] Fetch failed, using cached version.'));

        return cachedResponse;
      }

      // Si no está en la caché, buscar en la red
      return fetch(event.request);
    })
  );
});
