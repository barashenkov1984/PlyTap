const CACHE_NAME = 'plytap-app-v1';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './plytap-db.js',
  './plytap-charts.js',
  './plytap-pdf-report.js',
  './plytap-archive.js',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-512-maskable.png',
  'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',
  'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js',
  'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js'
];

// Установка: кэшируем файлы по одному, чтобы ошибка одного не ломала всё приложение
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return Promise.all(
        APP_SHELL.map((url) => {
          return cache.add(url).catch((err) => {
            console.warn('Не удалось загрузить в кэш:', url, err);
          });
        })
      );
    })
  );
});

// Активация: удаляем старые версии кэша (v1 и ниже) и сразу активируем новый кэш
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      );
    }).then(() => self.clients.claim())
  );
});

// Перехват запросов: СНАЧАЛА КЭШ, ПОТОМ СЕТЬ
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  event.respondWith(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.match(event.request).then((cachedResponse) => {
        // 1. Если файл есть в актуальном кэше — отдаем его (работает офлайн)
        if (cachedResponse) {
          return cachedResponse;
        }

        // 2. Если файла нет в кэше — идем в сеть и кэшируем на лету
        return fetch(event.request).then((networkResponse) => {
          if (!networkResponse || networkResponse.status !== 200) {
            return networkResponse;
          }
          cache.put(event.request, networkResponse.clone());
          return networkResponse;
        }).catch(() => {
          // 3. Если нет сети и нет в кэше для навигации — отдаем index.html
          if (event.request.mode === 'navigate') {
            return cache.match('./index.html');
          }
        });
      });
    })
  );
});