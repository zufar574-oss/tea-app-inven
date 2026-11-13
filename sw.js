/* ============================================================
   Service Worker — Чай: приёмка и инвентаризация (PWA-lite)
   Назначение: офлайн после первой загрузки + «Добавить на главный экран».
   Стратегия:
     • install  — precache app shell (статика);
     • activate — вычистить старые версии кэша;
     • fetch   — навигация: сеть→при сбое закэшенный index.html;
                 прочие GET того же origin: cache-first, фоном обновляем кэш.
   Примечание: загружаемая пользователем база .xlsx через <input type=file>
   идёт мимо fetch — в кэш не попадает, что и нужно (она в localStorage).
   Все пути относительные ('./…'), чтобы работать и на GitHub Pages subpath.
   ============================================================ */

const CACHE = 'tea-app-v4';
const SHELL = [
  './',
  './index.html',
  './manifest.json',
  './css/style.css',
  './js/storage.js',
  './js/data.js',
  './js/ui.js',
  './js/main.js',
  './vendor/xlsx.full.min.js',
  './assets/logo.png',
  './assets/icon-192.png',
  './assets/icon-512.png',
  './assets/apple-touch-icon.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE)
      .then((cache) => cache.addAll(SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Кэшируем только same-origin GET.
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  // Навигация: сеть, при сбое — закэшенный index.html (офлайн-режим).
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req).catch(() =>
        caches.match(req).then((r) => r || caches.match('./index.html'))
      )
    );
    return;
  }

  // Прочие GET: cache-first, фоновое обновление кэша.
  event.respondWith(
    caches.match(req).then((cached) => {
      const network = fetch(req).then((res) => {
        if (res && res.status === 200 && res.type === 'basic') {
          const copy = res.clone();
          caches.open(CACHE).then((cache) => cache.put(req, copy));
        }
        return res;
      }).catch(() => cached);
      return cached || network;
    })
  );
});
