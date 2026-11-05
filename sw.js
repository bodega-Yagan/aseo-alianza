/* Pedidos de Aseo Alianza: permite instalar la app. La página siempre se pide primero a internet;
   la copia guardada solo se usa sin conexión. Las consultas a la planilla no pasan por aquí. */
const CACHE = 'aseo-alianza-v1';
const BASE = ['./', './index.html', './manifest.json', './icons/icon-192.png', './icons/icon-512.png'];
self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(BASE.map(u =>
    fetch(u, { cache: 'reload' }).then(r => r.ok && c.put(u, r)).catch(() => {})))));
});
self.addEventListener('activate', e => e.waitUntil(
  caches.keys().then(ks => Promise.all(ks.filter(k => k.indexOf('aseo-alianza') === 0 && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim())));
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || u.origin !== location.origin) return;
  const nav = r.mode === 'navigate', key = nav ? './index.html' : r;
  e.respondWith(fetch(r.url, { cache: 'no-store', credentials: 'same-origin' }).then(res => {
    if (res.ok) { const cp = res.clone(); caches.open(CACHE).then(c => c.put(key, cp)); }
    return res;
  }).catch(() => caches.match(key).then(m => m || caches.match('./'))));
});
