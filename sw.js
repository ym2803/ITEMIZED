/* Service worker ITEMIZED: agar halaman & pustaka kamera tetap terbuka saat sinyal hilang.
   Taruh di folder yang sama dengan index.html. Data scan tidak lewat sini (sudah diantrekan di perangkat). */
const V = 'itemized-v8';
const LIBS = [
  'https://cdnjs.cloudflare.com/ajax/libs/html5-qrcode/2.3.8/html5-qrcode.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];
const CDN = /^https:\/\/(cdnjs\.cloudflare\.com|fonts\.googleapis\.com|fonts\.gstatic\.com)\//;

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(V).then(c => Promise.all(
    ['./'].concat(LIBS).map(u => fetch(u).then(r => r.ok && c.put(u, r)).catch(() => {}))
  )));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET' || url.hostname.endsWith('script.google.com') || url.hostname.endsWith('googleusercontent.com')) return;
  if (url.origin === location.origin){                      // halaman: utamakan versi terbaru, cadangan dari cache saat offline
    e.respondWith(fetch(req).then(r => { const cp = r.clone(); caches.open(V).then(c => c.put(req, cp)); return r; })
      .catch(() => caches.match(req, {ignoreSearch: true}).then(m => m || caches.match('./'))));
  }else if (CDN.test(req.url)){                             // pustaka & font: cache dulu
    e.respondWith(caches.match(req).then(m => m || fetch(req).then(r => { const cp = r.clone(); caches.open(V).then(c => c.put(req, cp)); return r; })));
  }
});
