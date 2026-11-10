/* Byliny s rozumem – offline režim. Soubory aplikace a písma se ukládají do tabletu. */
const CACHE = 'bsr-tablet-v3';
const FILES = ['./', 'index.html', 'tablet.js', 'aplikace.html', 'xlsx.mini.min.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png'];
const FONTS = 'https://fonts.googleapis.com/css2?family=Alegreya:ital,wght@0,500;0,700;1,400;1,500&family=Alegreya+Sans:wght@400;500;700&family=IBM+Plex+Mono:wght@400;500&display=swap';

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES).then(() => c.add(FONTS).catch(() => {}))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const own = url.origin === location.origin;
  const font = /^fonts\.(googleapis|gstatic)\.com$/.test(url.hostname);
  if (!own && !font) return; // data z tabulky řeší aplikace sama
  const key = req.mode === 'navigate' ? new Request(new URL('./', location).href) : req;
  e.respondWith(caches.open(CACHE).then(async c => {
    const hit = await c.match(key, { ignoreSearch: own });
    const net = fetch(req).then(r => { if (r && (r.ok || r.type === 'opaque')) c.put(key, r.clone()); return r; });
    if (hit) { e.waitUntil(net.catch(() => {})); return hit; }
    return net;
  }));
});
