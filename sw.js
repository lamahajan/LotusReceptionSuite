/* ═══════════════════════════════════════════════════════════════
   Reception Suite — service worker

   Do kaam karta hai:
     1) App ka shell (index.html, icons, manifest) cache me rakhta hai,
        taaki net na ho to bhi app khul jaye.
     2) Firebase / API requests ko KABHI cache nahi karta — data
        hamesha live aana chahiye, purana nahi.

   Update kaise hota hai:
     index.html network se pehle try hota hai. Naya version mila to
     wahi dikhta hai aur cache bhi update ho jata hai. Net na ho to
     cache wala purana khul jata hai.

   Version badalne par purana cache apne aap delete ho jata hai —
   file badlo to CACHE ka number bhi badha dena.
   ═══════════════════════════════════════════════════════════════ */

const CACHE = 'reception-suite-v1';

const SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './favicon-64.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE)
      // addAll fail ho jata hai agar ek bhi file na mile — isliye ek-ek karke
      .then(cache => Promise.all(SHELL.map(url => cache.add(url).catch(() => {}))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;

  // Sirf GET cache hota hai
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // Apne origin ke bahar ka kuch bhi (Firebase, Google Fonts, CDN scripts) —
  // seedha network, koi cache nahi.
  if (url.origin !== self.location.origin) return;

  // Kisi bhi live-data JSON ko kabhi cache na karo
  if (url.pathname.endsWith('.json') && url.pathname !== '/manifest.json' && !url.pathname.endsWith('manifest.json')) return;

  // HTML: network pehle, warna cache (taaki update turant mile)
  if (req.mode === 'navigate' || (req.headers.get('accept') || '').includes('text/html')) {
    event.respondWith(
      fetch(req)
        .then(res => {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match(req).then(hit => hit || caches.match('./index.html')))
    );
    return;
  }

  // Baaki assets: cache pehle, background me refresh
  event.respondWith(
    caches.match(req).then(hit => {
      const network = fetch(req)
        .then(res => {
          if (res && res.status === 200) {
            const copy = res.clone();
            caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
          }
          return res;
        })
        .catch(() => hit);
      return hit || network;
    })
  );
});

// App se "turant update karo" ka message
self.addEventListener('message', event => {
  if (event.data === 'skipWaiting') self.skipWaiting();
});
