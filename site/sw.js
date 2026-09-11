// TPAHA Books service worker: caches the app shell so the pages open offline.
// Data never passes through here: Microsoft Graph and sign-in requests are cross-origin and go straight to the network.
const VERSION = new URL(self.location.href).searchParams.get('v') || 'dev';
const CACHE = `tpaha-books-${VERSION}`;
const ASSETS = [
  './', './index.html', './ledger.html', './storage.html', './diagnostics.html', './css/app.css', './manifest.webmanifest',
  './js/config.js', './js/auth.js', './js/store.js', './js/ui.js', './js/xlsx-export.js', './js/pwa.js', './js/diagnostics.js',
  './js/ledger/model.js', './js/ledger/app.js', './js/storage/model.js', './js/storage/app.js',
  './js/workbook/excel-values.js', './js/workbook/excel-client.js', './js/workbook/mock-excel.js', './js/workbook/sample-workbook.js',
  './js/workbook/ledger-workbook.js', './js/workbook/locate.js', './js/save/save-controller.js', './js/save/drafts.js',
  './vendor/msal-browser.min.js', './vendor/xlsx.full.min.js',
  './icons/icon-192.png', './icons/icon-512.png', './icons/icon-maskable-512.png', './icons/apple-touch-icon.png',
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k.startsWith('tpaha-books-') && k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim()));
});

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;
  const isPage = req.mode === 'navigate' || url.pathname.endsWith('.html') || url.pathname.endsWith('/');
  if (isPage) {
    // Network first so a new release shows up; cached copy when offline.
    event.respondWith(fetch(req).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return res; })
      .catch(() => caches.match(req, { ignoreSearch: true }).then(hit => hit || caches.match('./index.html'))));
    return;
  }
  // Everything else: cache first, refresh in the background.
  event.respondWith(caches.match(req).then(hit => {
    const refresh = fetch(req).then(res => { if (res.ok) caches.open(CACHE).then(c => c.put(req, res.clone())); return res; }).catch(() => hit);
    return hit || refresh;
  }));
});
