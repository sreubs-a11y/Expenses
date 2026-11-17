/*
 * Keeps the app itself (this page, icons, manifest) on the device so it opens instantly and offline.
 * Your data is never stored here - requests to your Apps Script go straight to the network.
 * When you upload a new version to GitHub, bump VERSION: devices fetch the new files fresh
 * (skipping GitHub's 10-minute cache), and an open app reloads itself onto the new version.
 */
const VERSION = 'v4';
const CACHE = 'bills-app-' + VERSION;
const SHELL = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png', './vendor/pdf.min.mjs', './vendor/pdf.worker.min.mjs'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE)
    .then(c => Promise.all(SHELL.map(u => fetch(new Request(u, { cache: 'reload' })).then(r => { if (r.ok) return c.put(u === './' ? './index.html' : u, r); }))))
    .then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k.startsWith('bills-app-') && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;   // API calls etc. go to the network untouched
  if (url.pathname.endsWith('/sw.js')) return;
  const key = req.mode === 'navigate' ? './index.html' : req;
  e.respondWith(caches.open(CACHE).then(cache =>
    cache.match(key).then(hit => {
      const fresh = fetch(req, { cache: 'no-cache' }).then(res => {
        if (res && res.ok && res.type === 'basic') cache.put(key, res.clone());
        return res;
      }).catch(() => hit);
      return hit || fresh;          // cached copy now, refreshed copy next time
    })
  ));
});
