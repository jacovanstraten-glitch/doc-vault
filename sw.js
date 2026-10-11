// App shell cache: network first, cached copy when offline. Only same-origin GET requests.
const CACHE = "dv-shell-v8";
const SHELL = ["./", "index.html", "styles.css", "app.js", "config.js", "manifest.webmanifest", "icons/icon.svg", "icons/icon-192.png",
  "vendor/msal-browser.min.js", "vendor/pdf.min.js", "vendor/pdf.worker.min.js", "vendor/jspdf.umd.min.js", "vendor/pdf-lib.min.js", "vendor/tesseract.min.js", "vendor/tesseract-worker.min.js"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(r => {
    if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
    return r;
  }).catch(() => caches.match(e.request, {ignoreSearch: true}).then(r => r || caches.match("index.html"))));
});
