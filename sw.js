// Keeps the app shell available offline and picks up new versions automatically.
// Bump VERSION whenever any file changes so phones fetch the new files.
const VERSION = "budget-1.0.0";
const SHELL = [
  "./", "index.html", "styles.css", "manifest.webmanifest",
  "js/app.js", "js/auth.js", "js/data.js", "js/analytics.js", "js/financials.js", "js/prices.js", "js/util.js",
  "icons/icon-192.png", "icons/icon-512.png", "icons/apple-touch-icon.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

// Only this site's own files are cached. Microsoft sign-in, OneDrive and price requests always go to the network.
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  if (url.search.includes("code=") || url.search.includes("error=")) return; // sign-in redirect
  e.respondWith(
    fetch(e.request).then((r) => {
      if (r.ok) { const copy = r.clone(); caches.open(VERSION).then((c) => c.put(e.request, copy)); }
      return r;
    }).catch(() => caches.match(e.request, { ignoreSearch: true }).then((r) => r || caches.match("index.html"))),
  );
});
