// Fit Bench service worker: makes the app work offline and delivers updates.
// You never need to edit this file. To publish an update, change the number in version.js.
importScripts("version.js");

const CACHE = "fit-bench-" + self.APP_VERSION;
const FONT_CACHE = "fit-bench-fonts";
const ASSETS = [
  "./",
  "index.html",
  "version.js",
  "xlsx-lite.js",
  "manifest.webmanifest",
  "privacy.html",
  "icons/favicon-32.png",
  "icons/icon-192.png",
  "icons/icon-512.png",
  "icons/maskable-512.png"
];

self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE).then(cache => cache.addAll(ASSETS.map(u => new Request(u, { cache: "reload" }))))
  );
});

self.addEventListener("activate", event => {
  event.waitUntil((async () => {
    for (const key of await caches.keys()) {
      if (key.startsWith("fit-bench-") && key !== CACHE && key !== FONT_CACHE) await caches.delete(key);
    }
    await self.clients.claim();
  })());
});

self.addEventListener("message", event => {
  if (event.data && event.data.type === "SKIP_WAITING") self.skipWaiting();
});

self.addEventListener("fetch", event => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);

  // Google Fonts: use the saved copy, refresh it in the background.
  if (url.hostname === "fonts.googleapis.com" || url.hostname === "fonts.gstatic.com") {
    event.respondWith((async () => {
      const cache = await caches.open(FONT_CACHE);
      const hit = await cache.match(req);
      const net = fetch(req).then(res => { if (res && (res.ok || res.type === "opaque")) cache.put(req, res.clone()); return res; }).catch(() => null);
      return hit || (await net) || new Response("", { status: 504 });
    })());
    return;
  }
  if (url.origin !== self.location.origin) return;

  // Pages: always answer from the installed version so the app opens offline.
  if (req.mode === "navigate") {
    event.respondWith((async () => {
      const cache = await caches.open(CACHE);
      return (await cache.match("index.html")) || fetch(req);
    })());
    return;
  }

  // Everything else: installed copy first, network as a fallback.
  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const hit = await cache.match(req, { ignoreSearch: true });
    if (hit) return hit;
    const res = await fetch(req);
    if (res && res.ok) cache.put(req, res.clone());
    return res;
  })());
});
