/* Add To Calendar Creator — Service Worker
 * Cache-first for local assets, network-first for external resources.
 * Bump CACHE_VERSION to "v2", "v3", etc. to force a cache refresh on next visit.
 */

const CACHE_VERSION = "v1";
const CACHE_NAME    = "atc-cache-" + CACHE_VERSION;

const LOCAL_ASSETS = [
  "./",
  "./index.html",
  "./script.js",
  "./style.css",
  "./manifest.json",
  "./icon-192.svg",
  "./icon-512.svg",
];

// ── Install: pre-cache all local assets ──
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(LOCAL_ASSETS))
  );
  self.skipWaiting();
});

// ── Activate: purge old caches ──
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((k) => k.startsWith("atc-cache-") && k !== CACHE_NAME)
          .map((k) => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

// ── Fetch: route by origin ──
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Skip non-GET requests and browser extensions
  if (event.request.method !== "GET") return;
  if (!url.protocol.startsWith("http")) return;

  const isSameOrigin = url.origin === self.location.origin;
  const isExternal   = url.hostname === "www.gstatic.com" ||
                       url.hostname === "res.cdn.office.net";

  if (isSameOrigin) {
    // Cache-first for local files
    event.respondWith(cacheFirst(event.request));
  } else if (isExternal) {
    // Network-first for decorative external logos (fallback to cache when offline)
    event.respondWith(networkFirst(event.request));
  }
  // All other requests (calendar URLs, data: URIs) pass through normally
});

async function cacheFirst(request) {
  const cached = await caches.match(request);
  if (cached) return cached;
  const response = await fetch(request);
  if (response.ok) {
    const cache = await caches.open(CACHE_NAME);
    cache.put(request, response.clone());
  }
  return response;
}

async function networkFirst(request) {
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    const cached = await caches.match(request);
    return cached || new Response("", { status: 503 });
  }
}
