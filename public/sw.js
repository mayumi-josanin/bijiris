const CACHE_NAME = "bijiris-respondent-v20260406-1";
const APP_SHELL_URLS = [
  "/f/",
  "/assets/style.css",
  "/assets/respondent.js",
  "/assets/manifest.webmanifest",
  "/assets/icons/apple-touch-icon.png",
  "/assets/icons/icon-192.png",
  "/assets/icons/icon-512.png",
  "/assets/icons/favicon-32.png",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches
      .open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL_URLS))
      .catch(() => undefined)
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))
      )
      .then(() => self.clients.claim())
  );
});

self.addEventListener("message", (event) => {
  if (event.data?.type === "SKIP_WAITING") {
    self.skipWaiting();
  }
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  if (request.method !== "GET") {
    return;
  }

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) {
    return;
  }

  if (url.pathname.startsWith("/api/admin/")) {
    return;
  }

  if (request.mode === "navigate") {
    event.respondWith(networkFirst(request, "/f/"));
    return;
  }

  if (url.pathname.startsWith("/assets/") || url.pathname === "/sw.js") {
    event.respondWith(staleWhileRevalidate(request));
    return;
  }

  if (url.pathname.startsWith("/api/public/") || url.pathname.startsWith("/f/")) {
    event.respondWith(networkFirst(request));
  }
});

async function networkFirst(request, fallbackKey = "") {
  const cache = await caches.open(CACHE_NAME);
  try {
    const response = await fetch(request);
    if (isCacheableResponse(response)) {
      cache.put(request, response.clone());
    }
    return response;
  } catch (error) {
    const cached = await caches.match(request);
    if (cached) {
      return cached;
    }
    if (fallbackKey) {
      const fallback = await caches.match(fallbackKey);
      if (fallback) {
        return fallback;
      }
    }
    throw error;
  }
}

async function staleWhileRevalidate(request) {
  const cache = await caches.open(CACHE_NAME);
  const cached = await caches.match(request);
  const networkPromise = fetch(request)
    .then((response) => {
      if (isCacheableResponse(response)) {
        cache.put(request, response.clone());
      }
      return response;
    })
    .catch(() => undefined);

  if (cached) {
    eventSafeRevalidate(networkPromise);
    return cached;
  }

  const network = await networkPromise;
  if (network) {
    return network;
  }
  throw new Error("offline");
}

function eventSafeRevalidate(promise) {
  promise.catch(() => undefined);
}

function isCacheableResponse(response) {
  return !!response && response.ok && response.type === "basic";
}
