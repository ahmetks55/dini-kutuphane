const CACHE = "dini-kutuphane-v35";
const CORE = [
  "./",
  "./index.html",
  "./styles.css?v=23",
  "./app.js?v=34",
  "./manifest.json",
  "./icon.svg",
  "./fonts/AmiriQuran-Regular.ttf",
  "./fonts/NotoNaskhArabic-Regular.ttf"
];

self.addEventListener("install", (e) => {
  e.waitUntil(
    caches.open(CACHE).then(async (c) => {
      await c.addAll(CORE);
      try {
        const res = await fetch("./api/tree?path=", { cache: "no-store" });
        if (res.ok) await c.put("./api/tree?path=", res);
      } catch (_) {}
    }).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

async function staleWhileRevalidate(req) {
  const cache = await caches.open(CACHE);
  const hit = await cache.match(req);
  const revalidate = fetch(req).then((res) => {
    if (res && res.ok) {
      const copy = res.clone();
      cache.put(req, copy);
    }
    return res;
  }).catch(() => null);
  return hit || revalidate;
}

self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  const url = new URL(e.request.url);
  if (url.origin !== location.origin) return;

  if (url.pathname.startsWith("/api/")) {
    e.respondWith(staleWhileRevalidate(e.request));
    return;
  }

  if (e.request.mode === "navigate") {
    e.respondWith(
      fetch(e.request)
        .then((res) => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put("./index.html", copy));
          }
          return res;
        })
        .catch(() => caches.match("./index.html"))
    );
    return;
  }

  e.respondWith(staleWhileRevalidate(e.request));
});
