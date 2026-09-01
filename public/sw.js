const CACHE = "dini-kutuphane-v55";
const CORE = [
  "./",
  "./index.html",
  "./exit.html",
  "./styles.css?v=32",
  "./app.js?v=55",
  "./manifest.json",
  "./icon-192.png",
  "./apple-touch-icon.png",
  "./fonts/AmiriQuran-Regular.ttf",
  "./fonts/NotoNaskhArabic-Regular.ttf"
];

self.addEventListener("install", (e) => {
  e.waitUntil(
    caches.open(CACHE).then(async (c) => {
      await c.addAll(CORE);
      try {
        const res = await fetch("/api/tree?path=", { cache: "no-store" });
        if (res.ok) await c.put("/api/tree?path=", res);
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

let precaching = false;

async function precacheAll(client) {
  if (precaching) return;
  precaching = true;
  const cache = await caches.open(CACHE);
  const report = (done, total) => {
    if (client) client.postMessage({ type: "PRECACHE_PROGRESS", done, total });
  };
  const files = [];

  try {
    async function walk(rel) {
      const url = "/api/tree?path=" + encodeURIComponent(rel);
      let data;
      const cachedTree = await cache.match(url);
      if (cachedTree) {
        data = await cachedTree.json();
      } else {
        const res = await fetch(url);
        if (res.ok) await cache.put(url, res.clone());
        data = await res.clone().json();
      }
      for (const item of data.items || []) {
        const childRel = rel ? rel + "/" + item.name : item.name;
        if (item.type === "folder") await walk(childRel);
        else files.push(childRel);
      }
    }

    report(0, 1);
    await walk("");
    const total = files.length;
    report(0, total);
    if (total === 0) {
      if (client) client.postMessage({ type: "PRECACHE_DONE", total: 0 });
      return;
    }

    let idx = 0;
    let done = 0;
    async function worker() {
      while (idx < total) {
        const rel = files[idx++];
        const ext = rel.split(".").pop().toLowerCase();
        const isText = ext === "txt" || ext === "md";
        const url = (isText ? "/api/read?path=" : "/api/file?path=") + encodeURIComponent(rel);
        try {
          if (!(await cache.match(url))) {
            const r = await fetch(url);
            if (r.ok) await cache.put(url, r);
          }
        } catch (_) {}
        done++;
        report(done, total);
      }
    }
    await Promise.all(Array.from({ length: Math.min(6, total) }, () => worker()));
    if (client) client.postMessage({ type: "PRECACHE_DONE", total });
  } finally {
    precaching = false;
  }
}

self.addEventListener("message", (e) => {
  if (e.data && e.data.type === "PRECACHE_ALL") {
    e.waitUntil(precacheAll(e.source));
  }
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
