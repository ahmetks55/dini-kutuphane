/* Dini Kutuphane - Offline veri katmani (v70)
 *
 * Capacitor (APK) icinde calisirken /api/* isteklerini yerel olarak karsilar:
 *   - Temel kitaplik: public/offline-data/ (APK'ya gomulu) + index.json manifest
 *   - Kullanici degisiklikleri (yukleme/yeni metin/duzenleme/taşıma/silme):
 *     IndexedDB overlay ("dk-offline")
 *   - OCR / PDF / DOCX: gomulu vendor paketleri (tesseract.js, pdf-lib, docx)
 * Web (tarayici/PWA) tarafinda dapi() normal fetch'e duser, davranis degismez.
 */
(function () {
  "use strict";

  const IS_NATIVE = !!(window.Capacitor && window.Capacitor.isNativePlatform && window.Capacitor.isNativePlatform());
  window.__OFFLINE__ = IS_NATIVE;
  if (!IS_NATIVE) {
    window.dapi = function dapi(path, opts) { return fetch(path, opts); };
    return;
  }

  /* ---------- yardimcilar ---------- */
  const DB_NAME = "dk-offline";
  const DB_VER = 1;
  const trColl = new Intl.Collator("tr", { sensitivity: "base" });

  function dbOpen() {
    return new Promise((resolve, reject) => {
      const rq = indexedDB.open(DB_NAME, DB_VER);
      rq.onupgradeneeded = () => {
        const db = rq.result;
        if (!db.objectStoreNames.contains("files")) db.createObjectStore("files");
        if (!db.objectStoreNames.contains("meta")) db.createObjectStore("meta");
      };
      rq.onsuccess = () => resolve(rq.result);
      rq.onerror = () => reject(rq.error);
    });
  }
  let _dbPromise = null;
  function db() { if (!_dbPromise) _dbPromise = dbOpen(); return _dbPromise; }

  function idbGet(store, key) {
    return db().then((d) => new Promise((res, rej) => {
      const tx = d.transaction(store, "readonly");
      const rq = tx.objectStore(store).get(key);
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => rej(rq.error);
    }));
  }
  function idbPut(store, key, val) {
    return db().then((d) => new Promise((res, rej) => {
      const tx = d.transaction(store, "readwrite");
      tx.objectStore(store).put(val, key);
      tx.oncomplete = () => res();
      tx.onerror = () => rej(tx.error);
    }));
  }
  function idbDelete(store, key) {
    return db().then((d) => new Promise((res, rej) => {
      const tx = d.transaction(store, "readwrite");
      tx.objectStore(store).delete(key);
      tx.oncomplete = () => res();
      tx.onerror = () => rej(tx.error);
    }));
  }

  /* overlay dosya islemleri */
  function overlayGet(rel) { return idbGet("files", rel); }
  function overlayPut(rel, blob, ext, size) {
    bustCaches();
    return idbPut("files", rel, { blob, ext: ext || '', size: size == null ? blob.size : size, mtime: Date.now() });
  }
  function overlayDelete(rel) { bustCaches(); return idbDelete("files", rel); }

  /* meta islemleri: 'del' tombstone, 'dir' olusturulan klasor, 'via:<rel>' alias */
  function metaGet(rel) { return idbGet("meta", rel); }
  function metaSet(rel, val) { bustCaches(); return idbPut("meta", rel, val); }
  function metaDelete(rel) { bustCaches(); return idbDelete("meta", rel); }

  /* ---------- performans onbellekleri ----------
   * Tema: her agac acilisi icin IndexedDB'yi tek tek sormak ve tum overlay
   * blob'larini okumak yerine, yalnizca ANAHTARLARI (key-only) ve meta
   * haritasini bir kez yukleyip agac sonuclarini tembel + onbellekli kurariz.
   * overlayPut/Delete ve metaSet/Delete her seferinde bustCaches() cagirdigi
   * icin kullanicinin yaptigi degisiklikler sonrasi onbellek kendini bozar.
   */
  let _fileKeySet = null; // overlay dosya anahtarlari (blob'suz)
  let _metaMap = null;    // meta haritasi (kucuk)
  let treeCache = new Map();

  function bustCaches() {
    treeCache.clear();
    _fileKeySet = null;
    _metaMap = null;
  }

  function fileKeySet() {
    if (_fileKeySet) return Promise.resolve(_fileKeySet);
    return db().then((d) => new Promise((res, rej) => {
      const tx = d.transaction("files", "readonly");
      const out = new Set();
      const cur = tx.objectStore("files").openKeyCursor();
      cur.onsuccess = () => { const c = cur.result; if (c) { out.add(c.key); c.continue(); } else { _fileKeySet = out; res(out); } };
      cur.onerror = () => rej(cur.error);
    }));
  }

  function metaMap() {
    if (_metaMap) return Promise.resolve(_metaMap);
    return overlayAllMeta().then((all) => {
      const m = new Map();
      for (const [k, v] of all) m.set(k, v);
      _metaMap = m;
      return m;
    });
  }

  /* rel icin bosluk duyarlisi yardimcilar (kac akim olmadan) */
  function viaSplit(rel, mm) {
    const parts = rel.split("/");
    for (let i = parts.length; i >= 0; i--) {
      const p = parts.slice(0, i).join("/");
      const v = mm.get(p);
      if (typeof v === "string" && v.startsWith("via:")) return [p, v.slice(4)];
    }
    return null;
  }
  function relDel(rel, mm) {
    const parts = rel.split("/");
    for (let i = parts.length; i >= 0; i--) {
      if (mm.get(parts.slice(0, i).join("/")) === "del") return true;
    }
    return false;
  }
  function nodeTypeM(rel, mm, ks) {
    const via = viaSplit(rel, mm);
    if (via) return nodeTypeM(via[1], mm, ks);
    if (ks.has(rel)) return "file";
    if (baseFiles.has(rel)) return "file";
    if (baseFolders.has(rel)) return "folder";
    if (mm.get(rel) === "dir") return "folder";
    return null;
  }

  function overlayAllFiles() {
    return db().then((d) => new Promise((res, rej) => {
      const tx = d.transaction("files", "readonly");
      const out = [];
      const cur = tx.objectStore("files").openCursor();
      cur.onsuccess = () => { const c = cur.result; if (c) { out.push([c.key, c.value]); c.continue(); } else res(out); };
      cur.onerror = () => rej(cur.error);
    }));
  }
  function overlayAllMeta() {
    return db().then((d) => new Promise((res, rej) => {
      const tx = d.transaction("meta", "readonly");
      const out = [];
      const cur = tx.objectStore("meta").openCursor();
      cur.onsuccess = () => { const c = cur.result; if (c) { out.push([c.key, c.value]); c.continue(); } else res(out); };
      cur.onerror = () => rej(cur.error);
    }));
  }

  /* ---------- manifest index ---------- */
  let _manifestPromise = null;
  const baseFiles = new Map();   // rel -> {ext,size,mtime}
  const baseFolders = new Set(); // klasor rels

  function manifest() {
    if (!_manifestPromise) {
      _manifestPromise = fetch("offline-data/index.json")
        .then((r) => { if (!r.ok) throw new Error("manifest yok: " + r.status); return r.json(); })
        .then((m) => {
          baseFiles.clear();
          baseFolders.clear();
          for (const rel of Object.keys(m.files || {})) {
            baseFiles.set(rel, m.files[rel]);
            const parts = rel.split("/");
            for (let i = 1; i < parts.length; i++) baseFolders.add(parts.slice(0, i).join("/"));
          }
          return m;
        })
        .catch(() => null);
    }
    return _manifestPromise;
  }

  /* rel -> kaynak cozumleme (alias zinciri + overlay oncelik) */
  async function resolveRel(rel) {
    const ks = await fileKeySet();
    if (ks.has(rel)) return { kind: "overlay", rel };
    const via = viaSplit(rel, await metaMap());
    if (via) {
      const mapped = via[1] + rel.slice(via[0].length);
      return resolveRel(mapped);
    }
    if (ks.has(rel)) return { kind: "overlay", rel };
    return { kind: "bundle", rel };
  }

  async function isDeleted(rel) {
    return relDel(rel, await metaMap());
  }

  async function nodeType(rel) {
    const [mm, ks] = await Promise.all([metaMap(), fileKeySet()]);
    return nodeTypeM(rel, mm, ks);
  }

  function extOf(name) {
    const i = name.lastIndexOf(".");
    return i > 0 ? name.slice(i).toLowerCase() : "";
  }

  /* ---------- tree ---------- */
  function baseChildrenOf(folderRel) {
    const out = new Map(); // name -> item
    const prefix = folderRel ? folderRel + "/" : "";
    if (baseFiles.has(folderRel)) { /* file, no children */ }
    for (const rel of baseFiles.keys()) {
      if (!rel.startsWith(prefix)) continue;
      const rest = rel.slice(prefix.length);
      const first = rest.indexOf("/");
      if (first === -1) {
        const name = rest;
        if (!out.has(name)) out.set(name, { name, type: "file", rel, ext: baseFiles.get(rel).ext, size: baseFiles.get(rel).size });
      } else {
        const name = rest.slice(0, first);
        if (!out.has(name)) out.set(name, { name, type: "folder", rel: prefix + name, folder: true });
      }
    }
    return out;
  }

  async function treeOf(folderRel) {
    await manifest();
    const cacheKey = folderRel || "";
    if (treeCache.has(cacheKey)) return treeCache.get(cacheKey);
    const mm = await metaMap();
    const ks = await fileKeySet();
    const items = [];
    const seen = new Set();
    const inFolder = (rel) => {
      if (!folderRel) return rel.indexOf("/") === -1;
      return rel.startsWith(folderRel + "/") && rel.slice(folderRel.length + 1).indexOf("/") === -1;
    };

    /* alias klasor acilinca asil icerige yonlendir (dosya alias'lari gibi) */
    const viaSelf = viaSplit(folderRel || "", mm);
    if (viaSelf && folderRel) {
      const mapped = viaSelf[1] + folderRel.slice(viaSelf[0].length);
      if (mapped !== folderRel) return treeOf(mapped);
    }

    const base = baseChildrenOf(folderRel);
    for (const name of base.keys()) {
      const it = base.get(name);
      const rel = it.rel;
      if (relDel(rel, mm)) continue;
      if (!it.folder) {
        let size = baseFiles.get(rel) ? baseFiles.get(rel).size : null;
        let ext = it.ext || extOf(it.name);
        if (ks.has(rel)) {
          const rec = await overlayGet(rel);
          if (rec) { size = rec.size; ext = rec.ext || ext; }
        }
        items.push({ name: it.name, type: "file", rel, ext, size: size || null });
        seen.add(rel);
      } else if (!viaSplit(rel, mm)) {
        items.push({ name: it.name, type: "folder", rel, size: null });
        seen.add(rel);
      }
    }
    /* overlay'deki dosyalar (yalnizca bu klasoru tarar, blob okumaz) */
    for (const rel of ks) {
      if (!inFolder(rel) || seen.has(rel)) continue;
      if (relDel(rel, mm)) continue;
      const rec = await overlayGet(rel);
      if (!rec) continue;
      const name = rel.split("/").pop();
      items.push({ name, type: "file", rel, ext: rec.ext || extOf(name), size: rec.size || null });
      seen.add(rel);
    }
    /* olusturulan klasorler + alias ogeleri */
    for (const [rel, val] of mm) {
      if (!inFolder(rel) || seen.has(rel)) continue;
      const name = rel.split("/").pop();
      if (val === "dir") {
        items.push({ name, type: "folder", rel, size: null });
        seen.add(rel);
      } else if (typeof val === "string" && val.startsWith("via:")) {
        const t = nodeTypeM(rel, mm, ks);
        if (!t) continue;
        if (t === "file") {
          const r = await resolveRel(rel);
          const rec = r.kind === "overlay" ? await overlayGet(r.rel) : baseFiles.get(r.rel);
          items.push({ name, type: "file", rel, ext: rec ? rec.ext : extOf(name), size: rec ? rec.size : null });
        } else {
          items.push({ name, type: "folder", rel, size: null });
        }
        seen.add(rel);
      }
    }
    items.sort((a, b) => {
      if (a.type !== b.type) return a.type === "folder" ? -1 : 1;
      return trColl.compare(a.name, b.name);
    });
    treeCache.set(cacheKey, items);
    return items;
  }

  /* ---------- okuma / dosya erisimi ---------- */
  async function readBlob(rel) {
    await manifest();
    const r = await resolveRel(rel);
    if (r.kind === "overlay") {
      const rec = await overlayGet(r.rel);
      return rec.blob;
    }
    const res = await fetch("offline-data/" + r.rel, { cache: "no-store" });
    if (!res.ok) throw new Error("Dosya bulunamadi");
    return await res.blob();
  }

  function relUrlPath(rel) {
    return rel.split("/").map(encodeURIComponent).join("/");
  }
  async function offlineFileUrl(rel) {
    await manifest();
    const r = await resolveRel(rel);
    if (r.kind === "overlay") {
      const rec = await overlayGet(r.rel);
      return rec ? URL.createObjectURL(rec.blob) : "/offline-data/" + relUrlPath(rel);
    }
    return "offline-data/" + relUrlPath(r.rel);
  }

  /* ---------- arama ---------- */
  async function doSearch(q) {
    await manifest();
    const mm = await metaMap();
    const lower = q.toLowerCase();
    const out = [];
    const push = (it) => { if (out.length < 200) out.push(it); };
    for (const rel of baseFiles.keys()) {
      if (relDel(rel, mm)) continue;
      if (rel.toLowerCase().includes(lower)) {
        const meta = baseFiles.get(rel);
        push({ name: rel.split("/").pop(), type: "file", rel, ext: meta.ext, size: meta.size });
      } else {
        const ext = extOf(rel);
        if ((ext === ".txt" || ext === ".md") && metaBaseSmall(rel)) {
          try {
            const r = await resolveRel(rel);
            let text = "";
            if (r.kind === "overlay") { const rec = await overlayGet(r.rel); text = await rec.blob.text(); }
            else { const res = await fetch("offline-data/" + r.rel, { cache: "no-store" }); text = await res.text(); }
            if (text && text.toLowerCase().includes(lower)) {
              const idx = text.toLowerCase().indexOf(lower);
              const start = Math.max(0, idx - 45);
              const snippet = text.slice(start, idx + lower.length + 70).replace(/\s+/g, " ").trim();
              push({ name: rel.split("/").pop(), type: "file", rel, ext, size: (baseFiles.get(rel) || {}).size, snippet });
            }
          } catch (_) {}
        }
      }
    }
    const ks = await fileKeySet();
    for (const rel of ks) {
      if (relDel(rel, mm)) continue;
      const name = rel.split("/").pop();
      if (name.toLowerCase().includes(lower)) {
        const rec = await overlayGet(rel);
        push({ name, type: "file", rel, ext: rec ? rec.ext : extOf(name), size: rec ? rec.size : null });
      } else if (((extOf(rel) === ".txt") || (extOf(rel) === ".md")) && await smallOverlay(rel)) {
        try {
          const rec = await overlayGet(rel);
          const text = await rec.blob.text();
          if (text.toLowerCase().includes(lower)) {
            const idx = text.toLowerCase().indexOf(lower);
            const start = Math.max(0, idx - 45);
            const snippet = text.slice(start, idx + lower.length + 70).replace(/\s+/g, " ").trim();
            push({ name, type: "file", rel, ext: extOf(rel), size: rec.size, snippet });
          }
        } catch (_) {}
      }
    }
    for (const [rel, val] of mm) {
      if (val === "dir" && rel.toLowerCase().includes(lower)) {
        push({ name: rel.split("/").pop(), type: "folder", rel });
      }
    }
    return out.sort((a, b) => trColl.compare(a.rel, b.rel));
  }

  function metaBaseSmall(rel) { return (baseFiles.get(rel) || {}).size < 3 * 1024 * 1024; }
  async function smallOverlay(rel) {
    const rec = await overlayGet(rel);
    return !!rec && rec.size < 3 * 1024 * 1024;
  }

  /* ---------- varlik kontrolu ---------- */
  async function exists(rel) {
    await manifest();
    if (await isDeleted(rel)) return false;
    const t = await nodeType(rel);
    return t === "file" || t === "folder";
  }

  /* ---------- degisiklik islemleri ---------- */
  function validName(name) { return !!name && !/[\\/:*?"<>|]/.test(name); }

  async function opWrite(rel, name, content) {
    const fileName = /\.\w+$/.test(name) ? name : name + ".txt";
    const targetRel = rel ? rel + "/" + fileName : fileName;
    if (await exists(targetRel)) return { status: 409, error: "Bu isimde bir dosya zaten var" };
    await overlayPut(targetRel, new Blob([String(content == null ? "" : content)], { type: "text/plain;charset=utf-8" }), ".txt");
    return { ok: true, name };
  }

  async function opCreateFolder(rel, name) {
    if (!validName(name)) return { status: 400, error: "Gecersiz klasor adi" };
    const targetRel = rel ? rel + "/" + name : name;
    if (await exists(targetRel)) return { status: 409, error: "Bu isimde bir klasor zaten var" };
    await metaSet(targetRel, "dir");
    return { ok: true };
  }

  async function opRename(rel, newName) {
    if (!validName(newName)) return { status: 400, error: "Gecersiz ad" };
    if (!await exists(rel)) return { status: 404, error: "Oge bulunamadi" };
    const parent = rel.includes("/") ? rel.slice(0, rel.lastIndexOf("/")) : "";
    const newRel = parent ? parent + "/" + newName : newName;
    if (await exists(newRel)) return { status: 409, error: "Bu isimde bir oge zaten var" };
    await metaSet(newRel, "via:" + rel);
    await metaSet(rel, "del");
    return { ok: true, name: newName };
  }

  async function opMove(rel, dest) {
    if (!await exists(rel)) return { status: 404, error: "Oge bulunamadi" };
    if (dest && rel.indexOf(dest + "/") === 0) return { status: 400, error: "Klasor kendi icine tasinamaz" };
    const name = rel.split("/").pop();
    const newRel = dest ? dest + "/" + name : name;
    if (dest && !(await exists(dest))) return { status: 400, error: "Hedef klasor bulunamadi" };
    if (await exists(newRel)) return { status: 409, error: "Hedefte ayni isimde bir ogeler zaten var" };
    await metaSet(newRel, "via:" + rel);
    await metaSet(rel, "del");
    return { ok: true };
  }

  async function opDelete(rel) {
    const t = await nodeType(rel);
    if (!t) return { status: 404, error: "Bulunamadi" };
    await metaSet(rel, "del");
    /* overlay'de gercek blob varsa temizle (anahtar taramasi, blob okumaz) */
    const ks = await fileKeySet();
    for (const k of ks) {
      if (k === rel || (t === "folder" && k.startsWith(rel + "/"))) {
        await overlayDelete(k);
        await metaDelete(k);
      }
    }
    return { ok: true };
  }

  async function opUpload(rel, file) {
    const name = file.name || "dosya";
    const targetRel = rel ? rel + "/" + name : name;
    if (await isDeleted(targetRel)) await metaDelete(targetRel);
    const blob = new Blob([await file.arrayBuffer()], { type: file.type || "application/octet-stream" });
    await overlayPut(targetRel, blob, extOf(name));
    return { ok: true, name };
  }

  async function opUpdate(rel, content, format) {
    if (!(await exists(rel))) return { status: 404, error: "Dosya bulunamadi" };
    const text = String(content == null ? "" : content);
    if (format === "docx") {
      const blob = await docxFromText(text);
      await overlayPut(rel, blob, ".docx");
    } else {
      await overlayPut(rel, new Blob([text], { type: "text/plain;charset=utf-8" }), rel.endsWith(".md") ? ".md" : ".txt");
    }
    return { ok: true };
  }

  /* ---------- DOCX uretimi (offline edit + scan) ---------- */
  const ARABIC_RE = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      if (window.docx) return resolve(window.docx);
      const s = document.createElement("script");
      s.src = src;
      s.onload = () => resolve(window.docx);
      s.onerror = () => reject(new Error("Kutuphane yuklenemedi"));
      document.head.appendChild(s);
    });
  }
  async function docxFromText(text) {
    const docx = await loadScript("/vendor/docx/docx.umd.js");
    const nodes = text.split(/\r?\n/).map((ln) => {
      const lineText = ln === "" ? "\u00A0" : ln;
      const isRtl = ARABIC_RE.test(lineText);
      const run = isRtl
        ? new docx.TextRun({ text: lineText, font: { name: "Traditional Arabic", hint: "eastAsia" }, size: 24 })
        : new docx.TextRun({ text: lineText, font: { name: "Segoe UI" }, size: 24 });
      return new docx.Paragraph({ children: [run], bidirectional: isRtl, spacing: { after: 160 } });
    });
    const paraLeft = document.createElement("p");
    const doc = new docx.Document({ sections: [{ children: nodes }] });
    const buf = await docx.Packer.toBuffer(doc);
    return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  }

  /* ---------- Scan (camera + OCR) offline ---------- */
  let _tesseractPromise = null;
  let _ocrWorker = null;
  function loadTesseract() {
    if (!_tesseractPromise) {
      _tesseractPromise = new Promise((resolve, reject) => {
        const s = document.createElement("script");
        s.src = "/vendor/tesseract/tesseract.min.js";
        s.onload = () => resolve(window.Tesseract);
        s.onerror = () => reject(new Error("tesseract yuklenemedi"));
        document.head.appendChild(s);
      });
    }
    return _tesseractPromise;
  }
  async function ocrBlob(blob, lang) {
    const Tesseract = await loadTesseract();
    const worker = await Tesseract.createWorker(lang, 1, {
      workerPath: "/vendor/tesseract/worker.min.js",
      corePath: "/vendor/tesseract/core",
      langPath: "/vendor/tesseract/langdata",
      cacheMethod: "none",
      gzip: false,
    });
    try {
      await worker.setParameters({
        tessedit_pageseg_mode: Tesseract.PSM ? Tesseract.PSM.AUTO : 3,
        preserve_interword_spaces: "1",
      });
      const { data } = await worker.recognize(blob);
      return data.text || "";
    } finally {
      await worker.terminate();
    }
  }

  async function pdfFromImages(files) {
    const script = document.createElement("script");
    const ok = await new Promise((res) => {
      const s = document.createElement("script");
      s.src = "/vendor/pdf-lib/pdf-lib.min.js";
      s.onload = () => res(true);
      s.onerror = () => res(false);
      document.head.appendChild(s);
    });
    if (!ok) throw new Error("pdf-lib yuklenemedi");
    const PDFLib = window.PDFLib;
    const pdfDoc = await PDFLib.PDFDocument.create();
    const pageW = 595.28, pageH = 841.89, margin = 40;
    const maxW = pageW - margin * 2, maxH = pageH - margin * 2;
    for (const f of files) {
      const bytes = await f.arrayBuffer();
      const isPng = /\.png$/i.test(f.name || "");
      const img = isPng ? await pdfDoc.embedPng(bytes) : await pdfDoc.embedJpg(bytes);
      const page = pdfDoc.addPage([pageW, pageH]);
      const scale = Math.min(maxW / img.width, maxH / img.height);
      const w = img.width * scale, h = img.height * scale;
      page.drawImage(img, { x: (pageW - w) / 2, y: (pageH - h) / 2, width: w, height: h });
    }
    return new Blob([await pdfDoc.save()], { type: "application/pdf" });
  }

  async function handleScan(req) {
    const isForm = req.headers && String(req.headers["content-type"] || "").startsWith("multipart/");
    const isJson = req.headers && String(req.headers["content-type"] || "").includes("application/json");
    let body = {};
    let files = [];
    if (isForm) {
      const fd = req.body instanceof FormData ? req.body : new FormData();
      files = fd.getAll("file");
      body.path = fd.get("path") || "";
      body.name = fd.get("name") || "";
      body.format = fd.get("format") || "txt";
      body.lang = fd.get("lang") || "tur";
      body.mode = fd.get("mode") || "";
    } else if (isJson) {
      body = jsonBody(req);
    }
    const fmt = ["pdf", "docx", "txt"].includes(body.format) ? body.format : "txt";

    if (body.mode === "preview") {
      if (!files[0]) return { status: 400, error: "Gorsel secilmedi" };
      const lang = ["tur", "eng", "ara"].includes(body.lang) ? body.lang : "tur";
      const text = await ocrBlob(files[0], lang);
      return { ok: true, text };
    }

    const name = String(body.name || "").trim();
    if (!name) return { status: 400, error: "Dosya adi gerekli" };
    const baseName = name.replace(/\.(pdf|docx|txt)$/i, "");
    const targetRel = (body.path ? body.path + "/" : "") + baseName + "." + fmt;
    if (await exists(targetRel)) return { status: 409, error: "Bu isimde bir dosya zaten var" };

    if (fmt === "pdf") {
      if (!files.length) return { status: 400, error: "Gorsel secilmedi" };
      const blob = await pdfFromImages(files);
      await overlayPut(targetRel, blob, ".pdf");
    } else if (body.content != null && String(body.content).length > 0) {
      const blob = fmt === "docx" ? await docxFromText(String(body.content)) : new Blob([String(body.content)], { type: "text/plain;charset=utf-8" });
      await overlayPut(targetRel, blob, "." + fmt);
    } else {
      if (!files[0]) return { status: 400, error: "Gorsel secilmedi" };
      const lang = ["tur", "eng", "ara"].includes(body.lang) ? body.lang : "tur";
      const text = await ocrBlob(files[0], lang);
      const blob = fmt === "docx" ? await docxFromText(text) : new Blob([text], { type: "text/plain;charset=utf-8" });
      await overlayPut(targetRel, blob, "." + fmt);
    }
    return { ok: true, name: baseName + "." + fmt };
  }

  /* ---------- HTTP benzeri yanit ---------- */
  function jsonBody(req) {
    const b = req.body;
    if (typeof b === "string") {
      try { return JSON.parse(b) || {}; } catch (_) { return {}; }
    }
    return b || {};
  }
  function localRes(status, payload, isBlob) {
    return {
      ok: status >= 200 && status < 300,
      status,
      json: () => Promise.resolve(payload),
      text: () => Promise.resolve(isBlob ? "" : (typeof payload === "string" ? payload : JSON.stringify(payload))),
      arrayBuffer: async () => {
        if (isBlob) return await payload.arrayBuffer();
        return new TextEncoder().encode(typeof payload === "string" ? payload : JSON.stringify(payload)).buffer;
      },
      blob: () => Promise.resolve(isBlob ? payload : new Blob([typeof payload === "string" ? payload : JSON.stringify(payload)])),
    };
  }
  function errRes(status, error) { return localRes(status, { error }); }

  /* ---------- dapi router ---------- */
  window.dapi = async function dapi(input, opts) {
    const url = new URL(typeof input === "string" ? input : input.url, location.href);
    const method = ((opts && opts.method) || "GET").toUpperCase();
    const pathname = url.pathname;
    const req = { url, method, body: (opts && opts.body) || null, headers: (opts && opts.headers) || {} };

    try {
      await manifest();
      if (pathname === "/api/tree") {
        const rel = (url.searchParams.get("path") || "").replace(/^\//, "");
        if (!await exists(rel) && rel !== "") return errRes(404, "Klasor bulunamadi");
        const items = await treeOf(rel);
        return localRes(200, { path: rel, items });
      }
      if (pathname === "/api/search") {
        const q = (url.searchParams.get("q") || "").trim();
        if (!q) return localRes(200, { items: [] });
        const items = await doSearch(q);
        return localRes(200, { items });
      }
      if (pathname === "/api/read") {
        const rel = (url.searchParams.get("path") || "").replace(/^\//, "");
        const blob = await readBlob(rel);
        return localRes(200, await blob.text());
      }
      if (pathname === "/api/file") {
        const rel = (url.searchParams.get("path") || "").replace(/^\//, "");
        const blob = await readBlob(rel);
        const ext = extOf(rel);
        const type = { ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".pdf": "application/pdf", ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif", ".webp": "image/webp", ".bmp": "image/bmp", ".txt": "text/plain;charset=utf-8", ".md": "text/plain;charset=utf-8" }[ext] || "application/octet-stream";
        return localRes(200, blob, true);
      }
      if (pathname === "/api/folder" && method === "POST") {
        const b = jsonBody(req);
        const r = await opCreateFolder((b.path || "").replace(/^\//, ""), b.name);
        return r.status ? errRes(r.status, r.error) : localRes(200, { ok: true });
      }
      if (pathname === "/api/write" && method === "POST") {
        const b = jsonBody(req);
        const r = await opWrite((b.path || "").replace(/^\//, ""), b.name, b.content);
        return r.status ? errRes(r.status, r.error) : localRes(200, { ok: true, name: b.name });
      }
      if (pathname === "/api/upload" && method === "POST") {
        const fd = req.body instanceof FormData ? req.body : new FormData();
        const file = fd.get("file");
        if (!file) return errRes(400, "Dosya secilmedi");
        const rel = String(fd.get("path") || "").replace(/^\//, "");
        const r = await opUpload(rel, file);
        return localRes(200, { ok: true, name: file.name });
      }
      if (pathname === "/api/rename" && method === "POST") {
        const b = jsonBody(req);
        const r = await opRename((b.path || "").replace(/^\//, ""), b.newName);
        return r.status ? errRes(r.status, r.error) : localRes(200, { ok: true, name: b.newName });
      }
      if (pathname === "/api/move" && method === "POST") {
        const b = jsonBody(req);
        const r = await opMove((b.path || "").replace(/^\//, ""), (b.dest || "").replace(/^\//, ""));
        return r.status ? errRes(r.status, r.error) : localRes(200, { ok: true });
      }
      if (pathname === "/api/update" && method === "POST") {
        const b = jsonBody(req);
        const r = await opUpdate((b.path || "").replace(/^\//, ""), b.content, b.format);
        return r.status ? errRes(r.status, r.error) : localRes(200, { ok: true });
      }
      if (pathname === "/api/item" && method === "DELETE") {
        const rel = (url.searchParams.get("path") || "").replace(/^\//, "");
        const r = await opDelete(rel);
        return r.status ? errRes(r.status, r.error) : localRes(200, { ok: true });
      }
      if (pathname === "/api/scan" && method === "POST") {
        const r = await handleScan(req);
        if (r.status) return errRes(r.status, r.error);
        return localRes(200, r);
      }
      /* bilinmeyen /api* -> yerel olarak cevap yok */
      return errRes(404, "Bulunamadi");
    } catch (e) {
      return errRes(400, e.message || "Yerel islem hatasi");
    }
  };

  /* app.js'in kullandigi yardimcilar */
  window.offlineFileUrl = offlineFileUrl;
  window.offlineReadText = async function (rel) {
    const blob = await readBlob((rel || "").replace(/^\//, ""));
    return await blob.text();
  };
  window.offlineReadBlob = async function (rel) {
    return readBlob((rel || "").replace(/^\//, ""));
  };
  window.offlineIsNative = IS_NATIVE;
  window.offlineHasData = function () { return manifest().then(Boolean); };

  /* hizli sayilar: ana sayfa kartlari taban+overlay dosya adedi */
  window.offlineCountFiles = async function (prefix) {
    await manifest();
    const mm = await metaMap();
    const ks = await fileKeySet();
    const p = prefix ? prefix + "/" : "";
    let n = 0;
    for (const rel of baseFiles.keys()) {
      if (rel.startsWith(p) && !relDel(rel, mm)) n++;
    }
    for (const rel of ks) {
      if (rel.startsWith(p) && !relDel(rel, mm)) n++;
    }
    return n;
  };

  /* tasima/folder secici icin tum klasor listesi */
  window.offlineFolders = async function () {
    await manifest();
    const mm = await metaMap();
    const ks = await fileKeySet();
    const out = [];
    const push = (rel) => {
      if (!rel) return;
      out.push({ name: rel, label: rel.split("/").join(" › ") });
    };
    for (const f of baseFolders) push(f);
    for (const [rel, val] of mm) {
      if (val === "dir") push(rel);
      else if (typeof val === "string" && val.startsWith("via:")) {
        if (nodeTypeM(rel, mm, ks) === "folder") push(rel);
      }
    }
    out.sort((a, b) => trColl.compare(a.name, b.name));
    return out;
  };

  /* indirme islemleri gizli anchor ile */
  window.offlineDownload = async function (rel, name) {
    const url = await offlineFileUrl((rel || "").replace(/^\//, ""));
    const a = document.createElement("a");
    a.href = url;
    a.download = name || rel.split("/").pop() || "dosya";
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { a.remove(); }, 100);
  };
})();