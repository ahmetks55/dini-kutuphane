const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const os = require("os");
const http = require("http");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const { PDFDocument } = require("pdf-lib");
const { createWorker, PSM } = require("tesseract.js");

const PORT = process.env.PORT || 3000;
const STORAGE = path.join(__dirname, "storage");
const OCR_CACHE = path.join(__dirname, ".ocr-cache");

if (!fs.existsSync(STORAGE)) fs.mkdirSync(STORAGE, { recursive: true });
if (!fs.existsSync(OCR_CACHE)) fs.mkdirSync(OCR_CACHE, { recursive: true });

const app = express();
app.use(express.json());

app.get("/sw.js", (req, res) => {
  res.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
  res.sendFile(path.join(__dirname, "public", "sw.js"));
});

app.use(express.static(path.join(__dirname, "public")));
app.use(
  "/vendor/mammoth/mammoth.browser.min.js",
  express.static(path.join(__dirname, "node_modules", "mammoth", "mammoth.browser.min.js"))
);
app.use(
  "/vendor/pdfjs",
  express.static(path.join(__dirname, "node_modules", "pdfjs-dist", "build"))
);

const MIME = {
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".pdf": "application/pdf",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".webp": "image/webp",
  ".bmp": "image/bmp",
  ".txt": "text/plain; charset=utf-8",
  ".md": "text/plain; charset=utf-8",
};

function resolveSafe(rel) {
  const base = path.resolve(STORAGE);
  const target = path.resolve(base, rel || "");
  if (target !== base && !target.startsWith(base + path.sep)) {
    throw new Error("Gecersiz yol");
  }
  return target;
}

function compareNames(a, b) {
  return a.name.localeCompare(b.name, "tr", { sensitivity: "base" });
}

const TEXT_EXTS = new Set([".txt", ".md"]);

function listTree(dir) {
  const entries = fs.readdirSync(dir, { withFileTypes: true }).sort((a, b) => {
    if (a.isDirectory() !== b.isDirectory()) return a.isDirectory() ? -1 : 1;
    return compareNames(a, b);
  });
  return entries.map((e) => {
    const full = path.join(dir, e.name);
    const st = fs.statSync(full);
    if (e.isDirectory()) {
      return {
        name: e.name,
        type: "folder",
        size: null,
        mtime: st.mtimeMs,
        children: listTree(full),
      };
    }
    const ext = path.extname(e.name).toLowerCase();
    return {
      name: e.name,
      type: "file",
      ext,
      size: st.size,
      mtime: st.mtimeMs,
      category: MIME[ext] || "unknown",
    };
  });
}

app.get("/api/tree", (req, res) => {
  try {
    const target = resolveSafe(req.query.path || "");
    if (!fs.existsSync(target)) return res.status(404).json({ error: "Klasor bulunamadi" });
    const rel = path.relative(STORAGE, target);
    const relSlash = rel === "" ? "" : rel.split(path.sep).join("/");
    const items = listTree(target);
    res.json({ path: relSlash, items });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

function searchTree(dir, q) {
  const out = [];
  const lower = q.toLowerCase();
  const entries = fs.readdirSync(dir, { withFileTypes: true }).sort((a, b) => compareNames(a, b));
  for (const e of entries) {
    const full = path.join(dir, e.name);
    const rel = path.relative(STORAGE, full).split(path.sep).join("/");
    if (e.isDirectory()) {
      if (e.name.toLowerCase().includes(lower)) out.push({ name: e.name, type: "folder", rel });
      out.push(...searchTree(full, q));
    } else {
      const st = fs.statSync(full);
      const ext = path.extname(e.name).toLowerCase();
      const nameMatch = e.name.toLowerCase().includes(lower);
      if (nameMatch) {
        out.push({ name: e.name, type: "file", rel, ext, size: st.size });
      } else if (TEXT_EXTS.has(ext) && st.size < 3 * 1024 * 1024) {
        try {
          const content = fs.readFileSync(full, "utf8");
          const idx = content.toLowerCase().indexOf(lower);
          if (idx >= 0) {
            const start = Math.max(0, idx - 45);
            const snippet = content.slice(start, idx + lower.length + 70).replace(/\s+/g, " ").trim();
            out.push({ name: e.name, type: "file", rel, ext, size: st.size, snippet });
          }
        } catch (_) {}
      }
    }
  }
  return out;
}

app.get("/api/search", (req, res) => {
  try {
    const q = String(req.query.q || "").trim();
    if (!q) return res.json({ items: [] });
    res.json({ items: searchTree(STORAGE, q) });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.post("/api/folder", (req, res) => {
  try {
    const { path: rel, name } = req.body || {};
    if (!name || /[\\/:*?"<>|]/.test(name)) return res.status(400).json({ error: "Gecersiz klasor adi" });
    const target = path.join(resolveSafe(rel || ""), name);
    if (fs.existsSync(target)) return res.status(409).json({ error: "Bu isimde bir klasor zaten var" });
    fs.mkdirSync(target, { recursive: true });
    res.json({ ok: true });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.post("/api/write", (req, res) => {
  try {
    const { path: rel, name, content } = req.body || {};
    if (!name || /[\\/:*?"<>|]/.test(name)) return res.status(400).json({ error: "Gecersiz dosya adi" });
    const fileName = /\.\w+$/.test(name) ? name : name + ".txt";
    const target = path.join(resolveSafe(rel || ""), fileName);
    if (fs.existsSync(target)) return res.status(409).json({ error: "Bu isimde bir dosya zaten var" });
    fs.writeFileSync(target, String(content == null ? "" : content), "utf8");
    res.json({ ok: true, name });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => {
      try {
        cb(null, resolveSafe(req.body.path || ""));
      } catch (e) {
        cb(e);
      }
    },
    filename: (req, file, cb) => {
      const original = Buffer.from(file.originalname, "latin1").toString("utf8");
      cb(null, original);
    },
  }),
  limits: { fileSize: 100 * 1024 * 1024 },
});

app.post("/api/upload", (req, res) => {
  upload.single("file")(req, res, (err) => {
    if (err) return res.status(400).json({ error: "Yukleme hatasi: " + err.message });
    if (!req.file) return res.status(400).json({ error: "Dosya secilmedi" });
    res.json({ ok: true, name: req.file.filename });
  });
});

app.post("/api/rename", (req, res) => {
  try {
    const { path: rel, newName } = req.body || {};
    if (!newName || /[\\/:*?"<>|]/.test(newName)) return res.status(400).json({ error: "Gecersiz ad" });
    const oldPath = resolveSafe(rel || "");
    if (!fs.existsSync(oldPath)) return res.status(404).json({ error: "Oge bulunamadi" });
    const newPath = path.join(path.dirname(oldPath), newName);
    if (fs.existsSync(newPath)) return res.status(409).json({ error: "Bu isimde bir oge zaten var" });
    fs.renameSync(oldPath, newPath);
    res.json({ ok: true, name: newName });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.get("/api/read", (req, res) => {
  try {
    const target = resolveSafe(req.query.path || "");
    if (!fs.existsSync(target) || !fs.statSync(target).isFile()) return res.status(404).json({ error: "Dosya bulunamadi" });
    res.send(fs.readFileSync(target, "utf8"));
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.post("/api/update", async (req, res) => {
  try {
    const { path: rel, content, format } = req.body || {};
    const target = resolveSafe(rel || "");
    if (!fs.existsSync(target) || !fs.statSync(target).isFile()) return res.status(404).json({ error: "Dosya bulunamadi" });
    if (format === "docx") {
      await writeTextDoc(target, String(content == null ? "" : content), "docx");
    } else {
      fs.writeFileSync(target, String(content == null ? "" : content), "utf8");
    }
    res.json({ ok: true });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.delete("/api/item", (req, res) => {
  try {
    const target = resolveSafe(req.query.path || "");
    if (target === STORAGE) return res.status(400).json({ error: "Ana kitaplik silinemez" });
    if (!fs.existsSync(target)) return res.status(404).json({ error: "Bulunamadi" });
    const st = fs.statSync(target);
    if (st.isDirectory()) fs.rmSync(target, { recursive: true, force: true });
    else fs.unlinkSync(target);
    res.json({ ok: true });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.post("/api/move", (req, res) => {
  try {
    const { path: rel, dest } = req.body || {};
    if (!rel) return res.status(400).json({ error: "Kaynak yol gerekli" });
    const src = resolveSafe(rel);
    if (src === STORAGE) return res.status(400).json({ error: "Ana kitaplik tasinamaz" });
    if (!fs.existsSync(src)) return res.status(404).json({ error: "Bulunamadi" });
    const destDir = resolveSafe(dest);
    if (!fs.existsSync(destDir) || !fs.statSync(destDir).isDirectory()) {
      return res.status(400).json({ error: "Hedef klasor bulunamadi" });
    }
    const name = path.basename(src);
    const destFull = path.join(destDir, name);
    if (fs.existsSync(destFull)) return res.status(409).json({ error: "Hedefte ayni isimde bir ogeler zaten var" });
    const st = fs.statSync(src);
    if (st.isDirectory() && destFull.startsWith(src + path.sep)) {
      return res.status(400).json({ error: "Klasor kendi icine tasinamaz" });
    }
    fs.renameSync(src, destFull);
    res.json({ ok: true });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

app.get("/api/file", (req, res) => {
  try {
    const target = resolveSafe(req.query.path || "");
    if (!fs.existsSync(target)) return res.status(404).send("Dosya bulunamadi");
    const ext = path.extname(target).toLowerCase();
    const type = MIME[ext] || "application/octet-stream";
    res.setHeader("Content-Type", type);
    res.setHeader("Cache-Control", "no-store");
    fs.createReadStream(target).pipe(res);
  } catch (e) {
    res.status(400).send(e.message);
  }
});

app.get("/api/download", (req, res) => {
  try {
    const target = resolveSafe(req.query.path || "");
    if (!fs.existsSync(target)) return res.status(404).send("Dosya bulunamadi");
    res.download(target, path.basename(target));
  } catch (e) {
    res.status(400).send(e.message);
  }
});

/* ---------------- Belge tarama (OCR) ---------------- */
const SCAN_LANGS = ["tur", "eng", "ara"];
const ARABIC_RE = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
const scanUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 30 * 1024 * 1024 },
});

async function imageToPdf(imagePaths, pdfPath) {
  const pdfDoc = await PDFDocument.create();
  const pageW = 595.28;
  const pageH = 841.89;
  const margin = 40;
  const maxW = pageW - margin * 2;
  const maxH = pageH - margin * 2;
  for (const imagePath of imagePaths) {
    const ext = path.extname(imagePath).toLowerCase();
    const bytes = fs.readFileSync(imagePath);
    const img = ext === ".png" ? await pdfDoc.embedPng(bytes) : await pdfDoc.embedJpg(bytes);
    const page = pdfDoc.addPage([pageW, pageH]);
    const scale = Math.min(maxW / img.width, maxH / img.height);
    const w = img.width * scale;
    const h = img.height * scale;
    page.drawImage(img, { x: (pageW - w) / 2, y: (pageH - h) / 2, width: w, height: h });
  }
  fs.writeFileSync(pdfPath, await pdfDoc.save());
}

async function ocrText(imagePath, lang) {
  const worker = await createWorker(lang, 1, { cachePath: OCR_CACHE });
  try {
    await worker.setParameters({
      tessedit_pageseg_mode: PSM.AUTO,
      preserve_interword_spaces: "1",
    });
    const { data } = await worker.recognize(imagePath);
    return data.text;
  } finally {
    await worker.terminate();
  }
}

async function writeTextDoc(target, text, fmt) {
  if (fmt === "txt") {
    fs.writeFileSync(target, text, "utf8");
  } else {
    const paras = text.split(/\r?\n/).map((ln) => {
      const isRtl = ARABIC_RE.test(ln);
      const run = isRtl
        ? new TextRun({ text: ln, font: { name: "Traditional Arabic", hint: "eastAsia" }, size: 24 })
        : new TextRun({ text: ln, font: { name: "Segoe UI" }, size: 24 });
      return new Paragraph({ children: [run], bidirectional: isRtl, spacing: { after: 160 } });
    });
    fs.writeFileSync(target, await Packer.toBuffer(new Document({ sections: [{ children: paras }] })));
  }
}

function tempImage(file) {
  const ext = path.extname(file.originalname).toLowerCase();
  const tmp = path.join(os.tmpdir(), "dini-scan-" + Date.now() + (ext === ".png" ? ".png" : ".jpg"));
  fs.writeFileSync(tmp, file.buffer);
  return tmp;
}

app.post("/api/scan", scanUpload.array("file", 20), async (req, res) => {
  const tmps = [];
  try {
    const body = req.body || {};
    const { path: rel, name, format, lang, mode, content } = body;
    const fmt = ["pdf", "docx", "txt"].includes(format) ? format : "txt";
    const files = req.files || [];

    if (mode === "preview") {
      if (!files[0]) return res.status(400).json({ error: "Gorsel secilmedi" });
      const ocrLang = SCAN_LANGS.includes(lang) ? lang : "tur";
      const tmp = tempImage(files[0]);
      tmps.push(tmp);
      const text = await ocrText(tmp, ocrLang);
      return res.json({ ok: true, text });
    }

    if (!name || !String(name).trim()) return res.status(400).json({ error: "Dosya adi gerekli" });
    const baseName = String(name).replace(/\.(pdf|docx|txt)$/i, "");
    const target = path.join(resolveSafe(rel || ""), baseName + "." + fmt);
    if (fs.existsSync(target)) return res.status(409).json({ error: "Bu isimde bir dosya zaten var" });

    if (fmt === "pdf") {
      if (files.length === 0) return res.status(400).json({ error: "Gorsel secilmedi" });
      for (const f of files) tmps.push(tempImage(f));
      await imageToPdf(tmps, target);
    } else if (content != null && String(content).length > 0) {
      await writeTextDoc(target, String(content), fmt);
    } else {
      if (!files[0]) return res.status(400).json({ error: "Gorsel secilmedi" });
      const ocrLang = SCAN_LANGS.includes(lang) ? lang : "tur";
      const tmp = tempImage(files[0]);
      tmps.push(tmp);
      await writeTextDoc(target, await ocrText(tmp, ocrLang), fmt);
    }
    res.json({ ok: true, name: path.basename(target) });
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    for (const t of tmps) {
      if (t && fs.existsSync(t)) {
        try { fs.unlinkSync(t); } catch {}
      }
    }
  }
});

function lanAddresses() {
  const nets = os.networkInterfaces();
  const out = [];
  for (const name of Object.keys(nets)) {
    for (const net of nets[name] || []) {
      if (net.family === "IPv4" && !net.internal) out.push(net.address);
    }
  }
  return out;
}

const server = http.createServer(app);
server.listen(PORT, "0.0.0.0", () => {
  console.log("========================================");
  console.log("  DINI KUTUPHANE UYGULAMASI");
  console.log("========================================");
  console.log("  Bu bilgisayarda:  http://localhost:" + PORT);
  lanAddresses().forEach((ip) => {
    console.log("  Telefon/Tablet:   http://" + ip + ":" + PORT);
  });
  console.log("========================================");
});
