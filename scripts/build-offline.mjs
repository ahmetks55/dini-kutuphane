import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, "..");
const PUB = path.join(ROOT, "public");
const STORAGE = path.join(ROOT, "storage");
const NM = path.join(ROOT, "node_modules");

const OFFLINE_DATA = path.join(PUB, "offline-data");
const VENDOR = path.join(PUB, "vendor");

const LANGS = ["tur", "eng", "ara"];
const TESSDATA_URL = (lang) => `https://tessdata.projectnaptha.com/4.0.0/${lang}.traineddata.gz`;

function rmrf(p) {
  fs.rmSync(p, { recursive: true, force: true });
}
function copyFile(src, dest) {
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.copyFileSync(src, dest);
}

function walk(dir, rel, out) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, e.name);
    const childRel = rel ? rel + "/" + e.name : e.name;
    if (e.isDirectory()) walk(full, childRel, out);
    else out.push(childRel);
  }
}

async function main() {
  console.log("== offline-data bundle ==");
  rmrf(OFFLINE_DATA);
  fs.mkdirSync(OFFLINE_DATA, { recursive: true });
  const files = [];
  walk(STORAGE, "", files);
  const manifest = { version: 1, files: {} };
  for (const rel of files) {
    const full = path.join(STORAGE, rel);
    const st = fs.statSync(full);
    const dest = path.join(OFFLINE_DATA, rel);
    fs.mkdirSync(path.dirname(dest), { recursive: true });
    fs.copyFileSync(full, dest);
    manifest.files[rel] = { ext: path.extname(rel).toLowerCase(), size: st.size, mtime: st.mtimeMs };
  }
  fs.writeFileSync(path.join(OFFLINE_DATA, "index.json"), JSON.stringify(manifest));
  const mb = (fs.statSync(path.join(OFFLINE_DATA, "index.json")).size + fs.readdirSync(OFFLINE_DATA).length) / 1;
  console.log(`  ${files.length} dosya + index.json -> public/offline-data (${(offlineSize() / 1024 / 1024).toFixed(1)} MB)`);

  console.log("== vendor bundle ==");
  rmrf(VENDOR);
  const copies = [
    ["mammoth/mammoth.browser.min.js", "vendor/mammoth/mammoth.browser.min.js"],
    ["pdfjs-dist/build/pdf.min.js", "vendor/pdfjs/pdf.min.js"],
    ["pdfjs-dist/build/pdf.worker.min.js", "vendor/pdfjs/pdf.worker.min.js"],
    ["pdf-lib/dist/pdf-lib.min.js", "vendor/pdf-lib/pdf-lib.min.js"],
    ["docx/build/index.umd.js", "vendor/docx/docx.umd.js"],
    ["tesseract.js/dist/tesseract.min.js", "vendor/tesseract/tesseract.min.js"],
    ["tesseract.js/dist/worker.min.js", "vendor/tesseract/worker.min.js"],
  ];
  for (const [srcRel, destRel] of copies) {
    const src = path.join(NM, srcRel);
    if (!fs.existsSync(src)) throw new Error("Eksik vendor: " + srcRel);
    copyFile(src, path.join(PUB, destRel));
    console.log("  + " + destRel);
  }
  const coreDst = path.join(VENDOR, "tesseract", "core");
  fs.mkdirSync(coreDst, { recursive: true });
  for (const name of ["tesseract-core-lstm.wasm.js", "tesseract-core-simd-lstm.wasm.js", "tesseract-core-relaxedsimd-lstm.wasm.js"]) {
    const src = path.join(NM, "tesseract.js-core", name);
    if (!fs.existsSync(src)) throw new Error("Eksik tesseract core: " + name);
    copyFile(src, path.join(coreDst, name));
    console.log("  + vendor/tesseract/core/" + name);
  }

  console.log("== tessdata indirme ==");
  const langDst = path.join(VENDOR, "tesseract", "langdata");
  fs.mkdirSync(langDst, { recursive: true });
  /* Dikkat: Android AAPT2, assets icine .gz uzantisini paketlerken siyirir.
   * Bu yuzden dosyayi UZANTISIZ sakliyoruz (<lang>.traineddata), icine gz
   * byte'lari yaziliyor. tesseract.js worker'i gzip bayraghi false ve magic
   * byte kontroluyle (1F 8B) bunu acar. */
  for (const lang of LANGS) {
    const dest = path.join(langDst, lang + ".traineddata");
    if (fs.existsSync(dest)) {
      console.log(`  ${lang}.traineddata zaten var (${(fs.statSync(dest).size / 1024 / 1024).toFixed(1)} MB)`);
      continue;
    }
    const res = await fetch(TESSDATA_URL(lang));
    if (!res.ok) throw new Error(`tessdata ${lang} indirilemedi: ${res.status}`);
    const buf = Buffer.from(await res.arrayBuffer());
    fs.writeFileSync(dest, buf);
    console.log(`  + ${lang}.traineddata (gz, ${(buf.length / 1024 / 1024).toFixed(1)} MB)`);
  }
  console.log("== tamam ==");
}

function offlineSize() {
  const all = [];
  walk(OFFLINE_DATA, "", all);
  return all.reduce((s, f) => s + fs.statSync(path.join(OFFLINE_DATA, f)).size, 0);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});