const state = {
  path: "",
  items: [],
  currentFile: null,
  uploadPath: "",
  scanPath: "",
  scanFile: null,
  scanOcrFile: null,
  scanTextReady: false,
  scanFiles: [],
  moveItem: null,
  moveRel: "",
};

const MAIN_CATS = ["Dualar", "Sureler", "Mevlid", "Ilahiler", "Kasideler", "Salavatlar"];
const CAT_ICONS = { Dualar: "🤲", Sureler: "📖", Mevlid: "🕌", Ilahiler: "📄", Kasideler: "📜", Salavatlar: "🕌" };

function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function joinPath(base, name) {
  return base ? base + "/" + name : name;
}

function iconFor(item) {
  if (item.type === "folder") return "📁";
  const ext = (item.ext || "").toLowerCase();
  if (ext === ".docx" || ext === ".doc") return "📄";
  if (ext === ".pdf") return "📕";
  if ([".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"].includes(ext)) return "🖼️";
  return "📝";
}

function fmtSize(bytes) {
  if (bytes == null) return "";
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

function toast(msg, type) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.classList.toggle("error", type === "error");
  t.classList.add("show");
  clearTimeout(t._h);
  t._h = setTimeout(() => t.classList.remove("show"), 2400);
}

/* ---------------- Category menu (hamburger) ---------------- */
function renderCategories() {
  const menu = document.getElementById("catMenu");
  menu.innerHTML = "";
  const chips = [{ name: "", label: "🏠 Ana Kitaplik" }, ...MAIN_CATS.map((c) => ({ name: c, label: (CAT_ICONS[c] || "📁") + " " + c }))];
  chips.forEach((ch) => {
    const b = document.createElement("button");
    b.className = "cat-menu-item" + (state.path === ch.name ? " active" : "");
    b.textContent = ch.label;
    b.onclick = () => { closeCatMenu(); go(ch.name); };
    menu.appendChild(b);
  });
  (state.items || []).forEach((it) => {
    if (it.type === "folder" && !MAIN_CATS.includes(it.name)) {
      const b = document.createElement("button");
      b.className = "cat-menu-item";
      b.textContent = "📁 " + it.name;
      b.onclick = () => { closeCatMenu(); go(it.name); };
      menu.appendChild(b);
    }
  });
}

function toggleCatMenu() {
  const menu = document.getElementById("catMenu");
  menu.hidden = !menu.hidden;
  if (!menu.hidden) renderCategories();
}
function closeCatMenu() {
  document.getElementById("catMenu").hidden = true;
}

/* ---------------- Breadcrumb ---------------- */
function renderBreadcrumb() {
  const el = document.getElementById("breadcrumb");
  if (!state.path) { el.innerHTML = ""; return; }
  const parts = state.path.split("/");
  let html = '<a onclick="goHome()">🏠 Ana Kitaplik</a>';
  let acc = "";
  parts.forEach((p, i) => {
    acc = acc ? acc + "/" + p : p;
    html += ' <span class="sep">›</span> ';
    html += i === parts.length - 1 ? `<span><strong>${esc(p)}</strong></span>` : `<a onclick="openFolder('${esc(acc)}')">${esc(p)}</a>`;
  });
  el.innerHTML = html;
}

/* ---------------- Load & render ---------------- */
async function load() {
  renderCategories();
  renderBreadcrumb();
  const grid = document.getElementById("grid");
  grid.innerHTML = '<p class="empty">Yukleniyor...</p>';
  try {
    const res = await fetch("/api/tree?path=" + encodeURIComponent(state.path));
    if (!res.ok) throw new Error();
    const data = await res.json();
    state.items = data.items;
    renderGrid();
  } catch (e) {
    grid.innerHTML = '<p class="empty">Klasor okunamadi. Sunucu calisiyor mu kontrol edin.</p>';
  }
}

function renderGrid() {
  const grid = document.getElementById("grid");
  document.getElementById("folderTitle").textContent = state.path ? state.path.split("/").pop() : "Ana Kitaplik";

  if (!state.path) {
    document.getElementById("itemCount").textContent = state.items.length + " ogeler";
    renderHome();
    return;
  }

  document.getElementById("itemCount").textContent = state.items.length + " ogeler";
  if (state.items.length === 0) {
    grid.innerHTML = '<p class="empty">Bu klasorde hicbir sey yok.<br>Yukaridan "Klasor" veya "Yukle" butonu ile ekleyebilirsiniz.</p>';
    return;
  }

  grid.className = "grid";
  grid.innerHTML = "";
  state.items.forEach((it) => {
    grid.appendChild(makeItem(it));
  });
}

function renderHome() {
  const grid = document.getElementById("grid");
  grid.className = "grid";
  grid.innerHTML = "";

  const topLevel = state.items;
  MAIN_CATS.forEach((cat) => {
    const found = topLevel.find((f) => f.name === cat && f.type === "folder");
    const count = found && found.children ? found.children.filter((c) => c.type === "file").length : 0;
    const card = document.createElement("div");
    card.className = "home-card";
    card.innerHTML = `<div class="ic">${CAT_ICONS[cat] || "📁"}</div><div class="nm">${cat}</div><div class="ct">${count} dosya</div>`;
    card.onclick = () => { go(cat); };
    grid.appendChild(card);
  });

  topLevel.forEach((it) => {
    if (it.type === "folder" && !MAIN_CATS.includes(it.name)) {
      const card = document.createElement("div");
      card.className = "home-card";
      const cnt = it.children ? it.children.filter((c) => c.type === "file").length : 0;
      card.innerHTML = `<div class="ic">📁</div><div class="nm">${esc(it.name)}</div><div class="ct">${cnt} dosya</div>`;
      card.onclick = () => { go(it.name); };
      grid.appendChild(card);
    }
  });

  topLevel.forEach((it) => {
    if (it.type === "file") grid.appendChild(makeItem(it));
  });

  if (grid.children.length === 0) {
    grid.innerHTML = '<p class="empty">Ana kitaplik bos.</p>';
  }
}

function makeItem(it) {
  const div = document.createElement("div");
  div.className = "item" + (it.type === "folder" ? " folder" : "");
  const rel = joinPath(state.path, it.name);

  const more = document.createElement("button");
  more.className = "more";
  more.textContent = "⋮";
  more.title = "Islemler";
  more.onclick = (e) => {
    e.stopPropagation();
    if (activeMenu && activeMenu._btn === more) closeItemMenu();
    else openItemMenu(it, rel, more);
  };

  const thumb = document.createElement("div");
  thumb.className = "thumb";
  thumb.textContent = iconFor(it);
  thumb.onclick = () => (it.type === "folder" ? openFolder(rel) : openFile(it, rel));

  const meta = document.createElement("div");
  meta.className = "meta";
  meta.onclick = () => (it.type === "folder" ? openFolder(rel) : openFile(it, rel));
  const sub = it.type === "folder" ? "Klasor" : fmtSize(it.size);
  meta.innerHTML = `<div class="nm">${esc(it.name)}</div><div class="sz">${sub}</div>`;

  div.appendChild(more);
  div.appendChild(thumb);
  div.appendChild(meta);
  return div;
}

/* ---------------- Item menu (3 nokta) ---------------- */
let activeMenu = null;

function closeItemMenu() {
  if (activeMenu) {
    activeMenu.remove();
    activeMenu = null;
  }
}

function openItemMenu(it, rel, anchor) {
  closeItemMenu();
  const menu = document.createElement("div");
  menu.className = "item-menu";

  const isEditable = it.type === "file" && /\.(txt|md|docx)$/i.test(it.name);

  const adlandir = document.createElement("button");
  adlandir.className = "menu-item";
  adlandir.textContent = "✏️ Yeniden Adlandır";
  adlandir.onclick = (e) => {
    e.stopPropagation();
    closeItemMenu();
    openRename(it, rel);
  };
  menu.appendChild(adlandir);

  if (isEditable) {
    const duzenle = document.createElement("button");
    duzenle.className = "menu-item";
    duzenle.textContent = "📝 Düzenle";
    duzenle.onclick = (e) => {
      e.stopPropagation();
      closeItemMenu();
      openEdit(it, rel);
    };
    menu.appendChild(duzenle);
  }

  const tasi = document.createElement("button");
  tasi.className = "menu-item";
  tasi.textContent = "📦 Tasi";
  tasi.onclick = (e) => {
    e.stopPropagation();
    closeItemMenu();
    openMoveModal(it, rel);
  };
  menu.appendChild(tasi);

  const sil = document.createElement("button");
  sil.className = "menu-item danger";
  sil.textContent = "🗑 Sil";
  sil.onclick = (e) => {
    e.stopPropagation();
    closeItemMenu();
    deleteItem(it, rel);
  };
  menu.appendChild(sil);

  document.body.appendChild(menu);
  menu._btn = anchor;
  activeMenu = menu;
  const r = anchor.getBoundingClientRect();
  menu.style.position = "fixed";
  menu.style.top = r.bottom + 6 + "px";
  menu.style.left = Math.min(r.left, window.innerWidth - menu.offsetWidth - 8) + "px";
}

/* ---------------- Rename ---------------- */
let renameRel = "";
function openRename(it, rel) {
  renameRel = rel;
  document.getElementById("renameInput").value = it.name;
  document.getElementById("renameModal").hidden = false;
  document.getElementById("renameInput").focus();
  document.getElementById("renameInput").select();
}
function closeRename() { document.getElementById("renameModal").hidden = true; }

async function doRename() {
  const newName = document.getElementById("renameInput").value.trim();
  if (!newName) { toast("Yeni ad gerekli", "error"); return; }
  const res = await fetch("/api/rename", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: renameRel, newName }),
  });
  if (res.ok) { toast("Yeniden adlandirildi"); closeRename(); load(); }
  else {
    const data = await res.json().catch(() => ({}));
    toast(data.error || "Yeniden adlandirilamadi", "error");
  }
}

/* ---------------- Edit (metin dosyalari) ---------------- */
let editRel = "";
function openEdit(it, rel) {
  editRel = rel;
  document.getElementById("editTitle").textContent = "Düzenle: " + it.name;
  const ta = document.getElementById("editContent");
  ta.value = "Yukleniyor...";
  document.getElementById("editModal").hidden = false;
  if (/\.docx$/i.test(it.name)) {
    fetch(fileUrl(rel))
      .then((r) => (r.ok ? r.arrayBuffer() : Promise.reject()))
      .then((buf) => mammoth.convertToHtml({ arrayBuffer: buf }))
      .then((res) => {
        const tmp = document.createElement("div");
        tmp.innerHTML = res.value;
        tmp.querySelectorAll("br").forEach((b) => b.replaceWith(document.createTextNode("\n")));
        tmp.querySelectorAll("p, li, h1, h2, h3, h4, h5, h6, tr, div").forEach((el) => {
          el.appendChild(document.createTextNode("\n"));
        });
        ta.value = tmp.textContent || "";
      })
      .catch(() => {
        ta.value = "";
        toast("Icerik okunamadi", "error");
      });
  } else {
    fetch("/api/read?path=" + encodeURIComponent(rel))
      .then((r) => (r.ok ? r.text() : Promise.reject()))
      .then((t) => { ta.value = t; })
      .catch(() => {
        ta.value = "";
        toast("Icerik okunamadi", "error");
      });
  }
}
function closeEdit() { document.getElementById("editModal").hidden = true; }

async function doEdit() {
  const content = document.getElementById("editContent").value;
  const isDocx = /\.docx$/i.test(editRel);
  const res = await fetch("/api/update", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: editRel, content, format: isDocx ? "docx" : undefined }),
  });
  if (res.ok) { toast("Kaydedildi"); closeEdit(); }
  else {
    const data = await res.json().catch(() => ({}));
    toast(data.error || "Kaydedilemedi", "error");
  }
}

/* ---------------- Move item ---------------- */
function openMoveModal(it, rel) {
  state.moveItem = it;
  state.moveRel = rel;
  const folderSelect = document.getElementById("moveFolder");
  folderSelect.innerHTML = '<option value="">Yukleniyor...</option>';
  document.getElementById("moveTarget").textContent = rel;
  document.getElementById("moveModal").hidden = false;
  loadFolderOptions();
}

function collectFolders(items, acc, prefix) {
  items.forEach((f) => {
    if (f.type === "folder") {
      const rel = prefix ? prefix + "/" + f.name : f.name;
      acc.push({ name: rel, label: (prefix ? prefix.replace(/\//g, " › ") + " › " : "") + f.name });
      if (f.children) collectFolders(f.children, acc, rel);
    }
  });
  return acc;
}

async function loadFolderOptions() {
  const select = document.getElementById("moveFolder");
  try {
    const res = await fetch("/api/tree?path=");
    const data = await res.json().catch(() => ({}));
    const opts = [{ name: "", label: "🏠 Ana Kitaplik" }, ...collectFolders(data.items || [], [], "")];
    const cur = state.moveRel.split("/").slice(0, -1).join("/");
    select.innerHTML = "";
    opts.forEach((o) => {
      const option = document.createElement("option");
      option.value = o.name;
      option.textContent = o.label;
      if (o.name === cur) option.selected = true;
      select.appendChild(option);
    });
  } catch (e) {
    select.innerHTML = '<option value="">Klasorler yuklenemedi</option>';
  }
}

function closeMoveModal() { document.getElementById("moveModal").hidden = true; }

async function doMove() {
  const dest = document.getElementById("moveFolder").value;
  if (!state.moveItem) return;
  const res = await fetch("/api/move", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: state.moveRel, dest }),
  });
  const data = await res.json().catch(() => ({}));
  if (res.ok) { toast("Tasindi: " + state.moveItem.name); closeMoveModal(); load(); }
  else toast(data.error || "Tasinamadi", "error");
}

/* ---------------- Navigation ---------------- */
function go(p) {
  const h = p ? "/" + p : "/";
  if (location.hash === "#" + h) {
    state.path = p;
    load();
  } else {
    location.hash = h;
  }
}
function openFolder(rel) { go(rel); }
function goHome() { go("/"); }
window.addEventListener("hashchange", () => {
  const raw = decodeURIComponent(location.hash.replace(/^#/, ""));
  const [pathPart, filePart] = raw.split(":f=");
  state.path = pathPart.replace(/^\//, "");
  if (filePart !== undefined) {
    if (document.getElementById("viewerModal").hidden) {
      const f = (state.items || []).find((i) => i.type === "file" && i.name === decodeURIComponent(filePart));
      if (f) openFile(f, joinPath(state.path, f.name));
    }
  } else {
    closeViewer();
    load();
  }
});
window.addEventListener("popstate", () => {
  const openModal = Array.from(document.querySelectorAll(".modal-overlay")).find((m) => !m.hidden);
  if (openModal) {
    const id = openModal.id;
    if (id === "viewerModal") { closeViewer(); return; }
    if (id === "searchModal") closeSearch();
    else if (id === "uploadModal") closeUpload();
    else if (id === "scanModal") closeScan();
    else if (id === "cropModal") closeCrop();
    else if (id === "moveModal") closeMoveModal();
    else if (id === "newTextModal") closeNewText();
    else if (id === "renameModal") closeRename();
    else if (id === "editModal") closeEdit();
    history.forward();
    return;
  }
  const raw = decodeURIComponent(location.hash.replace(/^#/, ""));
  const [pathPart] = raw.split(":f=");
  if (!pathPart || pathPart === "/") {
    if (history.state && history.state.__guard) return;
    history.pushState({ __guard: true }, "", "#/");
    state.path = "";
    load();
  }
});

async function deleteItem(it, rel) {
  if (!confirm(`"${it.name}" silinsin mi?\nBu islem geri alinamaz.`)) return;
  const res = await fetch("/api/item?path=" + encodeURIComponent(rel), { method: "DELETE" });
  if (res.ok) { toast("Silindi: " + it.name); load(); }
  else toast("Silme hatasi", "error");
}

/* ---------------- Folders ---------------- */
function promptNewFolder() {
  const name = prompt("Yeni klasor adi:");
  if (!name || !name.trim()) return;
  createFolder(name.trim());
}

async function createFolder(name) {
  const res = await fetch("/api/folder", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: state.path, name }),
  });
  if (res.ok) { toast("Klasor olusturuldu: " + name); load(); }
  else {
    const data = await res.json().catch(() => ({}));
    toast(data.error || "Klasor olusturulamadi", "error");
  }
}

/* ---------------- Upload menu (Yükle dropdown) ---------------- */
function toggleUploadMenu() {
  const menu = document.getElementById("uploadMenu");
  menu.hidden = !menu.hidden;
}
function closeUploadMenu() {
  document.getElementById("uploadMenu").hidden = true;
}

/* ---------------- Upload ---------------- */
function openUpload() {
  state.uploadPath = state.path;
  document.getElementById("uploadTarget").textContent = state.path ? state.path.split("/").join(" › ") : "Ana Kitaplik";
  document.getElementById("fileInput").value = "";
  document.getElementById("uploadModal").hidden = false;
}
function closeUpload() { document.getElementById("uploadModal").hidden = true; }

async function uploadFiles() {
  const input = document.getElementById("fileInput");
  const files = Array.from(input.files || []);
  if (files.length === 0) { toast("Dosya secilmedi", "error"); return; }
  let ok = 0;
  for (const f of files) {
    const fd = new FormData();
    fd.append("path", state.uploadPath);
    fd.append("file", f);
    const res = await fetch("/api/upload", { method: "POST", body: fd });
    if (res.ok) ok++;
  }
  toast(ok + "/" + files.length + " dosya yuklendi");
  closeUpload();
  load();
}

/* ---------------- New text file ---------------- */
function openNewText() {
  document.getElementById("newTextTarget").textContent = state.path ? state.path.split("/").join(" › ") : "Ana Kitaplik";
  document.getElementById("newTextName").value = "";
  document.getElementById("newTextContent").value = "";
  document.getElementById("newTextModal").hidden = false;
  document.getElementById("newTextName").focus();
}
function closeNewText() { document.getElementById("newTextModal").hidden = true; }

async function saveNewText() {
  const nameEl = document.getElementById("newTextName");
  const name = nameEl.value.trim();
  if (!name) { toast("Dosya adi gerekli", "error"); nameEl.focus(); return; }
  const content = document.getElementById("newTextContent").value;
  const res = await fetch("/api/write", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: state.path, name, content }),
  });
  if (res.ok) { toast("Kaydedildi: " + name + ".txt"); closeNewText(); load(); }
  else {
    const data = await res.json().catch(() => ({}));
    toast(data.error || "Kaydedilemedi", "error");
  }
}

/* ---------------- Scan (camera + OCR) ---------------- */
function openScan() {
  state.scanPath = state.path;
  state.scanFile = null;
  state.scanOcrFile = null;
  state.scanTextReady = false;
  state.scanFiles = [];
  document.getElementById("scanTarget").textContent = state.path ? state.path.split("/").join(" › ") : "Ana Kitaplik";
  const d = new Date();
  const stamp = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0") + "-" + String(d.getHours()).padStart(2, "0") + String(d.getMinutes()).padStart(2, "0");
  document.getElementById("scanName").value = "Tarama-" + stamp;
  const stack = document.getElementById("scanStack");
  stack.innerHTML = "";
  stack.hidden = true;
  document.getElementById("scanInput").value = "";
  document.querySelector('input[name="sfmt"][value="docx"]').checked = true;
  document.getElementById("scanProgress").hidden = true;
  document.getElementById("scanSaveBtn").disabled = false;
  document.getElementById("scanOcrBtn").disabled = false;
  const ta = document.getElementById("scanText");
  ta.value = "";
  ta.hidden = true;
  toggleScanFormat();
  document.getElementById("scanModal").hidden = false;
}
function closeScan() { document.getElementById("scanModal").hidden = true; }

function toggleScanFormat() {
  const fmt = document.querySelector('input[name="sfmt"]:checked').value;
  const needsOcr = fmt === "docx" || fmt === "txt";
  document.getElementById("scanLangRow").style.display = needsOcr ? "" : "none";
  document.getElementById("scanOcrHint").hidden = !needsOcr;
  document.getElementById("scanOcrBtn").hidden = !needsOcr;
  const ta = document.getElementById("scanText");
  if (!needsOcr) {
    state.scanTextReady = false;
    ta.hidden = true;
  } else if (state.scanTextReady) {
    ta.hidden = false;
  }
}

async function runScanOcr() {
  if (!state.scanFile) { toast("Once kamera ile gorsel cekin", "error"); return; }
  const lang = document.getElementById("scanLang").value;
  const btn = document.getElementById("scanOcrBtn");
  const prog = document.getElementById("scanProgress");
  btn.disabled = true;
  prog.hidden = false;
  prog.textContent = "Belge taranıyor (OCR), ilk kullanimda dil verisi indirilir...";
  try {
    const fd = new FormData();
    fd.append("path", state.scanPath);
    fd.append("format", "txt");
    fd.append("lang", lang);
    fd.append("mode", "preview");
    fd.append("file", state.scanOcrFile || state.scanFile);
    const res = await fetch("/api/scan", { method: "POST", body: fd });
    const data = await res.json().catch(() => ({}));
    if (res.ok) {
      state.scanTextReady = true;
      const ta = document.getElementById("scanText");
      const prev = ta.value.trim();
      const add = (data.text || "").trim();
      ta.value = prev ? prev + "\n\n" + add : add;
      ta.hidden = false;
      prog.textContent = "Metin tanindi ve alttakine eklendi. Baska sayfa icin tekrar cekip onizleyin, sonra Kaydet'e basin.";
    } else {
      toast(data.error || "Tarama basarisiz", "error");
      prog.hidden = true;
    }
  } catch (e) {
    toast("Baglanti hatasi", "error");
  } finally {
    btn.disabled = false;
    prog.hidden = prog.textContent === "Belge taranıyor (OCR), ilk kullanimda dil verisi indirilir..." ? true : prog.hidden;
  }
}

async function saveScan() {
  const fmt = document.querySelector('input[name="sfmt"]:checked').value;
  const name = document.getElementById("scanName").value.trim();
  if (!name) { toast("Dosya adi gerekli", "error"); return; }
  const btn = document.getElementById("scanSaveBtn");
  const prog = document.getElementById("scanProgress");
  btn.disabled = true;
  try {
    let res;
    if (fmt === "pdf") {
      if (state.scanFiles.length === 0) { toast("Once kamera ile gorsel cekin", "error"); return; }
      const fd = new FormData();
      fd.append("path", state.scanPath);
      fd.append("name", name);
      fd.append("format", "pdf");
      for (const f of state.scanFiles) fd.append("file", f);
      prog.hidden = false;
      prog.textContent = "PDF olusturuluyor (" + state.scanFiles.length + " sayfa)...";
      res = await fetch("/api/scan", { method: "POST", body: fd });
    } else {
      if (!state.scanTextReady) { toast("Once Tara (Onizleme) butonuna basin", "error"); return; }
      const text = document.getElementById("scanText").value;
      prog.hidden = false;
      prog.textContent = "Kaydediliyor...";
      res = await fetch("/api/scan", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ path: state.scanPath, name, format: fmt, content: text }),
      });
    }
    const data = await res.json().catch(() => ({}));
    if (res.ok) { toast("Kaydedildi: " + (data.name || name)); closeScan(); load(); }
    else toast(data.error || "Kayit basarisiz", "error");
  } catch (e) {
    toast("Baglanti hatasi", "error");
  } finally {
    btn.disabled = false;
    prog.hidden = true;
  }
}

/* ---------------- Search ---------------- */
let searchTimer = null;

function openSearch() {
  document.getElementById("searchInput").value = "";
  document.getElementById("searchResults").innerHTML = "";
  document.getElementById("searchModal").hidden = false;
  setTimeout(() => document.getElementById("searchInput").focus(), 50);
}
function closeSearch() { document.getElementById("searchModal").hidden = true; }

function onSearchInput() {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(runSearch, 250);
}

async function runSearch() {
  const q = document.getElementById("searchInput").value.trim();
  const box = document.getElementById("searchResults");
  if (!q) { box.innerHTML = ""; return; }
  box.innerHTML = '<p class="search-loading">Araniyor...</p>';
  try {
    const res = await fetch("/api/search?q=" + encodeURIComponent(q));
    const data = await res.json().catch(() => ({}));
    if (!res.ok) { box.innerHTML = '<p class="empty">Arama yapilamadi.</p>'; return; }
    const items = data.items || [];
    if (items.length === 0) {
      box.innerHTML = '<p class="empty">Sonuc bulunamadi.</p>';
      return;
    }
    box.innerHTML = "";
    items.slice(0, 50).forEach((it) => {
      const row = document.createElement("button");
      row.className = "search-result";
      const parent = it.rel.split("/").slice(0, -1).join("/");
      row.innerHTML = `<span class="sr-icon">${it.type === "folder" ? "📁" : iconFor(it)}</span>
        <span class="sr-main"><span class="sr-name">${esc(it.name)}</span>
        <span class="sr-path">${esc(parent || "Ana Kitaplik")}</span>${it.snippet ? `<span class="sr-snippet">…${esc(it.snippet)}…</span>` : ""}</span>`;
      row.onclick = () => {
        closeSearch();
        if (it.type === "folder") openFolder(it.rel);
        else openFile(it, it.rel);
      };
      box.appendChild(row);
    });
  } catch (e) {
    box.innerHTML = '<p class="empty">Baglanti hatasi.</p>';
  }
}

/* ---------------- Viewer ---------------- */
let viewingRel = null;

function fileUrl(rel) { return "/api/file?path=" + encodeURIComponent(rel); }

function openFile(item, rel) {
  viewingRel = rel;
  document.getElementById("viewerTitle").textContent = item.name;
  const body = document.getElementById("viewerBody");
  body.innerHTML = '<p class="viewer-loading">Yukleniyor...</p>';
  document.getElementById("viewerModal").hidden = false;
  document.body.style.overflow = "hidden";
  location.hash = (state.path ? "/" + state.path : "/") + ":f=" + encodeURIComponent(item.name);

  const ext = (item.ext || "").toLowerCase();
  if ([".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"].includes(ext)) {
    body.innerHTML = `<img src="${fileUrl(rel)}" alt="${esc(item.name)}" />`;
  } else if (ext === ".pdf") {
    viewPdf(rel);
  } else if (ext === ".docx") {
    viewDocx(rel);
  } else {
    body.innerHTML = `<iframe src="${fileUrl(rel)}"></iframe>`;
  }
}

async function viewPdf(rel) {
  const body = document.getElementById("viewerBody");
  try {
    if (typeof pdfjsLib === "undefined") throw new Error("PDF kutuphanesi yuklenmedi");
    pdfjsLib.GlobalWorkerOptions.workerSrc = "/vendor/pdfjs/pdf.worker.min.js";
    const res = await fetch(fileUrl(rel));
    const buf = await res.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    body.innerHTML = '<div class="pdf-body"></div>';
    const wrap = body.querySelector(".pdf-body");
    const pages = await Promise.all(
      Array.from({ length: pdf.numPages }, (_, i) => pdf.getPage(i + 1))
    );
    for (const page of pages) {
      const vp = page.getViewport({ scale: 1.5 });
      const canvas = document.createElement("canvas");
      canvas.width = vp.width;
      canvas.height = vp.height;
      wrap.appendChild(canvas);
      const ctx = canvas.getContext("2d");
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
    }
  } catch (e) {
    body.innerHTML = '<div class="pdf-body"><p class="viewer-loading">PDF goruntulenemedi. <a class="btn btn-primary" href="' + fileUrl(rel) + '" download>Indir</a></p></div>';
  }
}

async function viewDocx(rel) {
  const body = document.getElementById("viewerBody");
  try {
    const res = await fetch(fileUrl(rel));
    const buf = await res.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer: buf });
    const isOtt = /osmanlica|osmanlıca/i.test(rel);
    body.innerHTML = '<div class="doc-body' + (isOtt ? " ott-mode" : "") + '">' + result.value + "</div>";
    beautifyArabic(body);
    body.querySelectorAll("p").forEach((p) => {
      const text = p.textContent.trim();
      if (/^(Meal|Meâl):/i.test(text) || /^(Mesaj|Not):/i.test(text)) p.classList.add("meal");
      if (/^[A-ZĞÜŞİÖÇ].*Bahri|BÖLÜM|BAHIR|DUASI|DASI|SURESİ|Ayet|Âyet/i.test(text) && text.length < 60) {
        p.classList.add("section");
      }
    });
  } catch (e) {
    body.innerHTML = '<p class="viewer-loading">Dokuman goruntulenemedi. Indirerek acabilirsiniz.</p>';
  }
}

const AR_RE = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/g;
const OTT_RE = /[پچژگڭڤۆۈۉۋۏەھۀۂ]/;

function beautifyArabic(container) {
  const blocks = container.querySelectorAll("p, li, div, h1, h2, h3, h4");
  blocks.forEach((el) => {
    if (el.classList.contains("ar")) return;
    const text = el.textContent || "";
    const letters = text.replace(/\s/g, "");
    if (!letters) return;
    const arCount = (letters.match(AR_RE) || []).length;
    if (arCount > 3 && arCount / letters.length > 0.3) {
      el.classList.add("ar");
      if (OTT_RE.test(text)) el.classList.add("ott");
    }
  });
}

function closeViewer() {
  document.getElementById("viewerModal").hidden = true;
  document.body.style.overflow = "";
  document.getElementById("viewerBody").innerHTML = "";
  if (location.hash.includes(":f=")) {
    const h = state.path ? "/" + state.path : "/";
    history.replaceState(null, "", "#" + h);
  }
}
function downloadCurrent() {
  if (viewingRel) window.open("/api/download?path=" + encodeURIComponent(viewingRel), "_blank");
}

/* ---------------- Init ---------------- */
function preprocessImage(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      const MAX = 4000;
      const scale = Math.min(1, MAX / Math.max(img.width, img.height));
      const w = Math.round(img.width * scale);
      const h = Math.round(img.height * scale);
      const cv = document.createElement("canvas");
      cv.width = w;
      cv.height = h;
      const ctx = cv.getContext("2d");
      ctx.drawImage(img, 0, 0, w, h);
      const data = ctx.getImageData(0, 0, w, h);
      const d = data.data;
      const contrast = 1.35;
      const factor = (259 * (contrast + 255)) / (255 * (259 - contrast));
      for (let i = 0; i < d.length; i += 4) {
        let v = factor * (d[i] - 128) + 128;
        v = v < 0 ? 0 : v > 255 ? 255 : v;
        const gray = 0.299 * v + 0.587 * d[i + 1] + 0.114 * d[i + 2];
        d[i] = gray;
        d[i + 1] = gray;
        d[i + 2] = gray;
      }
      ctx.putImageData(data, 0, 0);
      cv.toBlob((blob) => resolve(blob), "image/png");
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Gorsel okunamadi"));
    };
    img.src = url;
  });
}

/* ---------------- Scan page stack + crop ---------------- */
function renderScanStack() {
  const stack = document.getElementById("scanStack");
  stack.innerHTML = "";
  if (!state.scanFiles.length) { stack.hidden = true; return; }
  stack.hidden = false;
  state.scanFiles.forEach((f, i) => {
    const page = document.createElement("div");
    page.className = "scan-page";
    const head = document.createElement("div");
    head.className = "scan-page-head";
    const lbl = document.createElement("span");
    lbl.textContent = (i + 1) + ". sayfa";
    const btns = document.createElement("span");
    const cropB = document.createElement("button");
    cropB.className = "btn btn-ghost";
    cropB.textContent = "✂ Kırp";
    cropB.onclick = (ev) => { ev.stopPropagation(); openCrop(i); };
    const delB = document.createElement("button");
    delB.className = "btn btn-ghost";
    delB.textContent = "🗑 Sil";
    delB.onclick = (ev) => { ev.stopPropagation(); removeScanPage(i); };
    btns.appendChild(cropB);
    btns.appendChild(delB);
    head.appendChild(lbl);
    head.appendChild(btns);
    const img = document.createElement("img");
    img.src = URL.createObjectURL(f);
    page.appendChild(head);
    page.appendChild(img);
    stack.appendChild(page);
  });
}

function removeScanPage(i) {
  if (i === state.scanFiles.length - 1) {
    state.scanFile = null;
    state.scanOcrFile = null;
    state.scanTextReady = false;
  }
  state.scanFiles.splice(i, 1);
  renderScanStack();
}

let cropState = { index: -1, scaleX: 1, scaleY: 1 };

function openCrop(i) {
  cropState.index = i;
  const f = state.scanFiles[i];
  if (!f) return;
  const img = document.getElementById("cropImage");
  const rect = document.getElementById("cropRect");
  img.onload = () => {
    const stage = document.getElementById("cropStage");
    const sw = stage.clientWidth;
    const dispH = img.clientHeight || (sw * img.naturalHeight / img.naturalWidth);
    cropState.scaleX = img.naturalWidth / sw;
    cropState.scaleY = img.naturalHeight / dispH;
    const m = 24;
    rect.style.left = m + "px";
    rect.style.top = m + "px";
    rect.style.width = Math.max(40, sw - m * 2) + "px";
    rect.style.height = Math.max(40, dispH - m * 2) + "px";
    rect.style.display = "block";
  };
  img.src = URL.createObjectURL(f);
  document.getElementById("cropModal").hidden = false;
}

function closeCrop() {
  document.getElementById("cropModal").hidden = true;
  document.getElementById("cropRect").style.display = "none";
  document.getElementById("cropImage").removeAttribute("src");
  cropState.index = -1;
}

function confirmCrop() {
  const i = cropState.index;
  if (i < 0) return;
  const rect = document.getElementById("cropRect");
  const img = document.getElementById("cropImage");
  const x = parseFloat(rect.style.left) * cropState.scaleX;
  const y = parseFloat(rect.style.top) * cropState.scaleY;
  const w = parseFloat(rect.style.width) * cropState.scaleX;
  const h = parseFloat(rect.style.height) * cropState.scaleY;
  const cv = document.createElement("canvas");
  cv.width = Math.round(w);
  cv.height = Math.round(h);
  cv.getContext("2d").drawImage(img, x, y, w, h, 0, 0, cv.width, cv.height);
  cv.toBlob(async (blob) => {
    blob.name = (state.scanFiles[i] && state.scanFiles[i].name) || "crop.png";
    state.scanFiles[i] = blob;
    if (i === state.scanFiles.length - 1) {
      state.scanFile = blob;
      state.scanOcrFile = blob;
    }
    closeCrop();
    renderScanStack();
    toast("Gorsel kirpildi");
  }, "image/png");
}

let cropDrag = null;

function cropPos(e) {
  const r = document.getElementById("cropStage").getBoundingClientRect();
  return { x: e.clientX - r.left, y: e.clientY - r.top };
}

document.getElementById("cropRect").addEventListener("pointerdown", (e) => {
  e.preventDefault();
  const rect = document.getElementById("cropRect");
  const rr = rect.getBoundingClientRect();
  const p = { x: e.clientX - rr.left, y: e.clientY - rr.top };
  const d = 22;
  const near = (v, t) => Math.abs(v - t) <= d;
  let mode = "move";
  if (near(p.x, 0) && near(p.y, 0)) mode = "nw";
  else if (near(p.x, rr.width) && near(p.y, 0)) mode = "ne";
  else if (near(p.x, 0) && near(p.y, rr.height)) mode = "sw";
  else if (near(p.x, rr.width) && near(p.y, rr.height)) mode = "se";
  cropDrag = {
    mode,
    startX: p.x,
    startY: p.y,
    startL: parseFloat(rect.style.left),
    startT: parseFloat(rect.style.top),
    startW: parseFloat(rect.style.width),
    startH: parseFloat(rect.style.height),
  };
  document.getElementById("cropStage").setPointerCapture(e.pointerId);
});

document.getElementById("cropStage").addEventListener("pointermove", (e) => {
  if (!cropDrag) return;
  e.preventDefault();
  const p = cropPos(e);
  const dd = cropDrag;
  const stage = document.getElementById("cropStage");
  const rect = document.getElementById("cropRect");
  const maxW = stage.clientWidth;
  const maxH = stage.clientHeight;
  const min = 40;
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const dx = p.x - dd.startX;
  const dy = p.y - dd.startY;
  let l = dd.startL, t = dd.startT, w = dd.startW, h = dd.startH;
  if (dd.mode === "move") {
    l = clamp(dd.startL + dx, 0, maxW - w);
    t = clamp(dd.startT + dy, 0, maxH - h);
  } else if (dd.mode === "se") {
    w = clamp(dd.startW + dx, min, maxW - dd.startL);
    h = clamp(dd.startH + dy, min, maxH - dd.startT);
  } else if (dd.mode === "nw") {
    l = clamp(dd.startL + dx, 0, dd.startL + dd.startW - min);
    t = clamp(dd.startT + dy, 0, dd.startT + dd.startH - min);
    w = dd.startL + dd.startW - l;
    h = dd.startT + dd.startH - t;
  } else if (dd.mode === "ne") {
    w = clamp(dd.startW + dx, min, maxW - dd.startL);
    t = clamp(dd.startT + dy, 0, dd.startT + dd.startH - min);
    h = dd.startT + dd.startH - t;
  } else if (dd.mode === "sw") {
    l = clamp(dd.startL + dx, 0, dd.startL + dd.startW - min);
    w = dd.startL + dd.startW - l;
    h = clamp(dd.startH + dy, min, maxH - dd.startT);
  }
  rect.style.left = l + "px";
  rect.style.top = t + "px";
  rect.style.width = w + "px";
  rect.style.height = h + "px";
});

function cropEnd() { cropDrag = null; }
document.getElementById("cropStage").addEventListener("pointerup", cropEnd);
document.getElementById("cropStage").addEventListener("pointercancel", cropEnd);

document.getElementById("scanInput").addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  state.scanFile = f;
  state.scanOcrFile = null;
  state.scanTextReady = false;
  state.scanFiles.push(f);
  renderScanStack();
  const prog = document.getElementById("scanProgress");
  prog.hidden = false;
  prog.textContent = "Gorsel isleniyor (" + state.scanFiles.length + ". sayfa)...";
  try {
    state.scanOcrFile = await preprocessImage(f);
    prog.hidden = true;
  } catch (err) {
    prog.textContent = "Gorsel islenemedi, orijinal kullanilacak.";
    setTimeout(() => { prog.hidden = true; }, 2500);
  }
});
const initialHash = decodeURIComponent(location.hash.replace(/^#/, ""));
const initialPath = initialHash.split(":f=")[0].replace(/^\//, "");
state.path = initialPath;
if (!location.hash) history.replaceState(null, "", "#/");
document.querySelectorAll('input[name="sfmt"]').forEach((r) => r.addEventListener("change", toggleScanFormat));
document.getElementById("viewerModal").addEventListener("click", (e) => {
  if (e.target.id === "viewerModal") closeViewer();
});
document.getElementById("uploadModal").addEventListener("click", (e) => {
  if (e.target.id === "uploadModal") closeUpload();
});
document.getElementById("scanModal").addEventListener("click", (e) => {
  if (e.target.id === "scanModal") closeScan();
});
document.getElementById("cropModal").addEventListener("click", (e) => {
  if (e.target.id === "cropModal") closeCrop();
});
document.getElementById("searchModal").addEventListener("click", (e) => {
  if (e.target.id === "searchModal") closeSearch();
});
document.getElementById("moveModal").addEventListener("click", (e) => {
  if (e.target.id === "moveModal") closeMoveModal();
});
document.getElementById("newTextModal").addEventListener("click", (e) => {
  if (e.target.id === "newTextModal") closeNewText();
});
document.getElementById("renameModal").addEventListener("click", (e) => {
  if (e.target.id === "renameModal") closeRename();
});
document.getElementById("editModal").addEventListener("click", (e) => {
  if (e.target.id === "editModal") closeEdit();
});
document.getElementById("searchInput").addEventListener("input", onSearchInput);
document.getElementById("searchInput").addEventListener("keydown", (e) => {
  if (e.key === "Enter") { e.preventDefault(); runSearch(); }
});
document.addEventListener("click", (e) => {
  if (activeMenu && !activeMenu.contains(e.target)) closeItemMenu();
  if (!document.getElementById("catMenu").hidden && !e.target.closest(".topbar") && !e.target.closest("#catMenu")) closeCatMenu();
  if (!document.getElementById("uploadMenu").hidden && !e.target.closest(".upload-wrap")) closeUploadMenu();
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") { closeViewer(); closeUpload(); closeScan(); closeCrop(); closeItemMenu(); closeSearch(); closeMoveModal(); closeNewText(); closeRename(); closeEdit(); closeCatMenu(); closeUploadMenu(); }
});

load();
