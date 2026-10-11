/* Document Vault — scan, OCR-search and file documents in your own OneDrive.
   Storage: OneDrive app folder (Apps/Document Vault). index.json holds all metadata + searchable text.
   Text: PDF text layer first, then on-device OCR (Tesseract), then optional AI fallback (Claude API key). */
(() => {
"use strict";

/* ================= utilities ================= */
const $ = (s, el = document) => el.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const norm = s => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
const abs = p => new URL(p, location.href).href;
const sleep = ms => new Promise(r => setTimeout(r, ms));
const lsGet = (k, d = null) => { try { const v = localStorage.getItem(k); return v === null ? d : JSON.parse(v); } catch { return d; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} };
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2)).replace(/-/g, "").slice(0, 20);

const ICON = {
  back:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 18l-6-6 6-6"/></svg>',
  rotate:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 11a8 8 0 1 0-2.3 5.7"/><path d="M20 4v7h-7"/></svg>',
  crop:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 2v14a2 2 0 0 0 2 2h14"/><path d="M2 6h14a2 2 0 0 1 2 2v14"/></svg>',
  trash:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14"/></svg>',
  camera:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/></svg>',
  image:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="m21 16-5-5-9 9"/></svg>',
  spark:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z"/></svg>',
  text:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7V5h16v2M9 19h6M12 5v14"/></svg>',
  open:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 4h6v6M20 4l-9 9"/><path d="M19 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1h5"/></svg>',
  x:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M6 6l12 12M18 6 6 18"/></svg>',
  check:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12l5 5 9-10"/></svg>',
  cal:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/></svg>',
  share:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><path d="m8.6 13.5 6.8 4M15.4 6.5l-6.8 4"/></svg>',
  download:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v12M7 11l5 5 5-5"/><path d="M4 20h16"/></svg>',
  inbox:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 13h5l1 3h6l1-3h5"/><path d="M5 5h14l2 8v6H3v-6z"/></svg>'
};

const DEFAULT_CATS = {
  private: ["Insurance","Taxes","Bank & Finance","Housing","Utilities & Energy","Car & Bike","Health","Identity & Legal","Purchases & Warranties","Work & Income","Other"],
  business: ["Invoices received","Invoices sent","Receipts & Expenses","Contracts & Agreements","Tax & Accounting","Banking","Payroll & HR","Shipping & Customs","Permits & Licences","Quotes & Orders","Other"]
};
const SPACE_LABEL = {private: "Private", business: "Business"};
const CURRENCIES = ["EUR","USD","GBP","ETB","KES","TZS"];
const INBOX = "Inbox";
const SCOPES = ["Files.ReadWrite"];
const G = "https://graph.microsoft.com/v1.0";
const BASE = ((window.DV_CONFIG && window.DV_CONFIG.folder) || "Document Vault").replace(/^\/+|\/+$/g, "");

let toastTimer;
function toast(msg, ms = 3400) {
  const t = $("#toast"); t.textContent = msg; t.hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.hidden = true, ms);
}
function fmtDate(iso) {
  if (!iso) return "";
  const d = new Date(iso.length === 10 ? iso + "T12:00:00" : iso);
  return isNaN(d) ? iso : d.toLocaleDateString("en-GB", {day: "numeric", month: "short", year: "numeric"});
}
function fmtAmount(a, cur) {
  if (a === null || a === undefined || a === "" || isNaN(Number(a))) return "";
  try { return new Intl.NumberFormat("nl-NL", {style: "currency", currency: cur || "EUR"}).format(Number(a)); }
  catch { return (cur ? cur + " " : "") + Number(a).toFixed(2); }
}
function daysUntil(iso) {
  if (!iso) return null;
  const d = new Date(iso + "T00:00:00"), n = new Date(); n.setHours(0, 0, 0, 0);
  return isNaN(d) ? null : Math.round((d - n) / 86400000);
}
function dueInfo(doc) {
  if (doc.handled || !doc.dueDate) return null;
  const n = daysUntil(doc.dueDate);
  if (n === null || n > 14) return null;
  if (n < 0) return {n, label: `Overdue ${-n} d`};
  if (n === 0) return {n, label: "Due today"};
  return {n, label: `Due in ${n} d`};
}
const docDate = d => d.date || (d.createdAt ? new Date(d.createdAt).toISOString().slice(0, 10) : "");
function safeSeg(s, max = 90) {
  return String(s || "").replace(/[\\/:*?"<>|#%\u0000-\u001f]+/g, " ").replace(/\s+/g, " ").trim().replace(/[. ]+$/, "").slice(0, max).trim() || "Untitled";
}
const encPath = p => p.split("/").map(encodeURIComponent).join("/");
function loadScript(src) {
  return new Promise((res, rej) => {
    const ex = [...document.scripts].find(s => s.dataset.src === src);
    if (ex) { if (ex.dataset.ok) return res(); ex.addEventListener("load", () => res()); ex.addEventListener("error", rej); return; }
    const s = document.createElement("script"); s.src = src; s.dataset.src = src;
    s.onload = () => { s.dataset.ok = "1"; res(); }; s.onerror = () => rej(new Error("Could not load " + src));
    document.head.appendChild(s);
  });
}
const loadImg = src => new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; });
const toBlob = (c, q = .82) => new Promise(r => c.toBlob(r, "image/jpeg", q));
async function sha256(buf) {
  try { const h = await crypto.subtle.digest("SHA-256", buf); return [...new Uint8Array(h)].map(b => b.toString(16).padStart(2, "0")).join(""); }
  catch { return ""; }
}
function blobToB64(blob) {
  return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(",")[1]); r.onerror = rej; r.readAsDataURL(blob); });
}

/* ================= tiny IndexedDB key-value store ================= */
const idb = (() => {
  let dbp = null;
  const open = () => dbp || (dbp = new Promise((res, rej) => {
    const r = indexedDB.open("document-vault", 1);
    r.onupgradeneeded = () => { r.result.createObjectStore("kv"); r.result.createObjectStore("thumbs"); };
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  }));
  const tx = async (store, mode, fn) => {
    try {
      const db = await open();
      return await new Promise((res, rej) => { const t = db.transaction(store, mode); const q = fn(t.objectStore(store)); t.oncomplete = () => res(q && q.result); t.onerror = () => rej(t.error); });
    } catch { return undefined; }
  };
  return {
    get: (store, k) => tx(store, "readonly", s => s.get(k)),
    set: (store, k, v) => tx(store, "readwrite", s => s.put(v, k)),
    del: (store, k) => tx(store, "readwrite", s => s.delete(k)),
    clear: store => tx(store, "readwrite", s => s.clear())
  };
})();

/* ================= state ================= */
const state = {
  index: emptyIndex(), etag: null, loaded: false,
  space: "all", cat: null, q: "", year: "", amt: "", selecting: false, sel: new Set(), online: navigator.onLine,
  syncing: false, dirty: false, syncError: ""
};
function emptyIndex() { return {app: "document-vault", version: 1, docs: {}, categories: structuredClone(DEFAULT_CATS), catsUpdatedAt: 0, updatedAt: 0}; }
const docs = () => Object.values(state.index.docs).filter(d => !d.deleted);
function catsFor(space) { const c = state.index.categories?.[space]; return c && c.length ? c : DEFAULT_CATS[space]; }
const getDoc = id => { const d = state.index.docs[id]; return d && !d.deleted ? d : null; };

/* ================= auth (MSAL) ================= */
let msalApp = null, account = null;
const clientId = () => (window.DV_CONFIG && window.DV_CONFIG.clientId) || lsGet("dv.clientId", "");
const redirectUri = () => location.origin + location.pathname.replace(/index\.html$/, "");

async function initAuth() {
  msalApp = new msal.PublicClientApplication({
    auth: {clientId: clientId(), authority: "https://login.microsoftonline.com/consumers", redirectUri: redirectUri(), navigateToLoginRequestUrl: false},
    cache: {cacheLocation: "localStorage"}
  });
  await msalApp.initialize();
  const res = await msalApp.handleRedirectPromise();
  account = (res && res.account) || msalApp.getAllAccounts()[0] || null;
  return account;
}
async function getToken() {
  try {
    const r = await msalApp.acquireTokenSilent({scopes: SCOPES, account});
    return r.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError || /interaction_required|login_required|consent_required|no_tokens_found/i.test(String(e && e.errorCode))) {
      await msalApp.acquireTokenRedirect({scopes: SCOPES, account});
      await new Promise(() => {}); // page navigates away
    }
    throw e;
  }
}

/* ================= Microsoft Graph ================= */
class GraphError extends Error { constructor(status, msg, code) { super(msg); this.status = status; this.code = code; } }
async function g(method, path, {body, headers = {}, raw = false} = {}) {
  const url = path.startsWith("http") ? path : G + path;
  for (let attempt = 0; attempt < 4; attempt++) {
    const token = await getToken();
    const h = {Authorization: "Bearer " + token, ...headers};
    let b = body;
    if (body && !(body instanceof Blob) && !(body instanceof ArrayBuffer) && typeof body !== "string") { b = JSON.stringify(body); h["Content-Type"] = "application/json"; }
    let r;
    try { r = await fetch(url, {method, headers: h, body: b}); }
    catch { throw new GraphError(0, "You seem to be offline.", "offline"); }
    if (r.status === 429 || r.status === 503 || r.status === 504) { await sleep((Number(r.headers.get("Retry-After")) || (2 ** attempt)) * 1000); continue; }
    if (raw) return r;
    if (r.status === 204) return null;
    if (!r.ok) {
      let msg = r.statusText, code = "";
      try { const j = await r.json(); msg = j.error?.message || msg; code = j.error?.code || ""; } catch {}
      throw new GraphError(r.status, msg, code);
    }
    const ct = r.headers.get("content-type") || "";
    return ct.includes("json") ? r.json() : r.blob();
  }
  throw new GraphError(429, "OneDrive is busy. Try again in a minute.", "throttled");
}
const full = p => p ? `${BASE}/${p}` : BASE;
const itemRef = p => `/me/drive/root:/${encPath(full(p))}:`;

async function uploadFile(folder, name, blob, {replace = false, ifMatch = null} = {}) {
  const p = folder ? `${folder}/${name}` : name;
  const cb = replace ? "replace" : "rename";
  if (blob.size <= 3.5 * 1024 * 1024) {
    const headers = {"Content-Type": blob.type || "application/octet-stream"};
    if (ifMatch) headers["If-Match"] = ifMatch;
    return g("PUT", `${itemRef(p)}/content?@microsoft.graph.conflictBehavior=${ifMatch ? "replace" : cb}`, {body: blob, headers});
  }
  const headers = ifMatch ? {"If-Match": ifMatch} : {};
  const s = await g("POST", `${itemRef(p)}/createUploadSession`, {body: {item: {"@microsoft.graph.conflictBehavior": ifMatch ? "replace" : cb}}, headers});
  const CH = 320 * 1024 * 10; // 3.2 MB, a multiple of 320 KiB
  let res = null;
  for (let start = 0; start < blob.size; start += CH) {
    const end = Math.min(blob.size, start + CH);
    const r = await fetch(s.uploadUrl, {method: "PUT", headers: {"Content-Range": `bytes ${start}-${end - 1}/${blob.size}`}, body: blob.slice(start, end)});
    if (!r.ok && r.status !== 202) throw new GraphError(r.status, "Upload failed.", "upload");
    if (r.status === 200 || r.status === 201) res = await r.json();
  }
  return res;
}
async function replaceContent(itemId, blob) {
  if (blob.size <= 3.5 * 1024 * 1024)
    return g("PUT", `/me/drive/items/${itemId}/content`, {body: blob, headers: {"Content-Type": blob.type || "application/octet-stream"}});
  const s = await g("POST", `/me/drive/items/${itemId}/createUploadSession`, {body: {item: {"@microsoft.graph.conflictBehavior": "replace"}}});
  const CH = 320 * 1024 * 10; let res = null;
  for (let start = 0; start < blob.size; start += CH) {
    const end = Math.min(blob.size, start + CH);
    const r = await fetch(s.uploadUrl, {method: "PUT", headers: {"Content-Range": `bytes ${start}-${end - 1}/${blob.size}`}, body: blob.slice(start, end)});
    if (!r.ok && r.status !== 202) throw new GraphError(r.status, "Upload failed.", "upload");
    if (r.status === 200 || r.status === 201) res = await r.json();
  }
  return res;
}
const folderCache = new Map();
async function ensureFolder(path) {
  if (folderCache.has(path)) return folderCache.get(path);
  try { const it = await g("GET", itemRef(path)); folderCache.set(path, it.id); return it.id; }
  catch (e) { if (e.status !== 404) throw e; }
  if (!path) {
    try { const it = await g("POST", "/me/drive/root/children", {body: {name: BASE, folder: {}, "@microsoft.graph.conflictBehavior": "fail"}}); folderCache.set("", it.id); return it.id; }
    catch (e) { if (e.status !== 409) throw e; const it = await g("GET", itemRef("")); folderCache.set("", it.id); return it.id; }
  }
  const parts = path.split("/"), name = parts.pop(), parent = parts.join("/");
  const parentPath = `/me/drive/items/${await ensureFolder(parent)}/children`;
  try {
    const it = await g("POST", parentPath, {body: {name, folder: {}, "@microsoft.graph.conflictBehavior": "fail"}});
    folderCache.set(path, it.id); return it.id;
  } catch (e) {
    if (e.status !== 409) throw e;
    const it = await g("GET", itemRef(path)); folderCache.set(path, it.id); return it.id;
  }
}
async function moveItem(itemId, folder, name) {
  const parentId = await ensureFolder(folder);
  return g("PATCH", `/me/drive/items/${itemId}?@microsoft.graph.conflictBehavior=rename`, {body: {name, parentReference: {id: parentId}}});
}
async function downloadItem(itemId) {
  const it = await g("GET", `/me/drive/items/${itemId}`);
  const url = it["@microsoft.graph.downloadUrl"];
  if (!url) throw new GraphError(404, "File not found in OneDrive.", "notfound");
  const r = await fetch(url);
  if (!r.ok) throw new GraphError(r.status, "Download failed.", "download");
  return r.blob();
}

/* ================= index sync ================= */
const INDEX = "index.json";
function mergeIndex(a, b) {
  const out = emptyIndex();
  out.docs = {...a.docs};
  for (const [id, d] of Object.entries(b.docs || {})) {
    const x = out.docs[id];
    if (!x || (d.updatedAt || 0) > (x.updatedAt || 0)) out.docs[id] = d;
  }
  const ca = a.catsUpdatedAt || 0, cb = b.catsUpdatedAt || 0;
  out.categories = cb > ca ? b.categories : a.categories;
  out.catsUpdatedAt = Math.max(ca, cb);
  out.updatedAt = Math.max(a.updatedAt || 0, b.updatedAt || 0);
  if (!out.categories || !out.categories.private) out.categories = structuredClone(DEFAULT_CATS);
  return out;
}
async function fetchRemoteIndex() {
  let meta;
  try { meta = await g("GET", itemRef(INDEX)); }
  catch (e) { if (e.status === 404) return {index: null, etag: null}; throw e; }
  const r = await fetch(meta["@microsoft.graph.downloadUrl"], {cache: "no-store"});
  if (!r.ok) throw new GraphError(r.status, "Could not read the index.", "index");
  const idx = await r.json();
  return {index: idx && idx.docs ? idx : emptyIndex(), etag: meta.eTag};
}
async function cacheLocal() { await idb.set("kv", "index", {index: state.index, etag: state.etag, dirty: state.dirty}); }
async function loadIndex() {
  const cached = await idb.get("kv", "index");
  if (cached && cached.index) { state.index = mergeIndex(emptyIndex(), cached.index); state.etag = cached.etag; state.dirty = !!cached.dirty; state.loaded = true; render(); }
  await refreshIndex();
}
async function refreshIndex() {
  if (state.syncing) return;
  try {
    setSync("Syncing…");
    const {index, etag} = await fetchRemoteIndex();
    if (index && etag !== state.etag) state.index = mergeIndex(state.index, index);
    if (index) state.etag = etag;
    state.loaded = true; state.syncError = "";
    await cacheLocal(); render();
    if (state.dirty || !index) scheduleSave(0); else setSync("Synced");
  } catch (e) {
    state.loaded = true; render();
    syncFailed(e);
  }
}
let saveTimer = null, saving = null;
function scheduleSave(ms = 700) {
  state.dirty = true; cacheLocal(); setSync("Saving…");
  clearTimeout(saveTimer); saveTimer = setTimeout(() => { saving = (saving || Promise.resolve()).then(saveIndex); }, ms);
}
async function saveIndex() {
  if (!state.dirty) return;
  state.syncing = true;
  try {
    for (let attempt = 0; attempt < 4; attempt++) {
      state.index.updatedAt = Date.now();
      const blob = new Blob([JSON.stringify(state.index)], {type: "application/json"});
      try {
        const it = await uploadFile("", INDEX, blob, {replace: !state.etag ? false : true, ifMatch: state.etag});
        if (!state.etag && it && it.name !== INDEX) { // created a renamed copy: someone else made index.json first
          await g("DELETE", `/me/drive/items/${it.id}`).catch(() => {});
          throw new GraphError(409, "conflict", "conflict");
        }
        state.etag = it.eTag; state.dirty = false; state.syncError = "";
        await cacheLocal(); setSync("Synced");
        writeCsv();
        return;
      } catch (e) {
        if (e.status === 412 || e.status === 409) {
          const {index, etag} = await fetchRemoteIndex();
          if (index) { state.index = mergeIndex(state.index, index); state.etag = etag; render(); }
          continue;
        }
        throw e;
      }
    }
    throw new GraphError(409, "Could not save after several tries.", "conflict");
  } catch (e) { syncFailed(e); }
  finally { state.syncing = false; }
}
function syncFailed(e) {
  console.error(e);
  state.syncError = e && e.code === "offline" ? "offline" : "error";
  setSync(state.syncError === "offline" ? "Offline" : "Sync error", true);
}
function setSync(t, err = false) { const s = $("#sync"); if (s) { s.textContent = t; s.classList.toggle("err", err); } }
let csvTimer;
function writeCsv() {
  clearTimeout(csvTimer);
  csvTimer = setTimeout(async () => {
    const cell = v => `"${String(v ?? "").replace(/"/g, '""')}"`;
    const rows = [["Space","Category","Date","Title","From","Reference","Amount","Currency","Due or expires","Tags","Notes","OneDrive path"].map(cell).join(";")];
    for (const d of docs().sort((a, b) => docDate(b).localeCompare(docDate(a))))
      rows.push([d.filed ? SPACE_LABEL[d.space] : INBOX, d.filed ? d.category : "", d.date, d.title, d.sender, d.reference, d.amount ?? "", d.currency, d.dueDate, (d.tags || []).join(", "), d.notes, d.path].map(cell).join(";"));
    try { await uploadFile("", "index.csv", new Blob(["﻿" + rows.join("\r\n")], {type: "text/csv"}), {replace: true}); } catch {}
  }, 4000);
}
function putDoc(d) { d.updatedAt = Date.now(); state.index.docs[d.id] = d; render(); scheduleSave(); }

/* ================= image processing ================= */
function applyEnhance(ctx, w, h, mode) {
  const img = ctx.getImageData(0, 0, w, h), d = img.data, hist = new Uint32Array(256);
  const step = Math.max(1, Math.floor((w * h) / 200000)); let count = 0;
  for (let i = 0; i < d.length; i += 4 * step) { hist[(d[i] * .299 + d[i + 1] * .587 + d[i + 2] * .114) | 0]++; count++; }
  const pct = p => { let acc = 0; const t = count * p; for (let v = 0; v < 256; v++) { acc += hist[v]; if (acc >= t) return v; } return 255; };
  // paper -> white, ink -> black; never stretch a nearly uniform page
  const hi = pct(mode === "bw" ? .80 : .97), lo = Math.min(pct(mode === "bw" ? .04 : .01), 110, hi - 60);
  if (hi - lo < 60) return;
  const lut = new Uint8ClampedArray(256);
  for (let v = 0; v < 256; v++) { const x = Math.min(1, Math.max(0, (v - lo) / (hi - lo))); lut[v] = Math.round(255 * Math.pow(x, mode === "bw" ? 1.15 : 1)); }
  if (mode === "bw") for (let i = 0; i < d.length; i += 4) { const gg = lut[(d[i] * .299 + d[i + 1] * .587 + d[i + 2] * .114) | 0]; d[i] = d[i + 1] = d[i + 2] = gg; }
  else for (let i = 0; i < d.length; i += 4) { d[i] = lut[d[i]]; d[i + 1] = lut[d[i + 1]]; d[i + 2] = lut[d[i + 2]]; }
  ctx.putImageData(img, 0, 0);
}
function renderPage(p, mode, {maxSide = 2200, crop = true, enhance = true} = {}) {
  const bm = p.bitmap, w = bm.width, h = bm.height;
  const rot = ((p.rot % 360) + 360) % 360, swap = rot === 90 || rot === 270;
  const rw = swap ? h : w, rh = swap ? w : h;
  let cx = 0, cy = 0, cw = rw, ch = rh;
  if (crop && p.crop) { cx = p.crop.x * rw; cy = p.crop.y * rh; cw = p.crop.w * rw; ch = p.crop.h * rh; }
  const scale = Math.min(1, maxSide / Math.max(cw, ch));
  const c = document.createElement("canvas");
  c.width = Math.max(1, Math.round(cw * scale)); c.height = Math.max(1, Math.round(ch * scale));
  const ctx = c.getContext("2d", {willReadFrequently: true});
  ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, c.width, c.height);
  ctx.save(); ctx.scale(scale, scale); ctx.translate(-cx, -cy); ctx.translate(rw / 2, rh / 2);
  ctx.rotate(rot * Math.PI / 180); ctx.drawImage(bm, -w / 2, -h / 2); ctx.restore();
  if (enhance && mode !== "original") applyEnhance(ctx, c.width, c.height, mode);
  return c;
}
function scaled(canvas, maxSide) {
  const s = Math.min(1, maxSide / Math.max(canvas.width, canvas.height));
  if (s === 1) return canvas;
  const t = document.createElement("canvas"); t.width = Math.round(canvas.width * s); t.height = Math.round(canvas.height * s);
  t.getContext("2d").drawImage(canvas, 0, 0, t.width, t.height); return t;
}

/* ================= PDF ================= */
async function pdfLib() {
  await loadScript("vendor/pdf.min.js");
  window.pdfjsLib.GlobalWorkerOptions.workerSrc = abs("vendor/pdf.worker.min.js");
  return window.pdfjsLib;
}
async function openPdf(buf) {
  const lib = await pdfLib();
  try { return await lib.getDocument({data: new Uint8Array(buf), isEvalSupported: false}).promise; }
  catch (e) { throw {userMsg: e && e.name === "PasswordException" ? "This PDF is password protected." : "This PDF could not be opened."}; }
}
function textFromContent(tc) {
  let out = "", lastY = null;
  for (const it of tc.items) {
    if (typeof it.str !== "string") continue;
    const y = it.transform ? it.transform[5] : null;
    if (lastY !== null && y !== null && Math.abs(y - lastY) > 2) out += "\n";
    else if (out && !/\s$/.test(out) && it.str && !/^\s/.test(it.str)) out += " ";
    out += it.str;
    if (it.hasEOL) { out += "\n"; lastY = null; } else lastY = y;
  }
  return out.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}
async function pdfText(pdf) {
  const parts = [];
  for (let i = 1; i <= Math.min(pdf.numPages, 300); i++) {
    const page = await pdf.getPage(i);
    try { parts.push(textFromContent(await page.getTextContent())); } catch {}
    page.cleanup();
  }
  return parts.join("\n\n").trim();
}
async function pdfPageCanvas(pdf, i, maxSide = 2000) {
  const page = await pdf.getPage(i);
  const v0 = page.getViewport({scale: 1}), sc = Math.min(3, maxSide / Math.max(v0.width, v0.height)), v = page.getViewport({scale: sc});
  const c = document.createElement("canvas"); c.width = Math.ceil(v.width); c.height = Math.ceil(v.height);
  const ctx = c.getContext("2d"); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, c.width, c.height);
  await page.render({canvasContext: ctx, viewport: v}).promise; page.cleanup();
  return c;
}
async function buildScanPdf(pageBlobs, pageTexts, meta) {
  await loadScript("vendor/jspdf.umd.min.js");
  const {jsPDF} = window.jspdf; let pdf = null;
  for (let i = 0; i < pageBlobs.length; i++) {
    const bytes = new Uint8Array(await pageBlobs[i].arrayBuffer());
    const bm = await createImageBitmap(pageBlobs[i]);
    const w = bm.width, h = bm.height, o = w > h ? "l" : "p"; bm.close && bm.close();
    if (!pdf) pdf = new jsPDF({orientation: o, unit: "mm", format: "a4", compress: true}); else pdf.addPage("a4", o);
    const pw = pdf.internal.pageSize.getWidth(), ph = pdf.internal.pageSize.getHeight(), m = 6, s = Math.min((pw - 2 * m) / w, (ph - 2 * m) / h);
    pdf.addImage(bytes, "JPEG", (pw - w * s) / 2, (ph - h * s) / 2, w * s, h * s, undefined, "NONE");
    const t = (pageTexts[i] || "").replace(/[^\x09\x0A\x0D\x20-\x7E -ÿ€]/g, " ").trim();
    if (t) { try { pdf.setFontSize(4); const lines = pdf.splitTextToSize(t, pw - 10); pdf.text(lines.slice(0, Math.floor((ph - 10) / 1.6)), 5, 6, {renderingMode: "invisible", lineHeightFactor: 1.1}); } catch {} }
  }
  pdf.setProperties({title: meta.title || "Document", keywords: (meta.tags || []).join(", "), creator: "Document Vault"});
  return pdf.output("blob");
}

/* ================= OCR (Tesseract, on this device) ================= */
let ocrWorker = null, ocrIdle = null, ocrProgress = null;
async function getOcr() {
  clearTimeout(ocrIdle);
  if (!ocrWorker) {
    await loadScript("vendor/tesseract.min.js");
    ocrWorker = Tesseract.createWorker(["nld", "eng"], 1, {
      workerPath: abs("vendor/tesseract-worker.min.js"),
      corePath: abs("vendor/tesseract-core/"),
      langPath: abs("vendor/tessdata"),
      logger: m => { if (ocrProgress && m && m.status) ocrProgress(m); }
    }).catch(e => { ocrWorker = null; throw e; });
  }
  return ocrWorker;
}
function releaseOcr() {
  clearTimeout(ocrIdle);
  ocrIdle = setTimeout(async () => { const w = ocrWorker; ocrWorker = null; try { (await w)?.terminate(); } catch {} }, 90000);
}
async function ocrCanvas(canvas) {
  const w = await getOcr();
  try {
    const {data} = await w.recognize(scaled(canvas, 2400));
    return {text: (data.text || "").trim(), conf: Math.round(data.confidence || 0)};
  } finally { releaseOcr(); }
}
const isWeak = (text, conf) => text.replace(/\s/g, "").length < 40 || (conf !== null && conf !== undefined && conf < 70);

/* ================= AI fallback (optional, your own API key) ================= */
const ai = {
  get key() { return lsGet("dv.aiKey", ""); },
  get model() { return lsGet("dv.aiModel", "") || "claude-haiku-5-5"; },
  get auto() { return lsGet("dv.aiAuto", true); },
  get on() { return !!this.key; }
};
function aiPrompt(n) {
  return `You are reading ${n} scanned page image(s) of ONE document (invoice, receipt, letter, contract, warranty or similar) for a personal archive. Read every page carefully, including small and faded print.
Private categories: ${catsFor("private").join(" | ")}
Business categories (his horticulture companies, suppliers, customers, staff, shipping, company finance): ${catsFor("business").join(" | ")}
Reply with only one JSON object with these fields:
"text": faithful transcription of all readable text in reading order, original language, with line breaks, at most 20000 characters,
"title": short descriptive title, max 60 characters,
"space": "private" or "business",
"category": one category of that space, copied exactly,
"date": document date "YYYY-MM-DD" or "",
"sender": issuing company or person, or "",
"reference": invoice, receipt, policy or order number, or "",
"amount": total amount as a number, or null,
"currency": ISO code like "EUR", or "",
"dueDate": payment deadline or warranty end "YYYY-MM-DD", or "".`;
}
async function aiRead(blobs) {
  const imgs = [];
  for (const b of blobs.slice(0, 5)) {
    const url = URL.createObjectURL(b); const img = await loadImg(url); URL.revokeObjectURL(url);
    const c = document.createElement("canvas"), s = Math.min(1, 1600 / Math.max(img.naturalWidth, img.naturalHeight));
    c.width = Math.round(img.naturalWidth * s); c.height = Math.round(img.naturalHeight * s);
    c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
    imgs.push(await blobToB64(await toBlob(c, .85)));
  }
  let r;
  try {
    r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {"content-type": "application/json", "x-api-key": ai.key, "anthropic-version": "2023-06-01", "anthropic-dangerous-direct-browser-access": "true"},
      body: JSON.stringify({model: ai.model, max_tokens: 8000, messages: [{role: "user", content: [
        ...imgs.map(data => ({type: "image", source: {type: "base64", media_type: "image/jpeg", data}})),
        {type: "text", text: aiPrompt(imgs.length)}]}]})
    });
  } catch { throw {userMsg: "Could not reach the AI service. Check your connection."}; }
  if (!r.ok) {
    let m = ""; try { m = (await r.json()).error?.message || ""; } catch {}
    if (r.status === 401) throw {userMsg: "The AI key was not accepted. Check it in Settings."};
    if (r.status === 429) throw {userMsg: "The AI service is busy or your credit is used up. Try again later."};
    if (r.status === 404 || /model/i.test(m)) throw {userMsg: "The AI model name was not recognised. Check it in Settings."};
    throw {userMsg: "AI reading failed" + (m ? ": " + m : ".")};
  }
  const j = await r.json();
  const raw = (j.content || []).filter(c => c.type === "text").map(c => c.text).join("");
  let obj = null;
  try { obj = JSON.parse(raw); } catch {
    const s = raw.indexOf("{"), e = raw.lastIndexOf("}");
    if (s >= 0 && e > s) { try { obj = JSON.parse(raw.slice(s, e + 1)); } catch {} }
  }
  if (!obj || typeof obj !== "object") throw {userMsg: "The AI answer could not be read. Try again."};
  const space = obj.space === "business" ? "business" : "private";
  const cat = catsFor(space).find(c => norm(c) === norm(obj.category)) || "";
  const iso = v => typeof v === "string" && /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : "";
  return {
    text: String(obj.text || "").slice(0, 40000), title: String(obj.title || "").slice(0, 120), space, category: cat,
    date: iso(obj.date), sender: String(obj.sender || ""), reference: String(obj.reference || ""),
    amount: obj.amount === null || obj.amount === "" || isNaN(Number(obj.amount)) ? null : Number(obj.amount),
    currency: String(obj.currency || "").toUpperCase().slice(0, 3), dueDate: iso(obj.dueDate)
  };
}

/* ================= guessing fields without AI ================= */
const MONTHS = {jan:1,januari:1,january:1,feb:2,februari:2,february:2,mrt:3,maart:3,mar:3,march:3,apr:4,april:4,mei:5,may:5,jun:6,juni:6,june:6,jul:7,juli:7,july:7,aug:8,augustus:8,august:8,sep:9,sept:9,september:9,okt:10,oct:10,oktober:10,october:10,nov:11,november:11,dec:12,december:12};
function findDates(text) {
  const out = [], now = new Date().getFullYear();
  const push = (y, m, d, idx) => { y = +y; m = +m; d = +d; if (y < 100) y += 2000; if (y < 1990 || y > now + 6 || m < 1 || m > 12 || d < 1 || d > 31) return; out.push({iso: `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`, idx}); };
  let m;
  const r1 = /\b(\d{1,2})[-./](\d{1,2})[-./](\d{4}|\d{2})\b/g; while ((m = r1.exec(text))) push(m[3], m[2], m[1], m.index);
  const r2 = /\b(\d{4})-(\d{2})-(\d{2})\b/g; while ((m = r2.exec(text))) push(m[1], m[2], m[3], m.index);
  const r3 = /\b(\d{1,2})\s+([a-z]{3,9})\.?\s+(\d{4})\b/gi; while ((m = r3.exec(text))) { const mo = MONTHS[m[2].toLowerCase()]; if (mo) push(m[3], mo, m[1], m.index); }
  return out;
}
function parseAmount(s) {
  s = s.replace(/\s/g, "");
  if (/,\d{2}$/.test(s)) s = s.replace(/\./g, "").replace(",", ".");
  else s = s.replace(/,/g, "");
  const n = Number(s); return isNaN(n) ? null : n;
}
function guessFields(text) {
  const t = text || "", low = norm(t), g = {};
  const dates = findDates(t);
  if (dates.length) {
    const key = low.search(/(factuurdatum|invoice date|datum|date|aankoopdatum|kassabon)/);
    const near = key >= 0 ? dates.filter(d => d.idx >= key && d.idx - key < 80) : [];
    g.date = (near[0] || dates[0]).iso;
    const dk = low.search(/(vervaldatum|uiterlijk|due date|te betalen voor|betalen voor|garantie tot|valid until|geldig tot)/);
    const dn = dk >= 0 ? dates.filter(d => d.idx >= dk && d.idx - dk < 80) : [];
    if (dn[0] && dn[0].iso !== g.date) g.dueDate = dn[0].iso;
  }
  const amtRe = /(?:€|eur|euro)?\s*(\d{1,3}(?:[.\s]\d{3})*,\d{2}|\d+,\d{2}|\d{1,3}(?:,\d{3})*\.\d{2}|\d+\.\d{2})\b/gi;
  let m, best = null, max = null;
  while ((m = amtRe.exec(t))) {
    const v = parseAmount(m[1]); if (v === null || v > 1e7) continue;
    const before = norm(t.slice(Math.max(0, m.index - 40), m.index));
    if (/(totaal|total|te betalen|amount due|bedrag|to pay|grand total)/.test(before)) best = v;
    if (max === null || v > max) max = v;
  }
  g.amount = best ?? max;
  if (g.amount !== null && g.amount !== undefined) g.currency = /\$|usd/.test(low) && !/€|eur/.test(low) ? "USD" : /£|gbp/.test(low) ? "GBP" : /\betb\b|birr/.test(low) ? "ETB" : "EUR";
  const ref = t.match(/(?:factuurnummer|factuurnr\.?|invoice (?:no|number)\.?|polisnummer|ordernummer|order (?:no|number)\.?|bonnummer|booking ?code|booking reference|boekingscode|reserveringsnummer|dossiernummer)[:\s#]*([A-Z0-9][A-Z0-9\-/.]{2,24})/i);
  if (ref) g.reference = ref[1];
  const first = t.split("\n").map(s => s.trim()).find(s => /[a-z]{3,}/i.test(s) && s.length >= 3 && s.length <= 60);
  if (first) g.sender = niceCase(first.replace(/[|_~*]+/g, " ").replace(/\s+/g, " ").trim());
  return g;
}
function suggestFromHistory(text, exceptId) {
  const nt = norm(text);
  if (!nt) return null;
  let best = null;
  for (const d of docs()) {
    if (!d.filed || d.id === exceptId || !d.sender) continue;
    const s = norm(d.sender).trim();
    if (s.length < 4 || !nt.includes(s)) continue;
    if (!best || s.length > best.len || (s.length === best.len && d.updatedAt > best.doc.updatedAt)) best = {doc: d, len: s.length};
  }
  return best ? {space: best.doc.space, category: best.doc.category, sender: best.doc.sender} : null;
}

/* ================= list and search ================= */
// "2026" = year, ">500" / "<50" / "100-500" = amount; the rest is text
function parseQuery(q) {
  const terms = [], f = {};
  const num = v => parseFloat(String(v).replace(/\./g, "").replace(",", "."));
  for (const raw of norm(q).replace(/€\s+/g, "€").split(/\s+/).filter(Boolean)) {
    let m;
    if (/^(19|20)\d{2}$/.test(raw)) { f.year = raw; continue; }
    if ((m = raw.match(/^€?(>=|<=|>|<)€?(\d[\d.,]*)$/))) { if (m[1][0] === ">") f.min = num(m[2]); else f.max = num(m[2]); continue; }
    if ((m = raw.match(/^€?(\d[\d.,]*)\+$/))) { f.min = num(m[1]); continue; }
    if ((m = raw.match(/^€?(\d[\d.,]*)-€?(\d[\d.,]*)$/)) && num(m[1]) < num(m[2]) && !/^(19|20)\d{2}$/.test(m[1])) { f.min = num(m[1]); f.max = num(m[2]); continue; }
    terms.push(raw);
  }
  return {terms, f};
}
const AMT_FILTERS = {"": null, "lt50": {max: 50}, "gt100": {min: 100}, "gt500": {min: 500}, "gt1000": {min: 1000}};
function filtered() {
  const {terms, f} = parseQuery(state.q);
  const year = f.year || state.year, af = {...(AMT_FILTERS[state.amt] || {}), ...("min" in f ? {min: f.min} : {}), ...("max" in f ? {max: f.max} : {})};
  const out = [];
  for (const d of docs()) {
    if (year && docDate(d).slice(0, 4) !== year) continue;
    if ("min" in af || "max" in af) {
      const a = Number(d.amount);
      if (d.amount === null || d.amount === undefined || d.amount === "" || isNaN(a)) continue;
      if ("min" in af && a < af.min) continue;
      if ("max" in af && a > af.max) continue;
    }
    if (state.space !== "all" && (!d.filed || d.space !== state.space)) continue;
    if (state.cat === "__due") { if (!dueInfo(d)) continue; }
    else if (state.cat === "__inbox") { if (d.filed) continue; }
    else if (state.cat === "__weak") { if (!d.weak) continue; }
    else if (state.cat && (!d.filed || d.category !== state.cat)) continue;
    if (terms.length) {
      const head = norm([d.title, d.sender, d.category, d.reference, (d.tags || []).join(" "), d.notes, d.name, fmtAmount(d.amount, d.currency), d.date].join(" "));
      const body = norm(d.text); let score = 0, ok = true;
      for (const t of terms) {
        const inH = head.includes(t), inB = body.includes(t);
        if (!inH && !inB) { ok = false; break; }
        score += (norm(d.title).includes(t) ? 5 : 0) + (inH ? 2 : 0) + (inB ? 1 : 0);
      }
      if (ok) out.push({d, score});
    } else out.push({d, score: 0});
  }
  if (terms.length) out.sort((a, b) => b.score - a.score || docDate(b.d).localeCompare(docDate(a.d)));
  else out.sort((a, b) => docDate(b.d).localeCompare(docDate(a.d)) || (b.d.createdAt || 0) - (a.d.createdAt || 0));
  return {list: out.map(o => o.d), terms};
}
function highlight(text, terms) {
  if (!terms.length) return esc(text);
  const nt = norm(text), ranges = [];
  for (const t of terms) { let i = 0; while ((i = nt.indexOf(t, i)) !== -1) { ranges.push([i, i + t.length]); i += t.length; } }
  if (!ranges.length) return esc(text);
  ranges.sort((a, b) => a[0] - b[0]);
  let res = "", pos = 0;
  for (const [s, e] of ranges) { if (s < pos) continue; res += esc(text.slice(pos, s)) + "<mark>" + esc(text.slice(s, e)) + "</mark>"; pos = e; }
  return res + esc(text.slice(pos));
}
function snippet(text, terms) {
  if (!text || !terms.length) return "";
  const nt = norm(text); let i = -1;
  for (const t of terms) { i = nt.indexOf(t); if (i >= 0) break; }
  if (i < 0) return "";
  const s = Math.max(0, i - 50), raw = text.slice(s, s + 160).replace(/\s+/g, " ");
  return (s > 0 ? "…" : "") + highlight(raw, terms) + "…";
}
function renderChips() {
  const all = docs(), inSpace = all.filter(d => state.space === "all" || (d.filed && d.space === state.space));
  const counts = new Map(); let due = 0, inbox = 0, weak = 0;
  for (const d of inSpace) { if (d.filed) counts.set(d.category, (counts.get(d.category) || 0) + 1); if (dueInfo(d)) due++; if (d.weak) weak++; }
  for (const d of all) if (!d.filed) inbox++;
  const order = state.space === "all" ? [...new Set([...catsFor("private"), ...catsFor("business")])] : catsFor(state.space);
  const known = order.filter(c => counts.has(c)), extra = [...counts.keys()].filter(c => !order.includes(c));
  if (state.cat && !state.cat.startsWith("__") && !counts.has(state.cat)) state.cat = null;
  if ((state.cat === "__inbox" && !inbox) || (state.cat === "__due" && !due) || (state.cat === "__weak" && !weak)) state.cat = null;
  let h = `<button class="chip" data-cat="" aria-pressed="${!state.cat}">All<span class="n">${inSpace.length}</span></button>`;
  if (inbox) h += `<button class="chip due" data-cat="__inbox" aria-pressed="${state.cat === "__inbox"}">To file<span class="n">${inbox}</span></button>`;
  if (due) h += `<button class="chip due" data-cat="__due" aria-pressed="${state.cat === "__due"}">Due soon<span class="n">${due}</span></button>`;
  if (weak) h += `<button class="chip weak" data-cat="__weak" aria-pressed="${state.cat === "__weak"}">Hard to read<span class="n">${weak}</span></button>`;
  for (const c of [...known, ...extra]) h += `<button class="chip" data-cat="${esc(c)}" aria-pressed="${state.cat === c}">${esc(c)}<span class="n">${counts.get(c)}</span></button>`;
  $("#chips").innerHTML = h;
}
function render() {
  if ($("#app").hidden) return;
  renderChips();
  const L = $("#list");
  if (!state.loaded) { L.innerHTML = `<div class="empty"><p>Loading your documents from OneDrive…</p></div>`; return; }
  if (!docs().length) {
    L.innerHTML = `<div class="empty"><h2>Your vault is empty</h2><p>Scan letters, invoices, receipts and warranties, or add PDFs and photos. Every file is stored in your own OneDrive and the text inside is made searchable.</p>
      <ol class="steps"><li><span><b>Scan</b> pages with your phone camera, or drop PDFs and photos here on your PC.</span></li>
      <li><span><b>File</b> it under Private or Business. Documents from a sender you filed before are filed automatically.</span></li>
      <li><span><b>Find</b> it later by any word in the document, such as “Trek” on a bike shop receipt.</span></li></ol>
      <div class="btnrow"><button class="btn primary" data-act="scan">${ICON.camera}Scan a document</button><label class="btn" for="fileIn">Add files</label></div></div>`;
    return;
  }
  const {list, terms} = filtered();
  renderFilterBar(list);
  if (!list.length) { L.innerHTML = `<div class="empty"><h2>No matches</h2><p>Nothing matches ${state.q ? `“${esc(state.q)}”` : "this filter"}. Try a different word, another year or amount, or switch to All.</p></div>`; return; }
  let h = "", lastMonth = "";
  for (const d of list.slice(0, 400)) {
    if (!terms.length) {
      const dd = docDate(d), m = dd ? new Date(dd + "T12:00:00").toLocaleDateString("en-GB", {month: "long", year: "numeric"}) : "Undated";
      if (m !== lastMonth) { h += `<div class="month">${esc(m)}</div>`; lastMonth = m; }
    }
    const due = dueInfo(d), amt = fmtAmount(d.amount, d.currency);
    const sub = [d.sender, fmtDate(d.date)].filter(Boolean).join(" · ") || d.name || "";
    const sn = terms.length ? snippet(d.text, terms) : "";
    const sel = state.selecting && state.sel.has(d.id);
    h += `<button class="doc${state.selecting ? " selecting" : ""}${sel ? " selected" : ""}" data-id="${esc(d.id)}" ${state.selecting ? `aria-pressed="${sel}"` : ""}>
      ${state.selecting ? `<span class="tick" aria-hidden="true">${sel ? ICON.check : ""}</span>` : ""}<img class="thumb" data-thumb="${esc(d.itemId || "")}" data-v="${esc(String(d.fileVersion || 0))}" alt="">
      <span class="meta"><span class="t">${terms.length ? highlight(d.title || "Untitled", terms) : esc(d.title || "Untitled")}</span>
      <span class="sub">${esc(sub)}</span>
      <span class="tags">${d.filed ? `<span class="pill ${d.space}">${SPACE_LABEL[d.space]}</span><span class="cat">${esc(d.category)}</span>` : `<span class="due">To file</span>`}${d.weak ? `<span class="weakpill">Hard to read</span>` : ""}${due ? `<span class="due">${due.label}</span>` : ""}${amt ? `<span class="amt mono">${esc(amt)}</span>` : ""}</span>
      ${sn ? `<span class="snip">${sn}</span>` : ""}</span></button>`;
  }
  if (list.length > 400) h += `<p class="hint">Showing 400 of ${list.length}. Search to narrow down.</p>`;
  L.innerHTML = h;
  observeThumbs();
  renderSelBar();
}
function sumAmounts(list) {
  const by = new Map();
  for (const d of list) { const a = Number(d.amount); if (d.amount === null || d.amount === undefined || d.amount === "" || isNaN(a)) continue; const c = d.currency || "EUR"; by.set(c, (by.get(c) || 0) + a); }
  return [...by.entries()].sort((a, b) => (b[0] === "EUR") - (a[0] === "EUR")).map(([c, v]) => fmtAmount(v, c)).join(" + ");
}
function renderFilterBar(list) {
  const years = [...new Set(docs().map(d => docDate(d).slice(0, 4)).filter(Boolean))].sort().reverse();
  const yOpts = `<option value="">All years</option>` + years.map(y => `<option${y === state.year ? " selected" : ""}>${y}</option>`).join("");
  const aOpts = [["", "Any amount"], ["lt50", "Under €50"], ["gt100", "Over €100"], ["gt500", "Over €500"], ["gt1000", "Over €1,000"]].map(([v, l]) => `<option value="${v}"${v === state.amt ? " selected" : ""}>${l}</option>`).join("");
  const total = sumAmounts(list);
  $("#filterbar").innerHTML = `<label class="sr" for="fYear">Year</label><select id="fYear">${yOpts}</select>
    <label class="sr" for="fAmt">Amount</label><select id="fAmt">${aOpts}</select>
    <span class="fsum">${list.length} document${list.length === 1 ? "" : "s"}${total ? ` · total <b class="mono">${esc(total)}</b>` : ""}</span>
    <span class="fbtns"><button class="btn" data-act="select">${state.selecting ? "Cancel selection" : "Select"}</button><button class="btn" data-act="export" ${list.length ? "" : "disabled"}>${ICON.download}Export these</button></span>`;
}
function renderSelBar() {
  const bar = $("#selbar"), n = state.sel.size;
  $("#fab").hidden = state.selecting;
  if (!state.selecting) { bar.hidden = true; return; }
  bar.hidden = false;
  bar.innerHTML = `<span><b>${n}</b> selected</span><button class="btn" data-act="selall">Select all shown</button><button class="btn primary" data-act="merge" ${n >= 2 ? "" : "disabled"}>Merge into one</button><button class="btn" data-act="selexport" ${n ? "" : "disabled"}>${ICON.download}Export</button><button class="btn ghost" data-act="selcancel">Cancel</button>`;
}

/* thumbnails from OneDrive, cached on this device */
const thumbMem = new Map();
let thumbObs = null, thumbQueue = [], thumbActive = 0;
function observeThumbs() {
  if (!thumbObs) thumbObs = new IntersectionObserver(es => { for (const e of es) if (e.isIntersecting) { thumbObs.unobserve(e.target); queueThumb(e.target); } }, {rootMargin: "300px"});
  document.querySelectorAll("img[data-thumb]").forEach(img => {
    const k = img.dataset.thumb + ":" + img.dataset.v;
    if (!img.dataset.thumb) return;
    if (thumbMem.has(k)) { img.src = thumbMem.get(k); return; }
    thumbObs.observe(img);
  });
}
function queueThumb(img) { thumbQueue.push(img); pumpThumbs(); }
async function pumpThumbs() {
  while (thumbActive < 4 && thumbQueue.length) {
    const img = thumbQueue.shift(); thumbActive++;
    loadThumb(img.dataset.thumb, img.dataset.v).then(u => { if (u) img.src = u; }).finally(() => { thumbActive--; pumpThumbs(); });
  }
}
async function loadThumb(itemId, v) {
  const k = itemId + ":" + v;
  if (thumbMem.has(k)) return thumbMem.get(k);
  let blob = await idb.get("thumbs", k);
  if (!blob && state.online) {
    try {
      const r = await g("GET", `/me/drive/items/${itemId}/thumbnails?select=medium`);
      const url = r.value && r.value[0] && r.value[0].medium && r.value[0].medium.url;
      if (url) {
        try { const x = await fetch(url); if (x.ok) { blob = await x.blob(); idb.set("thumbs", k, blob); } }
        catch { thumbMem.set(k, url); return url; }
      }
    } catch {}
  }
  if (!blob) return null;
  const u = URL.createObjectURL(blob); thumbMem.set(k, u); return u;
}

/* ================= panels ================= */
const panel = $("#panel");
function showPanel(title, bodyHtml, {right = "", docId = ""} = {}) {
  panel.innerHTML = `<div class="phead"><div class="wrap"><button class="iconbtn" data-act="close" aria-label="Back">${ICON.back}</button><h2>${esc(title)}</h2>${right}</div></div><div class="wrap pbody">${bodyHtml}</div>`;
  panel.hidden = false; panel.scrollTop = 0; document.body.style.overflow = "hidden";
  if (docId) panel.dataset.docId = docId; else delete panel.dataset.docId;
}
function closePanel() { panel.hidden = true; panel.innerHTML = ""; panel.onclick = null; document.body.style.overflow = ""; delete panel.dataset.docId; }

/* ================= form ================= */
function fieldsFrom(d = {}) {
  return {title: d.title || "", space: d.space || (state.space !== "all" ? state.space : "private"), category: d.category || "",
    date: d.date || "", sender: d.sender || "", reference: d.reference || "", amount: d.amount ?? null, currency: d.currency || "EUR",
    dueDate: d.dueDate || "", tags: d.tags || [], notes: d.notes || "", text: d.text || ""};
}
function formHtml(f, {textOpen = false} = {}) {
  const cats = catsFor(f.space);
  const opts = cats.map(c => `<option${c === f.category ? " selected" : ""}>${esc(c)}</option>`).join("");
  const curOpts = [...new Set([...CURRENCIES, f.currency].filter(Boolean))].map(c => `<option${c === f.currency ? " selected" : ""}>${esc(c)}</option>`).join("");
  return `<form class="form" id="docForm" novalidate>
    <div class="field"><label for="f_title">Title</label><input id="f_title" name="title" value="${esc(f.title)}" placeholder="e.g. Trek bike, Piet Fiets" required></div>
    <div class="field"><span class="lbl">File under</span><div class="seg" id="f_space">
      <button type="button" data-sp="private" aria-pressed="${f.space === "private"}">Private</button><button type="button" data-sp="business" aria-pressed="${f.space === "business"}">Business</button></div></div>
    <div class="field"><label for="f_cat">Category</label><select id="f_cat" name="category">${opts}</select></div>
    <div class="two"><div class="field"><label for="f_date">Document date</label><input id="f_date" name="date" type="date" value="${esc(f.date)}"></div>
      <div class="field"><label for="f_sender">From</label><input id="f_sender" name="sender" value="${esc(f.sender)}" placeholder="Company or person"></div></div>
    <div class="two"><div class="field"><label for="f_amount">Amount</label><div class="amtrow"><input id="f_amount" name="amount" class="mono" inputmode="decimal" value="${f.amount ?? ""}" placeholder="0.00"><select id="f_cur" name="currency" aria-label="Currency">${curOpts}</select></div></div>
      <div class="field"><label for="f_due">Due date or warranty end</label><input id="f_due" name="dueDate" type="date" value="${esc(f.dueDate)}"></div></div>
    <div class="two"><div class="field"><label for="f_ref">Reference number</label><input id="f_ref" name="reference" class="mono" value="${esc(f.reference)}" placeholder="Invoice or policy no."></div>
      <div class="field"><label for="f_tags">Tags</label><input id="f_tags" name="tags" value="${esc(f.tags.join(", "))}" placeholder="comma separated"></div></div>
    <div class="field"><label for="f_notes">Notes</label><textarea id="f_notes" name="notes" rows="2" placeholder="Anything you want to find this by later">${esc(f.notes)}</textarea></div>
    <details${textOpen || !f.text ? " open" : ""}><summary>Text in the document${f.text ? ` <span class="hint">(${f.text.length.toLocaleString("en-GB")} characters, used for search)</span>` : ""}</summary>
      <div class="field" style="margin-top:10px"><textarea id="f_text" name="text" rows="8" class="mono" placeholder="No text was found. Type key words here, such as the brand or product, to make it searchable.">${esc(f.text)}</textarea></div></details>
  </form>`;
}
function wireForm(f) {
  const form = $("#docForm", panel);
  $("#f_space", panel).addEventListener("click", e => {
    const b = e.target.closest("[data-sp]"); if (!b) return;
    f.space = b.dataset.sp;
    for (const x of $("#f_space", panel).children) x.setAttribute("aria-pressed", x === b);
    const cur = $("#f_cat", panel).value;
    $("#f_cat", panel).innerHTML = catsFor(f.space).map(c => `<option${c === cur ? " selected" : ""}>${esc(c)}</option>`).join("");
  });
  return () => {
    const v = n => form.elements[n].value.trim();
    const amtRaw = v("amount").replace(/\s/g, "");
    return {title: v("title") || "Untitled document", space: f.space, category: v("category"), date: v("date"), sender: v("sender"),
      reference: v("reference"), amount: amtRaw === "" ? null : parseAmount(amtRaw.includes(",") || amtRaw.includes(".") ? amtRaw : amtRaw + ".00"), currency: v("currency"),
      dueDate: v("dueDate"), tags: v("tags").split(",").map(t => t.trim().toLowerCase()).filter(Boolean).slice(0, 12),
      notes: v("notes"), text: form.elements.text.value.slice(0, 60000)};
  };
}
const folderOf = d => d.filed ? `${SPACE_LABEL[d.space]}/${safeSeg(d.category, 60)}` : INBOX;
// "2025-03-14 Piet Fiets - Trek FX 3 bike.pdf": date, sender, title; sender left out when the title already names it
const niceCase = s => /[a-z]/.test(s) ? s : String(s).toLowerCase().replace(/(^|[\s\-/&.(])(\p{L})/gu, (m, a, b) => a + b.toUpperCase());
function fileNameOf(d, ext) {
  const cut = (t, n) => t.length <= n ? t : (t.slice(0, n).replace(/[\s,;:-]+\S*$/, "") || t.slice(0, n));
  const title = cut(safeSeg(niceCase(d.title || ""), 200), 70).replace(/\b\d{1,2} [A-Z][a-z]{2} \d{4}\b/g, "").replace(/[\s,;-]+$/, "").trim() || "Document";
  const sender = d.sender ? niceCase(d.sender).replace(/[\\/:*?"<>|#%\u0000-\u001f]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 40).trim() : "";
  const parts = [d.date || "undated"];
  if (sender && !norm(title).includes(norm(sender))) parts.push(`${sender} -`);
  parts.push(title);
  return parts.join(" ") + ext;
}
const extOf = name => (String(name || "").match(/\.[a-z0-9]{2,5}$/i) || [".pdf"])[0].toLowerCase();

/* ================= scan flow (phone camera) ================= */
let draft = null;
function startScan(openCamera) {
  draft = {pages: [], mode: "enhanced"}; renderScan();
  if (openCamera) $("#camIn").click();
}
async function addScanFiles(files, fromCamera = false) {
  if (!draft) return;
  draft.justAdded = fromCamera && files.length > 0;
  for (const f of files) {
    try {
      const bitmap = await createImageBitmap(f, {imageOrientation: "from-image"});
      const p = {bitmap, rot: 0, crop: null, preview: null};
      draft.pages.push(p); await refreshPreview(p);
    } catch { toast("That file could not be opened as an image."); }
  }
  renderScan();
}
async function refreshPreview(p) {
  const b = await toBlob(renderPage(p, draft.mode, {maxSide: 700}), .8);
  if (p.preview) URL.revokeObjectURL(p.preview);
  p.preview = URL.createObjectURL(b);
}
$("#camIn").addEventListener("change", e => { const f = [...e.target.files]; e.target.value = ""; addScanFiles(f, true); });
$("#galIn").addEventListener("change", e => { const f = [...e.target.files]; e.target.value = ""; addScanFiles(f); });

function renderScan() {
  const n = draft.pages.length;
  const pagesHtml = draft.pages.map((p, i) => `<div class="pg"><img src="${p.preview}" alt="Page ${i + 1}"><div class="bar"><span>Page ${i + 1}</span><div>
      <button class="iconbtn" data-act="rot" data-i="${i}" aria-label="Rotate page ${i + 1}">${ICON.rotate}</button>
      <button class="iconbtn" data-act="crop" data-i="${i}" aria-label="Crop page ${i + 1}">${ICON.crop}</button>
      <button class="iconbtn" data-act="del" data-i="${i}" aria-label="Remove page ${i + 1}">${ICON.trash}</button></div></div></div>`).join("");
  const ask = n && draft.justAdded;
  showPanel("New scan", `
    ${ask ? `<div class="card nextpage"><h3>Page ${n} added. Does the document have more pages?</h3>
      <div class="btnrow"><button class="btn primary grow" data-act="cam">${ICON.camera}Scan page ${n + 1}</button><button class="btn grow" data-act="donepages">${ICON.check}No, that's all</button></div></div>` : ""}
    ${n ? `<div class="pages">${pagesHtml}</div>` : `<div class="card"><p class="hint">Take a photo of each page. Lay the document flat in good light and fill the frame. You can add more pages after the first one.</p></div>`}
    ${ask ? "" : `<div class="btnrow"><button class="btn grow" data-act="cam">${ICON.camera}${n ? "Add page" : "Take photo"}</button><button class="btn grow" data-act="gal">${ICON.image}From gallery</button></div>`}
    ${n ? `<div class="field"><span class="label">Image style</span><div class="seg" id="modeSeg">
        ${[["original", "Original"], ["enhanced", "Enhanced"], ["bw", "Black & white"]].map(([k, l]) => `<button data-mode="${k}" aria-pressed="${draft.mode === k}">${l}</button>`).join("")}</div></div>
      ${ask ? "" : `<div class="btnrow"><button class="btn primary grow" data-act="read">${ICON.text}Read text and continue (${n} page${n > 1 ? "s" : ""})</button></div>`}
      <p class="hint">The text is read on this phone. ${ai.on && ai.auto ? "If a page is hard to read, it is sent to AI for a second try." : ""}</p>` : ""}`);
  panel.onclick = onScanClick;
}
async function onScanClick(e) {
  const b = e.target.closest("[data-act],[data-mode]"); if (!b) return;
  if (b.dataset.mode) { draft.mode = b.dataset.mode; for (const p of draft.pages) await refreshPreview(p); renderScan(); return; }
  const i = Number(b.dataset.i), act = b.dataset.act;
  if (act === "close") { draft = null; return closePanel(); }
  if (act === "cam") return $("#camIn").click();
  if (act === "gal") return $("#galIn").click();
  if (act === "donepages") { draft.justAdded = false; return renderScan(); }
  draft.justAdded = false;
  if (act === "rot") { const p = draft.pages[i]; p.rot = (p.rot + 90) % 360; p.crop = null; await refreshPreview(p); return renderScan(); }
  if (act === "del") { const p = draft.pages.splice(i, 1)[0]; if (p.preview) URL.revokeObjectURL(p.preview); return renderScan(); }
  if (act === "crop") return openCrop(i);
  if (act === "read") return readScan();
}
function progressPanel(title, steps, foot = "") {
  showPanel(title, `<div class="progress">${steps.map((s, i) => `<div class="stepline" data-s="${i}"><span class="dot"></span><span class="txt">${esc(s)}</span></div>`).join("")}</div>${foot}`);
  return (i, cls, txt) => { const el = panel.querySelector(`[data-s="${i}"]`); if (!el) return; el.className = "stepline " + cls; if (txt) el.querySelector(".txt").textContent = txt; };
}
async function readScan() {
  const n = draft.pages.length;
  const step = progressPanel("Reading document", [`Preparing ${n} page${n > 1 ? "s" : ""}`, "Reading the text", "Checking the result"]);
  panel.onclick = e => { if (e.target.closest('[data-act="close"]')) renderScan(); };
  step(0, "on");
  const blobs = [], canvases = [];
  for (const p of draft.pages) { const c = renderPage(p, draft.mode); canvases.push(c); blobs.push(await toBlob(c)); }
  step(0, "done"); step(1, "on");
  const texts = []; let confSum = 0;
  try {
    for (let i = 0; i < canvases.length; i++) {
      ocrProgress = m => { if (m.status === "recognizing text") step(1, "on", `Reading page ${i + 1} of ${n} · ${Math.round(m.progress * 100)}%`); else if (/loading|initializing/.test(m.status)) step(1, "on", "Loading the text reader (first time only)…"); };
      const r = await ocrCanvas(canvases[i]); texts.push(r.text); confSum += r.conf;
    }
  } catch (err) { console.error(err); texts.length = 0; toast("The text reader could not start. You can type key words yourself."); }
  ocrProgress = null;
  step(1, "done"); step(2, "on");
  let text = texts.join("\n\n").trim(), conf = texts.length ? Math.round(confSum / texts.length) : 0, source = "ocr", weak = isWeak(text, conf), aiData = null, aiErr = "";
  if (weak && ai.on && ai.auto) {
    step(2, "on", "Hard to read. Asking AI for a second try…");
    try { aiData = await aiRead(blobs); text = aiData.text || text; source = "ai"; weak = false; texts.splice(0, texts.length, ...[text]); }
    catch (err) { aiErr = err.userMsg || "AI reading failed."; }
  }
  step(2, "done");
  showScanReview({blobs, texts, text, conf, source, weak, aiData, aiErr});
}
function showScanReview(r) {
  const guess = guessFields(r.text), hist = suggestFromHistory(r.text);
  const f = fieldsFrom({...guess, ...(hist || {}), ...(r.aiData ? Object.fromEntries(Object.entries(r.aiData).filter(([, v]) => v !== "" && v !== null)) : {}), text: r.text});
  if (!f.title) f.title = f.sender || "Scanned document";
  if (!catsFor(f.space).includes(f.category)) f.category = catsFor(f.space).at(-1);
  const urls = r.blobs.map(b => URL.createObjectURL(b));
  showPanel("Check details", `
    ${r.weak ? `<div class="flag weak">${ICON.text}<span>${r.aiErr ? esc(r.aiErr) + " " : ""}The text was hard to read. Add key words below so you can find it later${ai.on ? ", or try AI" : ""}.</span>${ai.on ? `<button class="btn" data-act="ai">${ICON.spark}Read with AI</button>` : ""}</div>`
      : r.source === "ai" ? `<div class="flag info">${ICON.spark}<span>AI read this document. Check the details, then save.</span></div>`
      : hist ? `<div class="flag info">${ICON.check}<span>Filed like your earlier documents from ${esc(hist.sender)}. Check and save.</span></div>` : ""}
    <div class="viewer">${urls.map((u, i) => `<img src="${u}" alt="Page ${i + 1}">`).join("")}</div>
    ${formHtml(f)}
    <div class="btnrow"><button class="btn primary grow" data-act="save">${ICON.check}Save to OneDrive</button></div>`);
  const getVals = wireForm(f);
  panel.onclick = async e => {
    const b = e.target.closest("[data-act]"); if (!b) return;
    if (b.dataset.act === "close") return renderScan();
    if (b.dataset.act === "ai") {
      b.disabled = true; b.textContent = "Reading…";
      try { const a = await aiRead(r.blobs); showScanReview({...r, aiData: a, text: a.text || r.text, texts: [a.text || r.text], source: "ai", weak: false, aiErr: ""}); }
      catch (err) { toast(err.userMsg || "AI reading failed."); b.disabled = false; b.innerHTML = ICON.spark + "Read with AI"; }
      return;
    }
    if (b.dataset.act === "save") {
      b.disabled = true; b.textContent = "Saving…";
      try {
        const vals = getVals();
        const meta = {id: uid(), ...vals, filed: true, weak: r.weak && !vals.text.trim() ? true : r.weak && vals.text === r.text, textSource: r.source, ocrConf: r.conf,
          pages: r.blobs.length, mime: "application/pdf", source: "scan", createdAt: Date.now(), handled: false};
        const pdf = await buildScanPdf(r.blobs, r.source === "ai" ? [r.text] : r.texts, meta);
        const it = await uploadFile(folderOf(meta), fileNameOf(meta, ".pdf"), pdf);
        Object.assign(meta, {itemId: it.id, name: it.name, path: `${folderOf(meta)}/${it.name}`, size: it.size, webUrl: it.webUrl, fileVersion: 1});
        putDoc(meta); urls.forEach(u => URL.revokeObjectURL(u)); draft = null; closePanel(); toast("Saved to OneDrive");
      } catch (err) { console.error(err); b.disabled = false; b.innerHTML = ICON.check + "Save to OneDrive"; toast(errText(err)); }
    }
  };
}
function errText(e) {
  if (e && e.userMsg) return e.userMsg;
  if (e && e.code === "offline") return "You are offline. Connect to the internet and try again.";
  if (e && e.status === 507) return "Your OneDrive is full.";
  if (e && (e.status === 401 || e.status === 403)) return "OneDrive refused access. Sign out and in again in Settings.";
  const detail = e && (e.message || e.code || e.status);
  return "Saving failed" + (detail ? `: ${String(detail).slice(0, 140)}` : ". Try again.");
}

/* crop */
async function openCrop(i) {
  const p = draft.pages[i];
  const url = URL.createObjectURL(await toBlob(renderPage(p, "original", {maxSide: 1100, crop: false, enhance: false}), .85));
  let box = p.crop ? {...p.crop} : {x: .04, y: .04, w: .92, h: .92};
  showPanel(`Crop page ${i + 1}`, `<p class="hint">Drag the corners to the edges of the paper.</p>
    <div class="cropwrap" id="cw"><img src="${url}" alt="Page ${i + 1}" draggable="false"><div class="cropclip"><div class="cropbox" id="cb"></div></div>
    <div id="hd" style="position:absolute"><span class="handle tl" data-h="tl"></span><span class="handle tr" data-h="tr"></span><span class="handle bl" data-h="bl"></span><span class="handle br" data-h="br"></span></div></div>
    <div class="btnrow"><button class="btn" data-act="reset">Use whole photo</button><button class="btn primary grow" data-act="apply">${ICON.check}Apply crop</button></div>`);
  const cw = $("#cw", panel), cb = $("#cb", panel), hd = $("#hd", panel);
  const place = () => { for (const el of [cb, hd]) { el.style.left = box.x * 100 + "%"; el.style.top = box.y * 100 + "%"; el.style.width = box.w * 100 + "%"; el.style.height = box.h * 100 + "%"; } };
  place();
  let drag = null;
  cw.addEventListener("pointerdown", e => {
    const h = e.target.dataset.h || (e.target === cb ? "move" : null); if (!h) return;
    e.preventDefault(); cw.setPointerCapture(e.pointerId); drag = {h, sx: e.clientX, sy: e.clientY, b: {...box}};
  });
  cw.addEventListener("pointermove", e => {
    if (!drag) return;
    const r = cw.getBoundingClientRect(), dx = (e.clientX - drag.sx) / r.width, dy = (e.clientY - drag.sy) / r.height, b = drag.b, min = .08;
    let {x, y, w, h} = b;
    if (drag.h === "move") { x = Math.min(1 - w, Math.max(0, b.x + dx)); y = Math.min(1 - h, Math.max(0, b.y + dy)); }
    else {
      let x1 = b.x, y1 = b.y, x2 = b.x + b.w, y2 = b.y + b.h;
      if (drag.h.includes("l")) x1 = Math.min(x2 - min, Math.max(0, b.x + dx));
      if (drag.h.includes("r")) x2 = Math.max(x1 + min, Math.min(1, b.x + b.w + dx));
      if (drag.h.includes("t")) y1 = Math.min(y2 - min, Math.max(0, b.y + dy));
      if (drag.h.includes("b")) y2 = Math.max(y1 + min, Math.min(1, b.y + b.h + dy));
      x = x1; y = y1; w = x2 - x1; h = y2 - y1;
    }
    box = {x, y, w, h}; place();
  });
  const end = () => drag = null;
  cw.addEventListener("pointerup", end); cw.addEventListener("pointercancel", end);
  panel.onclick = async e => {
    const b = e.target.closest("[data-act]"); if (!b) return;
    if (b.dataset.act === "close") { URL.revokeObjectURL(url); return renderScan(); }
    if (b.dataset.act === "reset") { box = {x: 0, y: 0, w: 1, h: 1}; place(); return; }
    if (b.dataset.act === "apply") { p.crop = box.w > .995 && box.h > .995 ? null : box; URL.revokeObjectURL(url); await refreshPreview(p); renderScan(); }
  };
}

/* ================= PC import: PDFs and images ================= */
const queue = []; let queueRunning = false;
const isPdfFile = f => f.type === "application/pdf" || /\.pdf$/i.test(f.name || "");
const isMailFile = f => /\.(msg|eml)$/i.test(f.name || "") || f.type === "message/rfc822" || f.type === "application/vnd.ms-outlook";
function enqueueFiles(files) {
  const ok = files.filter(f => isPdfFile(f) || isMailFile(f) || /^image\//.test(f.type));
  if (files.length > ok.length) toast(`${files.length - ok.length} file${files.length - ok.length > 1 ? "s" : ""} skipped. Only PDF, image and e-mail files can be added.`);
  if (!ok.length) return;
  const imgs = ok.filter(f => !isPdfFile(f) && !isMailFile(f));
  if (imgs.length >= 2) return askCombine(ok, imgs);
  for (const f of ok) queue.push({file: f, name: f.name || "Pasted image.png", status: "waiting", note: ""});
  renderImport(); runQueue();
}
let pendingCombine = null;
function askCombine(all, imgs) {
  pendingCombine = {all, imgs};
  $("#combineBox").innerHTML = `<div class="card nextpage"><h3>${imgs.length} photos selected. Are they pages of one document?</h3>
    <div class="btnrow"><button class="btn primary grow" data-act="combine">${ICON.text}Yes, one document of ${imgs.length} pages</button><button class="btn grow" data-act="separate">No, ${imgs.length} separate documents</button></div></div>`;
  $("#combineBox").scrollIntoView({block: "nearest"});
}
$("#combineBox").addEventListener("click", e => {
  const b = e.target.closest("[data-act]"); if (!b || !pendingCombine) return;
  const {all, imgs} = pendingCombine; pendingCombine = null; $("#combineBox").innerHTML = "";
  if (b.dataset.act === "combine") {
    for (const f of all) if (isPdfFile(f) || isMailFile(f)) queue.push({file: f, name: f.name, status: "waiting", note: ""});
    queue.push({combine: imgs, name: `${imgs.length} photos as one document`, status: "waiting", note: ""});
  } else for (const f of all) queue.push({file: f, name: f.name || "Pasted image.png", status: "waiting", note: ""});
  renderImport(); runQueue();
});
/* ---- e-mails (.msg from Outlook, .eml) ---- */
const winAnsi = t => String(t || "").normalize("NFKC").replace(/[\u2018\u2019]/g, "'").replace(/[\u201C\u201D]/g, '"').replace(/[\u2013\u2014]/g, "-").replace(/\u2026/g, "...")
  .replace(/[^\x09\x0A\x0D\x20-\x7E\u00A0-\u00FF\u20AC]/g, " ");
function cleanMailBody(t) {
  return String(t || "").replace(/\r/g, "").replace(/<(https?:[^>\s]+\.(gif|png|jpe?g)(\?[^>]*)?)>/gi, "").replace(/<mailto:[^>]+>/gi, "")
    .replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}
function htmlToText(h) {
  const d = new DOMParser().parseFromString(String(h || ""), "text/html");
  d.querySelectorAll("script,style,head").forEach(n => n.remove());
  d.querySelectorAll("br").forEach(n => n.replaceWith("\n"));
  d.querySelectorAll("p,div,tr,li,h1,h2,h3,h4,table").forEach(n => n.append("\n"));
  return (d.body ? d.body.textContent : "").replace(/\u00A0/g, " ");
}
async function parseMail(f, buf) {
  await loadScript("vendor/vault-mail.min.js");
  const {MsgReader, PostalMime} = window.VaultMail;
  if (/\.msg$/i.test(f.name || "") || f.type === "application/vnd.ms-outlook") {
    const r = new MsgReader(buf), d = r.getFileData();
    if (d.error) throw {userMsg: "This Outlook e-mail could not be read."};
    const atts = (d.attachments || []).filter(a => !a.attachmentHidden && a.fileName).map(a => {
      const x = r.getAttachment(a);
      return {name: x.fileName || a.fileName, type: a.attachMimeTag || "", data: x.content};
    });
    const date = d.messageDeliveryTime || d.clientSubmitTime || d.creationTime;
    return {subject: d.subject || "", fromName: d.senderName || "", fromEmail: d.senderEmail || d.senderSmtpAddress || "",
      to: (d.recipients || []).map(x => x.name || x.email).filter(Boolean).join(", "), date: date ? new Date(date) : null,
      body: cleanMailBody(d.body || htmlToText(d.bodyHtml || "")), attachments: atts};
  }
  const m = await PostalMime.parse(buf);
  return {subject: m.subject || "", fromName: m.from?.name || "", fromEmail: m.from?.address || "",
    to: (m.to || []).map(x => x.name || x.address).join(", "), date: m.date ? new Date(m.date) : null,
    body: cleanMailBody(m.text || htmlToText(m.html || "")),
    attachments: (m.attachments || []).filter(a => a.disposition !== "inline" || /pdf/i.test(a.mimeType || "")).map(a => ({name: a.filename || "attachment", type: a.mimeType || "", data: new Uint8Array(a.content)}))};
}
async function mailPdf(mail, title) {
  await loadScript("vendor/jspdf.umd.min.js");
  const {jsPDF} = window.jspdf;
  const pdf = new jsPDF({unit: "mm", format: "a4", compress: true});
  const W = pdf.internal.pageSize.getWidth(), H = pdf.internal.pageSize.getHeight(), M = 18;
  let y = M;
  const line = (txt, size, style = "normal", gap = 1.5) => {
    pdf.setFont("helvetica", style); pdf.setFontSize(size);
    for (const l of pdf.splitTextToSize(winAnsi(txt), W - 2 * M)) {
      if (y > H - M) { pdf.addPage(); y = M; }
      pdf.text(l, M, y); y += size * 0.3528 * 1.35;
    }
    y += gap;
  };
  line(mail.subject || "(no subject)", 15, "bold", 3);
  pdf.setTextColor(90);
  line(`From: ${mail.fromName}${mail.fromEmail ? ` <${mail.fromEmail}>` : ""}`, 9.5, "normal", 0);
  if (mail.to) line(`To: ${mail.to}`, 9.5, "normal", 0);
  if (mail.date) line(`Date: ${mail.date.toLocaleString("en-GB", {dateStyle: "long", timeStyle: "short"})}`, 9.5, "normal", 0);
  if (mail.attachments.length) line(`Attachments: ${mail.attachments.map(a => a.name).join(", ")}`, 9.5, "normal", 0);
  pdf.setTextColor(0); y += 3; pdf.setDrawColor(200); pdf.line(M, y - 2, W - M, y - 2); y += 3;
  line(mail.body || "(no text)", 10.5, "normal", 0);
  pdf.setProperties({title, subject: mail.subject, author: mail.fromName, creator: "Document Vault"});
  return pdf.output("blob");
}
async function importEmail(it, f, buf, hash) {
  it.note = "Reading e-mail"; renderImport();
  const mail = await parseMail(f, buf);
  const iso = mail.date && !isNaN(mail.date) ? mail.date.toISOString().slice(0, 10) : "";
  const text = [`Subject: ${mail.subject}`, `From: ${mail.fromName} ${mail.fromEmail}`, mail.to ? `To: ${mail.to}` : "", mail.attachments.length ? `Attachments: ${mail.attachments.map(a => a.name).join(", ")}` : "", "", mail.body].filter(x => x !== "").join("\n");
  const guess = guessFields(mail.body), hist = suggestFromHistory(`${mail.fromName}\n${mail.fromEmail}\n${text}`);
  const base = fieldsFrom({...guess, ...(hist || {}), title: mail.subject || "E-mail", sender: niceCase(mail.fromName || mail.fromEmail), date: iso || guess.date, text});
  base.tags = [...new Set([...(base.tags || []), "e-mail"])];
  if (mail.attachments.length) base.notes = `E-mail with ${mail.attachments.length} attachment${mail.attachments.length > 1 ? "s" : ""}: ${mail.attachments.map(a => a.name).join(", ")}.`;
  const filed = !!hist;
  if (!catsFor(base.space).includes(base.category)) base.category = catsFor(base.space).at(-1);
  const meta = {id: uid(), ...base, filed, autoFiled: filed ? "history" : "", weak: false, textSource: "email", ocrConf: null,
    pages: 1, hash, mime: "application/pdf", source: "email", originalName: f.name || "", createdAt: Date.now(), handled: false};
  it.note = "Making a PDF of the e-mail"; renderImport();
  const pdf = await mailPdf(mail, meta.title);
  it.note = "Uploading to OneDrive"; renderImport();
  const up = await uploadFile(folderOf(meta), fileNameOf(meta, ".pdf"), pdf);
  Object.assign(meta, {itemId: up.id, name: up.name, path: `${folderOf(meta)}/${up.name}`, size: up.size, webUrl: up.webUrl, fileVersion: 1});
  // keep the original e-mail file for Outlook
  try {
    const ext = /\.eml$/i.test(f.name || "") ? ".eml" : ".msg";
    const oname = fileNameOf(meta, ext);
    const orig = it.remote ? await moveItem(it.remote.id, "E-mail originals", oname) : await uploadFile("E-mail originals", oname, new Blob([buf], {type: f.type || "application/octet-stream"}));
    meta.originalItemId = orig.id; meta.originalPath = `E-mail originals/${orig.name}`;
  } catch (e) { console.warn("original e-mail not kept", e); }
  putDoc(meta);
  const parent = {title: meta.title, sender: meta.sender, date: meta.date, filed, space: meta.space, category: meta.category};
  let added = 0;
  for (const a of mail.attachments) {
    const type = a.type || (/\.pdf$/i.test(a.name) ? "application/pdf" : /\.(jpe?g)$/i.test(a.name) ? "image/jpeg" : /\.png$/i.test(a.name) ? "image/png" : "");
    const af = new File([a.data], a.name, {type});
    if (!(isPdfFile(af) || /^image\//.test(type))) continue;
    if (/^image\//.test(type) && a.data.length < 25000) continue; // logos and signatures
    queue.push({file: af, name: `${a.name} (from e-mail)`, status: "waiting", note: "", parent}); added++;
  }
  it.status = "done"; it.docId = meta.id;
  it.note = (filed ? `${SPACE_LABEL[meta.space]} · ${meta.category}` : "To file") + (added ? ` · ${added} attachment${added > 1 ? "s" : ""} next` : "");
}

async function importCombined(it) {
  const files = it.combine, blobs = [], texts = []; let cs = 0;
  const parts = [];
  for (let i = 0; i < files.length; i++) {
    const buf = await files[i].arrayBuffer(); parts.push(new Uint8Array(buf));
    let bm;
    try { bm = await createImageBitmap(new Blob([buf], {type: files[i].type}), {imageOrientation: "from-image"}); }
    catch { throw {userMsg: `${files[i].name} could not be opened.`}; }
    const c = renderPage({bitmap: bm, rot: 0, crop: null}, "original", {enhance: false, maxSide: 2200});
    blobs.push(await toBlob(c, .82));
    ocrProgress = m => { if (m.status === "recognizing text") { it.note = `Reading text, page ${i + 1} of ${files.length} · ${Math.round(m.progress * 100)}%`; renderImport(); } };
    const r = await ocrCanvas(c); texts.push(r.text); cs += r.conf;
  }
  ocrProgress = null;
  const all = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0; for (const p of parts) { all.set(p, o); o += p.length; }
  const hash = await sha256(all.buffer);
  if (hash && docs().find(d => d.hash === hash)) { it.status = "skipped"; it.note = "Already in your vault"; return; }
  let text = texts.join("\n\n").trim(), conf = Math.round(cs / files.length), source = "ocr", weak = isWeak(text, conf), aiData = null;
  if (weak && ai.on && ai.auto) {
    it.note = "Hard to read, asking AI"; renderImport();
    try { aiData = await aiRead(blobs); text = aiData.text || text; source = "ai"; weak = false; } catch {}
  }
  const guess = guessFields(text), hist = suggestFromHistory(text);
  const base = fieldsFrom({...guess, ...(hist || {}), ...(aiData ? Object.fromEntries(Object.entries(aiData).filter(([, v]) => v !== "" && v !== null)) : {}), text});
  if (!base.title) base.title = base.sender || files[0].name.replace(/\.[^.]+$/, "");
  const filed = !!(hist || (aiData && aiData.category));
  if (!catsFor(base.space).includes(base.category)) base.category = catsFor(base.space).at(-1);
  const meta = {id: uid(), ...base, filed, autoFiled: filed ? (hist ? "history" : "ai") : "", weak, textSource: source, ocrConf: conf,
    pages: files.length, hash, mime: "application/pdf", source: "import", originalName: files.map(f => f.name).join(", ").slice(0, 200), createdAt: Date.now(), handled: false};
  it.note = "Making the PDF"; renderImport();
  const pdf = await buildScanPdf(blobs, source === "ai" ? [text] : texts, meta);
  it.note = "Uploading to OneDrive"; renderImport();
  const up = await uploadFile(folderOf(meta), fileNameOf(meta, ".pdf"), pdf);
  Object.assign(meta, {itemId: up.id, name: up.name, path: `${folderOf(meta)}/${up.name}`, size: up.size, webUrl: up.webUrl, fileVersion: 1});
  putDoc(meta);
  it.status = "done"; it.docId = meta.id;
  it.note = filed ? `${SPACE_LABEL[meta.space]} · ${meta.category}` : "To file";
}
async function runQueue() {
  if (queueRunning) return; queueRunning = true;
  for (let it; (it = queue.find(q => q.status === "waiting"));) {
    it.status = "working"; it.note = "Opening"; renderImport();
    try { await importOne(it); }
    catch (e) { console.error(e); it.status = "failed"; it.note = errText(e); }
    it.file = null; it.combine = null; ocrProgress = null; renderImport();
  }
  queueRunning = false; renderImport();
}
let impRaf = 0;
function renderImport() {
  if (impRaf) return;
  impRaf = requestAnimationFrame(() => {
    impRaf = 0;
    const box = $("#importBox");
    if (!queue.length) { box.innerHTML = ""; return; }
    const finished = queue.filter(q => ["done", "failed", "skipped"].includes(q.status)).length;
    const active = finished < queue.length, ok = queue.filter(q => q.status === "done").length;
    const toFile = queue.filter(q => q.status === "done" && q.docId && getDoc(q.docId) && !getDoc(q.docId).filed).length;
    box.innerHTML = `<div class="card imp"><div class="imphead"><b>${active ? `Adding files · ${finished} of ${queue.length} done` : `Added ${ok} of ${queue.length} file${queue.length > 1 ? "s" : ""}`}</b>
      <span class="btnrow">${!active && toFile ? `<button class="btn primary" data-act="filenow">${ICON.inbox}File ${toFile} now</button>` : ""}${active ? "" : `<button class="btn ghost" data-act="clearq">Close</button>`}</span></div>
      <div class="pbar"><span style="width:${(finished / queue.length * 100).toFixed(1)}%"></span></div>
      <ul class="implist">${queue.map(q => `<li class="${q.status}"><span class="fn" title="${esc(q.name)}">${esc(q.name)}</span><span class="st">${q.status === "done" && q.docId ? `<button data-open="${esc(q.docId)}">${esc(q.note)}</button>` : esc(q.status === "waiting" ? "Waiting" : q.note)}</span></li>`).join("")}</ul></div>`;
  });
}
$("#importBox").addEventListener("click", e => {
  const o = e.target.closest("[data-open]"); if (o) return openDetail(o.dataset.open);
  if (e.target.closest('[data-act="clearq"]')) { queue.length = 0; renderImport(); }
  if (e.target.closest('[data-act="filenow"]')) { const n = nextToFile(); if (n) openEdit(getDoc(n), {filing: true}); }
});
async function scanInbox() {
  if (!account || !state.online) return;
  let list;
  try { list = await g("GET", `${itemRef(INBOX)}/children?$top=200&select=id,name,file,size,webUrl`); }
  catch (e) { return; }
  const known = new Set(docs().map(d => d.itemId)), queued = new Set(queue.map(q => q.remote && q.remote.id));
  const fresh = (list.value || []).filter(x => x.file && !known.has(x.id) && !queued.has(x.id) && (isPdfFile({name: x.name, type: x.file.mimeType}) || isMailFile({name: x.name, type: x.file.mimeType}) || /^image\//.test(x.file.mimeType || "")));
  if (!fresh.length) return;
  for (const x of fresh) queue.push({remote: x, name: x.name, status: "waiting", note: ""});
  renderImport(); runQueue();
}
async function importOne(it) {
  if (it.combine) return importCombined(it);
  let f = it.file;
  if (it.remote) {
    it.note = "Downloading from OneDrive"; renderImport();
    const blob = await downloadItem(it.remote.id);
    f = new File([blob], it.remote.name, {type: it.remote.file.mimeType || blob.type});
  }
  const buf = await f.arrayBuffer();
  const hash = await sha256(buf);
  const dup = hash && docs().find(d => d.hash === hash);
  if (dup) { it.status = "skipped"; it.note = "Already in your vault"; return; }
  if (isMailFile(f)) return importEmail(it, f, buf, hash);
  const pdfFile = isPdfFile(f);
  let text = "", source = "", conf = null, pages = 1; const aiBlobs = [];
  const ocrTick = (i, n) => m => { if (m.status === "recognizing text") { it.note = `Reading text${n > 1 ? `, page ${i} of ${n}` : ""} · ${Math.round(m.progress * 100)}%`; renderImport(); } else if (/loading|initializing/.test(m.status)) { it.note = "Loading the text reader…"; renderImport(); } };
  if (pdfFile) {
    it.note = "Reading PDF"; renderImport();
    const pdf = await openPdf(buf.slice(0)); pages = pdf.numPages;
    text = await pdfText(pdf);
    if (text.replace(/\s/g, "").length >= 60) source = "pdf";
    else {
      const n = Math.min(pages, 20), texts = []; let cs = 0;
      for (let i = 1; i <= n; i++) {
        const c = await pdfPageCanvas(pdf, i);
        ocrProgress = ocrTick(i, n);
        const r = await ocrCanvas(c); texts.push(r.text); cs += r.conf;
        if (aiBlobs.length < 5) aiBlobs.push(await toBlob(scaled(c, 1800), .85));
        c.width = c.height = 0;
      }
      text = texts.join("\n\n").trim(); conf = Math.round(cs / n); source = "ocr";
    }
    await pdf.destroy();
  } else {
    let bm;
    try { bm = await createImageBitmap(new Blob([buf], {type: f.type}), {imageOrientation: "from-image"}); }
    catch { throw {userMsg: "This image could not be opened."}; }
    const c = renderPage({bitmap: bm, rot: 0, crop: null}, "original", {enhance: false, maxSide: 2600});
    ocrProgress = ocrTick(1, 1);
    const r = await ocrCanvas(c); text = r.text; conf = r.conf; source = "ocr";
    aiBlobs.push(await toBlob(scaled(c, 1800), .85));
  }
  ocrProgress = null;
  let weak = source === "ocr" && isWeak(text, conf), aiData = null;
  if (weak && ai.on && ai.auto) {
    it.note = "Hard to read, asking AI"; renderImport();
    try { aiData = await aiRead(aiBlobs); text = aiData.text || text; source = "ai"; weak = false; } catch (e) { it.aiErr = e.userMsg; }
  }
  const guess = guessFields(text), hist = suggestFromHistory(text);
  const base = fieldsFrom({...guess, ...(hist || {}), ...(aiData ? Object.fromEntries(Object.entries(aiData).filter(([, v]) => v !== "" && v !== null)) : {}), text});
  if (!base.title) base.title = (aiData && aiData.title) || f.name.replace(/\.[^.]+$/, "").replace(/[_]+/g, " ").trim();
  const par = it.parent;
  if (par) {
    base.notes = `Attachment of the e-mail “${par.title}”${par.sender ? ` from ${par.sender}` : ""}${par.date ? `, ${fmtDate(par.date)}` : ""}.`;
    base.tags = [...new Set([...(base.tags || []), "e-mail attachment"])];
    if (!base.date) base.date = par.date;
    if (!hist && par.filed) { base.space = par.space; base.category = par.category; }
  }
  const filed = !!(hist || (aiData && aiData.category) || (par && par.filed));
  if (!catsFor(base.space).includes(base.category)) base.category = catsFor(base.space).at(-1);
  const meta = {id: uid(), ...base, filed, autoFiled: filed ? (hist ? "history" : (aiData && aiData.category) ? "ai" : "email") : "", weak, textSource: source, ocrConf: conf,
    pages, hash, mime: f.type || (pdfFile ? "application/pdf" : "image/jpeg"), source: "import", originalName: f.name || "", createdAt: Date.now(), handled: false};
  const ext = pdfFile ? ".pdf" : (extOf(f.name) || ".jpg");
  let up;
  if (it.remote) {
    meta.source = "inbox";
    it.note = "Filing in OneDrive"; renderImport();
    up = filed ? await moveItem(it.remote.id, folderOf(meta), fileNameOf(meta, ext)) : it.remote;
  } else {
    it.note = "Uploading to OneDrive"; renderImport();
    up = await uploadFile(folderOf(meta), fileNameOf(meta, ext), new Blob([buf], {type: meta.mime}));
  }
  Object.assign(meta, {itemId: up.id, name: up.name, path: `${folderOf(meta)}/${up.name}`, size: up.size, webUrl: up.webUrl, fileVersion: 1});
  putDoc(meta);
  it.status = "done"; it.docId = meta.id;
  it.note = filed ? `${SPACE_LABEL[meta.space]} · ${meta.category}` : weak ? "To file · hard to read" : "To file";
}
$("#btnAdd").addEventListener("click", () => $("#fileIn").click());
$("#fileIn").addEventListener("change", e => { const f = [...e.target.files]; e.target.value = ""; enqueueFiles(f); });
{
  let depth = 0;
  const hasFiles = e => [...(e.dataTransfer?.types || [])].includes("Files");
  const ov = $("#dropOverlay");
  window.addEventListener("dragenter", e => { if (!hasFiles(e) || $("#app").hidden) return; e.preventDefault(); depth++; ov.hidden = false; });
  window.addEventListener("dragover", e => { if (hasFiles(e)) { e.preventDefault(); e.dataTransfer.dropEffect = "copy"; } });
  window.addEventListener("dragleave", e => { if (!hasFiles(e)) return; depth = Math.max(0, depth - 1); if (!depth) ov.hidden = true; });
  window.addEventListener("drop", e => { if (!hasFiles(e)) return; e.preventDefault(); depth = 0; ov.hidden = true; if (!$("#app").hidden) enqueueFiles([...e.dataTransfer.files]); });
  window.addEventListener("paste", e => {
    if ($("#app").hidden || (e.target.closest && e.target.closest("input,textarea"))) return;
    let files = [...(e.clipboardData?.files || [])];
    if (!files.length) files = [...(e.clipboardData?.items || [])].filter(i => i.kind === "file").map(i => i.getAsFile()).filter(Boolean);
    if (!files.length) return;
    e.preventDefault();
    const stamp = new Date().toLocaleString("sv-SE").slice(0, 16).replace(":", ".");
    enqueueFiles(files.map((f, i) => /^image\//.test(f.type) && (!f.name || /^image\.\w+$/i.test(f.name))
      ? new File([f], `Screenshot ${stamp}${files.length > 1 ? ` (${i + 1})` : ""}.${(f.type.split("/")[1] || "png").replace("jpeg", "jpg")}`, {type: f.type}) : f));
  });
}
const nextToFile = (except) => docs().filter(d => !d.filed && d.id !== except).sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0))[0]?.id;

/* ================= zoom ================= */
let zoomEl = null;
function closeZoom() { if (zoomEl) { zoomEl.remove(); zoomEl = null; } }
function openZoom(src) {
  closeZoom();
  const ov = document.createElement("div"); ov.className = "zoomov"; zoomEl = ov;
  ov.innerHTML = `<img alt="Zoomed page" draggable="false"><button class="iconbtn zoomclose" aria-label="Close">${ICON.x}</button><span class="zoomhint">Pinch, scroll or double-tap to zoom</span>`;
  const img = ov.querySelector("img"); img.src = src;
  document.body.appendChild(ov);
  let s = 1, x = 0, y = 0, start = null, lastTap = 0;
  const pts = new Map();
  const apply = () => { img.style.transform = `translate(${x}px, ${y}px) scale(${s})`; };
  const zoomAt = (cx, cy, ns) => {
    ns = Math.min(8, Math.max(1, ns));
    const ox = cx - window.innerWidth / 2, oy = cy - window.innerHeight / 2;
    x = ox - (ox - x) * (ns / s); y = oy - (oy - y) * (ns / s); s = ns;
    if (s === 1) { x = 0; y = 0; }
    apply();
  };
  const snapshot = () => ({s, x, y, pts: new Map([...pts].map(([k, v]) => [k, {...v}]))});
  ov.addEventListener("pointerdown", e => {
    if (e.target.closest(".zoomclose")) return;
    ov.setPointerCapture(e.pointerId); pts.set(e.pointerId, {x: e.clientX, y: e.clientY}); start = snapshot();
  });
  ov.addEventListener("pointermove", e => {
    if (!pts.has(e.pointerId) || !start) return;
    pts.set(e.pointerId, {x: e.clientX, y: e.clientY});
    const now = [...pts.values()], was = [...start.pts.values()];
    if (now.length === 1 && was.length === 1) { x = start.x + now[0].x - was[0].x; y = start.y + now[0].y - was[0].y; apply(); }
    else if (now.length >= 2 && was.length >= 2) {
      const d = (a, b) => Math.hypot(a.x - b.x, a.y - b.y);
      const mid = (a, b) => ({x: (a.x + b.x) / 2, y: (a.y + b.y) / 2});
      const m0 = mid(was[0], was[1]), m1 = mid(now[0], now[1]);
      s = start.s; x = start.x + m1.x - m0.x; y = start.y + m1.y - m0.y;
      zoomAt(m1.x, m1.y, start.s * d(now[0], now[1]) / Math.max(1, d(was[0], was[1])));
    }
  });
  const up = e => {
    const p0 = start && start.pts.get(e.pointerId), p1 = pts.get(e.pointerId);
    pts.delete(e.pointerId); start = pts.size ? snapshot() : null;
    if (p0 && p1 && !pts.size && Math.hypot(p1.x - p0.x, p1.y - p0.y) < 8) {
      const t = Date.now();
      if (t - lastTap < 300) { zoomAt(p1.x, p1.y, s > 1 ? 1 : 2.5); lastTap = 0; } else lastTap = t;
    }
  };
  ov.addEventListener("pointerup", up); ov.addEventListener("pointercancel", up);
  ov.addEventListener("wheel", e => { e.preventDefault(); zoomAt(e.clientX, e.clientY, s * Math.exp(-e.deltaY / 300)); }, {passive: false});
  ov.querySelector(".zoomclose").addEventListener("click", closeZoom);
  return img;
}
panel.addEventListener("click", async e => {
  const t = e.target.closest(".viewer img, .viewer canvas, .pg img");
  if (!t || e.target.closest("[data-rot]")) return;
  e.stopPropagation();
  const img = openZoom(t.tagName === "CANVAS" ? t.toDataURL("image/jpeg", .92) : t.src);
  const fig = t.closest(".vpage"), d = panel.dataset.docId && getDoc(panel.dataset.docId);
  if (t.tagName === "CANVAS" && fig && d && isPdfDoc(d)) {
    try {
      const pdf = await openPdf(await (await getFileBlob(d)).arrayBuffer());
      const c = await pdfPageCanvas(pdf, Number(fig.dataset.page) + 1, 3200); pdf.destroy();
      if (img.isConnected) img.src = URL.createObjectURL(await toBlob(c, .9));
    } catch {}
  }
}, true);
document.addEventListener("keydown", e => { if (e.key === "Escape" && zoomEl) closeZoom(); });
// Used by the Android app's back button: close zoom, then an open screen.
window.vaultBack = () => {
  if (zoomEl) { closeZoom(); return true; }
  if (!panel.hidden) { const b = panel.querySelector('.phead [data-act="close"]'); if (b) { b.click(); return true; } }
  return false;
};

/* ================= Android app updates ================= */
const APK_URL = "https://github.com/jacovanstraten-glitch/doc-vault/releases/latest/download/DocumentVault.apk";
async function checkAndroidUpdate() {
  const isApp = /; wv\)/.test(navigator.userAgent) || !!window.VaultAndroid;
  if (!isApp) return;
  let installed = 0;
  try { installed = window.VaultAndroid && window.VaultAndroid.versionCode ? Number(window.VaultAndroid.versionCode()) : 0; } catch {}
  let latest = 0;
  try {
    const r = await fetch("https://api.github.com/repos/jacovanstraten-glitch/doc-vault/releases/latest", {cache: "no-store"});
    if (r.ok) latest = Number(((await r.json()).tag_name || "").replace(/\D+/g, "")) || 0;
  } catch {}
  if (latest && installed < latest) {
    $("#notice").innerHTML = `<div class="notice update"><span>A new version of the Document Vault app is ready${installed ? "" : ", with sharing and saving to Downloads"}.</span><a class="btn primary" href="${APK_URL}">Install update</a></div>`;
  }
}

/* ================= download, share, rotate ================= */
const bridge = () => (window.VaultAndroid && typeof window.VaultAndroid.shareFile === "function") ? window.VaultAndroid : null;
const inOldAndroidApp = () => !bridge() && /; wv\)/.test(navigator.userAgent);
const canDownload = () => !!bridge() || !inOldAndroidApp();
function canShare() {
  if (bridge()) return true;
  try { return !!(navigator.canShare && navigator.canShare({files: [new File(["x"], "x.pdf", {type: "application/pdf"})]})); } catch { return false; }
}
async function docFile(d) {
  const blob = await getFileBlob(d);
  const name = d.name || fileNameOf(d, isPdfDoc(d) ? ".pdf" : ".jpg");
  return new File([blob], name, {type: d.mime || blob.type || "application/octet-stream"});
}
async function busy(b, label, fn) {
  const old = b.innerHTML; b.disabled = true; b.textContent = label;
  try { await fn(); } catch (e) { if (!(e && e.name === "AbortError")) { console.error(e); toast(e.userMsg || errText(e)); } }
  finally { b.disabled = false; b.innerHTML = old; }
}
async function saveBlob(f) {
  const br = bridge();
  if (br) { br.saveFile(await blobToB64(f), f.name, f.type); return; }
  const u = URL.createObjectURL(f), a = document.createElement("a");
  a.href = u; a.download = f.name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(u), 60000);
}
async function shareBlob(f, title) {
  const br = bridge();
  if (br) { br.shareFile(await blobToB64(f), f.name, f.type, title || f.name); return; }
  await navigator.share({files: [f], title: title || f.name});
}
function downloadDoc(d, b) { return busy(b, "Preparing…", async () => saveBlob(await docFile(d))); }
function shareDoc(d, b) { return busy(b, "Preparing…", async () => shareBlob(await docFile(d), d.title)); }

/* ---- create a new document from a file (split, cut out, merge) ---- */
async function createDoc(blob, fields, ext = ".pdf") {
  const meta = {id: uid(), ...fields, createdAt: Date.now(), handled: false, hash: ""};
  if (!catsFor(meta.space).includes(meta.category)) meta.category = catsFor(meta.space).at(-1);
  const up = await uploadFile(folderOf(meta), fileNameOf(meta, ext), blob);
  Object.assign(meta, {itemId: up.id, name: up.name, path: `${folderOf(meta)}/${up.name}`, size: up.size, webUrl: up.webUrl, fileVersion: 1});
  putDoc(meta);
  return meta;
}
async function deleteDoc(d) {
  try { await g("DELETE", `/me/drive/items/${d.itemId}`); } catch (err) { if (err.status !== 404) throw err; }
  if (d.originalItemId) { try { await g("DELETE", `/me/drive/items/${d.originalItemId}`); } catch {} }
  putDoc({id: d.id, deleted: true, createdAt: d.createdAt});
}
const inherit = d => ({space: d.space, category: d.category, filed: d.filed, sender: d.sender, date: d.date, reference: "", amount: null, currency: d.currency || "EUR",
  dueDate: "", tags: d.tags || [], notes: "", weak: false, autoFiled: ""});

/* ---- export for the accountant ---- */
function exportLabel() {
  const {f} = parseQuery(state.q);
  const parts = [];
  if (state.space !== "all") parts.push(SPACE_LABEL[state.space]);
  if (state.cat && !state.cat.startsWith("__")) parts.push(state.cat);
  const y = f.year || state.year; if (y) parts.push(y);
  return parts.length ? parts.join(" ") : `Documents ${new Date().toISOString().slice(0, 10)}`;
}
function openExport(list, label) {
  if (!list.length) return toast("There is nothing to export.");
  list = [...list].sort((a, b) => docDate(a).localeCompare(docDate(b)));
  const byCat = new Map();
  for (const d of list) { const k = d.filed ? `${SPACE_LABEL[d.space]} · ${d.category}` : "To file"; byCat.set(k, [...(byCat.get(k) || []), d]); }
  showPanel("Export", `
    <div class="field"><label for="x_name">Name of the export</label><input id="x_name" value="${esc(label)}"></div>
    <div class="card"><h3>${list.length} document${list.length > 1 ? "s" : ""}${sumAmounts(list) ? ` · ${esc(sumAmounts(list))}` : ""}</h3>
      <table class="xtable"><tbody>${[...byCat.entries()].map(([k, v]) => `<tr><td>${esc(k)}</td><td class="mono">${v.length}</td><td class="mono">${esc(sumAmounts(v))}</td></tr>`).join("")}</tbody></table>
      <p class="hint">You get a ZIP with every file in a folder per category, and <span class="mono">Overview.xlsx</span>: one row per document with date, sender, reference and amount, plus a sheet with totals per category.</p></div>
    <label class="toggle"><input type="checkbox" id="x_od" checked>Also save the ZIP in OneDrive › ${esc(BASE)} › Exports</label>
    <div class="btnrow"><button class="btn primary grow" data-act="make">${ICON.download}Make export</button></div>
    <div id="x_out"></div>`);
  panel.onclick = async e => {
    const b = e.target.closest("[data-act]"); if (!b) return;
    if (b.dataset.act === "close") return closePanel();
    if (b.dataset.act === "make") {
      const name = safeSeg($("#x_name", panel).value || label, 80), toOd = $("#x_od", panel).checked;
      b.disabled = true;
      try {
        const zip = await buildExport(list, name, t => { b.textContent = t; });
        const file = new File([zip], `${name}.zip`, {type: "application/zip"});
        let where = "";
        if (toOd) { b.textContent = "Saving to OneDrive…"; const it = await uploadFile("Exports", file.name, file, {replace: true}); where = `Saved in OneDrive › ${BASE} › Exports › ${it.name}.`; }
        b.innerHTML = `${ICON.check}Export ready`;
        $("#x_out", panel).innerHTML = `<div class="flag info">${ICON.check}<span>${esc(where || "The export is ready.")}</span></div>
          <div class="btnrow">${canDownload() ? `<button class="btn primary grow" data-act="xdl">${ICON.download}Download ZIP</button>` : ""}${canShare() ? `<button class="btn grow" data-act="xshare">${ICON.share}Share ZIP</button>` : ""}</div>`;
        panel.dataset.xfile = "1"; panel._xfile = file;
      } catch (err) { console.error(err); b.disabled = false; b.innerHTML = `${ICON.download}Make export`; toast(err.userMsg || errText(err)); }
    }
    if (b.dataset.act === "xdl") return busy(b, "Preparing…", () => saveBlob(panel._xfile));
    if (b.dataset.act === "xshare") return busy(b, "Preparing…", () => shareBlob(panel._xfile, panel._xfile.name));
  };
}
async function buildExport(list, name, progress) {
  await loadScript("vendor/jszip.min.js"); await loadScript("vendor/xlsx.full.min.js");
  const zip = new JSZip(), root = zip.folder(name), used = new Set(), rows = [];
  for (let i = 0; i < list.length; i++) {
    const d = list[i];
    progress(`Collecting ${i + 1} of ${list.length}…`);
    const folder = d.filed ? safeSeg(`${SPACE_LABEL[d.space]} - ${d.category}`, 80) : "To file";
    let fname = d.name || fileNameOf(d, isPdfDoc(d) ? ".pdf" : ".jpg"), k = 2;
    while (used.has(`${folder}/${fname}`)) fname = (d.name || "document").replace(/(\.[^.]+)?$/, m => ` (${k++})${m}`);
    used.add(`${folder}/${fname}`);
    root.file(`${folder}/${fname}`, await getFileBlob(d));
    const amt = d.amount === null || d.amount === undefined || d.amount === "" || isNaN(Number(d.amount)) ? null : Number(d.amount);
    rows.push([d.date ? new Date(d.date + "T12:00:00") : null, d.filed ? SPACE_LABEL[d.space] : "", d.filed ? d.category : "To file", d.sender || "", d.title || "",
      d.reference || "", amt, d.currency || (amt !== null ? "EUR" : ""), d.dueDate ? new Date(d.dueDate + "T12:00:00") : null, `${folder}/${fname}`]);
  }
  progress("Making the overview…");
  const head = ["Date", "Space", "Category", "From", "Title", "Reference", "Amount", "Currency", "Due / ends", "File in ZIP"];
  const ws = XLSX.utils.aoa_to_sheet([head, ...rows], {cellDates: true});
  ws["!cols"] = [12, 10, 26, 26, 40, 18, 12, 9, 12, 60].map(w => ({wch: w}));
  for (let r = 2; r <= rows.length + 1; r++) {
    if (ws[`A${r}`]) ws[`A${r}`].z = "dd-mm-yyyy";
    if (ws[`I${r}`]) ws[`I${r}`].z = "dd-mm-yyyy";
    if (ws[`G${r}`]) ws[`G${r}`].z = "#,##0.00";
  }
  ws["!autofilter"] = {ref: `A1:J${rows.length + 1}`};
  const cats = [...new Set(rows.map(r => r[2]))];
  const n = rows.length + 1;
  const tot = [["Category", "Documents", "Total EUR"], ...cats.map((c, i) => [c, {f: `COUNTIF(Documents!C2:C${n},A${i + 2})`}, {f: `SUMIFS(Documents!G2:G${n},Documents!C2:C${n},A${i + 2},Documents!H2:H${n},"EUR")`, z: "#,##0.00"}]),
    ["Total", {f: `SUM(B2:B${cats.length + 1})`}, {f: `SUM(C2:C${cats.length + 1})`, z: "#,##0.00"}]];
  const ws2 = XLSX.utils.aoa_to_sheet(tot); ws2["!cols"] = [{wch: 34}, {wch: 12}, {wch: 14}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Documents"); XLSX.utils.book_append_sheet(wb, ws2, "Totals");
  root.file("Overview.xlsx", XLSX.write(wb, {bookType: "xlsx", type: "array"}));
  progress("Packing the ZIP…");
  return zip.generateAsync({type: "blob", compression: "STORE"});
}

/* ---- calendar reminders ---- */
const isWarranty = d => /warrant|garantie/i.test([d.category, d.title, d.notes, (d.tags || []).join(" ")].join(" "));
const calTitle = d => `${isWarranty(d) ? "Warranty ends" : "Payment due"}: ${d.title || "document"}`;
const calDetails = d => [d.sender && `From: ${d.sender}`, d.amount !== null && d.amount !== undefined && d.amount !== "" && `Amount: ${fmtAmount(d.amount, d.currency)}`,
  d.reference && `Reference: ${d.reference}`, d.path && `File: OneDrive/${BASE}/${d.path}`].filter(Boolean).join("\n");
const ymd = (iso, add = 0) => { const t = new Date(iso + "T12:00:00"); t.setDate(t.getDate() + add); return t.toISOString().slice(0, 10).replace(/-/g, ""); };
const googleCalUrl = d => `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent(calTitle(d))}&dates=${ymd(d.dueDate)}/${ymd(d.dueDate, 1)}&details=${encodeURIComponent(calDetails(d))}`;
function icsFor(list) {
  const e = t => String(t).replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/([,;])/g, "\\$1");
  const stamp = new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d+/, "");
  const ev = list.map(d => ["BEGIN:VEVENT", `UID:dv-${d.id}@document-vault`, `DTSTAMP:${stamp}`, `DTSTART;VALUE=DATE:${ymd(d.dueDate)}`, `DTEND;VALUE=DATE:${ymd(d.dueDate, 1)}`,
    `SUMMARY:${e(calTitle(d))}`, `DESCRIPTION:${e(calDetails(d))}`, "TRANSP:TRANSPARENT",
    "BEGIN:VALARM", `TRIGGER:-P${isWarranty(d) ? 30 : 7}D`, "ACTION:DISPLAY", `DESCRIPTION:${e(calTitle(d))}`, "END:VALARM", "END:VEVENT"].join("\r\n"));
  return new File([["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//van Straten//Document Vault//EN", "CALSCALE:GREGORIAN", ...ev, "END:VCALENDAR"].join("\r\n") + "\r\n"],
    list.length === 1 ? `${safeSeg(calTitle(list[0]), 60)}.ics` : "Document Vault reminders.ics", {type: "text/calendar"});
}
async function icsOut(list, b) {
  return busy(b, "Preparing…", async () => { const f = icsFor(list); if (bridge() || !canDownload()) return shareBlob(f, f.name); return saveBlob(f); });
}

/* ---- split, cut out, merge ---- */
async function sourcePages(d, maxSide) {
  const blob = await getFileBlob(d);
  if (isPdfDoc(d)) {
    const pdf = await openPdf(await blob.arrayBuffer()), out = [];
    for (let i = 1; i <= Math.min(pdf.numPages, 60); i++) {
      const page = await pdf.getPage(i); let text = "";
      try { text = textFromContent(await page.getTextContent()); } catch {}
      out.push({canvas: await pdfPageCanvas(pdf, i, maxSide), text});
    }
    pdf.destroy(); return out;
  }
  const bm = await createImageBitmap(blob, {imageOrientation: "from-image"});
  return [{canvas: renderPage({bitmap: bm, rot: 0, crop: null}, "original", {enhance: false, maxSide: Math.max(maxSide, 2600)}), text: d.text || ""}];
}
async function openSplit(d) {
  showPanel("Split or cut out", `<p class="hint">Loading pages…</p>`, {docId: d.id});
  let pages;
  try { pages = await sourcePages(d, 2400); } catch (e) { return toast(errText(e)); }
  const cuts = pages.map((_, i) => i < pages.length - 1);
  const thumbs = pages.map(p => scaled(p.canvas, 360).toDataURL("image/jpeg", .8));
  let saved = 0;
  const draw = () => {
    const groups = cuts.filter(Boolean).length + 1;
    showPanel("Split or cut out", `
      ${pages.length > 1 && isPdfDoc(d) ? `<div class="card"><h3>Split into separate documents</h3>
        <p class="hint">Tap the scissors between pages to choose where to cut.</p>
        <div class="splitrow">${thumbs.map((t, i) => `<figure class="sthumb"><img src="${t}" alt="Page ${i + 1}"><figcaption>${i + 1}</figcaption></figure>${i < pages.length - 1 ? `<button class="iconbtn cutbtn${cuts[i] ? " on" : ""}" data-cut="${i}" aria-pressed="${cuts[i]}" aria-label="Cut after page ${i + 1}">✂</button>` : ""}`).join("")}</div>
        <div class="btnrow"><button class="btn primary grow" data-act="dosplit" ${groups > 1 ? "" : "disabled"}>Split into ${groups} documents</button></div></div>` : ""}
      <div class="card"><h3>Cut out parts of a page</h3>
        <p class="hint">For one scan with several receipts: pick the page, draw a box around one receipt and save it as its own document. Repeat for the next one.</p>
        <div class="splitrow">${thumbs.map((t, i) => `<button class="sthumb pick" data-pick="${i}" aria-label="Cut parts out of page ${i + 1}"><img src="${t}" alt=""><span>${i + 1}</span></button>`).join("")}</div>
        ${saved ? `<p class="hint">${saved} part${saved > 1 ? "s" : ""} saved as new documents.</p>` : ""}</div>
      <label class="toggle"><input type="checkbox" id="s_del" ${saved ? "checked" : ""}>Move the original to the OneDrive recycle bin when done</label>
      <div class="btnrow"><button class="btn grow" data-act="done">Done</button></div>`, {docId: d.id});
  };
  draw();
  const finish = async () => {
    if ($("#s_del", panel)?.checked) { try { await deleteDoc(d); toast("Original moved to the recycle bin"); } catch (e) { toast(errText(e)); } }
    closePanel();
  };
  panel.onclick = async e => {
    const c = e.target.closest("[data-cut]"); if (c) { cuts[c.dataset.cut] = !cuts[c.dataset.cut]; return draw(); }
    const pk = e.target.closest("[data-pick]"); if (pk) return cutOut(Number(pk.dataset.pick));
    const b = e.target.closest("[data-act]"); if (!b) return;
    if (b.dataset.act === "close") return openDetail(d.id);
    if (b.dataset.act === "done") return finish();
    if (b.dataset.act === "dosplit") {
      const del = $("#s_del", panel).checked;
      return busy(b, "Splitting…", async () => {
        await loadScript("vendor/pdf-lib.min.js");
        const {PDFDocument} = window.PDFLib;
        const src = await PDFDocument.load(await (await getFileBlob(d)).arrayBuffer(), {ignoreEncryption: true});
        const groups = [[]]; pages.forEach((_, i) => { groups.at(-1).push(i); if (cuts[i]) groups.push([]); });
        const allEmpty = pages.every(p => !p.text.trim());
        let k = 0;
        for (const grp of groups) {
          k++; b.textContent = `Saving part ${k} of ${groups.length}…`;
          const nd = await PDFDocument.create();
          for (const pg of await nd.copyPages(src, grp)) nd.addPage(pg);
          const text = allEmpty ? (d.text || "") : grp.map(i => pages[i].text).join("\n\n").trim();
          const g0 = guessFields(text);
          await createDoc(new Blob([await nd.save()], {type: "application/pdf"}), {...inherit(d), title: `${d.title} (part ${k})`, text, textSource: d.textSource,
            amount: g0.amount ?? null, reference: g0.reference || "", pages: grp.length, mime: "application/pdf", source: "split"});
        }
        if (del) { await deleteDoc(d); await tidyNames(); }
        toast(`Split into ${groups.length} documents`); closePanel();
      });
    }
  };
  function cutOut(i) {
    const src = pages[i].canvas;
    const url = scaled(src, 1100).toDataURL("image/jpeg", .85);
    let box = {x: .05, y: .05, w: .9, h: .4};
    showPanel(`Cut out from page ${i + 1}`, `<p class="hint">Drag the box around one receipt or part, then save it.</p>
      <div class="cropwrap" id="cw"><img src="${url}" alt="Page ${i + 1}" draggable="false"><div class="cropclip"><div class="cropbox" id="cb"></div></div>
      <div id="hd" style="position:absolute"><span class="handle tl" data-h="tl"></span><span class="handle tr" data-h="tr"></span><span class="handle bl" data-h="bl"></span><span class="handle br" data-h="br"></span></div></div>
      <div class="btnrow"><button class="btn primary grow" data-act="savepart">${ICON.check}Save this part as a new document</button><button class="btn grow" data-act="back">Back</button></div>`, {docId: d.id});
    const cw = $("#cw", panel), cb = $("#cb", panel), hd = $("#hd", panel);
    const place = () => { for (const el of [cb, hd]) { el.style.left = box.x * 100 + "%"; el.style.top = box.y * 100 + "%"; el.style.width = box.w * 100 + "%"; el.style.height = box.h * 100 + "%"; } };
    place();
    let drag = null;
    cw.addEventListener("pointerdown", e => { const h = e.target.dataset.h || (e.target === cb ? "move" : null); if (!h) return; e.preventDefault(); cw.setPointerCapture(e.pointerId); drag = {h, sx: e.clientX, sy: e.clientY, b: {...box}}; });
    cw.addEventListener("pointermove", e => {
      if (!drag) return;
      const r = cw.getBoundingClientRect(), dx = (e.clientX - drag.sx) / r.width, dy = (e.clientY - drag.sy) / r.height, b = drag.b, min = .05;
      if (drag.h === "move") box = {...b, x: Math.min(1 - b.w, Math.max(0, b.x + dx)), y: Math.min(1 - b.h, Math.max(0, b.y + dy))};
      else {
        let x1 = b.x, y1 = b.y, x2 = b.x + b.w, y2 = b.y + b.h;
        if (drag.h.includes("l")) x1 = Math.min(x2 - min, Math.max(0, b.x + dx));
        if (drag.h.includes("r")) x2 = Math.max(x1 + min, Math.min(1, b.x + b.w + dx));
        if (drag.h.includes("t")) y1 = Math.min(y2 - min, Math.max(0, b.y + dy));
        if (drag.h.includes("b")) y2 = Math.max(y1 + min, Math.min(1, b.y + b.h + dy));
        box = {x: x1, y: y1, w: x2 - x1, h: y2 - y1};
      }
      place();
    });
    const end = () => drag = null; cw.addEventListener("pointerup", end); cw.addEventListener("pointercancel", end);
    panel.onclick = async e => {
      const b = e.target.closest("[data-act]"); if (!b) return;
      if (b.dataset.act === "close" || b.dataset.act === "back") { draw(); return rebind(); }
      if (b.dataset.act === "savepart") return busy(b, "Reading the text…", async () => {
        const c = document.createElement("canvas");
        c.width = Math.max(1, Math.round(box.w * src.width)); c.height = Math.max(1, Math.round(box.h * src.height));
        c.getContext("2d").drawImage(src, box.x * src.width, box.y * src.height, c.width, c.height, 0, 0, c.width, c.height);
        const blob = await toBlob(c, .88);
        let text = "", conf = null;
        try { const r = await ocrCanvas(c); text = r.text; conf = r.conf; } catch {}
        const g0 = guessFields(text); saved++;
        const fields = {...inherit(d), title: `${d.title} (part ${saved})`, text, textSource: "ocr", ocrConf: conf, weak: isWeak(text, conf),
          sender: d.sender || g0.sender || "", date: g0.date || d.date, amount: g0.amount ?? null, currency: g0.currency || d.currency || "EUR", reference: g0.reference || "",
          pages: 1, mime: "application/pdf", source: "cut"};
        b.textContent = "Saving…";
        await createDoc(await buildScanPdf([blob], [text], fields), fields);
        toast(`Part ${saved} saved`);
        box = {x: .05, y: Math.min(.55, box.y + box.h), w: .9, h: .4}; place();
      });
    };
  }
  const handler = panel.onclick;
  const rebind = () => { panel.onclick = handler; };
}
function openMerge(list) {
  let order = [...list].sort((a, b) => docDate(a).localeCompare(docDate(b)));
  const draw = () => showPanel("Merge into one document", `
    <p class="hint">The documents are combined into one PDF in this order.</p>
    <ol class="mergelist">${order.map((d, i) => `<li><span>${esc(d.title || "Untitled")}<small>${esc([d.sender, fmtDate(d.date)].filter(Boolean).join(" · "))}</small></span>
      <button class="iconbtn" data-up="${i}" aria-label="Move up" ${i ? "" : "disabled"}>↑</button><button class="iconbtn" data-down="${i}" aria-label="Move down" ${i < order.length - 1 ? "" : "disabled"}>↓</button></li>`).join("")}</ol>
    <div class="field"><label for="m_title">Title of the new document</label><input id="m_title" value="${esc(order[0].title || "")}"></div>
    <label class="toggle"><input type="checkbox" id="m_del" checked>Move the ${order.length} originals to the OneDrive recycle bin</label>
    <div class="btnrow"><button class="btn primary grow" data-act="domerge">Merge ${order.length} documents</button></div>`);
  draw();
  panel.onclick = async e => {
    const u = e.target.closest("[data-up]"), dn = e.target.closest("[data-down]");
    if (u) { const i = +u.dataset.up; [order[i - 1], order[i]] = [order[i], order[i - 1]]; return draw(); }
    if (dn) { const i = +dn.dataset.down; [order[i + 1], order[i]] = [order[i], order[i + 1]]; return draw(); }
    const b = e.target.closest("[data-act]"); if (!b) return;
    if (b.dataset.act === "close") return closePanel();
    if (b.dataset.act !== "domerge") return;
    const title = $("#m_title", panel).value.trim() || order[0].title, del = $("#m_del", panel).checked;
    return busy(b, "Merging…", async () => {
      await loadScript("vendor/pdf-lib.min.js");
      const {PDFDocument} = window.PDFLib;
      const out = await PDFDocument.create();
      for (let i = 0; i < order.length; i++) {
        const d = order[i]; b.textContent = `Adding ${i + 1} of ${order.length}…`;
        const blob = await getFileBlob(d);
        if (isPdfDoc(d)) {
          const src = await PDFDocument.load(await blob.arrayBuffer(), {ignoreEncryption: true});
          for (const pg of await out.copyPages(src, src.getPageIndices())) out.addPage(pg);
        } else {
          let bytes = new Uint8Array(await blob.arrayBuffer()), img;
          if (/png/.test(blob.type || d.mime || "")) img = await out.embedPng(bytes);
          else {
            if (!/jpe?g/.test(blob.type || d.mime || "")) { const bm = await createImageBitmap(blob); const c = document.createElement("canvas"); c.width = bm.width; c.height = bm.height; c.getContext("2d").drawImage(bm, 0, 0); bytes = new Uint8Array(await (await toBlob(c, .9)).arrayBuffer()); }
            img = await out.embedJpg(bytes);
          }
          const land = img.width > img.height, W = land ? 841.89 : 595.28, H = land ? 595.28 : 841.89, m = 18, sc = Math.min((W - 2 * m) / img.width, (H - 2 * m) / img.height);
          out.addPage([W, H]).drawImage(img, {x: (W - img.width * sc) / 2, y: (H - img.height * sc) / 2, width: img.width * sc, height: img.height * sc});
        }
      }
      const first = order[0];
      const fields = {...inherit(first), title, date: order.map(d => d.date).filter(Boolean).sort()[0] || "", amount: first.amount ?? null, currency: first.currency || "EUR",
        reference: first.reference || "", dueDate: order.map(d => d.dueDate).filter(Boolean).sort()[0] || "",
        text: order.map(d => d.text || "").join("\n\n").trim(), textSource: "merged", weak: order.some(d => d.weak),
        tags: [...new Set(order.flatMap(d => d.tags || []))], notes: `Merged from: ${order.map(d => d.title).join("; ")}`.slice(0, 500),
        pages: out.getPageCount(), mime: "application/pdf", source: "merge", filed: order.every(d => d.filed) ? first.filed : false};
      const nd = await createDoc(new Blob([await out.save()], {type: "application/pdf"}), fields);
      if (del) { for (const d of order) await deleteDoc(d); await tidyNames(); }
      state.selecting = false; state.sel.clear();
      toast("Merged into one document"); openDetail(nd.id);
    });
  };
}
async function rotateDocPage(d, index, btn) {
  btn.disabled = true; btn.classList.add("spin");
  try {
    const blob = await getFileBlob(d);
    let out;
    if (isPdfDoc(d)) {
      await loadScript("vendor/pdf-lib.min.js");
      const {PDFDocument, degrees} = window.PDFLib;
      const doc = await PDFDocument.load(await blob.arrayBuffer(), {ignoreEncryption: true});
      const page = doc.getPages()[index];
      page.setRotation(degrees(((page.getRotation().angle || 0) + 90) % 360));
      out = new Blob([await doc.save()], {type: "application/pdf"});
    } else {
      const bm = await createImageBitmap(blob, {imageOrientation: "from-image"});
      const c = document.createElement("canvas"); c.width = bm.height; c.height = bm.width;
      const ctx = c.getContext("2d"); ctx.translate(c.width / 2, c.height / 2); ctx.rotate(Math.PI / 2); ctx.drawImage(bm, -bm.width / 2, -bm.height / 2);
      const type = /png/.test(d.mime || "") ? "image/png" : "image/jpeg";
      out = await new Promise(r => c.toBlob(r, type, .9));
    }
    const it = await replaceContent(d.itemId, out);
    const v = (d.fileVersion || 1) + 1;
    fileCache.set(d.itemId + ":" + v, out);
    putDoc({...d, fileVersion: v, size: it && it.size || out.size});
    openDetail(d.id);
    toast("Page rotated");
  } catch (e) {
    console.error(e); btn.disabled = false; btn.classList.remove("spin");
    toast(e && e.message && /encrypt/i.test(e.message) ? "This PDF is protected and cannot be rotated." : errText(e));
  }
}

/* ================= document detail ================= */
const fileCache = new Map();
async function getFileBlob(d) {
  const k = d.itemId + ":" + (d.fileVersion || 0);
  if (fileCache.has(k)) return fileCache.get(k);
  const b = await downloadItem(d.itemId);
  fileCache.set(k, b); if (fileCache.size > 6) fileCache.delete(fileCache.keys().next().value);
  return b;
}
const isPdfDoc = d => (d.mime || "").includes("pdf") || /\.pdf$/i.test(d.name || "");
function pageWrap(node, i, canRotate) {
  const w = document.createElement("figure"); w.className = "vpage"; w.dataset.page = i;
  w.appendChild(node);
  if (canRotate) w.insertAdjacentHTML("beforeend", `<button class="iconbtn rotbtn" data-rot="${i}" aria-label="Rotate page ${i + 1}" title="Rotate page">${ICON.rotate}</button>`);
  return w;
}
async function fillViewer(d, el, {rotate = false} = {}) {
  try {
    const blob = await getFileBlob(d);
    if (!el.isConnected) return;
    if (isPdfDoc(d)) {
      const pdf = await openPdf(await blob.arrayBuffer());
      el.innerHTML = "";
      const n = Math.min(pdf.numPages, 15);
      for (let i = 1; i <= n; i++) { if (!el.isConnected) break; const c = await pdfPageCanvas(pdf, i, 1500); c.setAttribute("aria-label", `Page ${i}`); el.appendChild(pageWrap(c, i - 1, rotate)); }
      if (pdf.numPages > n && el.isConnected) el.insertAdjacentHTML("beforeend", `<span class="loading">${pdf.numPages - n} more pages. Open the file to see all.</span>`);
      pdf.destroy();
    } else {
      const img = new Image(); img.src = URL.createObjectURL(blob); img.alt = "Document";
      el.innerHTML = ""; el.appendChild(pageWrap(img, 0, rotate));
    }
  } catch (e) { if (el.isConnected) el.innerHTML = `<span class="loading">${esc(e.code === "offline" ? "Connect to the internet to view this file." : "The file could not be loaded from OneDrive.")}</span>`; }
}
async function pageImagesFor(d, max = 20) {
  const blob = await getFileBlob(d), out = [];
  if ((d.mime || "").includes("pdf") || /\.pdf$/i.test(d.name || "")) {
    const pdf = await openPdf(await blob.arrayBuffer());
    for (let i = 1; i <= Math.min(pdf.numPages, max); i++) out.push(await pdfPageCanvas(pdf, i));
    pdf.destroy();
  } else {
    const bm = await createImageBitmap(blob, {imageOrientation: "from-image"});
    out.push(renderPage({bitmap: bm, rot: 0, crop: null}, "original", {enhance: false, maxSide: 2600}));
  }
  return out;
}
function openDetail(id, confirmDelete = false) {
  const d = getDoc(id); if (!d) return closePanel();
  const terms = norm(state.q).split(/\s+/).filter(Boolean);
  const due = dueInfo(d), amt = fmtAmount(d.amount, d.currency);
  const row = (k, v) => v ? `<dt>${k}</dt><dd>${v}</dd>` : "";
  showPanel(d.title || "Untitled", `
    ${!d.filed ? `<div class="flag weak">${ICON.inbox}<span>This document is not filed yet.</span><button class="btn" data-act="edit">File it</button></div>` : ""}
    ${d.autoFiled ? `<div class="flag info">${ICON.check}<span>Filed automatically${d.autoFiled === "history" ? " based on earlier documents from this sender" : d.autoFiled === "email" ? " like the e-mail it came with" : " by AI"}.</span><button class="btn" data-act="okfiled">Looks good</button></div>` : ""}
    ${d.weak ? `<div class="flag weak">${ICON.text}<span>The text was hard to read, so search may miss this document.</span>${ai.on ? `<button class="btn" data-act="ai">${ICON.spark}Read with AI</button>` : `<button class="btn" data-act="edit">Add key words</button>`}</div>` : ""}
    <div class="viewer" id="viewer"><span class="loading">Loading from OneDrive…</span></div>
    <div class="tags">${d.filed ? `<span class="pill ${d.space}">${SPACE_LABEL[d.space]}</span><span class="cat">${esc(d.category)}</span>` : ""}${due ? `<span class="due">${due.label}</span>` : ""}${d.handled && d.dueDate ? `<span class="cat">Handled</span>` : ""}</div>
    <dl class="facts">
      ${row("Date", esc(fmtDate(d.date)))}${row("From", highlight(d.sender || "", terms))}${row("Amount", amt ? `<span class="mono">${esc(amt)}</span>` : "")}
      ${row("Due / ends", esc(fmtDate(d.dueDate)))}${row("Reference", d.reference ? `<span class="mono">${highlight(d.reference, terms)}</span>` : "")}
      ${row("Tags", esc((d.tags || []).join(", ")))}${row("Notes", d.notes ? highlight(d.notes, terms) : "")}${row("Pages", String(d.pages || 1))}
      ${row("OneDrive", `<span class="mono">${esc(BASE)}/${esc(d.path || "")}</span>`)}${row("Text from", esc({email: "E-mail", pdf: "PDF text", ocr: `Text recognition${d.ocrConf ? ` (${d.ocrConf}% sure)` : ""}`, ai: "AI", manual: "Typed"}[d.textSource] || ""))}
    </dl>
    ${d.text ? `<details${terms.length ? " open" : ""}><summary>Text in the document</summary><div class="ocr">${highlight(d.text, terms)}</div></details>` : ""}
    <div class="btnrow">
      ${canShare() ? `<button class="btn primary grow" data-act="share">${ICON.share}Share</button>` : ""}
      ${canDownload() ? `<button class="btn grow" data-act="download">${ICON.download}Download</button>` : ""}
      ${d.webUrl ? `<a class="btn grow" href="${esc(d.webUrl)}" target="_blank" rel="noopener">${ICON.open}Open in OneDrive</a>` : ""}
      <button class="btn grow" data-act="edit">Edit details</button>
      ${d.dueDate ? `<button class="btn grow" data-act="handled">${ICON.check}${d.handled ? "Mark as open" : "Mark as handled"}</button>` : ""}
      ${d.dueDate ? `<button class="btn grow" data-act="cal">${ICON.cal}Add to calendar</button>` : ""}
      <button class="btn grow" data-act="split">${ICON.crop}Split or cut out</button>
      <button class="btn grow" data-act="reocr">${ICON.text}Read text again</button>
      ${ai.on && !d.weak ? `<button class="btn grow" data-act="ai">${ICON.spark}Read with AI</button>` : ""}
    </div>
    <div id="calbox"></div>
    ${confirmDelete ? `<div class="confirm"><span>Delete this document? The file moves to your OneDrive recycle bin.</span><button class="btn danger" data-act="delyes">Delete</button><button class="btn ghost" data-act="delno">Cancel</button></div>`
      : `<div class="btnrow"><button class="btn danger" data-act="delask">${ICON.trash}Delete</button></div>`}`, {docId: id});
  fillViewer(d, $("#viewer", panel), {rotate: true});
  panel.onclick = async e => {
    const rb = e.target.closest("[data-rot]");
    if (rb) return rotateDocPage(d, Number(rb.dataset.rot), rb);
    const b = e.target.closest("[data-act]"); if (!b) return;
    const a = b.dataset.act;
    if (a === "close") return closePanel();
    if (a === "download") return downloadDoc(d, b);
    if (a === "share") return shareDoc(d, b);
    if (a === "delask") return openDetail(id, true);
    if (a === "delno") return openDetail(id, false);
    if (a === "delyes") {
      b.disabled = true;
      try { await deleteDoc(d); } catch (err) { b.disabled = false; return toast(errText(err)); }
      closePanel(); toast("Deleted. The file is in your OneDrive recycle bin."); return;
    }
    if (a === "edit") return openEdit(d, {filing: !d.filed});
    if (a === "okfiled") { putDoc({...d, autoFiled: ""}); return openDetail(id); }
    if (a === "handled") { putDoc({...d, handled: !d.handled}); return openDetail(id); }
    if (a === "split") return openSplit(d);
    if (a === "ics") return icsOut([d], b);
    if (a === "cal") {
      const box = $("#calbox", panel);
      box.innerHTML = box.innerHTML ? "" : `<div class="card"><h3>${esc(calTitle(d))} · ${esc(fmtDate(d.dueDate))}</h3>
        <div class="btnrow"><a class="btn primary grow" href="${esc(googleCalUrl(d))}" target="_blank" rel="noopener">${ICON.cal}Google Calendar</a><button class="btn grow" data-act="ics">${ICON.download}Outlook or other (.ics)</button></div>
        <p class="hint">The .ics file reminds you ${isWarranty(d) ? "30" : "7"} days before.</p></div>`;
      return;
    }
    if (a === "reocr" || a === "ai") return rereadDoc(d, b, a === "ai");
  };
}
async function rereadDoc(d, b, useAi) {
  if (!useAi && d.textSource === "pdf") return toast("This PDF already contains its own text, so it is fully searchable.");
  const label = b.innerHTML; b.disabled = true; b.textContent = "Reading…";
  try {
    const canvases = await pageImagesFor(d, useAi ? 5 : 20);
    let text = "", conf = null, source, aiData = null;
    if (useAi) {
      const blobs = []; for (const c of canvases) blobs.push(await toBlob(scaled(c, 1800), .85));
      aiData = await aiRead(blobs); text = aiData.text; source = "ai";
    } else {
      const texts = []; let cs = 0;
      for (let i = 0; i < canvases.length; i++) { ocrProgress = m => { if (m.status === "recognizing text") b.textContent = `Reading ${i + 1}/${canvases.length} · ${Math.round(m.progress * 100)}%`; }; const r = await ocrCanvas(canvases[i]); texts.push(r.text); cs += r.conf; }
      ocrProgress = null; text = texts.join("\n\n").trim(); conf = Math.round(cs / canvases.length); source = "ocr";
    }
    openEdit({...d, text: text || d.text, textSource: source, ocrConf: conf, weak: source === "ocr" ? isWeak(text, conf) : false}, {
      ai: aiData, note: useAi ? "AI read the document again. Check the values, then save." : `Text read again${conf !== null ? ` (${conf}% sure)` : ""}. Save to keep it.`});
  } catch (err) { console.error(err); toast(err.userMsg || errText(err)); b.disabled = false; b.innerHTML = label; }
}
function openEdit(d, {filing = false, ai: aiData = null, note = ""} = {}) {
  const merged = aiData ? {...d, ...Object.fromEntries(Object.entries(aiData).filter(([k, v]) => v !== "" && v !== null && (k === "text" || !d[k] || !d.filed)))} : d;
  const f = fieldsFrom(merged);
  if (!catsFor(f.space).includes(f.category)) { const hist = suggestFromHistory(d.text, d.id); if (hist) { f.space = hist.space; f.category = hist.category; } else f.category = catsFor(f.space).at(-1); }
  const more = filing ? nextToFile(d.id) : null;
  showPanel(filing ? "File document" : "Edit details", `
    ${note ? `<div class="flag info">${ICON.check}<span>${esc(note)}</span></div>` : ""}
    ${filing ? `<div class="viewer" id="viewer"><span class="loading">Loading from OneDrive…</span></div>` : ""}
    ${formHtml(f, {textOpen: !!note})}
    <div class="btnrow"><button class="btn primary grow" data-act="save">${ICON.check}${filing ? "File it" : "Save changes"}</button>${more ? `<button class="btn grow" data-act="savenext">File and next</button>` : ""}<button class="btn grow" data-act="cancel">Cancel</button></div>
    <p class="hint">Changing the title, date or category also renames or moves the file in OneDrive.</p>`, {docId: d.id});
  if (filing) fillViewer(d, $("#viewer", panel));
  const getVals = wireForm(f);
  panel.onclick = async e => {
    const b = e.target.closest("[data-act]"); if (!b) return;
    const a = b.dataset.act;
    if (a === "close" || a === "cancel") return getDoc(d.id) ? openDetail(d.id) : closePanel();
    if (a === "save" || a === "savenext") {
      b.disabled = true; const label = b.innerHTML; b.textContent = "Saving…";
      try {
        const cur = getDoc(d.id) || d, vals = getVals();
        const userEdited = vals.text !== (d.text || "");
        const next = {...cur, ...vals, filed: true, autoFiled: "", textSource: userEdited ? "manual" : (d.textSource ?? cur.textSource),
          ocrConf: d.ocrConf ?? cur.ocrConf, weak: userEdited ? vals.text.replace(/\s/g, "").length < 20 : (d.weak ?? cur.weak)};
        const wantFolder = folderOf(next), wantName = fileNameOf(next, extOf(cur.name));
        if (cur.itemId && (wantFolder !== folderOf(cur) || wantName !== cur.name)) {
          const it = await moveItem(cur.itemId, wantFolder, wantName);
          Object.assign(next, {name: it.name, path: `${wantFolder}/${it.name}`, webUrl: it.webUrl || cur.webUrl});
        }
        putDoc(next); toast(filing ? "Filed" : "Changes saved");
        const n = a === "savenext" ? nextToFile(d.id) : null;
        if (n) openEdit(getDoc(n), {filing: true}); else openDetail(d.id);
      } catch (err) { console.error(err); b.disabled = false; b.innerHTML = label; toast(errText(err)); }
    }
  };
}

/* bring file names and folders of filed documents in line with the naming rule */
async function tidyNames() {
  const todo = docs().filter(d => d.filed && d.itemId && d.name && (d.name !== fileNameOf(d, extOf(d.name)) || !String(d.path || "").startsWith(folderOf(d) + "/"))).slice(0, 40);
  for (const d of todo) {
    try {
      const it = await moveItem(d.itemId, folderOf(d), fileNameOf(d, extOf(d.name)));
      putDoc({...d, name: it.name, path: `${folderOf(d)}/${it.name}`, webUrl: it.webUrl || d.webUrl});
    } catch (e) { console.warn("rename skipped", e); }
  }
}

/* ================= settings ================= */
function openSettings() {
  const catBlock = sp => `<div class="card"><h3>${SPACE_LABEL[sp]} categories</h3><div class="catlist">${catsFor(sp).map((c, i) => `<span class="catitem">${esc(c)}<button data-act="rmcat" data-sp="${sp}" data-i="${i}" aria-label="Remove ${esc(c)}">${ICON.x}</button></span>`).join("")}</div>
    <div class="addrow"><input id="add_${sp}" placeholder="New category" maxlength="40"><button class="btn" data-act="addcat" data-sp="${sp}">Add</button></div></div>`;
  const n = docs().length;
  showPanel("Settings", `
    <div class="card"><h3>OneDrive</h3>
      <p class="hint">Signed in as <b>${esc(account?.username || "")}</b>. Your ${n} document${n === 1 ? "" : "s"} are stored in <span class="mono">OneDrive › ${esc(BASE)}</span>, sorted into Private and Business folders by category. An <span class="mono">index.csv</span> there lists them all for Excel. Files you save into its <span class="mono">Inbox</span> folder on your PC are picked up and made searchable automatically.</p>
      <div class="btnrow"><button class="btn grow" data-act="openfolder">${ICON.open}Open folder in OneDrive</button><button class="btn grow" data-act="sync">Sync now</button><button class="btn grow" data-act="signout">Sign out</button></div></div>
    <div class="card"><h3>AI for hard-to-read documents</h3>
      <p class="hint">Optional. Text is always read on your own device first. Only when a page is hard to read, such as a faded receipt, is it sent to AI. You need your own API key from console.anthropic.com; you pay per use, usually a few cents per document. The key is stored only on this device.</p>
      <div class="field"><label for="s_key">API key</label><input id="s_key" type="password" autocomplete="off" placeholder="sk-ant-…" value="${esc(ai.key)}"></div>
      <div class="field"><label for="s_model">Model</label><input id="s_model" class="mono" value="${esc(lsGet("dv.aiModel", "") || "")}" placeholder="claude-haiku-5-5"></div>
      <label class="toggle"><input type="checkbox" id="s_auto" ${ai.auto ? "checked" : ""}>Use AI automatically when a document is hard to read</label>
      <div class="btnrow"><button class="btn primary grow" data-act="saveai">Save</button>${ai.on ? `<button class="btn grow" data-act="clearai">Remove key</button>` : ""}</div></div>
    <div class="card"><h3>Reminders</h3><p class="hint">Put every upcoming payment date and warranty end in your calendar at once. Open the file with Outlook, Google Calendar or your phone's calendar.</p>
      <div class="btnrow"><button class="btn grow" data-act="allics">${ICON.cal}All upcoming dates (.ics)</button></div></div>
    <div class="card"><h3>Export for your accountant</h3><p class="hint">Choose Business, a year and if you like a category in the main screen, then tap <b>Export these</b>. You get a ZIP with the files and an Excel overview with totals.</p></div>
    ${catBlock("private")}${catBlock("business")}
    <div class="card"><h3>About</h3><p class="hint">Document Vault keeps a copy of the index on this device so search also works offline. Saving and viewing files needs an internet connection.</p></div>`);
  panel.onclick = async e => {
    const b = e.target.closest("[data-act]"); if (!b) return;
    const a = b.dataset.act, sp = b.dataset.sp;
    if (a === "close") return closePanel();
    if (a === "allics") {
      const today = new Date().toISOString().slice(0, 10);
      const list = docs().filter(d => d.dueDate && d.dueDate >= today && !d.handled);
      if (!list.length) return toast("There are no upcoming dates.");
      return icsOut(list, b);
    }
    if (a === "openfolder") { try { const r = await g("GET", itemRef("")); window.open(r.webUrl, "_blank", "noopener"); } catch (err) { toast(errText(err)); } return; }
    if (a === "sync") { await refreshIndex(); scanInbox(); return toast(state.syncError ? "Sync failed. Check your connection." : "Up to date"); }
    if (a === "signout") { await idb.clear("kv"); await idb.clear("thumbs"); return msalApp.logoutRedirect({account, postLogoutRedirectUri: redirectUri()}); }
    if (a === "saveai") {
      lsSet("dv.aiKey", $("#s_key", panel).value.trim()); lsSet("dv.aiModel", $("#s_model", panel).value.trim()); lsSet("dv.aiAuto", $("#s_auto", panel).checked);
      toast(ai.on ? "AI settings saved on this device" : "AI is off"); return openSettings();
    }
    if (a === "clearai") { lsSet("dv.aiKey", ""); toast("Key removed from this device"); return openSettings(); }
    if (a === "addcat") {
      const inp = $("#add_" + sp, panel), v = inp.value.trim();
      if (!v) return inp.focus();
      if (catsFor(sp).some(c => norm(c) === norm(v))) return toast("That category already exists.");
      const list = [...catsFor(sp)], oi = list.indexOf("Other"); oi >= 0 ? list.splice(oi, 0, v) : list.push(v);
      saveCats(sp, list); return openSettings();
    }
    if (a === "rmcat") {
      const list = [...catsFor(sp)], c = list[Number(b.dataset.i)];
      const used = docs().filter(d => d.filed && d.space === sp && d.category === c).length;
      if (used) return toast(`${used} document${used > 1 ? "s use" : " uses"} “${c}”. Move them to another category first.`);
      if (list.length <= 1) return toast("Keep at least one category.");
      list.splice(Number(b.dataset.i), 1); saveCats(sp, list); return openSettings();
    }
  };
}
function saveCats(sp, list) {
  state.index.categories = {...state.index.categories, [sp]: list}; state.index.catsUpdatedAt = Date.now();
  render(); scheduleSave();
}

/* ================= wiring ================= */
$("#spaceSeg").addEventListener("click", e => {
  const b = e.target.closest("button"); if (!b) return;
  state.space = b.dataset.space; state.cat = null;
  for (const x of $("#spaceSeg").children) x.setAttribute("aria-pressed", x === b);
  render();
});
$("#chips").addEventListener("click", e => { const b = e.target.closest(".chip"); if (!b) return; state.cat = b.dataset.cat || null; render(); });
let qTimer; $("#q").addEventListener("input", e => { clearTimeout(qTimer); qTimer = setTimeout(() => { state.q = e.target.value; render(); }, 120); });
$("#list").addEventListener("click", e => {
  if (e.target.closest('[data-act="scan"]')) return startScan(true);
  const b = e.target.closest(".doc"); if (!b) return;
  if (state.selecting) { const id = b.dataset.id; state.sel.has(id) ? state.sel.delete(id) : state.sel.add(id); return render(); }
  openDetail(b.dataset.id);
});
$("#filterbar").addEventListener("change", e => {
  if (e.target.id === "fYear") state.year = e.target.value;
  if (e.target.id === "fAmt") state.amt = e.target.value;
  render();
});
$("#filterbar").addEventListener("click", e => {
  const b = e.target.closest("[data-act]"); if (!b) return;
  if (b.dataset.act === "select") { state.selecting = !state.selecting; state.sel.clear(); return render(); }
  if (b.dataset.act === "export") return openExport(filtered().list, exportLabel());
});
$("#selbar").addEventListener("click", e => {
  const b = e.target.closest("[data-act]"); if (!b) return;
  const a = b.dataset.act;
  if (a === "selcancel") { state.selecting = false; state.sel.clear(); return render(); }
  if (a === "selall") { for (const d of filtered().list) state.sel.add(d.id); return render(); }
  const chosen = [...state.sel].map(getDoc).filter(Boolean);
  if (a === "selexport") return openExport(chosen, `Selection ${new Date().toISOString().slice(0, 10)}`);
  if (a === "merge") return openMerge(chosen);
});
$("#fab").addEventListener("click", () => startScan(true));
$("#btnSettings").addEventListener("click", openSettings);
window.addEventListener("online", () => { state.online = true; refreshIndex(); });
window.addEventListener("offline", () => { state.online = false; setSync("Offline", true); });
let lastRefresh = 0;
document.addEventListener("visibilitychange", () => { if (!document.hidden && Date.now() - lastRefresh > 30000 && account) { lastRefresh = Date.now(); refreshIndex().then(scanInbox); } });
setInterval(() => { if (!document.hidden && account && !state.dirty) { lastRefresh = Date.now(); refreshIndex(); } }, 120000);
window.addEventListener("beforeunload", e => { if (queueRunning || state.dirty) { e.preventDefault(); e.returnValue = ""; } });

/* ================= start ================= */
function gate(html) { const g0 = $("#gate"); g0.innerHTML = `<div class="box">${html}</div>`; g0.hidden = false; $("#app").hidden = true; return g0; }
function showSetup(err = "") {
  const g0 = gate(`<h1>Document Vault</h1><p>One-time setup: paste the Application (client) ID of your Microsoft app registration.</p>
    ${err ? `<div class="notice">${esc(err)}</div>` : ""}
    <div class="field"><label for="cid">Application (client) ID</label><input id="cid" placeholder="00000000-0000-0000-0000-000000000000" value="${esc(clientId())}"></div>
    <div class="btnrow"><button class="btn primary grow" id="cidSave">Continue</button></div>`);
  $("#cidSave", g0).onclick = () => {
    const v = $("#cid", g0).value.trim();
    if (!/^[0-9a-f-]{36}$/i.test(v)) return showSetup("That does not look like a client ID. It has 36 characters, such as 1a2b3c4d-1234-…");
    lsSet("dv.clientId", v); location.reload();
  };
}
function showSignIn(err = "") {
  const g0 = gate(`<h1>Document Vault</h1><p>Your documents are stored in your own OneDrive. Sign in with your Microsoft account to continue.</p>
    ${err ? `<div class="notice">${esc(err)}</div>` : ""}
    <div class="btnrow"><button class="btn primary grow" id="signIn">Sign in with Microsoft</button></div>
    <p class="hint">Everything is stored in the folder OneDrive › ${esc(BASE)}. Microsoft asks once to allow access to your OneDrive files; the app only reads and writes in that folder.</p>`);
  $("#signIn", g0).onclick = () => msalApp.loginRedirect({scopes: SCOPES, prompt: "select_account"});
}
async function start() {
  if ("serviceWorker" in navigator && location.protocol === "https:") navigator.serviceWorker.register("sw.js").catch(() => {});
  if (!clientId()) return showSetup();
  try { await initAuth(); }
  catch (e) { console.error(e); return showSignIn("Sign-in did not work: " + (e.errorMessage || e.message || "unknown error") + ". If this keeps happening, check the app registration."); }
  if (!account) return showSignIn();
  $("#gate").hidden = true; $("#app").hidden = false;
  render();
  await loadIndex();
  lastRefresh = Date.now();
  if (!state.syncError) { await tidyNames(); scanInbox(); }
  checkAndroidUpdate();
}
start();
})();
