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
  space: "all", cat: null, q: "", online: navigator.onLine,
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
  const lo = pct(mode === "bw" ? .04 : .015), hi = Math.max(lo + 30, pct(mode === "bw" ? .80 : .985));
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
    const url = URL.createObjectURL(pageBlobs[i]); const img = await loadImg(url); URL.revokeObjectURL(url);
    const w = img.naturalWidth, h = img.naturalHeight, o = w > h ? "l" : "p";
    if (!pdf) pdf = new jsPDF({orientation: o, unit: "mm", format: "a4", compress: true}); else pdf.addPage("a4", o);
    const pw = pdf.internal.pageSize.getWidth(), ph = pdf.internal.pageSize.getHeight(), m = 6, s = Math.min((pw - 2 * m) / w, (ph - 2 * m) / h);
    pdf.addImage(img, "JPEG", (pw - w * s) / 2, (ph - h * s) / 2, w * s, h * s);
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
  const ref = t.match(/(?:factuurnummer|factuurnr\.?|invoice (?:no|number)\.?|polisnummer|ordernummer|order (?:no|number)\.?|bonnummer)[:\s#]*([A-Z0-9][A-Z0-9\-/.]{2,24})/i);
  if (ref) g.reference = ref[1];
  const first = t.split("\n").map(s => s.trim()).find(s => /[a-z]{3,}/i.test(s) && s.length >= 3 && s.length <= 60);
  if (first) g.sender = first.replace(/[|_~*]+/g, " ").trim();
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
function filtered() {
  const terms = norm(state.q).split(/\s+/).filter(Boolean);
  const out = [];
  for (const d of docs()) {
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
  if (!list.length) { L.innerHTML = `<div class="empty"><h2>No matches</h2><p>Nothing matches ${state.q ? `“${esc(state.q)}”` : "this filter"}. Try a different word or switch to All.</p></div>`; return; }
  let h = "", lastMonth = "";
  for (const d of list.slice(0, 400)) {
    if (!terms.length) {
      const dd = docDate(d), m = dd ? new Date(dd + "T12:00:00").toLocaleDateString("en-GB", {month: "long", year: "numeric"}) : "Undated";
      if (m !== lastMonth) { h += `<div class="month">${esc(m)}</div>`; lastMonth = m; }
    }
    const due = dueInfo(d), amt = fmtAmount(d.amount, d.currency);
    const sub = [d.sender, fmtDate(d.date)].filter(Boolean).join(" · ") || d.name || "";
    const sn = terms.length ? snippet(d.text, terms) : "";
    h += `<button class="doc" data-id="${esc(d.id)}">
      <img class="thumb" data-thumb="${esc(d.itemId || "")}" data-v="${esc(String(d.fileVersion || 0))}" alt="">
      <span class="meta"><span class="t">${terms.length ? highlight(d.title || "Untitled", terms) : esc(d.title || "Untitled")}</span>
      <span class="sub">${esc(sub)}</span>
      <span class="tags">${d.filed ? `<span class="pill ${d.space}">${SPACE_LABEL[d.space]}</span><span class="cat">${esc(d.category)}</span>` : `<span class="due">To file</span>`}${d.weak ? `<span class="weakpill">Hard to read</span>` : ""}${due ? `<span class="due">${due.label}</span>` : ""}${amt ? `<span class="amt mono">${esc(amt)}</span>` : ""}</span>
      ${sn ? `<span class="snip">${sn}</span>` : ""}</span></button>`;
  }
  if (list.length > 400) h += `<p class="hint">Showing 400 of ${list.length}. Search to narrow down.</p>`;
  L.innerHTML = h;
  observeThumbs();
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
const fileNameOf = (d, ext) => `${d.date || "undated"} ${safeSeg(d.title, 80)}${ext}`;
const extOf = name => (String(name || "").match(/\.[a-z0-9]{2,5}$/i) || [".pdf"])[0].toLowerCase();

/* ================= scan flow (phone camera) ================= */
let draft = null;
function startScan(openCamera) {
  draft = {pages: [], mode: "enhanced"}; renderScan();
  if (openCamera) $("#camIn").click();
}
async function addScanFiles(files) {
  if (!draft) return;
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
$("#camIn").addEventListener("change", e => { const f = [...e.target.files]; e.target.value = ""; addScanFiles(f); });
$("#galIn").addEventListener("change", e => { const f = [...e.target.files]; e.target.value = ""; addScanFiles(f); });

function renderScan() {
  const n = draft.pages.length;
  const pagesHtml = draft.pages.map((p, i) => `<div class="pg"><img src="${p.preview}" alt="Page ${i + 1}"><div class="bar"><span>Page ${i + 1}</span><div>
      <button class="iconbtn" data-act="rot" data-i="${i}" aria-label="Rotate page ${i + 1}">${ICON.rotate}</button>
      <button class="iconbtn" data-act="crop" data-i="${i}" aria-label="Crop page ${i + 1}">${ICON.crop}</button>
      <button class="iconbtn" data-act="del" data-i="${i}" aria-label="Remove page ${i + 1}">${ICON.trash}</button></div></div></div>`).join("");
  showPanel("New scan", `
    ${n ? `<div class="pages">${pagesHtml}</div>` : `<div class="card"><p class="hint">Take a photo of each page. Lay the document flat in good light and fill the frame.</p></div>`}
    <div class="btnrow"><button class="btn grow" data-act="cam">${ICON.camera}${n ? "Add page" : "Take photo"}</button><button class="btn grow" data-act="gal">${ICON.image}From gallery</button></div>
    ${n ? `<div class="field"><span class="label">Image style</span><div class="seg" id="modeSeg">
        ${[["original", "Original"], ["enhanced", "Enhanced"], ["bw", "Black & white"]].map(([k, l]) => `<button data-mode="${k}" aria-pressed="${draft.mode === k}">${l}</button>`).join("")}</div></div>
      <div class="btnrow"><button class="btn primary grow" data-act="read">${ICON.text}Read text and continue</button></div>
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
  if (!f.title) f.title = [f.sender, f.date ? fmtDate(f.date) : ""].filter(Boolean).join(", ");
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
  return "Saving failed. Check your connection and try again.";
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
function enqueueFiles(files) {
  const ok = files.filter(f => isPdfFile(f) || /^image\//.test(f.type));
  if (files.length > ok.length) toast(`${files.length - ok.length} file${files.length - ok.length > 1 ? "s" : ""} skipped. Only PDF and image files can be added.`);
  if (!ok.length) return;
  for (const f of ok) queue.push({file: f, name: f.name || "Pasted image.png", status: "waiting", note: ""});
  renderImport(); runQueue();
}
async function runQueue() {
  if (queueRunning) return; queueRunning = true;
  for (let it; (it = queue.find(q => q.status === "waiting"));) {
    it.status = "working"; it.note = "Opening"; renderImport();
    try { await importOne(it); }
    catch (e) { console.error(e); it.status = "failed"; it.note = errText(e); }
    it.file = null; ocrProgress = null; renderImport();
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
  const fresh = (list.value || []).filter(x => x.file && !known.has(x.id) && !queued.has(x.id) && (isPdfFile({name: x.name, type: x.file.mimeType}) || /^image\//.test(x.file.mimeType || "")));
  if (!fresh.length) return;
  for (const x of fresh) queue.push({remote: x, name: x.name, status: "waiting", note: ""});
  renderImport(); runQueue();
}
async function importOne(it) {
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
  const filed = !!(hist || (aiData && aiData.category));
  if (!catsFor(base.space).includes(base.category)) base.category = catsFor(base.space).at(-1);
  const meta = {id: uid(), ...base, filed, autoFiled: filed ? (hist ? "history" : "ai") : "", weak, textSource: source, ocrConf: conf,
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
    const files = [...(e.clipboardData?.files || [])]; if (files.length) { e.preventDefault(); enqueueFiles(files); }
  });
}
const nextToFile = (except) => docs().filter(d => !d.filed && d.id !== except).sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0))[0]?.id;

/* ================= document detail ================= */
const fileCache = new Map();
async function getFileBlob(d) {
  const k = d.itemId + ":" + (d.fileVersion || 0);
  if (fileCache.has(k)) return fileCache.get(k);
  const b = await downloadItem(d.itemId);
  fileCache.set(k, b); if (fileCache.size > 6) fileCache.delete(fileCache.keys().next().value);
  return b;
}
async function fillViewer(d, el) {
  try {
    const blob = await getFileBlob(d);
    if (!el.isConnected) return;
    if ((d.mime || "").includes("pdf") || /\.pdf$/i.test(d.name || "")) {
      const pdf = await openPdf(await blob.arrayBuffer());
      el.innerHTML = "";
      const n = Math.min(pdf.numPages, 15);
      for (let i = 1; i <= n; i++) { if (!el.isConnected) break; const c = await pdfPageCanvas(pdf, i, 1500); c.setAttribute("aria-label", `Page ${i}`); el.appendChild(c); }
      if (pdf.numPages > n && el.isConnected) el.insertAdjacentHTML("beforeend", `<span class="loading">${pdf.numPages - n} more pages. Open the file to see all.</span>`);
      pdf.destroy();
    } else {
      const u = URL.createObjectURL(blob); el.innerHTML = `<img src="${u}" alt="Document">`;
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
    ${d.autoFiled ? `<div class="flag info">${ICON.check}<span>Filed automatically${d.autoFiled === "history" ? " based on earlier documents from this sender" : " by AI"}.</span><button class="btn" data-act="okfiled">Looks good</button></div>` : ""}
    ${d.weak ? `<div class="flag weak">${ICON.text}<span>The text was hard to read, so search may miss this document.</span>${ai.on ? `<button class="btn" data-act="ai">${ICON.spark}Read with AI</button>` : `<button class="btn" data-act="edit">Add key words</button>`}</div>` : ""}
    <div class="viewer" id="viewer"><span class="loading">Loading from OneDrive…</span></div>
    <div class="tags">${d.filed ? `<span class="pill ${d.space}">${SPACE_LABEL[d.space]}</span><span class="cat">${esc(d.category)}</span>` : ""}${due ? `<span class="due">${due.label}</span>` : ""}${d.handled && d.dueDate ? `<span class="cat">Handled</span>` : ""}</div>
    <dl class="facts">
      ${row("Date", esc(fmtDate(d.date)))}${row("From", highlight(d.sender || "", terms))}${row("Amount", amt ? `<span class="mono">${esc(amt)}</span>` : "")}
      ${row("Due / ends", esc(fmtDate(d.dueDate)))}${row("Reference", d.reference ? `<span class="mono">${highlight(d.reference, terms)}</span>` : "")}
      ${row("Tags", esc((d.tags || []).join(", ")))}${row("Notes", d.notes ? highlight(d.notes, terms) : "")}${row("Pages", String(d.pages || 1))}
      ${row("OneDrive", `<span class="mono">${esc(BASE)}/${esc(d.path || "")}</span>`)}${row("Text from", esc({pdf: "PDF text", ocr: `Text recognition${d.ocrConf ? ` (${d.ocrConf}% sure)` : ""}`, ai: "AI", manual: "Typed"}[d.textSource] || ""))}
    </dl>
    ${d.text ? `<details${terms.length ? " open" : ""}><summary>Text in the document</summary><div class="ocr">${highlight(d.text, terms)}</div></details>` : ""}
    <div class="btnrow">
      ${d.webUrl ? `<a class="btn grow" href="${esc(d.webUrl)}" target="_blank" rel="noopener">${ICON.open}Open in OneDrive</a>` : ""}
      <button class="btn grow" data-act="edit">Edit details</button>
      ${d.dueDate ? `<button class="btn grow" data-act="handled">${ICON.check}${d.handled ? "Mark as open" : "Mark as handled"}</button>` : ""}
      <button class="btn grow" data-act="reocr">${ICON.text}Read text again</button>
      ${ai.on && !d.weak ? `<button class="btn grow" data-act="ai">${ICON.spark}Read with AI</button>` : ""}
    </div>
    ${confirmDelete ? `<div class="confirm"><span>Delete this document? The file moves to your OneDrive recycle bin.</span><button class="btn danger" data-act="delyes">Delete</button><button class="btn ghost" data-act="delno">Cancel</button></div>`
      : `<div class="btnrow"><button class="btn danger" data-act="delask">${ICON.trash}Delete</button></div>`}`, {docId: id});
  fillViewer(d, $("#viewer", panel));
  panel.onclick = async e => {
    const b = e.target.closest("[data-act]"); if (!b) return;
    const a = b.dataset.act;
    if (a === "close") return closePanel();
    if (a === "delask") return openDetail(id, true);
    if (a === "delno") return openDetail(id, false);
    if (a === "delyes") {
      b.disabled = true;
      try { await g("DELETE", `/me/drive/items/${d.itemId}`); } catch (err) { if (err.status !== 404) { b.disabled = false; return toast(errText(err)); } }
      putDoc({id: d.id, deleted: true, createdAt: d.createdAt}); closePanel(); toast("Deleted. The file is in your OneDrive recycle bin."); return;
    }
    if (a === "edit") return openEdit(d, {filing: !d.filed});
    if (a === "okfiled") { putDoc({...d, autoFiled: ""}); return openDetail(id); }
    if (a === "handled") { putDoc({...d, handled: !d.handled}); return openDetail(id); }
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
    ${catBlock("private")}${catBlock("business")}
    <div class="card"><h3>About</h3><p class="hint">Document Vault keeps a copy of the index on this device so search also works offline. Saving and viewing files needs an internet connection.</p></div>`);
  panel.onclick = async e => {
    const b = e.target.closest("[data-act]"); if (!b) return;
    const a = b.dataset.act, sp = b.dataset.sp;
    if (a === "close") return closePanel();
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
  const b = e.target.closest(".doc"); if (b) openDetail(b.dataset.id);
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
  if (!state.syncError) scanInbox();
}
start();
})();
