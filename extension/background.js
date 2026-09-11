// OfficeMini extension - service worker.
//
// When Chrome downloads a Word or Excel file, open it in the editor. The editor is the ordinary
// web build, bundled into the extension and served from chrome-extension://.
//
// Chrome does the download exactly as it would without us, and the file it wrote is read back
// once the download is complete. That is the one approach that works for every way a site hands
// a file out: a plain link, an export endpoint that names the file in Content-Disposition, a
// POST, a URL that only works once, or a spreadsheet built in the page and saved from a blob:
// URL. The first version cancelled the download and fetched the URL again itself, which handled
// the first of those and none of the rest - and admin panels use the rest.
//
// Reading the finished file needs "Allow access to file URLs" on the extension. Without it, a
// plain link is fetched again instead, and anything else opens file-access.html, which explains
// the switch and finishes the job once it is on.
//
// A local file opened into Chrome - dropped on a tab - arrives as a download of a file:// URL.
// Letting it run would put a second copy in Downloads, so it is cancelled and the original is
// opened instead. Cancelling loses nothing: the file is already on disk.
//
// The bytes are parked in the origin private file system under inbox/<token>; the editor tab
// shares the chrome-extension:// origin, so it picks them up by token (src/files-web.ts).
//
// MV3 service workers are stopped between events, so every listener is registered at the top
// level and all state lives in chrome.storage or the OPFS - never in a module variable.

import { loadSettings, activeExtensions } from "./settings.js";

/** A server that sends the right type but no usable name still gets recognised. */
const MIME_EXT = {
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
  "application/vnd.ms-word.document.macroenabled.12": "docm",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.template": "dotx",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
  "application/vnd.ms-excel.sheet.macroenabled.12": "xlsm",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.template": "xltx",
  "text/csv": "csv",
  "text/tab-separated-values": "tsv",
  "text/markdown": "md",
};

/** Formats that are zip packages. Anything claiming to be one that is not is an error page. */
const ZIP_EXT = new Set(["docx", "docm", "dotx", "xlsx", "xlsm", "xltx"]);

const INBOX_TTL_MS = 24 * 60 * 60 * 1000;
const PENDING_TTL_MS = 15 * 60 * 1000;
const FETCHABLE = /^(https?|data):/;

// ---- names ------------------------------------------------------------------
// Exported so that docs/WEB-TESTS.md can exercise them in a page.

export function extOf(name) {
  const i = name.lastIndexOf(".");
  return i > 0 ? name.slice(i + 1).toLowerCase() : "";
}

export function baseName(p) { return (p || "").split(/[\\/]/).pop() || ""; }

export function nameFromUrl(url) {
  try { return decodeURIComponent(new URL(url).pathname.split("/").pop() || ""); } catch { return ""; }
}

/** The file name a download will have, falling back to the URL and then to the MIME type. */
export function targetName(item) {
  const fromPath = baseName(item.filename);
  if (fromPath && extOf(fromPath)) return fromPath;
  const fromUrl = nameFromUrl(item.finalUrl || item.url || "");
  if (fromUrl && extOf(fromUrl)) return fromUrl;
  const byMime = MIME_EXT[(item.mime || "").toLowerCase()];
  if (byMime) return (fromPath || fromUrl || "Document") + "." + byMime;
  return fromPath || fromUrl || "";
}

/** A response's own idea of its name: Content-Disposition first, then the MIME type. */
function nameFromResponse(res, fallback) {
  const cd = res.headers.get("content-disposition") || "";
  const star = /filename\*\s*=\s*[^']*''([^;]+)/i.exec(cd);
  const plain = /filename\s*=\s*"?([^";]+)"?/i.exec(cd);
  let name = star ? decodeURIComponent(star[1]) : plain ? plain[1].trim() : fallback;
  if (!extOf(name || "")) {
    const byMime = MIME_EXT[(res.headers.get("content-type") || "").split(";")[0].trim().toLowerCase()];
    if (byMime) name = (name || "Document") + "." + byMime;
  }
  return name || fallback;
}

/** A path from the downloads API as a file:// URL. */
export function fileUrl(p) {
  const parts = p.replace(/\\/g, "/").split("/").map((seg, i) => (i === 0 && /^[A-Za-z]:$/.test(seg) ? seg : encodeURIComponent(seg)));
  return "file:///" + parts.join("/").replace(/^\/+/, "");
}

/** Bytes that are plausibly the document, rather than an error page served in its place. */
export function looksRight(ext, buf) {
  const b = new Uint8Array(buf, 0, Math.min(buf.byteLength, 256));
  if (!b.length) return false;
  if (ZIP_EXT.has(ext)) return b[0] === 0x50 && b[1] === 0x4b;
  const head = new TextDecoder().decode(b).trimStart().slice(0, 32).toLowerCase();
  return !head.startsWith("<!doctype html") && !head.startsWith("<html");
}

// ---- the inbox: bytes handed to the editor page -----------------------------

async function inboxDir(create) {
  const root = await navigator.storage.getDirectory();
  return root.getDirectoryHandle("inbox", { create });
}

/**
 * Read `url` - a web address or a file:// one - into the inbox under `token`. Returns the name
 * to open it under, or null when it could not be read or is not the document it claims to be.
 */
async function stash(token, url, name) {
  try {
    const res = await fetch(url, { credentials: "include" });
    if (!res.ok) return null;
    const final = extOf(name || "") ? name : nameFromResponse(res, name);
    const buf = await res.arrayBuffer();
    if (!looksRight(extOf(final || ""), buf)) return null;
    const w = await (await (await inboxDir(true)).getFileHandle(token, { create: true })).createWritable();
    await w.write(buf);
    await w.close();
    return final || "Document";
  } catch { return null; }
}

/**
 * Inbox copies are kept after the editor has read them, so that reloading the tab reopens the
 * same document. They are swept a day later.
 */
async function sweepInbox() {
  try {
    const dir = await inboxDir(false);
    const cutoff = Date.now() - INBOX_TTL_MS;
    const stale = [];
    for await (const [name, handle] of dir.entries()) {
      if (handle.kind !== "file") continue;
      if ((await handle.getFile()).lastModified < cutoff) stale.push(name);
    }
    for (const name of stale) await dir.removeEntry(name);
  } catch { /* no inbox yet */ }
}

// ---- opening ----------------------------------------------------------------

async function openEditor(query, openerTabId) {
  const opts = { url: chrome.runtime.getURL("index.html") + "?" + new URLSearchParams(query).toString() };
  if (openerTabId !== undefined) opts.openerTabId = openerTabId;
  await chrome.tabs.create(opts);
}

const fileAccess = () => chrome.extension.isAllowedFileSchemeAccess();

/** The one thing that needs the user: explain the switch, and remember what to open after. */
async function askForFileAccess(name, retry) {
  await chrome.storage.local.set({ pending: { ...retry, name, at: Date.now() } });
  await chrome.tabs.create({ url: chrome.runtime.getURL("file-access.html") + "?" + new URLSearchParams({ name }).toString() });
}

/** A file that is already on disk, opened into Chrome. */
async function openLocal(url, name) {
  if (!(await fileAccess())) return askForFileAccess(name, { kind: "local", url });
  const token = crypto.randomUUID();
  const final = await stash(token, url, name);
  if (final) await openEditor({ inbox: token, name: final });
}

/** A finished download: read the file Chrome wrote, or fetch the link again if we may not. */
async function openDownload(item, s) {
  const name = targetName(item);
  const src = item.finalUrl || item.url || "";
  const token = crypto.randomUUID();
  const access = await fileAccess();
  let final = access && item.filename ? await stash(token, fileUrl(item.filename), name) : null;
  if (!final && FETCHABLE.test(src)) final = await stash(token, src, name);
  if (!final) {
    if (!access) await askForFileAccess(name, { kind: "download", id: item.id });
    return;
  }
  const query = { inbox: token, name: final };
  if (FETCHABLE.test(src)) query.src = src;   // the page's fallback if the inbox is ever missing
  await openEditor(query);
  if (!s.keepCopy) {
    // The editor holds its own copy now, so the one in Downloads can go - that is what makes this
    // "opened instead of downloaded" rather than "downloaded, then opened as well".
    try { await chrome.downloads.removeFile(item.id); } catch { /* moved or opened elsewhere */ }
    try { await chrome.downloads.erase({ id: item.id }); } catch { /* the row can stay */ }
  }
}

/**
 * Finish what file-access.html was waiting for, once the switch is on. Several things can ask at
 * once - the worker starting, the explainer page, the popup - so the in-flight promise is shared:
 * two callers must never open the same file twice. (That is the one module variable here, and it
 * only lives as long as the call.)
 */
let resuming = null;
function resumePending() {
  if (!resuming) resuming = doResume().finally(() => { resuming = null; });
  return resuming;
}

async function doResume() {
  const { pending } = await chrome.storage.local.get("pending");
  if (!pending) return false;
  if (Date.now() - pending.at > PENDING_TTL_MS) { await chrome.storage.local.remove("pending"); return false; }
  if (!(await fileAccess())) return false;
  await chrome.storage.local.remove("pending");
  if (pending.kind === "local") await openLocal(pending.url, pending.name);
  else {
    const [item] = await chrome.downloads.search({ id: pending.id });
    if (item && item.state === "complete" && item.exists !== false) await openDownload(item, await loadSettings());
  }
  return true;
}

// ---- events -----------------------------------------------------------------

// A local file opened into Chrome. Web downloads are not touched here: they finish first.
chrome.downloads.onCreated.addListener(async (item) => {
  const url = item.finalUrl || item.url || "";
  if (!url.startsWith("file:")) return;
  const s = await loadSettings();
  if (!s.enabled) return;
  const name = targetName(item);
  if (!activeExtensions(s).has(extOf(name))) return;
  try { await chrome.downloads.cancel(item.id); } catch { /* finished already */ }
  try { await chrome.downloads.erase({ id: item.id }); } catch { /* the row can stay */ }
  await openLocal(url, name);
});

// A web download that has finished - by now Chrome knows its real name, whatever the URL said.
chrome.downloads.onChanged.addListener(async (delta) => {
  if (!delta.state || delta.state.current !== "complete") return;
  const s = await loadSettings();
  if (!s.enabled) return;
  const [item] = await chrome.downloads.search({ id: delta.id });
  if (!item || (item.finalUrl || item.url || "").startsWith("file:")) return;
  if (!activeExtensions(s).has(extOf(targetName(item)))) return;
  await openDownload(item, s);
});

chrome.runtime.onMessage.addListener((msg, _sender, reply) => {
  if (msg && msg.type === "resume") { resumePending().then(reply); return true; }
  return false;
});

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({ id: "open-link", title: "Open link in OfficeMini", contexts: ["link"], targetUrlPatterns: ["http://*/*", "https://*/*"] });
  });
  sweepInbox();
});

chrome.runtime.onStartup.addListener(sweepInbox);

// Not a download at all, so this is the one path that fetches the URL itself.
chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  if (info.menuItemId !== "open-link" || !info.linkUrl) return;
  const url = info.linkUrl;
  const token = crypto.randomUUID();
  const final = await stash(token, url, nameFromUrl(url));
  // Nothing parked: hand the page the URL, so that it tries itself and says what went wrong.
  await openEditor(final ? { inbox: token, src: url, name: final } : { src: url, name: nameFromUrl(url) || "Document" }, tab && tab.id);
});

// Every start of the worker checks for a file that was waiting on the file-access switch. Chrome
// reloads the extension when that switch changes, and which of our events follow differs between
// browsers - so the check runs whenever the worker runs at all, not on one particular event.
resumePending();
