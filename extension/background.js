// OfficeMini extension - service worker.
//
// The one thing this extension exists to do: when Chrome is about to write a .docx or .xlsx to
// the Downloads folder, stop it and open the file in the editor instead. The editor itself is
// the ordinary web build, bundled into the extension and served from chrome-extension://.
//
// The hand-off has to survive a download URL that only works once (Drive, SharePoint and most
// "export" endpoints), so the bytes are fetched here, while the request is still fresh and the
// site's cookies still apply, and parked in the origin private file system. The editor page
// shares that origin, so it picks them up by token without a second request. `?src=` is the
// fallback when that fetch fails, and the only thing passed for a context-menu click.
//
// MV3 service workers are killed between events, so every listener is registered at the top
// level and all state lives in chrome.storage or the OPFS - never in a module variable.

import { loadSettings, activeExtensions } from "./settings.js";

/** A server that sends the right type but a nameless URL still gets recognised. */
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

const INBOX_TTL_MS = 24 * 60 * 60 * 1000;

export function extOf(name) {
  const i = name.lastIndexOf(".");
  return i > 0 ? name.slice(i + 1).toLowerCase() : "";
}

export function nameFromUrl(url) {
  try { return decodeURIComponent(new URL(url).pathname.split("/").pop() || ""); } catch { return ""; }
}

/**
 * The file name Chrome was about to use, falling back to the URL and then to the MIME type.
 * Exported, with the two helpers above, so that docs/WEB-TESTS.md can exercise them in a page.
 */
export function targetName(item) {
  const fromPath = item.filename ? item.filename.split(/[\\/]/).pop() : "";
  if (fromPath && extOf(fromPath)) return fromPath;
  const fromUrl = nameFromUrl(item.finalUrl || item.url || "");
  if (fromUrl && extOf(fromUrl)) return fromUrl;
  const byMime = MIME_EXT[(item.mime || "").toLowerCase()];
  if (byMime) return (fromPath || fromUrl || "Document") + "." + byMime;
  return fromPath || fromUrl || "";
}

// ---- the inbox: bytes handed to the editor page -----------------------------

async function inboxDir(create) {
  const root = await navigator.storage.getDirectory();
  return root.getDirectoryHandle("inbox", { create });
}

/** Fetch the document and park it for the editor page. Returns false to fall back to `?src=`. */
async function stash(token, url) {
  try {
    const res = await fetch(url, { credentials: "include" });
    if (!res.ok) return false;
    const buf = await res.arrayBuffer();
    if (!buf.byteLength) return false;
    const dir = await inboxDir(true);
    const w = await (await dir.getFileHandle(token, { create: true })).createWritable();
    await w.write(buf);
    await w.close();
    return true;
  } catch { return false; }
}

/**
 * Inbox copies are kept after the editor has read them, so that reloading the tab reopens the
 * same document rather than re-requesting a URL that may be spent. They are swept a day later.
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

const FETCHABLE = /^(https?|data):/;

async function openInEditor(url, name, openerTabId) {
  const token = crypto.randomUUID();
  const stashed = await stash(token, url);
  if (!stashed && !FETCHABLE.test(url)) return false;

  const q = new URLSearchParams();
  if (stashed) q.set("inbox", token);
  // A blob: URL belongs to the page that made it and cannot be re-fetched here, so it is only
  // ever usable through the stash above.
  if (FETCHABLE.test(url)) q.set("src", url);
  q.set("name", name || nameFromUrl(url) || "Document");

  const opts = { url: chrome.runtime.getURL("index.html") + "?" + q.toString() };
  if (openerTabId !== undefined) opts.openerTabId = openerTabId;
  await chrome.tabs.create(opts);
  return true;
}

// ---- events -----------------------------------------------------------------

chrome.downloads.onCreated.addListener(async (item) => {
  const s = await loadSettings();
  if (!s.enabled) return;
  const url = item.finalUrl || item.url || "";
  if (url.startsWith("blob:")) return;          // not fetchable from a service worker
  const name = targetName(item);
  if (!name || !activeExtensions(s).has(extOf(name))) return;

  // Cancel first: the point is that nothing reaches the Downloads folder. Chrome removes the
  // partial file, and erase() takes the row out of the downloads list too.
  try { await chrome.downloads.cancel(item.id); } catch { /* already finished; open it anyway */ }
  try { await chrome.downloads.erase({ id: item.id }); } catch { /* the row can stay */ }
  await openInEditor(url, name);
});

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({ id: "open-link", title: "Open link in OfficeMini", contexts: ["link"], targetUrlPatterns: ["http://*/*", "https://*/*"] });
  });
  sweepInbox();
});

chrome.runtime.onStartup.addListener(sweepInbox);

chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  if (info.menuItemId !== "open-link" || !info.linkUrl) return;
  await openInEditor(info.linkUrl, nameFromUrl(info.linkUrl), tab && tab.id);
});
