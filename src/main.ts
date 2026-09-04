// OfficeMini application shell: menus, file handling, editor wiring.
import { EditorView } from "prosemirror-view";
import { EditorState, NodeSelection, TextSelection } from "prosemirror-state";
import { Node as PMNode } from "prosemirror-model";
import { selectAll } from "prosemirror-commands";
import { undo, redo } from "prosemirror-history";
import { schema, SectProps, clearCssCaches } from "./schema";
import { loadDocx, LoadedDoc, resolveDeferredMedia } from "./docx/parse";
import { writeDocx } from "./docx/write";
import { blankDocxBytes } from "./docx/template";
import { ctx } from "./docx/styles";
import { markdownToDoc, docToMarkdown } from "./md/markdown";
import { createEditor, EditorHandle, requestPlainPaste, selectPos, redrawView, setSmartTyping } from "./editor/editor";
import { numberingKey } from "./editor/lists";
import { setDarkMode } from "./docx/props";
import { AppActions, Shortcut, keyLabel } from "./editor/keymap";
import * as C from "./editor/commands";
import * as T from "./editor/tables";
import { getFind, setFindQuery, closeFind, findStep, replaceCurrent, replaceAll, FindOptions, DEFAULT_FIND_OPTIONS, textareaFind, textareaReplace } from "./editor/find";
import { setZoomFactor, setHeaderFooters, setViewMode, paginationKey, pageAt, pageStartPos, relayout, resetHeaderFooterCache, setDefaultTabStop } from "./editor/pagination";
import { showContextMenu, tableMenu, pasteFromClipboard, ContextActions } from "./editor/contextmenu";
import { buildToolbar, ToolbarHandle, colorPopup } from "./ui/toolbar";
import { buildStatusbar, StatusHandle } from "./ui/statusbar";
import { el, showMenu, MenuItem, tooltip, closeAllPopups, icon } from "./ui/widgets";
import { showDialog, linkDialog, tableDialog, pageSetupDialog, paragraphDialog, goToPageDialog, shortcutsDialog, aboutDialog, closeDialog } from "./ui/dialogs";
import { imageSize } from "./docx/images";
import { twipsToPx } from "./docx/units";
import { setPrintPageSize, printDocument } from "./print";
import { checkForUpdates } from "./updater";
import * as F from "./files";

type Kind = "docx" | "md" | "new";

const app = {
  handle: null as EditorHandle | null,
  loaded: null as LoadedDoc | null,
  path: null as string | null,
  kind: "new" as Kind,
  dirty: false,
  zoom: 1,
  mode: "page" as "page" | "continuous",
  showMarks: false,
  theme: "light" as "light" | "dark",
  source: false,
  welcome: false,
  recoveryId: null as string | null,
  pages: 1,
  settings: {} as F.Settings,
  shortcuts: [] as Shortcut[],
  toolbar: null as ToolbarHandle | null,
  status: null as StatusHandle | null,
  shown: false,
  loading: false,
};

const $ = (id: string) => document.getElementById(id)!;
const view = () => app.handle!.view;
const run = (cmd: (s: EditorState, d?: any, v?: EditorView) => boolean, focus = true) => {
  const v = view();
  cmd(v.state, v.dispatch, v);
  if (focus) v.focus();
};

// ---------------------------------------------------------------------------
// Title / dirty state

function docName(): string {
  return app.path ? F.basename(app.path) : "Untitled";
}
function updateTitle() {
  F.setWindowTitle(`${docName()}${app.dirty ? " •" : ""} – OfficeMini`);
}
function setDirty(d: boolean) {
  if (app.dirty === d) return;
  app.dirty = d;
  updateTitle();
}

// ---------------------------------------------------------------------------
// Loading documents

async function loadBlank(): Promise<LoadedDoc> {
  return loadDocx(blankDocxBytes());
}

function installDocument(loaded: LoadedDoc, doc: PMNode, path: string | null, kind: Kind) {
  app.loaded = loaded;
  app.path = path;
  app.kind = kind;
  resetHeaderFooterCache();
  setHeaderFooters({ headers: loaded.headers, footers: loaded.footers });
  setDefaultTabStop(loaded.ctx.defaultTabStop);
  setPrintPageSize(doc.attrs.sect as SectProps);
  app.handle!.setDocument(doc);
  setDirty(false);
  updateTitle();
  view().focus();
  const st = paginationKey.getState(view().state);
  if (st && st.mode !== app.mode) setViewMode(view(), app.mode);
  if (loaded.warnings.length) console.warn("Document warnings:", loaded.warnings);
}

// ---------------------------------------------------------------------------
// Welcome screen (start without a document) and crash recovery

interface RecoveryEntry { id: string; path: string | null; name: string; file: string; savedAt: number; kind: "docx" | "md"; }

function fileKind(path: string): "docx" | "md" { return F.extname(path) === "docx" ? "docx" : "md"; }

async function buildWelcome(recovered: RecoveryEntry[]) {
  const w = $("welcome");
  w.innerHTML = "";
  const card = el("div", { class: "welcome-card" });
  card.append(el("h1", null, "OfficeMini"), el("div", { class: "sub" }, "Open a document or start a new one. Word and Markdown files, fast."));
  const actions = el("div", { class: "welcome-actions" });
  const newBtn = el("button", { class: "primary" }, icon("new"), "New document");
  newBtn.addEventListener("click", () => { hideWelcome(); view().focus(); });
  const openBtn = el("button", null, icon("open"), "Open…");
  openBtn.addEventListener("click", () => openFile());
  actions.append(newBtn, openBtn);
  card.append(actions);
  if (recovered.length) {
    const box = el("div", { class: "welcome-recovery" });
    box.append(el("h2", null, "Recovered documents"));
    for (const r of recovered) {
      const item = el("div", { class: "recent-item", title: r.path || "Unsaved document" });
      const ic = el("span", { class: "ri-icon " + r.kind }, r.kind === "docx" ? "W" : "M");
      const name = el("span", { class: "ri-name" }, r.name);
      const when = el("span", { class: "ri-path", style: { direction: "ltr" } }, "autosaved " + new Date(r.savedAt).toLocaleString());
      const x = el("span", { class: "ri-x", title: "Discard recovered copy" }, "✕");
      x.addEventListener("click", async (e) => { e.stopPropagation(); await discardRecovery(r); item.remove(); if (!box.querySelector(".recent-item")) box.remove(); });
      item.append(ic, name, when, x);
      item.addEventListener("click", () => openRecovered(r));
      box.append(item);
    }
    card.append(box);
  }
  const recent = app.settings.recent || [];
  if (recent.length) {
    card.append(el("h2", null, "Recent"));
    const list = el("div", { class: "recent-list" });
    for (const p of recent) {
      const exists = F.isTauri ? await F.fileExists(p) : true;
      const item = el("div", { class: "recent-item" + (exists ? "" : " missing"), title: p });
      const ic = el("span", { class: "ri-icon " + fileKind(p) }, fileKind(p) === "docx" ? "W" : "M");
      const name = el("span", { class: "ri-name" }, F.basename(p));
      const dir = el("span", { class: "ri-path" }, F.dirname(p));
      const x = el("span", { class: "ri-x", title: "Remove from list" }, "✕");
      x.addEventListener("click", (e) => { e.stopPropagation(); app.settings.recent = (app.settings.recent || []).filter((q) => q !== p); F.saveSettings(app.settings); item.remove(); });
      item.append(ic, name, dir, x);
      item.addEventListener("click", () => { if (exists) openPath(p); else app.status?.flash("File not found: " + p); });
      list.append(item);
    }
    card.append(list);
  }
  card.append(el("div", { class: "welcome-hint" }, "Ctrl+O open · Ctrl+N new · Ctrl+/ all shortcuts · drop a file on the window to open it"));
  w.append(card);
}
function showWelcome() { app.welcome = true; $("welcome").hidden = false; }
function hideWelcome() { if (!app.welcome) return; app.welcome = false; $("welcome").hidden = true; }

function recoveryIdFor(path: string | null): string {
  if (!path) return "untitled-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  let h = 0;
  for (let i = 0; i < path.length; i++) h = (h * 31 + path.charCodeAt(i)) >>> 0;
  return "doc-" + h.toString(16) + "-" + F.basename(path).replace(/[^\w.-]+/g, "_").slice(0, 40);
}

/** After a successful DOCX write, the written package becomes the base for further saves. */
function rebaseLoaded(bytes: Uint8Array) {
  const reloaded = loadDocx(bytes);
  app.loaded = { ...reloaded, doc: view().state.doc };
}

/** Write a recovery copy of a dirty document (called every minute). */
async function autosaveTick() {
  if (!F.isTauri || !app.dirty || app.settings.autosave === false || !app.handle) return;
  const dir = await F.recoveryDir();
  if (!dir) return;
  if (!app.recoveryId) app.recoveryId = recoveryIdFor(app.path);
  const kind: "docx" | "md" = app.kind === "md" ? "md" : "docx";
  try {
    let bytes: Uint8Array;
    if (kind === "md") {
      const text = app.source ? sourceTextarea().value : docToMarkdown(view().state.doc, () => null).markdown;
      bytes = new TextEncoder().encode(text);
    } else {
      if (!app.loaded) app.loaded = await loadBlank();
      bytes = writeDocx(app.loaded, view().state.doc);
      rebaseLoaded(bytes);
    }
    await F.writeFile(F.joinPath(dir, app.recoveryId + "." + kind), bytes);
    await F.writeFile(F.joinPath(dir, app.recoveryId + ".json"), new TextEncoder().encode(JSON.stringify({ path: app.path, name: docName(), savedAt: Date.now(), kind })));
    app.status?.flash("Recovery copy saved");
  } catch (e) { console.warn("autosave failed", e); }
}

async function clearRecovery() {
  if (!F.isTauri || !app.recoveryId) return;
  const dir = await F.recoveryDir();
  if (!dir) return;
  for (const ext of ["docx", "md", "json"]) await F.deleteFile(F.joinPath(dir, app.recoveryId + "." + ext));
  app.recoveryId = null;
}

async function discardRecovery(r: RecoveryEntry) {
  const dir = await F.recoveryDir();
  if (!dir) return;
  for (const ext of ["docx", "md", "json"]) await F.deleteFile(F.joinPath(dir, r.id + "." + ext));
}

/** Recovery copies whose original was not saved afterwards. */
async function findRecoveries(): Promise<RecoveryEntry[]> {
  if (!F.isTauri) return [];
  const dir = await F.recoveryDir();
  if (!dir) return [];
  const files = await F.listFiles(dir);
  const out: RecoveryEntry[] = [];
  for (const f of files) {
    if (!f.name.endsWith(".json")) continue;
    try {
      const meta = JSON.parse(await F.readTextFile(f.path));
      const id = f.name.slice(0, -5);
      const data = files.find((x) => x.name === id + "." + meta.kind);
      if (!data) { await F.deleteFile(f.path); continue; }
      // Stale if the original has been saved since the recovery copy was written.
      if (meta.path) {
        const mt = await F.fileMtime(meta.path);
        if (mt !== null && mt * 1000 >= meta.savedAt - 1500) { await discardRecovery({ id, path: meta.path, name: "", file: data.path, savedAt: 0, kind: meta.kind }); continue; }
      }
      out.push({ id, path: meta.path || null, name: meta.name || data.name, file: data.path, savedAt: meta.savedAt || f.mtime * 1000, kind: meta.kind === "md" ? "md" : "docx" });
    } catch { /* ignore broken sidecar */ }
  }
  return out.sort((a, b) => b.savedAt - a.savedAt);
}

async function openRecovered(r: RecoveryEntry) {
  await openPath(r.file);
  app.path = r.path;
  app.kind = r.path ? fileKind(r.path) : "new";
  app.recoveryId = r.id;
  if (r.path) { const list = (app.settings.recent || []).filter((p) => p !== r.file); app.settings.recent = list; F.saveSettings(app.settings); }
  setDirty(true);
  updateTitle();
  app.status?.flash("Recovered copy opened - save to keep it");
}

/** When a document opens and a newer recovery copy of it exists, offer it. */
async function offerRecoveryFor(path: string) {
  const entries = await findRecoveries();
  const r = entries.find((e) => e.path === path);
  if (!r) return;
  showDialog("Recovered version found", el("div", null,
    el("p", null, `An autosaved copy of "${F.basename(path)}" from ${new Date(r.savedAt).toLocaleString()} is newer than the file on disk.`),
    el("p", { style: { color: "var(--ui-muted)" } }, "Open the recovered copy, or discard it and keep the file as it is.")), [
    { label: "Discard copy", action: () => { discardRecovery(r); } },
    { label: "Open recovered copy", primary: true, action: () => { openRecovered(r); } },
  ]);
}

async function openPath(path: string) {
  const t0 = performance.now();
  app.loading = true;
  try {
    const ext = F.extname(path);
    const bytes = await F.readFile(path);
    if (ext === "docx" || ext === "docm" || ext === "dotx") {
      const loaded = loadDocx(bytes);
      if (app.source) { app.source = false; $("workspace").classList.remove("source"); }
      installDocument(loaded, loaded.doc, path, "docx");
      loadDeferredImages(loaded);
    } else if (ext === "md" || ext === "markdown" || ext === "txt") {
      const blank = await loadBlank();
      const text = new TextDecoder("utf-8").decode(bytes);
      const doc = ext === "txt" ? textToDoc(text) : markdownToDoc(text);
      if (app.source) { app.source = false; $("workspace").classList.remove("source"); }
      installDocument(blank, doc, path, "md");
      if (ext !== "txt") await resolveRelativeImages(F.dirname(path));
    } else {
      throw new Error("Unsupported file type: ." + ext);
    }
    addRecent(path);
    hideWelcome();
    app.recoveryId = null;
    app.status?.flash(`Opened in ${Math.round(performance.now() - t0)} ms`);
    offerRecoveryFor(path);
  } catch (e) {
    console.error(e);
    await F.showMessage("Could not open the file.\n\n" + (e as Error).message, "OfficeMini", "error");
  } finally {
    app.loading = false;
  }
}

/** Large images are inflated after first paint; swap their placeholders for the real bitmaps. */
function loadDeferredImages(loaded: LoadedDoc) {
  if (!Array.from(loaded.media.values()).some((m) => m.deferred)) return;
  resolveDeferredMedia(loaded, (rId, m) => {
    if (app.loaded !== loaded) return;
    const v = view();
    const tr = v.state.tr;
    v.state.doc.descendants((node, pos) => {
      if (node.type === schema.nodes.image && node.attrs.rId === rId && node.attrs.src !== m.url) tr.setNodeMarkup(pos, undefined, { ...node.attrs, src: m.url });
      return true;
    });
    if (tr.docChanged) { app.loading = true; v.dispatch(tr.setMeta("addToHistory", false)); app.loading = false; }
  });
}

function mediaBytes(a: any): { bytes: Uint8Array; ext: string } | null {
  if (a.media) return { bytes: a.media.bytes, ext: a.media.ext };
  if (a.rId && app.loaded) {
    const m = app.loaded.media.get(a.rId);
    if (m) {
      const bytes = m.bytes || app.loaded.pkg.get(m.part) || null;
      if (bytes) return { bytes, ext: m.ext };
    }
  }
  return null;
}

function textToDoc(text: string): PMNode {
  const paras = text.replace(/\r\n?/g, "\n").split("\n").map((line) => {
    const inline: PMNode[] = [];
    line.split("\t").forEach((seg, i) => { if (i > 0) inline.push(schema.nodes.tab.create()); if (seg) inline.push(schema.text(seg)); });
    return schema.nodes.paragraph.create({ props: {} }, inline);
  });
  return schema.nodes.doc.create({ sect: app.handle ? view().state.doc.attrs.sect : undefined }, paras.length ? paras : [schema.nodes.paragraph.create()]);
}

async function newDocument() {
  if (app.dirty || app.path) { await F.openInNewWindow(); return; }
  const blank = await loadBlank();
  installDocument(blank, blank.doc, null, "new");
  app.recoveryId = null;
  hideWelcome();
}

async function openFile() {
  const paths = await F.openDialog();
  if (!paths || !paths.length) return;
  // Reuse this window only when it holds an untouched empty document; otherwise open new windows.
  const reusable = !app.dirty && app.kind === "new" && !C.docHasContent(view().state.doc);
  for (let i = 0; i < paths.length; i++) {
    if (i === 0 && reusable) await openPath(paths[0]);
    else await F.openInNewWindow(paths[i]);
  }
}

// ---------------------------------------------------------------------------
// Saving

async function save(): Promise<boolean> {
  if (!app.path || app.kind === "new") return saveAs();
  return writeTo(app.path);
}

async function saveAs(): Promise<boolean> {
  const defaultName = app.path ? F.basename(app.path) : "Document.docx";
  const filters = app.kind === "md"
    ? [{ name: "Markdown", extensions: ["md"] }, { name: "Word Document", extensions: ["docx"] }]
    : [{ name: "Word Document", extensions: ["docx"] }, { name: "Markdown", extensions: ["md"] }];
  const target = await F.saveDialog(app.path ? app.path : defaultName, filters);
  if (!target) return false;
  const ok = await writeTo(target);
  if (ok) { app.path = target; app.kind = F.extname(target) === "md" ? "md" : "docx"; updateTitle(); addRecent(target); }
  return ok;
}

async function writeTo(path: string): Promise<boolean> {
  const t0 = performance.now();
  try {
    const ext = F.extname(path);
    if (app.source && (ext === "md" || ext === "markdown")) {
      // Source mode: the text area is the truth; write it as-is (plus any images it references).
      const text = sourceTextarea().value;
      await F.writeFile(path, new TextEncoder().encode(text));
      const dir = F.dirname(path);
      for (const [name, m] of sourceMediaMap) if (text.includes(name)) await F.writeFile(F.joinPath(dir, name), m.bytes);
      setDirty(false);
      app.status?.flash(`Saved in ${Math.round(performance.now() - t0)} ms`);
      return true;
    }
    if (app.source) leaveSourceView();
    const doc = view().state.doc;
    if (ext === "md" || ext === "markdown") {
      const dir = F.dirname(path);
      const base = F.basename(path).replace(/\.(md|markdown)$/i, "");
      let n = 0;
      const res = docToMarkdown(doc, (node) => {
        const mb = mediaBytes(node.attrs);
        if (!mb) return null;
        n++;
        return { name: `${base}_files/image${n}.${mb.ext}`, bytes: mb.bytes };
      });
      await F.writeFile(path, new TextEncoder().encode(res.markdown));
      for (const a of res.assets) await F.writeFile(F.joinPath(dir, a.name.replace("/", dir.includes("\\") ? "\\" : "/")), a.bytes);
    } else {
      if (!app.loaded) app.loaded = await loadBlank();
      const bytes = writeDocx(app.loaded, doc);
      await F.writeFile(path, bytes);
      // The written package is now the base for future saves (fresh rels/media ids).
      rebaseLoaded(bytes);
    }
    setDirty(false);
    clearRecovery();
    app.status?.flash(`Saved in ${Math.round(performance.now() - t0)} ms`);
    return true;
  } catch (e) {
    console.error(e);
    await F.showMessage("Could not save the file.\n\n" + (e as Error).message, "OfficeMini", "error");
    return false;
  }
}

/** Returns true when it is OK to discard the current document. */
async function confirmDiscard(): Promise<boolean> {
  if (!app.dirty) return true;
  return new Promise((resolve) => {
    let decided = false;
    showDialog("Save changes?", el("div", null, el("p", null, `Save changes to "${docName()}"?`), el("p", { style: { color: "var(--ui-muted)" } }, "Your changes will be lost if you don't save them.")), [
      { label: "Cancel", action: () => { decided = true; resolve(false); } },
      { label: "Don't Save", action: () => { decided = true; clearRecovery(); resolve(true); } },
      { label: "Save", primary: true, action: () => { decided = true; save().then(resolve); } },
    ], { onClose: () => { if (!decided) resolve(false); } });
  });
}

// ---------------------------------------------------------------------------
// Recent files

function addRecent(path: string) {
  const list = (app.settings.recent || []).filter((p) => p !== path);
  list.unshift(path);
  app.settings.recent = list.slice(0, 12);
  F.saveSettings(app.settings);
}

// ---------------------------------------------------------------------------
// Zoom / view

let zoomSaveTimer: ReturnType<typeof setTimeout> | null = null;
/** Apply a zoom factor; `fine` keeps two decimals (pinch) instead of snapping to whole percents. */
function setZoom(z: number, fine = false) {
  z = Math.max(0.5, Math.min(3, fine ? Math.round(z * 1000) / 1000 : Math.round(z * 100) / 100));
  if (z === app.zoom) return;
  app.zoom = z;
  document.documentElement.style.setProperty("--zoom", String(z));
  setZoomFactor(z);
  app.settings.zoom = z;
  // A pinch or wheel burst produces dozens of events; write the settings file once it settles.
  if (zoomSaveTimer) clearTimeout(zoomSaveTimer);
  zoomSaveTimer = setTimeout(() => { zoomSaveTimer = null; F.saveSettings(app.settings); }, 400);
  updateStatus();
}

function setMode(mode: "page" | "continuous") {
  app.mode = mode;
  $("workspace").classList.toggle("continuous", mode === "continuous");
  setViewMode(view(), mode);
  app.settings.view = mode;
  F.saveSettings(app.settings);
  updateStatus();
}

function toggleMarks() {
  app.showMarks = !app.showMarks;
  $("editor").classList.toggle("show-marks", app.showMarks);
  app.settings.showMarks = app.showMarks;
  F.saveSettings(app.settings);
}

/** Switch the colour theme. Document colours are computed in toDOM, so the view is redrawn. */
function setTheme(theme: "light" | "dark", persist = true) {
  app.theme = theme;
  document.documentElement.setAttribute("data-theme", theme);
  setDarkMode(theme === "dark");
  clearCssCaches(); // run styles are cached per mark and depend on the theme
  if (app.handle) {
    redrawView(view());
    // Headers/footers are rendered from a cache with baked-in colours: rebuild them.
    resetHeaderFooterCache();
    $("pages").removeAttribute("data-key");
    $("pages").removeAttribute("data-hf");
    view().dispatch(view().state.tr.setMeta(numberingKey, "refresh").setMeta("addToHistory", false));
    relayout(view());
  }
  if (persist) { app.settings.theme = theme; F.saveSettings(app.settings); }
}
function toggleTheme() { setTheme(app.theme === "dark" ? "light" : "dark"); }

function updateStatus() {
  if (!app.status || !app.handle) return;
  const state = view().state;
  const { words } = C.countWords(state.doc);
  const sel = state.selection;
  let selWords = 0;
  if (!sel.empty) { const m = state.doc.textBetween(sel.from, sel.to, " ", " ").match(/[^\s ]+/g); selWords = m ? m.length : 0; }
  app.status.update({ page: pageAt(state, state.selection.from), pages: app.pages, words, selWords, zoom: app.zoom, mode: app.mode, dirty: app.dirty, isMd: app.kind === "md", source: app.source, dark: app.theme === "dark" });
}

// ---------------------------------------------------------------------------
// Find bar

let findVisible = false;
let findReplaceMode = false;
function buildFindbar() {
  const bar = $("findbar");
  bar.innerHTML = "";
  const opts: FindOptions = { ...DEFAULT_FIND_OPTIONS, ...(app.settings.findOptions || {}) };
  let inSelection = false;
  const input = el("input", { type: "text", placeholder: "Find", "aria-label": "Find" });
  const count = el("span", { class: "count" }, "");
  const prev = el("button", { class: "tb-btn", type: "button" }, "▲");
  const next = el("button", { class: "tb-btn", type: "button" }, "▼");
  tooltip(prev, "Previous match", "Shift+Enter / Shift+F3");
  tooltip(next, "Next match", "Enter / F3");
  const cb = (label: string, key: string, get: () => boolean, set: (v: boolean) => void) => {
    const c = el("input", { type: "checkbox" });
    c.checked = get();
    c.addEventListener("change", () => { set(c.checked); persist(); apply(); refresh(); });
    const l = el("label", null, c, label);
    tooltip(l, label, key);
    return { el: l, input: c };
  };
  const caseCb = cb("Match case", "Alt+C", () => opts.caseSensitive, (v) => (opts.caseSensitive = v));
  const wordCb = cb("Whole word", "Alt+W", () => opts.wholeWord, (v) => (opts.wholeWord = v));
  const regexCb = cb("Regex", "Alt+R", () => opts.regex, (v) => (opts.regex = v));
  const selCb = cb("In selection", "Alt+S", () => inSelection, (v) => { inSelection = v; });
  const close = el("button", { class: "tb-btn", type: "button" }, "✕");
  tooltip(close, "Close", "Esc");
  const replaceIn = el("input", { type: "text", placeholder: "Replace with", "aria-label": "Replace with" });
  const replBtn = el("button", { class: "tb-btn textbtn", type: "button" }, "Replace");
  const replAllBtn = el("button", { class: "tb-btn textbtn", type: "button" }, "Replace all");
  tooltip(replBtn, "Replace this match and go to the next", "Enter (in the replace box)");
  tooltip(replAllBtn, "Replace every match", "Ctrl+Enter / Alt+A");
  const caseKeepCb = cb("Preserve case", "Alt+P", () => opts.preserveCase, (v) => (opts.preserveCase = v));
  const replaceRow = el("span", { style: { display: "inline-flex", gap: "6px", alignItems: "center" } }, replaceIn, replBtn, replAllBtn, caseKeepCb.el);
  const hint = el("span", { class: "hint" }, "");
  bar.append(input, prev, next, count, caseCb.el, wordCb.el, regexCb.el, selCb.el, replaceRow, hint, el("span", { style: { flex: "1" } }), close);
  const persist = () => { app.settings.findOptions = { ...opts }; F.saveSettings(app.settings); };
  let timer = 0;
  const scope = () => {
    if (!inSelection) return null;
    const sel = view().state.selection;
    return sel.empty ? null : { from: sel.from, to: sel.to };
  };
  let lockedScope: { from: number; to: number } | null = null;
  const apply = () => {
    if (app.source) { refresh(); return; }
    if (inSelection && !lockedScope) lockedScope = scope();
    if (!inSelection) lockedScope = null;
    setFindQuery(view(), { query: input.value, opts: { ...opts }, scope: lockedScope });
  };
  const step = (dir: 1 | -1) => {
    if (app.source) {
      const r = textareaFind(sourceTextarea(), input.value, opts, dir);
      count.textContent = r.error ? "Invalid pattern" : r.count ? `${r.index + 1} of ${r.count}` : (input.value ? "No results" : "");
      hint.textContent = r.wrapped ? (dir > 0 ? "Wrapped to top" : "Wrapped to end") : "";
      input.focus();
      return;
    }
    const r = findStep(view(), dir);
    hint.textContent = r.wrapped ? (dir > 0 ? "Wrapped to top" : "Wrapped to end") : "";
    refresh();
  };
  input.addEventListener("input", () => { clearTimeout(timer); timer = window.setTimeout(() => { apply(); refresh(); }, 80); });
  const keys = (e: KeyboardEvent) => {
    if (e.altKey && !e.ctrlKey) {
      const k = e.key.toLowerCase();
      const map: Record<string, HTMLInputElement> = { c: caseCb.input, w: wordCb.input, r: regexCb.input, s: selCb.input, p: caseKeepCb.input };
      if (map[k]) { e.preventDefault(); map[k].checked = !map[k].checked; map[k].dispatchEvent(new Event("change")); return true; }
      if (k === "a") { e.preventDefault(); doReplaceAll(); return true; }
    }
    if (e.key === "Escape") { e.preventDefault(); hideFind(); return true; }
    if (e.key === "F3") { e.preventDefault(); step(e.shiftKey ? -1 : 1); return true; }
    return false;
  };
  input.addEventListener("keydown", (e) => {
    if (keys(e)) return;
    if (e.key === "Enter") { e.preventDefault(); if (e.ctrlKey) doReplaceAll(); else step(e.shiftKey ? -1 : 1); }
    else if (e.key === "Tab" && !e.shiftKey && findReplaceMode) { e.preventDefault(); replaceIn.focus(); replaceIn.select(); }
  });
  replaceIn.addEventListener("keydown", (e) => {
    if (keys(e)) return;
    if (e.key === "Enter") { e.preventDefault(); if (e.ctrlKey) doReplaceAll(); else doReplace(); }
    else if (e.key === "Tab" && e.shiftKey) { e.preventDefault(); input.focus(); input.select(); }
  });
  const doReplace = () => {
    if (app.source) { const n = textareaReplace(sourceTextarea(), input.value, replaceIn.value, opts, false); if (!n) step(1); else step(1); return; }
    replaceCurrent(view(), replaceIn.value);
    refresh();
  };
  const doReplaceAll = () => {
    const n = app.source ? textareaReplace(sourceTextarea(), input.value, replaceIn.value, opts, true) : replaceAll(view(), replaceIn.value);
    app.status?.flash(`Replaced ${n} occurrence${n === 1 ? "" : "s"}`);
    hint.textContent = `Replaced ${n}`;
    refresh();
  };
  prev.addEventListener("click", () => step(-1));
  next.addEventListener("click", () => step(1));
  replBtn.addEventListener("click", doReplace);
  replAllBtn.addEventListener("click", doReplaceAll);
  close.addEventListener("click", hideFind);
  const refresh = () => {
    if (app.source) { if (!input.value) count.textContent = ""; return; }
    const st = getFind(view().state);
    input.classList.toggle("invalid", !!st.error);
    if (!st.query) { count.textContent = ""; hint.textContent = ""; return; }
    count.textContent = st.error ? "Invalid pattern" : st.matches.length ? `${st.current + 1} of ${st.matches.length}` : "No results";
    if (st.error) hint.textContent = st.error;
  };
  return {
    show(replace: boolean) {
      findVisible = true; findReplaceMode = replace;
      bar.hidden = false;
      replaceRow.style.display = replace ? "inline-flex" : "none";
      hint.textContent = "";
      if (!app.source) {
        const sel = view().state.selection;
        const selected = !sel.empty && sel.to - sel.from < 200 ? view().state.doc.textBetween(sel.from, sel.to, " ") : "";
        // A multi-line selection becomes the search scope; a short one becomes the query.
        if (!sel.empty && (sel.to - sel.from >= 200 || selected.includes("\n"))) { inSelection = true; selCb.input.checked = true; lockedScope = { from: sel.from, to: sel.to }; }
        else if (selected.trim()) { input.value = selected; inSelection = false; selCb.input.checked = false; lockedScope = null; }
      } else {
        const ta = sourceTextarea();
        const s = ta.value.slice(ta.selectionStart, ta.selectionEnd);
        if (s && s.length < 200 && !s.includes("\n")) input.value = s;
      }
      apply();
      refresh();
      (replace && input.value ? replaceIn : input).focus();
      (replace && input.value ? replaceIn : input).select();
    },
    hide() {
      findVisible = false; bar.hidden = true; inSelection = false; lockedScope = null; selCb.input.checked = false;
      if (app.source) sourceTextarea().focus(); else { closeFind(view()); view().focus(); }
    },
    refresh,
    step,
  };
}
let findbar: ReturnType<typeof buildFindbar>;
function showFind(replace: boolean) { findbar.show(replace); }
function hideFind() { findbar.hide(); }

// ---------------------------------------------------------------------------
// Insert helpers

async function insertImageFile() {
  const paths = await F.openDialog([{ name: "Images", extensions: ["png", "jpg", "jpeg", "gif", "bmp", "webp", "svg"] }]);
  if (!paths || !paths[0]) return;
  const bytes = await F.readFile(paths[0]);
  await insertImageBytes(bytes, F.extname(paths[0]) || "png", F.basename(paths[0]));
}

async function insertImageBytes(bytes: Uint8Array, ext: string, name = "Picture") {
  ext = ext.toLowerCase().replace("jpg", "jpeg");
  const mime = ext === "svg" ? "image/svg+xml" : "image/" + ext;
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime }));
  let w = 300, h = 200;
  try { const s = await imageSize(url); w = s.w; h = s.h; } catch { /* keep defaults */ }
  const sect = view().state.doc.attrs.sect as SectProps;
  const maxW = twipsToPx(sect.pgW - sect.marL - sect.marR);
  if (w > maxW) { h = h * (maxW / w); w = maxW; }
  const node = schema.nodes.image.create({ src: url, w: Math.round(w), h: Math.round(h), rId: null, raw: null, kind: "inline", name, alt: name, media: { ext, bytes }, ext, origW: Math.round(w), origH: Math.round(h) });
  view().dispatch(view().state.tr.replaceSelectionWith(node).scrollIntoView());
  view().focus();
}

async function insertTableDialog() {
  const r = await tableDialog();
  if (!r) return;
  run(T.insertTable(r.rows, r.cols));
}

async function insertLink() {
  const state = view().state;
  const existing = C.linkAtSelection(state);
  const sel = state.selection;
  const selectedText = sel.empty ? "" : state.doc.textBetween(sel.from, sel.to, " ");
  const r = await linkDialog(existing?.attrs.href || "", selectedText, !sel.empty);
  if (!r) return;
  if (sel.empty && !existing) {
    const text = r.text || r.href;
    const tr = state.tr.insertText(text, sel.from, sel.to);
    tr.addMark(sel.from, sel.from + text.length, schema.marks.link.create({ href: r.href }));
    view().dispatch(tr);
  } else if (existing && sel.empty) {
    // Update the link mark around the cursor.
    run(C.unsetLink, false);
    const s2 = view().state;
    // re-find the range: unsetLink removed the mark; simplest is to re-add over the word range
    const $f = s2.selection.$from;
    const parent = $f.parent, base = $f.start();
    let from = $f.pos, to = $f.pos;
    parent.forEach((child, off) => { const s = base + off, e = s + child.nodeSize; if (s <= $f.pos && e >= $f.pos && child.isText) { from = s; to = e; } });
    view().dispatch(s2.tr.addMark(from, to, schema.marks.link.create({ href: r.href })));
  } else {
    run(C.setLink(r.href), false);
  }
  view().focus();
}

async function showParagraphDialog() {
  const eff = C.selectionParaProps(view().state).eff;
  const r = await paragraphDialog({
    before: Math.round((eff.spBefore || 0) / 20), after: Math.round((eff.spAfter || 0) / 20),
    line: eff.spLine ?? 240, rule: eff.spLineRule || "auto",
    left: Math.round((eff.indLeft || 0) / 20), right: Math.round((eff.indRight || 0) / 20),
    firstLine: Math.round((eff.indFirstLine || 0) / 20), hanging: Math.round((eff.indHanging || 0) / 20),
  });
  if (!r) return;
  run(C.setParaProps(r));
}

async function showPageSetup() {
  const sect = view().state.doc.attrs.sect as SectProps;
  const r = await pageSetupDialog(sect);
  if (!r) return;
  view().dispatch(view().state.tr.setDocAttribute("sect", r));
  setPrintPageSize(r);
  view().focus();
}

async function goToPage() {
  const r = await goToPageDialog(app.pages);
  if (!r) return;
  selectPos(view(), pageStartPos(view().state, r));
}

function cellShadingPopup() {
  const anchor = document.activeElement instanceof HTMLElement ? document.activeElement : $("toolbar");
  colorPopup(anchor, null, (hex) => run(T.setCellShading(hex)), { auto: "No color" });
}

// ---------------------------------------------------------------------------
// Printing

async function print() {
  const prevMode = app.mode, prevZoom = app.zoom, prevTheme = app.theme;
  if (app.source) leaveSourceView();
  await printDocument(async () => {
    closeAllPopups();
    if (app.theme !== "light") setTheme("light", false);
    if (app.mode !== "page") setMode("page");
    if (app.zoom !== 1) setZoom(1);
    setPrintPageSize(view().state.doc.attrs.sect as SectProps);
    relayout(view());
    await new Promise((r) => setTimeout(r, 160));
  }, () => {
    if (prevTheme !== app.theme) setTheme(prevTheme, false);
    if (prevMode !== app.mode) setMode(prevMode);
    if (prevZoom !== app.zoom) setZoom(prevZoom);
  });
}

// ---------------------------------------------------------------------------
// Markdown source view (raw text editing of .md files)

let sourceEl: HTMLTextAreaElement | null = null;
let sourceMediaMap = new Map<string, { bytes: Uint8Array; ext: string; src: string }>();

function sourceTextarea(): HTMLTextAreaElement {
  if (!sourceEl) {
    sourceEl = document.getElementById("source") as HTMLTextAreaElement;
    sourceEl.addEventListener("input", () => setDirty(true));
    sourceEl.addEventListener("keydown", (e) => {
      if (e.key === "Tab" && !e.ctrlKey && !e.altKey) {
        e.preventDefault();
        const s = sourceEl!.selectionStart, en = sourceEl!.selectionEnd;
        if (e.shiftKey) {
          // outdent current line
          const lineStart = sourceEl!.value.lastIndexOf("\n", s - 1) + 1;
          const line = sourceEl!.value.slice(lineStart, en);
          const m = /^( {1,2}|\t)/.exec(line);
          if (m) { sourceEl!.setRangeText("", lineStart, lineStart + m[0].length, "end"); sourceEl!.setSelectionRange(Math.max(lineStart, s - m[0].length), Math.max(lineStart, en - m[0].length)); }
        } else sourceEl!.setRangeText("  ", s, en, "end");
        setDirty(true);
      } else if (e.key === "Enter" && !e.shiftKey && !e.ctrlKey) {
        // continue list markers
        const s = sourceEl!.selectionStart;
        const lineStart = sourceEl!.value.lastIndexOf("\n", s - 1) + 1;
        const line = sourceEl!.value.slice(lineStart, s);
        const m = /^(\s*)([-*+]|\d+\.)\s+(.*)$/.exec(line);
        if (m) {
          e.preventDefault();
          if (!m[3].trim()) { sourceEl!.setRangeText("", lineStart, s, "end"); }
          else {
            const marker = /^\d+\.$/.test(m[2]) ? `${parseInt(m[2], 10) + 1}.` : m[2];
            sourceEl!.setRangeText(`\n${m[1]}${marker} `, s, sourceEl!.selectionEnd, "end");
          }
          setDirty(true);
        }
      }
    });
  }
  return sourceEl;
}

/** Serialize the live document to Markdown for the source editor, remembering image bytes by name. */
function docToSource(): string {
  sourceMediaMap = new Map();
  let n = 0;
  const res = docToMarkdown(view().state.doc, (node) => {
    const a = node.attrs;
    const mb = mediaBytes(a);
    if (!mb) return null;
    n++;
    const name = a.name && /\.(png|jpe?g|gif|svg|webp)$/i.test(a.name) ? a.name : `image${n}.${mb.ext}`;
    sourceMediaMap.set(name, { bytes: mb.bytes, ext: mb.ext, src: a.src });
    return { name, bytes: mb.bytes };
  });
  sourceBlockLines = res.blockLines;
  lastSourceText = res.markdown;
  return res.markdown;
}

/** Parse Markdown from the source editor back into the document, restoring known images. */
function sourceToDoc(text: string): PMNode {
  const doc = markdownToDoc(text);
  const media = sourceMediaMap;
  if (!media.size) return doc;
  const json = doc.toJSON();
  const fix = (node: any) => {
    if (node.type === "image" && node.attrs && media.has(node.attrs.src)) {
      const m = media.get(node.attrs.src)!;
      node.attrs = { ...node.attrs, src: m.src, media: { bytes: m.bytes, ext: m.ext }, ext: m.ext };
    }
    if (node.content) node.content.forEach(fix);
  };
  fix(json);
  return schema.nodeFromJSON(json);
}

let sourceBlockLines: number[] = [];

function enterSourceView() {
  if (app.source) return;
  if (findVisible) hideFind();
  const ta = sourceTextarea();
  const v = view();
  // Which top-level block holds the caret? Put the text cursor on the matching line.
  const $from = v.state.selection.$from;
  const blockIndex = $from.depth ? $from.index(0) : 0;
  ta.value = docToSource();
  const line = sourceBlockLines[Math.min(blockIndex, sourceBlockLines.length - 1)] ?? 0;
  app.source = true;
  $("workspace").classList.add("source");
  $("toolbar").classList.add("disabled");
  ta.focus();
  let offset = 0;
  const parts = ta.value.split("\n");
  for (let i = 0; i < line && i < parts.length; i++) offset += parts[i].length + 1;
  ta.setSelectionRange(offset, offset);
  const lh = parseFloat(getComputedStyle(ta).lineHeight) || 20;
  ta.scrollTop = Math.max(0, line * lh - ta.clientHeight / 3);
  updateStatus();
}

function leaveSourceView() {
  if (!app.source) return;
  const ta = sourceTextarea();
  const wasDirty = app.dirty;
  // Caret line -> block index so the rendered view opens at the same place.
  const caretLine = ta.value.slice(0, ta.selectionStart).split("\n").length - 1;
  let blockIndex = 0;
  for (let i = 0; i < sourceBlockLines.length; i++) if (sourceBlockLines[i] <= caretLine) blockIndex = i;
  app.loading = true;
  try {
    const doc = sourceToDoc(ta.value);
    const v = view();
    // Replace the content in one undoable step (keeps history and plugin state).
    const tr = v.state.tr.replaceWith(0, v.state.doc.content.size, doc.content);
    let pos = 1;
    if (blockIndex < tr.doc.childCount) { pos = 0; for (let i = 0; i < blockIndex; i++) pos += tr.doc.child(i).nodeSize; pos += 1; }
    tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(pos, tr.doc.content.size))));
    v.dispatch(tr.scrollIntoView());
  } finally { app.loading = false; }
  app.source = false;
  $("workspace").classList.remove("source");
  $("toolbar").classList.remove("disabled");
  setDirty(wasDirty || ta.value !== lastSourceText);
  view().focus();
  updateStatus();
}
let lastSourceText = "";

function toggleSourceView() {
  if (app.source) { leaveSourceView(); return; }
  if (app.kind !== "md") {
    F.showMessage("Markdown source view is available for Markdown files. Use File > Save as… to save this document as Markdown first.", "OfficeMini", "info");
    return;
  }
  enterSourceView();
}

/** Load images referenced by relative paths in a Markdown file. */
async function resolveRelativeImages(baseDir: string) {
  const v = view();
  const targets: { pos: number; node: PMNode }[] = [];
  v.state.doc.descendants((node, pos) => { if (node.type === schema.nodes.image && node.attrs.src && !/^(https?:|data:|blob:)/i.test(node.attrs.src)) targets.push({ pos, node }); return true; });
  if (!targets.length) return;
  const tr = v.state.tr;
  for (const t of targets) {
    const rel = decodeURIComponent(t.node.attrs.src).replace(/\//g, baseDir.includes("\\") ? "\\" : "/");
    const path = /^([a-z]:)?[\\/]/i.test(rel) ? rel : F.joinPath(baseDir, rel);
    try {
      const bytes = await F.readFile(path);
      const ext = F.extname(path) || "png";
      const mime = ext === "svg" ? "image/svg+xml" : "image/" + ext.replace("jpg", "jpeg");
      const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime }));
      let w = t.node.attrs.w, h = t.node.attrs.h;
      try { const s = await imageSize(url); const sect = v.state.doc.attrs.sect as SectProps; const maxW = twipsToPx(sect.pgW - sect.marL - sect.marR); w = s.w; h = s.h; if (w > maxW) { h = h * (maxW / w); w = maxW; } } catch { /* keep */ }
      tr.setNodeMarkup(t.pos, undefined, { ...t.node.attrs, src: url, media: { bytes, ext }, ext, w: Math.round(w), h: Math.round(h), name: F.basename(path), origW: Math.round(w), origH: Math.round(h) });
    } catch { /* leave the broken reference visible */ }
  }
  if (tr.docChanged) { app.loading = true; v.dispatch(tr.setMeta("addToHistory", false)); app.loading = false; setDirty(false); }
}

// ---------------------------------------------------------------------------
// Actions exposed to keymap / toolbar / menus

const actions: ContextActions = {
  newDocument: () => { newDocument(); },
  openFile: () => { openFile(); },
  save: () => { save(); },
  saveAs: () => { saveAs(); },
  print: () => { print(); },
  closeWindow: () => { closeWindowRequest(); },
  find: () => showFind(false),
  replace: () => showFind(true),
  findNext: () => { if (!findVisible) showFind(false); else findbar.step(1); },
  findPrev: () => { if (!findVisible) showFind(false); else findbar.step(-1); },
  closeFind: () => {
    if (findVisible) { hideFind(); return true; }
    closeAllPopups(); closeDialog();
    const sel = view().state.selection;
    if (sel instanceof NodeSelection || !sel.empty) { view().dispatch(view().state.tr.setSelection(TextSelection.create(view().state.doc, sel.to))); return true; }
    return false;
  },
  zoomIn: () => setZoom(app.zoom + 0.1),
  zoomOut: () => setZoom(app.zoom - 0.1),
  zoomReset: () => setZoom(1),
  toggleMarks,
  toggleFullscreen: () => F.toggleFullscreen(),
  showShortcuts: () => shortcutsDialog(app.shortcuts),
  focusFontFamily: () => app.toolbar?.focusFont(),
  focusFontSize: () => app.toolbar?.focusSize(),
  insertLink: () => { insertLink(); },
  insertImage: () => { insertImageFile(); },
  insertTable: () => { insertTableDialog(); },
  goToPage: () => { goToPage(); },
  pastePlain: () => { requestPlainPaste(); pasteFromClipboard(view(), true); },
  selectAllCmd: () => run(selectAll),
  paragraphDialog: () => { showParagraphDialog(); },
  editLink: () => { insertLink(); },
  cellShading: () => cellShadingPopup(),
  insertImageFromBytes: (bytes, ext) => { insertImageBytes(bytes, ext); },
  toggleTheme: () => toggleTheme(),
  toggleSource: () => { toggleSourceView(); },
};

async function closeWindowRequest() {
  if (await confirmDiscard()) F.closeWindow();
}

// ---------------------------------------------------------------------------
// Menu bar

function buildMenubar() {
  const bar = $("menubar");
  bar.innerHTML = "";
  const key = (id: string) => { const s = app.shortcuts.find((x) => x.id === id); return s ? keyLabel(s.keys[0]) : undefined; };
  const menus: { title: string; alt: string; items: () => MenuItem[] }[] = [
    { title: "File", alt: "f", items: () => [
      { label: "New", key: key("new"), action: () => newDocument() },
      { label: "New window", action: () => F.openInNewWindow() },
      { label: "Open…", key: key("open"), action: () => openFile() },
      { label: "Open recent", submenu: (app.settings.recent || []).length ? (app.settings.recent || []).map((p) => ({ label: F.basename(p), action: () => (!app.dirty && (!app.path || app.kind !== "new") ? openPath(p) : F.openInNewWindow(p)) })) : [{ label: "(empty)", disabled: true }] },
      { sep: true },
      { label: "Save", key: key("save"), action: () => save() },
      { label: "Save as…", key: key("saveas"), action: () => saveAs() },
      { sep: true },
      { label: "Page setup…", action: () => showPageSetup() },
      { label: "Print…", key: key("print"), action: () => print() },
      { sep: true },
      { label: "Close window", key: key("close"), action: () => closeWindowRequest() },
    ] },
    { title: "Edit", alt: "e", items: () => [
      { label: "Undo", key: key("undo"), action: () => run(undo) },
      { label: "Redo", key: key("redo"), action: () => run(redo) },
      { sep: true },
      { label: "Cut", key: "Ctrl+X", action: () => { view().focus(); document.execCommand("cut"); } },
      { label: "Copy", key: "Ctrl+C", action: () => { view().focus(); document.execCommand("copy"); } },
      { label: "Paste", key: "Ctrl+V", action: () => pasteFromClipboard(view(), false, actions.insertImageFromBytes) },
      { label: "Paste without formatting", key: key("pasteplain"), action: () => pasteFromClipboard(view(), true) },
      { sep: true },
      { label: "Select all", key: key("selectall"), action: () => run(selectAll) },
      { sep: true },
      { label: "Find…", key: key("find"), action: () => showFind(false) },
      { label: "Find and replace…", key: key("replace"), action: () => showFind(true) },
      { label: "Go to page…", key: key("goto"), action: () => goToPage() },
    ] },
    { title: "View", alt: "v", items: () => [
      { label: "Page view", checked: app.mode === "page", action: () => setMode("page") },
      { label: "Continuous", checked: app.mode === "continuous", action: () => setMode("continuous") },
      { sep: true },
      { label: "Zoom in", key: key("zoominx"), action: () => setZoom(app.zoom + 0.1) },
      { label: "Zoom out", key: key("zoomout"), action: () => setZoom(app.zoom - 0.1) },
      { label: "Zoom 100%", key: key("zoom0"), action: () => setZoom(1) },
      { label: "Fit page width", action: () => fitWidth() },
      { sep: true },
      { label: "Dark mode", key: key("theme"), checked: app.theme === "dark", action: () => toggleTheme() },
      { label: "Markdown source", key: key("source"), checked: app.source, disabled: app.kind !== "md" && !app.source, action: () => toggleSourceView() },
      { sep: true },
      { label: "Formatting marks", key: key("marks"), checked: app.showMarks, action: () => toggleMarks() },
      { label: "Spell check", checked: view().dom.getAttribute("spellcheck") === "true", action: () => { const on = view().dom.getAttribute("spellcheck") === "true"; view().dom.setAttribute("spellcheck", on ? "false" : "true"); } },
      { label: "Full screen", key: key("fullscreen"), action: () => F.toggleFullscreen() },
    ] },
    { title: "Insert", alt: "i", items: () => [
      { label: "Image…", key: key("image"), action: () => insertImageFile() },
      { label: "Table…", key: key("table"), action: () => insertTableDialog() },
      { label: "Link…", key: key("link"), action: () => insertLink() },
      { sep: true },
      { label: "Page break", key: key("pagebreak"), action: () => run(C.insertPageBreak) },
      { label: "Line break", key: key("linebreak"), action: () => run(C.insertLineBreak) },
      { label: "Tab character", action: () => run(C.insertTab) },
      { sep: true },
      { label: "Non-breaking space", key: key("nbsp"), action: () => run(C.insertTextCmd(String.fromCharCode(0xa0))) },
      { label: "Non-breaking hyphen", key: key("nbhyphen"), action: () => run(C.insertTextCmd(String.fromCharCode(0x2011))) },
      { label: "Symbol", submenu: ["—", "–", "…", "•", "©", "®", "™", "€", "£", "₺", "°", "±", "×", "÷", "≤", "≥", "≠", "→", "←", "↑", "↓", "✓", "✗", "§", "¶"].map((s) => ({ label: s, action: () => run(C.insertTextCmd(s)) })) },
    ] },
    { title: "Format", alt: "o", items: () => {
      const rp = C.selectionRunProps(view().state);
      const pp = C.selectionParaProps(view().state).eff;
      return [
        { label: "Bold", key: key("bold"), checked: !!rp.b, action: () => run(C.toggleBold) },
        { label: "Italic", key: key("italic"), checked: !!rp.i, action: () => run(C.toggleItalic) },
        { label: "Underline", key: key("underline"), checked: !!rp.u && rp.u !== "none", action: () => run(C.toggleUnderline) },
        { label: "Strikethrough", key: key("strike"), checked: !!rp.strike, action: () => run(C.toggleStrike) },
        { label: "Superscript", key: key("sup"), checked: rp.vertAlign === "superscript", action: () => run(C.toggleSuperscript) },
        { label: "Subscript", key: key("sub"), checked: rp.vertAlign === "subscript", action: () => run(C.toggleSubscript) },
        { label: "All caps", checked: !!rp.caps, action: () => run(C.toggleCaps) },
        { label: "Small caps", checked: !!rp.smallCaps, action: () => run(C.toggleSmallCaps) },
        { sep: true },
        { label: "Font…", key: key("fontbox"), action: () => app.toolbar?.focusFont() },
        { label: "Paragraph…", action: () => showParagraphDialog() },
        { sep: true },
        { label: "Align", submenu: [
          { label: "Left", key: key("left"), checked: !pp.jc || pp.jc === "left", action: () => run(C.setAlign("left")) },
          { label: "Center", key: key("center"), checked: pp.jc === "center", action: () => run(C.setAlign("center")) },
          { label: "Right", key: key("right"), checked: pp.jc === "right", action: () => run(C.setAlign("right")) },
          { label: "Justify", key: key("justify"), checked: pp.jc === "both", action: () => run(C.setAlign("both")) },
        ] },
        { label: "Line spacing", submenu: [
          { label: "Single", key: key("single"), action: () => run(C.setLineSpacing(1)) },
          { label: "1.15", action: () => run(C.setLineSpacing(1.15)) },
          { label: "1.5", key: key("onehalf"), action: () => run(C.setLineSpacing(1.5)) },
          { label: "Double", key: key("double"), action: () => run(C.setLineSpacing(2)) },
        ] },
        { label: "Lists", submenu: [
          { label: "Bullets", key: key("bullets"), action: () => run(C.toggleList("bullet")) },
          { label: "Numbering", key: key("numbers"), action: () => run(C.toggleList("decimal")) },
          { label: "Increase indent", key: key("indent"), action: () => run(C.indentParagraphs(1)) },
          { label: "Decrease indent", key: key("outdent"), action: () => run(C.indentParagraphs(-1)) },
        ] },
        { label: "Style", submenu: [
          { label: "Normal", key: key("normal"), action: () => run(C.setParaStyle(null)) },
          { label: "Heading 1", key: key("h1"), action: () => run(C.setParaStyle(styleId("Heading1", "heading 1"))) },
          { label: "Heading 2", key: key("h2"), action: () => run(C.setParaStyle(styleId("Heading2", "heading 2"))) },
          { label: "Heading 3", key: key("h3"), action: () => run(C.setParaStyle(styleId("Heading3", "heading 3"))) },
          { label: "Title", action: () => run(C.setParaStyle(styleId("Title", "Title"))) },
          { label: "Quote", action: () => run(C.setParaStyle(styleId("Quote", "Quote"))) },
        ] },
        { sep: true },
        { label: "Clear formatting", key: key("clear"), action: () => run(C.clearFormatting) },
        { sep: true },
        { label: "Smart quotes and dashes while typing", checked: app.settings.smartQuotes !== false, action: () => { app.settings.smartQuotes = app.settings.smartQuotes === false; setSmartTyping(app.settings.smartQuotes); F.saveSettings(app.settings); } },
        { label: "Autosave recovery copy every minute", checked: app.settings.autosave !== false, action: () => { app.settings.autosave = app.settings.autosave === false; F.saveSettings(app.settings); } },
      ];
    } },
    { title: "Table", alt: "t", items: () => {
      const inTable = C.isInTable(view().state);
      return [
        { label: "Insert table…", key: key("table"), action: () => insertTableDialog() },
        { sep: true },
        ...(inTable ? tableMenu(view(), actions) : [{ label: "(place the cursor in a table for more options)", disabled: true }]),
      ];
    } },
    { title: "Help", alt: "h", items: () => [
      { label: "Keyboard shortcuts", key: key("help"), action: () => shortcutsDialog(app.shortcuts) },
      { label: "Check for updates…", action: () => { checkForUpdates(true); } },
      { label: "About OfficeMini", action: () => aboutDialog() },
    ] },
  ];
  let openIdx = -1;
  const titles: HTMLElement[] = [];
  const openMenu = (i: number) => {
    closeAllPopups();
    openIdx = i;
    titles.forEach((t, j) => t.classList.toggle("open", j === i));
    showMenu(titles[i], menus[i].items(), { onClose: () => { if (openIdx === i) { openIdx = -1; titles[i].classList.remove("open"); } } });
  };
  menus.forEach((m, i) => {
    const t = el("div", { class: "menu-title" }, m.title);
    t.addEventListener("mousedown", (e) => { e.preventDefault(); if (openIdx === i) { closeAllPopups(); } else openMenu(i); });
    t.addEventListener("mouseenter", () => { if (openIdx >= 0 && openIdx !== i) openMenu(i); });
    titles.push(t);
    bar.appendChild(t);
  });
  // Alt+letter opens menus
  window.addEventListener("keydown", (e) => {
    if (e.altKey && !e.ctrlKey && !e.metaKey) {
      const i = menus.findIndex((m) => m.alt === e.key.toLowerCase());
      if (i >= 0) { e.preventDefault(); openMenu(i); }
    }
    if (openIdx >= 0 && (e.key === "ArrowLeft" || e.key === "ArrowRight")) {
      e.preventDefault();
      openMenu((openIdx + (e.key === "ArrowRight" ? 1 : menus.length - 1)) % menus.length);
    }
  });
}

function styleId(id: string, name: string): string {
  return ctx.styles.has(id) ? id : ctx.styleIdByName(name) || id;
}

function fitWidth() {
  const sect = view().state.doc.attrs.sect as SectProps;
  const pageW = twipsToPx(sect.pgW) + 48;
  const avail = $("workspace").clientWidth - 20;
  setZoom(avail / pageW);
}

// ---------------------------------------------------------------------------
// Global key handling outside the editor & browser-default suppression

function installGlobalKeys() {
  window.addEventListener("keydown", (e) => {
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key.toLowerCase();
    // Block webview defaults that would break the app.
    if ((mod && ["p", "s", "o", "f", "h", "g", "r", "u", "j", "d", "n", "w", "q", "+", "=", "-", "0"].includes(k)) || k === "f5" || (mod && e.shiftKey && k === "i") || k === "f3" || k === "f7" || (e.altKey && k === "arrowleft")) {
      e.preventDefault();
    }
    const inEditor = view().dom.contains(e.target as Node);
    if (inEditor) return; // ProseMirror keymap handles it
    // Toolbar inputs / dialogs: route app shortcuts.
    if (mod && !e.shiftKey && !e.altKey) {
      switch (k) {
        case "s": save(); break;
        case "o": openFile(); break;
        case "p": print(); break;
        case "n": newDocument(); break;
        case "f": showFind(false); break;
        case "h": showFind(true); break;
        case "w": case "q": closeWindowRequest(); break;
        case "=": case "+": setZoom(app.zoom + 0.1); break;
        case "-": setZoom(app.zoom - 0.1); break;
        case "0": setZoom(1); break;
        case "/": shortcutsDialog(app.shortcuts); break;
        default: return;
      }
      e.preventDefault();
    } else if (mod && e.shiftKey && k === "s") { e.preventDefault(); saveAs(); }
    else if (k === "f3") { e.preventDefault(); if (findVisible) findbar.step(e.shiftKey ? -1 : 1); else showFind(false); }
    else if (k === "f1") { e.preventDefault(); shortcutsDialog(app.shortcuts); }
    else if (k === "f11") { e.preventDefault(); F.toggleFullscreen(); }
    else if (k === "escape" && !findVisible) { closeAllPopups(); view().focus(); }
  });
  // Ctrl+wheel and touchpad pinch zoom. The listener is passive so plain scrolling never waits for
  // JavaScript (a non-passive wheel listener makes WebKit scroll only after the handler has run,
  // which shows as laggy touchpad scrolling). The webview's own zoom hotkeys are disabled in
  // tauri.conf.json, so there is no default action to prevent.
  $("workspace").addEventListener("wheel", (e) => {
    if (!e.ctrlKey) return;
    // Pinch gestures arrive as ctrl+wheel with small pixel deltas: zoom proportionally and
    // continuously. A real mouse wheel (line deltas or large pixel steps) zooms in 10% steps.
    const pinch = e.deltaMode === 0 && Math.abs(e.deltaY) < 40;
    if (pinch) setZoom(app.zoom * Math.exp(-e.deltaY * 0.01), true);
    else setZoom(app.zoom + (e.deltaY < 0 ? 0.1 : -0.1));
  }, { passive: true });
  // Middle-click / plain click on the gray area focuses the editor.
  $("workspace").addEventListener("mousedown", (e) => {
    if (e.target === $("workspace") || e.target === $("pagearea") || (e.target as HTMLElement).classList?.contains("page-bg")) {
      if (e.button !== 0) return;
      e.preventDefault();
      // Put the caret at the end of the document when clicking below content.
      const v = view();
      v.focus();
    }
  });
  // Keep the "paste plain" flag in sync with Ctrl+Shift+V pressed inside the editor.
  view().dom.addEventListener("keydown", (e) => { if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === "v") requestPlainPaste(); }, true);
  // Paste images from the clipboard (files / bitmaps)
  view().dom.addEventListener("paste", (e) => {
    const items = e.clipboardData?.items;
    if (!items) return;
    for (const it of Array.from(items)) {
      if (it.kind === "file" && it.type.startsWith("image/")) {
        const f = it.getAsFile();
        if (f) { e.preventDefault(); f.arrayBuffer().then((b) => insertImageBytes(new Uint8Array(b), it.type.split("/")[1], f.name || "Pasted image")); return; }
      }
    }
  }, true);
}

// ---------------------------------------------------------------------------
// Boot

async function boot() {
  app.settings = await F.loadSettings();
  app.zoom = app.settings.zoom || 1;
  app.mode = app.settings.view || "page";
  app.showMarks = !!app.settings.showMarks;
  app.theme = app.settings.theme || "light";
  document.documentElement.setAttribute("data-theme", app.theme);
  setDarkMode(app.theme === "dark");
  document.documentElement.style.setProperty("--zoom", String(app.zoom));
  setZoomFactor(app.zoom);
  $("workspace").classList.toggle("continuous", app.mode === "continuous");
  $("editor").classList.toggle("show-marks", app.showMarks);

  const blank = await loadBlank();
  app.loaded = blank;
  const handle = createEditor(blank.doc, {
    mount: $("editor"),
    pagesEl: $("pages"),
    actions,
    onUpdate: (v) => {
      if (v.state.doc !== lastDoc) { lastDoc = v.state.doc; if (!app.loading) { setDirty(true); hideWelcome(); } }
      app.toolbar?.update(v.state);
      updateStatus();
      if (findVisible) findbar.refresh();
    },
    onPages: (n) => { app.pages = n; updateStatus(); if (!app.shown) revealWindow(); },
    onContextMenu: (v, e) => showContextMenu(v, e, actions, app.shortcuts),
  });
  app.handle = handle;
  app.shortcuts = handle.shortcuts;
  let lastDoc = handle.view.state.doc;
  if (app.mode !== "page") setViewMode(handle.view, app.mode);

  app.toolbar = buildToolbar($("toolbar"), { view: () => app.handle?.view || null, run: (cmd, focus) => run(cmd, focus), actions, shortcuts: app.shortcuts });
  app.status = buildStatusbar($("statusbar"), { onZoom: setZoom, onViewMode: setMode, onGoToPage: () => goToPage(), onSource: () => toggleSourceView(), onTheme: () => toggleTheme() });
  findbar = buildFindbar();
  buildMenubar();
  installGlobalKeys();
  app.toolbar.update(handle.view.state);
  updateStatus();

  // Which file to open: ?file= (new windows) or CLI args (first window).
  const params = new URLSearchParams(location.search);
  let file = params.get("file");
  if (!file) { const args = await F.cliArgs(); if (args.length) { file = args[0]; for (const extra of args.slice(1)) F.openInNewWindow(extra); } }
  setSmartTyping(app.settings.smartQuotes !== false);
  if (file) await openPath(file);
  else {
    // No document requested: show the welcome screen with recent files and any recovered copies.
    const recovered = await findRecoveries();
    await buildWelcome(recovered);
    showWelcome();
  }
  lastDoc = handle.view.state.doc;
  setDirty(false);
  handle.view.focus();
  setTimeout(revealWindow, 400);
  setInterval(() => { autosaveTick(); }, 60000);
  // Quiet update check a few seconds after start (never blocks opening a document).
  setTimeout(() => { checkForUpdates(false); }, 8000);

  F.onCloseRequested(confirmDiscard);
  F.onFileDrop((paths) => {
    for (const p of paths) {
      const ext = F.extname(p);
      if (["png", "jpg", "jpeg", "gif", "bmp", "webp", "svg"].includes(ext)) F.readFile(p).then((b) => insertImageBytes(b, ext, F.basename(p)));
      else if (!app.dirty && (!app.path || app.kind === "new")) openPath(p);
      else F.openInNewWindow(p);
    }
  }, (over) => $("workspace").classList.toggle("drop-target", over));

  // Dev helpers (browser mode): expose the app for inspection.
  (window as any).om = { app, view: () => view(), openPath, ctx, writeDocx, loadDocx, schema, writeTo, save, setMode, setZoom, print };
}

function revealWindow() {
  if (app.shown) return;
  app.shown = true;
  F.showWindow();
}

boot().catch((e) => { console.error(e); document.body.innerHTML = `<pre style="padding:20px;color:#b00">Failed to start: ${(e as Error).stack || e}</pre>`; F.showWindow(); });

void icon; void tooltip;
