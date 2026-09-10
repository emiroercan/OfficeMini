// Sheets application shell: menus, toolbar, formula bar, cell editing,
// clipboard, rows/columns, sorting, filtering, find/replace, sheet tabs,
// files, recovery and printing. Mirrors the Words shell in main.ts.
import { Grid, GridEvents, Hit } from "./grid";
import {
  Workbook, Sheet, Range, Cell, key, rowOf, colOf, cellRef, rangeRef, parseRef, colName, MAXR, MAXC, newSheet, setColWidth, setRowHeight, pxToChars,
  charsToPx, colWidthChars, colHidden, rowHeightPx, mergeAt, inRange, normRange, quoteSheet, isError, Hyperlink, Value,
} from "./model";
import { loadXlsx } from "./xlsx-read";
import { writeXlsx } from "./xlsx-write";
import { csvToWorkbook, blankWorkbook, sheetToCsv, encodeCsv } from "./csv";
import { StyleResolver } from "./render-style";
import { Engine } from "./formula/engine";
import { History, Entry, CellChange, cellsEntry, styleEntry, StylePatch, inputCell, structuralEntry, insertDelete, renameSheet, fillChanges, sortChanges, clearChanges, boundRange, xfWith, SelSnapshot, movedCell } from "./edit";
import { copyRange, pasteInternal, pasteExternal, parseExternal, isOurHtml, internalClip, clearInternalClip, cancelCut, PasteMode } from "./clipboard";
import { setLocale, locale, detectLocale, LOCALES, presets, formatValue, editText, todaySerial, nowSerial, isDateFormat } from "./numfmt";
import { CellEditor } from "./celledit";
import { dataBlock, looksLikeHeader, applyFilters, filterState, clearFilterState, showFilterPopup, displayOf, columnFiltered } from "./filter";
import { printDialog, printSheet } from "./print";
import { pivotDialog, computePivot, pivotDefs, PivotDef } from "./pivot";
import { setDarkMode } from "../docx/props";
import { el, showMenu, MenuItem, tooltip, closeAllPopups, icon, showPopup, showNotice } from "../ui/widgets";
import { showDialog, promptDialog, closeDialog, dialogOpen } from "../ui/dialog-core";
import { checkForUpdates } from "../updater";
import * as F from "../files";

type Kind = "xlsx" | "csv" | "new";
const SHEET_EXT = ["xlsx", "xlsm", "xltx", "csv", "tsv"];
const isMac = /Mac|iPhone|iPad/.test(navigator.platform);
const MOD = isMac ? "⌘" : "Ctrl";

const app = {
  wb: null as Workbook | null,
  grid: null as Grid | null,
  styles: null as StyleResolver | null,
  engine: null as Engine | null,
  history: new History(),
  editor: null as CellEditor | null,
  editSheet: -1,
  path: null as string | null,
  kind: "new" as Kind,
  dirty: false,
  zoom: 1,
  theme: "light" as "light" | "dark",
  settings: {} as F.Settings,
  welcome: false,
  recoveryId: null as string | null,
  shown: false,
  tabStartCol: null as number | null,
  toolbar: null as ReturnType<typeof buildToolbar> | null,
  status: null as ReturnType<typeof buildStatusbar> | null,
  findbar: null as ReturnType<typeof buildFindbar> | null,
};

const $ = (id: string) => document.getElementById(id)!;
const wb = () => app.wb!;
const grid = () => app.grid!;
const sheet = () => grid().sheet();
const styles = () => app.styles!;

// DOM pieces created in buildWorkspace()
let namebox: HTMLInputElement, finput: HTMLTextAreaElement, gridHost: HTMLElement, celled: HTMLTextAreaElement, tabsEl: HTMLElement;
/**
 * Invisible textarea that holds keyboard focus while the grid is "focused". Copy / cut / paste
 * then arrive as native clipboard events in every webview (Chromium, WebKitGTK, WebKit) without
 * needing the async clipboard API or permissions, and typing starts a cell edit from keydown.
 */
let keyProxy: HTMLTextAreaElement;
function focusGrid() { keyProxy.focus({ preventScroll: true }); keyProxy.select(); }

// ---------------------------------------------------------------------------
// Title / dirty

function docName(): string { return app.path ? F.basename(app.path) : app.kind === "csv" ? "Untitled.csv" : "Untitled.xlsx"; }
function updateTitle() {
  const t = `${docName()}${app.dirty ? " •" : ""} - OfficeMini`;
  document.title = t;
  F.setWindowTitle(t);
}
function setDirty(d: boolean) {
  if (app.dirty === d) return;
  app.dirty = d;
  updateTitle();
  updateStatus();
}
function flash(msg: string) { app.status?.flash(msg); }

// ---------------------------------------------------------------------------
// Workspace DOM

function buildWorkspace() {
  const ws = $("workspace");
  ws.classList.add("sheets");
  $("pagearea").style.display = "none";
  $("source").style.display = "none";
  namebox = el("input", { class: "namebox", type: "text", spellcheck: "false", autocomplete: "off", "aria-label": "Name box" });
  tooltip(namebox, "Name box: type a cell or range (e.g. B3 or A1:C10) and press Enter");
  finput = el("textarea", { class: "finput", rows: "1", spellcheck: "false", "aria-label": "Formula bar" });
  const fbar = el("div", { id: "fbar" }, namebox, el("span", { class: "fx" }, "fx"), finput);
  gridHost = el("div", { tabindex: "-1" });
  celled = el("textarea", { id: "celled", spellcheck: "false", "aria-label": "Cell editor" });
  keyProxy = el("textarea", { id: "keyproxy", "aria-label": "Spreadsheet", spellcheck: "false", autocomplete: "off", tabindex: "0", style: { position: "absolute", left: "0", top: "0", width: "1px", height: "1px", opacity: "0", padding: "0", border: "0", resize: "none", overflow: "hidden", zIndex: "-1" } });
  keyProxy.value = " ";
  keyProxy.addEventListener("focus", () => keyProxy.select());
  keyProxy.addEventListener("input", () => { keyProxy.value = " "; keyProxy.select(); });
  const gridwrap = el("div", { id: "gridwrap" }, keyProxy, gridHost, celled);
  tabsEl = el("div", { id: "tabs" });
  const welcome = $("welcome");
  ws.insertBefore(fbar, welcome);
  ws.insertBefore(gridwrap, welcome);
  ws.insertBefore(tabsEl, welcome);
  namebox.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); goToRef(namebox.value.trim()); focusGrid(); }
    else if (e.key === "Escape") { e.preventDefault(); updateNameBox(); focusGrid(); }
    e.stopPropagation();
  });
  namebox.addEventListener("focus", () => namebox.select());
}

function goToRef(text: string) {
  if (!text) return;
  let ref = parseRef(text.toUpperCase().replace(/İ/g, "I") === text.toUpperCase() ? text : text);
  if (!ref) {
    const dn = wb().definedNames.find((d) => d.name.toLowerCase() === text.toLowerCase());
    if (dn) ref = parseRef(dn.ref.replace(/\$/g, ""));
  }
  if (!ref) { flash("Not a valid reference: " + text); return; }
  if (ref.sheet) {
    const i = wb().sheets.findIndex((s) => s.name.toLowerCase() === ref!.sheet!.toLowerCase());
    if (i < 0) { flash("No sheet named " + ref.sheet); return; }
    if (i !== grid().sheetIdx()) switchSheet(i);
  }
  grid().setRanges([{ r1: ref.r1, c1: ref.c1, r2: ref.r2, c2: ref.c2 }], { r: ref.r1, c: ref.c1 });
  grid().ensureVisible(ref.r1, ref.c1);
}

// ---------------------------------------------------------------------------
// Workbook lifecycle

function installWorkbook(wbk: Workbook, path: string | null, kind: Kind) {
  if (app.editor?.active) app.editor.finish(null, false);
  app.wb = wbk; app.path = path; app.kind = kind;
  app.styles = new StyleResolver(wbk);
  app.engine = new Engine(wbk);
  app.history.clear();
  app.recoveryId = null;
  app.tabStartCol = null;
  app.editSheet = -1;
  clearInternalClip();
  if (!app.grid) {
    app.grid = new Grid(gridHost, wbk, app.styles, gridEvents);
    app.grid.zoom = app.zoom;
    app.grid.hidePageBreaks = true;
    installGridExtras();
  } else {
    app.grid.wb = wbk; app.grid.styles = app.styles; app.grid.clipRange = null;
    app.grid.setSheet(wbk.active);
  }
  app.grid.invalidate();
  app.dirty = false;
  updateTitle();
  renderTabs();
  updateFbar();
  updateStatus();
  updateToolbarState();
  hideWelcome();
}

function untouched(): boolean { return !app.dirty && app.kind === "new" && !app.path; }

async function openPath(path: string) {
  const ext = F.extname(path);
  if (!SHEET_EXT.includes(ext) && ext !== "txt") {
    // Word / Markdown files belong to the Words editor.
    if (untouched()) { location.search = "?file=" + encodeURIComponent(path); } else F.openInNewWindow(path);
    return;
  }
  const t0 = performance.now();
  try {
    const bytes = await F.readFile(path);
    const t1 = performance.now();
    let w: Workbook;
    let kind: Kind;
    if (ext === "csv" || ext === "tsv" || ext === "txt") { w = csvToWorkbook(bytes, F.basename(path)); kind = "csv"; }
    else { w = loadXlsx(bytes, path); kind = "xlsx"; }
    w.path = path;
    const t2 = performance.now();
    installWorkbook(w, path, kind);
    const t3 = performance.now();
    addRecent(path);
    flash(`Opened in ${Math.round(t3 - t0)} ms (read ${Math.round(t1 - t0)}, parse ${Math.round(t2 - t1)}, show ${Math.round(t3 - t2)})`);
    if (F.isTauri) offerRecoveryFor(path);
  } catch (e) {
    console.error(e);
    await F.showMessage("Could not open the file.\n\n" + (e as Error).message, "OfficeMini", "error");
  }
}

function newWorkbook(kind: "xlsx" | "csv" = "xlsx") {
  const w = blankWorkbook();
  if (kind === "csv") { w.kind = "csv"; w.csvOptions = { delimiter: locale().decimal === "," ? ";" : ",", encoding: "utf-8", bom: true }; }
  installWorkbook(w, null, "new");
  if (kind === "csv") app.kind = "csv";
  updateTitle();
  focusGrid();
}

async function newSpreadsheet() {
  if (untouched()) { newWorkbook(); return; }
  F.openInNewWindow("new:xlsx");
}

const SHEET_FILTERS: F.FileFilter[] = [
  { name: "Spreadsheets", extensions: ["xlsx", "xlsm", "csv", "tsv", "txt"] },
  { name: "Excel Workbook", extensions: ["xlsx", "xlsm"] },
  { name: "CSV", extensions: ["csv", "tsv", "txt"] },
  { name: "Documents", extensions: ["docx", "md", "markdown"] },
  { name: "All files", extensions: ["*"] },
];

async function openFile() {
  const paths = await F.openDialog(SHEET_FILTERS, true);
  if (!paths || !paths.length) return;
  const [first, ...rest] = paths;
  if (untouched()) await openPath(first); else F.openInNewWindow(first);
  for (const p of rest) F.openInNewWindow(p);
}

async function save(): Promise<boolean> {
  if (!app.path || app.kind === "new") return saveAs();
  return writeTo(app.path);
}

async function saveAs(): Promise<boolean> {
  const isCsv = app.kind === "csv" || (app.kind === "new" && wb().kind === "csv");
  const defaultName = app.path ? F.basename(app.path) : isCsv ? "Untitled.csv" : "Untitled.xlsx";
  const filters = isCsv
    ? [{ name: "CSV", extensions: ["csv"] }, { name: "Excel Workbook", extensions: ["xlsx"] }]
    : [{ name: "Excel Workbook", extensions: ["xlsx"] }, { name: "CSV (current sheet)", extensions: ["csv"] }];
  const target = await F.saveDialog(app.path ? app.path : defaultName, filters);
  if (!target) return false;
  const ok = await writeTo(target);
  if (ok) { app.path = target; app.kind = F.extname(target) === "csv" ? "csv" : "xlsx"; wb().path = target; updateTitle(); addRecent(target); }
  return ok;
}

function csvText(): string {
  const s = sheet();
  const opts = wb().csvOptions;
  const delim = opts?.delimiter || (locale().decimal === "," ? ";" : ",");
  return sheetToCsv(s, delim, (r, c) => { const cell = s.cells.get(key(r, c)); if (!cell) return ""; if (typeof cell.v === "string") return cell.v; return displayOf(cell, styles()); });
}

async function writeTo(path: string): Promise<boolean> {
  if (app.editor?.active) app.editor.finish(null, true);
  const t0 = performance.now();
  try {
    const ext = F.extname(path);
    if (ext === "csv" || ext === "tsv" || ext === "txt") {
      const text = csvText();
      const opts = wb().csvOptions;
      await F.writeFile(path, encodeCsv(text, opts?.encoding || "utf-8", opts ? opts.bom : true));
    } else {
      wb().pkg.materialize();
      const bytes = writeXlsx(wb());
      await F.writeFile(path, bytes);
    }
    setDirty(false);
    clearRecovery();
    flash(`Saved in ${Math.round(performance.now() - t0)} ms`);
    return true;
  } catch (e) {
    console.error(e);
    await F.showMessage("Could not save the file.\n\n" + (e as Error).message, "OfficeMini", "error");
    return false;
  }
}

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

/**
 * Closing asks nothing: a recovery copy is written and the window goes (see the same
 * comment in main.ts). Cancel or Esc keeps the window; after CLOSE_COPY_TIMEOUT it
 * closes whether the copy finished or not.
 */
const CLOSE_COPY_TIMEOUT = 5000;
let closePending = false;

async function requestClose(): Promise<boolean> {
  if (!app.dirty) return true;
  if (!F.isTauri || app.settings.autosave === false) return confirmDiscard();
  if (closePending) return false;
  closePending = true;
  return new Promise<boolean>((resolve) => {
    let done = false;
    const finish = (close: boolean) => {
      if (done) return;
      done = true;
      closePending = false;
      clearTimeout(timer);
      document.removeEventListener("keydown", onKey, true);
      notice.close();
      resolve(close);
    };
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); finish(false); } };
    const notice = showNotice(`Closing "${docName()}" - saving a copy you can recover.`, { label: "Cancel", onClick: () => finish(false) });
    const timer = setTimeout(() => finish(true), CLOSE_COPY_TIMEOUT);
    document.addEventListener("keydown", onKey, true);
    writeRecoveryCopy().then(() => finish(true), (e) => { console.warn("close copy failed", e); finish(true); });
  });
}

async function closeWindowRequest() { if (await requestClose()) F.closeWindow(); }

function addRecent(path: string) {
  const list = (app.settings.recent || []).filter((p) => p !== path);
  list.unshift(path);
  app.settings.recent = list.slice(0, 12);
  F.saveSettings(app.settings);
}

// ---------------------------------------------------------------------------
// Recovery (autosave copies) and welcome screen

interface RecoveryEntry { id: string; path: string | null; name: string; file: string; savedAt: number; kind: string; }

function recoveryIdFor(path: string | null): string {
  if (!path) return "untitled-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  let h = 0;
  for (let i = 0; i < path.length; i++) h = (h * 31 + path.charCodeAt(i)) >>> 0;
  return "doc-" + h.toString(16) + "-" + F.basename(path).replace(/[^\w.-]+/g, "_").slice(0, 40);
}

/** Write a recovery copy of the current workbook; rejects when it cannot be written. */
async function writeRecoveryCopy(): Promise<void> {
  const dir = await F.recoveryDir();
  if (!dir) throw new Error("no recovery directory");
  if (!app.recoveryId) app.recoveryId = recoveryIdFor(app.path);
  const kind = app.kind === "csv" ? "csv" : "xlsx";
  const bytes = kind === "csv" ? new TextEncoder().encode(csvText()) : (wb().pkg.materialize(), writeXlsx(wb()));
  await F.writeFile(F.joinPath(dir, app.recoveryId + "." + kind), bytes);
  await F.writeFile(F.joinPath(dir, app.recoveryId + ".json"), new TextEncoder().encode(JSON.stringify({ path: app.path, name: docName(), savedAt: Date.now(), kind })));
}

async function autosaveTick() {
  if (!F.isTauri || !app.dirty || app.settings.autosave === false || !app.wb) return;
  try { await writeRecoveryCopy(); flash("Recovery copy saved"); }
  catch (e) { console.warn("autosave failed", e); }
}

async function clearRecovery() {
  if (!F.isTauri || !app.recoveryId) return;
  const dir = await F.recoveryDir();
  if (!dir) return;
  for (const ext of ["xlsx", "csv", "json"]) await F.deleteFile(F.joinPath(dir, app.recoveryId + "." + ext));
  app.recoveryId = null;
}

async function discardRecovery(r: RecoveryEntry) {
  const dir = await F.recoveryDir();
  if (!dir) return;
  for (const ext of [r.kind, "json"]) await F.deleteFile(F.joinPath(dir, r.id + "." + ext));
}

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
      if (meta.kind !== "xlsx" && meta.kind !== "csv") continue;
      const id = f.name.slice(0, -5);
      const data = files.find((x) => x.name === id + "." + meta.kind);
      if (!data) { await F.deleteFile(f.path); continue; }
      if (meta.path) {
        const mt = await F.fileMtime(meta.path);
        if (mt !== null && mt * 1000 >= meta.savedAt - 1500) { await discardRecovery({ id, path: meta.path, name: "", file: data.path, savedAt: 0, kind: meta.kind }); continue; }
      }
      out.push({ id, path: meta.path || null, name: meta.name || data.name, file: data.path, savedAt: meta.savedAt || f.mtime * 1000, kind: meta.kind });
    } catch { /* ignore */ }
  }
  return out.sort((a, b) => b.savedAt - a.savedAt);
}

async function openRecovered(r: RecoveryEntry) {
  await openPath(r.file);
  app.path = r.path;
  app.kind = r.path ? (F.extname(r.path) === "csv" ? "csv" : "xlsx") : "new";
  if (app.wb) app.wb.path = r.path;
  app.recoveryId = r.id;
  if (r.path) { app.settings.recent = (app.settings.recent || []).filter((p) => p !== r.file); F.saveSettings(app.settings); }
  setDirty(true);
  updateTitle();
  flash("Recovered copy opened - save to keep it");
}

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

function iconFor(path: string): { cls: string; letter: string } {
  const ext = F.extname(path);
  if (ext === "docx") return { cls: "docx", letter: "W" };
  if (ext === "md" || ext === "markdown") return { cls: "md", letter: "M" };
  if (ext === "csv" || ext === "tsv" || ext === "txt") return { cls: "csv", letter: "C" };
  return { cls: "xlsx", letter: "S" };
}

async function buildWelcome(recovered: RecoveryEntry[]) {
  const w = $("welcome");
  w.innerHTML = "";
  const card = el("div", { class: "welcome-card" });
  card.append(el("h1", null, "OfficeMini Sheets"), el("div", { class: "sub" }, "Open a spreadsheet or start a new one. Excel and CSV files, fast."));
  const actions = el("div", { class: "welcome-actions" });
  const newBtn = el("button", { class: "primary" }, icon("new"), "New spreadsheet");
  newBtn.addEventListener("click", () => { hideWelcome(); focusGrid(); });
  const newDoc = el("button", null, icon("new"), "New document");
  newDoc.addEventListener("click", () => { location.search = "?file=new%3Adocx"; });
  const openBtn = el("button", null, icon("open"), "Open…");
  openBtn.addEventListener("click", () => openFile());
  actions.append(newBtn, newDoc, openBtn);
  card.append(actions);
  if (recovered.length) {
    const box = el("div", { class: "welcome-recovery" });
    box.append(el("h2", null, "Recovered spreadsheets"));
    for (const r of recovered) {
      const item = el("div", { class: "recent-item", title: r.path || "Unsaved spreadsheet" });
      const ic = iconFor(r.path || r.name);
      const x = el("span", { class: "ri-x", title: "Discard recovered copy" }, "✕");
      x.addEventListener("click", async (e) => { e.stopPropagation(); await discardRecovery(r); item.remove(); if (!box.querySelector(".recent-item")) box.remove(); });
      item.append(el("span", { class: "ri-icon " + ic.cls }, ic.letter), el("span", { class: "ri-name" }, r.name), el("span", { class: "ri-path", style: { direction: "ltr" } }, "autosaved " + new Date(r.savedAt).toLocaleString()), x);
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
      const ic = iconFor(p);
      const x = el("span", { class: "ri-x", title: "Remove from list" }, "✕");
      x.addEventListener("click", (e) => { e.stopPropagation(); app.settings.recent = (app.settings.recent || []).filter((q) => q !== p); F.saveSettings(app.settings); item.remove(); });
      item.append(el("span", { class: "ri-icon " + ic.cls }, ic.letter), el("span", { class: "ri-name" }, F.basename(p)), el("span", { class: "ri-path" }, F.dirname(p)), x);
      item.addEventListener("click", () => { if (exists) openPath(p); else flash("File not found: " + p); });
      list.append(item);
    }
    card.append(list);
  }
  card.append(el("div", { class: "welcome-hint" }, `${MOD}+O open · ${MOD}+N new · ${MOD}+/ all shortcuts · drop a file on the window to open it`));
  w.append(card);
}
function showWelcome() { app.welcome = true; $("welcome").hidden = false; }
function hideWelcome() { if (!app.welcome) return; app.welcome = false; $("welcome").hidden = true; }

// ---------------------------------------------------------------------------
// Theme / zoom / locale

function setTheme(theme: "light" | "dark", persist = true) {
  app.theme = theme;
  document.documentElement.setAttribute("data-theme", theme);
  setDarkMode(theme === "dark");
  if (app.grid) { styles().invalidate(); app.grid.readColors(); app.grid.invalidate(); }
  if (persist) { app.settings.theme = theme; F.saveSettings(app.settings); }
  updateStatus();
}
function toggleTheme() { setTheme(app.theme === "dark" ? "light" : "dark"); }

let zoomSaveTimer: ReturnType<typeof setTimeout> | null = null;
function setZoom(z: number, fine = false) {
  z = Math.max(0.5, Math.min(3, fine ? Math.round(z * 1000) / 1000 : Math.round(z * 100) / 100));
  if (z === app.zoom) return;
  app.zoom = z;
  grid().zoom = z;
  grid().invalidate();
  app.editor?.placeOverlay();
  updateStatus();
  if (zoomSaveTimer) clearTimeout(zoomSaveTimer);
  zoomSaveTimer = setTimeout(() => { app.settings.sheetZoom = app.zoom; F.saveSettings(app.settings); }, 400);
}

function setLocaleId(id: string, persist = true) {
  setLocale(id);
  if (persist) { app.settings.locale = id; F.saveSettings(app.settings); }
  if (app.grid) { styles().invalidate(); grid().invalidate(); updateFbar(); }
  updateStatus();
}

function scrollSpeed(): number {
  const s = app.settings.scrollSpeed;
  if (typeof s === "number" && s > 0) return s;
  return /Linux/.test(navigator.platform) ? 2 : 1;
}

// ---------------------------------------------------------------------------
// Grid events, selection helpers

function selSnapshot(): SelSnapshot { const s = grid().selection; return { sheet: grid().sheetIdx(), ranges: s.ranges.map((r) => ({ ...r })), active: { ...s.active } }; }
function restoreSel(s: SelSnapshot) {
  if (s.sheet !== grid().sheetIdx() && s.sheet < wb().sheets.length) switchSheet(s.sheet);
  grid().setRanges(s.ranges.map((r) => ({ ...r })), { ...s.active });
  grid().ensureVisible(s.active.r, s.active.c);
}

/** The current selection, whole-row/column selections bounded to the used area. */
function selRange(): Range { const s = grid().selection; return boundRange(sheet(), s.ranges[s.ranges.length - 1]); }
function selRanges(): Range[] { return grid().selection.ranges.map((r) => boundRange(sheet(), r)); }
function isWholeCols(rg: Range) { return rg.r1 === 0 && rg.r2 >= MAXR - 1; }
function isWholeRows(rg: Range) { return rg.c1 === 0 && rg.c2 >= MAXC - 1; }
function rawSel(): Range { const s = grid().selection; return s.ranges[s.ranges.length - 1]; }

const gridEvents: GridEvents = {
  onSelect() { if (!app.grid) return; if (!app.editor?.active) updateFbar(); updateNameBox(); updateStatus(); updateToolbarState(); if (app.findbar?.visible) app.findbar.refresh(); },
  onEdit(seed) { beginEdit(seed, seed !== null); },
  onContextMenu(e, hit) { showContextMenu(e, hit); },
  onColResize(c, w, all) { structural("Resize column", () => { for (const cc of all) setColWidth(sheet(), cc, pxToChars(w)); }); void c; },
  onRowResize(r, h, all) { structural("Resize row", () => { for (const rr of all) setRowHeight(sheet(), rr, Math.round((h * 72) / 96 * 100) / 100); }); void r; },
  onAutoFit(kind, index) { autoFit(kind, [index]); },
  onFill(src, dst) { const ch = fillChanges(wb(), sheet(), src, dst); applyChanges(ch, "Fill"); grid().setRanges([{ r1: Math.min(src.r1, dst.r1), c1: Math.min(src.c1, dst.c1), r2: Math.max(src.r2, dst.r2), c2: Math.max(src.c2, dst.c2) }], { ...grid().selection.active }); },
  onLink(hl) { openLink(hl); },
  onFilterButton(c, r, x, y) { filterPopup(c, r, x, y); },
  onMoveRange(src, dst, copy) { moveRange(src, dst, copy); },
};

function installGridExtras() {
  const g = grid();
  // Overlay follows scrolling; Ctrl+wheel zooms (passive so scrolling never waits).
  g.host.addEventListener("scroll", () => { if (app.editor?.active) app.editor.placeOverlay(); }, { passive: true });
  g.host.addEventListener("wheel", (e) => {
    if (!e.ctrlKey) {
      const f = scrollSpeed();
      if (f !== 1 && e.deltaMode === 0 && !e.shiftKey) { g.host.scrollTop += e.deltaY * (f - 1); g.host.scrollLeft += e.deltaX * (f - 1); }
      return;
    }
    const pinch = e.deltaMode === 0 && Math.abs(e.deltaY) < 40;
    if (pinch) setZoom(app.zoom * Math.exp(-e.deltaY * 0.01), true); else setZoom(app.zoom + (e.deltaY < 0 ? 0.1 : -0.1));
  }, { passive: true });
  // Formula reference pointing: clicks on the grid insert references while a formula is being typed.
  g.host.addEventListener("mousedown", (e) => {
    const ed = app.editor;
    if (!ed || !ed.active || e.button !== 0 || !ed.canPointNow()) return;
    const hit = g.hitTest(e);
    if (hit.type !== "cell" && hit.type !== "colHeader" && hit.type !== "rowHeader") return;
    e.preventDefault(); e.stopPropagation();
    ed.suppressBlur = true;
    const other = grid().sheetIdx() !== app.editSheet ? sheet().name : null;
    const rangeFor = (h: Hit, start: Hit): Range => {
      if (h.type === "colHeader" || start.type === "colHeader") { const c1 = (start as any).c, c2 = (h as any).c ?? c1; return { r1: 0, c1: Math.min(c1, c2), r2: MAXR - 1, c2: Math.max(c1, c2) }; }
      if (h.type === "rowHeader" || start.type === "rowHeader") { const r1 = (start as any).r, r2 = (h as any).r ?? r1; return { r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAXC - 1 }; }
      return normRange({ r: (start as any).r, c: (start as any).c }, { r: (h as any).r, c: (h as any).c });
    };
    const start = hit;
    ed.pointWithMouse(rangeFor(hit, start), other);
    const mm = (ev: MouseEvent) => { const h = g.hitTest(ev); if (h.type === "cell" || h.type === "colHeader" || h.type === "rowHeader") ed.pointWithMouse(rangeFor(h, start), other); };
    const mu = () => { window.removeEventListener("mousemove", mm); window.removeEventListener("mouseup", mu); ed.suppressBlur = false; ed.endPointing(); };
    window.addEventListener("mousemove", mm);
    window.addEventListener("mouseup", mu);
  }, true);
  g.host.addEventListener("dragover", (e) => e.preventDefault());
  // The grid host is not the keyboard target: clicks and Tab-focus land on the proxy textarea.
  g.host.addEventListener("mousedown", () => { if (!app.editor?.active && !dialogOpen()) focusGrid(); });
  g.host.addEventListener("focusin", (e) => { if (e.target === g.host) focusGrid(); });
  g.filteredCols = (c) => columnFiltered(sheet(), c);
}

function updateNameBox() {
  if (document.activeElement === namebox) return;
  const rg = rawSel();
  const s = grid().selection;
  let text: string;
  if (isWholeCols(rg) && isWholeRows(rg)) text = "A1:" + colName(MAXC - 1) + MAXR;
  else if (isWholeCols(rg)) text = colName(rg.c1) + ":" + colName(rg.c2);
  else if (isWholeRows(rg)) text = (rg.r1 + 1) + ":" + (rg.r2 + 1);
  else if (rg.r1 !== rg.r2 || rg.c1 !== rg.c2) text = `${rg.r2 - rg.r1 + 1}R × ${rg.c2 - rg.c1 + 1}C  ${rangeRef(rg)}`;
  else text = cellRef(s.active.r, s.active.c);
  namebox.value = text;
}

function cellEditText(cell: Cell | undefined): string {
  if (!cell) return "";
  if (cell.f !== undefined) return "=" + cell.f;
  return editText(cell.v, styles().get(cell.s).numFmt, wb().date1904);
}

function updateFbar() {
  if (!app.grid || app.editor?.active) return;
  const { r, c } = grid().selection.active;
  finput.value = cellEditText(sheet().cells.get(key(r, c)));
  updateNameBox();
}

// ---------------------------------------------------------------------------
// Editing

function beginEdit(seed: string | null, arrowsNavigate: boolean) {
  const ed = app.editor!;
  if (ed.active) return;
  if (app.welcome) hideWelcome();
  const { r, c } = grid().selection.active;
  app.editSheet = grid().sheetIdx();
  ed.arrowsNavigate = arrowsNavigate;
  ed.begin(seed, false, cellEditText(sheet().cells.get(key(r, c))));
}

const editorHost = {
  get grid() { return grid(); },
  onBegin() { app.editSheet = grid().sheetIdx(); },
  pointSheetName() { return app.editSheet >= 0 && app.editSheet !== grid().sheetIdx() ? sheet().name : null; },
  commit(text: string, move: { dr: number; dc: number } | null) {
    const ed = app.editor!;
    const pos = ed.pos!;
    if (app.editSheet >= 0 && app.editSheet !== grid().sheetIdx()) switchSheet(app.editSheet);
    const s = sheet();
    const cell = inputCell(wb(), s, pos.r, pos.c, text);
    const old = s.cells.get(key(pos.r, pos.c)) || null;
    const same = (!old && !cell) || (!!old && !!cell && old.s === cell.s && old.f === cell.f && (old.f !== undefined || old.v === cell.v));
    if (!same) applyChanges([{ r: pos.r, c: pos.c, cell }], "Edit " + cellRef(pos.r, pos.c));
    focusGrid();
    if (move) {
      if (move.dr === 1 && move.dc === 0 && app.tabStartCol !== null) { grid().setActive(pos.r + 1, app.tabStartCol); app.tabStartCol = null; }
      else { if (move.dc) { if (app.tabStartCol === null) app.tabStartCol = pos.c; } else app.tabStartCol = null; grid().setActive(pos.r + move.dr, pos.c + move.dc); }
    } // no move (blur / click elsewhere): the grid already holds the selection the user chose
    updateFbar();
    app.status?.update();
  },
  cancel() { if (app.editSheet >= 0 && app.editSheet !== grid().sheetIdx()) switchSheet(app.editSheet); focusGrid(); updateFbar(); app.status?.update(); },
  onInput(text: string) {
    // live formula preview in the status bar
    if (text.startsWith("=") && text.length > 1 && app.editor?.pos && app.editSheet >= 0) {
      try {
        const v = app.engine!.evaluateFormula(text.slice(1), wb().sheets[app.editSheet], app.editor.pos.r, app.editor.pos.c);
        app.status?.preview(v === null ? "" : isError(v) ? v.e : typeof v === "number" ? formatValue(v, "General").text : String(v));
      } catch { app.status?.preview(""); }
    } else app.status?.preview("");
  },
};

function pushEntry(e: Entry, before: SelSnapshot) { e.before = before; e.after = selSnapshot(); app.history.push(e); }

function applyChanges(changes: CellChange[], label: string) {
  if (!changes.length) return;
  const before = selSnapshot();
  const e = cellsEntry(sheet(), changes, label);
  pushEntry(e, before);
  afterModel();
}

function afterModel() {
  app.engine!.recalcAll();
  if (sheet().autoFilter) applyFilters(sheet(), styles());
  grid().invalidate();
  setDirty(true);
  updateFbar();
  updateStatus();
  updateToolbarState();
  if (app.findbar?.visible) app.findbar.refresh();
}

function structural(label: string, mutate: () => void, othersToo = false) {
  const before = selSnapshot();
  const e = structuralEntry(wb(), grid().sheetIdx(), label, mutate, othersToo);
  pushEntry(e, before);
  styles().invalidate();
  afterModel();
}

function undo() {
  if (app.editor?.active) { app.editor.finish(null, false); }
  const e = app.history.undo();
  if (!e) { flash("Nothing to undo"); return; }
  styles().invalidate();
  if (e.before) restoreSel(e.before);
  afterModel();
  flash("Undo: " + e.label);
}
function redo() {
  if (app.editor?.active) { app.editor.finish(null, false); }
  const e = app.history.redo();
  if (!e) { flash("Nothing to redo"); return; }
  styles().invalidate();
  if (e.after) restoreSel(e.after);
  afterModel();
  flash("Redo: " + e.label);
}

function activeCell(): Cell | undefined { const { r, c } = grid().selection.active; return sheet().cells.get(key(r, c)); }
function activeStyle() { const cell = activeCell(); return styles().get(cell ? cell.s : (sheet().rows.get(grid().selection.active.r)?.style ?? 0)); }

function applyStyle(patch: StylePatch, label = "Format") {
  if (app.editor?.active && !app.editor.isFormula()) { /* formatting while editing applies to the cell being edited */ }
  const before = selSnapshot();
  const e = styleEntry(wb(), sheet(), grid().selection.ranges, patch, label);
  pushEntry(e, before);
  styles().invalidate();
  afterModel();
}

function toggleFont(prop: "bold" | "italic" | "strike" | "underline") {
  const cs = activeStyle();
  if (prop === "underline") applyStyle({ font: { underline: cs.underline ? null : "single" } }, "Underline");
  else applyStyle({ font: { [prop]: !cs[prop] } }, prop[0].toUpperCase() + prop.slice(1));
}

function currencyCode(dec = 2): string {
  const l = locale();
  const num = dec ? "#,##0.00" : "#,##0";
  return l.currencyPos === "prefix" ? `"${l.currency}"${num}` : `${num} "${l.currency}"`;
}
function dateCode(): string { const l = locale(); return l.dateOrder === "mdy" ? "m/d/yyyy" : l.dateOrder === "ymd" ? "yyyy-mm-dd" : `dd${l.dateSep}mm${l.dateSep}yyyy`; }

function changeDecimals(delta: number) {
  const code = activeStyle().numFmt;
  let out: string;
  const secs = code.split(";");
  const first = secs[0];
  if (code === "General") {
    const v = activeCell()?.v;
    const cur = typeof v === "number" ? Math.min(10, (String(v).split(".")[1] || "").length) : 0;
    const n = Math.max(0, cur + delta);
    out = n ? "0." + "0".repeat(n) : "0";
  } else {
    const adjust = (sec: string) => {
      const m = /(#|0)(\.(0+|#+))?(?=[^0#.]*(%|E\+|E-|$|"|\)|_|\\| ))/.exec(sec);
      if (!m) return sec;
      const digits = m[3] ? m[3].length : 0;
      const n = Math.max(0, digits + delta);
      const rep = m[1] + (n ? "." + "0".repeat(n) : "");
      return sec.slice(0, m.index) + rep + sec.slice(m.index + m[0].length);
    };
    out = secs.map(adjust).join(";");
    if (out === code && delta > 0 && !/[0#]\.[0#]/.test(first)) out = first.replace(/([0#])(?![0#.])/, "$1.0");
  }
  applyStyle({ numFmt: out }, "Decimals");
}

function clearFormat() { applyStyle({ reset: true }, "Clear formatting"); }

/** Border presets need per-cell patches (outer edges differ per cell). */
function applyBorders(kind: "all" | "outer" | "inner" | "top" | "bottom" | "left" | "right" | "none" | "thickOuter", color = "000000") {
  const st = wb().styles;
  const changes: CellChange[] = [];
  const side = (style: string) => ({ style, color: { rgb: "FF" + color } });
  for (const rg of selRanges()) {
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
      const cell = sheet().cells.get(key(r, c));
      const p: StylePatch = { border: {} };
      const b = p.border!;
      const top = r === rg.r1, bottom = r === rg.r2, left = c === rg.c1, right = c === rg.c2;
      const thin = side("thin"), medium = side("medium");
      switch (kind) {
        case "all": b.top = thin; b.bottom = thin; b.left = thin; b.right = thin; break;
        case "outer": if (top) b.top = thin; if (bottom) b.bottom = thin; if (left) b.left = thin; if (right) b.right = thin; break;
        case "thickOuter": if (top) b.top = medium; if (bottom) b.bottom = medium; if (left) b.left = medium; if (right) b.right = medium; break;
        case "inner": if (!top) b.top = thin; if (!bottom) b.bottom = thin; if (!left) b.left = thin; if (!right) b.right = thin; break;
        case "top": if (top) b.top = thin; break;
        case "bottom": if (bottom) b.bottom = thin; break;
        case "left": if (left) b.left = thin; break;
        case "right": if (right) b.right = thin; break;
        case "none": b.top = null; b.bottom = null; b.left = null; b.right = null; break;
      }
      if (!Object.keys(b).length) continue;
      const s = xfWith(st, cell ? cell.s : 0, p);
      if (cell) { if (cell.s !== s) changes.push({ r, c, cell: { ...cell, s } }); }
      else if (kind !== "none") changes.push({ r, c, cell: { v: null, s } });
    }
  }
  styles().invalidate();
  applyChanges(changes, "Borders");
}

function insertDateTime(time: boolean) {
  const { r, c } = grid().selection.active;
  const v = time ? nowSerial() - todaySerial() : todaySerial();
  const cell = sheet().cells.get(key(r, c));
  const cs = styles().get(cell ? cell.s : 0);
  const fmt = time ? "hh:mm" : dateCode();
  const s = isDateFormat(cs.numFmt) ? (cell ? cell.s : 0) : xfWith(wb().styles, cell ? cell.s : 0, { numFmt: fmt });
  styles().invalidate();
  applyChanges([{ r, c, cell: { v, s } }], time ? "Insert time" : "Insert date");
}

// ---------------------------------------------------------------------------
// Clipboard

let pendingCopy: { text: string; html: string } | null = null;
let pendingPasteMode: PasteMode = "all";
/** Native clipboard events: the grid proxy (or a menu-triggered fallback) supplies the data synchronously. */
function onCopyCut(e: ClipboardEvent, cut: boolean) {
  if (!e.clipboardData) return;
  const t = e.target as HTMLElement;
  if (pendingCopy) { e.preventDefault(); e.clipboardData.setData("text/plain", pendingCopy.text); e.clipboardData.setData("text/html", pendingCopy.html); pendingCopy = null; return; }
  if (t !== keyProxy || !app.grid || app.editor?.active) return; // find bar, name box, cell editor: native behaviour
  e.preventDefault();
  const clip = prepareCopy(cut);
  if (clip) { e.clipboardData.setData("text/plain", clip.text); e.clipboardData.setData("text/html", clip.html); }
}
document.addEventListener("copy", (e) => onCopyCut(e, false));
document.addEventListener("cut", (e) => onCopyCut(e, true));
document.addEventListener("paste", (e) => {
  const t = e.target as HTMLElement;
  if (t !== keyProxy && (t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement)) return;
  if (!app.grid || app.welcome || app.editor?.active) return;
  e.preventDefault();
  const cd = e.clipboardData;
  if (!cd) return;
  const mode = pendingPasteMode; pendingPasteMode = "all";
  doPaste(cd.getData("text/html") || null, cd.getData("text/plain") || "", mode);
});

/** Snapshot the selection into the internal clipboard; returns the system-clipboard payload. */
function prepareCopy(cut: boolean): { text: string; html: string } | null {
  const rg = rawSel();
  const bounded = boundRange(sheet(), rg);
  if ((bounded.r2 - bounded.r1 + 1) * (bounded.c2 - bounded.c1 + 1) > 2_000_000) { flash("Selection too large to copy"); return null; }
  const res = copyRange(wb(), sheet(), rg, cut, styles());
  grid().clipRange = { ...bounded };
  grid().schedule();
  flash(cut ? "Cut - paste to move" : "Copied");
  return res;
}

async function writeClipboard(text: string, html: string) {
  try {
    const item = new ClipboardItem({ "text/plain": new Blob([text], { type: "text/plain" }), "text/html": new Blob([html], { type: "text/html" }) });
    await navigator.clipboard.write([item]);
  } catch {
    pendingCopy = { text, html };
    const ta = el("textarea", { style: { position: "fixed", left: "-1000px", top: "0" } });
    ta.value = text || " ";
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand("copy"); } catch { /* ignore */ }
    ta.remove();
    focusGrid();
  }
}

/** Menu / context-menu copy: same as Ctrl+C but we have to write the clipboard ourselves. */
async function copySelection(cut: boolean) {
  if (app.editor?.active) return; // textarea handles its own copy
  const clip = prepareCopy(cut);
  if (!clip) return;
  await writeClipboard(clip.text, clip.html);
}

async function pasteFromSystem(mode: PasteMode = "all") {
  if (app.editor?.active) return;
  let html: string | null = null, text = "";
  try {
    const items: ClipboardItem[] = await (navigator.clipboard as any).read();
    for (const it of items) {
      if (it.types.includes("text/html")) html = await (await it.getType("text/html")).text();
      if (it.types.includes("text/plain")) text = await (await it.getType("text/plain")).text();
    }
  } catch {
    try { text = await navigator.clipboard.readText(); }
    catch {
      // No clipboard API: try the native paste command on the proxy; if nothing arrives, use what we copied ourselves.
      pendingPasteMode = mode; focusGrid();
      let ok = false;
      try { ok = document.execCommand("paste"); } catch { /* ignore */ }
      if (!ok && internalClip()) { pendingPasteMode = "all"; doPaste(null, "", mode); }
      return;
    }
  }
  doPaste(html, text, mode);
}

function doPaste(html: string | null, text: string, mode: PasteMode) {
  const target = rawSel();
  const bounded: Range = { r1: target.r1, c1: target.c1, r2: isWholeCols(target) ? target.r1 : target.r2, c2: isWholeRows(target) ? target.c1 : target.c2 };
  const clip = internalClip();
  if (clip && (isOurHtml(html) || (!html && !text) || (!html && clip.text !== undefined && text.replace(/\r\n/g, "\n").replace(/\n$/, "") === clip.text))) {
    const res = pasteInternal(wb(), sheet(), bounded, mode);
    const before = selSnapshot();
    const entries: Entry[] = [];
    if (res.cutSource && !(mode === "formats")) {
      const srcSheet = wb().sheets.find((s) => s.name.toLowerCase() === res.cutSource!.sheet.toLowerCase());
      if (srcSheet) {
        const rg = res.cutSource.range;
        const clear: CellChange[] = [];
        for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
          if (srcSheet === sheet() && inRange(res.resultRange, r, c)) continue; // overwritten by the paste anyway
          if (srcSheet.cells.has(key(r, c))) clear.push({ r, c, cell: null });
        }
        if (clear.length) entries.push(cellsEntry(srcSheet, clear, "Cut"));
      }
      clearInternalClip();
      grid().clipRange = null;
    }
    entries.push(cellsEntry(sheet(), res.changes, "Paste"));
    const combined: Entry = { label: "Paste", undo() { for (const e of [...entries].reverse()) e.undo(); }, redo() { for (const e of entries) e.redo(); } };
    pushEntry(combined, before);
    styles().invalidate();
    afterModel();
    grid().setRanges([res.resultRange], { r: res.resultRange.r1, c: res.resultRange.c1 });
    return;
  }
  if (!text && !html) return;
  const rows = parseExternal(html, text);
  if (!rows.length) return;
  const single = rows.length === 1 && rows[0].length === 1;
  const res = pasteExternal(wb(), sheet(), { r: bounded.r1, c: bounded.c1 }, rows, single && (bounded.r1 !== bounded.r2 || bounded.c1 !== bounded.c2) ? bounded : null);
  styles().invalidate();
  applyChanges(res.changes, "Paste");
  grid().setRanges([res.resultRange], { r: res.resultRange.r1, c: res.resultRange.c1 });
}

function moveRange(src: Range, dst: Range, copy: boolean) {
  const s = sheet();
  const changes: CellChange[] = [];
  const moved = new Map<number, Cell | null>();
  for (let r = src.r1; r <= src.r2; r++) for (let c = src.c1; c <= src.c2; c++) {
    const cell = s.cells.get(key(r, c)) || null;
    moved.set(key(dst.r1 + (r - src.r1), dst.c1 + (c - src.c1)), cell ? { ...cell } : null);
    if (!copy && cell) changes.push({ r, c, cell: null });
  }
  for (const [k, cell] of moved) { const r = rowOf(k), c = colOf(k); if (r >= MAXR || c >= MAXC) continue; changes.push({ r, c, cell }); }
  applyChanges(changes, copy ? "Copy cells" : "Move cells");
  grid().setRanges([dst], { r: dst.r1, c: dst.c1 });
}

function fillDirection(dir: "down" | "right") {
  const rg = selRange();
  if (dir === "down") {
    if (rg.r1 === rg.r2) { const src = { ...rg, r1: rg.r1 - 1, r2: rg.r1 - 1 }; if (src.r1 < 0) return; applyChanges(fillChanges(wb(), sheet(), src, rg), "Fill down"); return; }
    applyChanges(fillChanges(wb(), sheet(), { ...rg, r2: rg.r1 }, { ...rg, r1: rg.r1 + 1 }), "Fill down");
  } else {
    if (rg.c1 === rg.c2) { const src = { ...rg, c1: rg.c1 - 1, c2: rg.c1 - 1 }; if (src.c1 < 0) return; applyChanges(fillChanges(wb(), sheet(), src, rg), "Fill right"); return; }
    applyChanges(fillChanges(wb(), sheet(), { ...rg, c2: rg.c1 }, { ...rg, c1: rg.c1 + 1 }), "Fill right");
  }
}

function clearSelection(what: "contents" | "formats" | "all") {
  const ch = clearChanges(sheet(), grid().selection.ranges, what);
  applyChanges(ch, what === "contents" ? "Delete" : what === "formats" ? "Clear formatting" : "Clear");
}

// ---------------------------------------------------------------------------
// Rows / columns / merge / freeze

function selectedRows(): [number, number] { const rg = rawSel(); return [rg.r1, isWholeCols(rg) ? rg.r1 : rg.r2]; }
function selectedCols(): [number, number] { const rg = rawSel(); return [rg.c1, isWholeRows(rg) ? rg.c1 : rg.c2]; }

function insertRows(at: number, count: number) { structural(`Insert ${count} row${count > 1 ? "s" : ""}`, () => insertDelete(wb(), grid().sheetIdx(), "row", at, count), true); grid().selectRows(at, at + count - 1); }
function insertCols(at: number, count: number) { clearFilterState(sheet()); structural(`Insert ${count} column${count > 1 ? "s" : ""}`, () => insertDelete(wb(), grid().sheetIdx(), "col", at, count), true); grid().selectCols(at, at + count - 1); }
function deleteRows(r1: number, r2: number) { structural(`Delete row${r2 > r1 ? "s" : ""}`, () => insertDelete(wb(), grid().sheetIdx(), "row", r1, -(r2 - r1 + 1)), true); grid().setActive(r1, grid().selection.active.c); }
function deleteCols(c1: number, c2: number) { clearFilterState(sheet()); structural(`Delete column${c2 > c1 ? "s" : ""}`, () => insertDelete(wb(), grid().sheetIdx(), "col", c1, -(c2 - c1 + 1)), true); grid().setActive(grid().selection.active.r, c1); }

function hideRowsCols(kind: "row" | "col", hidden: boolean, a?: number, b?: number) {
  const [x1, x2] = kind === "row" ? (a !== undefined ? [a, b ?? a] : selectedRows()) : (a !== undefined ? [a, b ?? a] : selectedCols());
  structural(hidden ? `Hide ${kind}s` : `Unhide ${kind}s`, () => {
    for (let i = x1; i <= x2; i++) { if (kind === "row") setRowHeight(sheet(), i, undefined, hidden); else setColWidth(sheet(), i, undefined, hidden); }
  });
}
function unhideAll(kind: "row" | "col") {
  structural(`Unhide all ${kind}s`, () => {
    if (kind === "row") { for (const [r, info] of sheet().rows) if (info.hidden) sheet().rows.set(r, { ...info, hidden: false }); }
    else sheet().cols = sheet().cols.map((c) => ({ ...c, hidden: false }));
    sheet().dirty = true;
  });
}

function autoFit(kind: "col" | "row", indices: number[]) {
  structural(kind === "col" ? "Autofit column" : "Autofit row", () => {
    for (const i of indices) {
      if (kind === "col") { const px = grid().measureColumn(i); if (px > 0) setColWidth(sheet(), i, pxToChars(Math.min(1200, px + 6))); }
      else { const px = grid().measureRow(i); if (px > 0) setRowHeight(sheet(), i, Math.round((px * 72) / 96 * 100) / 100); }
    }
  });
}

async function columnWidthDialog() {
  const [c1, c2] = selectedCols();
  const cur = Math.round(charsToPx(colWidthChars(sheet(), c1)));
  const v = await promptDialog("Column width", `Width in pixels for column${c2 > c1 ? "s" : ""} ${colName(c1)}${c2 > c1 ? "–" + colName(c2) : ""}`, String(cur));
  if (v === null) return;
  const px = parseFloat(v);
  if (!(px > 0)) return;
  structural("Column width", () => { for (let c = c1; c <= c2; c++) setColWidth(sheet(), c, pxToChars(px)); });
}
async function rowHeightDialog() {
  const [r1, r2] = selectedRows();
  const cur = Math.round(rowHeightPx(sheet(), r1));
  const v = await promptDialog("Row height", `Height in pixels for row${r2 > r1 ? "s" : ""} ${r1 + 1}${r2 > r1 ? "–" + (r2 + 1) : ""}`, String(cur));
  if (v === null) return;
  const px = parseFloat(v);
  if (!(px > 0)) return;
  structural("Row height", () => { for (let r = r1; r <= r2; r++) setRowHeight(sheet(), r, Math.round((px * 72) / 96 * 100) / 100); });
}

function toggleMerge() {
  const rg = selRange();
  const s = sheet();
  const existing = s.merges.find((m) => m.r1 === rg.r1 && m.c1 === rg.c1 && m.r2 === rg.r2 && m.c2 === rg.c2);
  if (existing || (rg.r1 === rg.r2 && rg.c1 === rg.c2)) {
    const hits = s.merges.filter((m) => !(m.r2 < rg.r1 || m.r1 > rg.r2 || m.c2 < rg.c1 || m.c1 > rg.c2));
    if (!hits.length) { flash("Select two or more cells to merge"); return; }
    structural("Unmerge cells", () => { s.merges = s.merges.filter((m) => !hits.includes(m)); s.dirty = true; });
    return;
  }
  structural("Merge cells", () => {
    s.merges = s.merges.filter((m) => m.r2 < rg.r1 || m.r1 > rg.r2 || m.c2 < rg.c1 || m.c1 > rg.c2);
    s.merges.push({ ...rg });
    // keep the top-left value only
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) { if (r === rg.r1 && c === rg.c1) continue; const k = key(r, c); const cell = s.cells.get(k); if (cell && (cell.v !== null || cell.f)) s.cells.set(k, { v: null, s: cell.s }); }
    s.dirty = true;
  });
  grid().setRanges([rg], { r: rg.r1, c: rg.c1 });
}

function setFreeze(rows: number, cols: number) {
  const s = sheet();
  s.freeze = rows || cols ? { rows, cols } : null;
  s.autoFreeze = false;
  s.dirty = true;
  setDirty(true);
  grid().invalidate();
  updateToolbarState();
  flash(rows || cols ? `Frozen: ${rows} row${rows === 1 ? "" : "s"}, ${cols} column${cols === 1 ? "" : "s"}` : "Panes unfrozen");
}

// ---------------------------------------------------------------------------
// Sort / filter

/** Range to sort/filter: the selection when it spans cells, else the data block around the active cell. */
function tableRange(): Range {
  const rg = rawSel();
  const s = sheet();
  if (s.autoFilter && inRange(s.autoFilter, grid().selection.active.r, grid().selection.active.c)) return boundRange(s, s.autoFilter);
  if (rg.r1 !== rg.r2 || rg.c1 !== rg.c2) return boundRange(s, rg);
  return dataBlock(s, rg.r1, rg.c1);
}

function sortBy(col: number, ascending: boolean, range?: Range, header?: boolean) {
  const rg = range || tableRange();
  const hasHeader = header ?? (!!sheet().autoFilter && sheet().autoFilter!.r1 === rg.r1 || looksLikeHeader(sheet(), rg));
  const ch = sortChanges(sheet(), rg, col, ascending, hasHeader);
  if (!ch.length) { flash("Already sorted"); return; }
  applyChanges(ch, ascending ? "Sort A → Z" : "Sort Z → A");
  grid().setRanges([rg], { r: rg.r1, c: col });
}

async function sortDialog() {
  const rg = tableRange();
  const s = sheet();
  const hasHeader = looksLikeHeader(s, rg) || (!!s.autoFilter && s.autoFilter.r1 === rg.r1);
  const cols: { c: number; label: string }[] = [];
  for (let c = rg.c1; c <= rg.c2; c++) { const t = hasHeader ? displayOf(s.cells.get(key(rg.r1, c)), styles()) : ""; cols.push({ c, label: `${colName(c)}${t ? " - " + t : ""}` }); }
  const colSel = el("select", null, ...cols.map((x) => el("option", { value: String(x.c) }, x.label)));
  colSel.value = String(Math.max(rg.c1, Math.min(rg.c2, grid().selection.active.c)));
  const dirSel = el("select", null, el("option", { value: "asc" }, "A → Z (ascending)"), el("option", { value: "desc" }, "Z → A (descending)"));
  const headerCb = el("input", { type: "checkbox" }); headerCb.checked = hasHeader;
  const col2 = el("select", null, el("option", { value: "-1" }, "(none)"), ...cols.map((x) => el("option", { value: String(x.c) }, x.label)));
  const dir2 = el("select", null, el("option", { value: "asc" }, "A → Z"), el("option", { value: "desc" }, "Z → A"));
  const body = el("div", null,
    el("p", { style: { color: "var(--ui-muted)", margin: "0 0 8px" } }, `Sort ${rangeRef(rg)}`),
    el("label", { style: { display: "flex", alignItems: "center", gap: "6px", color: "var(--ui-fg)" } }, headerCb, "Data has a header row"),
    el("div", { class: "grid2", style: { marginTop: "8px" } }, el("label", null, "Sort by"), colSel, el("label", null, "Order"), dirSel, el("label", null, "Then by"), col2, el("label", null, "Order"), dir2));
  showDialog("Sort range", body, [
    { label: "Cancel" },
    { label: "Sort", primary: true, action: () => {
      const before = selSnapshot();
      const entries: Entry[] = [];
      const c2 = parseInt(col2.value, 10);
      if (c2 >= 0) { const ch = sortChanges(s, rg, c2, dir2.value === "asc", headerCb.checked); if (ch.length) entries.push(cellsEntry(s, ch, "Sort")); }
      const ch = sortChanges(s, rg, parseInt(colSel.value, 10), dirSel.value === "asc", headerCb.checked);
      if (ch.length) entries.push(cellsEntry(s, ch, "Sort"));
      if (!entries.length) { flash("Already sorted"); return; }
      pushEntry({ label: "Sort", undo() { for (const e of [...entries].reverse()) e.undo(); }, redo() { for (const e of entries) e.redo(); } }, before);
      afterModel();
      grid().setRanges([rg], { r: rg.r1, c: rg.c1 });
    } },
  ]);
}

function toggleFilter() {
  const s = sheet();
  if (s.autoFilter) {
    clearFilterState(s);
    s.autoFilter = null; s.dirty = true;
    setDirty(true);
    grid().invalidate();
    flash("Filter removed");
  } else {
    const rg = tableRange();
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && !s.cells.get(key(rg.r1, rg.c1))) { flash("Select the data to filter first"); return; }
    s.autoFilter = { ...rg }; s.dirty = true;
    setDirty(true);
    grid().invalidate();
    flash("Filter added - click the ▾ in a header cell");
  }
  updateToolbarState();
}

function clearFilters() {
  const s = sheet();
  if (!s.autoFilter) return;
  clearFilterState(s);
  grid().invalidate();
  flash("Filters cleared");
}

function filterPopup(c: number, r: number, x: number, y: number) {
  const s = sheet();
  if (!s.autoFilter) return;
  const header = displayOf(s.cells.get(key(r, c)), styles());
  showFilterPopup({ sheet: s, col: c, styles: styles(), x, y, header, onApply: () => { grid().invalidate(); updateStatus(); const n = s.hiddenRowsByFilter.size; flash(n ? `${n} row${n === 1 ? "" : "s"} hidden by filter` : "Showing all rows"); }, onSort: (asc) => sortBy(c, asc, boundRange(s, s.autoFilter!), true) });
}

function removeDuplicates() {
  const rg = tableRange();
  const s = sheet();
  const hasHeader = looksLikeHeader(s, rg) || (!!s.autoFilter && s.autoFilter.r1 === rg.r1);
  const seen = new Set<string>();
  const keep: number[] = [];
  let removed = 0;
  for (let r = rg.r1 + (hasHeader ? 1 : 0); r <= rg.r2; r++) {
    const sig: string[] = [];
    for (let c = rg.c1; c <= rg.c2; c++) { const v = s.cells.get(key(r, c))?.v ?? null; sig.push(v === null ? "" : typeof v === "string" ? "s" + v.toLocaleLowerCase("tr") : typeof v + ":" + String(isError(v) ? v.e : v)); }
    const k = sig.join("\u0001");
    if (seen.has(k)) { removed++; continue; }
    seen.add(k); keep.push(r);
  }
  if (!removed) { flash("No duplicate rows found"); return; }
  const changes: CellChange[] = [];
  let target = rg.r1 + (hasHeader ? 1 : 0);
  for (const r of keep) {
    if (r !== target) for (let c = rg.c1; c <= rg.c2; c++) { const cell = s.cells.get(key(r, c)); changes.push({ r: target, c, cell: cell ? movedCell(cell, target - r) : null }); }
    target++;
  }
  for (let r = target; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) if (s.cells.has(key(r, c))) changes.push({ r, c, cell: null });
  applyChanges(changes, "Remove duplicates");
  flash(`${removed} duplicate row${removed === 1 ? "" : "s"} removed`);
}

function trimWhitespace() {
  const changes: CellChange[] = [];
  for (const rg of selRanges()) for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
    const cell = sheet().cells.get(key(r, c));
    if (cell && typeof cell.v === "string" && !cell.f) { const t = cell.v.replace(/\s+/g, " ").trim(); if (t !== cell.v) changes.push({ r, c, cell: { ...cell, v: t } }); }
  }
  applyChanges(changes, "Trim whitespace");
  flash(changes.length ? `${changes.length} cell${changes.length === 1 ? "" : "s"} trimmed` : "Nothing to trim");
}

// ---------------------------------------------------------------------------
// Pivot tables

async function createPivot(refreshSheet?: Sheet) {
  const existing = refreshSheet ? pivotDefs.get(refreshSheet) : undefined;
  let srcSheet = sheet();
  if (existing) { const found = wb().sheets.find((s) => s.name.toLowerCase() === existing.sourceSheet.toLowerCase()); if (!found) { flash(`Pivot source sheet "${existing.sourceSheet}" no longer exists`); return; } srcSheet = found; }
  const def = await pivotDialog(wb(), srcSheet, existing ? existing.source : rawSel(), styles(), existing);
  if (!def) return;
  writePivot(def, refreshSheet);
}

function writePivot(def: PivotDef, target?: Sheet) {
  const res = computePivot(wb(), def, styles());
  let out = target;
  if (!out) { out = addSheet(uniqueSheetName("Pivot")); }
  const idx = wb().sheets.indexOf(out);
  if (idx !== grid().sheetIdx()) switchSheet(idx);
  const st = wb().styles;
  const boldXf = xfWith(st, 0, { font: { bold: true } });
  const fmtXf = new Map<string, number>();
  structural("Pivot table", () => {
    out!.cells.clear();
    for (const c of res.cells) {
      if (c.v === null && !c.bold) continue;
      let s = c.bold ? boldXf : 0;
      if (c.fmt && !c.bold) { let x = fmtXf.get(c.fmt); if (x === undefined) { x = xfWith(st, 0, { numFmt: c.fmt }); fmtXf.set(c.fmt, x); } s = x; }
      else if (c.fmt && c.bold) { let x = fmtXf.get("b" + c.fmt); if (x === undefined) { x = xfWith(st, boldXf, { numFmt: c.fmt }); fmtXf.set("b" + c.fmt, x); } s = x; }
      out!.cells.set(key(c.r, c.c), { v: c.v, s });
    }
    out!.maxRow = res.height - 1; out!.maxCol = res.width - 1;
    out!.freeze = { rows: def.cols.length ? 2 : 1, cols: def.rows.length }; out!.autoFreeze = false;
    out!.dirty = true;
  });
  pivotDefs.set(out, def);
  for (let c = 0; c < res.width; c++) { const px = grid().measureColumn(c); if (px > 0) setColWidth(out, c, pxToChars(Math.min(400, px + 10))); }
  grid().invalidate();
  grid().setActive(0, 0);
  flash(`Pivot table written to "${out.name}" (Data → Refresh pivot table to update)`);
}

// ---------------------------------------------------------------------------
// Links

function openLink(hl: Hyperlink) {
  if (hl.target) { F.openExternal(hl.target); return; }
  if (hl.location) {
    const ref = parseRef(hl.location.replace(/\$/g, ""));
    if (ref) { goToRef(hl.location.replace(/\$/g, "")); return; }
    const dn = wb().definedNames.find((d) => d.name === hl.location);
    if (dn) goToRef(dn.ref.replace(/\$/g, ""));
  }
}

async function insertLink() {
  const cell = activeCell();
  const { r, c } = grid().selection.active;
  const existing = grid().hyperlinkAt(r, c);
  const curUrl = existing?.target || (cell?.f && /^HYPERLINK\("([^"]*)"/i.exec(cell.f)?.[1]) || "";
  const curText = cell ? displayOf(cell, styles()) : "";
  const urlIn = el("input", { type: "text", value: curUrl, placeholder: "https://…", style: { width: "100%" } });
  const textIn = el("input", { type: "text", value: curText, style: { width: "100%" } });
  showDialog("Insert link", el("div", null, el("label", null, "Link"), urlIn, el("label", null, "Text"), textIn), [
    { label: "Cancel" },
    ...(curUrl ? [{ label: "Remove link", action: () => { applyChanges([{ r, c, cell: { v: curText, s: cell ? cell.s : 0 } }], "Remove link"); removeSheetHyperlink(r, c); } }] : []),
    { label: "OK", primary: true, action: () => {
      let url = urlIn.value.trim();
      if (!url) { urlIn.focus(); return false; }
      if (!/^[a-z]+:/i.test(url)) url = "https://" + url;
      const text = textIn.value || url;
      removeSheetHyperlink(r, c);
      applyChanges([{ r, c, cell: { v: text, f: `HYPERLINK("${url.replace(/"/g, '""')}","${text.replace(/"/g, '""')}")`, s: cell ? cell.s : 0 } }], "Insert link");
    } },
  ]);
}
function removeSheetHyperlink(r: number, c: number) {
  const s = sheet();
  const n = s.hyperlinks.length;
  s.hyperlinks = s.hyperlinks.filter((h) => { const p = parseRef(h.ref); return !p || !inRange({ r1: p.r1, c1: p.c1, r2: p.r2, c2: p.c2 }, r, c); });
  if (s.hyperlinks.length !== n) s.dirty = true;
}

// ---------------------------------------------------------------------------
// Sheet tabs

function uniqueSheetName(base: string): string {
  const names = new Set(wb().sheets.map((s) => s.name.toLowerCase()));
  if (base === "Sheet") { let n = wb().sheets.length + 1; while (names.has(`sheet${n}`)) n++; return `Sheet${n}`; }
  if (!names.has(base.toLowerCase())) return base;
  let n = 2;
  while (names.has(`${base} ${n}`.toLowerCase())) n++;
  return `${base} ${n}`;
}

function addSheet(name?: string, at?: number): Sheet {
  const w = wb();
  const id = Math.max(0, ...w.sheets.map((s) => s.sheetId)) + 1;
  const s = newSheet(name || uniqueSheetName("Sheet"), id, "", "");
  s.dirty = true;
  const idx = at ?? w.sheets.length;
  w.sheets.splice(idx, 0, s);
  setDirty(true);
  switchSheet(idx);
  renderTabs();
  return s;
}

/** Index of the nearest visible sheet at or after i (or before it). */
function nearestVisible(i: number): number {
  const w = wb();
  for (let k = i; k < w.sheets.length; k++) if (w.sheets[k].state === "visible") return k;
  for (let k = i - 1; k >= 0; k--) if (w.sheets[k].state === "visible") return k;
  return i;
}

function switchSheet(i: number) {
  const w = wb();
  if (i < 0 || i >= w.sheets.length) return;
  if (w.sheets[i].state !== "visible") { w.sheets[i].state = "visible"; w.sheets[i].dirty = true; }
  const wasEditing = app.editor?.active;
  if (wasEditing && !app.editor!.isFormula()) app.editor!.finish(null, true);
  w.sheets.forEach((s, j) => { s.view.tabSelected = j === i; });
  w.active = i;
  grid().setSheet(i);
  renderTabs();
  updateFbar();
  updateStatus();
  updateToolbarState();
  if (app.editor?.active) app.editor.placeOverlay();
}

async function deleteSheet(i: number) {
  const w = wb();
  const s = w.sheets[i];
  if (w.sheets.filter((x) => x.state === "visible" && x !== s).length < 1) { flash("A workbook needs at least one visible sheet"); return; }
  if (s.cells.size && !(await F.askYesNo(`Delete sheet "${s.name}"? This cannot be undone.`))) return;
  w.sheets.splice(i, 1);
  if (s.rId) (w.removed ||= []).push({ part: s.part, rId: s.rId });
  app.history.clear();
  // the grid may still point at the removed index: switch before anything reads sheet()
  switchSheet(nearestVisible(Math.min(i, w.sheets.length - 1)));
  setDirty(true);
}

function duplicateSheet(i: number) {
  const w = wb();
  const src = w.sheets[i];
  const id = Math.max(0, ...w.sheets.map((s) => s.sheetId)) + 1;
  const copy = newSheet(uniqueSheetName("Copy of " + src.name), id, "", "");
  copy.cells = new Map(Array.from(src.cells.entries()).map(([k, c]) => [k, { ...c }]));
  copy.rows = new Map(Array.from(src.rows.entries()).map(([k, v]) => [k, { ...v }]));
  copy.cols = src.cols.map((c) => ({ ...c }));
  copy.merges = src.merges.map((m) => ({ ...m }));
  copy.freeze = src.freeze ? { ...src.freeze } : null; copy.autoFreeze = src.autoFreeze;
  copy.view = { ...src.view, tabSelected: false };
  copy.defaultRowHeight = src.defaultRowHeight; copy.defaultColWidth = src.defaultColWidth;
  copy.autoFilter = src.autoFilter ? { ...src.autoFilter } : null;
  copy.hyperlinks = src.hyperlinks.filter((h) => !h.rId).map((h) => ({ ...h }));
  copy.tabColor = src.tabColor;
  copy.maxRow = src.maxRow; copy.maxCol = src.maxCol;
  const keep = new Set(["sheetPr", "sheetFormatPr", "printOptions", "pageMargins", "pageSetup", "headerFooter", "sheetProtection", "conditionalFormatting", "dataValidations", "rowBreaks", "colBreaks"]);
  copy.rawBlocks = src.rawBlocks.filter((b) => keep.has(b.name)).map((b) => ({ name: b.name, xml: b.xml.replace(/\s+r:id="[^"]*"/g, "") }));
  copy.rawRootAttrs = src.rawRootAttrs;
  copy.dirty = true;
  w.sheets.splice(i + 1, 0, copy);
  setDirty(true);
  switchSheet(i + 1);
}

function moveSheet(from: number, to: number) {
  const w = wb();
  if (to < 0 || to >= w.sheets.length || from === to) return;
  const [s] = w.sheets.splice(from, 1);
  w.sheets.splice(to, 0, s);
  setDirty(true);
  switchSheet(to);
}

function renameSheetInline(i: number, tab: HTMLElement) {
  const w = wb();
  const s = w.sheets[i];
  const input = el("input", { type: "text", value: s.name });
  tab.innerHTML = "";
  tab.appendChild(input);
  input.focus(); input.select();
  let done = false;
  const finish = (commit: boolean) => {
    if (done) return; done = true;
    const name = input.value.trim();
    if (commit && name && name !== s.name) {
      if (/[\[\]:*?/\\]/.test(name) || name.length > 31) flash("Sheet names cannot contain [ ] : * ? / \\ and must be at most 31 characters");
      else if (w.sheets.some((x, j) => j !== i && x.name.toLowerCase() === name.toLowerCase())) flash("A sheet with that name already exists");
      else {
        const old = s.name;
        const before = selSnapshot();
        const apply = (from: string, to: string) => { renameSheet(w, i, to); for (const sh of w.sheets) { const pd = pivotDefs.get(sh); if (pd && pd.sourceSheet === from) pd.sourceSheet = to; } renderTabs(); updateStatus(); };
        apply(old, name);
        pushEntry({ label: "Rename sheet", undo() { apply(name, old); }, redo() { apply(old, name); } }, before);
        setDirty(true);
      }
    }
    renderTabs();
    focusGrid();
  };
  input.addEventListener("keydown", (e) => { e.stopPropagation(); if (e.key === "Enter") { e.preventDefault(); finish(true); } else if (e.key === "Escape") { e.preventDefault(); finish(false); } });
  input.addEventListener("blur", () => finish(true));
}

const TAB_COLORS = ["", "C00000", "FF0000", "FFC000", "FFFF00", "92D050", "00B050", "00B0F0", "0070C0", "002060", "7030A0", "7F7F7F"];

function tabMenu(i: number, anchor: HTMLElement | { x: number; y: number }) {
  const w = wb();
  const s = w.sheets[i];
  const hidden = w.sheets.filter((x) => x.state !== "visible");
  const items: MenuItem[] = [
    { label: "Rename", action: () => { const tab = tabsEl.querySelector<HTMLElement>(`[data-idx="${i}"]`); if (tab) renameSheetInline(i, tab); } },
    { label: "Duplicate", action: () => duplicateSheet(i) },
    { label: "Delete", action: () => deleteSheet(i), disabled: w.sheets.length <= 1 },
    { sep: true },
    { label: "Move left", action: () => moveSheet(i, i - 1), disabled: i === 0 },
    { label: "Move right", action: () => moveSheet(i, i + 1), disabled: i >= w.sheets.length - 1 },
    { sep: true },
    { label: "Tab colour", submenu: TAB_COLORS.map((c) => ({ label: c ? "#" + c : "None", checked: (s.tabColor || "") === c, action: () => { s.tabColor = c || null; s.tabColorChanged = true; s.dirty = true; setDirty(true); renderTabs(); } })) },
    { label: "Hide sheet", action: () => { if (w.sheets.filter((x) => x.state === "visible").length <= 1) { flash("Cannot hide the only visible sheet"); return; } s.state = "hidden"; s.dirty = true; setDirty(true); switchSheet(nearestVisible(i)); } },
    { label: "Unhide sheet", submenu: hidden.length ? hidden.map((h) => ({ label: h.name, action: () => { h.state = "visible"; h.dirty = true; setDirty(true); switchSheet(w.sheets.indexOf(h)); } })) : [{ label: "(none hidden)", disabled: true }] },
    { sep: true },
    ...(pivotDefs.get(s) ? [{ label: "Refresh pivot table", action: () => createPivot(s) }] : []),
    { label: "Insert sheet", action: () => addSheet(undefined, i + 1) },
  ];
  showMenu(anchor, items);
}

function renderTabs() {
  const w = wb();
  tabsEl.innerHTML = "";
  const active = grid().sheetIdx();
  w.sheets.forEach((s, i) => {
    if (s.state !== "visible") return;
    const tab = el("div", { class: "sheet-tab" + (i === active ? " active" : ""), "data-idx": String(i), title: s.name }, s.name);
    if (s.tabColor) tab.appendChild(el("span", { class: "tab-color", style: { background: "#" + s.tabColor } }));
    tab.addEventListener("mousedown", (e) => { if ((e.target as HTMLElement).tagName === "INPUT") return; if (e.button === 0) { tabDragFrom = i; e.preventDefault(); if (i !== active) switchSheet(i); } });
    tab.addEventListener("mouseup", (e) => { if (e.button === 0 && tabDragFrom >= 0 && tabDragFrom !== i) { const from = tabDragFrom; tabDragFrom = -1; moveSheet(from, i); } tabDragFrom = -1; });
    tab.addEventListener("dblclick", () => renameSheetInline(i, tab));
    tab.addEventListener("contextmenu", (e) => { e.preventDefault(); tabMenu(i, { x: e.clientX, y: e.clientY }); });
    tabsEl.appendChild(tab);
  });
  const add = el("span", { class: "tab-add" }, "+");
  tooltip(add, "Add sheet", "Shift+F11");
  add.addEventListener("click", () => addSheet());
  tabsEl.appendChild(add);
  const hiddenCount = w.sheets.filter((s) => s.state !== "visible").length;
  if (hiddenCount) { const h = el("span", { class: "tab-add", style: { fontSize: "12px" } }, `${hiddenCount} hidden`); h.addEventListener("click", (e) => tabMenu(active, { x: e.clientX, y: e.clientY })); tabsEl.appendChild(h); }
  tabsEl.querySelector(".active")?.scrollIntoView({ inline: "nearest", block: "nearest" });
}
let tabDragFrom = -1;
window.addEventListener("mouseup", () => { setTimeout(() => { tabDragFrom = -1; }, 0); });

// ---------------------------------------------------------------------------
// Number formats

function formatPopup(anchor: HTMLElement) {
  showPopup(anchor, (popup, close) => {
    popup.style.minWidth = "260px";
    const cur = activeStyle().numFmt;
    for (const g of presets()) {
      popup.appendChild(el("div", { class: "item", style: { color: "var(--ui-muted)", fontSize: "11px", padding: "6px 10px 2px" } }, g.group));
      for (const p of g.items) {
        const sample = p.sample !== undefined ? formatValue(p.sample, p.code, wb().date1904).text : "";
        const row = el("div", { class: "item" + (cur === p.code ? " sel" : ""), style: { display: "flex", gap: "12px" } }, el("span", { style: { flex: "1" } }, p.label), el("span", { style: { color: "var(--ui-muted)", fontFamily: "Consolas, monospace", fontSize: "12px" } }, sample));
        row.addEventListener("click", () => { close(); applyStyle({ numFmt: p.code }, "Number format"); });
        popup.appendChild(row);
      }
    }
    popup.appendChild(el("div", { class: "menu-sep" }));
    const custom = el("div", { class: "item" }, "Custom number format…");
    custom.addEventListener("click", () => { close(); customFormatDialog(); });
    popup.appendChild(custom);
  });
}

function customFormatDialog() {
  const cur = activeStyle().numFmt;
  const input = el("input", { type: "text", value: cur, style: { width: "100%", fontFamily: "Consolas, monospace" } });
  const preview = el("div", { style: { marginTop: "8px", color: "var(--ui-muted)" } });
  const sampleV = (() => { const v = activeCell()?.v; return typeof v === "number" ? v : 1234.5678; })();
  const upd = () => { try { preview.textContent = "Preview: " + formatValue(sampleV, input.value || "General", wb().date1904).text; } catch { preview.textContent = "Invalid format"; } };
  input.addEventListener("input", upd); upd();
  const examples = ['#,##0.00', '0.0%', '"₺"#,##0.00', '#,##0.00 "€"', 'dd.mm.yyyy', 'd mmmm yyyy', 'hh:mm', '[h]:mm', '0.00E+00', '#,##0;[Red]-#,##0', '@'];
  const ex = el("div", { style: { display: "flex", flexWrap: "wrap", gap: "4px", marginTop: "8px" } });
  for (const e of examples) { const b = el("button", { type: "button", style: { fontFamily: "Consolas, monospace", fontSize: "12px" } }, e); b.addEventListener("click", () => { input.value = e; upd(); }); ex.appendChild(b); }
  showDialog("Custom number format", el("div", null, el("label", null, "Format code (Excel syntax)"), input, preview, ex), [
    { label: "Cancel" },
    { label: "Apply", primary: true, action: () => { applyStyle({ numFmt: input.value || "General" }, "Number format"); } },
  ], { width: "460px" });
}

// ---------------------------------------------------------------------------
// Colours

const PALETTE = [
  ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"],
  ["F2F2F2", "808080", "D0CECE", "D6DCE4", "D9E2F3", "FBE5D5", "EDEDED", "FFF2CC", "DEEBF6", "E2EFD9"],
  ["D8D8D8", "595959", "AEABAB", "ADB9CA", "B4C6E7", "F7CBAC", "DBDBDB", "FFE599", "BDD7EE", "C5E0B3"],
  ["BFBFBF", "3F3F3F", "757070", "8496B0", "8EAADB", "F4B183", "C9C9C9", "FFD965", "9DC3E6", "A8D08D"],
  ["A5A5A5", "262626", "3A3838", "323F4F", "2F5496", "C55A11", "7B7B7B", "BF9000", "2E75B5", "538135"],
  ["7F7F7F", "0C0C0C", "171616", "222A35", "1F3864", "833C0B", "525252", "7F6000", "1E4E79", "375623"],
];
const STANDARD = ["C00000", "FF0000", "FFC000", "FFFF00", "92D050", "00B050", "00B0F0", "0070C0", "002060", "7030A0"];

function colorPopup(anchor: HTMLElement, onPick: (hex: string | null) => void, autoLabel: string) {
  showPopup(anchor, (popup, close) => {
    const pick = (hex: string | null) => { close(); onPick(hex); };
    popup.appendChild(el("div", { class: "item", onclick: () => pick(null) }, autoLabel));
    const grid = el("div", { class: "color-grid" });
    for (const row of PALETTE) for (const c of row) grid.appendChild(el("div", { class: "color-cell", style: { background: "#" + c }, title: "#" + c, onclick: () => pick(c) }));
    popup.appendChild(grid);
    popup.appendChild(el("div", { class: "item", style: { color: "var(--ui-muted)", fontSize: "11px", padding: "2px 10px" } }, "Standard colours"));
    const std = el("div", { class: "color-grid" });
    for (const c of STANDARD) std.appendChild(el("div", { class: "color-cell", style: { background: "#" + c }, title: "#" + c, onclick: () => pick(c) }));
    popup.appendChild(std);
    const input = el("input", { type: "color", value: "#000000" });
    input.addEventListener("change", () => pick(input.value.slice(1).toUpperCase()));
    popup.appendChild(el("div", { class: "row" }, el("span", null, "More colours…"), input));
  });
}

// ---------------------------------------------------------------------------
// Find / replace

function buildFindbar() {
  const bar = $("findbar");
  bar.innerHTML = "";
  const opts = { caseSensitive: false, wholeCell: false, regex: false, formulas: false, allSheets: false, ...(app.settings.sheetFind || {}) };
  const input = el("input", { type: "text", placeholder: "Find", "aria-label": "Find" });
  const count = el("span", { class: "count" }, "");
  const prev = el("button", { class: "tb-btn", type: "button" }, "▲");
  const next = el("button", { class: "tb-btn", type: "button" }, "▼");
  tooltip(prev, "Previous match", "Shift+Enter / Shift+F3");
  tooltip(next, "Next match", "Enter / F3");
  const cb = (label: string, k: string, get: () => boolean, set: (v: boolean) => void) => {
    const c = el("input", { type: "checkbox" });
    c.checked = get();
    c.addEventListener("change", () => { set(c.checked); persist(); refresh(); });
    const l = el("label", null, c, label);
    tooltip(l, label, k);
    return { el: l, input: c };
  };
  const caseCb = cb("Match case", "Alt+C", () => opts.caseSensitive, (v) => (opts.caseSensitive = v));
  const cellCb = cb("Entire cell", "Alt+W", () => opts.wholeCell, (v) => (opts.wholeCell = v));
  const regexCb = cb("Regex", "Alt+R", () => opts.regex, (v) => (opts.regex = v));
  const formCb = cb("In formulas", "Alt+F", () => opts.formulas, (v) => (opts.formulas = v));
  const allCb = cb("All sheets", "Alt+S", () => opts.allSheets, (v) => (opts.allSheets = v));
  const close = el("button", { class: "tb-btn", type: "button" }, "✕");
  tooltip(close, "Close", "Esc");
  const replaceIn = el("input", { type: "text", placeholder: "Replace with", "aria-label": "Replace with" });
  const replBtn = el("button", { class: "tb-btn textbtn", type: "button" }, "Replace");
  const replAllBtn = el("button", { class: "tb-btn textbtn", type: "button" }, "Replace all");
  tooltip(replBtn, "Replace this match and go to the next", "Enter (in the replace box)");
  tooltip(replAllBtn, "Replace every match", `${MOD}+Enter / Alt+A`);
  const replaceRow = el("span", { style: { display: "inline-flex", gap: "6px", alignItems: "center" } }, replaceIn, replBtn, replAllBtn);
  const hint = el("span", { class: "hint" }, "");
  bar.append(input, prev, next, count, caseCb.el, cellCb.el, regexCb.el, formCb.el, allCb.el, replaceRow, hint, el("span", { style: { flex: "1" } }), close);
  const persist = () => { app.settings.sheetFind = { ...opts }; F.saveSettings(app.settings); };

  let matches: { sheet: number; r: number; c: number }[] = [];
  let cur = -1;
  let matcher: ((text: string) => boolean) | null = null;
  const buildMatcher = (): boolean => {
    const q = input.value;
    if (!q) { matcher = null; return true; }
    try {
      if (opts.regex) { const re = new RegExp(opts.wholeCell ? `^(?:${q})$` : q, opts.caseSensitive ? "u" : "iu"); matcher = (t) => re.test(t); }
      else {
        const fold = (s: string) => opts.caseSensitive ? s : s.toLocaleLowerCase("tr").replace(/i̇/g, "i");
        const fq = fold(q);
        matcher = opts.wholeCell ? (t) => fold(t) === fq : (t) => fold(t).includes(fq);
      }
      input.classList.remove("invalid");
      return true;
    } catch { input.classList.add("invalid"); matcher = null; return false; }
  };
  const cellText = (s: Sheet, cell: Cell): string => opts.formulas && cell.f !== undefined ? "=" + cell.f : (typeof cell.v === "string" ? cell.v : displayOf(cell, styles()));
  const scan = () => {
    matches = [];
    if (!matcher) return;
    const sheets = opts.allSheets ? wb().sheets.map((s, i) => [s, i] as const) : [[sheet(), grid().sheetIdx()] as const];
    for (const [s, i] of sheets) {
      if (s.state !== "visible") continue;
      const keys = Array.from(s.cells.keys()).sort((a, b) => a - b);
      for (const k of keys) { const cell = s.cells.get(k)!; if (cell.v === null && cell.f === undefined) continue; if (cell.f !== undefined && !opts.formulas) continue; if (matcher(cellText(s, cell))) matches.push({ sheet: i, r: rowOf(k), c: colOf(k) }); }
    }
  };
  const showCount = () => {
    count.textContent = input.classList.contains("invalid") ? "Invalid pattern" : matches.length ? `${cur + 1} of ${matches.length}` : input.value ? "No results" : "";
  };
  const goTo = (i: number) => {
    if (!matches.length) { cur = -1; showCount(); return; }
    cur = (i + matches.length) % matches.length;
    const m = matches[cur];
    if (m.sheet !== grid().sheetIdx()) switchSheet(m.sheet);
    grid().setActive(m.r, m.c);
    showCount();
  };
  const refresh = () => {
    buildMatcher(); scan();
    // keep position: the match at/after the active cell
    const a = grid().selection.active, si = grid().sheetIdx();
    const idx = matches.findIndex((m) => m.sheet > si || (m.sheet === si && (m.r > a.r || (m.r === a.r && m.c >= a.c))));
    cur = matches.length ? (idx >= 0 ? idx : 0) : -1;
    showCount();
    hint.textContent = "";
  };
  const step = (dir: 1 | -1) => {
    if (!matches.length) { refresh(); if (!matches.length) return; goTo(cur < 0 ? 0 : cur); return; }
    const a = grid().selection.active, si = grid().sheetIdx();
    const at = matches.findIndex((m) => m.sheet === si && m.r === a.r && m.c === a.c);
    let ni: number;
    if (at >= 0) ni = at + dir;
    else { const idx = matches.findIndex((m) => m.sheet > si || (m.sheet === si && (m.r > a.r || (m.r === a.r && m.c > a.c)))); ni = dir > 0 ? (idx >= 0 ? idx : 0) : (idx >= 0 ? idx - 1 : matches.length - 1); }
    const wrapped = ni >= matches.length || ni < 0;
    goTo(ni);
    hint.textContent = wrapped ? (dir > 0 ? "Wrapped to start" : "Wrapped to end") : "";
  };
  const replacement = (text: string): string => {
    const q = input.value, rep = replaceIn.value;
    if (opts.regex) { const re = new RegExp(opts.wholeCell ? `^(?:${q})$` : q, (opts.caseSensitive ? "" : "i") + "gu"); return text.replace(re, rep); }
    if (opts.wholeCell) return rep;
    const fold = (s: string) => opts.caseSensitive ? s : s.toLocaleLowerCase("tr");
    let out = "", i = 0; const ft = fold(text), fq = fold(q);
    while (i <= text.length) { const j = ft.indexOf(fq, i); if (j < 0 || !fq) { out += text.slice(i); break; } out += text.slice(i, j) + rep; i = j + q.length; }
    return out;
  };
  const replaceAt = (m: { sheet: number; r: number; c: number }): CellChange | null => {
    const s = wb().sheets[m.sheet];
    const cell = s.cells.get(key(m.r, m.c));
    if (!cell) return null;
    if (cell.f !== undefined && !opts.formulas) return null; // never turn a formula into a constant silently
    const text = cellText(s, cell);
    if (!matcher || !matcher(text)) return null;
    const nt = replacement(text);
    if (nt === text) return null;
    const isFormula = opts.formulas && cell.f !== undefined;
    const nc = isFormula ? inputCell(wb(), s, m.r, m.c, nt) : inputCell(wb(), s, m.r, m.c, typeof cell.v === "string" && /^[=+\-@']/.test(nt) ? "'" + nt : nt);
    return { r: m.r, c: m.c, cell: nc };
  };
  const replaceOne = () => {
    if (!matches.length) { refresh(); if (!matches.length) return; }
    const a = grid().selection.active, si = grid().sheetIdx();
    let at = matches.findIndex((m) => m.sheet === si && m.r === a.r && m.c === a.c);
    if (at < 0) { step(1); return; }
    const m = matches[at];
    const ch = replaceAt(m);
    if (ch) { if (m.sheet !== grid().sheetIdx()) switchSheet(m.sheet); applyChanges([ch], "Replace"); }
    refresh();
    // move to the next match after the current cell
    const idx = matches.findIndex((x) => x.sheet > si || (x.sheet === si && (x.r > a.r || (x.r === a.r && x.c > a.c))));
    if (matches.length) goTo(idx >= 0 ? idx : 0);
  };
  const replaceAll = () => {
    refresh();
    if (!matches.length) return;
    const bySheet = new Map<number, CellChange[]>();
    for (const m of matches) { const ch = replaceAt(m); if (ch) { let l = bySheet.get(m.sheet); if (!l) { l = []; bySheet.set(m.sheet, l); } l.push(ch); } }
    let n = 0;
    const before = selSnapshot();
    const entries: Entry[] = [];
    for (const [si, list] of bySheet) { n += list.length; entries.push(cellsEntry(wb().sheets[si], list, "Replace all")); }
    if (!n) { flash("Nothing replaced"); return; }
    pushEntry({ label: "Replace all", undo() { for (const e of [...entries].reverse()) e.undo(); }, redo() { for (const e of entries) e.redo(); } }, before);
    afterModel();
    refresh();
    flash(`Replaced ${n} cell${n === 1 ? "" : "s"}`);
  };
  let timer = 0;
  input.addEventListener("input", () => { clearTimeout(timer); timer = window.setTimeout(() => { refresh(); if (matches.length) goTo(cur < 0 ? 0 : cur); }, 120); });
  const keyHandler = (e: KeyboardEvent) => {
    const mod = e.ctrlKey || e.metaKey;
    if (e.key === "Enter" && e.target === input) { e.preventDefault(); step(e.shiftKey ? -1 : 1); }
    else if (e.key === "Enter" && e.target === replaceIn) { e.preventDefault(); if (mod) replaceAll(); else replaceOne(); }
    else if (e.key === "Escape") { e.preventDefault(); handle.hide(); }
    else if (e.altKey && !mod) {
      const k = e.key.toLowerCase();
      const map: Record<string, HTMLInputElement | (() => void)> = { c: caseCb.input, w: cellCb.input, r: regexCb.input, f: formCb.input, s: allCb.input, a: replaceAll };
      const t = map[k];
      if (t) { e.preventDefault(); if (typeof t === "function") t(); else { t.checked = !t.checked; t.dispatchEvent(new Event("change")); } }
    } else if (e.key === "F3") { e.preventDefault(); step(e.shiftKey ? -1 : 1); }
    else if (mod && e.key.toLowerCase() === "h") { e.preventDefault(); handle.show(true); }
    else if (mod && e.key.toLowerCase() === "f") { e.preventDefault(); handle.show(false); }
    e.stopPropagation();
  };
  bar.addEventListener("keydown", keyHandler);
  prev.addEventListener("click", () => step(-1));
  next.addEventListener("click", () => step(1));
  replBtn.addEventListener("click", replaceOne);
  replAllBtn.addEventListener("click", replaceAll);
  close.addEventListener("click", () => handle.hide());
  const handle = {
    visible: false,
    show(replace: boolean) {
      bar.hidden = false; this.visible = true;
      replaceRow.style.display = replace ? "inline-flex" : "none";
      // seed with the active cell's text when the box is empty
      if (!input.value) { const cell = activeCell(); if (cell && typeof cell.v === "string" && cell.v.length < 60) input.value = cell.v; }
      input.focus(); input.select();
      refresh();
    },
    hide() { bar.hidden = true; this.visible = false; focusGrid(); },
    step, refresh,
  };
  return handle;
}

// ---------------------------------------------------------------------------
// Context menus

function pasteSpecialMenu(): MenuItem[] {
  return [
    { label: "Paste values only", key: `${MOD}+Shift+V`, action: () => pasteFromSystem("values") },
    { label: "Paste format only", action: () => pasteFromSystem("formats") },
    { label: "Paste formulas only", action: () => pasteFromSystem("formulas") },
    { label: "Paste transposed", action: () => pasteFromSystem("transpose") },
  ];
}
function freezeMenu(): MenuItem[] {
  const s = sheet();
  const f = s.freeze || { rows: 0, cols: 0 };
  const { r, c } = grid().selection.active;
  return [
    { label: "No rows", checked: !f.rows, action: () => setFreeze(0, f.cols) },
    { label: "1 row", checked: f.rows === 1, action: () => setFreeze(1, f.cols) },
    { label: "2 rows", checked: f.rows === 2, action: () => setFreeze(2, f.cols) },
    { label: `Up to row ${r + 1}`, action: () => setFreeze(r + 1, f.cols) },
    { sep: true },
    { label: "No columns", checked: !f.cols, action: () => setFreeze(f.rows, 0) },
    { label: "1 column", checked: f.cols === 1, action: () => setFreeze(f.rows, 1) },
    { label: "2 columns", checked: f.cols === 2, action: () => setFreeze(f.rows, 2) },
    { label: `Up to column ${colName(c)}`, action: () => setFreeze(f.rows, c + 1) },
  ];
}
function numberFormatMenu(): MenuItem[] {
  const cur = activeStyle().numFmt;
  const out: MenuItem[] = [];
  for (const g of presets()) {
    out.push({ label: g.group, disabled: true });
    for (const p of g.items) out.push({ label: "   " + p.label, checked: cur === p.code, action: () => applyStyle({ numFmt: p.code }, "Number format") });
  }
  out.push({ sep: true }, { label: "Custom format…", action: () => customFormatDialog() });
  return out;
}

function showContextMenu(e: MouseEvent, hit: Hit) {
  closeAllPopups();
  const at = { x: e.clientX, y: e.clientY };
  const s = sheet();
  if (hit.type === "colHeader") {
    if (!grid().selection.ranges.some((rg) => isWholeCols(rg) && hit.c >= rg.c1 && hit.c <= rg.c2)) grid().selectCols(hit.c, hit.c);
    const [c1, c2] = selectedCols(); const n = c2 - c1 + 1;
    showMenu(at, [
      { label: "Cut", key: `${MOD}+X`, action: () => copySelection(true) }, { label: "Copy", key: `${MOD}+C`, action: () => copySelection(false) }, { label: "Paste", key: `${MOD}+V`, action: () => pasteFromSystem() },
      { sep: true },
      { label: `Insert ${n} column${n > 1 ? "s" : ""} left`, action: () => insertCols(c1, n) },
      { label: `Insert ${n} column${n > 1 ? "s" : ""} right`, action: () => insertCols(c2 + 1, n) },
      { label: `Delete column${n > 1 ? "s" : ""} ${colName(c1)}${n > 1 ? "–" + colName(c2) : ""}`, action: () => deleteCols(c1, c2) },
      { label: `Clear column${n > 1 ? "s" : ""}`, action: () => clearSelection("all") },
      { sep: true },
      { label: `Hide column${n > 1 ? "s" : ""}`, action: () => hideRowsCols("col", true) },
      { label: "Unhide all columns", action: () => unhideAll("col") },
      { label: "Resize column…", action: () => columnWidthDialog() },
      { label: "Fit to data", action: () => autoFit("col", Array.from({ length: n }, (_, i) => c1 + i)) },
      { sep: true },
      { label: "Sort sheet A → Z", action: () => sortBy(hit.c, true, dataBlock(s, s.autoFilter ? s.autoFilter.r1 : 0, hit.c)) },
      { label: "Sort sheet Z → A", action: () => sortBy(hit.c, false, dataBlock(s, s.autoFilter ? s.autoFilter.r1 : 0, hit.c)) },
      { label: s.autoFilter ? "Remove filter" : "Create a filter", action: () => toggleFilter() },
      { label: `Freeze up to column ${colName(hit.c)}`, action: () => setFreeze(s.freeze?.rows || 0, hit.c + 1) },
      { sep: true },
      { label: "Number format", submenu: numberFormatMenu() },
    ]);
    return;
  }
  if (hit.type === "rowHeader") {
    if (!grid().selection.ranges.some((rg) => isWholeRows(rg) && hit.r >= rg.r1 && hit.r <= rg.r2)) grid().selectRows(hit.r, hit.r);
    const [r1, r2] = selectedRows(); const n = r2 - r1 + 1;
    showMenu(at, [
      { label: "Cut", key: `${MOD}+X`, action: () => copySelection(true) }, { label: "Copy", key: `${MOD}+C`, action: () => copySelection(false) }, { label: "Paste", key: `${MOD}+V`, action: () => pasteFromSystem() },
      { sep: true },
      { label: `Insert ${n} row${n > 1 ? "s" : ""} above`, action: () => insertRows(r1, n) },
      { label: `Insert ${n} row${n > 1 ? "s" : ""} below`, action: () => insertRows(r2 + 1, n) },
      { label: `Delete row${n > 1 ? "s" : ""} ${r1 + 1}${n > 1 ? "–" + (r2 + 1) : ""}`, action: () => deleteRows(r1, r2) },
      { label: `Clear row${n > 1 ? "s" : ""}`, action: () => clearSelection("all") },
      { sep: true },
      { label: `Hide row${n > 1 ? "s" : ""}`, action: () => hideRowsCols("row", true) },
      { label: "Unhide all rows", action: () => unhideAll("row") },
      { label: "Resize row…", action: () => rowHeightDialog() },
      { label: "Fit to data", action: () => autoFit("row", Array.from({ length: n }, (_, i) => r1 + i)) },
      { sep: true },
      { label: `Freeze up to row ${hit.r + 1}`, action: () => setFreeze(hit.r + 1, s.freeze?.cols || 0) },
    ]);
    return;
  }
  if (hit.type !== "cell") return;
  const rg = rawSel();
  const nr = isWholeCols(rg) ? 1 : rg.r2 - rg.r1 + 1, nc = isWholeRows(rg) ? 1 : rg.c2 - rg.c1 + 1;
  const merged = s.merges.some((m) => m.r1 === rg.r1 && m.c1 === rg.c1 && m.r2 === rg.r2 && m.c2 === rg.c2);
  const link = grid().hyperlinkAt(hit.r, hit.c);
  showMenu(at, [
    { label: "Cut", key: `${MOD}+X`, action: () => copySelection(true) },
    { label: "Copy", key: `${MOD}+C`, action: () => copySelection(false) },
    { label: "Paste", key: `${MOD}+V`, action: () => pasteFromSystem() },
    { label: "Paste special", submenu: pasteSpecialMenu() },
    { sep: true },
    { label: `Insert ${nr} row${nr > 1 ? "s" : ""} above`, action: () => insertRows(rg.r1, nr) },
    { label: `Insert ${nc} column${nc > 1 ? "s" : ""} left`, action: () => insertCols(rg.c1, nc) },
    { label: `Delete row${nr > 1 ? "s" : ""}`, action: () => deleteRows(rg.r1, rg.r1 + nr - 1) },
    { label: `Delete column${nc > 1 ? "s" : ""}`, action: () => deleteCols(rg.c1, rg.c1 + nc - 1) },
    { sep: true },
    { label: "Delete contents", key: "Delete", action: () => clearSelection("contents") },
    { label: "Clear formatting", key: `${MOD}+\\`, action: () => clearSelection("formats") },
    { sep: true },
    { label: link ? "Edit link…" : "Insert link…", key: `${MOD}+K`, action: () => insertLink() },
    ...(link ? [{ label: "Open link", action: () => openLink(link) }] : []),
    { sep: true },
    { label: "Sort range A → Z", action: () => sortBy(grid().selection.active.c, true) },
    { label: "Sort range Z → A", action: () => sortBy(grid().selection.active.c, false) },
    { label: "Sort range…", action: () => sortDialog() },
    { label: s.autoFilter ? "Remove filter" : "Create a filter", key: `${MOD}+Shift+L`, action: () => toggleFilter() },
    { sep: true },
    { label: merged ? "Unmerge cells" : "Merge cells", action: () => toggleMerge() },
    { label: "Freeze", submenu: freezeMenu() },
    { label: "Number format", submenu: numberFormatMenu() },
    { sep: true },
    { label: "Column width…", action: () => columnWidthDialog() },
    { label: "Row height…", action: () => rowHeightDialog() },
  ]);
}

// ---------------------------------------------------------------------------
// Menu bar

function buildMenubar() {
  const bar = $("menubar");
  bar.innerHTML = "";
  const menus: { title: string; alt: string; items: () => MenuItem[] }[] = [
    { title: "File", alt: "f", items: () => [
      { label: "New spreadsheet", key: `${MOD}+N`, action: () => newSpreadsheet() },
      { label: "New document (Words)", action: () => F.openInNewWindow("new:docx") },
      { label: "New window", action: () => F.openInNewWindow("new:xlsx") },
      { label: "Open…", key: `${MOD}+O`, action: () => openFile() },
      { label: "Open recent", submenu: (app.settings.recent || []).length ? (app.settings.recent || []).map((p) => ({ label: F.basename(p), action: () => (untouched() ? openPath(p) : F.openInNewWindow(p)) })) : [{ label: "(empty)", disabled: true }] },
      { sep: true },
      { label: "Save", key: `${MOD}+S`, action: () => save() },
      { label: "Save as…", key: `${MOD}+Shift+S`, action: () => saveAs() },
      { label: "Export current sheet as CSV…", action: async () => { const t = await F.saveDialog((app.path ? app.path.replace(/\.[^.]+$/, "") : sheet().name) + ".csv", [{ name: "CSV", extensions: ["csv"] }]); if (t) { const text = csvText(); await F.writeFile(t, encodeCsv(text, "utf-8", true)); flash("Exported " + F.basename(t)); } } },
      { sep: true },
      { label: "Print…", key: `${MOD}+P`, action: () => print() },
      { sep: true },
      { label: "Close window", key: `${MOD}+W`, action: () => closeWindowRequest() },
    ] },
    { title: "Edit", alt: "e", items: () => [
      { label: "Undo" + (app.history.undoLabel() ? " " + app.history.undoLabel() : ""), key: `${MOD}+Z`, action: () => undo(), disabled: !app.history.canUndo() },
      { label: "Redo" + (app.history.redoLabel() ? " " + app.history.redoLabel() : ""), key: `${MOD}+Y`, action: () => redo(), disabled: !app.history.canRedo() },
      { sep: true },
      { label: "Cut", key: `${MOD}+X`, action: () => copySelection(true) },
      { label: "Copy", key: `${MOD}+C`, action: () => copySelection(false) },
      { label: "Paste", key: `${MOD}+V`, action: () => pasteFromSystem() },
      { label: "Paste special", submenu: pasteSpecialMenu() },
      { sep: true },
      { label: "Fill down", key: `${MOD}+D`, action: () => fillDirection("down") },
      { label: "Fill right", key: `${MOD}+R`, action: () => fillDirection("right") },
      { label: "Insert date", key: `${MOD}+;`, action: () => insertDateTime(false) },
      { label: "Insert time", key: `${MOD}+Shift+;`, action: () => insertDateTime(true) },
      { sep: true },
      { label: "Delete contents", key: "Delete", action: () => clearSelection("contents") },
      { label: "Clear formatting", key: `${MOD}+\\`, action: () => clearSelection("formats") },
      { label: "Clear all", action: () => clearSelection("all") },
      { sep: true },
      { label: "Select all", key: `${MOD}+A`, action: () => selectAllSmart() },
      { sep: true },
      { label: "Find…", key: `${MOD}+F`, action: () => app.findbar!.show(false) },
      { label: "Find and replace…", key: `${MOD}+H`, action: () => app.findbar!.show(true) },
    ] },
    { title: "View", alt: "v", items: () => [
      { label: "Freeze", submenu: freezeMenu() },
      { label: "Unfreeze", action: () => setFreeze(0, 0), disabled: !sheet().freeze },
      { sep: true },
      { label: "Show gridlines", checked: sheet().view.showGridLines, action: () => { sheet().view.showGridLines = !sheet().view.showGridLines; sheet().dirty = true; setDirty(true); grid().invalidate(); } },
      { label: "Show formulas", checked: grid().showFormulas, action: () => { grid().showFormulas = !grid().showFormulas; grid().invalidate(); } },
      { label: "Hidden rows and columns", submenu: [{ label: "Unhide all rows", action: () => unhideAll("row") }, { label: "Unhide all columns", action: () => unhideAll("col") }] },
      { sep: true },
      { label: "Zoom in", key: `${MOD}++`, action: () => setZoom(app.zoom + 0.1) },
      { label: "Zoom out", key: `${MOD}+-`, action: () => setZoom(app.zoom - 0.1) },
      { label: "Zoom 100%", key: `${MOD}+0`, action: () => setZoom(1) },
      { label: "Zoom", submenu: [50, 75, 90, 100, 125, 150, 200].map((z) => ({ label: z + "%", checked: Math.round(app.zoom * 100) === z, action: () => setZoom(z / 100) })) },
      { sep: true },
      { label: "Dark mode", key: `${MOD}+Shift+D`, checked: app.theme === "dark", action: () => toggleTheme() },
      { label: "Full screen", key: "F11", action: () => F.toggleFullscreen() },
    ] },
    { title: "Insert", alt: "i", items: () => {
      const [r1, r2] = selectedRows(); const [c1, c2] = selectedCols();
      const nr = r2 - r1 + 1, nc = c2 - c1 + 1;
      return [
        { label: `${nr} row${nr > 1 ? "s" : ""} above`, key: `${MOD}+Alt+=`, action: () => insertRows(r1, nr) },
        { label: `${nr} row${nr > 1 ? "s" : ""} below`, action: () => insertRows(r2 + 1, nr) },
        { label: `${nc} column${nc > 1 ? "s" : ""} left`, action: () => insertCols(c1, nc) },
        { label: `${nc} column${nc > 1 ? "s" : ""} right`, action: () => insertCols(c2 + 1, nc) },
        { sep: true },
        { label: "Sheet", key: "Shift+F11", action: () => addSheet() },
        { label: "Link…", key: `${MOD}+K`, action: () => insertLink() },
        { label: "Function", submenu: ["SUM", "AVERAGE", "COUNT", "MAX", "MIN", "IF", "XLOOKUP", "VLOOKUP", "CONCAT", "LEFT", "RIGHT", "TEXT", "TODAY", "SUMIF", "COUNTIF"].map((f) => ({ label: f, action: () => insertFunction(f) })) },
        { sep: true },
        { label: "Pivot table…", action: () => createPivot() },
      ];
    } },
    { title: "Format", alt: "o", items: () => {
      const cs = activeStyle();
      return [
        { label: "Number", submenu: numberFormatMenu() },
        { label: "Currency", key: `${MOD}+Shift+4`, action: () => applyStyle({ numFmt: currencyCode() }, "Currency") },
        { label: "Percent", key: `${MOD}+Shift+5`, action: () => applyStyle({ numFmt: "0.00%" }, "Percent") },
        { label: "Date", key: `${MOD}+Shift+3`, action: () => applyStyle({ numFmt: dateCode() }, "Date") },
        { label: "Plain text", action: () => applyStyle({ numFmt: "@" }, "Plain text") },
        { label: "Automatic", action: () => applyStyle({ numFmt: "General" }, "Automatic") },
        { sep: true },
        { label: "Bold", key: `${MOD}+B`, checked: cs.bold, action: () => toggleFont("bold") },
        { label: "Italic", key: `${MOD}+I`, checked: cs.italic, action: () => toggleFont("italic") },
        { label: "Underline", key: `${MOD}+U`, checked: cs.underline, action: () => toggleFont("underline") },
        { label: "Strikethrough", key: `${MOD}+5`, checked: cs.strike, action: () => toggleFont("strike") },
        { sep: true },
        { label: "Align", submenu: [
          { label: "Left", key: `${MOD}+Shift+L`, checked: cs.halign === "left", action: () => applyStyle({ halign: "left" }, "Align left") },
          { label: "Centre", key: `${MOD}+Shift+E`, checked: cs.halign === "center", action: () => applyStyle({ halign: "center" }, "Align centre") },
          { label: "Right", key: `${MOD}+Shift+R`, checked: cs.halign === "right", action: () => applyStyle({ halign: "right" }, "Align right") },
          { label: "General", checked: cs.halign === "general", action: () => applyStyle({ halign: null }, "Align") },
          { sep: true },
          { label: "Top", checked: cs.valign === "top", action: () => applyStyle({ valign: "top" }, "Align top") },
          { label: "Middle", checked: cs.valign === "center", action: () => applyStyle({ valign: "center" }, "Align middle") },
          { label: "Bottom", checked: cs.valign === "bottom", action: () => applyStyle({ valign: null }, "Align bottom") },
        ] },
        { label: "Wrap text", checked: cs.wrap, action: () => applyStyle({ wrap: !cs.wrap }, "Wrap") },
        { label: "Increase indent", action: () => applyStyle({ indent: cs.indent + 1 }, "Indent") },
        { label: "Decrease indent", action: () => applyStyle({ indent: Math.max(0, cs.indent - 1) }, "Indent") },
        { sep: true },
        { label: "Merge cells", action: () => toggleMerge() },
        { label: "Borders", submenu: bordersMenu() },
        { sep: true },
        { label: "Column width…", action: () => columnWidthDialog() },
        { label: "Row height…", action: () => rowHeightDialog() },
        { label: "Fit columns to data", action: () => { const [c1, c2] = selectedCols(); autoFit("col", Array.from({ length: c2 - c1 + 1 }, (_, i) => c1 + i)); } },
        { sep: true },
        { label: "Clear formatting", key: `${MOD}+\\`, action: () => clearFormat() },
      ];
    } },
    { title: "Data", alt: "d", items: () => [
      { label: "Sort range A → Z", action: () => sortBy(grid().selection.active.c, true) },
      { label: "Sort range Z → A", action: () => sortBy(grid().selection.active.c, false) },
      { label: "Sort range…", action: () => sortDialog() },
      { sep: true },
      { label: sheet().autoFilter ? "Remove filter" : "Create a filter", key: `${MOD}+Shift+L`, action: () => toggleFilter() },
      { label: "Clear filter criteria", action: () => clearFilters(), disabled: !sheet().autoFilter },
      { sep: true },
      { label: "Pivot table…", action: () => createPivot() },
      { label: "Refresh pivot table", action: () => createPivot(sheet()), disabled: !pivotDefs.get(sheet()) },
      { sep: true },
      { label: "Remove duplicates", action: () => removeDuplicates() },
      { label: "Trim whitespace", action: () => trimWhitespace() },
      { sep: true },
      { label: "Number locale", submenu: Object.values(LOCALES).map((l) => ({ label: localeLabel(l.id), checked: locale().id === l.id, action: () => setLocaleId(l.id) })) },
      { label: "Recalculate now", key: "F9", action: () => { const n = app.engine!.recalcAll(); grid().invalidate(); flash(`Recalculated (${n} changed)`); } },
    ] },
    { title: "Help", alt: "h", items: () => [
      { label: "Keyboard shortcuts", key: `${MOD}+/`, action: () => shortcutsDialog() },
      { label: "Check for updates…", action: () => { checkForUpdates(true); } },
      { label: "Check for updates automatically", checked: app.settings.autoUpdate !== false, action: () => { app.settings.autoUpdate = app.settings.autoUpdate === false; F.saveSettings(app.settings); } },
      { sep: true },
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
    t.addEventListener("mousedown", (e) => { e.preventDefault(); if (openIdx === i) closeAllPopups(); else openMenu(i); });
    t.addEventListener("mouseenter", () => { if (openIdx >= 0 && openIdx !== i) openMenu(i); });
    titles.push(t);
    bar.appendChild(t);
  });
  window.addEventListener("keydown", (e) => {
    if (e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey && !app.editor?.active && !dialogOpen()) {
      const i = menus.findIndex((m) => m.alt === e.key.toLowerCase());
      if (i >= 0) { e.preventDefault(); openMenu(i); }
    }
    if (openIdx >= 0 && (e.key === "ArrowLeft" || e.key === "ArrowRight")) { e.preventDefault(); openMenu((openIdx + (e.key === "ArrowRight" ? 1 : menus.length - 1)) % menus.length); }
  });
}

function localeLabel(id: string): string { return { tr: "Turkish (1.234,56 ₺ · 09.09.2026)", us: "US (1,234.56 $ · 9/9/2026)", eu: "European (1.234,56 € · 09.09.2026)", uk: "UK (1,234.56 £ · 09/09/2026)" }[id] || id; }

function bordersMenu(): MenuItem[] {
  return [
    { label: "All borders", action: () => applyBorders("all") },
    { label: "Outer borders", action: () => applyBorders("outer") },
    { label: "Thick outer border", action: () => applyBorders("thickOuter") },
    { label: "Inner borders", action: () => applyBorders("inner") },
    { sep: true },
    { label: "Top border", action: () => applyBorders("top") },
    { label: "Bottom border", action: () => applyBorders("bottom") },
    { label: "Left border", action: () => applyBorders("left") },
    { label: "Right border", action: () => applyBorders("right") },
    { sep: true },
    { label: "Clear borders", action: () => applyBorders("none") },
  ];
}

function insertFunction(name: string) {
  const rg = selRange();
  // Σ on a column/row of numbers: put the formula below/right of the selection
  if (["SUM", "AVERAGE", "COUNT", "MAX", "MIN"].includes(name) && (rg.r1 !== rg.r2 || rg.c1 !== rg.c2)) {
    const changes: CellChange[] = [];
    if (rg.r2 > rg.r1) { for (let c = rg.c1; c <= rg.c2; c++) changes.push({ r: rg.r2 + 1, c, cell: inputCell(wb(), sheet(), rg.r2 + 1, c, `=${name}(${cellRef(rg.r1, c)}:${cellRef(rg.r2, c)})`) }); }
    else { changes.push({ r: rg.r1, c: rg.c2 + 1, cell: inputCell(wb(), sheet(), rg.r1, rg.c2 + 1, `=${name}(${cellRef(rg.r1, rg.c1)}:${cellRef(rg.r1, rg.c2)})`) }); }
    applyChanges(changes, name);
    return;
  }
  if (["SUM", "AVERAGE", "COUNT", "MAX", "MIN"].includes(name)) {
    // guess the block of numbers above (or left of) the cell
    const { r, c } = grid().selection.active;
    const has = (rr: number, cc: number) => { const cell = sheet().cells.get(key(rr, cc)); return !!cell && typeof cell.v === "number"; };
    let r0 = r - 1; while (r0 >= 0 && has(r0, c)) r0--;
    if (r0 < r - 1) { beginEdit(`=${name}(${cellRef(r0 + 1, c)}:${cellRef(r - 1, c)})`, true); return; }
    let c0 = c - 1; while (c0 >= 0 && has(r, c0)) c0--;
    if (c0 < c - 1) { beginEdit(`=${name}(${cellRef(r, c0 + 1)}:${cellRef(r, c - 1)})`, true); return; }
  }
  beginEdit(`=${name}(`, true);
}

function selectAllSmart() {
  const s = sheet();
  const { r, c } = grid().selection.active;
  const block = dataBlock(s, r, c);
  const cur = rawSel();
  if (cur.r1 === block.r1 && cur.c1 === block.c1 && cur.r2 === block.r2 && cur.c2 === block.c2 || (block.r1 === block.r2 && block.c1 === block.c2)) grid().selectAll();
  else grid().setRanges([block], { r, c });
}

// ---------------------------------------------------------------------------
// Toolbar

const FONTS = ["Aptos", "Arial", "Calibri", "Cambria", "Consolas", "Courier New", "Georgia", "Segoe UI", "Tahoma", "Times New Roman", "Verdana"];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];

function buildToolbar() {
  const tb = $("toolbar");
  tb.innerHTML = "";
  const buttons = new Map<string, HTMLButtonElement>();
  const btn = (id: string, label: string, content: Node | string, onClick: (b: HTMLButtonElement) => void, keyHint?: string) => {
    const b = el("button", { class: "tb-btn" + (typeof content === "string" ? " textsm" : ""), type: "button", tabindex: -1 }, content);
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", () => onClick(b));
    tooltip(b, label, keyHint);
    buttons.set(id, b);
    tb.appendChild(b);
    return b;
  };
  const sep = () => tb.appendChild(el("div", { class: "tb-sep" }));
  btn("new", "New spreadsheet", icon("new"), () => newSpreadsheet(), `${MOD}+N`);
  btn("open", "Open", icon("open"), () => openFile(), `${MOD}+O`);
  btn("save", "Save", icon("save"), () => save(), `${MOD}+S`);
  btn("print", "Print", icon("print"), () => print(), `${MOD}+P`);
  sep();
  btn("undo", "Undo", icon("undo"), () => undo(), `${MOD}+Z`);
  btn("redo", "Redo", icon("redo"), () => redo(), `${MOD}+Y`);
  sep();
  btn("fmt", "Number formats", "123 ▾", (b) => formatPopup(b));
  btn("cur", "Currency (uses the number locale)", "₺", () => applyStyle({ numFmt: currencyCode() }, "Currency"), `${MOD}+Shift+4`);
  btn("pct", "Percent", "%", () => applyStyle({ numFmt: "0.00%" }, "Percent"), `${MOD}+Shift+5`);
  btn("decm", "Decrease decimal places", ".0", () => changeDecimals(-1));
  btn("decp", "Increase decimal places", ".00", () => changeDecimals(1));
  sep();
  // font family
  const fontInput = el("input", { type: "text", spellcheck: "false", autocomplete: "off" });
  const fontBtn = el("button", { type: "button", tabindex: -1 }, "▾");
  const fontCombo = el("div", { class: "tb-combo", id: "font-family" }, fontInput, fontBtn);
  tooltip(fontCombo, "Font");
  tb.appendChild(fontCombo);
  const applyFont = () => { const v = fontInput.value.trim(); if (v) applyStyle({ font: { name: v } }, "Font"); focusGrid(); };
  fontInput.addEventListener("keydown", (e) => { e.stopPropagation(); if (e.key === "Enter") { e.preventDefault(); applyFont(); } else if (e.key === "Escape") focusGrid(); });
  fontInput.addEventListener("focus", () => fontInput.select());
  const fontList = () => { const used = new Set<string>(FONTS); for (const f of wb().styles.fonts) if (f.name) used.add(f.name); return Array.from(used).sort((a, b) => a.localeCompare(b)); };
  fontBtn.addEventListener("mousedown", (e) => e.preventDefault());
  fontBtn.addEventListener("click", () => showPopup(fontCombo, (popup, close) => { for (const f of fontList()) popup.appendChild(el("div", { class: "item", style: { fontFamily: `"${f}"` }, onclick: () => { close(); applyStyle({ font: { name: f } }, "Font"); } }, f)); }));
  // font size
  const sizeInput = el("input", { type: "text", spellcheck: "false", autocomplete: "off" });
  const sizeBtn = el("button", { type: "button", tabindex: -1 }, "▾");
  const sizeCombo = el("div", { class: "tb-combo", id: "font-size" }, sizeInput, sizeBtn);
  tooltip(sizeCombo, "Font size");
  tb.appendChild(sizeCombo);
  const applySize = () => { const v = parseFloat(sizeInput.value.replace(",", ".")); if (v > 0 && v < 400) applyStyle({ font: { size: v } }, "Font size"); focusGrid(); };
  sizeInput.addEventListener("keydown", (e) => { e.stopPropagation(); if (e.key === "Enter") { e.preventDefault(); applySize(); } else if (e.key === "Escape") focusGrid(); });
  sizeInput.addEventListener("focus", () => sizeInput.select());
  sizeBtn.addEventListener("mousedown", (e) => e.preventDefault());
  sizeBtn.addEventListener("click", () => showPopup(sizeCombo, (popup, close) => { for (const s of SIZES) popup.appendChild(el("div", { class: "item", onclick: () => { close(); applyStyle({ font: { size: s } }, "Font size"); } }, String(s))); }));
  sep();
  btn("bold", "Bold", icon("bold"), () => toggleFont("bold"), `${MOD}+B`);
  btn("italic", "Italic", icon("italic"), () => toggleFont("italic"), `${MOD}+I`);
  btn("underline", "Underline", icon("underline"), () => toggleFont("underline"), `${MOD}+U`);
  btn("strike", "Strikethrough", icon("strike"), () => toggleFont("strike"), `${MOD}+5`);
  sep();
  const colorBtn = btn("color", "Text colour", icon("color"), (b) => colorPopup(b, (hex) => applyStyle({ font: { color: hex } }, "Text colour"), "Automatic"));
  colorBtn.classList.add("tb-color");
  const colorSw = el("span", { class: "swatch", style: { background: "#000" } }); colorBtn.appendChild(colorSw);
  const fillBtn = btn("fill", "Fill colour", icon("highlight"), (b) => colorPopup(b, (hex) => applyStyle({ fill: hex }, "Fill colour"), "No fill"));
  fillBtn.classList.add("tb-color");
  const fillSw = el("span", { class: "swatch", style: { background: "#ff0" } }); fillBtn.appendChild(fillSw);
  btn("borders", "Borders", "⊞ ▾", (b) => showMenu(b, bordersMenu()));
  sep();
  btn("halign", "Horizontal align", icon("alignLeft"), (b) => showMenu(b, [
    { label: "Left", action: () => applyStyle({ halign: "left" }, "Align left"), key: `${MOD}+Shift+L` }, { label: "Centre", action: () => applyStyle({ halign: "center" }, "Align centre"), key: `${MOD}+Shift+E` }, { label: "Right", action: () => applyStyle({ halign: "right" }, "Align right"), key: `${MOD}+Shift+R` }, { label: "General (numbers right, text left)", action: () => applyStyle({ halign: null }, "Align") },
  ]));
  btn("valign", "Vertical align", "⇅ ▾", (b) => showMenu(b, [{ label: "Top", action: () => applyStyle({ valign: "top" }, "Align top") }, { label: "Middle", action: () => applyStyle({ valign: "center" }, "Align middle") }, { label: "Bottom", action: () => applyStyle({ valign: null }, "Align bottom") }]));
  btn("wrap", "Wrap text", "↵", () => applyStyle({ wrap: !activeStyle().wrap }, "Wrap"));
  btn("merge", "Merge / unmerge cells", "⊟", () => toggleMerge());
  sep();
  btn("freeze", "Freeze rows and columns", "❄ ▾", (b) => showMenu(b, freezeMenu()));
  btn("sort", "Sort", "⇅ A-Z ▾", (b) => showMenu(b, [{ label: "Sort range A → Z", action: () => sortBy(grid().selection.active.c, true) }, { label: "Sort range Z → A", action: () => sortBy(grid().selection.active.c, false) }, { label: "Sort range…", action: () => sortDialog() }]));
  btn("filter", "Create / remove filter", "▽", () => toggleFilter(), `${MOD}+Shift+L`);
  sep();
  btn("sum", "Functions", "Σ ▾", (b) => showMenu(b, [
    ...["SUM", "AVERAGE", "COUNT", "MAX", "MIN"].map((f) => ({ label: f, action: () => insertFunction(f) })),
    { sep: true },
    ...["IF", "XLOOKUP", "VLOOKUP", "INDEX", "MATCH", "CONCAT", "LEFT", "RIGHT", "MID", "TEXT", "SUMIF", "COUNTIF", "IFERROR", "TODAY", "EOMONTH"].map((f) => ({ label: f, action: () => insertFunction(f) })),
  ]));
  btn("link", "Insert link", icon("link"), () => insertLink(), `${MOD}+K`);
  btn("find", "Find and replace", icon("find"), () => app.findbar!.show(true), `${MOD}+H`);
  return {
    update() {
      if (!app.grid) return;
      const cs = activeStyle();
      buttons.get("bold")!.classList.toggle("active", cs.bold);
      buttons.get("italic")!.classList.toggle("active", cs.italic);
      buttons.get("underline")!.classList.toggle("active", cs.underline);
      buttons.get("strike")!.classList.toggle("active", cs.strike);
      buttons.get("wrap")!.classList.toggle("active", cs.wrap);
      buttons.get("filter")!.classList.toggle("active", !!sheet().autoFilter);
      buttons.get("freeze")!.classList.toggle("active", !!sheet().freeze && !sheet().autoFreeze);
      const rg = rawSel();
      buttons.get("merge")!.classList.toggle("active", sheet().merges.some((m) => m.r1 === rg.r1 && m.c1 === rg.c1 && m.r2 === rg.r2 && m.c2 === rg.c2));
      buttons.get("undo")!.disabled = !app.history.canUndo();
      buttons.get("redo")!.disabled = !app.history.canRedo();
      buttons.get("cur")!.firstChild!.textContent = locale().currency;
      if (document.activeElement !== fontInput) fontInput.value = cs.fontName;
      if (document.activeElement !== sizeInput) sizeInput.value = String(cs.fontSize);
      colorSw.style.background = cs.color || "var(--paper-text)";
      fillSw.style.background = cs.fill || "transparent";
      const ha = buttons.get("halign")!;
      ha.replaceChild(icon(cs.halign === "center" ? "alignCenter" : cs.halign === "right" ? "alignRight" : "alignLeft"), ha.firstChild!);
    },
  };
}
function updateToolbarState() { app.toolbar?.update(); }

// ---------------------------------------------------------------------------
// Status bar

function buildStatusbar() {
  const bar = $("statusbar");
  bar.innerHTML = "";
  const stats = el("span", { class: "sb-stats" }, "");
  const msg = el("span", { style: { color: "var(--ui-accent)" } }, "");
  const prev = el("span", { style: { color: "var(--ui-muted)" } }, "");
  const locBox = el("span", { class: "sb-loc" });
  const locBtns = new Map<string, HTMLElement>();
  for (const id of ["tr", "us", "eu", "uk"]) {
    const b = el("span", { class: "sb-btn" }, id.toUpperCase());
    tooltip(b, localeLabel(id));
    b.addEventListener("click", () => setLocaleId(id));
    locBtns.set(id, b);
    locBox.appendChild(b);
  }
  const themeBtn = el("span", { class: "sb-btn" }, "☾");
  tooltip(themeBtn, "Dark mode on/off", `${MOD}+Shift+D`);
  themeBtn.addEventListener("click", () => toggleTheme());
  const zoomOut = el("span", { class: "sb-btn" }, "−");
  const zoomIn = el("span", { class: "sb-btn" }, "+");
  const range = el("input", { type: "range", min: "50", max: "300", step: "5", value: "100" });
  const zoomLabel = el("span", { class: "sb-btn", style: { minWidth: "44px", textAlign: "center" }, title: `Reset zoom (${MOD}+0)` }, "100%");
  range.addEventListener("input", () => setZoom(parseInt(range.value, 10) / 100));
  zoomOut.addEventListener("click", () => setZoom(app.zoom - 0.1));
  zoomIn.addEventListener("click", () => setZoom(app.zoom + 0.1));
  zoomLabel.addEventListener("click", () => setZoom(1));
  tooltip(zoomOut, "Zoom out", `${MOD}+-`);
  tooltip(zoomIn, "Zoom in", `${MOD}++`);
  const info = el("span", { style: { color: "var(--ui-muted)" } }, "");
  bar.append(stats, msg, prev, el("span", { class: "grow" }), info, locBox, themeBtn, el("span", { class: "zoom" }, zoomOut, range, zoomIn, zoomLabel));
  let flashTimer = 0;
  const fmtNum = (n: number) => formatValue(n, Number.isInteger(n) ? "#,##0" : "#,##0.00", false).text;
  return {
    update() {
      if (!app.grid) return;
      // selection statistics
      let sum = 0, n = 0, cnt = 0, min = Infinity, max = -Infinity;
      const s = sheet();
      const ranges = selRanges();
      const seen = ranges.length > 1 ? new Set<number>() : null; // overlapping ranges must not count twice
      const hidden = s.hiddenRowsByFilter;
      let cells = 0;
      const add = (v: Value) => { if (v !== null && v !== "") cnt++; if (typeof v === "number") { sum += v; n++; if (v < min) min = v; if (v > max) max = v; } };
      for (const rg of ranges) {
        const size = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
        cells += size;
        if (size > s.cells.size) {
          // sparse iteration: walk the cell map instead of the (larger) rectangle
          for (const [k, cell] of s.cells) { const r = (k / MAXC) | 0, c = k - r * MAXC; if (r < rg.r1 || r > rg.r2 || c < rg.c1 || c > rg.c2 || (hidden.size && hidden.has(r))) continue; if (seen) { if (seen.has(k)) continue; seen.add(k); } add(cell.v); }
          continue;
        }
        for (let r = rg.r1; r <= rg.r2; r++) { if (hidden.size && hidden.has(r)) continue; for (let c = rg.c1; c <= rg.c2; c++) { const k = key(r, c); const cell = s.cells.get(k); if (!cell) continue; if (seen) { if (seen.has(k)) continue; seen.add(k); } add(cell.v); } }
      }
      if (n >= 2) stats.textContent = `Sum: ${fmtNum(sum)}   Avg: ${fmtNum(sum / n)}   Min: ${fmtNum(min)}   Max: ${fmtNum(max)}   Count: ${cnt}`;
      else if (cells > 1) stats.textContent = cnt ? `Count: ${cnt}` : "";
      else stats.textContent = "";
      const nHidden = s.hiddenRowsByFilter.size;
      info.textContent = `${s.maxRow + 1} rows × ${s.maxCol + 1} cols${nHidden ? ` · ${nHidden} filtered` : ""}${app.dirty ? " · modified" : ""}`;
      for (const [id, b] of locBtns) b.classList.toggle("active", locale().id === id);
      themeBtn.textContent = app.theme === "dark" ? "☀" : "☾";
      const z = Math.round(app.zoom * 100);
      range.value = String(z); zoomLabel.textContent = z + "%";
    },
    flash(m: string) { msg.textContent = m; clearTimeout(flashTimer); flashTimer = window.setTimeout(() => { msg.textContent = ""; }, 3000); },
    preview(t: string) { prev.textContent = t ? "= " + t : ""; },
  };
}
function updateStatus() { app.status?.update(); }

// ---------------------------------------------------------------------------
// Print

async function print() {
  if (app.editor?.active) app.editor.finish(null, true);
  const rg = rawSel();
  const hasSel = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
  const o = await printDialog(hasSel);
  if (!o) return;
  closeDialog();
  await printSheet(sheet(), styles(), o.range === "selection" ? selRange() : null, o, wb().date1904);
  grid().invalidate();
}

// ---------------------------------------------------------------------------
// Help

const SHORTCUTS: { group: string; rows: [string, string][] }[] = [
  { group: "File", rows: [[`${MOD}+N`, "New spreadsheet"], [`${MOD}+O`, "Open"], [`${MOD}+S`, "Save"], [`${MOD}+Shift+S`, "Save as"], [`${MOD}+P`, "Print"], [`${MOD}+W`, "Close window"]] },
  { group: "Editing", rows: [["Enter / F2", "Edit the cell (Enter commits and moves down, Shift+Enter up)"], ["Type", "Start editing with the typed text; arrows commit and move"], ["Tab / Shift+Tab", "Commit and move right / left (Enter afterwards returns to the starting column)"], ["Esc", "Cancel the edit"], ["Alt+Enter", "New line inside a cell"], ["F4", "Cycle $ anchors of the reference at the caret"], ["Arrows while typing a formula", "Insert / extend a reference (Shift to extend)"], ["Delete / Backspace", "Clear contents"], [`${MOD}+Z / ${MOD}+Y`, "Undo / redo"], [`${MOD}+D / ${MOD}+R`, "Fill down / fill right"], [`${MOD}+;  /  ${MOD}+Shift+;`, "Insert today's date / current time"], [`${MOD}+K`, "Insert link"], ["F9", "Recalculate"]] },
  { group: "Clipboard", rows: [[`${MOD}+C / ${MOD}+X / ${MOD}+V`, "Copy / cut / paste (formulas shift like Google Sheets; cut keeps references)"], [`${MOD}+Shift+V`, "Paste values only"], ["Drag the selection border", "Move cells (hold Ctrl to copy)"], ["Drag the fill handle", "Fill series / copy formulas"]] },
  { group: "Selection", rows: [["Shift+Arrows", "Extend selection"], [`${MOD}+Arrows`, "Jump to the edge of the data"], [`${MOD}+Shift+Arrows`, "Extend to the edge of the data"], [`${MOD}+A`, "Select the data block, then the whole sheet"], [`${MOD}+Space / Shift+Space`, "Select column / row"], [`${MOD}+Home / ${MOD}+End`, "First / last cell"], ["Home / End", "First / last column of the row"], ["Page Up / Page Down", "Scroll a screen"], [`${MOD}+Page Up / Page Down`, "Previous / next sheet"], [`${MOD}+click`, "Add to selection"], ["Shift+click", "Extend selection"], ["Click a header", "Select column / row"], ["Double-click a column edge", "Fit column to its contents"]] },
  { group: "Formatting", rows: [[`${MOD}+B / I / U`, "Bold / italic / underline"], [`${MOD}+5`, "Strikethrough"], [`${MOD}+Shift+1`, "Number 1,234.56"], [`${MOD}+Shift+2`, "Time"], [`${MOD}+Shift+3`, "Date"], [`${MOD}+Shift+4`, "Currency"], [`${MOD}+Shift+5`, "Percent"], [`${MOD}+Shift+6`, "Scientific"], [`${MOD}+\\`, "Clear formatting"], [`${MOD}+Shift+L / E / R`, "Align left / centre / right"]] },
  { group: "Rows, columns, sheets", rows: [[`${MOD}+Alt+=`, "Insert rows (or columns when columns are selected)"], [`${MOD}+Alt+-`, "Delete selected rows / columns"], [`${MOD}+Shift+L`, "Create / remove filter"], ["Shift+F11", "New sheet"], ["Right-click a header or tab", "Row / column / sheet menu"], ["Double-click a tab", "Rename sheet"], ["Drag a tab", "Reorder sheets"]] },
  { group: "View", rows: [[`${MOD}+F / ${MOD}+H`, "Find / find and replace"], ["F3 / Shift+F3", "Next / previous match"], [`${MOD}++ / ${MOD}+- / ${MOD}+0`, "Zoom"], [`${MOD}+wheel`, "Zoom"], ["Shift+wheel", "Horizontal scroll"], [`${MOD}+Shift+D`, "Dark mode"], ["F11", "Full screen"], ["Alt+F, E, V, I, O, D, H", "Open menus"], [`${MOD}+/`, "This list"]] },
];

function shortcutsDialog() {
  const table = el("table", { class: "keys" });
  for (const g of SHORTCUTS) {
    table.appendChild(el("tr", null, el("th", { colspan: "2" }, g.group)));
    for (const [k, l] of g.rows) table.appendChild(el("tr", null, el("td", null, k), el("td", null, l)));
  }
  showDialog("Keyboard shortcuts and mouse", el("div", { style: { maxHeight: "70vh", overflow: "auto" } }, table), [{ label: "Close", primary: true }], { width: "600px" });
}

function aboutDialog() {
  showDialog("About OfficeMini", el("div", null,
    el("p", null, el("b", null, `OfficeMini Sheets ${__APP_VERSION__}`), " — a small, fast editor for Excel (.xlsx) and CSV files."),
    el("p", null, "Files are saved with a round-trip strategy: sheets you did not touch, charts, drawings and other parts are preserved exactly as they were."),
    el("p", { style: { color: "var(--ui-muted)" } }, `Press ${MOD}+/ for keyboard shortcuts.`),
    el("p", { style: { color: "var(--ui-muted)", fontSize: "12px" } }, `Build ${__BUILD_INFO__}`),
  ), [{ label: "Close", primary: true }]);
}

// ---------------------------------------------------------------------------
// Keyboard

function installGlobalKeys() {
  window.addEventListener("keydown", (e) => {
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    const kl = k.toLowerCase();
    if ((mod && ["p", "s", "o", "f", "h", "g", "r", "u", "j", "d", "n", "w", "q", "+", "=", "-", "0", "k", "l"].includes(kl)) || k === "F5" || k === "F3" || k === "F7" || (mod && e.shiftKey && kl === "i") || (e.altKey && k === "ArrowLeft")) e.preventDefault();
    if (dialogOpen()) return;
    const t = e.target as HTMLElement;
    const inText = t !== keyProxy && (t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement);
    const editing = !!app.editor?.active;
    if (inText && !editing) {
      // find bar, name box, toolbar combos: only the app-wide file shortcuts
      if (mod && !e.shiftKey && !e.altKey && ["s", "o", "p", "n", "w", "q"].includes(kl)) { e.preventDefault(); ({ s: save, o: openFile, p: print, n: newSpreadsheet, w: closeWindowRequest, q: closeWindowRequest } as Record<string, () => void>)[kl](); }
      return;
    }
    if (editing) {
      // the cell editor's own handler stops propagation; anything arriving here targets another element
      return;
    }
    if (!app.grid) return;
    const g = grid();
    const shift = e.shiftKey, alt = e.altKey;
    const prevent = () => e.preventDefault();
    // ---- navigation
    if (k.startsWith("Arrow") && !alt) {
      prevent();
      const dr = k === "ArrowUp" ? -1 : k === "ArrowDown" ? 1 : 0, dc = k === "ArrowLeft" ? -1 : k === "ArrowRight" ? 1 : 0;
      g.moveActive(dr, dc, shift, mod);
      app.tabStartCol = null;
      return;
    }
    if (alt && (k === "ArrowUp" || k === "ArrowDown") && !mod) { prevent(); nextSheet(k === "ArrowDown" ? 1 : -1); return; }
    if (k === "Tab") { prevent(); if (app.tabStartCol === null) app.tabStartCol = g.selection.active.c; g.moveActive(0, shift ? -1 : 1, false, false); return; }
    if (k === "Enter") {
      prevent();
      if (shift) { g.moveActive(-1, 0, false, false); return; }
      if (alt || mod) { beginEdit(null, false); return; }
      if (app.tabStartCol !== null && !g.selection.ranges[0] || false) { /* unreachable */ }
      beginEdit(null, false);
      return;
    }
    if (k === "F2") { prevent(); beginEdit(null, false); return; }
    if (k === "Escape") { prevent(); closeAllPopups(); if (app.findbar?.visible) app.findbar.hide(); g.clipRange = null; cancelCut(); g.setActive(g.selection.active.r, g.selection.active.c); return; }
    if (k === "Delete" || (k === "Backspace" && !mod)) { prevent(); clearSelection("contents"); return; }
    if (k === "Home") { prevent(); if (mod) g.setActive(0, 0, shift); else g.setActive(g.selection.active.r, 0, shift); return; }
    if (k === "End") { prevent(); const s = sheet(); if (mod) g.setActive(Math.max(0, s.maxRow), Math.max(0, s.maxCol), shift); else { let c = Math.max(0, s.maxCol); while (c > 0 && !s.cells.has(key(g.selection.active.r, c))) c--; g.setActive(g.selection.active.r, c, shift); } return; }
    if (k === "PageDown" || k === "PageUp") { prevent(); if (mod) nextSheet(k === "PageDown" ? 1 : -1); else if (alt) { g.host.scrollLeft += (k === "PageDown" ? 1 : -1) * (g.host.clientWidth - 60); } else g.pageMove(k === "PageDown" ? 1 : -1, shift); return; }
    if (k === " " && (mod || shift)) { prevent(); const rg = rawSel(); if (mod && shift) g.selectAll(); else if (mod) g.selectCols(rg.c1, rg.c2); else g.selectRows(rg.r1, rg.r2); return; }
    if (k === "F9") { prevent(); const n = app.engine!.recalcAll(); g.invalidate(); flash(`Recalculated (${n} changed)`); return; }
    if (k === "F11" && shift) { prevent(); addSheet(); return; }
    if (k === "F11") { prevent(); F.toggleFullscreen(); return; }
    if (k === "F1") { prevent(); shortcutsDialog(); return; }
    if (k === "F3") { prevent(); if (app.findbar?.visible) app.findbar.step(shift ? -1 : 1); else app.findbar!.show(false); return; }
    if (k === "F4" && !mod) { prevent(); redo(); return; }
    // ---- with modifier
    if (mod && !alt) {
      if (shift) {
        switch (kl) {
          case "v": prevent(); pasteFromSystem("values"); return;
          case "c": case "x": return; // native copy/cut events on the proxy
          case "z": prevent(); redo(); return;
          case "s": prevent(); saveAs(); return;
          case "d": prevent(); toggleTheme(); return;
          case "l": prevent(); toggleFilter(); return;
          case "e": prevent(); applyStyle({ halign: "center" }, "Align centre"); return;
          case "r": prevent(); applyStyle({ halign: "right" }, "Align right"); return;
          case ";": case ":": prevent(); insertDateTime(true); return;
          case "arrowdown": case "arrowup": return;
        }
        // digit shortcuts arrive as symbols on US layouts and as digits with e.code elsewhere
        const digit = /^Digit(\d)$/.exec(e.code)?.[1] || ({ "!": "1", "@": "2", "#": "3", "$": "4", "%": "5", "^": "6" } as Record<string, string>)[k];
        if (digit) {
          prevent();
          const codes: Record<string, [string, string]> = { "1": ["#,##0.00", "Number"], "2": ["hh:mm", "Time"], "3": [dateCode(), "Date"], "4": [currencyCode(), "Currency"], "5": ["0.00%", "Percent"], "6": ["0.00E+00", "Scientific"] };
          const c = codes[digit];
          if (c) applyStyle({ numFmt: c[0] }, c[1]);
          return;
        }
        return;
      }
      switch (kl) {
        case "s": prevent(); save(); return;
        case "o": prevent(); openFile(); return;
        case "p": prevent(); print(); return;
        case "n": prevent(); newSpreadsheet(); return;
        case "w": case "q": prevent(); closeWindowRequest(); return;
        case "f": prevent(); app.findbar!.show(false); return;
        case "h": prevent(); app.findbar!.show(true); return;
        case "z": prevent(); undo(); return;
        case "y": prevent(); redo(); return;
        case "c": case "x": case "v":
          // Let the native copy/cut/paste events fire on the proxy textarea (synchronous, no permissions).
          if (t === keyProxy) return;
          prevent(); if (kl === "v") pasteFromSystem("all"); else copySelection(kl === "x"); return;
        case "a": prevent(); selectAllSmart(); return;
        case "b": prevent(); toggleFont("bold"); return;
        case "i": prevent(); toggleFont("italic"); return;
        case "u": prevent(); toggleFont("underline"); return;
        case "5": prevent(); toggleFont("strike"); return;
        case "d": prevent(); fillDirection("down"); return;
        case "r": prevent(); fillDirection("right"); return;
        case "k": prevent(); insertLink(); return;
        case ";": prevent(); insertDateTime(false); return;
        case "\\": prevent(); clearFormat(); return;
        case "=": case "+": prevent(); setZoom(app.zoom + 0.1); return;
        case "-": prevent(); setZoom(app.zoom - 0.1); return;
        case "0": prevent(); setZoom(1); return;
        case "/": prevent(); shortcutsDialog(); return;
        case "l": prevent(); applyStyle({ halign: "left" }, "Align left"); return;
        case "e": prevent(); applyStyle({ halign: "center" }, "Align centre"); return;
        case "pageup": case "pagedown": return;
      }
      return;
    }
    if (mod && alt) {
      if (k === "=" || k === "+") { prevent(); const rg = rawSel(); if (isWholeCols(rg)) insertCols(rg.c1, rg.c2 - rg.c1 + 1); else { const [r1, r2] = selectedRows(); insertRows(r1, r2 - r1 + 1); } return; }
      if (k === "-") { prevent(); const rg = rawSel(); if (isWholeCols(rg)) deleteCols(rg.c1, rg.c2); else { const [r1, r2] = selectedRows(); deleteRows(r1, r2); } return; }
      // AltGr on Windows / Linux arrives as Ctrl+Alt: characters such as @ ₺ # $ [ ] { } \ | start an edit
      if (e.ctrlKey && e.altKey && !e.metaKey && k.length === 1 && !/[\x00-\x1f]/.test(k)) { prevent(); if (app.welcome) hideWelcome(); beginEdit(k, true); return; }
      return;
    }
    if (alt) return;
    // ---- typing starts an edit
    if (k.length === 1 && !mod) { prevent(); if (app.welcome) hideWelcome(); beginEdit(k, true); return; }
  });
}

function nextSheet(dir: 1 | -1) {
  const w = wb();
  let i = grid().sheetIdx();
  for (let n = 0; n < w.sheets.length; n++) { i = (i + dir + w.sheets.length) % w.sheets.length; if (w.sheets[i].state === "visible") { switchSheet(i); return; } }
}

// ---------------------------------------------------------------------------
// Boot

async function boot() {
  app.settings = await F.loadSettings();
  app.theme = app.settings.theme || "light";
  document.documentElement.setAttribute("data-theme", app.theme);
  setDarkMode(app.theme === "dark");
  setLocaleId(app.settings.locale || detectLocale(), false);
  app.zoom = app.settings.sheetZoom || 1;
  buildWorkspace();
  app.editor = new CellEditor(editorHost, celled, finput);
  installWorkbook(blankWorkbook(), null, "new");
  app.toolbar = buildToolbar();
  app.status = buildStatusbar();
  app.findbar = buildFindbar();
  buildMenubar();
  installGlobalKeys();
  updateToolbarState();
  updateStatus();
  gridHost.addEventListener("mousedown", () => { if (app.welcome) hideWelcome(); });

  const openPaths = (paths: string[]) => { for (const p of paths) { if (untouched()) openPath(p); else F.openInNewWindow(p); } };
  await F.onBackendEvent("open-files", async () => openPaths(await F.takePendingOpens()));
  await F.onBackendEvent<string>("menu", (id) => { if (id === "quit") closeWindowRequest(); else if (id === "undo") undo(); else if (id === "redo") redo(); });

  const params = new URLSearchParams(location.search);
  let file = params.get("file");
  const isNew = params.has("new") || (file && file.startsWith("new:"));
  if (isNew) file = null;
  else if (!file) { const args = await F.cliArgs(); if (args.length) { file = args[0]; for (const extra of args.slice(1)) F.openInNewWindow(extra); } }
  if (file) await openPath(file);
  else if (!isNew) { const recovered = await findRecoveries(); await buildWelcome(recovered); showWelcome(); }
  if (isNew && (params.get("new") === "csv" || file === "new:csv")) newWorkbook("csv");
  app.dirty = false; updateTitle();
  focusGrid();
  requestAnimationFrame(() => revealWindow());
  setTimeout(revealWindow, 400);
  setInterval(() => { autosaveTick(); }, 60000);
  if (app.settings.autoUpdate !== false) setTimeout(() => { checkForUpdates(false); }, 8000);
  F.onCloseRequested(requestClose);
  F.onFileDrop((paths) => openPaths(paths), (over) => $("workspace").classList.toggle("drop-target", over));
  (window as any).om = { app, openPath, save, writeTo, wb: () => app.wb, grid: () => app.grid, sheet, applyChanges, setZoom, setTheme, setLocaleId, loadXlsx, writeXlsx };
}

function revealWindow() { if (app.shown) return; app.shown = true; F.showWindow(); }

boot().catch((e) => { console.error(e); document.body.innerHTML = `<pre style="padding:20px;color:#b00">Failed to start: ${(e as Error).stack || e}</pre>`; F.showWindow(); });

export {};
void quoteSheet; void mergeAt; void rowHeightPx; void colHidden;
