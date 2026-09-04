// Bridge to the Tauri backend, with a browser fallback used during development
// (so the editor can be exercised in a plain browser tab).

export const isTauri = typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;

type Invoke = (cmd: string, args?: any, options?: any) => Promise<any>;
let _invoke: Invoke | null = null;
async function invoke(): Promise<Invoke> {
  if (!_invoke) {
    const m = await import("@tauri-apps/api/core");
    _invoke = m.invoke as Invoke;
  }
  return _invoke;
}

export interface FileFilter { name: string; extensions: string[]; }
export const DOC_FILTERS: FileFilter[] = [
  { name: "Documents", extensions: ["docx", "md", "markdown", "txt"] },
  { name: "Word Document", extensions: ["docx"] },
  { name: "Markdown", extensions: ["md", "markdown"] },
  { name: "All files", extensions: ["*"] },
];

// Browser fallback state
const browserFiles = new Map<string, File>();

export function basename(path: string): string {
  const i = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
  return i >= 0 ? path.slice(i + 1) : path;
}
export function dirname(path: string): string {
  const i = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
  return i >= 0 ? path.slice(0, i) : "";
}
export function extname(path: string): string {
  const b = basename(path);
  const i = b.lastIndexOf(".");
  return i >= 0 ? b.slice(i + 1).toLowerCase() : "";
}
export function joinPath(dir: string, name: string): string {
  if (!dir) return name;
  const sep = dir.includes("\\") ? "\\" : "/";
  return dir.endsWith(sep) ? dir + name : dir + sep + name;
}

export async function readFile(path: string): Promise<Uint8Array> {
  if (isTauri) {
    const inv = await invoke();
    const buf: ArrayBuffer = await inv("read_file", { path });
    return new Uint8Array(buf);
  }
  const f = browserFiles.get(path);
  if (f) return new Uint8Array(await f.arrayBuffer());
  const res = await fetch(path);
  if (!res.ok) throw new Error("Cannot load " + path + " (" + res.status + ")");
  return new Uint8Array(await res.arrayBuffer());
}

export async function readTextFile(path: string): Promise<string> {
  const bytes = await readFile(path);
  return new TextDecoder("utf-8").decode(bytes);
}

export async function writeFile(path: string, data: Uint8Array): Promise<void> {
  if (isTauri) {
    const inv = await invoke();
    await inv("write_file", data, { headers: { "x-path": encodeURIComponent(path) } });
    return;
  }
  // Browser (dev): post to the Vite dev-save endpoint, which writes samples/out/<name>.
  (window as any).__lastSaved = { path, data };
  const res = await fetch("/__save?name=" + encodeURIComponent(basename(path)), { method: "POST", body: data as BlobPart });
  if (!res.ok) throw new Error("dev save failed: " + res.status);
}

export async function fileExists(path: string): Promise<boolean> {
  if (isTauri) { const inv = await invoke(); return inv("file_exists", { path }); }
  return browserFiles.has(path);
}

export async function openDialog(filters: FileFilter[] = DOC_FILTERS, multiple = false): Promise<string[] | null> {
  if (isTauri) {
    const { open } = await import("@tauri-apps/plugin-dialog");
    const r = await open({ multiple, filters, title: "Open document" });
    if (!r) return null;
    return Array.isArray(r) ? r : [r];
  }
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.multiple = multiple;
    input.accept = filters.flatMap((f) => f.extensions.filter((e) => e !== "*").map((e) => "." + e)).join(",");
    input.onchange = () => {
      const files = Array.from(input.files || []);
      if (!files.length) return resolve(null);
      const paths = files.map((f) => { const p = "browser:" + f.name; browserFiles.set(p, f); return p; });
      resolve(paths);
    };
    input.click();
  });
}

export async function saveDialog(defaultPath: string | null, filters: FileFilter[]): Promise<string | null> {
  if (isTauri) {
    const { save } = await import("@tauri-apps/plugin-dialog");
    const r = await save({ defaultPath: defaultPath || undefined, filters, title: "Save document" });
    return r || null;
  }
  const name = prompt("Save as (file name)", defaultPath ? basename(defaultPath) : "document.docx");
  return name ? "browser:" + name : null;
}

export async function askYesNo(message: string, title = "OfficeMini"): Promise<boolean> {
  if (isTauri) {
    const { ask } = await import("@tauri-apps/plugin-dialog");
    return ask(message, { title, kind: "warning" });
  }
  return confirm(message);
}

/** Yes / No / Cancel. Returns "yes" | "no" | "cancel". */
export async function askSaveChanges(name: string): Promise<"yes" | "no" | "cancel"> {
  if (isTauri) {
    const { ask } = await import("@tauri-apps/plugin-dialog");
    const save = await ask(`Save changes to "${name}"?\n\nYour changes will be lost if you don't save them.`, {
      title: "OfficeMini", kind: "warning", okLabel: "Save", cancelLabel: "Don't Save",
    });
    return save ? "yes" : "no";
  }
  const r = confirm(`Save changes to "${name}"? OK = save, Cancel = discard`);
  return r ? "yes" : "no";
}

export async function showMessage(message: string, title = "OfficeMini", kind: "info" | "warning" | "error" = "info"): Promise<void> {
  if (isTauri) {
    const { message: msg } = await import("@tauri-apps/plugin-dialog");
    await msg(message, { title, kind });
    return;
  }
  alert(message);
}

export async function cliArgs(): Promise<string[]> {
  if (isTauri) { const inv = await invoke(); return inv("cli_args"); }
  return [];
}

export async function openInNewWindow(path?: string): Promise<void> {
  if (isTauri) { const inv = await invoke(); await inv("open_window", { path: path || null }); return; }
  window.open(location.pathname + (path ? "?file=" + encodeURIComponent(path) : ""), "_blank");
}

export async function openExternal(url: string): Promise<void> {
  if (isTauri) { const inv = await invoke(); await inv("open_external", { url }); return; }
  window.open(url, "_blank", "noopener");
}

export interface Settings {
  recent?: string[];
  zoom?: number;
  view?: "page" | "continuous";
  showMarks?: boolean;
  theme?: "light" | "dark";
  findOptions?: { caseSensitive?: boolean; wholeWord?: boolean; regex?: boolean; preserveCase?: boolean };
}

export async function loadSettings(): Promise<Settings> {
  try {
    if (isTauri) { const inv = await invoke(); return JSON.parse(await inv("load_settings")) || {}; }
    return JSON.parse(localStorage.getItem("officemini.settings") || "{}");
  } catch { return {}; }
}

export async function saveSettings(s: Settings): Promise<void> {
  try {
    const json = JSON.stringify(s);
    if (isTauri) { const inv = await invoke(); await inv("save_settings", { json }); return; }
    localStorage.setItem("officemini.settings", json);
  } catch { /* ignore */ }
}

// ---- Window control ---------------------------------------------------------

export async function appWindow() {
  if (!isTauri) return null;
  const { getCurrentWindow } = await import("@tauri-apps/api/window");
  return getCurrentWindow();
}

export async function setWindowTitle(title: string) {
  document.title = title;
  const w = await appWindow();
  if (w) await w.setTitle(title);
}

export async function showWindow() {
  const w = await appWindow();
  if (w) { await w.show(); await w.setFocus(); }
}

export async function closeWindow() {
  const w = await appWindow();
  if (w) await w.destroy();
  else window.close();
}

/** Register a close-request handler; return false from the callback to keep the window open. */
export async function onCloseRequested(handler: () => Promise<boolean>) {
  const w = await appWindow();
  if (!w) {
    window.addEventListener("beforeunload", (e) => { /* browser: nothing reliable */ void e; });
    return;
  }
  await w.onCloseRequested(async (event) => {
    event.preventDefault();
    if (await handler()) await w.destroy();
  });
}

/** Files dropped on the window (Tauri) -> callback with paths. */
export async function onFileDrop(cb: (paths: string[]) => void, hover?: (over: boolean) => void) {
  const w = await appWindow();
  if (!w) return;
  await w.onDragDropEvent((event) => {
    const p = event.payload as any;
    if (p.type === "enter" || p.type === "over") hover?.(true);
    else if (p.type === "leave") hover?.(false);
    else if (p.type === "drop") { hover?.(false); if (p.paths?.length) cb(p.paths); }
  });
}

export function toggleFullscreen() {
  appWindow().then(async (w) => {
    if (!w) { if (document.fullscreenElement) document.exitFullscreen(); else document.documentElement.requestFullscreen(); return; }
    const fs = await w.isFullscreen();
    await w.setFullscreen(!fs);
  });
}
