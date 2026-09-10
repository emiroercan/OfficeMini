// File System Access backend for the browser build (`vite --mode web`, __WEB_BUILD__).
//
// The rest of the app passes file *paths* around as strings, which the browser has no notion
// of. Every handle the user grants is therefore registered under a synthetic path that still
// looks like one - "/web/3/Report.docx" - so basename, dirname, extname and joinPath keep
// working untouched and nothing outside this file needs to know the difference. Recovery
// copies live in the origin private file system under "/opfs/...", which needs no permission
// and survives a reload.
//
// Nothing here is imported by the Tauri build: __WEB_BUILD__ is a compile-time false there,
// so the calls fold away and the module is dropped from the bundle.

const OPFS = "/opfs";

interface FilePickerType { description: string; accept: Record<string, string[]> }
interface OpenPickerOptions { multiple?: boolean; types?: FilePickerType[]; excludeAcceptAllOption?: boolean }
interface SavePickerOptions { suggestedName?: string; types?: FilePickerType[] }
type Perm = "granted" | "denied" | "prompt";
interface Handle {
  name: string;
  getFile(): Promise<File>;
  createWritable(): Promise<{ write(data: BufferSource): Promise<void>; close(): Promise<void> }>;
  queryPermission?(d: { mode: string }): Promise<Perm>;
  requestPermission?(d: { mode: string }): Promise<Perm>;
}
type Picker = {
  showOpenFilePicker?(o?: OpenPickerOptions): Promise<Handle[]>;
  showSaveFilePicker?(o?: SavePickerOptions): Promise<Handle>;
};

export function supported(): boolean {
  return typeof window !== "undefined" && typeof (window as unknown as Picker).showOpenFilePicker === "function";
}

// ---- handle registry --------------------------------------------------------

const handles = new Map<string, Handle>();
let seq = 0;

function register(h: Handle): string {
  const path = `/web/${++seq}/${h.name}`;
  handles.set(path, h);
  return path;
}

/** The handle behind a path, or null when the path is not one we handed out. */
export function handleFor(path: string): Handle | null { return handles.get(path) || null; }

async function writable(h: Handle) {
  if (h.queryPermission && (await h.queryPermission({ mode: "readwrite" })) !== "granted") {
    if (!h.requestPermission || (await h.requestPermission({ mode: "readwrite" })) !== "granted") {
      throw new Error("Permission to write this file was not granted");
    }
  }
  return h.createWritable();
}

// ---- origin private file system (recovery copies) ---------------------------

async function opfsDir(create = true): Promise<FileSystemDirectoryHandle> {
  const root = await navigator.storage.getDirectory();
  return root.getDirectoryHandle("recovery", { create });
}

function opfsName(path: string): string { return path.slice(OPFS.length + 1).replace(/^recovery\//, ""); }

export function isOpfs(path: string): boolean { return path === OPFS || path.startsWith(OPFS + "/"); }

export async function recoveryDir(): Promise<string> { await opfsDir(); return OPFS + "/recovery"; }

export interface FileInfo { name: string; path: string; size: number; mtime: number }

export async function listFiles(dir: string): Promise<FileInfo[]> {
  if (!isOpfs(dir)) return [];
  const out: FileInfo[] = [];
  try {
    const d = await opfsDir(false);
    // values() is an async iterator on the directory handle
    for await (const entry of (d as unknown as { values(): AsyncIterable<FileSystemHandle> }).values()) {
      if (entry.kind !== "file") continue;
      const f = await (entry as unknown as Handle).getFile();
      out.push({ name: entry.name, path: dir + "/" + entry.name, size: f.size, mtime: Math.floor(f.lastModified / 1000) });
    }
  } catch { /* no recovery directory yet */ }
  return out;
}

export async function deleteFile(path: string): Promise<void> {
  if (!isOpfs(path)) return;
  try { (await opfsDir(false)).removeEntry(opfsName(path)); } catch { /* already gone */ }
}

export async function fileMtime(path: string): Promise<number | null> {
  const h = handles.get(path);
  if (h) { try { return Math.floor((await h.getFile()).lastModified / 1000); } catch { return null; } }
  if (!isOpfs(path)) return null;
  try {
    const d = await opfsDir(false);
    const f = await (await d.getFileHandle(opfsName(path))).getFile();
    return Math.floor(f.lastModified / 1000);
  } catch { return null; }
}

// ---- reading and writing ----------------------------------------------------

export async function readFile(path: string): Promise<Uint8Array> {
  if (isOpfs(path)) {
    const d = await opfsDir(false);
    const f = await (await d.getFileHandle(opfsName(path))).getFile();
    return new Uint8Array(await f.arrayBuffer());
  }
  const h = handles.get(path);
  if (h) return new Uint8Array(await (await h.getFile()).arrayBuffer());
  // Not a granted handle: a URL the page can fetch (the bundled samples, a ?file= link).
  const res = await fetch(path);
  if (!res.ok) throw new Error("Cannot load " + path + " (" + res.status + ")");
  return new Uint8Array(await res.arrayBuffer());
}

export async function writeFile(path: string, data: Uint8Array): Promise<void> {
  if (isOpfs(path)) {
    const d = await opfsDir();
    const w = await (await d.getFileHandle(opfsName(path), { create: true })).createWritable();
    await w.write(data as BufferSource);
    await w.close();
    return;
  }
  const h = handles.get(path);
  if (!h) {
    // Side files (a Markdown document's images) sit next to the document, and a file handle
    // gives no way to reach its folder. Writing those needs showDirectoryPicker - see
    // docs/WEB-HANDOFF.md; until then the document itself saves and the assets do not.
    throw new Error("No write permission for " + path + " (open or Save as it first)");
  }
  const w = await writable(h);
  await w.write(data as BufferSource);
  await w.close();
}

export async function fileExists(path: string): Promise<boolean> {
  if (handles.has(path)) return true;
  if (!isOpfs(path)) return false;
  try { await (await opfsDir(false)).getFileHandle(opfsName(path)); return true; } catch { return false; }
}

// ---- pickers ----------------------------------------------------------------

export interface FileFilter { name: string; extensions: string[] }

const MIME: Record<string, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
  md: "text/markdown", markdown: "text/markdown", txt: "text/plain",
  csv: "text/csv", tsv: "text/tab-separated-values",
};

function pickerTypes(filters: FileFilter[]): FilePickerType[] {
  return filters
    .filter((f) => !f.extensions.includes("*"))
    .map((f) => {
      const accept: Record<string, string[]> = {};
      for (const ext of f.extensions) {
        const mime = MIME[ext] || "application/octet-stream";
        (accept[mime] ||= []).push("." + ext);
      }
      return { description: f.name, accept };
    });
}

export async function openDialog(filters: FileFilter[], multiple: boolean): Promise<string[] | null> {
  const picker = window as unknown as Picker;
  if (!picker.showOpenFilePicker) return null;
  try {
    const picked = await picker.showOpenFilePicker({ multiple, types: pickerTypes(filters) });
    return picked.length ? picked.map(register) : null;
  } catch { return null; }   // the user dismissed the picker
}

export async function saveDialog(defaultPath: string | null, filters: FileFilter[]): Promise<string | null> {
  const picker = window as unknown as Picker;
  if (!picker.showSaveFilePicker) return null;
  const name = defaultPath ? defaultPath.split(/[/\\]/).pop() || undefined : undefined;
  try {
    return register(await picker.showSaveFilePicker({ suggestedName: name, types: pickerTypes(filters) }));
  } catch { return null; }
}
