# OfficeMini Web — handoff

State as of 2026-09-10, on top of v0.3.3. A working prototype of a browser build lives in this
repo behind a build flag. It is **not** wired into the desktop app in any way, and the desktop
build is byte-for-byte unaffected. This document is what you need to take it further, either
here or in a repo of its own.

## What OfficeMini is

A Tauri 2 desktop app: Rust shell (`src-tauri/`) plus a TypeScript frontend (`src/`) with no UI
framework. Two editors share one shell — Words (`src/main.ts`, ProseMirror, `.docx`/`.md`) and
Sheets (`src/sheets/app.ts`, canvas grid, `.xlsx`/`.csv`) — picked by `src/boot.ts` from the
file extension.

The important fact for this project: **the document engine is already browser code.** DOCX
parse/write, XLSX parse/write, the ProseMirror layer, the canvas grid, the formula engine, find
and replace, pagination and printing contain zero references to Tauri. `npm run dev` has always
served the whole app in a plain browser tab. The only coupling is 36 `isTauri` branches, 21 of
them inside `src/files.ts`, which exists to be that seam.

```
$ grep -rln "isTauri\|__TAURI\|@tauri-apps" src/
src/files.ts      21 branches   ← the bridge
src/main.ts        6
src/sheets/app.ts  6
src/updater.ts     3
```

## What the prototype does

| Capability | State |
|---|---|
| Open a file through the native picker, no upload | works |
| `Ctrl+S` writes back to that same file, no download | works |
| Save as… to a new file | works |
| Autosave/recovery copies | works, in the origin private file system |
| Recovery offered on the next start | works within a session; see *Handle persistence* |
| Print | unchanged — it was always `window.print()` |
| Everything else (editing, formulas, filters, find…) | unchanged, it never touched Tauri |

Verified end to end by opening a real 61 KB `.xlsx` through a `FileSystemFileHandle`, editing a
cell and saving: 63,506 bytes written back to the same handle, still a valid zip, re-reading the
file returns the edited cell. No download at any point.

## How it is put together

### The build flag

`__WEB_BUILD__` is a Vite `define`, i.e. a compile-time constant, set from the build mode:

```
npm run dev        # Tauri frontend, __WEB_BUILD__ === false   → dist/
npm run dev:web    # browser build,  __WEB_BUILD__ === true    → dist-web/
npm run build:web  # production browser build
```

`vite.config.ts` switches `outDir` on the mode, so the web build **can never overwrite `dist/`**,
which is what the Tauri bundler ships. Because the flag is a constant, each build folds the other
one's branches away. This is checked, not assumed:

```
$ grep -l showOpenFilePicker dist/assets/*.js       # → nothing. The app build has no web code.
$ grep -l showOpenFilePicker dist-web/assets/*.js   # → the web build does.
```

### The backend: `src/files-web.ts`

The app passes file **paths** around as strings; the browser has only handles. Every handle the
user grants is registered under a synthetic path that still looks like one:

```
/web/3/Report.docx     → a granted FileSystemFileHandle
/opfs/recovery/<id>.docx → the origin private file system
```

so `basename`, `dirname`, `extname` and `joinPath` keep working and nothing outside this file
knows the difference. `readFile`/`writeFile` dispatch on the prefix: `/opfs/` → OPFS, a
registered path → the handle, anything else → `fetch` (which is how the bundled samples and
`?file=` links still load).

OPFS was chosen for recovery copies because it needs no permission prompt, survives reloads and
is invisible to the user — exactly the properties an autosave sidecar wants.

### The seam: `src/files.ts`

Nine functions gained one line each, ahead of the existing `isTauri` branch:

```ts
export const isWeb = __WEB_BUILD__;
...
export async function readFile(path: string): Promise<Uint8Array> {
  if (isWeb) return web.readFile(path);
  if (isTauri) { ... }
```

`readFile`, `writeFile`, `fileExists`, `openDialog`, `saveDialog`, `recoveryDir`, `listFiles`,
`deleteFile`, `fileMtime`, plus `openInNewWindow`. The recovery guards in `main.ts` and
`sheets/app.ts` changed from `!F.isTauri` to `!F.isTauri && !F.isWeb`, which folds back to the
original expression in the desktop build.

**Rule for anything you add: the desktop build must not change.** Keep new code behind `isWeb`
or in `files-web.ts`, and re-run the two `grep -l` checks above.

## Known gaps

1. **Handle persistence.** Handles die with the tab. The recent-files list and cross-reload
   recovery need handles stored in IndexedDB (handles are structured-cloneable) and
   `requestPermission()` re-granted on the next visit — one click per file, unavoidable.
   Until then, "Recent" is dead weight in the web build.
2. **Side files.** Saving a `.md` with images writes `name_files/image1.png` next to the
   document. A file handle gives no route to its folder, so those writes throw with a clear
   message. The fix is `showDirectoryPicker()` for the folder, or embedding images as data URIs.
   `writeFile` in `files-web.ts` carries the note.
3. **New tabs open empty.** `openInNewWindow` cannot hand a handle to another tab, so it opens a
   blank one. Opening several files at once, or a second file while one is loaded, means picking
   it again in the new tab. A shared `SharedWorker`/BroadcastChannel handle registry could fix
   this; a simpler answer is to keep everything in one tab.
4. **Chromium only.** File System Access is Chrome/Edge. Firefox and Safari fall through to the
   existing `<input type="file">` path: documents open, but saving back to them is impossible —
   they would need a download. Decide whether that fallback ships or is refused with a message.
5. **`src/updater.ts`** is dead weight in the web build (it early-returns on `isTauri`), and it
   drags ~296 KB of shared chunk. See the note under *Worth doing early*.

## Not goals (from the owner, explicitly)

- **No Explorer/Finder double-click integration.** The desktop app does that job. Do not spend
  time on the PWA `file_handlers` manifest or `launchQueue`.
- **Do not change, limit or entangle the desktop app.** This is a separate product.

## Extracting it into its own repo

The prototype is deliberately small so that this is easy:

1. Copy `src/`, `index.html`, `vite.config.ts`, `tsconfig.json`, `package.json`.
2. Drop `src-tauri/`, `src/updater.ts`, and the `@tauri-apps/*` dependencies.
3. Delete the `isTauri` branches — with Tauri gone, `files.ts` collapses to the web backend plus
   the `<input type="file">` fallback, and `files-web.ts` can be merged back into it.
4. Keep `docs/WEB-GOALS.md` and `docs/WEB-TESTS.md`; this file stops being useful once the split
   has happened.

Do the split **after** the gaps above are settled, not before: while both live here, every fix
is testable against the desktop app's behaviour, which is the reference implementation.

## Worth doing early

- **Handle persistence (gap 1)** unlocks recent files and recovery, and is the difference between
  a demo and something usable daily.
- **Drop `updater.ts` from the web entry.** It imports `./ui/dialogs`, which pulls ProseMirror
  into a chunk the Sheets editor loads but never uses (~296 KB). Importing `./ui/dialog-core`
  instead fixes it for both builds; it is a one-line change that has been sitting unmade.
- **Decide the non-Chromium story** before building anything else on top of the picker.

## Gotchas learned the hard way

- The browser pane used for automated testing reports `innerWidth === 0` when it is collapsed,
  which makes every hit-test return `outside`. Anything geometric must be tested in a real
  window — see `docs/WEB-TESTS.md`.
- `showOpenFilePicker` needs transient user activation, so it cannot be called from an automated
  evaluate. Stub *only* the dialog and hand back a real handle; everything after it is then the
  app's own code. `WEB-TESTS.md` has the recipe.
- Vite's `define` values must be JSON: `__WEB_BUILD__: JSON.stringify(web)`, not `web`.
