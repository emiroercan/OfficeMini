# OfficeMini Web — handoff

State as of 2026-09-10, on top of v0.3.3. A working prototype of a browser build lives in this
repo behind a build flag, and a Chrome extension built from the same bundle lives beside it. It
is **not** wired into the desktop app in any way, and the desktop build is byte-for-byte
unaffected. This document is what you need to take it further, either here or in a repo of its
own. The extension has its own document: `docs/WEB-EXTENSION.md`.

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
| Chrome extension: a downloaded `.docx`/`.xlsx` — link, export button or blob — opens in the editor | works, `docs/WEB-EXTENSION.md` |
| Files dropped on the window | works in the browser builds too, with a writable handle |

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
npm run build:ext  # the same browser build, packaged as a Chrome extension → dist-ext/
npm run pack:ext   # …and zipped for the Chrome Web Store
```

`--mode ext` is `--mode web` with three differences: a relative `base` (an extension page is not
served from a root), `dist-ext`, and a plugin that copies `extension/` and the app's icons in
after the bundle. There is no third code path.

`vite.config.ts` switches `outDir` on the mode, so the web builds **can never overwrite `dist/`**,
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
/web/3/Report.docx       → a granted FileSystemFileHandle
/opfs/recovery/<id>.docx → the origin private file system
/net/1/Report.docx       → bytes the extension handed in from a link
```

so `basename`, `dirname`, `extname` and `joinPath` keep working and nothing outside this file
knows the difference. `readFile`/`writeFile` dispatch on the prefix: `/opfs/` → OPFS, a
registered path → the handle, `/net/` → the bytes or the URL they came from, anything else →
`fetch` (which is how the bundled samples and `?file=` links still load).

`adoptEntryUrl()` is the extension's door in: it turns `?inbox=<token>` or `?src=<url>` into a
`/net/` path and rewrites the query string to the `?file=` the rest of the app reads, before
either editor module loads.

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

The extension added `isRemote`, `remoteUrl` and `warnOnClose` here, used at three call sites in
each editor — `save()`, `addRecent()` and the close warning — each guarded by `__WEB_BUILD__`
rather than `F.isWeb`. **That distinction matters:** see the gotcha at the end of this file.

**Rule for anything you add: the desktop build must not change.** Keep new code behind `isWeb`
or in `files-web.ts`, and re-run `docs/WEB-TESTS.md` §1 — which now compares `dist/` against a
build of `HEAD` byte-for-byte, not just greps it.

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
   drags ~296 KB of shared chunk. See the note under *Worth doing early*. In the extension it is
   pure dead weight — Chrome does the updating.
6. **The extension needs *Allow access to file URLs*** for export buttons, blob downloads and
   local files: it opens a download by reading the file Chrome saved, and Chrome has no API to
   request that switch. Plain links work without it, and `extension/file-access.html` explains
   the rest and finishes the job once the switch is on.

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
- **Decide the non-Chromium story** before building anything else on top of the picker. The
  extension sidesteps it (it is a Chrome extension), the PWA does not.
- **Narrow the extension's `<all_urls>` host permission** to `optional_host_permissions`
  requested per site, before any Chrome Web Store submission. `docs/WEB-EXTENSION.md` has the
  rest of that list.

## Gotchas learned the hard way

- The browser pane used for automated testing reports `innerWidth === 0` when it is collapsed,
  which makes every hit-test return `outside`. Anything geometric must be tested in a real
  window — see `docs/WEB-TESTS.md`.
- `showOpenFilePicker` needs transient user activation, so it cannot be called from an automated
  evaluate. Stub *only* the dialog and hand back a real handle; everything after it is then the
  app's own code. `WEB-TESTS.md` has the recipe.
- Vite's `define` values must be JSON: `__WEB_BUILD__: JSON.stringify(web)`, not `web`.
- **`F.isWeb` does not fold; `__WEB_BUILD__` does.** Inside `files.ts` the const is local and
  rollup drops the dead branch — which is why the whole web backend disappears from `dist/`.
  From `main.ts` or `sheets/app.ts` it is an import from a *different chunk*, so it survives as
  a real call and the desktop bundle grows. Writing `if (F.isWeb) …` in an editor added 20
  bytes to two desktop chunks; `if (__WEB_BUILD__) …` made them byte-identical again.
- **Do not link `node_modules` into a comparison worktree.** `git worktree remove --force`
  follows a Windows junction and deletes the real directory behind it.
- **Stubbed extension tests prove very little.** A `chrome` stub accepted the first design, which
  then failed every export button, blob download and local file in a real browser. Test the
  extension in a browser: `npm run test:ext`.
- **Branded Chrome ignores `--load-extension`** since 137. `Extensions.loadUnpacked` over
  `--remote-debugging-pipe` (with `--enable-unsafe-extension-debugging`) still works, and the
  debugging port can stay open alongside. Edge still honours the flag.
- **Chrome disables a pipe-loaded extension when its file access is switched off**
  (`unsupportedDeveloperExtension`); it comes back when switched on. That path is checked by hand; testing is Chrome only.
- **A clean test profile hides what real profiles have.** Every test passed while dropping a file
  failed for the owner: Google's old *Office Editing for Docs, Sheets & Slides* claims Office
  types, so Chrome shows the file in the tab (in that profile, failing with "Couldn't load
  plugin") instead of downloading it, and no download event ever fires. Look at what is installed in the failing profile — its extension
  manifests are readable — before trusting a green run.
