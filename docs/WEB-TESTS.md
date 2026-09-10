# OfficeMini Web — tests

How to prove the browser build works, and how to prove it has not damaged the desktop app.
Written to be executed by an agent: every check states what to run and what a pass looks like.

There is no test runner in this project. Checks are either shell commands or scripts evaluated
in a live page — the techniques in §6 make the second kind reliable.

---

## 1. Isolation — the desktop app is untouched

Run first, and after every change. **A failure here outranks every other result in this file.**

```bash
npm run build          # → dist/
npm run build:web      # → dist-web/
npm run build:ext      # → dist-ext/
grep -l showOpenFilePicker dist/assets/*.js       # PASS: no output
grep -l showOpenFilePicker dist-web/assets/*.js   # PASS: one or more files
grep -l showOpenFilePicker dist-ext/assets/*.js   # PASS: one or more files
```

The grep proves the web backend is absent. The stronger check is that **nothing at all moved**:
build `HEAD` somewhere else and compare. A web-only change that alters one byte of `dist/` has
broken this rule, whatever the greps say.

```bash
git worktree add ../om-base HEAD
cd ../om-base && npm ci && npx vite build              # → ../om-base/dist
diff -r ../om-base/dist/assets <repo>/dist/assets      # PASS: no output
```

- Never link `node_modules` into that worktree. `git worktree remove --force` follows a symlink
  or a Windows junction and deletes what it points at.
- `npx tsc --noEmit` clean.
- `git diff --stat` on a web-only change touches no file under `src-tauri/`.
- `dist/`, `dist-web/` and `dist-ext/` are different directories; the Tauri bundler only ever
  ships `dist/`.

**What makes a branch fold.** `__WEB_BUILD__` is replaced with a literal, so `if (__WEB_BUILD__)`
disappears wherever it is written. `F.isWeb` only folds *inside* `files.ts`, where it is a local
const — read from `main.ts` or `sheets/app.ts` it is an import from another chunk and survives as
a real call, and the desktop bundle grows. Guard new call sites in the editors with
`__WEB_BUILD__` directly, and re-run the byte comparison above.

## 2. Backend — the file bridge

Serve the web build and evaluate in the page (§6). Every call goes through the **public** bridge
`src/files.ts`, not `files-web.ts`, so the seam is covered too.

```js
const F = await import('/src/files.ts');
F.isWeb === true && F.isTauri === false
```

**OPFS round trip** — the path autosave uses:

```js
const dir = await F.recoveryDir();                     // PASS: "/opfs/recovery"
const p = F.joinPath(dir, 'probe.docx');
await F.writeFile(p, new TextEncoder().encode('hello opfs'));
await F.fileExists(p);                                 // PASS: true
new TextDecoder().decode(await F.readFile(p));         // PASS: "hello opfs"
(await F.listFiles(dir)).map(f => f.name + ':' + f.size);  // PASS: ["probe.docx:10"]
typeof await F.fileMtime(p);                           // PASS: "number"
await F.deleteFile(p);
await F.fileExists(p);                                 // PASS: false
```

**Synthetic paths behave like paths** — this is what lets the rest of the app stay ignorant:

```js
F.basename('/web/3/Report.docx') === 'Report.docx'
F.extname('/web/3/Report.docx')  === 'docx'
F.dirname('/web/3/Report.docx')  === '/web/3'
```

## 3. The point of the whole project — open and save in place

The native picker needs user activation and cannot be driven automatically. Stub **only the
dialog** and hand back a genuine `FileSystemFileHandle`; everything after that line is the app's
own code, so the test is real.

```js
const F = await import('/src/files.ts');
const src = new Uint8Array(await (await fetch('/samples/sheets/<a real file>.xlsx')).arrayBuffer());
const root = await navigator.storage.getDirectory();
const h = await root.getFileHandle('picked.xlsx', { create: true });
const w = await h.createWritable(); await w.write(src); await w.close();

window.showOpenFilePicker = async () => [h];           // the ONLY stub
const paths = await F.openDialog([{ name: 'Spreadsheets', extensions: ['xlsx'] }], false);
// PASS: ["/web/1/picked.xlsx"]

await window.om.openPath(paths[0]);
window.om.applyChanges([{ r: 0, c: 0, cell: { v: 'WRITTEN BY THE WEB BUILD', s: 0 } }], 'probe');
await window.om.writeTo(paths[0]);                     // PASS: true

const after = new Uint8Array(await (await h.getFile()).arrayBuffer());
after[0] === 0x50 && after[1] === 0x4b                 // PASS: still a zip
after.length !== src.length                            // PASS: the file on disk changed
window.om.loadXlsx(after).sheets[0].cells.get(0).v      // PASS: "WRITTEN BY THE WEB BUILD"
```

Reference result from the prototype: 61,613 bytes in, 63,506 bytes back, round trip clean.

**Then do it once by hand**, because the stub skips the dialog and the permission grant:
`npm run dev:web`, open a real `.docx` from the picker, type, `Ctrl+S`, and confirm in Explorer
that the file's timestamp moved and Word still opens it. Confirm the Downloads folder stays
empty — that is the claim being tested.

## 4. Fidelity — files come back the way they went in

Same parsers as the desktop app, so this is a comparison, not new ground.

- Open each file in `samples/` and `samples/sheets/` in **both** builds, save, and compare the
  two outputs byte-for-byte. **PASS: identical.** Any difference is a web-backend bug, because
  the writer is shared.
- Validate a saved `.docx`/`.xlsx` with LibreOffice headless (see the desktop workflow in
  `README.md`) — it must open without repair prompts.
- A file the app only partly understands (charts, pivot tables, macros) must keep those parts.

## 5. Behaviour under the browser's rules

- `Ctrl+S` saves the document and never triggers the browser's Save Page.
- `Ctrl+P` prints the document, and the output matches the desktop app's.
- Reload with unsaved changes: recovery copy is in OPFS afterwards (`F.listFiles`).
- Deny the write permission when prompted: a clear error, no data loss, the document stays open.
- Firefox and Safari: whatever §*Open questions* in `WEB-GOALS.md` settles, the behaviour must be
  deliberate and explained on screen — never a silent failure.
- Known gaps that must fail *loudly*, not silently (see `WEB-HANDOFF.md`): saving a `.md` that
  has images, and opening a second document into a new tab.

## 6. How to run these

**Serving the build**

```bash
npm run dev:web     # Vite in web mode on :1420
```

**Evaluating in the page.** Two options:

- The in-app browser pane: `preview_start` with `url: "http://localhost:1420/index.html?file=new:xlsx"`,
  then `javascript_tool` to evaluate. Fine for logic, OPFS and the file bridge.
- A real window over the Chrome DevTools Protocol, needed for anything geometric.

**The pane reports `innerWidth === 0` when it is collapsed**, so the page has no layout and every
hit test returns `outside`. Column widths, header clicks, the filter button and autofit must be
tested in a real window:

```powershell
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS='--remote-debugging-port=9222'
Start-Process chrome.exe "--app=http://localhost:1420/"
```

then drive it over CDP — `http://127.0.0.1:9222/json/list` for the target, a WebSocket and
`Runtime.evaluate` (Node's global `WebSocket` is enough). This is how the desktop app's grid was
tested; the same script works against Chrome.

**Useful hooks**

`window.om` exposes the app's internals in both editors: `app`, `sheet()`, `grid()`, `openPath`,
`writeTo`, `save`, `applyChanges`, `loadXlsx`, `writeXlsx`, `view()`. Grid internals marked
`private` in TypeScript are plain properties at runtime — `om.grid().cellRect(r, c)`,
`hitAt(x, y)`, `filterBtnRect(...)` are all callable from a test.

Synthetic `KeyboardEvent`s on `#keyproxy` drive the Sheets keyboard; ProseMirror needs
`view.someProp('handleKeyDown')` instead.

## 7. The extension

`docs/WEB-EXTENSION.md` describes what is being tested. The page half runs against
`npm run dev:ext` in the browser pane; the Chrome half needs the unpacked extension.

**The entry points.** Both turn into an ordinary `?file=` before either editor loads.

```js
// ?src= - a URL the page fetches itself (the context menu, and the fallback)
location.href = '/index.html?src=' + encodeURIComponent('/samples/sheets/<a real file>.xlsx') + '&name=Test.xlsx';
// after load:
const F = await import('/src/files.ts');
window.om.app.path               // PASS: "/net/1/Test.xlsx"
F.isRemote(window.om.app.path)   // PASS: true
F.remoteUrl(window.om.app.path)  // PASS: the URL it came from
location.search                  // PASS: still carries src= and name=, so a reload reopens it
```

```js
// ?inbox= - bytes the service worker already fetched, handed over through the OPFS
const token = crypto.randomUUID();
const src = new Uint8Array(await (await fetch('/samples/<a real file>.docx')).arrayBuffer());
const dir = await (await navigator.storage.getDirectory()).getDirectoryHandle('inbox', { create: true });
const w = await (await dir.getFileHandle(token, { create: true })).createWritable();
await w.write(src); await w.close();
location.href = '/index.html?inbox=' + token + '&name=Test.docx';
// PASS: the document opens, and reloading the tab opens it again - the inbox copy is kept.
```

**Ctrl+S becomes Save as…** — the point of `isRemote`. Stub only the dialog, as in §3:

```js
const target = await (await navigator.storage.getDirectory()).getFileHandle('saved-as.xlsx', { create: true });
let calls = 0, suggested = null;
window.showSaveFilePicker = async (o) => { calls++; suggested = o?.suggestedName; return target; };
window.om.applyChanges([{ r: 0, c: 0, cell: { v: 'EDITED', s: 0 } }], 'probe');
await window.om.save();
calls === 1          // PASS: save() routed to saveAs(), it did not throw
suggested            // PASS: the link's file name
window.om.app.path   // PASS: "/web/1/saved-as.xlsx" - a real handle now
window.om.loadXlsx(new Uint8Array(await (await target.getFile()).arrayBuffer()))  // PASS: reads back, edited
JSON.parse(localStorage['officemini.settings']).recent   // PASS: no "/net/" entry was added
```

**The service worker's rules**, with a stubbed `chrome`. `extOf`, `nameFromUrl` and `targetName`
are exported from `background.js` for this.

```js
globalThis.chrome = { /* storage.sync, downloads, runtime, contextMenus, tabs stubs */ };
const bg = await import('/extension/background.js');
const s  = await import('/extension/settings.js');
const active = s.activeExtensions(await s.loadSettings());

bg.targetName({ url: 'https://x/a/Report.docx', filename: 'C:\\Users\\a\\Downloads\\Report.docx' })  // "Report.docx"
bg.targetName({ url: 'https://drive/export?id=9', filename: '', mime: '...spreadsheetml.sheet' })    // "export.xlsx"
bg.targetName({ url: 'https://x/get/Q3%20Plan.xlsx?token=abc', filename: '' })                       // "Q3 Plan.xlsx"
active.has('pdf') || active.has('zip')   // PASS: false - only Office types are taken over
active.has('md')                         // PASS: false by default
```

Then drive the whole handler, which must cancel, erase, stash and open exactly one tab:

```js
await globalThis.__onCreated({ id: 11, url: location.origin + '/samples/sheets/<file>.xlsx', filename: 'x.xlsx' });
// PASS: cancel(11) and erase(11) both called; one tabs.create; its ?inbox= token names an OPFS
//       file whose byte length equals the file on the server; ?src= is kept as the fallback.
await globalThis.__onCreated({ id: 12, url: 'https://x/manual.pdf', filename: 'manual.pdf' });   // PASS: ignored
await globalThis.__onCreated({ id: 13, url: 'blob:http://x/abc', filename: 'x.docx' });          // PASS: ignored
```

**In Chrome, by hand**, because none of the above touches the parts only Chrome has:

- Load `dist-ext` unpacked, click a real `.docx` link: a tab opens with the document, and the
  Downloads folder gets nothing. That is the whole claim.
- Right-click a link → *Open link in OfficeMini*.
- `Ctrl+S`, choose a location, confirm the file opens in Word. `Ctrl+S` again writes in place.
- Turn the master switch off in the popup: the same link downloads normally again.
- A `.pdf` and a `.zip` link still download.
- Reload the editor tab with unsaved changes: the browser's own prompt appears first, and after
  reloading the document is still there.

## 8. Before calling a change done

1. §1 isolation — the greps, and the byte-for-byte comparison against `HEAD`.
2. §2 backend round trip.
3. §3 stubbed open/save, plus the manual picker check if the picker path changed.
4. §4 on at least one `.docx` and one `.xlsx`.
5. §7 if anything under `extension/`, `files-web.ts` or either editor's `save()` moved.
6. `npx tsc --noEmit`.
