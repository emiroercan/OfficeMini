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
grep -l showOpenFilePicker dist/assets/*.js       # PASS: no output
grep -l showOpenFilePicker dist-web/assets/*.js   # PASS: one or more files
```

- `npx tsc --noEmit` clean.
- `git diff --stat` on a web-only change touches no file under `src-tauri/`.
- `dist/` and `dist-web/` are different directories; the Tauri bundler only ever ships `dist/`.

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

## 7. Before calling a change done

1. §1 isolation — both greps.
2. §2 backend round trip.
3. §3 stubbed open/save, plus the manual picker check if the picker path changed.
4. §4 on at least one `.docx` and one `.xlsx`.
5. `npx tsc --noEmit`.
