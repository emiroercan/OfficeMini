# OfficeMini Web — goals

## The one-sentence goal

The same editor, opening and saving real files in a Chromium tab, with no upload, no download
and no server — a browser build that a person could use for a day's work instead of the desktop
app, and not notice much.

## Why this is worth building

The document engine is already browser code; only file access ever needed the Rust shell. So the
web version is not a rewrite, it is a different backend for `src/files.ts`. The prototype proved
the hard part works: a picked file is opened, edited and written back **to the same file on
disk**, with the browser's own permission model and no download.

What the web build buys that the desktop app cannot:

- Nothing to install, nothing to update, no code signing, no antivirus warnings.
- A link opens the editor. Anywhere, on any machine, including one you do not administer.
- The natural home for the browser-extension idea: a `.docx` link on a website opens in the
  editor instead of landing in Downloads.

## Success criteria

Ordered. Each one is a release you could hand someone.

### 1. Daily-driver parity for a single document

- Open `.docx`, `.md`, `.xlsx`, `.csv` from the picker; edit; `Ctrl+S` saves back to the same
  file; Save as… writes a new one. **Done in the prototype.**
- Round trip fidelity identical to the desktop app — same parsers, so this is a matter of
  testing, not building. `docs/WEB-TESTS.md` §4.
- Autosave recovery survives a crash *and a reload*, which needs handle persistence.
- Print produces the same output as the desktop app.
- Clear, honest behaviour on Firefox and Safari, whatever is decided: either a read-only mode
  with an explicit banner, or a refusal that names the reason.

### 2. It remembers things

- Recent files that actually reopen: handles in IndexedDB, one permission click per file.
- Settings, zoom, theme, locale persist. (Already: the `localStorage` fallback carries them.)
- Recovery copies offered on the next visit, matched to their original file.

### 3. Multi-document

- Several documents open at once without the "new tab opens empty" wart, either by keeping
  everything in one tab with a document switcher, or by sharing handles between tabs.

### 4. The extension

- A Chrome extension that intercepts `.docx`/`.xlsx` links and downloads and renders them in the
  editor instead of writing them to disk. Shares the same bundle; only the entry differs.
  **Done as a prototype** — `npm run build:ext`, installable unpacked, `docs/WEB-EXTENSION.md`.
- This is where the web build does something the desktop app cannot, so treat it as the point of
  the project rather than a bonus.
- Still open before it could be published: narrowing `<all_urls>` to a permission asked for per
  site, and the store listing's own list.

## Non-goals

- **Explicitly excluded by the owner: double-clicking a file in Explorer/Finder.** The desktop
  app owns that. No PWA `file_handlers`, no `launchQueue`.
- **Any change to the desktop app's behaviour, size or build.** It is a shipping product on its
  own release train. The web build must remain invisible to it — see the two `grep` checks in
  `docs/WEB-TESTS.md` §1.
- No server, no accounts, no sync, no collaboration. Files stay on the user's machine.
- No new document formats. The formats the desktop app reads are the formats this reads.
- Not a mobile app. Chromium desktop is the target; a narrow screen is out of scope.

## Design constraints inherited from the desktop app

These are why people use it, and the web build does not get to trade them away.

1. **Opening is fast.** Under a second for a normal document. Watch the bundle: the web build
   currently ships a ~296 KB chunk containing ProseMirror to the Sheets editor, which does not
   use it. That is a bug, not a budget.
2. **Files come back the way they went in.** Every part of a `.docx`/`.xlsx` that the app does
   not understand is written back byte-for-byte. Nothing in the file layer may touch that.
3. **Printing matches the screen.**
4. **The keyboard comes first.** Every shortcut in the desktop app works in the browser build,
   with the browser's own conflicts (`Ctrl+P`, `Ctrl+S`, `Ctrl+W`) handled — `Ctrl+S` must save
   the document, never the page.

## Open questions for the owner

1. Firefox/Safari: read-only fallback, or refuse with an explanation?
2. Is `Ctrl+W`/tab-close worth a "you have unsaved changes" `beforeunload` prompt? The browser
   only allows a generic one, and the desktop app's Alt+F4 behaviour (write a copy, close, 5s
   cancel) is not reproducible in a tab. **Provisionally yes**: the extension build turns it on
   (`F.warnOnClose`), because a document opened from a link has no file to fall back on. It is
   one `__WEB_BUILD__` guard in each editor if that turns out to be the wrong call.
3. ~~Extension or PWA first?~~ **Answered: the extension**, and it is built. The PWA is still
   unstarted.
