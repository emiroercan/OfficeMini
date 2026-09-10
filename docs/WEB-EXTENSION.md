# OfficeMini for Chrome — the extension

State as of 2026-09-10, on top of v0.3.3. This is success criterion 4 of `docs/WEB-GOALS.md`:
the part of the browser build that the desktop app cannot do.

## What it does

Click a `.docx` or `.xlsx` link and it opens in the editor, in a tab, instead of landing in the
Downloads folder. The file is never written to disk unless you ask for it — `Ctrl+S` raises
Save as…, and from that point on it saves back to the file you chose, in place, like the
desktop app.

Three ways in:

| | |
|---|---|
| A download Chrome was about to start | cancelled and opened in the editor |
| Right-click a link → *Open link in OfficeMini* | opened without downloading |
| The toolbar button | new document, new spreadsheet, or the picker |

Nothing is uploaded. The document is fetched by the extension from the site you were already
on, and stays in the browser.

## Install it for yourself

```bash
npm run build:ext
```

Then in Chrome: `chrome://extensions` → turn on **Developer mode** → **Load unpacked** →
select the `dist-ext` folder. It appears as *OfficeMini*; pin it to the toolbar if you want the
popup one click away.

To try it immediately, open any page with a `.docx` link and click it — for example a raw file
link from a GitHub repository. The download bubble may flash for an instant before the tab
opens; that is the cancellation, and no file is left behind.

After a code change, `npm run build:ext` again and press the reload arrow on the extension's
card. A change to `background.js` also needs that reload; a change under `src/` only needs the
tab reopened.

`npm run dev:ext` serves the same build on `localhost:1420` for quick iteration on the editor
itself, but the extension plumbing (downloads, context menu, popup) only exists in Chrome.

## How the hand-off works

The awkward part is that a download URL is often single-use — Drive, SharePoint and most
`?export=` endpoints will not serve the same URL twice. So the bytes are fetched **once**, by
the service worker, at the moment the download is cancelled and while the site's cookies still
apply. They are parked in the origin private file system under `inbox/<token>`, and the editor
tab — same `chrome-extension://` origin, so the same OPFS — picks them up by token.

```
link click
   │
   ├─ chrome.downloads.onCreated              extension/background.js
   │     cancel + erase              nothing reaches the Downloads folder
   │     fetch(url, credentials)     one request, while it is still valid
   │     OPFS  inbox/<token>
   │
   └─ tabs.create  index.html?inbox=<token>&name=Report.docx&src=<url>
         │
         ├─ adoptEntryUrl()                   src/files-web.ts
         │     reads inbox/<token>, registers "/net/1/Report.docx",
         │     rewrites the query to ?file=/net/1/Report.docx
         │
         └─ boot.ts → Words or Sheets, unchanged
```

`src=` is the fallback: if the service worker's fetch failed, the page tries the URL itself.
A context-menu click passes only `src=`.

The inbox copy is **kept** after the editor reads it, so reloading the tab reopens the same
document rather than re-requesting a spent URL. `sweepInbox()` deletes copies older than a day
on the next browser start.

### `/net/` paths

A document from a link has no file behind it, but the rest of the app only ever deals in path
strings. So it gets one — `/net/1/Report.docx` — alongside the `/web/…` (granted handle) and
`/opfs/…` (recovery copy) paths that `files-web.ts` already invented. `basename`, `extname` and
`dirname` behave; `readFile` serves the bytes; `writeFile` refuses with a message.

Three places know that a path can be remote, all of them folded away in the desktop build:

- `save()` in both editors routes to `saveAs()` — there is nowhere to write back to.
- `addRecent()` skips it — a Recent entry that cannot reopen is worse than none.
- `warnOnClose()` puts the browser's generic unsaved-changes prompt on the tab.

## What is shared, and what is not

The editor is the ordinary `--mode web` build. `--mode ext` differs only in `base` (relative,
because an extension page is not served from a root), `outDir` (`dist-ext`), and a Vite plugin
that copies `extension/` and the app's icons in after the bundle. `__WEB_BUILD__` is true in
both, so there is no third code path to keep working.

```
extension/manifest.json    MV3
extension/background.js    the service worker: interception and the hand-off
extension/settings.js      which file types to take over (shared with the popup)
extension/popup.js/.html   the toolbar popup
```

## Settings

The popup has a master switch and four groups: Word, Excel, CSV/TSV, and Markdown/text.
Markdown and text are off by default — the browser shows those links rather than downloading
them, so the setting would mostly catch files someone deliberately asked to keep. Settings live
in `chrome.storage.sync`, so they follow the Chrome profile.

## Known gaps

1. **`blob:` downloads are left alone.** A page that builds the file in JavaScript and saves it
   from a blob URL cannot be intercepted: the URL belongs to that page and a service worker
   cannot fetch it. Those still download normally. A content script could relay the bytes; that
   is the only reason this extension would ever need one.
2. **Two requests to the server.** The download is cancelled and then re-fetched. Sites that
   count downloads will count one; sites that invalidate the URL on first *response* rather
   than first request would break, and none tested do.
3. **A brief flash in the download bubble** before the cancel lands. Cosmetic.
4. **One document per tab.** Inherited from the web build — gap 3 in `docs/WEB-HANDOFF.md`.
5. **Markdown side files.** Saving a `.md` with images still cannot write `name_files/` next to
   the document — gap 2 in `docs/WEB-HANDOFF.md`.
6. **Chromium only**, and deliberately so: this is a Chrome extension.

## Before the Chrome Web Store

The prototype is installable and complete; publishing it is a separate list.

- **`host_permissions: ["<all_urls>"]` is the hard part.** It shows as *"Read and change all
  your data on all websites"* and reviewers ask for a justification. It is genuinely needed —
  the file could be on any site — but the polite version is `optional_host_permissions`,
  requested per-site from the popup with `chrome.permissions.request()` on a click. Worth doing
  before submitting; it turns the install-time warning into a per-site one.
- **Permission justifications** to write for the listing: `downloads` (to cancel the download
  being replaced), `contextMenus` (the right-click item), `storage` (which file types to take
  over), host access (to fetch the document the link points at).
- **Data disclosure**: nothing is collected or transmitted. The listing has to say so
  explicitly, and a privacy policy URL is required once any permission is declared.
- **No remote code.** Everything ships in the package; the bundle has no `eval` or `new
  Function` and loads nothing from the network. Checked — this is what makes MV3's default
  content security policy work with no override in the manifest.
- **Single purpose**: "open Office documents from the web in an editor" is one purpose, which
  is what the policy asks for.
- **Assets**: a 128px icon (present), at least one 1280×800 screenshot, a short description.
- `npm run pack:ext` produces `officemini-extension-<version>.zip` with the manifest at the
  root, which is the shape the store's uploader wants. Bump `version` in `package.json` for
  every upload — the manifest takes its version from there at build time.

## Rules this must keep

From `docs/WEB-HANDOFF.md`, unchanged and non-negotiable: **the desktop app's build must not
change.** `docs/WEB-TESTS.md` §1 is the check, and it now compares the desktop bundle against
`HEAD` byte-for-byte rather than only grepping it.
