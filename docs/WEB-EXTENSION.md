# OfficeMini for Chrome — the extension

State as of 2026-09-11, on top of v0.3.4. This is success criterion 4 of `docs/WEB-GOALS.md`:
the part of the browser build that the desktop app cannot do.

## What it does

Download a Word or Excel file and it opens in the editor, in a tab — whether it came from a plain
link, an export button, or a spreadsheet the site built on the page. Once the editor has it, the
copy in Downloads is removed, unless *Keep a copy in Downloads too* is on. `Ctrl+S` raises
Save as…, and from then on saves back to the file you chose, in place, like the desktop app.

| | |
|---|---|
| A download of a `.docx`, `.xlsx`, `.csv`… | opened in the editor once Chrome has it |
| A file from your computer dropped onto any tab | opened as a copy; `Ctrl+S` saves somewhere new |
| A file dropped onto an OfficeMini tab | opened, and `Ctrl+S` saves back to it |
| Right-click a link → *Open link in OfficeMini* | opened without downloading |
| The toolbar button | new document, new spreadsheet, or the picker |

Nothing is uploaded. Files stay on the machine.

## Install it for yourself

```bash
npm run build:ext
```

(In Windows PowerShell with script execution disabled, `npm.cmd run build:ext`.)

Then in Chrome: `chrome://extensions` → turn on **Developer mode** → **Load unpacked** → select
the `dist-ext` folder. Then open the extension's **Details** and turn on **Allow access to file
URLs**. That is the one switch it needs for export buttons, blob downloads and local files;
plain links work without it. While it is off the popup says so, and the first file that needs it
opens a page explaining the switch — the file then opens by itself once it is on.

After a code change, build again and press the reload arrow on the extension's card.

## How it works

### Downloads

Chrome does the download exactly as it would without the extension. When it completes, the
service worker reads the file Chrome wrote — `fetch()` on its `file://` URL, which a worker may
do once file access is on — parks the bytes in the origin private file system under
`inbox/<token>`, and opens the editor. The editor tab shares the `chrome-extension://` origin,
so it picks them up by token.

```
download completes                      chrome.downloads.onChanged    extension/background.js
   │   (its real name is known by now, whatever the URL said)
   ├─ fetch(file:///…/Downloads/Export.xlsx)        one read of what Chrome saved
   ├─ OPFS  inbox/<token>
   ├─ tabs.create  index.html?inbox=<token>&name=Export.xlsx&src=<url>
   │     └─ adoptEntryUrl() → "/net/1/Export.xlsx" → Words or Sheets, unchanged
   └─ removeFile + erase                            unless "keep a copy" is on
```

**Why not intercept sooner.** The first version cancelled the download and fetched the URL again
itself. That handled a plain link and nothing else: an export endpoint names its file only in
`Content-Disposition`, unknown until the download is under way; a `blob:` URL belongs to the
page that made it; and many export URLs work exactly once. Admin panels use all three. Reading
what Chrome saved is the one approach that works for every case, and it never makes a second
request.

The copy in Downloads is removed only after the bytes are safely in the inbox. The inbox copy is
kept, so reloading the tab reopens the document; `sweepInbox()` deletes copies older than a day.

### Local files

A file dropped onto a tab that does not handle drops makes Chrome navigate to its `file://` URL.
What happens next depends on whether anything claims the type:

- **Nothing does** (a clean profile): the navigation becomes a download — a second copy in
  Downloads. `downloads.onCreated` cancels the copy (the original is already on disk, so nothing
  is lost), reads the original, and opens it the same way.
- **Another extension does**: Chrome *shows* the file in the tab and never downloads it. The
  real case is Google's old **Office Editing for Docs, Sheets & Slides**
  (`gbkeegbaiigmenfmjfclcdgdpimamgkj`, internally `qo_documents`), still installed in many
  profiles. It registers a MIME handler for Word, Excel and CSV and renders them in its own viewer
  — or, as in the owner's profile, fails and leaves *"Couldn't load plugin"* on a `file://` page.
  Either way `tabs.onUpdated` catches that tab once it has loaded and replaces it with the editor,
  in the same tab.

A drop can reach both listeners, so the first to *claim* the file's URL (in
`chrome.storage.session`, ten seconds) opens it and the other stands down. Chrome only tells an
extension a `file://` tab's URL when file access is on — which reading the file needs anyway.

### Without file access

- **A plain link** is fetched again from its URL. The bytes are checked first — a zip format has
  to start with `PK`, a CSV must not be an HTML error page — before anything is opened or removed.
- **Anything else** — an export endpoint that will not serve twice, a blob, a local file — opens
  `file-access.html`, which explains the switch and links to the extension's settings. The file
  is remembered in `chrome.storage.local` for fifteen minutes and opened as soon as the switch is
  on: the worker checks on every start, the explainer asks when it sees the switch flip, and so
  does the popup when it is opened. A shared in-flight promise stops two of those opening it twice.

### Dropped files (onto an OfficeMini tab)

Chromium's `DataTransferItem.getAsFileSystemHandle()` returns a real `FileSystemFileHandle`, and
the drop is itself the user gesture that a write-permission prompt needs — so a dropped document
opens *and* saves back to where it came from, exactly like one chosen from the picker. Browsers
without it fall back to the plain `File`, which opens read-only as a `/net/` path.

Preventing the browser's default matters as much as the feature: without it, dropping a file on
the editor navigates the tab to that file and throws the open document away.

### `/net/` paths

A document from a download has no file behind it that the editor may write to, but the rest of
the app only ever deals in path strings. So it gets one — `/net/1/Report.docx` — alongside the
`/web/…` (granted handle) and `/opfs/…` (recovery copy) paths that `files-web.ts` already
invented. `basename`, `extname` and `dirname` behave; `readFile` serves the bytes; `writeFile`
refuses with a message.

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
extension/manifest.json          MV3
extension/background.js          the service worker: downloads, local files, the hand-off
extension/settings.js            which file types to take over (shared with the popup)
extension/popup.js/.html         the toolbar popup
extension/file-access.js/.html   the page that explains "Allow access to file URLs"
```

## Settings

The popup has a master switch, four groups — Word, Excel, CSV/TSV, Markdown/text — and *Keep a
copy in Downloads too*, which is off by default: the point is "open it instead of downloading
it". Markdown and text are off by default too — the browser shows those links rather than
downloading them, so the setting would mostly catch files someone deliberately asked to keep.
Settings live in `chrome.storage.sync`, so they follow the Chrome profile.

## Known gaps

1. **File access is a manual switch.** Chrome has no API to request it. `file-access.html` walks
   through it, and plain links work without it. A local file *shown* in its tab (the plugin case
   above) is not even noticed while the switch is off: Chrome hides a `file://` tab's URL from
   extensions without it, so that tab just keeps its error page.
2. **A file that fails to open has already left Downloads** when *keep a copy* is off: the copy is
   removed once the bytes are parked, not once the editor has parsed them. Turn the extension's
   switch off and download it again, or keep copies.
3. **"Ask where to save each file"**: with that Chrome setting on and *keep a copy* off, a file
   saved to a folder you chose is removed once it opens. Turn *keep a copy* on if you use it.
4. **The download shows in the download bubble** until it is removed. Cosmetic.
5. **One document per tab.** Inherited from the web build — gap 3 in `docs/WEB-HANDOFF.md`.
6. **Markdown side files.** Saving a `.md` with images still cannot write `name_files/` next to
   the document — gap 2 in `docs/WEB-HANDOFF.md`.
7. **Chromium only**, and deliberately so: this is a Chrome extension.

## Testing

`npm run test:ext` builds the extension and runs `scripts/e2e-ext.mjs`: the real extension in real Chrome, headless, driven over the DevTools protocol, with generated fixtures and real mouse
clicks. It covers every download shape above, local files both downloaded and shown in their tab,
*keep a copy* and the master switch. The file-access-off path - the explainer, and the file
opening once the switch goes on - is checked by hand: Chrome disables a pipe-loaded extension
when that switch flips.
`docs/WEB-TESTS.md` §7 has the details and the reference result.

## Before the Chrome Web Store

The prototype is installable and complete; publishing it is a separate list.

- **Host permission.** `<all_urls>` shows as *"Read and change all your data on all websites"*.
  Since the redesign it only serves the no-file-access fallback and the context menu — the main
  path reads a local file. Worth checking whether `file:///*` plus `optional_host_permissions`,
  requested per site from the popup with `chrome.permissions.request()`, is enough; that would
  take the install-time warning away.
- **File access** cannot be requested by the extension. The listing should say that it asks for
  *Allow access to file URLs*, and why.
- **Permission justifications** to write for the listing: `downloads` (read a finished download,
  remove the copy, cancel the duplicate of a local file), `contextMenus` (the right-click item),
  `storage` (settings, and a file waiting on the switch), host access (the fallback fetch).
- **Data disclosure**: nothing is collected or transmitted. The listing has to say so
  explicitly, and a privacy policy URL is required once any permission is declared.
- **No remote code.** Everything ships in the package; the bundle has no `eval` or `new
  Function` and loads nothing from the network. Checked — this is what makes MV3's default
  content security policy work with no override in the manifest.
- **Single purpose**: "open Office documents you download in an editor" is one purpose, which is
  what the policy asks for.
- **Assets**: a 128px icon (present), at least one 1280×800 screenshot, a short description.
- `npm run pack:ext` produces `officemini-extension-<version>.zip` with the manifest at the
  root, which is the shape the store's uploader wants. Bump `version` in `package.json` for
  every upload — the manifest takes its version from there at build time.

## Rules this must keep

From `docs/WEB-HANDOFF.md`, unchanged and non-negotiable: **the desktop app's build must not
change.** `docs/WEB-TESTS.md` §1 is the check, and it compares the desktop bundle against `HEAD`
byte-for-byte rather than only grepping it.
