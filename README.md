# OfficeMini

A small, fast word processor for `.docx` and Markdown files, plus a spreadsheet editor for `.xlsx` and CSV files. Built with Tauri 2 (Rust shell, system webview) and ProseMirror. Opens in well under a second, keeps every part of a Word or Excel file it does not understand exactly as it was, and prints what you see on screen.

## Sheets (.xlsx / .csv)

One binary, two editors: opening a spreadsheet starts the Sheets editor, which shares the menus, toolbar, dark mode, start screen, recovery copies and updater with the word processor.

- Opens `.xlsx`, `.xlsm`, `.csv`, `.tsv` (UTF-8, UTF-16 and Windows-1254 Turkish CSVs are detected, so are `;` and tab delimiters); saves `.xlsx` and `.csv`. A 45 000-cell workbook opens in about 100 ms, a 9 MB CSV with 800 000 cells in about 1.5 s.
- Canvas grid with frozen panes; the header row is frozen by default when a file has no frozen rows (view-only, nothing is written unless you freeze it yourself). Freeze any number of rows and columns from the View menu, the ❄ toolbar button or the row/column header menus.
- Google Sheets controls: Enter or F2 edits, typing replaces, Enter/Tab commit and move (Enter returns to the column where a Tab run started), Shift+Enter moves up, Esc cancels, Alt+Enter inserts a line break, Delete clears, arrow keys and Ctrl+arrows jump across data, Ctrl+A selects the data block then the sheet, Ctrl+Space / Shift+Space select columns / rows, double-click a column edge to fit it.
- Formulas: `=` starts a formula; while typing, clicking or arrowing over cells inserts references, references are colour-coded on the grid, F4 cycles `$` anchors, function names autocomplete with signatures, and the status bar shows the result live. About 130 functions with Excel semantics: SUM/AVERAGE/COUNT/MIN/MAX, SUBTOTAL (respects filtered rows), IF/IFS/IFERROR/SWITCH/AND/OR, SUMIF(S)/COUNTIF(S)/AVERAGEIF, XLOOKUP/VLOOKUP/HLOOKUP/INDEX/MATCH/OFFSET/INDIRECT, CONCAT/TEXTJOIN/LEFT/RIGHT/MID/LEN/TRIM/SUBSTITUTE/TEXT/VALUE/SPLIT, ROUND family, DATE/TODAY/NOW/EDATE/EOMONTH/DATEDIF/NETWORKDAYS/WORKDAY (with holidays), UNIQUE, SUMPRODUCT and array arithmetic such as `SUMPRODUCT((A:A="x")*(B:B))`, LET, SEQUENCE, plus Turkish names (TOPLA, EĞER, DÜŞEYARA, ÇAPRAZARA…). Both `,` and `;` argument separators work; whole-row (`2:2`) and whole-column (`B:B`) references and unquoted sheet names (`Data!A1`) are understood. Unknown functions keep the value cached in the file, and Excel recalculates on open.
- Copy and paste behave like Google Sheets: relative references shift on copy and stay on cut, pasting into a larger selection tiles the clip, `Ctrl+Shift+V` pastes values only, the Edit menu has formats-only, formulas-only and transposed paste. TSV/HTML from other programs pastes with type detection; copies go out as tab-separated text and an HTML table. Drag the selection border to move cells (Ctrl to copy), drag the fill handle for series (numbers, dates, month and day names, "Item 1", formulas).
- Formatting: number format presets (numbers, percent, currency ₺ $ € £, dates in Turkish/European/US/ISO order, time, duration, accounting) and custom Excel format codes with live preview, decimal +/- buttons, font, size, bold/italic/underline/strikethrough, text and fill colour, borders, alignment, wrap, indent, merge. Ctrl+Shift+1..6 apply number, time, date, currency, percent and scientific formats like Excel.
- Number locale switch in the status bar (TR / US / EU / UK): decides how typed numbers and dates are read (`1.234,56` vs `1,234.56`, `09.09.2026` vs `9/9/2026`), how `General` numbers are shown, the currency button and month/day names. Typing `15%`, `₺1.234,56`, `12,5`, `9 Eylül 2026` or `14:30` picks a matching format automatically.
- Data: sort A→Z / Z→A / two-level sort dialog with header detection, Excel-style filters (funnel button in the header row, searchable unique-value list with counts, blanks), remove duplicates, trim whitespace, pivot tables (rows, second row field, columns, two values with SUM/COUNT/COUNTA/AVERAGE/MIN/MAX, filter, totals; written to a new sheet and refreshable), insert/delete/hide rows and columns with formula references adjusted across sheets, column width and row height dialogs, autofit.
- Sheet tabs: add, rename (double-click), duplicate, delete, reorder by dragging, tab colour, hide/unhide, Ctrl+PageUp/PageDown to switch.
- Find and replace across the sheet or the workbook, with match case, entire cell, regex, search in formulas, Turkish-aware case folding.
- Printing: whole sheet or selection, portrait/landscape, A4/Letter, fit to width, gridlines, row/column headers, frozen rows repeated on every page.
- Round trip: only the sheets you changed are regenerated; charts, drawings, images, comments, conditional formatting, data validation, pivot caches, tables, print settings and every other part are written back byte-for-byte. Shared strings and styles are extended, never rewritten from scratch.
- `.xlsx` and `.csv` files get their own icons in Explorer after installation; Ctrl+/ lists every shortcut.

## What it does

- Opens `.docx`, `.md`, `.txt`; saves `.docx` and `.md` (either can be saved as the other).
- Page view (default) or continuous view, formatting marks. Zoom: Ctrl+wheel, touchpad pinch, presets for 100%, page width, whole page and *Actual size* (paper at its physical dimensions, computed from the monitor's reported size; also the first-run default). 100% follows the Word convention of 96 pixels per inch, so it is not physical size on high-density screens. View menu also sets the scroll speed (default 2× on Linux, where touchpads map finger travel 1:1).
- Text formatting: font, size, bold/italic/underline/strikethrough, super/subscript, colour, highlight, caps.
- Paragraphs: styles (Normal, Heading 1‑6, Title, Quote…), alignment, indents, line and paragraph spacing, bullets and numbering, page breaks, tab stops.
- Tables: insert, add/delete rows and columns, merge/split, shading, borders, column resize by dragging, Tab between cells.
- Images: insert, paste, drag onto the window, resize by dragging corners, wrap left/right/inline. SVG, PNG, JPEG, GIF, and EMF/WMF with embedded bitmaps.
- Find and replace: match case, whole word, regular expressions with `$1` groups, search inside a selection, case-preserving replacement, Turkish-aware case folding (`i`, `İ`, `ı`, `I` match), wrap-around, live match count. Enter/Shift+Enter step, Ctrl+Enter replaces all, Alt+C/W/R/S/P toggle the options.
- Dark mode (`Ctrl+Shift+D`): the interface and the page go dark; dark text becomes light and light fills become dark while hues are kept, transparent images get a light plate so logos stay visible, photos are left alone. Printing always uses the light theme.
- Markdown files can be edited rendered or as raw source (`Ctrl+Alt+M`, or the "Markdown source" button in the status bar); images referenced by relative paths are loaded from the file's folder.
- Large embedded images are inflated after the text is on screen, so a 10 MB document still opens in about 100 ms.
- Start screen with recent files and recovered documents when launched without a file; window size and position are remembered.
- A recovery copy of unsaved work is written every minute (Format menu to switch off); after a crash the start screen offers it, and opening a file with a newer recovery copy asks which one to use.
- Smart quotes and dashes while typing (Format menu to switch off); the status bar shows the word count of the selection.
- Undo/redo, hyperlinks, headers and footers with page numbers, page setup, printing.
- Right-click context menus for text, links, images and tables; tooltips show every shortcut; `Ctrl+/` opens the full shortcut list.
- `.docx` and `.md` files get their own icons in Explorer after installation.

## Round-trip strategy

When a `.docx` is saved, only the paragraphs and runs you touched are rewritten, and even those keep every property the editor does not model. Styles, numbering definitions, headers, footers, footnotes, comments, content controls, drawings, fields and settings are written back byte-for-byte. Text boxes are editable; other shapes are shown as placeholders and preserved.

## Keyboard shortcuts (excerpt)

| Keys | Action |
| --- | --- |
| Ctrl+N / Ctrl+O / Ctrl+S / Ctrl+Shift+S | New / Open / Save / Save as |
| Ctrl+P | Print |
| Ctrl+Z / Ctrl+Y or Ctrl+Shift+Z | Undo / Redo |
| Ctrl+F / Ctrl+H / F3 / Shift+F3 | Find / Replace / Next / Previous |
| Ctrl+B / Ctrl+I / Ctrl+U / Ctrl+Shift+X | Bold / Italic / Underline / Strikethrough |
| Ctrl+= / Ctrl+Shift+= | Subscript / Superscript |
| Ctrl+Shift+> / Ctrl+Shift+< / Ctrl+] / Ctrl+[ | Grow / shrink font |
| Ctrl+L / Ctrl+E / Ctrl+R / Ctrl+J | Align left / centre / right / justify |
| Ctrl+Shift+L / Ctrl+Shift+7 | Bullets / Numbering |
| Ctrl+M / Ctrl+Shift+M, Tab / Shift+Tab in lists | Indent / Outdent |
| Ctrl+1 / Ctrl+5 / Ctrl+2 | Single / 1.5 / double line spacing |
| Ctrl+Alt+1..3, Ctrl+Shift+N | Heading 1..3, Normal |
| Ctrl+Space | Clear character formatting |
| Ctrl+K / Ctrl+Shift+I / Ctrl+Shift+T | Insert link / image / table |
| Ctrl+Enter / Shift+Enter | Page break / line break |
| Ctrl+Shift+Space / Ctrl+Shift+- | Non-breaking space / hyphen |
| Ctrl++ / Ctrl+- / Ctrl+0 / Ctrl+wheel | Zoom |
| Ctrl+Shift+8 | Formatting marks |
| Ctrl+Shift+G | Go to page |
| F11 | Full screen |
| Ctrl+/ or F1 | All shortcuts and mouse actions |
| Alt+F, Alt+E, Alt+V, Alt+I, Alt+O, Alt+T, Alt+H | Open menus |

Mouse: double-click selects a word, triple-click a paragraph, Shift+click extends, Ctrl+click opens links, right-click opens the context menu, dragging an image corner resizes it, dragging a cell border resizes the column, middle-click pastes the primary selection on Linux.

## Building

Requirements: Node.js 18+, Rust (stable), and the platform webview toolchain.

```bash
npm install
npm run tauri dev      # run in development
npm run tauri build    # installers in src-tauri/target/release/bundle
```

### Windows

Needs the Visual Studio Build Tools with the "Desktop development with C++" workload (MSVC + Windows SDK). WebView2 is part of Windows 11. The build produces an NSIS installer (`.exe`) and an MSI.

### macOS

Needs Xcode command line tools (`xcode-select --install`). `npm run tauri build` produces an `OfficeMini.app` and a `.dmg` under `src-tauri/target/release/bundle` for the Mac it runs on; `npm run tauri build -- --target universal-apple-darwin` (after `rustup target add aarch64-apple-darwin x86_64-apple-darwin`) makes one universal build for Apple silicon and Intel, which is what the release workflow ships.

The releases are only ad-hoc signed (no Apple Developer certificate), so Gatekeeper refuses the downloaded app at first: after copying it to Applications, either right-click the app and choose *Open*, or allow it under System Settings > Privacy & Security ("Open Anyway"), or clear the quarantine flag with `xattr -dr com.apple.quarantine /Applications/OfficeMini.app`. This is needed once; updates installed by the app itself are not quarantined. The app installs a native menu bar with Edit and Window menus (WebKit only accepts Cmd+C/V/X through them); the document menus stay in the window. Trackpad pinch zooms, Cmd replaces Ctrl in every shortcut, and files opened from Finder or with *Open with* land in the front window (or a new one when it holds unsaved work).

### Fedora Linux

```bash
sudo dnf install webkit2gtk4.1-devel openssl-devel dbus-devel curl wget file libappindicator-gtk3-devel librsvg2-devel gcc gcc-c++ make rpm-build patchelf
sudo dnf install google-carlito-fonts google-caladea-fonts liberation-fonts   # metric-compatible Calibri/Cambria/Arial replacements
npm install
NO_STRIP=true npm run tauri build
```

`NO_STRIP=true` is needed on current Fedora: the `strip` bundled inside linuxdeploy predates the `.relr.dyn` relocation sections that Fedora's libraries use, and the AppImage step fails without it. The RPM and DEB are unaffected. The "public key found, but no private key" message at the end only concerns the updater signature and can be ignored for local builds.

The build produces an `.rpm`, a `.deb` and an `AppImage` under `src-tauri/target/release/bundle`. Install the RPM with `sudo dnf install ./OfficeMini-*.rpm`; it registers `.docx`, `.md`, `.xlsx` and `.csv` file associations.

If scrolling and typing feel like they run at half speed on Linux, check the power profile: when the system reports the *power-saver* profile (power-profiles-daemon or tuned-ppd, also on AC power), WebKitGTK caps every page at 30 frames per second and coarsens timers, and the app has no way to opt out. Switch to *Balanced* in the KDE/GNOME power menu, or run `powerprofilesctl set balanced` (or `tuned-adm profile balanced`), and the editor runs at 60 fps again.

The app runs on the session's native backend. Wayland is preferred: it delivers per-pixel touchpad scrolling and pinch-to-zoom, which XWayland cannot. GTK3 renders at 2x on fractional-scale Wayland outputs, so the window is created "app paintable" to skip GTK's full-window background fill on every redraw (that fill alone cost ~2.5 s of main-thread CPU at startup on a 2560x1600 laptop). Set `GDK_BACKEND=x11` to force XWayland if needed.

## Releases and automatic updates

Installed copies check GitHub Releases quietly a few seconds after start (at most once every six hours) and offer to download and install a newer version; Help > Check for updates… does it on demand. Updates are signed: the app only installs bundles whose signature matches the public key in `src-tauri/tauri.conf.json`.

To publish a release, bump the version in `package.json`, `src-tauri/tauri.conf.json` and `src-tauri/Cargo.toml`, commit, then tag and push:

```bash
git tag v0.3.0 && git push origin main --tags
```

The `Release` workflow builds Windows, Linux and macOS installers, signs them with the repository secrets `TAURI_SIGNING_PRIVATE_KEY` / `TAURI_SIGNING_PRIVATE_KEY_PASSWORD`, creates the GitHub release and uploads `latest.json` for the updater. The private key lives outside the repository (`~/.tauri/officemini.key` on the build machine); losing it means existing installs can no longer verify new updates, so keep a backup.

## Browser development mode

`npm run dev` and open `http://localhost:1420/?file=/samples/<name>.docx` (or `.xlsx`/`.csv` for the Sheets editor, `?file=new:xlsx` for a blank workbook). In this mode "Save" posts the file to `samples/out/` through a dev-only endpoint, which is handy for round-trip testing with LibreOffice (`soffice --headless --convert-to pdf`). `window.om` exposes the running editor for scripting.

## Known limitations

- Screen pagination is computed from the rendered flow; it matches Word closely for ordinary documents but not for complex float-heavy layouts.
- Shapes other than pictures and text boxes are shown as placeholders (they are preserved in the file).
- Footnotes, comments and tracked changes are preserved but not shown or editable.
- Multi-section documents use the last section's page setup for all pages.
- Fonts not installed on the machine fall back to metric-compatible or generic families.
