# OfficeMini Sheets — plan

Goal: a fast, small spreadsheet editor for `.xlsx` and `.csv` that opens instantly, handles everyday editing and printing, keeps everything it does not understand intact, and feels exactly like the word processor.

## 1. One app, two editors

One binary, one installer, one updater, one start screen. The window loads the editor that matches the file type:

| File | Editor |
| --- | --- |
| .docx .md .txt | Words (existing) |
| .xlsx .xlsm .csv .tsv | Sheets (new) |

Shared shell (already written, to be factored out of `main.ts` into `src/shell/`): menu bar, toolbar framework, status bar, find/replace bar, dialogs, dark mode, settings, recent files, recovery copies, window state, updater, printing flow, keyboard/mouse conventions, context-menu framework. The sheet editor is loaded lazily so Words start-up time does not change.

Same look: identical menu bar (File Edit View Insert Format Data Sheet Help), same toolbar buttons where they mean the same thing (font, size, B/I/U, colours, alignment, undo/redo, find, print), same status bar (sheet tabs strip sits just above it), same dialogs and shortcuts. `.xlsx` and `.csv` get their own Explorer icons (green "S").

## 2. Rendering

- Grid drawn on a canvas with virtualisation: only visible rows/columns are painted, so 100k-cell sheets scroll at 60 fps. Fonts, fills, borders, alignment, wrap, merged cells, frozen panes, column/row headers, selection, fill handle are all painted.
- In-cell editor is an overlaid input; a formula bar sits under the toolbar with the name box (A1 jump).
- Printing renders pages as real HTML tables (vector text, exact borders), the same trick as Words: what the print preview shows is what prints. Dashed page-break lines on the grid, like Excel's page break preview.
- Dark mode reuses the Words colour mapping: light fills go dark, dark text goes light, hue kept.

## 3. Reading and writing (round-trip strategy)

Keep the zip; parse only what is rendered and edited; write back only touched sheets.

Parsed: workbook (sheets, defined names, calc settings, 1904 dates), each worksheet (cells, formulas, cached values, column widths, row heights, merges, freeze panes, hidden rows/cols, hyperlinks, autofilter range, page setup, margins, print options, header/footer, print area and titles), shared strings, styles (fonts, fills, borders, number formats, cell xfs, named styles), theme colours, tables (header styling), conditional formatting (simple rules rendered: cell value comparisons, colour scales as a later step).

Preserved untouched: charts, drawings, images, pivot tables and caches, comments, data validation, macros (`.xlsm`), external links, custom XML, everything else. Shown as a placeholder where they occupy a position on the sheet.

Written: edited worksheet XML regenerated (cells with values, formulas and cached values, styles, merges, widths, heights, panes, print settings), shared strings appended, styles appended and de-duplicated, `calcChain.xml` removed and `fullCalcOnLoad` set so Excel and LibreOffice recalculate everything on open. Unchanged sheets stay byte-for-byte.

## 4. Formulas

- Own parser and evaluator (MIT libraries considered for the function set), cross-sheet references, absolute/relative refs, ranges, arithmetic, comparison, text concatenation, percent, errors (`#DIV/0!`, `#REF!`, `#VALUE!`, `#N/A`, `#NAME?`).
- Function set for v1: SUM AVERAGE MIN MAX COUNT COUNTA COUNTBLANK SUMIF SUMIFS COUNTIF COUNTIFS AVERAGEIF ROUND ROUNDUP ROUNDDOWN INT MOD ABS SQRT POWER IF IFS AND OR NOT IFERROR ISBLANK ISNUMBER LEN LEFT RIGHT MID TRIM UPPER LOWER PROPER CONCAT CONCATENATE TEXTJOIN SUBSTITUTE FIND SEARCH TEXT VALUE VLOOKUP HLOOKUP XLOOKUP INDEX MATCH TODAY NOW DATE YEAR MONTH DAY DAYS EDATE EOMONTH WEEKDAY.
- Anything else keeps its cached value, is shown with a subtle marker, and is recalculated by Excel on open. Array/dynamic-array formulas are preserved, not evaluated.
- Recalculation on every edit with dependency tracking and cycle detection; whole-workbook recalc for small books.
- English function names; Turkish aliases (TOPLA, EĞER, DÜŞEYARA…) as a later addition. Decimal separator follows the system locale for typing and display.

## 5. Number formats and dates

Excel format codes rendered faithfully: General, fixed decimals, thousands separators, percent, currency (₺ $ € and any symbol), dates and times in all common patterns, text, scientific, colour sections and conditions in custom formats. Dates stored as serials (1900/1904 systems). Toolbar menu: General, Number, Currency, Percent, Date, Time, Text, increase/decrease decimals, thousands separator.

## 6. Editing features (v1)

Navigation: arrows, Tab/Shift+Tab, Enter/Shift+Enter, Ctrl+arrows to data edges, Home/End, Ctrl+Home/End, Page Up/Down, Shift and Ctrl+Shift extend, Ctrl+A, click and drag, row/column header click (whole row/column), corner click (all), name box, mouse wheel and Shift+wheel, Ctrl+wheel zoom.

Editing: type to replace, F2 or double-click to edit in place, formula bar editing, Escape cancels, Delete/Backspace clears, Ctrl+Z/Y undo/redo, cut/copy/paste including TSV to and from other apps and Excel, paste values only, fill handle (copy and series for numbers, dates, weekdays, months), Ctrl+D fill down and Ctrl+R fill right, insert/delete rows and columns, drag to resize rows and columns, double-click a header edge to auto-fit, merge and unmerge, wrap text, freeze panes at selection, sort a range ascending/descending, clear contents/formats/all.

Sheets: tab strip with switch, add, rename by double-click, delete, colour preserved, reorder by drag.

Formatting: font, size, bold, italic, underline, strikethrough, text colour, fill colour, borders (all, outer, inner, top, bottom, left, right, none, thick outer), horizontal and vertical alignment, indent, wrap, number formats as above, clear formatting, format painter (later).

Find and replace: the same bar as Words, searching values or formulas, whole cell, regex, within selection, replace all with count.

CSV/TSV: open with delimiter and encoding detection (comma, semicolon as used by Turkish Excel, tab; UTF-8 with or without BOM, Windows-1254), save as CSV with a delimiter choice.

Printing: page setup (paper, orientation, margins, scale or fit to N pages wide/tall, print area, repeat title rows, gridlines and headings on/off, header/footer text with page numbers), print preview lines on the grid, print the active sheet or the whole workbook.

Right-click menus on cells, row/column headers and sheet tabs. Tooltips with shortcuts everywhere; Ctrl+/ lists them.

## 7. Later (v2)

Autofilter and filter dropdowns, conditional formatting editor, data validation lists, images in cells, charts rendered from chart XML, comments shown as notes, hide/unhide rows and columns, split view, Turkish function names, format painter, multi-range selection, cell styles gallery, protection.

## 8. Performance targets

| Action | Target |
| --- | --- |
| Open a 50,000-cell workbook | under 300 ms |
| Scroll and zoom | 60 fps |
| Recalculate a typical book after an edit | under 100 ms |
| Save | under 500 ms |

## 9. Order of work

1. Shell refactor (Words keeps working), `New spreadsheet` on the start screen, file associations and icons.
2. XLSX reader and the canvas grid: open the sample workbooks with fonts, fills, borders, merges, widths, frozen panes, number formats. Compare against LibreOffice renders.
3. Editing core: cell editor, formula bar, undo, clipboard, fill handle, rows/columns, merge, freeze, sheet tabs.
4. Formatting toolbar and number formats.
5. Formula engine, recalculation, XLSX writer with round-trip validation in LibreOffice.
6. Find/replace, sort, CSV.
7. Printing and page setup.
8. Polish, shortcuts audit, dark mode check, recovery copies, release.

Roughly four working sessions; the first release comes after step 5, with printing in the second release.

## 10. Needed from you

- Five to ten real `.xlsx` files you work with (pricing lists, order sheets, reports), the same way the Word samples drove fidelity. They stay in the ignored samples folder.
- A note on which formulas you rely on most, so the function set covers them first.
