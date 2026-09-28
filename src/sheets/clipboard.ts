// Copy / cut / paste. Internal pastes carry cells (values, formulas, styles)
// and behave like Google Sheets: relative references shift with the paste
// offset, cut keeps them. External pastes use TSV (and HTML tables).
import { Workbook, Sheet, Cell, Range, key, Value, isError, MAXR, MAXC, visibleRows, visibleCols, visibleRowsFrom, visibleColsFrom } from "./model";
import { shiftFormula } from "./formula/tokens";
import { CellChange, boundRange, xfWith, inputCell, StylePatch } from "./edit";
import { formatValue, formatCodeFor, editText } from "./numfmt";
import { StyleResolver } from "./render-style";

/** r / c: position inside the (compacted) clip; sr / sc: the sheet cell it was copied from. */
export interface ClipCell { r: number; c: number; cell: Cell | null; sr: number; sc: number; }
export interface ClipData { sheet: string; range: Range; cells: ClipCell[]; cut: boolean; styles: { xf: any; font: any; fill: any; border: any; numFmt: string }[]; text?: string; }

let internal: ClipData | null = null;
let internalToken = "";

export function internalClip(): ClipData | null { return internal; }

/** Snapshot the selection into the internal clipboard and the system clipboard. */
export function copyRange(wb: Workbook, sheet: Sheet, range0: Range, cut: boolean, styles: StyleResolver, selected?: (r: number, c: number) => boolean): { text: string; html: string } {
  const range = boundRange(sheet, range0);
  // Multi-area selections copy the bounding box with cells outside the selection blanked out.
  const inSel = selected || (() => true);
  const cells: ClipCell[] = [];
  const styleIdx = new Map<number, number>();
  const styleList: ClipData["styles"] = [];
  // Hidden rows and columns (filtered out, or hidden by hand) are not copied: only the cells that
  // can be seen go to the clipboard, closed up into one block.
  const rows = visibleRows(sheet, range.r1, range.r2);
  const cols = visibleCols(sheet, range.c1, range.c2);
  const whole = rows.length === range.r2 - range.r1 + 1 && cols.length === range.c2 - range.c1 + 1;
  const rowIdx = new Map(rows.map((r, i) => [r, i]));
  const colIdx = new Map(cols.map((c, i) => [c, i]));
  for (const r of rows) for (const c of cols) {
    const cell = inSel(r, c) ? sheet.cells.get(key(r, c)) : undefined;
    if (cell && !styleIdx.has(cell.s)) {
      const xf = wb.styles.xfs[cell.s] || wb.styles.xfs[0];
      styleList.push({ xf, font: wb.styles.fonts[xf.fontId], fill: wb.styles.fills[xf.fillId], border: wb.styles.borders[xf.borderId], numFmt: formatCodeFor(xf.numFmtId, wb.styles.numFmts) });
      styleIdx.set(cell.s, styleList.length - 1);
    }
    cells.push({ r: rowIdx.get(r)!, c: colIdx.get(c)!, cell: cell ? { ...cell } : null, sr: r, sc: c });
  }
  const clipRange: Range = whole ? { ...range } : { r1: range.r1, c1: range.c1, r2: range.r1 + Math.max(1, rows.length) - 1, c2: range.c1 + Math.max(1, cols.length) - 1 };
  // A cut moves a block; with hidden cells in between there is no block to move, so it copies.
  internal = { sheet: sheet.name, range: clipRange, cells, cut: cut && whole, styles: styleList };
  internalToken = "om-sheets-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  // text/plain: TSV of displayed values; text/html: table with basic styles.
  // The table also says where it came from, in attributes: the marker that lets this window
  // recognise its own copy, and each formula with the cell it was copied from. Attributes are
  // used because they survive the clipboard - WebKit drops an HTML comment on the way - and
  // because they let another OfficeMini window, which has no access to this window's
  // clipboard, paste the formulas instead of their values.
  const lines: string[] = [];
  let html = `<!-- ${internalToken} --><table ${CLIP_ATTR}="${internalToken}">`;
  for (const r of rows) {
    const fields: string[] = [];
    html += "<tr>";
    for (const c of cols) {
      const cell = inSel(r, c) ? sheet.cells.get(key(r, c)) : undefined;
      const cs = styles.get(cell ? cell.s : 0);
      const text = cell ? styles.render(cell, cs).text : "";
      fields.push(text.replace(/\t/g, " ").replace(/\r?\n/g, " "));
      const st: string[] = [];
      if (cs.bold) st.push("font-weight:bold"); if (cs.italic) st.push("font-style:italic"); if (cs.color) st.push("color:" + cs.color); if (cs.fill) st.push("background:" + cs.fill);
      if (cs.halign && cs.halign !== "general") st.push("text-align:" + cs.halign);
      const src = cell && cell.f !== undefined ? ` ${FORMULA_ATTR}="${escapeAttr(cell.f)}" ${AT_ATTR}="${r},${c}"` : "";
      html += `<td${src}${st.length ? ` style="${st.join(";")}"` : ""}>${escapeHtml(text)}</td>`;
    }
    html += "</tr>";
    lines.push(fields.join("\t"));
  }
  html += "</table>";
  internal.text = lines.join("\n");
  return { text: internal.text, html };
}

function escapeHtml(s: string) { return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
function escapeAttr(s: string) { return escapeHtml(s).replace(/"/g, "&quot;"); }

const CLIP_ATTR = "data-om-clip", FORMULA_ATTR = "data-om-f", AT_ATTR = "data-om-at";

export function clearInternalClip() { internal = null; }
/** Esc after Ctrl+X: the data stays copied but the source is no longer deleted on paste. */
export function cancelCut() { if (internal) internal.cut = false; }

/** Is the HTML on the system clipboard the one we wrote last? */
export function isOurHtml(html: string | null): boolean {
  return !!html && !!internalToken && html.includes(internalToken);
}
/** Was the HTML copied from an OfficeMini sheet - this window or another one? */
export function isSheetHtml(html: string | null): boolean {
  return !!html && html.includes(CLIP_ATTR + "=");
}

/** A formula carried by a copied cell, with the cell it was copied from. */
export interface CopiedFormula { f: string; r: number; c: number; }

/**
 * The formulas in a table copied from an OfficeMini sheet, laid out like parseExternal()'s grid
 * (null where the cell holds a plain value). Null when the HTML carries none.
 */
export function parseFormulas(html: string | null): (CopiedFormula | null)[][] | null {
  if (!html || !html.includes(FORMULA_ATTR + "=")) return null;
  try {
    const table = new DOMParser().parseFromString(html, "text/html").querySelector("table");
    if (!table) return null;
    const out: (CopiedFormula | null)[][] = [];
    let any = false;
    for (const tr of Array.from(table.querySelectorAll("tr"))) {
      const row: (CopiedFormula | null)[] = [];
      for (const td of Array.from(tr.children)) {
        if (!/^t[dh]$/i.test(td.tagName)) continue;
        const f = td.getAttribute(FORMULA_ATTR), at = /^(\d+),(\d+)$/.exec(td.getAttribute(AT_ATTR) || "");
        if (f && at) { row.push({ f, r: parseInt(at[1], 10), c: parseInt(at[2], 10) }); any = true; } else row.push(null);
        const span = parseInt(td.getAttribute("colspan") || "1", 10) || 1;
        for (let i = 1; i < span; i++) row.push(null);
      }
      out.push(row);
    }
    return any ? out : null;
  } catch { return null; }
}

export type PasteMode = "all" | "values" | "formats" | "formulas" | "transpose";

/**
 * Compute cell changes for pasting the internal clipboard at `at` (top-left).
 * If the target selection is larger than the clip and a multiple of it, the
 * clip is tiled (like Sheets); a single-cell clip fills the whole selection.
 */
/**
 * A paste lands on visible cells only. Rows and columns that are filtered out or hidden are
 * skipped, so a column of values pasted into a filtered list fills the rows you can see instead
 * of the ones in between, and the size of the target selection is counted in visible cells.
 */
function pasteRows(sheet: Sheet, from: number, count: number): number[] { return visibleRowsFrom(sheet, from, count); }
function pasteCols(sheet: Sheet, from: number, count: number): number[] { return visibleColsFrom(sheet, from, count); }

export function pasteInternal(wb: Workbook, sheet: Sheet, target: Range, mode: PasteMode): { changes: CellChange[]; resultRange: Range; cutSource: { sheet: string; range: Range } | null } {
  const clip = internal!;
  const h = clip.range.r2 - clip.range.r1 + 1, w = clip.range.c2 - clip.range.c1 + 1;
  // the target's size is what can be seen of it
  const th = Math.max(1, visibleRows(sheet, target.r1, target.r2).length), tw = Math.max(1, visibleCols(sheet, target.c1, target.c2).length);
  const transpose = mode === "transpose";
  const ch = transpose ? w : h, cw = transpose ? h : w;
  const repR = th >= ch && th % ch === 0 && (th > ch || tw > cw) && !clip.cut ? th / ch : 1;
  const repC = tw >= cw && tw % cw === 0 && (th > ch || tw > cw) && !clip.cut ? tw / cw : 1;
  const changes: CellChange[] = [];
  const rows = pasteRows(sheet, target.r1, repR * ch);
  const cols = pasteCols(sheet, target.c1, repC * cw);
  const styleMap = new Map<number, number>();
  const mapStyle = (cell: Cell): number => {
    // Same workbook: style indices are valid as-is (clipboard survives sheet switches).
    return cell.s;
  };
  const srcSheetIsThis = clip.sheet.toLowerCase() === sheet.name.toLowerCase();
  for (let rr = 0; rr < repR; rr++) for (let cc = 0; cc < repC; cc++) {
    for (const cc0 of clip.cells) {
      const pr = transpose ? cc0.c : cc0.r, pc = transpose ? cc0.r : cc0.c;
      const r = rows[rr * ch + pr], c = cols[cc * cw + pc];
      if (r === undefined || c === undefined || r >= MAXR || c >= MAXC) continue;
      const existing = sheet.cells.get(key(r, c));
      const src = cc0.cell;
      if (!src) { if (mode !== "formats" && existing) changes.push({ r, c, cell: existing.s && mode === "values" ? { v: null, s: existing.s } : null }); continue; }
      // relative references shift by the distance from the cell it was really copied from
      const dr = r - cc0.sr, dc = c - cc0.sc;
      let cell: Cell;
      if (mode === "values") cell = { v: src.v, s: existing ? existing.s : 0 };
      else if (mode === "formats") { if (!existing) { changes.push({ r, c, cell: { v: null, s: mapStyle(src) } }); continue; } cell = { ...existing, s: mapStyle(src) }; }
      else if (mode === "formulas") cell = { v: src.f ? null : src.v, f: src.f ? (clip.cut ? src.f : shiftFormula(src.f, dr, dc)) : undefined, s: existing ? existing.s : 0 };
      else { cell = { ...src, s: mapStyle(src) }; if (src.f) { cell.f = clip.cut ? src.f : shiftFormula(src.f, dr, dc); if (!clip.cut) cell.v = null; } }
      if (cell.f === undefined) delete cell.f;
      delete cell.sh; delete cell.arr;
      changes.push({ r, c, cell });
    }
  }
  const resultRange: Range = { r1: rows[0] ?? target.r1, c1: cols[0] ?? target.c1, r2: rows[rows.length - 1] ?? target.r1, c2: cols[cols.length - 1] ?? target.c1 };
  void styleMap; void srcSheetIsThis;
  return { changes, resultRange, cutSource: clip.cut ? { sheet: clip.sheet, range: clip.range } : null };
}

/** Parse external clipboard content (HTML table preferred, else TSV) into a value grid. */
export function parseExternal(html: string | null, text: string): string[][] {
  if (html && /<table/i.test(html)) {
    try {
      const doc = new DOMParser().parseFromString(html, "text/html");
      const table = doc.querySelector("table");
      if (table) {
        const rows: string[][] = [];
        for (const tr of Array.from(table.querySelectorAll("tr"))) {
          const cells: string[] = [];
          for (const td of Array.from(tr.children)) {
            if (!/^t[dh]$/i.test(td.tagName)) continue;
            const span = parseInt(td.getAttribute("colspan") || "1", 10) || 1;
            cells.push((td as HTMLElement).innerText.replace(/ /g, " ").trim());
            for (let i = 1; i < span; i++) cells.push("");
          }
          rows.push(cells);
        }
        if (rows.length) return rows;
      }
    } catch { /* fall through */ }
  }
  const t = text.replace(/\r\n?/g, "\n");
  const lines = t.endsWith("\n") ? t.slice(0, -1).split("\n") : t.split("\n");
  // Quoted TSV fields with embedded newlines
  const rows: string[][] = [];
  let cur: string[] | null = null, field = "", q = false;
  for (const line of lines) {
    if (!cur) cur = [];
    let i = 0;
    while (i < line.length) {
      const ch = line[i];
      if (q) { if (ch === '"') { if (line[i + 1] === '"') { field += '"'; i += 2; continue; } q = false; i++; continue; } field += ch; i++; continue; }
      if (ch === '"' && field === "") { q = true; i++; continue; }
      if (ch === "\t") { cur.push(field); field = ""; i++; continue; }
      field += ch; i++;
    }
    if (q) { field += "\n"; continue; }
    cur.push(field); field = "";
    rows.push(cur); cur = null;
  }
  if (cur) { cur.push(field); rows.push(cur); }
  return rows;
}

/**
 * Paste a grid of text that came from outside this window. `formulas`, when given, holds the
 * formulas of cells copied from another OfficeMini window: those cells get the formula, with
 * its relative references moved by the distance from the cell it was copied from - the same
 * result as a paste inside one window.
 */
export function pasteExternal(wb: Workbook, sheet: Sheet, at: { r: number; c: number }, grid: string[][], target: Range | null, formulas?: (CopiedFormula | null)[][] | null): { changes: CellChange[]; resultRange: Range } {
  const changes: CellChange[] = [];
  let maxC = 0;
  const single = grid.length === 1 && grid[0].length === 1;
  const reps = single && target ? target : null;
  const put = (r: number, c: number, text: string, i: number, j: number) => {
    if (r >= MAXR || c >= MAXC) return;
    const from = formulas?.[i]?.[j];
    const cell = inputCell(wb, sheet, r, c, from ? "=" + shiftFormula(from.f, r - from.r, c - from.c) : text);
    changes.push({ r, c, cell });
  };
  if (reps) {
    for (const r of visibleRows(sheet, reps.r1, reps.r2)) for (const c of visibleCols(sheet, reps.c1, reps.c2)) put(r, c, grid[0][0], 0, 0);
    return { changes, resultRange: reps };
  }
  // Same rule as an internal paste: hidden rows and columns are skipped, not written through.
  for (const row of grid) maxC = Math.max(maxC, row.length);
  const rows = pasteRows(sheet, at.r, grid.length);
  const cols = pasteCols(sheet, at.c, maxC);
  grid.forEach((row, i) => { row.forEach((text, j) => { if (rows[i] !== undefined && cols[j] !== undefined) put(rows[i], cols[j], text, i, j); }); });
  return { changes, resultRange: { r1: rows[0] ?? at.r, c1: cols[0] ?? at.c, r2: rows[rows.length - 1] ?? at.r, c2: cols[cols.length - 1] ?? at.c } };
}

export { formatValue, editText, xfWith, isError };
export type { StylePatch, Value };
