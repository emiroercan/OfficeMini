// Editing operations on the workbook model with undo/redo. Every operation
// returns an Entry that can be reverted; structural operations snapshot the
// sheet they touch.
import {
  Workbook, Sheet, Cell, Range, Font, Fill, Border, BorderSide, Xf, Styles, ColInfo, RowInfo, Value,
  key, rowOf, colOf, updateExtent, recomputeExtent, MAXR, MAXC, inRange, cellRef, parseRef, DEFAULT_FONT, Hyperlink,
} from "./model";
import { adjustFormulaForInsertDelete, shiftFormula, renameSheetInFormula } from "./formula/tokens";
import { parseInput, isDateFormat, locale, serialToDate, dateToSerial } from "./numfmt";
import { formatCodeFor } from "./numfmt";

export interface SelSnapshot { sheet: number; ranges: Range[]; active: { r: number; c: number }; }
export interface Entry { label: string; undo(): void; redo(): void; before?: SelSnapshot; after?: SelSnapshot; }

export class History {
  private entries: Entry[] = [];
  private index = 0;
  onChange: (() => void) | null = null;
  push(e: Entry) { this.entries.length = this.index; this.entries.push(e); this.index++; if (this.entries.length > 500) { this.entries.shift(); this.index--; } this.onChange?.(); }
  canUndo() { return this.index > 0; }
  canRedo() { return this.index < this.entries.length; }
  undo(): Entry | null { if (!this.canUndo()) return null; const e = this.entries[--this.index]; e.undo(); this.onChange?.(); return e; }
  redo(): Entry | null { if (!this.canRedo()) return null; const e = this.entries[this.index++]; e.redo(); this.onChange?.(); return e; }
  clear() { this.entries = []; this.index = 0; this.onChange?.(); }
  undoLabel() { return this.canUndo() ? this.entries[this.index - 1].label : ""; }
  redoLabel() { return this.canRedo() ? this.entries[this.index].label : ""; }
}

export type CellChange = { r: number; c: number; cell: Cell | null };

/** Apply cell replacements; returns the inverse changes. */
export function applyCellChanges(sheet: Sheet, changes: CellChange[]): CellChange[] {
  const inverse: CellChange[] = [];
  let needExtent = false;
  for (const ch of changes) {
    const k = key(ch.r, ch.c);
    const old = sheet.cells.get(k) || null;
    inverse.push({ r: ch.r, c: ch.c, cell: old });
    sheet.dirtyRows.add(ch.r);
    if (sheet.formulaKeys) { if (ch.cell && ch.cell.f !== undefined) sheet.formulaKeys.add(k); else sheet.formulaKeys.delete(k); }
    if (ch.cell) { sheet.cells.set(k, ch.cell); updateExtent(sheet, ch.r, ch.c); }
    else { sheet.cells.delete(k); if (ch.r === sheet.maxRow || ch.c === sheet.maxCol) needExtent = true; }
  }
  if (needExtent) recomputeExtent(sheet);
  sheet.dirty = true;
  return inverse;
}

// ---- styles ------------------------------------------------------------------

export interface StylePatch {
  font?: { name?: string; size?: number; bold?: boolean; italic?: boolean; strike?: boolean; underline?: string | null; color?: string | null };
  fill?: string | null;                     // RRGGBB or null to clear
  border?: { top?: BorderSide | null; bottom?: BorderSide | null; left?: BorderSide | null; right?: BorderSide | null };
  numFmt?: string;
  halign?: string | null; valign?: string | null; wrap?: boolean; indent?: number; rotation?: number;
  reset?: boolean;                          // clear all formatting
}

const fontSig = (f: Font) => JSON.stringify([f.name, f.size, f.bold, f.italic, f.underline, f.strike, f.color, f.scheme]);
const fillSig = (f: Fill) => JSON.stringify([f.patternType, f.fg, f.bg]);
const borderSig = (b: Border) => JSON.stringify([b.left, b.right, b.top, b.bottom, b.diagonal, b.diagonalUp, b.diagonalDown]);
const xfSig = (x: Xf) => JSON.stringify([x.numFmtId, x.fontId, x.fillId, x.borderId, x.xfId, x.alignment, x.protection, x.quotePrefix]);

function findOrAdd<T>(list: T[], item: T, sig: (t: T) => string, start = 0): number {
  const s = sig(item);
  for (let i = start; i < list.length; i++) if (sig(list[i]) === s) return i;
  list.push(item);
  return list.length - 1;
}

export function numFmtIdFor(styles: Styles, code: string): number {
  for (const [id, c] of Object.entries({ General: 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10, "0.00E+00": 11, "m/d/yyyy": 14, "d-mmm-yy": 15, "d-mmm": 16, "mmm-yy": 17, "h:mm AM/PM": 18, "h:mm:ss AM/PM": 19, "h:mm": 20, "h:mm:ss": 21, "m/d/yyyy h:mm": 22, "mm:ss": 45, "[h]:mm:ss": 46, "@": 49 })) if (id === code) return c;
  for (const [id, c] of styles.numFmts) if (c === code) return id;
  let id = 164;
  for (const k of styles.numFmts.keys()) id = Math.max(id, k + 1);
  styles.numFmts.set(id, code);
  return id;
}

/** Style index derived from `base` with the patch applied (de-duplicated). */
export function xfWith(styles: Styles, base: number, patch: StylePatch): number {
  const bx = styles.xfs[base] || styles.xfs[0];
  if (patch.reset) return 0;
  const xf: Xf = { ...bx, alignment: bx.alignment ? { ...bx.alignment } : null, protection: bx.protection ? { ...bx.protection } : null };
  if (patch.font) {
    const bf = styles.fonts[bx.fontId] || styles.fonts[0];
    const f: Font = { ...bf, raw: undefined };
    const p = patch.font;
    if (p.name !== undefined) { f.name = p.name; delete f.scheme; }
    if (p.size !== undefined) f.size = p.size;
    if (p.bold !== undefined) f.bold = p.bold;
    if (p.italic !== undefined) f.italic = p.italic;
    if (p.strike !== undefined) f.strike = p.strike;
    if (p.underline !== undefined) f.underline = p.underline;
    if (p.color !== undefined) f.color = p.color ? { rgb: "FF" + p.color.replace("#", "").toUpperCase() } : null;
    xf.fontId = findOrAdd(styles.fonts, f, fontSig);
    xf.applyFont = true;
  }
  if (patch.fill !== undefined) {
    const fill: Fill = patch.fill ? { patternType: "solid", fg: { rgb: "FF" + patch.fill.replace("#", "").toUpperCase() }, bg: { indexed: 64 } } : { patternType: "none", fg: null, bg: null };
    xf.fillId = findOrAdd(styles.fills, fill, fillSig, 2);
    if (!patch.fill) xf.fillId = 0;
    xf.applyFill = true;
  }
  if (patch.border) {
    const bb = styles.borders[bx.borderId] || styles.borders[0];
    const b: Border = { left: bb.left, right: bb.right, top: bb.top, bottom: bb.bottom, diagonal: bb.diagonal, diagonalUp: bb.diagonalUp, diagonalDown: bb.diagonalDown };
    for (const side of ["top", "bottom", "left", "right"] as const) if (patch.border[side] !== undefined) b[side] = patch.border[side]!;
    xf.borderId = findOrAdd(styles.borders, b, borderSig);
    xf.applyBorder = true;
  }
  if (patch.numFmt !== undefined) { xf.numFmtId = numFmtIdFor(styles, patch.numFmt); xf.applyNumberFormat = true; }
  if (patch.halign !== undefined || patch.valign !== undefined || patch.wrap !== undefined || patch.indent !== undefined || patch.rotation !== undefined) {
    const al = { ...(xf.alignment || {}) };
    if (patch.halign !== undefined) { if (patch.halign) al.horizontal = patch.halign; else delete al.horizontal; }
    if (patch.valign !== undefined) { if (patch.valign) al.vertical = patch.valign; else delete al.vertical; }
    if (patch.wrap !== undefined) { if (patch.wrap) al.wrapText = true; else delete al.wrapText; }
    if (patch.indent !== undefined) { if (patch.indent > 0) al.indent = patch.indent; else delete al.indent; }
    if (patch.rotation !== undefined) { if (patch.rotation) al.textRotation = patch.rotation; else delete al.textRotation; }
    xf.alignment = Object.keys(al).length ? al : null;
    xf.applyAlignment = true;
  }
  return findOrAdd(styles.xfs, xf, xfSig);
}

// ---- entries -------------------------------------------------------------------

export function cellsEntry(sheet: Sheet, changes: CellChange[], label: string): Entry {
  let inverse = applyCellChanges(sheet, changes);
  const fwd = changes.map((c) => ({ ...c }));
  return {
    label,
    undo() { applyCellChanges(sheet, inverse); },
    redo() { inverse = applyCellChanges(sheet, fwd); },
  };
}

/** Bound whole-row/column selections to the used area plus a margin. */
export function boundRange(sheet: Sheet, rg: Range, extra = 0): Range {
  return { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, Math.max(sheet.maxRow + extra, rg.r1)), c2: Math.min(rg.c2, Math.max(sheet.maxCol + extra, rg.c1)) };
}

export function styleEntry(wb: Workbook, sheet: Sheet, ranges: Range[], patch: StylePatch, label = "Format"): Entry {
  const changes: CellChange[] = [];
  const colStyles: { c: number; before: ColInfo[]; }[] = [];
  const rowStyles: { r: number; before: RowInfo | undefined }[] = [];
  const xfCache = new Map<number, number>();
  const map = (base: number) => { let v = xfCache.get(base); if (v === undefined) { v = xfWith(wb.styles, base, patch); xfCache.set(base, v); } return v; };
  const seen = new Set<number>();
  for (const rg0 of ranges) {
    const wholeCols = rg0.r1 === 0 && rg0.r2 >= MAXR - 1;
    const wholeRows = rg0.c1 === 0 && rg0.c2 >= MAXC - 1;
    const rg = boundRange(sheet, rg0);
    // existing cells inside the (bounded) range
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
      const k = key(r, c);
      if (seen.has(k)) continue; seen.add(k);
      const cell = sheet.cells.get(k);
      if (cell) changes.push({ r, c, cell: { ...cell, s: map(cell.s) } });
      else if (!wholeCols && !wholeRows) { const rowStyle = sheet.rows.get(r)?.style || 0; changes.push({ r, c, cell: { v: null, s: map(rowStyle) } }); }
    }
    if (wholeCols) {
      if (!colStyles.length) colStyles.push({ c: rg0.c1, before: sheet.cols.map((x) => ({ ...x })) });
      setColStyleRange(sheet, rg0.c1, Math.min(rg0.c2, MAXC - 1), map);
    }
    // row styles only matter for cells typed later: stop a little past the used area (Ctrl+Shift+Down selects a million rows)
    if (wholeRows && !wholeCols) for (let r = rg0.r1; r <= Math.min(rg0.r2, MAXR - 1, Math.max(sheet.maxRow, 0) + 500); r++) {
      const before = sheet.rows.get(r);
      rowStyles.push({ r, before: before ? { ...before } : undefined });
      sheet.rows.set(r, { ...(before || {}), style: map(before?.style || 0) });
    }
  }
  const inverse = applyCellChanges(sheet, changes);
  const fwd = changes.map((c) => ({ ...c }));
  const colsAfter = sheet.cols.map((x) => ({ ...x }));
  const rowsAfter = rowStyles.map((rs) => ({ r: rs.r, after: { ...(sheet.rows.get(rs.r) || {}) } }));
  return {
    label,
    undo() {
      applyCellChanges(sheet, inverse);
      if (colStyles.length) sheet.cols = colStyles[0].before;
      for (const rs of rowStyles) { if (rs.before) sheet.rows.set(rs.r, rs.before); else sheet.rows.delete(rs.r); }
      sheet.dirty = true;
    },
    redo() {
      applyCellChanges(sheet, fwd);
      if (colStyles.length) sheet.cols = colsAfter.map((x) => ({ ...x }));
      for (const rs of rowsAfter) sheet.rows.set(rs.r, { ...rs.after });
      sheet.dirty = true;
    },
  };
}

/** Apply a style mapping to every column in [c1, c2], splitting existing <col> ranges only at the two boundaries. */
function setColStyleRange(sheet: Sheet, c1: number, c2: number, map: (base: number) => number) {
  const out: ColInfo[] = [];
  const covered: [number, number][] = [];
  for (const ci of sheet.cols) {
    if (ci.max < c1 || ci.min > c2) { out.push(ci); continue; }
    if (ci.min < c1) out.push({ ...ci, max: c1 - 1 });
    const lo = Math.max(ci.min, c1), hi = Math.min(ci.max, c2);
    out.push({ ...ci, min: lo, max: hi, style: map(ci.style || 0) });
    covered.push([lo, hi]);
    if (ci.max > c2) out.push({ ...ci, min: c2 + 1 });
  }
  covered.sort((a, b) => a[0] - b[0]);
  let cur = c1;
  const gapStyle = map(0);
  for (const [lo, hi] of covered) { if (lo > cur) out.push({ min: cur, max: lo - 1, style: gapStyle, width: sheet.defaultColWidth }); cur = Math.max(cur, hi + 1); }
  if (cur <= c2) out.push({ min: cur, max: c2, style: gapStyle, width: sheet.defaultColWidth });
  out.sort((a, b) => a.min - b.min);
  sheet.cols = out;
}

function setColStyle(sheet: Sheet, c: number, style: number) {
  const out: ColInfo[] = [];
  let found = false;
  for (const ci of sheet.cols) {
    if (c < ci.min || c > ci.max) { out.push(ci); continue; }
    if (ci.min < c) out.push({ ...ci, max: c - 1 });
    out.push({ ...ci, min: c, max: c, style });
    if (ci.max > c) out.push({ ...ci, min: c + 1 });
    found = true;
  }
  if (!found) out.push({ min: c, max: c, style, width: sheet.defaultColWidth });
  out.sort((a, b) => a.min - b.min);
  sheet.cols = out;
}

/** Set a value typed by the user; picks up formats for dates/percent/currency like Excel. */
export function inputCell(wb: Workbook, sheet: Sheet, r: number, c: number, text: string): Cell | null {
  const old = sheet.cells.get(key(r, c));
  const base = old ? old.s : (sheet.rows.get(r)?.style ?? sheet.cols.find((x) => c >= x.min && c <= x.max)?.style ?? 0);
  if (text === "") return old && old.s ? { v: null, s: old.s } : null;
  if (text.startsWith("=") && text.length > 1) return { v: old?.f === text.slice(1) ? old.v : null, f: text.slice(1), s: base };
  const currentFmt = formatCodeFor((wb.styles.xfs[base] || wb.styles.xfs[0]).numFmtId, wb.styles.numFmts);
  if (currentFmt === "@" || text.startsWith("'")) return { v: text.startsWith("'") ? text.slice(1) : text, s: base };
  const p = parseInput(text, locale(), wb.date1904);
  let s = base;
  if (p.format && typeof p.value === "number") {
    // keep an existing numeric/date format when it already fits; otherwise adopt the detected one
    const keep = currentFmt !== "General" && (isDateFormat(currentFmt) === isDateFormat(p.format));
    if (!keep) s = xfWith(wb.styles, base, { numFmt: p.format });
  }
  return { v: p.value, s };
}

// ---- structural ------------------------------------------------------------------

interface SheetSnapshot { cells: Map<number, Cell>; rows: Map<number, RowInfo>; cols: ColInfo[]; merges: Range[]; autoFilter: Range | null; hyperlinks: Hyperlink[]; freeze: Sheet["freeze"]; maxRow: number; maxCol: number; }
function snapshot(sheet: Sheet): SheetSnapshot {
  return { cells: new Map(sheet.cells), rows: new Map(sheet.rows), cols: sheet.cols.map((c) => ({ ...c })), merges: sheet.merges.map((m) => ({ ...m })), autoFilter: sheet.autoFilter ? { ...sheet.autoFilter } : null, hyperlinks: sheet.hyperlinks.map((h) => ({ ...h })), freeze: sheet.freeze ? { ...sheet.freeze } : null, maxRow: sheet.maxRow, maxCol: sheet.maxCol };
}
function restore(sheet: Sheet, s: SheetSnapshot) {
  sheet.formulaKeys = undefined;
  sheet.cells = new Map(s.cells); sheet.rows = new Map(s.rows); sheet.cols = s.cols.map((c) => ({ ...c })); sheet.merges = s.merges.map((m) => ({ ...m }));
  sheet.autoFilter = s.autoFilter ? { ...s.autoFilter } : null; sheet.hyperlinks = s.hyperlinks.map((h) => ({ ...h })); sheet.freeze = s.freeze ? { ...s.freeze } : null;
  sheet.maxRow = s.maxRow; sheet.maxCol = s.maxCol; sheet.dirty = true;
}

/** Wrap a mutation of one sheet (plus formula fixes on other sheets) into an entry using snapshots. */
export function structuralEntry(wb: Workbook, sheetIdx: number, label: string, mutate: () => void, othersToo = false): Entry {
  const sheet = wb.sheets[sheetIdx];
  const before = snapshot(sheet);
  const othersBefore = wb.sheets.map((s) => (s === sheet || !othersToo ? null : new Map(s.cells)));
  mutate();
  const after = snapshot(sheet);
  const othersAfter = wb.sheets.map((s) => (s === sheet || !othersToo ? null : new Map(s.cells)));
  return {
    label,
    undo() { restore(sheet, before); wb.sheets.forEach((s, i) => { if (othersBefore[i]) { s.cells = new Map(othersBefore[i]!); s.dirty = true; } }); },
    redo() { restore(sheet, after); wb.sheets.forEach((s, i) => { if (othersAfter[i]) { s.cells = new Map(othersAfter[i]!); s.dirty = true; } }); },
  };
}

function shiftRange(rg: Range, axis: "row" | "col", at: number, count: number): Range | null {
  const a = axis === "row" ? rg.r1 : rg.c1, b = axis === "row" ? rg.r2 : rg.c2;
  let na: number, nb: number;
  if (count > 0) { na = a >= at ? a + count : a; nb = b >= at ? b + count : b; }
  else {
    const end = at - count;
    if (a >= at && b < end) return null;
    na = a >= end ? a + count : a >= at ? at : a;
    nb = b >= end ? b + count : b >= at ? at - 1 : b;
    if (nb < na) return null;
  }
  return axis === "row" ? { ...rg, r1: na, r2: nb } : { ...rg, c1: na, c2: nb };
}

/** Insert (count>0) or delete (count<0) rows/columns at index `at`. */
export function insertDelete(wb: Workbook, sheetIdx: number, axis: "row" | "col", at: number, count: number) {
  const sheet = wb.sheets[sheetIdx];
  const cells = new Map<number, Cell>();
  for (const [k, cell] of sheet.cells) {
    const r = rowOf(k), c = colOf(k);
    const v = axis === "row" ? r : c;
    let nv: number | null = v;
    if (count > 0) { if (v >= at) nv = v + count; }
    else { const end = at - count; if (v >= at && v < end) nv = null; else if (v >= end) nv = v + count; }
    if (nv === null) continue;
    let ncell = cell;
    if (cell.f) { const f2 = adjustFormulaForInsertDelete(cell.f, sheet.name, sheet.name, axis, at, count); if (f2 !== cell.f) ncell = { ...cell, f: f2 }; }
    if (ncell.arr) { const ar = parseRef(ncell.arr); const s2 = ar ? shiftRange({ r1: ar.r1, c1: ar.c1, r2: ar.r2, c2: ar.c2 }, axis, at, count) : null; ncell = { ...ncell }; if (s2) ncell.arr = cellRef(s2.r1, s2.c1) + ":" + cellRef(s2.r2, s2.c2); else delete ncell.arr; }
    cells.set(axis === "row" ? key(nv, c) : key(r, nv), ncell);
  }
  sheet.cells = cells;
  sheet.formulaKeys = undefined;
  if (axis === "row") {
    const rows = new Map<number, RowInfo>();
    for (const [r, info] of sheet.rows) {
      if (count > 0) rows.set(r >= at ? r + count : r, info);
      else { const end = at - count; if (r >= at && r < end) continue; rows.set(r >= end ? r + count : r, info); }
    }
    sheet.rows = rows;
    // inserted rows inherit nothing; deleted filter hides
    sheet.hiddenRowsByFilter = new Set();
  } else {
    const cols: ColInfo[] = [];
    for (const ci of sheet.cols) {
      const s = shiftRange({ r1: 0, r2: 0, c1: ci.min, c2: ci.max }, "col", at, count);
      if (!s) continue;
      // splitting a range that straddles the insertion point
      if (count > 0 && ci.min < at && ci.max >= at) { cols.push({ ...ci, max: at - 1 }); cols.push({ ...ci, min: at + count, max: ci.max + count }); }
      else cols.push({ ...ci, min: s.c1, max: s.c2 });
    }
    sheet.cols = cols.sort((a, b) => a.min - b.min);
  }
  sheet.merges = sheet.merges.map((m) => shiftRange(m, axis, at, count)).filter((m): m is Range => !!m && (m.r1 !== m.r2 || m.c1 !== m.c2));
  if (sheet.autoFilter) sheet.autoFilter = shiftRange(sheet.autoFilter, axis, at, count);
  sheet.hyperlinks = sheet.hyperlinks.map((h) => { const p = parseRef(h.ref); if (!p) return h; const s = shiftRange({ r1: p.r1, c1: p.c1, r2: p.r2, c2: p.c2 }, axis, at, count); return s ? { ...h, ref: cellRef(s.r1, s.c1) + (s.r1 !== s.r2 || s.c1 !== s.c2 ? ":" + cellRef(s.r2, s.c2) : "") } : null; }).filter((h): h is Hyperlink => !!h);
  if (sheet.freeze) { const f = sheet.freeze; if (axis === "row" && at < f.rows) f.rows = Math.max(0, f.rows + (count > 0 ? count : Math.max(count, at - f.rows))); if (axis === "col" && at < f.cols) f.cols = Math.max(0, f.cols + (count > 0 ? count : Math.max(count, at - f.cols))); }
  recomputeExtent(sheet);
  sheet.dirty = true;
  // formulas on other sheets referring to this sheet
  for (const other of wb.sheets) {
    if (other === sheet) continue;
    let touched = false;
    for (const [k, cell] of other.cells) {
      if (!cell.f) continue;
      const f2 = adjustFormulaForInsertDelete(cell.f, other.name, sheet.name, axis, at, count);
      if (f2 !== cell.f) { cell.f = f2; touched = true; }
      void k;
    }
    if (touched) other.dirty = true;
  }
}

export function renameSheet(wb: Workbook, idx: number, name: string) {
  const old = wb.sheets[idx].name;
  wb.sheets[idx].name = name;
  wb.sheets[idx].dirty = true;
  for (const s of wb.sheets) {
    let touched = false;
    for (const cell of s.cells.values()) { if (!cell.f) continue; const f2 = renameSheetInFormula(cell.f, old, name); if (f2 !== cell.f) { cell.f = f2; touched = true; } }
    if (touched) s.dirty = true;
  }
  for (const d of wb.definedNames) d.ref = renameSheetInFormula(d.ref, old, name);
}

// ---- fill series ---------------------------------------------------------------------

/** Values for filling `dst` from `src` (Google Sheets rules: copy, linear series, dates, "Item 1", weekday/month names, formulas shift). */
export function fillChanges(wb: Workbook, sheet: Sheet, src: Range, dst: Range): CellChange[] {
  const changes: CellChange[] = [];
  const vertical = dst.r2 - dst.r1 > src.r2 - src.r1 || dst.r1 !== src.r1;
  const horizontal = dst.c2 - dst.c1 > src.c2 - src.c1 || dst.c1 !== src.c1;
  const axis: "row" | "col" = vertical && !horizontal ? "row" : horizontal && !vertical ? "col" : (dst.r2 - dst.r1) >= (dst.c2 - dst.c1) ? "row" : "col";
  const down = axis === "row" ? dst.r2 > src.r2 : dst.c2 > src.c2;
  const lanes = axis === "row" ? src.c2 - src.c1 + 1 : src.r2 - src.r1 + 1;
  const srcLen = axis === "row" ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
  for (let lane = 0; lane < lanes; lane++) {
    const srcCells: (Cell | undefined)[] = [];
    for (let i = 0; i < srcLen; i++) {
      const r = axis === "row" ? src.r1 + i : src.r1 + lane, c = axis === "row" ? src.c1 + lane : src.c1 + i;
      srcCells.push(sheet.cells.get(key(r, c)));
    }
    const series = detectSeries(wb, srcCells);
    const targets: { r: number; c: number; idx: number }[] = [];
    if (down) {
      const start = axis === "row" ? src.r2 + 1 : src.c2 + 1, end = axis === "row" ? dst.r2 : dst.c2;
      for (let p = start, i = 0; p <= end; p++, i++) targets.push({ r: axis === "row" ? p : src.r1 + lane, c: axis === "row" ? src.c1 + lane : p, idx: srcLen + i });
    } else {
      const start = axis === "row" ? src.r1 - 1 : src.c1 - 1, end = axis === "row" ? dst.r1 : dst.c1;
      for (let p = start, i = -1; p >= end; p--, i--) targets.push({ r: axis === "row" ? p : src.r1 + lane, c: axis === "row" ? src.c1 + lane : p, idx: i });
    }
    for (const t of targets) {
      const si = ((t.idx % srcLen) + srcLen) % srcLen;
      const proto = srcCells[si];
      const cell = series.next(t.idx, si, proto);
      if (cell && proto?.f) {
        const sr = axis === "row" ? src.r1 + si : src.r1 + lane, sc = axis === "row" ? src.c1 + lane : src.c1 + si;
        cell.f = shiftFormula(proto.f, t.r - sr, t.c - sc);
        cell.v = null;
      }
      changes.push({ r: t.r, c: t.c, cell });
    }
  }
  return changes;
}

interface Series { next(idx: number, srcIdx: number, proto: Cell | undefined): Cell | null; }

function detectSeries(wb: Workbook, src: (Cell | undefined)[]): Series {
  const copy: Series = { next: (_i, si, proto) => (proto ? { ...proto } : null) };
  const vals = src.map((c) => c?.v ?? null);
  const nums = vals.map((v) => (typeof v === "number" ? v : null));
  const allNum = nums.every((n) => n !== null) && nums.length > 0;
  const anyFormula = src.some((c) => c?.f);
  if (anyFormula) return copy;
  if (allNum) {
    const isDate = src.every((c) => c && isDateFormat(formatCodeFor((wb.styles.xfs[c.s] || wb.styles.xfs[0]).numFmtId, wb.styles.numFmts)));
    if (nums.length === 1) {
      if (isDate) return { next: (i, si, proto) => ({ ...proto!, v: (nums[0] as number) + (i - si), f: undefined }) };
      return copy;
    }
    // linear series (least squares like Excel's fill)
    const n = nums.length;
    const xs = nums.map((_, i) => i), ys = nums as number[];
    const mx = xs.reduce((a, b) => a + b, 0) / n, my = ys.reduce((a, b) => a + b, 0) / n;
    let num = 0, den = 0;
    for (let i = 0; i < n; i++) { num += (xs[i] - mx) * (ys[i] - my); den += (xs[i] - mx) ** 2; }
    const slope = den ? num / den : 0, intercept = my - slope * mx;
    return { next: (i, _si, proto) => ({ ...(proto || src[0]!), v: round(intercept + slope * i), f: undefined }) };
  }
  const strs = vals.map((v) => (typeof v === "string" ? v : null));
  if (strs.every((s) => s !== null)) {
    // month / weekday names
    const loc = locale();
    const lists = [loc.months, loc.monthsShort, loc.days, loc.daysShort];
    for (const list of lists) {
      const idxs = strs.map((s) => list.findIndex((x) => x.toLowerCase() === (s as string).toLowerCase()));
      if (idxs.every((i) => i >= 0)) {
        const step = idxs.length > 1 ? idxs[1] - idxs[0] || 1 : 1;
        const cap = strs[0] === (strs[0] as string).toUpperCase() ? "upper" : null;
        return { next: (i, _si, proto) => { let name = list[(((idxs[0] + step * i) % list.length) + list.length) % list.length]; if (cap === "upper") name = name.toUpperCase(); return { ...(proto || src[0]!), v: name }; } };
      }
    }
    // trailing number: "Item 1" -> "Item 2"
    const m = strs.map((s) => /^(.*?)(\d+)(\D*)$/.exec(s as string));
    if (m.every((x) => x) && m.length) {
      const base = m[0]![1], suffix = m[0]![3];
      const ns = m.map((x) => parseInt(x![2], 10));
      const step = ns.length > 1 ? ns[1] - ns[0] : 1;
      const width = m[0]![2].length;
      return { next: (i, _si, proto) => { const v = ns[0] + step * i; return { ...(proto || src[0]!), v: base + String(Math.max(0, v)).padStart(m[0]![2].startsWith("0") ? width : 0, "0") + suffix }; } };
    }
  }
  return copy;
}
function round(v: number): number { return Math.round(v * 1e10) / 1e10; }

// ---- sorting -----------------------------------------------------------------------

/** A cell relocated by (dr, dc): relative references shift like a copy, shared/array formula info is dropped. */
export function movedCell(cell: Cell, dr: number, dc = 0): Cell {
  const out: Cell = { ...cell };
  delete out.sh; delete out.arr;
  if (out.f !== undefined && (dr || dc)) { out.f = shiftFormula(out.f, dr, dc); out.v = null; }
  return out;
}

export function sortChanges(sheet: Sheet, range: Range, col: number, ascending: boolean, hasHeader: boolean): CellChange[] {
  const r1 = range.r1 + (hasHeader ? 1 : 0);
  const rows: { r: number; cells: Map<number, Cell | undefined> }[] = [];
  for (let r = r1; r <= range.r2; r++) {
    const m = new Map<number, Cell | undefined>();
    for (let c = range.c1; c <= range.c2; c++) m.set(c, sheet.cells.get(key(r, c)));
    rows.push({ r, cells: m });
  }
  const collator = new Intl.Collator(locale().id === "tr" ? "tr" : "en", { numeric: true, sensitivity: "base" });
  const keyOf = (row: typeof rows[number]) => row.cells.get(col)?.v ?? null;
  const rank = (v: Value) => (v === null || v === "" ? 3 : typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2);
  const sorted = [...rows].sort((a, b) => {
    const va = keyOf(a), vb = keyOf(b);
    const ra = rank(va), rb = rank(vb);
    if (ra !== rb) return ra - rb; // blanks last regardless of direction
    let cmp = 0;
    if (typeof va === "number" && typeof vb === "number") cmp = va - vb;
    else if (typeof va === "string" && typeof vb === "string") cmp = collator.compare(va, vb);
    else if (typeof va === "boolean" && typeof vb === "boolean") cmp = (va ? 1 : 0) - (vb ? 1 : 0);
    return ascending || ra === 3 ? cmp : -cmp;
  });
  const changes: CellChange[] = [];
  sorted.forEach((row, i) => {
    const target = r1 + i;
    if (row.r === target) return;
    for (let c = range.c1; c <= range.c2; c++) { const src = row.cells.get(c); changes.push({ r: target, c, cell: src ? movedCell(src, target - row.r) : null }); }
  });
  return changes;
}

// ---- clear -----------------------------------------------------------------------------

export function clearChanges(sheet: Sheet, ranges: Range[], what: "contents" | "formats" | "all"): CellChange[] {
  const changes: CellChange[] = [];
  const seen = new Set<number>();
  for (const rg0 of ranges) {
    const rg = boundRange(sheet, rg0);
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
      if (sheet.hiddenRowsByFilter.size && sheet.hiddenRowsByFilter.has(r)) continue;
      const k = key(r, c);
      if (seen.has(k)) continue; seen.add(k);
      const cell = sheet.cells.get(k);
      if (!cell) continue;
      if (what === "all") changes.push({ r, c, cell: null });
      else if (what === "contents") changes.push({ r, c, cell: cell.s ? { v: null, s: cell.s } : null });
      else changes.push({ r, c, cell: cell.v === null && !cell.f ? null : { v: cell.v, f: cell.f, s: 0 } });
    }
  }
  return changes;
}

export { DEFAULT_FONT, serialToDate, dateToSerial, inRange };
