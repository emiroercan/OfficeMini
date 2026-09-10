// Workbook model for the Sheets editor. Cells are stored sparsely per sheet in a
// Map keyed by row*MAXC+col (0-based); everything the editor does not model is
// kept as raw XML blocks so the file round-trips.
import type { Package } from "../docx/zip";

export const MAXC = 16384;      // Excel column limit (XFD)
export const MAXR = 1048576;    // Excel row limit

export interface CellError { e: string; }
export type Value = number | string | boolean | CellError | null;

export interface Cell {
  v: Value;                 // value or, for formulas, the cached/computed result
  f?: string;               // formula text without the leading "="
  s: number;                // style index into Styles.xfs (0 = default)
  sh?: { si: number; ref?: string; master?: boolean }; // shared formula group (only from files)
  arr?: string;             // array formula range (preserved, not evaluated)
  rich?: string;            // raw <is>/<si> rich-text XML when the string had formatting runs
}

export interface RowInfo { height?: number; hidden?: boolean; style?: number; customHeight?: boolean; level?: number; }
export interface ColInfo { min: number; max: number; width?: number; hidden?: boolean; style?: number; customWidth?: boolean; bestFit?: boolean; level?: number; }

export interface Range { r1: number; c1: number; r2: number; c2: number; }

export interface Hyperlink { ref: string; rId?: string; target?: string; location?: string; display?: string; tooltip?: string; }

export interface RawBlock { name: string; xml: string; }

export interface SheetView {
  showGridLines: boolean;
  zoom: number;              // percent
  tabSelected: boolean;
  topLeft?: string;          // topLeftCell
  active?: string;           // active cell
  selection?: string;        // sqref
  rightToLeft?: boolean;
}

export interface Sheet {
  name: string;
  sheetId: number;
  rId: string;
  part: string;              // "xl/worksheets/sheet1.xml"
  state: "visible" | "hidden" | "veryHidden";
  cells: Map<number, Cell>;
  rows: Map<number, RowInfo>;
  cols: ColInfo[];
  merges: Range[];
  freeze: { rows: number; cols: number } | null;
  autoFreeze: boolean;       // freeze added by us for viewing only (not saved)
  view: SheetView;
  defaultRowHeight: number;  // points
  defaultColWidth: number;   // character units
  autoFilter: Range | null;
  hyperlinks: Hyperlink[];
  tabColor: string | null;
  maxRow: number;            // 0-based, -1 when empty
  maxCol: number;
  rawBlocks: RawBlock[];     // unmodelled worksheet children, in document order
  rawRootAttrs: string;      // original <worksheet ...> attributes
  relsXml: string | null;    // original sheet rels part text
  dirty: boolean;            // needs regeneration on save
  hiddenRowsByFilter: Set<number>; // rows hidden by the in-app filter (not persisted as hidden)
  dirtyRows: Set<number>;    // rows whose cells changed since the grid last measured automatic heights
  formulaKeys?: Set<number>; // keys of formula cells (maintained by applyCellChanges; undefined = rebuild on next recalc)
  tabColorChanged?: boolean; // the user picked a tab colour (otherwise the original <tabColor>, theme or indexed, is kept)
}

// ---- styles -----------------------------------------------------------------

export interface Color { rgb?: string; theme?: number; tint?: number; indexed?: number; auto?: boolean; }
export interface Font { name: string; size: number; bold: boolean; italic: boolean; underline: string | null; strike: boolean; color: Color | null; family?: number; scheme?: string; charset?: number; raw?: string; }
export interface Fill { patternType: string; fg: Color | null; bg: Color | null; raw?: string; }
export interface BorderSide { style: string; color: Color | null; }
export interface Border { left: BorderSide | null; right: BorderSide | null; top: BorderSide | null; bottom: BorderSide | null; diagonal: BorderSide | null; diagonalUp?: boolean; diagonalDown?: boolean; raw?: string; }
export interface Alignment { horizontal?: string; vertical?: string; wrapText?: boolean; indent?: number; textRotation?: number; shrinkToFit?: boolean; readingOrder?: number; }
export interface Xf {
  numFmtId: number; fontId: number; fillId: number; borderId: number; xfId: number;
  alignment: Alignment | null;
  protection: { locked?: boolean; hidden?: boolean } | null;
  applyNumberFormat?: boolean; applyFont?: boolean; applyFill?: boolean; applyBorder?: boolean; applyAlignment?: boolean; applyProtection?: boolean;
  quotePrefix?: boolean;
}

export interface Styles {
  numFmts: Map<number, string>;
  fonts: Font[];
  fills: Fill[];
  borders: Border[];
  cellStyleXfs: Xf[];
  xfs: Xf[];
  rawTail: RawBlock[];       // cellStyles, dxfs, tableStyles, colors, extLst kept verbatim
  rootAttrs: string;
}

export const DEFAULT_FONT: Font = { name: "Calibri", size: 11, bold: false, italic: false, underline: null, strike: false, color: null };

export interface DefinedName { name: string; ref: string; localSheetId: number | null; hidden: boolean; raw: string; sheet?: Sheet | null; }

export interface Workbook {
  sheets: Sheet[];
  active: number;
  date1904: boolean;
  definedNames: DefinedName[];
  styles: Styles;
  sst: string[];             // plain text of shared strings
  sstRaw: string[];          // original <si> XML (for rich text round trip)
  sstIndex: Map<string, number>;
  themeColors: string[];     // 12 scheme colours RRGGBB (dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink)
  pkg: Package;
  workbookPart: string;
  workbookXml: string;
  workbookRels: string;
  calcChainPart: string | null;
  stylesPart: string;
  sstPart: string | null;
  path: string | null;
  kind: "xlsx" | "csv";
  csvOptions?: { delimiter: string; encoding: string; bom: boolean };
  removed?: { part: string; rId: string }[];   // sheets deleted in this session (parts to drop on save)
}

// ---- A1 helpers -------------------------------------------------------------

export function key(r: number, c: number): number { return r * MAXC + c; }
export function rowOf(k: number): number { return Math.floor(k / MAXC); }
export function colOf(k: number): number { return k % MAXC; }

export function colName(c: number): string {
  let s = "";
  c += 1;
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}
export function colIndex(name: string): number {
  let n = 0;
  for (let i = 0; i < name.length; i++) n = n * 26 + (name.charCodeAt(i) - 64);
  return n - 1;
}

export interface RefParts { sheet: string | null; r1: number; c1: number; r2: number; c2: number; absR1: boolean; absC1: boolean; absR2: boolean; absC2: boolean; isRange: boolean; colOnly?: boolean; rowOnly?: boolean; }

const REF_RE = /^(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.À-￿]+))!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?$/i;
const COL_RANGE_RE = /^(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.À-￿]+))!)?(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})$/i;
const ROW_RANGE_RE = /^(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.À-￿]+))!)?(\$?)(\d+):(\$?)(\d+)$/;

export function parseRef(text: string): RefParts | null {
  let m = REF_RE.exec(text);
  if (m) {
    const sheet = m[1] ? m[1].replace(/''/g, "'") : m[2] || null;
    const c1 = colIndex(m[4].toUpperCase()), r1 = parseInt(m[6], 10) - 1;
    if (m[8]) {
      const c2 = colIndex(m[8].toUpperCase()), r2 = parseInt(m[10], 10) - 1;
      return { sheet, r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2), absC1: !!m[3], absR1: !!m[5], absC2: !!m[7], absR2: !!m[9], isRange: true };
    }
    return { sheet, r1, c1, r2: r1, c2: c1, absC1: !!m[3], absR1: !!m[5], absC2: !!m[3], absR2: !!m[5], isRange: false };
  }
  m = COL_RANGE_RE.exec(text);
  if (m) {
    const sheet = m[1] ? m[1].replace(/''/g, "'") : m[2] || null;
    const c1 = colIndex(m[4].toUpperCase()), c2 = colIndex(m[6].toUpperCase());
    return { sheet, r1: 0, c1: Math.min(c1, c2), r2: MAXR - 1, c2: Math.max(c1, c2), absC1: !!m[3], absR1: true, absC2: !!m[5], absR2: true, isRange: true, colOnly: true };
  }
  m = ROW_RANGE_RE.exec(text);
  if (m) {
    const sheet = m[1] ? m[1].replace(/''/g, "'") : m[2] || null;
    const r1 = parseInt(m[4], 10) - 1, r2 = parseInt(m[6], 10) - 1;
    return { sheet, r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAXC - 1, absC1: true, absR1: !!m[3], absC2: true, absR2: !!m[5], isRange: true, rowOnly: true };
  }
  return null;
}

export function cellRef(r: number, c: number, absR = false, absC = false): string {
  return (absC ? "$" : "") + colName(c) + (absR ? "$" : "") + (r + 1);
}
export function rangeRef(rg: Range): string {
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return cellRef(rg.r1, rg.c1);
  return cellRef(rg.r1, rg.c1) + ":" + cellRef(rg.r2, rg.c2);
}
export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : "'" + name.replace(/'/g, "''") + "'";
}
export function normRange(a: { r: number; c: number }, b: { r: number; c: number }): Range {
  return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
}
export function inRange(rg: Range, r: number, c: number): boolean { return r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2; }
export function rangesIntersect(a: Range, b: Range): boolean { return !(a.r2 < b.r1 || b.r2 < a.r1 || a.c2 < b.c1 || b.c2 < a.c1); }

export function isError(v: Value): v is CellError { return !!v && typeof v === "object" && "e" in v; }

// ---- sheet helpers ----------------------------------------------------------

export function newSheet(name: string, sheetId: number, rId: string, part: string): Sheet {
  return {
    name, sheetId, rId, part, state: "visible",
    cells: new Map(), rows: new Map(), cols: [], merges: [], freeze: null, autoFreeze: false,
    view: { showGridLines: true, zoom: 100, tabSelected: false },
    defaultRowHeight: 15, defaultColWidth: 8.43, autoFilter: null, hyperlinks: [], tabColor: null,
    maxRow: -1, maxCol: -1, rawBlocks: [], rawRootAttrs: "", relsXml: null, dirty: false, hiddenRowsByFilter: new Set(), dirtyRows: new Set(),
  };
}

export function getCell(sheet: Sheet, r: number, c: number): Cell | undefined { return sheet.cells.get(key(r, c)); }

/** Column width in character units for column c (0-based). */
export function colWidthChars(sheet: Sheet, c: number): number {
  for (const ci of sheet.cols) if (c >= ci.min && c <= ci.max && ci.width !== undefined) return ci.width;
  return sheet.defaultColWidth;
}
export function colHidden(sheet: Sheet, c: number): boolean {
  for (const ci of sheet.cols) if (c >= ci.min && c <= ci.max && ci.hidden) return true;
  return false;
}
/** Convert Excel character-width units to pixels (Calibri 11: max digit width 7 px). */
export function charsToPx(w: number, mdw = 7): number { return Math.round(((w * mdw + 5) / mdw) * mdw); }
export function pxToChars(px: number, mdw = 7): number { return Math.round(((px - 5) / mdw) * 100) / 100; }

export function rowHeightPx(sheet: Sheet, r: number): number {
  const ri = sheet.rows.get(r);
  if (ri?.hidden || sheet.hiddenRowsByFilter.has(r)) return 0;
  const pt = ri?.height ?? sheet.defaultRowHeight;
  return Math.round(pt * 96 / 72);
}

/** Set a column width (character units), splitting existing <col> ranges as needed. */
export function setColWidth(sheet: Sheet, c: number, width: number | undefined, hidden?: boolean) {
  const out: ColInfo[] = [];
  let found = false;
  for (const ci of sheet.cols) {
    if (c < ci.min || c > ci.max) { out.push(ci); continue; }
    if (ci.min < c) out.push({ ...ci, max: c - 1 });
    out.push({ ...ci, min: c, max: c, width: width ?? ci.width, customWidth: width !== undefined ? true : ci.customWidth, hidden: hidden ?? ci.hidden });
    if (ci.max > c) out.push({ ...ci, min: c + 1 });
    found = true;
  }
  if (!found) out.push({ min: c, max: c, width, customWidth: width !== undefined, hidden });
  out.sort((a, b) => a.min - b.min);
  sheet.cols = out;
  sheet.dirty = true;
}

export function setRowHeight(sheet: Sheet, r: number, height: number | undefined, hidden?: boolean) {
  const ri = { ...(sheet.rows.get(r) || {}) };
  if (height !== undefined) { ri.height = height; ri.customHeight = true; }
  if (hidden !== undefined) ri.hidden = hidden;
  sheet.rows.set(r, ri);
  sheet.dirty = true;
}

export function updateExtent(sheet: Sheet, r: number, c: number) {
  if (r > sheet.maxRow) sheet.maxRow = r;
  if (c > sheet.maxCol) sheet.maxCol = c;
}

export function recomputeExtent(sheet: Sheet) {
  let mr = -1, mc = -1;
  for (const k of sheet.cells.keys()) { const r = rowOf(k), c = colOf(k); if (r > mr) mr = r; if (c > mc) mc = c; }
  sheet.maxRow = mr; sheet.maxCol = mc;
}

/** Merged range containing (r,c), if any. */
export function mergeAt(sheet: Sheet, r: number, c: number): Range | null {
  for (const m of sheet.merges) if (inRange(m, r, c)) return m;
  return null;
}
