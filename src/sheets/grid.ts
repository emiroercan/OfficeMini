// Canvas grid: virtualised rendering of a worksheet with frozen panes, native
// scrolling (a spacer element gives the scrollbars), selection, hit-testing,
// column/row resizing and the fill handle. Editing lives in app.ts.
import { Sheet, Workbook, Range, Cell, key, colName, colWidthChars, colHidden, charsToPx, rowHeightPx, mergeAt, inRange, normRange, Hyperlink, MAXC, MAXR, isError } from "./model";
import { StyleResolver, CellStyle } from "./render-style";
import { fontFamilyCss } from "./render-style";

export type Hit =
  | { type: "cell"; r: number; c: number }
  | { type: "colHeader"; c: number }
  | { type: "rowHeader"; r: number }
  | { type: "corner" }
  | { type: "colResize"; c: number }
  | { type: "rowResize"; r: number }
  | { type: "fillHandle" }
  | { type: "filterBtn"; c: number; r: number }
  | { type: "outside" };

export interface GridEvents {
  onSelect(): void;
  onEdit(seed: string | null): void;
  onContextMenu(e: MouseEvent, hit: Hit): void;
  onColResize(c: number, widthPx: number, all: number[]): void;
  onRowResize(r: number, heightPx: number, all: number[]): void;
  onAutoFit(kind: "col" | "row", index: number): void;
  onFill(src: Range, dst: Range): void;
  onLink(hl: Hyperlink): void;
  onFilterButton(c: number, r: number, x: number, y: number): void;
  onMoveRange(src: Range, dst: Range, copy: boolean): void;
}

export interface Selection { ranges: Range[]; active: { r: number; c: number }; anchor: { r: number; c: number }; }

interface Colors { paper: string; text: string; grid: string; headerBg: string; headerText: string; headerSel: string; accent: string; selFill: string; frozenLine: string; }

const HEADER_W = 46, HEADER_H = 22, FILL_HANDLE = 6, RESIZE_ZONE = 5;
/**
 * A sheet with nothing in it is 1000 rows by 26 columns, the way Sheets starts one: the
 * scrollbars then mean something instead of running to row 1048576. A sheet that holds more
 * than that is as big as its data, and moving the selection past the edge grows it.
 */
const DEFAULT_ROWS = 1000, DEFAULT_COLS = 26;
/** Smallest filter button we still draw, so tiny header cells keep a visible one. */
const FILTER_BTN_MIN = 9;

export class Grid {
  host: HTMLElement;
  private spacer: HTMLElement;
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  wb: Workbook;
  private sheetIndex = 0;
  styles: StyleResolver;
  ev: GridEvents;
  zoom = 1;
  selection: Selection = { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 }, anchor: { r: 0, c: 0 } };
  highlights: { range: Range; color: string; sheet?: string }[] = [];
  private colX: number[] = [];   // prefix sums (px, zoomed) for columns 0..ncols
  private rowY: number[] = [];   // prefix sums for rows 0..nrows
  private ncols = 30;
  private nrows = 200;
  private raf = 0;
  private colors!: Colors;
  private fontCache = new Map<string, string>();
  private textCache = new WeakMap<Cell, { key: string; text: string; color: string | null; align: string | null }>();
  private drag: null | { kind: "select" | "col" | "row" | "resizeCol" | "resizeRow" | "fill" | "move"; start: { r: number; c: number }; additive?: boolean; index?: number; startPos?: number; startSize?: number; fillTarget?: Range; moveTarget?: Range; moveCopy?: boolean } = null;
  private autoScroll = 0;
  private resizeObserver: ResizeObserver;
  private dpr = 1;
  private fitCache = new Map<string, number>();
  hidePageBreaks = true;
  pageBreaks: { rows: number[]; cols: number[] } | null = null;
  /** Range last copied/cut: drawn with a marching dashed outline until Esc / paste. */
  clipRange: Range | null = null;
  /** Show formula text instead of values. */
  showFormulas = false;
  /** Columns with an active filter get a filled funnel. */
  filteredCols: (c: number) => boolean = () => false;

  constructor(host: HTMLElement, wb: Workbook, styles: StyleResolver, ev: GridEvents) {
    this.host = host; this.wb = wb; this.styles = styles; this.ev = ev;
    host.classList.add("sheet-host");
    host.innerHTML = "";
    this.spacer = document.createElement("div");
    this.spacer.className = "sheet-spacer";
    this.canvas = document.createElement("canvas");
    this.canvas.className = "sheet-canvas";
    host.append(this.spacer, this.canvas);
    this.ctx = this.canvas.getContext("2d", { alpha: false })!;
    this.readColors();
    host.addEventListener("scroll", () => this.schedule(), { passive: true });
    this.resizeObserver = new ResizeObserver(() => { this.resizeCanvas(); this.schedule(); });
    this.resizeObserver.observe(host);
    this.canvas.addEventListener("mousedown", (e) => this.onMouseDown(e));
    this.canvas.addEventListener("mousemove", (e) => this.onMouseMove(e));
    this.canvas.addEventListener("dblclick", (e) => this.onDblClick(e));
    this.canvas.addEventListener("contextmenu", (e) => { e.preventDefault(); const hit = this.hitTest(e); if (hit.type === "cell" && !this.isSelected(hit.r, hit.c)) { this.setActive(hit.r, hit.c); } this.ev.onContextMenu(e, hit); });
    window.addEventListener("mouseup", (e) => this.onMouseUp(e));
    window.addEventListener("mousemove", (e) => { if (this.drag) this.onMouseMove(e); });
    this.setSheet(wb.active);
  }

  destroy() { this.resizeObserver.disconnect(); cancelAnimationFrame(this.raf); }

  readColors() {
    const cs = getComputedStyle(document.documentElement);
    const v = (n: string, d: string) => cs.getPropertyValue(n).trim() || d;
    const dark = document.documentElement.getAttribute("data-theme") === "dark";
    this.colors = {
      paper: v("--paper", "#ffffff"), text: v("--paper-text", "#000000"), grid: dark ? "#3a3b40" : "#e1e1e1",
      headerBg: v("--ui-bg", "#f6f6f7"), headerText: v("--ui-muted", "#6b6b73"), headerSel: dark ? "#33405f" : "#d3e3fd", accent: v("--ui-accent", "#2f6fed"),
      selFill: dark ? "rgba(109,156,255,.18)" : "rgba(47,111,237,.12)", frozenLine: dark ? "#6b6f7a" : "#9aa0a6",
    };
    this.fontCache.clear();
    this.textCache = new WeakMap();
  }

  sheet(): Sheet { return this.wb.sheets[this.sheetIndex]; }
  sheetIdx(): number { return this.sheetIndex; }

  setSheet(i: number) {
    this.sheetIndex = Math.max(0, Math.min(i, this.wb.sheets.length - 1));
    const s = this.sheet();
    this.textCache = new WeakMap();
    this.fitCache.clear();
    this.ncols = 0; this.nrows = 0;
    this.cursorExt = null;
    this.layout();
    this.selection = { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { r: 0, c: 0 }, anchor: { r: 0, c: 0 } };
    if (s.view.active) { const m = /^([A-Z]+)(\d+)$/.exec(s.view.active); if (m) { const c = colIndexOf(m[1]), r = parseInt(m[2], 10) - 1; this.selection = { ranges: [{ r1: r, c1: c, r2: r, c2: c }], active: { r, c }, anchor: { r, c } }; } }
    this.host.scrollTop = 0; this.host.scrollLeft = 0;
    this.ensureVisible(this.selection.active.r, this.selection.active.c);
    this.schedule();
    this.ev.onSelect();
  }

  // ---- layout ---------------------------------------------------------------

  /** Rows without an explicit height that contain wrapped / multi-line text get an automatic height (like Excel). */
  private autoHeights = new Map<number, number>();
  private autoSheet: Sheet | null = null;
  private autoStyleVersion = -1;
  private autoWrapIdx = new Set<number>();
  /**
   * Full scan when the sheet or the style table changed (cheap per cell, but 800k-cell sheets take
   * ~50 ms), otherwise only the rows edited since the last measurement are re-measured.
   */
  private computeAutoHeights() {
    const s = this.sheet();
    const st = this.wb.styles;
    const full = this.autoSheet !== s || this.autoStyleVersion !== this.styles.version;
    if (full) {
      this.autoHeights.clear();
      this.autoSheet = s; this.autoStyleVersion = this.styles.version;
      const wrapIdx = new Set<number>();
      st.xfs.forEach((xf, i) => { if (xf.alignment?.wrapText) wrapIdx.add(i); });
      this.autoWrapIdx = wrapIdx;
      const rows = new Set<number>();
      for (const [k, cell] of s.cells) {
        const v = cell.v;
        if (v === null) continue;
        if (!wrapIdx.has(cell.s) && !(typeof v === "string" && v.length > 1 && v.indexOf("\n") >= 0)) continue;
        const r = Math.floor(k / MAXC);
        const ri = s.rows.get(r);
        if (ri && (ri.height !== undefined || ri.hidden)) continue;
        rows.add(r);
        if (rows.size > 20000) break; // pathological sheets: leave the rest at default height
      }
      for (const r of rows) { const h = this.measureRow(r); if (h > rowHeightPx(s, r)) this.autoHeights.set(r, h); }
      s.dirtyRows.clear();
      return;
    }
    if (!s.dirtyRows.size) return;
    for (const r of s.dirtyRows) {
      this.autoHeights.delete(r);
      const ri = s.rows.get(r);
      if (ri && (ri.height !== undefined || ri.hidden)) continue;
      let needs = false;
      for (let c = 0; c <= s.maxCol && !needs; c++) { const cell = s.cells.get(key(r, c)); if (cell && cell.v !== null && (this.autoWrapIdx.has(cell.s) || (typeof cell.v === "string" && cell.v.indexOf("\n") >= 0))) needs = true; }
      if (!needs) continue;
      const h = this.measureRow(r);
      if (h > rowHeightPx(s, r)) this.autoHeights.set(r, h);
    }
    s.dirtyRows.clear();
  }
  /** Unzoomed row height honouring hidden rows, filters and automatic heights. */
  rowH(r: number): number {
    const s = this.sheet();
    const base = rowHeightPx(s, r);
    if (base === 0) return 0;
    const auto = this.autoHeights.get(r);
    return auto !== undefined && s.rows.get(r)?.height === undefined ? auto : base;
  }

  /** Recompute prefix sums; call after any size/visibility change. */
  layout() {
    const s = this.sheet();
    // keep any extent already grown (an edit must not snap the view back), reset by setSheet
    this.ncols = Math.min(MAXC, Math.max(this.ncols, s.maxCol + 1, DEFAULT_COLS));
    this.nrows = Math.min(MAXR, Math.max(this.nrows, s.maxRow + 1, DEFAULT_ROWS));
    const z = this.zoom;
    const cx: number[] = [0];
    for (let c = 0; c < this.ncols; c++) cx.push(cx[c] + (colHidden(s, c) ? 0 : Math.round(charsToPx(colWidthChars(s, c)) * z)));
    this.colX = cx;
    this.computeAutoHeights();
    const ry: number[] = [0];
    for (let r = 0; r < this.nrows; r++) ry.push(ry[r] + Math.round(this.rowH(r) * z));
    this.rowY = ry;
    this.spacer.style.width = (this.colX[this.ncols] + HEADER_W * z) + "px";
    this.spacer.style.height = (this.rowY[this.nrows] + HEADER_H * z) + "px";
  }

  /** Rows and columns the sheet currently reaches; the scroll area is exactly this big. */
  extent(): { rows: number; cols: number } { return { rows: this.nrows, cols: this.ncols }; }

  /**
   * Grow the sheet so (r, c) is inside it. Moving the selection past the last row or column
   * is what adds more - scrolling does not, or the sheet would creep outwards on its own.
   */
  ensureExtent(r: number, c: number) {
    const rows = Math.min(MAXR, Math.max(this.nrows, r + 1));
    const cols = Math.min(MAXC, Math.max(this.ncols, c + 1));
    if (rows === this.nrows && cols === this.ncols) return;
    this.nrows = rows; this.ncols = cols;
    this.layout();
  }

  colWidthPx(c: number): number { return c < this.ncols ? this.colX[c + 1] - this.colX[c] : Math.round(charsToPx(this.sheet().defaultColWidth) * this.zoom); }
  rowHeightPxZ(r: number): number { return r < this.nrows ? this.rowY[r + 1] - this.rowY[r] : Math.round(this.rowH(r) * this.zoom); }
  private colStart(c: number): number { return c < this.colX.length ? this.colX[Math.min(c, this.ncols)] : this.colX[this.ncols] + (c - this.ncols) * this.colWidthPx(c); }
  private rowStart(r: number): number { return r < this.rowY.length ? this.rowY[Math.min(r, this.nrows)] : this.rowY[this.nrows] + (r - this.nrows) * this.rowHeightPxZ(r); }

  private frozenRows(): number { return this.sheet().freeze?.rows || 0; }
  private frozenCols(): number { return this.sheet().freeze?.cols || 0; }
  private frozenW(): number { return this.colStart(this.frozenCols()); }
  private frozenH(): number { return this.rowStart(this.frozenRows()); }

  /** Row index at a y position in the scrollable content (px from top of all rows). */
  private rowIndexAt(y: number): number {
    const ry = this.rowY;
    let lo = 0, hi = this.nrows;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (ry[mid + 1] <= y) lo = mid + 1; else hi = mid; }
    return lo;
  }
  private colIndexAt(x: number): number {
    const cx = this.colX;
    let lo = 0, hi = this.ncols;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (cx[mid + 1] <= x) lo = mid + 1; else hi = mid; }
    return lo;
  }

  private resizeCanvas() {
    this.dpr = window.devicePixelRatio || 1;
    const w = this.host.clientWidth, h = this.host.clientHeight;
    this.canvas.width = Math.max(1, Math.round(w * this.dpr));
    this.canvas.height = Math.max(1, Math.round(h * this.dpr));
    this.canvas.style.width = w + "px";
    this.canvas.style.height = h + "px";
  }

  schedule() { if (!this.raf) this.raf = requestAnimationFrame(() => { this.raf = 0; this.render(); }); }
  invalidate() { this.textCache = new WeakMap(); this.fitCache.clear(); this.layout(); this.schedule(); }

  // ---- geometry helpers ------------------------------------------------------

  /** Rectangle of a cell in canvas (host client) coordinates, honouring frozen panes. */
  cellRect(r: number, c: number): { x: number; y: number; w: number; h: number; visible: boolean } {
    const z = this.zoom, hw = HEADER_W * z, hh = HEADER_H * z;
    const fr = this.frozenRows(), fc = this.frozenCols();
    const st = this.host.scrollTop, sl = this.host.scrollLeft;
    const x = c < fc ? hw + this.colStart(c) : hw + this.colStart(c) - sl;
    const y = r < fr ? hh + this.rowStart(r) : hh + this.rowStart(r) - st;
    const w = this.colWidthPx(c), h = this.rowHeightPxZ(r);
    const visible = x + w > hw && y + h > hh && x < this.host.clientWidth && y < this.host.clientHeight && (c >= fc ? x >= hw + this.frozenW() - 1 : true) && (r >= fr ? y >= hh + this.frozenH() - 1 : true);
    return { x, y, w, h, visible };
  }

  rangeRect(rg: Range) {
    const a = this.cellRect(rg.r1, rg.c1), b = this.cellRect(rg.r2, rg.c2);
    return { x: a.x, y: a.y, w: b.x + b.w - a.x, h: b.y + b.h - a.y };
  }

  ensureVisible(r: number, c: number) {
    const z = this.zoom;
    const fr = this.frozenRows(), fc = this.frozenCols();
    const viewW = this.host.clientWidth - HEADER_W * z, viewH = this.host.clientHeight - HEADER_H * z;
    if (c >= fc) {
      const x0 = this.colStart(c) - this.frozenW(), x1 = x0 + this.colWidthPx(c);
      const sl = this.host.scrollLeft;
      const avail = viewW - this.frozenW();
      if (x0 < sl) this.host.scrollLeft = x0;
      else if (x1 > sl + avail) this.host.scrollLeft = Math.max(0, x1 - avail);
    }
    if (r >= fr) {
      const y0 = this.rowStart(r) - this.frozenH(), y1 = y0 + this.rowHeightPxZ(r);
      const st = this.host.scrollTop;
      const avail = viewH - this.frozenH();
      if (y0 < st) this.host.scrollTop = y0;
      else if (y1 > st + avail) this.host.scrollTop = Math.max(0, y1 - avail);
    }
  }

  hitTest(e: MouseEvent): Hit {
    const rect = this.canvas.getBoundingClientRect();
    const px = e.clientX - rect.left, py = e.clientY - rect.top;
    return this.hitAt(px, py);
  }

  hitAt(px: number, py: number): Hit {
    const z = this.zoom, hw = HEADER_W * z, hh = HEADER_H * z;
    if (px < 0 || py < 0 || px > this.host.clientWidth || py > this.host.clientHeight) return { type: "outside" };
    const fr = this.frozenRows(), fc = this.frozenCols();
    const st = this.host.scrollTop, sl = this.host.scrollLeft;
    // column index for x
    const colAt = (x: number): number => {
      const rel = x - hw;
      if (rel < this.frozenW()) return this.colIndexAt(Math.max(0, rel));
      return this.colIndexAt(rel + sl);
    };
    const rowAt = (y: number): number => {
      const rel = y - hh;
      if (rel < this.frozenH()) return this.rowIndexAt(Math.max(0, rel));
      return this.rowIndexAt(rel + st);
    };
    if (px < hw && py < hh) return { type: "corner" };
    if (py < hh) {
      const c = colAt(px);
      // resize zone near the right edge of column c (or left edge of next)
      const right = this.cellRect(0, c).x + this.colWidthPx(c);
      if (Math.abs(px - right) <= RESIZE_ZONE) return { type: "colResize", c };
      const left = this.cellRect(0, c).x;
      if (c > 0 && Math.abs(px - left) <= RESIZE_ZONE) return { type: "colResize", c: c - 1 };
      return { type: "colHeader", c };
    }
    if (px < hw) {
      const r = rowAt(py);
      const bottom = this.cellRect(r, 0).y + this.rowHeightPxZ(r);
      if (Math.abs(py - bottom) <= RESIZE_ZONE) return { type: "rowResize", r };
      const top = this.cellRect(r, 0).y;
      if (r > 0 && Math.abs(py - top) <= RESIZE_ZONE) return { type: "rowResize", r: r - 1 };
      return { type: "rowHeader", r };
    }
    // fill handle?
    const sel = this.selection.ranges[0];
    const rr = this.rangeRect(sel);
    const hx = rr.x + rr.w, hy = rr.y + rr.h;
    if (Math.abs(px - hx) <= FILL_HANDLE && Math.abs(py - hy) <= FILL_HANDLE) return { type: "fillHandle" };
    const c = colAt(px), r = rowAt(py);
    // filter button
    const s = this.sheet();
    if (s.autoFilter && r === s.autoFilter.r1 && c >= s.autoFilter.c1 && c <= s.autoFilter.c2) {
      const cr = this.cellRect(r, c);
      const b = this.filterBtnRect(cr.x, cr.y, cr.w, cr.h);
      if (px >= b.x - 2 && py >= b.y - 2 && px <= b.x + b.s + 2 && py <= b.y + b.s + 2) return { type: "filterBtn", c, r };
    }
    void fr; void fc;
    return { type: "cell", r, c };
  }

  // ---- selection --------------------------------------------------------------

  isSelected(r: number, c: number): boolean { return this.selection.ranges.some((rg) => inRange(rg, r, c)); }

  setActive(r: number, c: number, extend = false) {
    r = Math.max(0, Math.min(MAXR - 1, r)); c = Math.max(0, Math.min(MAXC - 1, c));
    this.ensureExtent(r, c);
    if (!extend) this.cursorExt = null;
    const merge = mergeAt(this.sheet(), r, c);
    if (!extend) {
      this.selection.anchor = { r, c };
      this.selection.active = merge ? { r: merge.r1, c: merge.c1 } : { r, c };
      this.selection.ranges = [merge ? { ...merge } : { r1: r, c1: c, r2: r, c2: c }];
    } else {
      const a = this.selection.anchor;
      let rg = normRange(a, { r, c });
      rg = this.expandToMerges(rg);
      this.selection.ranges = [rg];
      this.selection.active = merge ? { r: merge.r1, c: merge.c1 } : { r, c };
    }
    this.ensureVisible(r, c);
    this.schedule();
    this.ev.onSelect();
  }

  setRanges(ranges: Range[], active?: { r: number; c: number }) {
    this.cursorExt = null;
    this.selection.ranges = ranges.map((rg) => this.expandToMerges(rg));
    if (active) this.selection.active = active; else this.selection.active = { r: ranges[0].r1, c: ranges[0].c1 };
    // Only the active cell grows the sheet: whole-column ranges run to the last row by design.
    this.ensureExtent(this.selection.active.r, this.selection.active.c);
    this.selection.anchor = { ...this.selection.active };
    this.schedule();
    this.ev.onSelect();
  }

  addRange(rg: Range) { this.selection.ranges.push(this.expandToMerges(rg)); this.schedule(); this.ev.onSelect(); }

  private expandToMerges(rg: Range): Range {
    const s = this.sheet();
    let out = { ...rg };
    let changed = true;
    while (changed) {
      changed = false;
      for (const m of s.merges) {
        if (m.r2 < out.r1 || m.r1 > out.r2 || m.c2 < out.c1 || m.c1 > out.c2) continue;
        const n = { r1: Math.min(out.r1, m.r1), c1: Math.min(out.c1, m.c1), r2: Math.max(out.r2, m.r2), c2: Math.max(out.c2, m.c2) };
        if (n.r1 !== out.r1 || n.c1 !== out.c1 || n.r2 !== out.r2 || n.c2 !== out.c2) { out = n; changed = true; }
      }
    }
    return out;
  }

  /** Arrow-key movement; `jump` = Ctrl (to the edge of a data block like Excel/Sheets). */
  moveActive(dr: number, dc: number, extend: boolean, jump: boolean) {
    const s = this.sheet();
    const base = extend ? this.selection.ranges[0] : null;
    let { r, c } = extend ? this.extendCursor() : this.selection.active;
    const startR = r, startC = c;
    // The sheet ends at its extent. A jump (Ctrl+arrow) stops on the last row/column;
    // a single step may land one past it, and setActive then grows the sheet by that much.
    const maxR = Math.min(MAXR - 1, jump ? this.nrows - 1 : this.nrows);
    const maxC = Math.min(MAXC - 1, jump ? this.ncols - 1 : this.ncols);
    if (!extend) {
      // step over merged cells
      const m = mergeAt(s, r, c);
      if (m) { if (dr > 0) r = m.r2; if (dc > 0) c = m.c2; if (dr < 0) r = m.r1; if (dc < 0) c = m.c1; }
    }
    if (jump) {
      const has = (rr: number, cc: number) => { const cell = s.cells.get(key(rr, cc)); return !!cell && cell.v !== null && cell.v !== ""; };
      const step = () => { r += dr; c += dc; };
      const inBounds = () => r >= 0 && c >= 0 && r <= maxR && c <= maxC;
      const cur = has(r, c);
      const r0 = r, c0 = c;
      const nr = r + dr, nc = c + dc;
      if (cur && nr >= 0 && nc >= 0 && has(nr, nc)) { while (inBounds() && has(r + dr, c + dc)) step(); }
      else { step(); while (inBounds() && !has(r, c)) { if ((dr > 0 && r >= Math.max(s.maxRow, 0)) || (dc > 0 && c >= Math.max(s.maxCol, 0)) || (dr < 0 && r <= 0) || (dc < 0 && c <= 0)) break; step(); } if (!inBounds()) { r = Math.max(0, Math.min(maxR, r)); c = Math.max(0, Math.min(maxC, c)); } if (!has(r, c)) { if (dr > 0) r = r0 >= Math.max(s.maxRow, 0) ? maxR : Math.max(s.maxRow, 0); if (dc > 0) c = c0 >= Math.max(s.maxCol, 0) ? maxC : Math.max(s.maxCol, 0); if (dr < 0) r = 0; if (dc < 0) c = 0; } }
    } else { r += dr; c += dc; }
    r = Math.max(0, Math.min(maxR, r)); c = Math.max(0, Math.min(maxC, c));
    // skip hidden rows/cols
    let guard = 0;
    while (guard++ < maxR + 1 && ((dr && this.rowH(r) === 0) || (dc && colHidden(s, c)))) { r += dr || 0; c += dc || 0; if (r < 0 || c < 0 || r > maxR || c > maxC) { r = Math.max(0, Math.min(maxR, r)); c = Math.max(0, Math.min(maxC, c)); break; } }
    // no visible row/column in that direction: stay put instead of landing on a hidden one
    if ((dr && this.rowH(r) === 0) || (dc && colHidden(s, c))) { r = startR; c = startC; }
    if (extend && base) { this.selection.ranges = [this.expandToMerges(normRange(this.selection.anchor, { r, c }))]; this.cursorExt = { r, c }; this.ensureVisible(r, c); this.schedule(); this.ev.onSelect(); }
    else this.setActive(r, c);
  }
  private lastResize: { kind: "col" | "row"; index: number; at: number } | null = null;
  private cursorExt: { r: number; c: number } | null = null;
  private extendCursor() { return this.cursorExt && this.selection.ranges.length ? this.cursorExt : { ...this.selection.active }; }

  pageMove(dir: 1 | -1, extend: boolean) {
    const z = this.zoom;
    const visible = Math.max(1, Math.floor((this.host.clientHeight - HEADER_H * z - this.frozenH()) / Math.max(1, this.rowHeightPxZ(this.selection.active.r))) - 1);
    this.moveActive(dir * visible, 0, extend, false);
  }

  selectAll() { const s = this.sheet(); this.setRanges([{ r1: 0, c1: 0, r2: Math.max(0, Math.max(s.maxRow, MAXR - 1)), c2: MAXC - 1 }], { r: 0, c: 0 }); }
  selectRows(r1: number, r2: number) { this.setRanges([{ r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAXC - 1 }], { r: Math.min(r1, r2), c: this.selection.active.c }); }
  selectCols(c1: number, c2: number) { this.setRanges([{ r1: 0, c1: Math.min(c1, c2), r2: MAXR - 1, c2: Math.max(c1, c2) }], { r: this.selection.active.r, c: Math.min(c1, c2) }); }

  /** Ctrl+click / Ctrl+drag on the headers: another column or row alongside what is selected. */
  addCols(c1: number, c2: number, replaceLast = false) {
    const rg = { r1: 0, c1: Math.min(c1, c2), r2: MAXR - 1, c2: Math.max(c1, c2) };
    if (replaceLast && this.selection.ranges.length) this.selection.ranges[this.selection.ranges.length - 1] = rg;
    else this.selection.ranges.push(rg);
    this.selection.active = { r: this.selection.active.r, c: rg.c1 };
    this.selection.anchor = { r: 0, c: rg.c1 };
    this.schedule();
    this.ev.onSelect();
  }

  addRows(r1: number, r2: number, replaceLast = false) {
    const rg = { r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAXC - 1 };
    if (replaceLast && this.selection.ranges.length) this.selection.ranges[this.selection.ranges.length - 1] = rg;
    else this.selection.ranges.push(rg);
    this.selection.active = { r: rg.r1, c: this.selection.active.c };
    this.selection.anchor = { r: rg.r1, c: 0 };
    this.schedule();
    this.ev.onSelect();
  }

  // ---- mouse -------------------------------------------------------------------

  private onMouseDown(e: MouseEvent) {
    if (e.button === 1) return;
    const hit = this.hitTest(e);
    this.host.focus();
    if (e.button === 2) return;
    const shift = e.shiftKey, ctrl = e.ctrlKey || e.metaKey;
    switch (hit.type) {
      case "corner": this.selectAll(); return;
      case "colResize": this.lastResize = { kind: "col", index: hit.c, at: Date.now() }; this.drag = { kind: "resizeCol", start: { r: 0, c: hit.c }, index: hit.c, startPos: e.clientX, startSize: this.colWidthPx(hit.c) }; e.preventDefault(); return;
      case "rowResize": this.lastResize = { kind: "row", index: hit.r, at: Date.now() }; this.drag = { kind: "resizeRow", start: { r: hit.r, c: 0 }, index: hit.r, startPos: e.clientY, startSize: this.rowHeightPxZ(hit.r) }; e.preventDefault(); return;
      case "colHeader":
        if (ctrl && !shift) this.addCols(hit.c, hit.c);
        else if (shift) this.selectCols(this.selection.anchor.c, hit.c);
        else { this.selectCols(hit.c, hit.c); this.selection.anchor = { r: 0, c: hit.c }; }
        this.drag = { kind: "col", start: { r: 0, c: hit.c }, additive: ctrl && !shift }; e.preventDefault(); return;
      case "rowHeader":
        if (ctrl && !shift) this.addRows(hit.r, hit.r);
        else if (shift) this.selectRows(this.selection.anchor.r, hit.r);
        else { this.selectRows(hit.r, hit.r); this.selection.anchor = { r: hit.r, c: 0 }; }
        this.drag = { kind: "row", start: { r: hit.r, c: 0 }, additive: ctrl && !shift }; e.preventDefault(); return;
      case "fillHandle": this.drag = { kind: "fill", start: { ...this.selection.active }, fillTarget: { ...this.selection.ranges[0] } }; e.preventDefault(); return;
      case "filterBtn": { const cr = this.cellRect(hit.r, hit.c); const br = this.canvas.getBoundingClientRect(); this.ev.onFilterButton(hit.c, hit.r, br.left + cr.x, br.top + cr.y + cr.h); return; }
      case "cell": {
        if (ctrl && !shift) {
          const hl = this.hyperlinkAt(hit.r, hit.c);
          if (hl) { this.ev.onLink(hl); return; }
          this.addRange({ r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c });
          this.selection.active = { r: hit.r, c: hit.c }; this.selection.anchor = { r: hit.r, c: hit.c };
          this.drag = { kind: "select", start: { r: hit.r, c: hit.c } };
          return;
        }
        // drag-move: mousedown on the selection border (not the interior) moves the range
        if (!shift && this.isSelected(hit.r, hit.c) && this.onSelectionBorder(e)) {
          this.drag = { kind: "move", start: { r: hit.r, c: hit.c }, moveTarget: { ...this.selection.ranges[0] }, moveCopy: ctrl };
          this.canvas.style.cursor = "move";
          e.preventDefault();
          return;
        }
        this.setActive(hit.r, hit.c, shift);
        this.cursorExt = null;
        this.drag = { kind: "select", start: shift ? this.selection.anchor : { r: hit.r, c: hit.c } };
        e.preventDefault();
        return;
      }
    }
  }

  private onSelectionBorder(e: MouseEvent): boolean {
    const rect = this.canvas.getBoundingClientRect();
    const px = e.clientX - rect.left, py = e.clientY - rect.top;
    const rr = this.rangeRect(this.selection.ranges[0]);
    const d = 4;
    const nearX = Math.abs(px - rr.x) <= d || Math.abs(px - (rr.x + rr.w)) <= d;
    const nearY = Math.abs(py - rr.y) <= d || Math.abs(py - (rr.y + rr.h)) <= d;
    const inside = px >= rr.x - d && px <= rr.x + rr.w + d && py >= rr.y - d && py <= rr.y + rr.h + d;
    return inside && (nearX || nearY) && (rr.w > 12 || rr.h > 12);
  }

  private onMouseMove(e: MouseEvent) {
    const rect = this.canvas.getBoundingClientRect();
    const px = e.clientX - rect.left, py = e.clientY - rect.top;
    if (!this.drag) {
      const hit = this.hitAt(px, py);
      let cursor = "cell";
      if (hit.type === "colResize") cursor = "col-resize";
      else if (hit.type === "rowResize") cursor = "row-resize";
      else if (hit.type === "fillHandle") cursor = "crosshair";
      else if (hit.type === "colHeader") cursor = "s-resize";
      else if (hit.type === "rowHeader") cursor = "e-resize";
      else if (hit.type === "cell" && (e.ctrlKey || e.metaKey) && this.hyperlinkAt(hit.r, hit.c)) cursor = "pointer";
      else if (hit.type === "cell" && this.isSelected(hit.r, hit.c) && this.onSelectionBorder(e)) cursor = "move";
      else if (hit.type === "filterBtn") cursor = "pointer";
      this.canvas.style.cursor = cursor === "cell" ? "cell" : cursor;
      return;
    }
    const d = this.drag;
    if (d.kind === "resizeCol") {
      const w = Math.max(4, d.startSize! + (e.clientX - d.startPos!));
      this.previewSize = { kind: "col", index: d.index!, size: w };
      this.schedule();
      return;
    }
    if (d.kind === "resizeRow") {
      const h = Math.max(4, d.startSize! + (e.clientY - d.startPos!));
      this.previewSize = { kind: "row", index: d.index!, size: h };
      this.schedule();
      return;
    }
    const hit = this.hitAt(Math.max(HEADER_W * this.zoom + 1, Math.min(px, this.host.clientWidth - 1)), Math.max(HEADER_H * this.zoom + 1, Math.min(py, this.host.clientHeight - 1)));
    if (hit.type !== "cell" && hit.type !== "fillHandle" && hit.type !== "filterBtn") return;
    const r = (hit as any).r as number, c = (hit as any).c as number;
    if (d.kind === "select") {
      const rg = this.expandToMerges(normRange(d.start, { r, c }));
      const last = this.selection.ranges.length - 1;
      this.selection.ranges[last] = rg;
      this.cursorExt = { r, c };
      this.schedule(); this.ev.onSelect();
    } else if (d.kind === "col") { if (d.additive) this.addCols(d.start.c, c, true); else this.selectCols(d.start.c, c); }
    else if (d.kind === "row") { if (d.additive) this.addRows(d.start.r, r, true); else this.selectRows(d.start.r, r); }
    else if (d.kind === "fill") {
      const src = this.selection.ranges[0];
      // extend vertically or horizontally, whichever is dominant
      const dr = r < src.r1 ? r - src.r1 : r > src.r2 ? r - src.r2 : 0;
      const dc = c < src.c1 ? c - src.c1 : c > src.c2 ? c - src.c2 : 0;
      let t: Range = { ...src };
      if (Math.abs(dr) >= Math.abs(dc)) { if (dr > 0) t.r2 = r; else if (dr < 0) t.r1 = r; }
      else { if (dc > 0) t.c2 = c; else if (dc < 0) t.c1 = c; }
      d.fillTarget = t;
      this.schedule();
    } else if (d.kind === "move") {
      const src = this.selection.ranges[0];
      const dr = r - d.start.r, dc = c - d.start.c;
      d.moveTarget = { r1: Math.max(0, src.r1 + dr), c1: Math.max(0, src.c1 + dc), r2: Math.max(0, src.r1 + dr) + (src.r2 - src.r1), c2: Math.max(0, src.c1 + dc) + (src.c2 - src.c1) };
      d.moveCopy = e.ctrlKey || e.metaKey;
      this.schedule();
    }
    // autoscroll when dragging near/outside edges
    this.autoScrollWith(px, py);
  }
  private previewSize: { kind: "col" | "row"; index: number; size: number } | null = null;
  private dashOffset = 0;
  private dashTimer = 0;

  private lastPointer: { px: number; py: number } | null = null;
  private autoScrollWith(px: number, py: number) {
    this.lastPointer = { px, py };
    const m = 24, hw = HEADER_W * this.zoom, hh = HEADER_H * this.zoom;
    let dx = 0, dy = 0;
    if (px > this.host.clientWidth - m) dx = Math.min(40, 8 + (px - (this.host.clientWidth - m))); else if (px < hw + m) dx = -Math.min(40, 8 + (hw + m - px));
    if (py > this.host.clientHeight - m) dy = Math.min(40, 8 + (py - (this.host.clientHeight - m))); else if (py < hh + m) dy = -Math.min(40, 8 + (hh + m - py));
    if (dx || dy) {
      this.host.scrollLeft += dx; this.host.scrollTop += dy;
      // keep scrolling while the button is held and the pointer stays parked at the edge
      if (!this.autoScroll) this.autoScroll = window.setInterval(() => {
        if (!this.drag || !this.lastPointer) { this.stopAutoScroll(); return; }
        const br = this.canvas.getBoundingClientRect();
        this.onMouseMove(new MouseEvent("mousemove", { clientX: br.left + this.lastPointer.px, clientY: br.top + this.lastPointer.py, ctrlKey: this.drag.kind === "move" ? !!this.drag.moveCopy : false }));
      }, 50);
    } else this.stopAutoScroll();
  }
  private stopAutoScroll() { if (this.autoScroll) { clearInterval(this.autoScroll); this.autoScroll = 0; } }

  private onMouseUp(e: MouseEvent) {
    const d = this.drag;
    this.stopAutoScroll();
    if (!d) return;
    this.drag = null;
    if (d.kind === "resizeCol" && this.previewSize) {
      const w = this.previewSize.size / this.zoom;
      const sel = this.selection.ranges[0];
      const wholeCols = sel.r1 === 0 && sel.r2 >= MAXR - 1 && d.index! >= sel.c1 && d.index! <= sel.c2;
      const all = wholeCols ? Array.from({ length: sel.c2 - sel.c1 + 1 }, (_, i) => sel.c1 + i) : [d.index!];
      this.previewSize = null;
      this.ev.onColResize(d.index!, w, all);
    } else if (d.kind === "resizeRow" && this.previewSize) {
      const h = this.previewSize.size / this.zoom;
      const sel = this.selection.ranges[0];
      const wholeRows = sel.c1 === 0 && sel.c2 >= MAXC - 1 && d.index! >= sel.r1 && d.index! <= sel.r2;
      const all = wholeRows ? Array.from({ length: sel.r2 - sel.r1 + 1 }, (_, i) => sel.r1 + i) : [d.index!];
      this.previewSize = null;
      this.ev.onRowResize(d.index!, h, all);
    } else if (d.kind === "fill" && d.fillTarget) {
      const src = this.selection.ranges[0];
      const t = d.fillTarget;
      if (t.r1 !== src.r1 || t.r2 !== src.r2 || t.c1 !== src.c1 || t.c2 !== src.c2) this.ev.onFill(src, t);
    } else if (d.kind === "move" && d.moveTarget) {
      const src = this.selection.ranges[0];
      const t = d.moveTarget;
      this.canvas.style.cursor = "cell";
      if (t.r1 !== src.r1 || t.c1 !== src.c1) this.ev.onMoveRange(src, t, !!d.moveCopy);
    }
    this.schedule();
    void e;
  }

  private onDblClick(e: MouseEvent) {
    const hit = this.hitTest(e);
    if (hit.type === "colResize") { this.ev.onAutoFit("col", hit.c); return; }
    if (hit.type === "rowResize") { this.ev.onAutoFit("row", hit.r); return; }
    // The pointer can drift a pixel or two off the edge between the two clicks; the press
    // that started this double-click knew which edge it was on, so trust that instead.
    if (this.lastResize && Date.now() - this.lastResize.at < 800) {
      this.ev.onAutoFit(this.lastResize.kind, this.lastResize.index);
      this.lastResize = null;
      return;
    }
    if (hit.type === "cell") { this.setActive(hit.r, hit.c); this.ev.onEdit(null); }
  }

  hyperlinkAt(r: number, c: number): Hyperlink | null {
    for (const h of this.sheet().hyperlinks) {
      const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(h.ref);
      if (!m) continue;
      const c1 = colIndexOf(m[1]), r1 = parseInt(m[2], 10) - 1;
      const c2 = m[3] ? colIndexOf(m[3]) : c1, r2 = m[4] ? parseInt(m[4], 10) - 1 : r1;
      if (r >= r1 && r <= r2 && c >= c1 && c <= c2) return h;
    }
    const cell = this.sheet().cells.get(key(r, c));
    if (cell && cell.f) {
      const m = /^HYPERLINK\(\s*"((?:[^"]|"")*)"/i.exec(cell.f);
      if (m) { const target = m[1].replace(/""/g, '"'); return /^#/.test(target) ? { ref: colName(c) + (r + 1), location: target.slice(1) } : { ref: colName(c) + (r + 1), target }; }
    }
    return null;
  }

  // ---- rendering ---------------------------------------------------------------

  private font(cs: CellStyle): string {
    const k = `${cs.italic ? 1 : 0}${cs.bold ? 1 : 0}|${cs.fontSize}|${cs.fontName}|${this.zoom}`;
    let f = this.fontCache.get(k);
    if (!f) {
      const px = cs.fontSize * this.zoom * 96 / 72;
      f = `${cs.italic ? "italic " : ""}${cs.bold ? "bold " : ""}${px.toFixed(2)}px ${fontFamilyCss(cs.fontName)}`;
      this.fontCache.set(k, f);
    }
    return f;
  }

  private displayText(cell: Cell, cs: CellStyle, styleIdx: number): { text: string; color: string | null; align: string | null } {
    const k = styleIdx + "|" + this.styles.version + (this.showFormulas ? "|f" : "");
    const cached = this.textCache.get(cell);
    if (cached && cached.key === k) return cached;
    const r = this.showFormulas && cell.f !== undefined ? { text: "=" + cell.f, color: null, align: "left" as const } : this.styles.render(cell, cs);
    const out = { key: k, text: r.text, color: r.color, align: r.align };
    this.textCache.set(cell, out);
    return out;
  }

  render() {
    const s = this.sheet();
    const ctx = this.ctx;
    const z = this.zoom, dpr = this.dpr;
    if (!this.canvas.width) this.resizeCanvas();

    const W = this.host.clientWidth, H = this.host.clientHeight;
    const hw = HEADER_W * z, hh = HEADER_H * z;
    const st = this.host.scrollTop, sl = this.host.scrollLeft;
    const fr = this.frozenRows(), fc = this.frozenCols();
    const fw = this.frozenW(), fh = this.frozenH();
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = this.colors.paper;
    ctx.fillRect(0, 0, W, H);

    // Visible row/col ranges for the scrolled region
    const r0 = this.rowIndexAt(st + fh), r1 = Math.min(this.nrows - 1, this.rowIndexAt(st + H - hh) + 1);
    const c0 = this.colIndexAt(sl + fw), c1 = Math.min(this.ncols - 1, this.colIndexAt(sl + W - hw) + 1);
    const regions: { rows: [number, number]; cols: [number, number]; clip: [number, number, number, number]; ox: number; oy: number }[] = [];
    // main
    regions.push({ rows: [Math.max(r0, fr), r1], cols: [Math.max(c0, fc), c1], clip: [hw + fw, hh + fh, W - hw - fw, H - hh - fh], ox: hw - sl, oy: hh - st });
    if (fr) regions.push({ rows: [0, fr - 1], cols: [Math.max(c0, fc), c1], clip: [hw + fw, hh, W - hw - fw, fh], ox: hw - sl, oy: hh });
    if (fc) regions.push({ rows: [Math.max(r0, fr), r1], cols: [0, fc - 1], clip: [hw, hh + fh, fw, H - hh - fh], ox: hw, oy: hh - st });
    if (fr && fc) regions.push({ rows: [0, fr - 1], cols: [0, fc - 1], clip: [hw, hh, fw, fh], ox: hw, oy: hh });

    for (const reg of regions) {
      ctx.save();
      ctx.beginPath(); ctx.rect(reg.clip[0], reg.clip[1], Math.max(0, reg.clip[2]), Math.max(0, reg.clip[3])); ctx.clip();
      this.drawCells(reg.rows, reg.cols, reg.ox, reg.oy);
      this.drawSelection(reg.ox, reg.oy);
      ctx.restore();
    }
    // frozen split lines
    if (fr || fc) {
      ctx.strokeStyle = this.colors.frozenLine; ctx.lineWidth = 1;
      if (fr) { ctx.beginPath(); ctx.moveTo(0, hh + fh + 0.5); ctx.lineTo(W, hh + fh + 0.5); ctx.stroke(); }
      if (fc) { ctx.beginPath(); ctx.moveTo(hw + fw + 0.5, 0); ctx.lineTo(hw + fw + 0.5, H); ctx.stroke(); }
    }
    this.drawHeaders(r0, r1, c0, c1);
    if (this.previewSize) {
      ctx.strokeStyle = this.colors.accent; ctx.setLineDash([4, 3]); ctx.lineWidth = 1;
      if (this.previewSize.kind === "col") { const x = this.cellRect(0, this.previewSize.index).x + this.previewSize.size; ctx.beginPath(); ctx.moveTo(x + 0.5, 0); ctx.lineTo(x + 0.5, H); ctx.stroke(); }
      else { const y = this.cellRect(this.previewSize.index, 0).y + this.previewSize.size; ctx.beginPath(); ctx.moveTo(0, y + 0.5); ctx.lineTo(W, y + 0.5); ctx.stroke(); }
      ctx.setLineDash([]);
    }
  }

  private drawHeaders(r0: number, r1: number, c0: number, c1: number) {
    const ctx = this.ctx, s = this.sheet(), z = this.zoom;
    const W = this.host.clientWidth, H = this.host.clientHeight;
    const hw = HEADER_W * z, hh = HEADER_H * z;
    const st = this.host.scrollTop, sl = this.host.scrollLeft;
    const fr = this.frozenRows(), fc = this.frozenCols(), fw = this.frozenW(), fh = this.frozenH();
    ctx.font = `${(11 * z).toFixed(1)}px ${fontFamilyCss("Segoe UI")}`;
    ctx.textBaseline = "middle";
    const sel = this.selection.ranges;
    const colSelected = (c: number) => sel.some((rg) => c >= rg.c1 && c <= rg.c2);
    const rowSelected = (r: number) => sel.some((rg) => r >= rg.r1 && r <= rg.r2);
    const fullColSel = (c: number) => sel.some((rg) => c >= rg.c1 && c <= rg.c2 && rg.r1 === 0 && rg.r2 >= MAXR - 1);
    const fullRowSel = (r: number) => sel.some((rg) => r >= rg.r1 && r <= rg.r2 && rg.c1 === 0 && rg.c2 >= MAXC - 1);
    // column headers
    ctx.fillStyle = this.colors.headerBg; ctx.fillRect(0, 0, W, hh);
    const drawColHeader = (c: number, ox: number) => {
      const x = ox + this.colStart(c), w = this.colWidthPx(c);
      if (w <= 0) return;
      if (colSelected(c)) { ctx.fillStyle = fullColSel(c) ? this.colors.accent : this.colors.headerSel; ctx.fillRect(x, 0, w, hh); }
      ctx.fillStyle = colSelected(c) && fullColSel(c) ? "#fff" : this.colors.headerText;
      ctx.textAlign = "center";
      const name = colName(c);
      if (w > 14) ctx.fillText(name, x + w / 2, hh / 2 + 1, w - 2);
      ctx.strokeStyle = this.colors.grid; ctx.beginPath(); ctx.moveTo(x + w + 0.5, 0); ctx.lineTo(x + w + 0.5, hh); ctx.stroke();
    };
    ctx.save(); ctx.beginPath(); ctx.rect(hw + fw, 0, W - hw - fw, hh); ctx.clip();
    for (let c = Math.max(c0, fc); c <= c1; c++) drawColHeader(c, hw - sl);
    ctx.restore();
    if (fc) { ctx.save(); ctx.beginPath(); ctx.rect(hw, 0, fw, hh); ctx.clip(); for (let c = 0; c < fc; c++) drawColHeader(c, hw); ctx.restore(); }
    // row headers
    ctx.fillStyle = this.colors.headerBg; ctx.fillRect(0, hh, hw, H - hh);
    const drawRowHeader = (r: number, oy: number) => {
      const y = oy + this.rowStart(r), h = this.rowHeightPxZ(r);
      if (h <= 0) return;
      if (rowSelected(r)) { ctx.fillStyle = fullRowSel(r) ? this.colors.accent : this.colors.headerSel; ctx.fillRect(0, y, hw, h); }
      ctx.fillStyle = rowSelected(r) && fullRowSel(r) ? "#fff" : this.colors.headerText;
      ctx.textAlign = "center";
      if (h >= 9) ctx.fillText(String(r + 1), hw / 2, y + h / 2 + 1, hw - 4);
      ctx.strokeStyle = this.colors.grid; ctx.beginPath(); ctx.moveTo(0, y + h + 0.5); ctx.lineTo(hw, y + h + 0.5); ctx.stroke();
    };
    ctx.save(); ctx.beginPath(); ctx.rect(0, hh + fh, hw, H - hh - fh); ctx.clip();
    for (let r = Math.max(r0, fr); r <= r1; r++) drawRowHeader(r, hh - st);
    ctx.restore();
    if (fr) { ctx.save(); ctx.beginPath(); ctx.rect(0, hh, hw, fh); ctx.clip(); for (let r = 0; r < fr; r++) drawRowHeader(r, hh); ctx.restore(); }
    // corner
    ctx.fillStyle = this.colors.headerBg; ctx.fillRect(0, 0, hw, hh);
    ctx.strokeStyle = this.colors.grid; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(0, hh + 0.5); ctx.lineTo(W, hh + 0.5); ctx.moveTo(hw + 0.5, 0); ctx.lineTo(hw + 0.5, H); ctx.stroke();
    void s;
  }

  private drawCells(rows: [number, number], cols: [number, number], ox: number, oy: number) {
    const ctx = this.ctx, s = this.sheet(), z = this.zoom;
    const [ra, rb] = rows, [ca, cb] = cols;
    if (ra > rb || ca > cb) return;
    // gridlines
    if (s.view.showGridLines) {
      ctx.strokeStyle = this.colors.grid; ctx.lineWidth = 1; ctx.beginPath();
      const x0 = ox + this.colStart(ca), x1 = ox + this.colStart(cb) + this.colWidthPx(cb);
      const y0 = oy + this.rowStart(ra), y1 = oy + this.rowStart(rb) + this.rowHeightPxZ(rb);
      for (let r = ra; r <= rb; r++) { const y = oy + this.rowStart(r) + this.rowHeightPxZ(r); ctx.moveTo(x0, y + 0.5); ctx.lineTo(x1, y + 0.5); }
      for (let c = ca; c <= cb; c++) { const x = ox + this.colStart(c) + this.colWidthPx(c); ctx.moveTo(x + 0.5, y0); ctx.lineTo(x + 0.5, y1); }
      ctx.stroke();
    }
    // Pass 1: fills (merged cells drawn from their anchor, possibly outside the visible window)
    const merges = s.merges.filter((m) => !(m.r2 < ra || m.r1 > rb || m.c2 < ca || m.c1 > cb));
    const inMergeNotAnchor = (r: number, c: number) => merges.some((m) => inRange(m, r, c) && !(m.r1 === r && m.c1 === c));
    const drawFill = (r: number, c: number, x: number, y: number, w: number, h: number, cell: Cell | undefined, cs: CellStyle) => {
      if (cs.fill) { ctx.fillStyle = cs.fill; ctx.fillRect(x, y, w + 1, h + 1); }
      void r; void c; void cell;
    };
    const cellBox = (r: number, c: number) => {
      const m = merges.find((mm) => mm.r1 === r && mm.c1 === c);
      const x = ox + this.colStart(c), y = oy + this.rowStart(r);
      if (m) return { x, y, w: ox + this.colStart(m.c2) + this.colWidthPx(m.c2) - x, h: oy + this.rowStart(m.r2) + this.rowHeightPxZ(m.r2) - y };
      return { x, y, w: this.colWidthPx(c), h: this.rowHeightPxZ(r) };
    };
    const styleOf = (r: number, c: number, cell: Cell | undefined): number => {
      if (cell) return cell.s;
      const ri = s.rows.get(r); if (ri?.style) return ri.style;
      for (const ci of s.cols) if (c >= ci.min && c <= ci.max && ci.style) return ci.style;
      return 0;
    };
    // merged anchors may start before the visible range
    const extraAnchors = merges.filter((m) => m.r1 < ra || m.c1 < ca).map((m) => ({ r: m.r1, c: m.c1 }));
    for (const { r, c } of extraAnchors) { const cell = s.cells.get(key(r, c)); const cs = this.styles.get(styleOf(r, c, cell)); const b = cellBox(r, c); drawFill(r, c, b.x, b.y, b.w, b.h, cell, cs); }
    for (let r = ra; r <= rb; r++) {
      const h = this.rowHeightPxZ(r); if (h <= 0) continue;
      for (let c = ca; c <= cb; c++) {
        const w = this.colWidthPx(c); if (w <= 0) continue;
        if (inMergeNotAnchor(r, c)) continue;
        const cell = s.cells.get(key(r, c));
        const si = styleOf(r, c, cell);
        if (!si && !cell) continue;
        const cs = this.styles.get(si);
        const b = cellBox(r, c);
        drawFill(r, c, b.x, b.y, b.w, b.h, cell, cs);
      }
    }
    // Pass 2: text
    ctx.textBaseline = "alphabetic";
    const overflowTargets: { x: number; y: number; w: number; h: number }[] = [];
    for (const { r, c } of extraAnchors) {
      const cell = s.cells.get(key(r, c));
      if (!cell || ((cell.v === null || cell.v === "") && !(this.showFormulas && cell.f !== undefined))) continue;
      const cs = this.styles.get(cell.s);
      const b = cellBox(r, c);
      this.drawText(r, c, cell, cs, b.x, b.y, b.w, b.h, overflowTargets);
    }
    for (let r = ra; r <= rb; r++) {
      const h = this.rowHeightPxZ(r); if (h <= 0) continue;
      for (let c = ca; c <= cb; c++) {
        if (inMergeNotAnchor(r, c)) continue;
        const cell = s.cells.get(key(r, c));
        if (!cell || ((cell.v === null || cell.v === "") && !(this.showFormulas && cell.f !== undefined))) continue;
        const w = this.colWidthPx(c); if (w <= 0) continue;
        const cs = this.styles.get(cell.s);
        const b = cellBox(r, c);
        this.drawText(r, c, cell, cs, b.x, b.y, b.w, b.h, overflowTargets);
      }
    }
    // Pass 3: borders (after text so they are not covered)
    for (const { r, c } of extraAnchors) {
      const cell = s.cells.get(key(r, c));
      const si = styleOf(r, c, cell);
      if (!si) continue;
      const bd = this.styles.get(si).borders;
      if (!bd.top && !bd.bottom && !bd.left && !bd.right) continue;
      const b = cellBox(r, c);
      this.drawBorders(bd, b.x, b.y, b.w, b.h);
    }
    for (let r = ra; r <= rb; r++) {
      const h = this.rowHeightPxZ(r); if (h <= 0) continue;
      for (let c = ca; c <= cb; c++) {
        const cell = s.cells.get(key(r, c));
        const si = styleOf(r, c, cell);
        if (!si) continue;
        const cs = this.styles.get(si);
        const bd = cs.borders;
        if (!bd.top && !bd.bottom && !bd.left && !bd.right) continue;
        const b = inMergeNotAnchor(r, c) ? null : cellBox(r, c);
        const x = b ? b.x : ox + this.colStart(c), y = b ? b.y : oy + this.rowStart(r);
        const w = b ? b.w : this.colWidthPx(c), hh2 = b ? b.h : h;
        this.drawBorders(bd, x, y, w, hh2);
      }
    }
    // filter buttons
    if (s.autoFilter && s.autoFilter.r1 >= ra && s.autoFilter.r1 <= rb) {
      const r = s.autoFilter.r1;
      for (let c = Math.max(ca, s.autoFilter.c1); c <= Math.min(cb, s.autoFilter.c2); c++) {
        const b = this.filterBtnRect(ox + this.colStart(c), oy + this.rowStart(r), this.colWidthPx(c), this.rowHeightPxZ(r));
        this.drawFilterButton(b.x, b.y, b.s, this.filteredCols(c));
      }
    }
    // page break lines
    if (this.pageBreaks && !this.hidePageBreaks) {
      ctx.strokeStyle = this.colors.accent; ctx.setLineDash([6, 4]); ctx.lineWidth = 1;
      for (const r of this.pageBreaks.rows) if (r >= ra && r <= rb + 1) { const y = oy + this.rowStart(r); ctx.beginPath(); ctx.moveTo(ox + this.colStart(ca), y + 0.5); ctx.lineTo(ox + this.colStart(cb) + this.colWidthPx(cb), y + 0.5); ctx.stroke(); }
      for (const c of this.pageBreaks.cols) if (c >= ca && c <= cb + 1) { const x = ox + this.colStart(c); ctx.beginPath(); ctx.moveTo(x + 0.5, oy + this.rowStart(ra)); ctx.lineTo(x + 0.5, oy + this.rowStart(rb) + this.rowHeightPxZ(rb)); ctx.stroke(); }
      ctx.setLineDash([]);
    }
  }

  /**
   * Where the filter button sits in a header cell (x/y = cell top-left, w/h = its size). The
   * button shrinks with the cell and is clamped inside it, so a narrow column or a short
   * header row still shows one instead of pushing it up into the row above.
   */
  private filterBtnRect(x: number, y: number, w: number, h: number): { x: number; y: number; s: number } {
    const s = Math.max(FILTER_BTN_MIN, Math.min(15 * this.zoom, w - 2, h - 2));
    return { x: Math.max(x + 1, x + w - s - 2), y: Math.max(y + 1, y + h - s - 2), s };
  }

  /** Funnel button over a header cell; filled in the accent colour while the column filters. */
  private drawFilterButton(x: number, y: number, s: number, filtered: boolean) {
    const ctx = this.ctx;
    const rad = Math.min(3, s / 4);
    ctx.beginPath();
    ctx.moveTo(x + rad, y);
    ctx.arcTo(x + s, y, x + s, y + s, rad);
    ctx.arcTo(x + s, y + s, x, y + s, rad);
    ctx.arcTo(x, y + s, x, y, rad);
    ctx.arcTo(x, y, x + s, y, rad);
    ctx.closePath();
    ctx.fillStyle = filtered ? this.colors.accent : this.colors.headerBg;
    ctx.fill();
    ctx.strokeStyle = filtered ? this.colors.accent : this.colors.frozenLine;
    ctx.lineWidth = 1;
    ctx.stroke();
    // Funnel drawn in a unit box so rim, cone and stem keep their proportions at any size.
    const px = (u: number) => x + u * s, py = (v: number) => y + v * s;
    ctx.fillStyle = filtered ? "#ffffff" : this.colors.headerText;
    ctx.beginPath();
    ctx.moveTo(px(0.17), py(0.25));
    ctx.lineTo(px(0.83), py(0.25));
    ctx.lineTo(px(0.58), py(0.53));
    ctx.lineTo(px(0.58), py(0.81));
    ctx.lineTo(px(0.42), py(0.71));
    ctx.lineTo(px(0.42), py(0.53));
    ctx.closePath();
    ctx.fill();
  }

  private drawBorders(bd: CellStyle["borders"], x: number, y: number, w: number, h: number) {
    const ctx = this.ctx;
    const side = (s: { style: string; color: string } | null, x1: number, y1: number, x2: number, y2: number) => {
      if (!s) return;
      const thick = /thick/.test(s.style) ? 3 : /medium|double/.test(s.style) ? 2 : 1;
      ctx.strokeStyle = s.color; ctx.lineWidth = thick;
      if (/dash/.test(s.style)) ctx.setLineDash([4, 2]); else if (/dot/.test(s.style)) ctx.setLineDash([1, 2]); else ctx.setLineDash([]);
      const off = thick % 2 ? 0.5 : 0;
      ctx.beginPath(); ctx.moveTo(x1 + off, y1 + off); ctx.lineTo(x2 + off, y2 + off); ctx.stroke();
      if (s.style === "double") { ctx.lineWidth = 1; ctx.strokeStyle = this.colors.paper; ctx.beginPath(); ctx.moveTo(x1 + off, y1 + off); ctx.lineTo(x2 + off, y2 + off); ctx.stroke(); }
    };
    side(bd.top, x, y, x + w, y);
    side(bd.bottom, x, y + h, x + w, y + h);
    side(bd.left, x, y, x, y + h);
    side(bd.right, x + w, y, x + w, y + h);
    ctx.setLineDash([]); ctx.lineWidth = 1;
  }

  private drawText(r: number, c: number, cell: Cell, cs: CellStyle, x: number, y: number, w: number, h: number, _ov: any[]) {
    const ctx = this.ctx, s = this.sheet(), z = this.zoom;
    const d = this.displayText(cell, cs, cell.s);
    let text = d.text;
    if (!text) return;
    ctx.font = this.font(cs);
    ctx.fillStyle = d.color || cs.color || this.colors.text;
    const pad = 3 * z;
    const isNum = typeof cell.v === "number" || typeof cell.v === "boolean" || isError(cell.v);
    let halign = cs.halign;
    if (halign === "general") halign = d.align === "right" ? "right" : d.align === "center" ? "center" : typeof cell.v === "number" ? "right" : "left";
    if (halign === "fill" || halign === "justify" || halign === "centerContinuous") halign = "left";
    const indent = cs.indent * 9 * z;
    const lines: string[] = [];
    const fontPx = cs.fontSize * z * 96 / 72;
    const lineH = Math.round(fontPx * 1.25);
    if (cs.wrap || text.includes("\n")) {
      for (const para of text.split("\n")) {
        if (!cs.wrap) { lines.push(para); continue; }
        // word wrap
        const words = para.split(/(\s+)/);
        let cur = "";
        for (const wd of words) {
          const trial = cur + wd;
          if (ctx.measureText(trial).width <= w - 2 * pad - indent || !cur) cur = trial; else { lines.push(cur.trimEnd()); cur = wd.trimStart(); }
        }
        lines.push(cur.trimEnd());
      }
    } else lines.push(text);
    // Horizontal overflow into empty neighbours (text only, single line)
    let clipX = x, clipW = w;
    let textW = ctx.measureText(lines[0]).width;
    if (!cs.wrap && lines.length === 1 && textW + 2 * pad + indent > w) {
      if (isNum) {
        // numbers never overflow: show ### like Excel
        if (typeof cell.v === "number") {
          text = "#"; while (ctx.measureText(text + "#").width <= w - 2 * pad && text.length < 20) text += "#";
          lines[0] = text; textW = ctx.measureText(text).width;
        }
      } else if (halign === "left" || halign === "center") {
        // extend right over empty cells
        let cc = c + 1, ext = 0;
        const merge = mergeAt(s, r, c);
        if (merge) cc = merge.c2 + 1;
        while (ext < textW + 2 * pad - w && cc < c + 60) {
          const nb = s.cells.get(key(r, cc));
          if (nb && nb.v !== null && nb.v !== "") break;
          if (mergeAt(s, r, cc)) break;
          ext += this.colWidthPx(cc); cc++;
        }
        clipW = w + ext;
        if (halign === "center") {
          // also extend left
          let lc = c - 1, lext = 0;
          while (lext < (textW + 2 * pad - w) / 2 && lc >= 0 && c - lc < 60) { const nb = s.cells.get(key(r, lc)); if ((nb && nb.v !== null && nb.v !== "") || mergeAt(s, r, lc)) break; lext += this.colWidthPx(lc); lc--; }
          clipX = x - lext; clipW = w + ext + lext;
        }
      } else if (halign === "right") {
        let lc = c - 1, lext = 0;
        while (lext < textW + 2 * pad - w && lc >= 0 && c - lc < 60) { const nb = s.cells.get(key(r, lc)); if ((nb && nb.v !== null && nb.v !== "") || mergeAt(s, r, lc)) break; lext += this.colWidthPx(lc); lc--; }
        clipX = x - lext; clipW = w + lext;
      }
    }
    ctx.save();
    ctx.beginPath(); ctx.rect(clipX, y, clipW, h); ctx.clip();
    const totalH = lines.length * lineH;
    let ty: number;
    if (cs.valign === "top") ty = y + pad / 2 + fontPx;
    else if (cs.valign === "center") ty = y + (h - totalH) / 2 + fontPx;
    else ty = y + h - totalH + fontPx - pad / 2 - (lineH - fontPx) / 2;
    ctx.textAlign = halign === "right" ? "right" : halign === "center" ? "center" : "left";
    const tx = halign === "right" ? clipX + clipW - pad - indent : halign === "center" ? clipX + clipW / 2 : clipX + pad + indent;
    for (const line of lines) {
      ctx.fillText(line, tx, ty);
      if (cs.underline || cs.strike) {
        const lw = ctx.measureText(line).width;
        const lx = halign === "right" ? tx - lw : halign === "center" ? tx - lw / 2 : tx;
        ctx.strokeStyle = ctx.fillStyle as string; ctx.lineWidth = Math.max(1, z);
        if (cs.underline) { ctx.beginPath(); ctx.moveTo(lx, ty + 2 * z); ctx.lineTo(lx + lw, ty + 2 * z); ctx.stroke(); }
        if (cs.strike) { ctx.beginPath(); ctx.moveTo(lx, ty - fontPx * 0.3); ctx.lineTo(lx + lw, ty - fontPx * 0.3); ctx.stroke(); }
      }
      ty += lineH;
      if (ty - fontPx > y + h) break;
    }
    ctx.restore();
  }

  private drawSelection(ox: number, oy: number) {
    const ctx = this.ctx, z = this.zoom;
    const sel = this.selection;
    // highlights (formula references)
    for (const hl of this.highlights) {
      if (hl.sheet && hl.sheet.toLowerCase() !== this.sheet().name.toLowerCase()) continue;
      const rg = hl.range;
      const x = ox + this.colStart(rg.c1), y = oy + this.rowStart(rg.r1);
      const w = ox + this.colStart(rg.c2) + this.colWidthPx(rg.c2) - x, h = oy + this.rowStart(rg.r2) + this.rowHeightPxZ(rg.r2) - y;
      ctx.fillStyle = hl.color + "22"; ctx.fillRect(x, y, w, h);
      ctx.strokeStyle = hl.color; ctx.lineWidth = 2; ctx.strokeRect(x + 1, y + 1, w - 2, h - 2);
    }
    if (this.clipRange) {
      const rg = this.clipRange;
      const x = ox + this.colStart(rg.c1), y = oy + this.rowStart(rg.r1);
      const w = ox + this.colStart(rg.c2) + this.colWidthPx(rg.c2) - x, h = oy + this.rowStart(rg.r2) + this.rowHeightPxZ(rg.r2) - y;
      ctx.save(); ctx.strokeStyle = this.colors.accent; ctx.lineWidth = 2; ctx.setLineDash([5, 3]); ctx.lineDashOffset = -this.dashOffset; ctx.strokeRect(x + 1, y + 1, w - 2, h - 2); ctx.restore();
      if (!this.dashTimer) this.dashTimer = window.setInterval(() => { if (!this.clipRange) { clearInterval(this.dashTimer); this.dashTimer = 0; return; } this.dashOffset = (this.dashOffset + 1) % 16; this.schedule(); }, 120);
    }
    const drawRange = (rg: Range, primary: boolean) => {
      const x = ox + this.colStart(rg.c1), y = oy + this.rowStart(rg.r1);
      const x2 = ox + this.colStart(Math.min(rg.c2, this.ncols - 1)) + this.colWidthPx(Math.min(rg.c2, this.ncols - 1));
      const y2 = oy + this.rowStart(Math.min(rg.r2, this.nrows - 1)) + this.rowHeightPxZ(Math.min(rg.r2, this.nrows - 1));
      const w = (rg.c2 >= this.ncols - 1 ? this.host.clientWidth * 4 : x2 - x), h = (rg.r2 >= this.nrows - 1 ? this.host.clientHeight * 4 : y2 - y);
      ctx.fillStyle = this.colors.selFill; ctx.fillRect(x, y, w, h);
      if (primary) {
        // active cell stays unfilled
        const a = sel.active;
        const am = mergeAt(this.sheet(), a.r, a.c);
        const ax = ox + this.colStart(a.c), ay = oy + this.rowStart(a.r);
        const aw = am ? ox + this.colStart(am.c2) + this.colWidthPx(am.c2) - ax : this.colWidthPx(a.c);
        const ah = am ? oy + this.rowStart(am.r2) + this.rowHeightPxZ(am.r2) - ay : this.rowHeightPxZ(a.r);
        if (rg.r1 !== rg.r2 || rg.c1 !== rg.c2) {
          ctx.save(); ctx.globalCompositeOperation = "destination-out"; ctx.fillStyle = "#000"; ctx.fillRect(ax, ay, aw, ah); ctx.restore();
          // re-fill any cell fill under the active cell
          const cell = this.sheet().cells.get(key(a.r, a.c));
          const cs = this.styles.get(cell ? cell.s : 0);
          ctx.fillStyle = cs.fill || this.colors.paper; ctx.fillRect(ax, ay, aw, ah);
          if (cell) this.drawText(a.r, a.c, cell, cs, ax, ay, aw, ah, []);
        }
        ctx.strokeStyle = this.colors.accent; ctx.lineWidth = 2 * Math.max(1, Math.min(z, 1.5));
        ctx.strokeRect(x + 1, y + 1, w - 2, h - 2);
        // fill handle
        const hs = FILL_HANDLE;
        ctx.fillStyle = this.colors.accent; ctx.fillRect(x + w - hs / 2 - 1, y + h - hs / 2 - 1, hs + 1, hs + 1);
        ctx.strokeStyle = this.colors.paper; ctx.lineWidth = 1; ctx.strokeRect(x + w - hs / 2 - 1.5, y + h - hs / 2 - 1.5, hs + 2, hs + 2);
      } else {
        ctx.strokeStyle = this.colors.accent; ctx.lineWidth = 1; ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
      }
    };
    sel.ranges.forEach((rg, i) => drawRange(rg, i === 0));
    if (this.drag?.kind === "fill" && this.drag.fillTarget) {
      const t = this.drag.fillTarget;
      const x = ox + this.colStart(t.c1), y = oy + this.rowStart(t.r1);
      const w = ox + this.colStart(t.c2) + this.colWidthPx(t.c2) - x, h = oy + this.rowStart(t.r2) + this.rowHeightPxZ(t.r2) - y;
      ctx.strokeStyle = this.colors.frozenLine; ctx.setLineDash([3, 3]); ctx.lineWidth = 1.5; ctx.strokeRect(x + 1, y + 1, w - 2, h - 2); ctx.setLineDash([]);
    }
    if (this.drag?.kind === "move" && this.drag.moveTarget) {
      const t = this.drag.moveTarget;
      const x = ox + this.colStart(t.c1), y = oy + this.rowStart(t.r1);
      const w = ox + this.colStart(t.c2) + this.colWidthPx(t.c2) - x, h = oy + this.rowStart(t.r2) + this.rowHeightPxZ(t.r2) - y;
      ctx.strokeStyle = this.colors.accent; ctx.lineWidth = 2; ctx.strokeRect(x + 1, y + 1, w - 2, h - 2);
    }
  }

  /** Pixel width needed to fit column c (for auto-fit). */
  measureColumn(c: number, maxRows = 5000): number {
    const s = this.sheet(); const ctx = this.ctx;
    let best = 0;
    let n = 0;
    for (let r = 0; r <= s.maxRow && n < maxRows; r++) {
      const cell = s.cells.get(key(r, c));
      if (!cell || cell.v === null || cell.v === "") continue;
      n++;
      const cs = this.styles.get(cell.s);
      const t = this.displayText(cell, cs, cell.s).text;
      ctx.font = this.font(cs);
      const wv = Math.max(...t.split("\n").map((l) => ctx.measureText(l).width));
      best = Math.max(best, wv + 8 * this.zoom + cs.indent * 9 * this.zoom);
    }
    return best / this.zoom;
  }

  /** Pixel height needed to fit row r (wrapped text). */
  measureRow(r: number): number {
    const s = this.sheet(); const ctx = this.ctx;
    let best = 0;
    for (let c = 0; c <= s.maxCol; c++) {
      const cell = s.cells.get(key(r, c));
      if (!cell || cell.v === null || cell.v === "") continue;
      const cs = this.styles.get(cell.s);
      const fontPx = cs.fontSize * 96 / 72;
      const lineH = Math.round(fontPx * 1.25);
      let lines = 1;
      const t = this.displayText(cell, cs, cell.s).text;
      if (cs.wrap) {
        ctx.font = this.font(cs);
        const w = (this.colX.length > c + 1 ? this.colWidthPx(c) / this.zoom : charsToPx(colWidthChars(s, c))) - 6;
        lines = 0;
        for (const para of t.split("\n")) { let cur = ""; let ln = 1; for (const wd of para.split(/(\s+)/)) { const trial = cur + wd; if (ctx.measureText(trial).width / this.zoom <= w || !cur) cur = trial; else { ln++; cur = wd.trimStart(); } } lines += ln; }
      } else lines = t.split("\n").length;
      best = Math.max(best, lines * lineH + 4);
    }
    return best || rowHeightPx(s, r);
  }
}

function colIndexOf(name: string): number { let n = 0; for (let i = 0; i < name.length; i++) n = n * 26 + (name.charCodeAt(i) - 64); return n - 1; }
