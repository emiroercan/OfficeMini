// In-cell editor overlay and formula bar. Both edit the same text; formulas get
// coloured reference highlights on the grid, F4 cycles $ anchors, arrow keys
// insert references while a formula is being typed (Google Sheets behaviour),
// and function names autocomplete.
import { Grid } from "./grid";
import { Range, parseRef, cellRef, rangeRef, rangeRefShort, quoteSheet, MAXR, MAXC } from "./model";
import { formulaRefs, tokenize } from "./formula/tokens";
import { FUNCTION_NAMES } from "./formula/engine";
import { el, showPopup, PopupHandle, closeAllPopups } from "../ui/widgets";
import { fontFamilyCss } from "./render-style";

export const REF_COLORS = ["#e0533d", "#3d7ee0", "#2ea043", "#b86ad9", "#e0a13d", "#20a2a8", "#d9489a", "#7a6ad9"];

export interface EditorHost {
  grid: Grid;
  commit(text: string, move: { dr: number; dc: number } | null): void;
  cancel(): void;
  /** An edit started (also from the formula bar): the app records which sheet it belongs to. */
  onBegin?(): void;
  /** Name of the sheet shown in the grid when it differs from the sheet being edited (references get qualified). */
  pointSheetName?(): string | null;
  /** Called on every keystroke so the app can preview / update status. */
  onInput?(text: string): void;
}

const FUNC_HELP: Record<string, string> = {
  SUM: "SUM(value1, [value2, ...])", AVERAGE: "AVERAGE(value1, [value2, ...])", COUNT: "COUNT(value1, [value2, ...])", COUNTA: "COUNTA(value1, ...)", MIN: "MIN(value1, ...)", MAX: "MAX(value1, ...)",
  IF: "IF(condition, value_if_true, [value_if_false])", IFS: "IFS(cond1, value1, [cond2, value2, ...])", IFERROR: "IFERROR(value, value_if_error)", AND: "AND(logical1, ...)", OR: "OR(logical1, ...)", NOT: "NOT(logical)",
  SUMIF: "SUMIF(range, criterion, [sum_range])", SUMIFS: "SUMIFS(sum_range, criteria_range1, criterion1, ...)", COUNTIF: "COUNTIF(range, criterion)", COUNTIFS: "COUNTIFS(range1, criterion1, ...)", AVERAGEIF: "AVERAGEIF(range, criterion, [average_range])",
  XLOOKUP: "XLOOKUP(search_key, lookup_range, result_range, [missing_value], [match_mode], [search_mode])", VLOOKUP: "VLOOKUP(search_key, range, index, [is_sorted])", HLOOKUP: "HLOOKUP(search_key, range, index, [is_sorted])", INDEX: "INDEX(reference, [row], [column])", MATCH: "MATCH(search_key, range, [search_type])",
  CONCAT: "CONCAT(value1, [value2, ...])", CONCATENATE: "CONCATENATE(string1, [string2, ...])", TEXTJOIN: "TEXTJOIN(delimiter, ignore_empty, text1, ...)", LEFT: "LEFT(string, [number_of_characters])", RIGHT: "RIGHT(string, [number_of_characters])", MID: "MID(string, starting_at, extract_length)",
  LEN: "LEN(text)", TRIM: "TRIM(text)", UPPER: "UPPER(text)", LOWER: "LOWER(text)", PROPER: "PROPER(text)", SUBSTITUTE: "SUBSTITUTE(text, search_for, replace_with, [occurrence])", FIND: "FIND(search_for, text, [starting_at])", SEARCH: "SEARCH(search_for, text, [starting_at])", TEXT: "TEXT(number, format)", VALUE: "VALUE(text)", SPLIT: "SPLIT(text, delimiter)", REPT: "REPT(text, times)",
  ROUND: "ROUND(value, [places])", ROUNDUP: "ROUNDUP(value, [places])", ROUNDDOWN: "ROUNDDOWN(value, [places])", INT: "INT(value)", ABS: "ABS(value)", MOD: "MOD(dividend, divisor)", POWER: "POWER(base, exponent)", SQRT: "SQRT(value)",
  TODAY: "TODAY()", NOW: "NOW()", DATE: "DATE(year, month, day)", YEAR: "YEAR(date)", MONTH: "MONTH(date)", DAY: "DAY(date)", EDATE: "EDATE(start_date, months)", EOMONTH: "EOMONTH(start_date, months)", DATEDIF: "DATEDIF(start_date, end_date, unit)", NETWORKDAYS: "NETWORKDAYS(start_date, end_date, [holidays])", WEEKDAY: "WEEKDAY(date, [type])",
  UNIQUE: "UNIQUE(range)", SUMPRODUCT: "SUMPRODUCT(array1, [array2, ...])", LET: "LET(name1, value1, ..., expression)", CHOOSE: "CHOOSE(index, choice1, [choice2, ...])",
};

/**
 * Finishing "=SUM(A1:A5" gives "=SUM(A1:A5)": every parenthesis left open is closed, and an
 * unterminated string is terminated first, the way Excel completes a half-typed formula.
 * Parentheses inside string literals are text, not nesting.
 */
export function closeOpenParens(text: string): string {
  if (!text.startsWith("=")) return text;
  let depth = 0, inStr = false;
  for (let i = 1; i < text.length; i++) {
    const ch = text[i];
    if (inStr) {
      if (ch === '"') { if (text[i + 1] === '"') i++; else inStr = false; }
      continue;
    }
    if (ch === '"') inStr = true;
    else if (ch === "(") depth++;
    else if (ch === ")") depth = Math.max(0, depth - 1);
  }
  return text + (inStr ? '"' : "") + ")".repeat(depth);
}

/**
 * A reference sitting just before the caret, which pointing replaces rather than appends to.
 * Whole-column (A:A) and whole-row (3:3) forms count: pointing at a header writes those, and
 * without them a second click would leave "=A:AH8" behind.
 */
const REF_BEFORE_CARET = /((?:'[^']*'|[A-Za-z0-9_.]+)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)$/;

export class CellEditor {
  el: HTMLTextAreaElement;
  fbar: HTMLTextAreaElement;
  active = false;
  /** Where the edit started; the formula bar edits this cell too. */
  pos: { r: number; c: number } | null = null;
  private host: EditorHost;
  private inFbar = false;
  private popup: PopupHandle | null = null;
  private popupItems: string[] = [];
  private popupSel = 0;
  /** Reference being pointed at with arrow keys / mouse while typing a formula. */
  private pointing: { start: number; end: number; anchor: { r: number; c: number }; cur: { r: number; c: number } } | null = null;
  private lastValue = "";
  /** Set by the app while a mouse-select happens during formula entry. */
  suppressBlur = false;

  constructor(host: EditorHost, editEl: HTMLTextAreaElement, fbarEl: HTMLTextAreaElement) {
    this.host = host; this.el = editEl; this.fbar = fbarEl;
    for (const [ta, isF] of [[editEl, false], [fbarEl, true]] as [HTMLTextAreaElement, boolean][]) {
      ta.addEventListener("keydown", (e) => this.onKey(e, isF));
      ta.addEventListener("input", () => this.onInput(isF));
      ta.addEventListener("blur", () => { if (this.suppressBlur) return; setTimeout(() => { if (this.active && document.activeElement !== this.el && document.activeElement !== this.fbar && !this.suppressBlur) this.finish(null, true); }, 0); });
      ta.addEventListener("mousedown", () => { this.pointing = null; });
      // Arrow keys move the selection *inside* an open completion list, so re-running the
      // completion on their keyup threw that selection away again a moment after every press -
      // the list appeared to snap back to the first entry. Only caret moves along the line
      // re-evaluate what is being completed.
      ta.addEventListener("keyup", (e) => {
        this.updateHighlights();
        const caretMoved = e.key === "Home" || e.key === "End" || e.key === "ArrowLeft" || e.key === "ArrowRight";
        if (this.popup && caretMoved) this.maybeAutocomplete();
      });
    }
    fbarEl.addEventListener("focus", () => { if (!this.active) this.begin(null, true); this.inFbar = true; });
    fbarEl.addEventListener("mouseup", () => this.updateHighlights());
  }

  get text(): string { return this.inFbar ? this.fbar.value : this.el.value; }

  private setText(v: string) { this.el.value = v; this.fbar.value = v; this.lastValue = v; }
  private current(): HTMLTextAreaElement { return this.inFbar ? this.fbar : this.el; }

  isFormula(): boolean { return this.active && this.text.startsWith("="); }

  /** Start editing the active cell. `seed` replaces the content (typing starts an edit); null keeps it. */
  begin(seed: string | null, fromFbar = false, initialText?: string) {
    const g = this.host.grid;
    const { r, c } = g.selection.active;
    this.pos = { r, c };
    this.active = true;
    this.inFbar = fromFbar;
    this.host.onBegin?.();
    const text = seed !== null ? seed : (initialText ?? this.fbar.value);
    this.setText(text);
    this.placeOverlay();
    this.el.style.display = "block";
    const ta = this.current();
    if (!fromFbar) ta.focus();
    const n = ta.value.length;
    ta.setSelectionRange(n, n);
    this.autosize();
    this.updateHighlights();
    if (seed !== null) this.maybeAutocomplete();
  }

  /** Position the overlay on the active cell (call after scroll/zoom). */
  placeOverlay() {
    if (!this.pos) return;
    const g = this.host.grid;
    const s = g.sheet();
    const cs = g.styles.get(s.cells.get(this.pos.r * MAXC + this.pos.c)?.s ?? 0);
    const rect = g.rangeRect(g.selection.ranges[0].r1 === this.pos.r && g.selection.ranges[0].c1 === this.pos.c ? g.selection.ranges[0] : { r1: this.pos.r, c1: this.pos.c, r2: this.pos.r, c2: this.pos.c });
    const hostRect = g.host.getBoundingClientRect();
    const wrap = this.el.parentElement!.getBoundingClientRect();
    const px = cs.fontSize * g.zoom * 96 / 72;
    this.el.style.font = `${cs.italic ? "italic " : ""}${cs.bold ? "bold " : ""}${px.toFixed(2)}px ${fontFamilyCss(cs.fontName)}`;
    this.el.style.left = (hostRect.left - wrap.left + rect.x - 1) + "px";
    this.el.style.top = (hostRect.top - wrap.top + rect.y - 1) + "px";
    this.el.style.minWidth = (rect.w + 2) + "px";
    this.el.style.minHeight = (rect.h + 2) + "px";
    this.el.style.width = (rect.w + 2) + "px";
    this.el.style.height = (rect.h + 2) + "px";
    this.el.style.textAlign = cs.halign === "right" || cs.halign === "center" ? cs.halign : "left";
    this.el.classList.toggle("wrap", cs.wrap);
    this.autosize();
  }

  private autosize() {
    if (!this.pos) return;
    const g = this.host.grid;
    const maxW = g.host.clientWidth - parseFloat(this.el.style.left) - 4;
    const maxH = g.host.clientHeight - parseFloat(this.el.style.top) - 4;
    this.el.style.height = "auto";
    this.el.style.width = this.el.style.minWidth;
    if (!this.el.classList.contains("wrap")) {
      const w = Math.min(maxW, Math.max(parseFloat(this.el.style.minWidth), this.el.scrollWidth + 8));
      this.el.style.width = w + "px";
    }
    const h = Math.min(maxH, Math.max(parseFloat(this.el.style.minHeight), this.el.scrollHeight + 2));
    this.el.style.height = h + "px";
  }

  /** Commit (move = where to go afterwards) or cancel. */
  finish(move: { dr: number; dc: number } | null, commit: boolean) {
    if (!this.active) return;
    const text = commit ? closeOpenParens(this.text) : this.text;
    this.active = false;
    this.pointing = null;
    this.closePopup();
    this.el.style.display = "none";
    this.host.grid.highlights = [];
    this.host.grid.schedule();
    this.inFbar = false;
    if (commit) this.host.commit(text, move); else this.host.cancel();
  }

  private onInput(isF: boolean) {
    this.inFbar = isF;
    const v = this.current().value;
    if (isF) this.el.value = v; else this.fbar.value = v;
    this.lastValue = v;
    this.pointing = null;
    this.autosize();
    this.updateHighlights();
    this.maybeAutocomplete();
    this.host.onInput?.(v);
  }

  /** Whether arrow keys should move the caret (text) or navigate cells (Sheets behaviour: navigate unless the edit started with F2/double-click). */
  arrowsNavigate = true;

  private onKey(e: KeyboardEvent, isF: boolean) {
    this.inFbar = isF;
    const ta = this.current();
    const mod = e.ctrlKey || e.metaKey;
    if (this.popup) {
      if (e.key === "ArrowDown") { e.preventDefault(); this.popupSel = (this.popupSel + 1) % this.popupItems.length; this.renderPopup(); return; }
      if (e.key === "ArrowUp") { e.preventDefault(); this.popupSel = (this.popupSel - 1 + this.popupItems.length) % this.popupItems.length; this.renderPopup(); return; }
      if (e.key === "Tab" || (e.key === "Enter" && !e.shiftKey && !mod)) { e.preventDefault(); this.acceptCompletion(); return; }
      if (e.key === "Escape") { e.preventDefault(); this.closePopup(); return; }
    }
    if (e.key === "Escape") { e.preventDefault(); this.finish(null, false); return; }
    if (e.key === "Enter") {
      if (e.altKey || (mod && !e.shiftKey)) { e.preventDefault(); this.insertText("\n"); return; }
      e.preventDefault();
      this.finish({ dr: e.shiftKey ? -1 : 1, dc: 0 }, true);
      return;
    }
    if (e.key === "Tab") { e.preventDefault(); this.finish({ dr: 0, dc: e.shiftKey ? -1 : 1 }, true); return; }
    if (e.key === "F4" && this.isFormula()) { e.preventDefault(); this.toggleAnchor(); return; }
    if (e.key === "F2") { e.preventDefault(); this.arrowsNavigate = false; return; }
    if (e.key.startsWith("Arrow") || e.key === "Home" || e.key === "End" || e.key === "PageUp" || e.key === "PageDown") {
      if (this.isFormula() && this.canPointArrows()) { e.preventDefault(); this.pointWithKeys(e); return; }
      // Inside a formula (past its start) arrows move the caret through the text - they neither
      // point at cells nor commit, so the formula can be edited like any other text.
      if (this.isFormula()) return;
      if (this.arrowsNavigate && !isF && !ta.value.includes("\n") && (e.key === "ArrowUp" || e.key === "ArrowDown" || ((e.key === "ArrowLeft" || e.key === "ArrowRight") && ta.selectionStart === ta.selectionEnd && (e.key === "ArrowLeft" ? ta.selectionStart === 0 : ta.selectionStart === ta.value.length)))) {
        e.preventDefault();
        const dr = e.key === "ArrowUp" ? -1 : e.key === "ArrowDown" ? 1 : 0, dc = e.key === "ArrowLeft" ? -1 : e.key === "ArrowRight" ? 1 : 0;
        this.finish({ dr, dc }, true);
        return;
      }
      return; // caret movement
    }
    if (mod && e.key.toLowerCase() === "z") { /* let the textarea handle undo inside the edit */ e.stopPropagation(); return; }
    if (mod && (e.key.toLowerCase() === "s" || e.key.toLowerCase() === "o" || e.key.toLowerCase() === "p" || e.key.toLowerCase() === "n")) { this.finish(null, true); return; }
    e.stopPropagation();
  }

  private insertText(t: string) {
    const ta = this.current();
    const s = ta.selectionStart, en = ta.selectionEnd;
    ta.value = ta.value.slice(0, s) + t + ta.value.slice(en);
    ta.setSelectionRange(s + t.length, s + t.length);
    this.onInput(this.inFbar);
  }

  // ---- reference pointing ------------------------------------------------------


  /**
   * Whether an arrow key should point at cells rather than move the caret. Only right at the start
   * of a formula (nothing typed after "=") or while a pointing gesture is already running - so
   * arrows select a cell when you begin a formula and, from then on, just edit the text. Mouse
   * pointing (canPoint) is unaffected and still works anywhere.
   */
  private canPointArrows(): boolean {
    if (this.pointing) return true;
    const ta = this.current();
    return /^=\s*$/.test(ta.value.slice(0, ta.selectionStart));
  }

  /** Can a reference be inserted at the caret (after an operator, "(", "," or at a ref we are already pointing at)? */
  private canPoint(): boolean {
    const ta = this.current();
    if (this.pointing) return true;
    const before = ta.value.slice(0, ta.selectionStart);
    return /[=(+\-*/^&,;<>:]\s*$/.test(before) || /[=(+\-*/^&,;<>:]\s*\$?[A-Za-z]{0,3}\$?\d*$/.test(before) && !/[)"']$/.test(before);
  }

  private pointWithKeys(e: KeyboardEvent) {
    const g = this.host.grid;
    const ta = this.current();
    if (!this.pointing) {
      // If the caret is right after a partially typed ref, replace it.
      const before = ta.value.slice(0, ta.selectionStart);
      const m = REF_BEFORE_CARET.exec(before);
      const start = m ? ta.selectionStart - m[0].length : ta.selectionStart;
      const base = this.pos!;
      this.pointing = { start, end: ta.selectionStart, anchor: { ...base }, cur: { ...base } };
    }
    const p = this.pointing;
    let { r, c } = p.cur;
    const dr = e.key === "ArrowUp" ? -1 : e.key === "ArrowDown" ? 1 : e.key === "PageDown" ? 20 : e.key === "PageUp" ? -20 : 0;
    const dc = e.key === "ArrowLeft" ? -1 : e.key === "ArrowRight" ? 1 : 0;
    if (e.key === "Home") c = 0; else if (e.key === "End") c = Math.max(0, g.sheet().maxCol);
    r = Math.max(0, Math.min(MAXR - 1, r + dr)); c = Math.max(0, Math.min(MAXC - 1, c + dc));
    p.cur = { r, c };
    if (!e.shiftKey) p.anchor = { r, c };
    const rg: Range = { r1: Math.min(p.anchor.r, r), c1: Math.min(p.anchor.c, c), r2: Math.max(p.anchor.r, r), c2: Math.max(p.anchor.c, c) };
    const other = this.host.pointSheetName?.() ?? null;
    this.setPointedRef((other ? quoteSheet(other) + "!" : "") + rangeRefShort(rg));
    g.ensureVisible(r, c);
  }

  /** Called by the app when the user clicks / drags cells while a formula is being edited. */
  pointWithMouse(range: Range, sheetName: string | null) {
    const ta = this.current();
    if (!this.pointing) {
      const before = ta.value.slice(0, ta.selectionStart);
      const m = REF_BEFORE_CARET.exec(before);
      const start = m ? ta.selectionStart - m[0].length : ta.selectionStart;
      this.pointing = { start, end: ta.selectionStart, anchor: { r: range.r1, c: range.c1 }, cur: { r: range.r2, c: range.c2 } };
    }
    this.setPointedRef((sheetName ? quoteSheet(sheetName) + "!" : "") + rangeRefShort(range));
  }
  /** After a mouse pointing gesture the next typed character continues after the reference. */
  endPointing() { if (this.pointing) { const ta = this.current(); ta.setSelectionRange(this.pointing.end, this.pointing.end); } }
  isPointing() { return !!this.pointing && this.isFormula(); }
  canPointNow() { return this.isFormula() && this.canPoint(); }

  private setPointedRef(ref: string) {
    const p = this.pointing!;
    const ta = this.current();
    const v = ta.value;
    ta.value = v.slice(0, p.start) + ref + v.slice(p.end);
    p.end = p.start + ref.length;
    ta.setSelectionRange(p.end, p.end);
    if (this.inFbar) this.el.value = ta.value; else this.fbar.value = ta.value;
    this.lastValue = ta.value;
    this.autosize();
    this.updateHighlights();
    this.host.onInput?.(ta.value);
  }

  /** F4: cycle $A$1 → A$1 → $A1 → A1 on the reference at the caret. */
  private toggleAnchor() {
    const ta = this.current();
    const v = ta.value;
    const pos = ta.selectionStart;
    const re = /\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(v))) {
      if (pos >= m.index && pos <= m.index + m[0].length) {
        const cycle = (ref: string) => {
          const mm = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(ref)!;
          const ac = !!mm[1], ar = !!mm[3];
          const next = ac && ar ? [false, true] : !ac && ar ? [true, false] : ac && !ar ? [false, false] : [true, true];
          return (next[0] ? "$" : "") + mm[2] + (next[1] ? "$" : "") + mm[4];
        };
        const out = m[0].split(":").map(cycle).join(":");
        ta.value = v.slice(0, m.index) + out + v.slice(m.index + m[0].length);
        const np = Math.min(pos, m.index + out.length);
        ta.setSelectionRange(np, np);
        this.onInput(this.inFbar);
        return;
      }
    }
  }

  // ---- highlights ---------------------------------------------------------------

  updateHighlights() {
    const g = this.host.grid;
    if (!this.isFormula()) { if (g.highlights.length) { g.highlights = []; g.schedule(); } return; }
    const refs = formulaRefs(this.text);
    const cur = g.sheet().name.toLowerCase();
    const hl: Grid["highlights"] = [];
    const seen = new Map<string, number>();
    let n = 0;
    for (const ref of refs) {
      const k = (ref.sheet || "") + "!" + ref.r1 + "," + ref.c1 + "," + ref.r2 + "," + ref.c2;
      let idx = seen.get(k);
      if (idx === undefined) { idx = n++; seen.set(k, idx); }
      if (ref.sheet && ref.sheet.toLowerCase() !== cur) continue;
      hl.push({ range: { r1: ref.r1, c1: ref.c1, r2: ref.r2, c2: ref.c2 }, color: REF_COLORS[idx % REF_COLORS.length] });
    }
    g.highlights = hl;
    g.schedule();
  }

  /** Colour for the reference at a formula position (for the formula bar colouring). */
  static refColors(formula: string): { start: number; end: number; color: string }[] {
    const out: { start: number; end: number; color: string }[] = [];
    if (!formula.startsWith("=")) return out;
    const seen = new Map<string, number>();
    let n = 0, pos = 1;
    for (const t of tokenize(formula.slice(1))) {
      const start = pos; pos += t.text.length;
      if (t.type !== "ref") continue;
      const rp = parseRef(t.text);
      if (!rp) continue;
      const k = (rp.sheet || "") + "!" + rp.r1 + "," + rp.c1 + "," + rp.r2 + "," + rp.c2;
      let idx = seen.get(k);
      if (idx === undefined) { idx = n++; seen.set(k, idx); }
      out.push({ start, end: pos, color: REF_COLORS[idx % REF_COLORS.length] });
    }
    return out;
  }

  // ---- autocomplete --------------------------------------------------------------

  private maybeAutocomplete() {
    if (!this.isFormula()) { this.closePopup(); return; }
    const ta = this.current();
    const before = ta.value.slice(0, ta.selectionStart);
    const m = /(?:^=|[(+\-*/^&,;<>=])\s*([A-Za-zÇĞİÖŞÜçğıöşü_][A-Za-z0-9ÇĞİÖŞÜçğıöşü_.]*)$/.exec(before);
    if (!m || m[1].length < 1) { this.closePopup(); this.showFunctionHint(); return; }
    const prefix = m[1].toUpperCase().replace(/İ/g, "I");
    const items = FUNCTION_NAMES.filter((f) => f.startsWith(prefix)).slice(0, 12);
    if (!items.length || (items.length === 1 && items[0] === prefix)) { this.closePopup(); this.showFunctionHint(); return; }
    // Keep the highlighted entry when the list itself has not changed, so re-evaluating the
    // completion never moves the selection out from under the arrow keys.
    const same = items.length === this.popupItems.length && items.every((f, i) => f === this.popupItems[i]);
    this.popupItems = items;
    if (!same) this.popupSel = 0;
    else this.popupSel = Math.min(this.popupSel, items.length - 1);
    this.renderPopup();
  }

  private renderPopup() {
    const ta = this.current();
    const anchor = this.inFbar ? this.fbar : this.el;
    const build = (popup: HTMLElement) => {
      popup.innerHTML = "";
      popup.classList.add("fn-popup");
      this.popupItems.forEach((name, i) => {
        const row = el("div", { class: "item" + (i === this.popupSel ? " sel" : "") }, name, el("small", null, (FUNC_HELP[name] || "").replace(/^[A-Z0-9._]+/, "")));
        row.addEventListener("mousedown", (e) => { e.preventDefault(); this.popupSel = i; this.acceptCompletion(); });
        popup.appendChild(row);
      });
    };
    if (this.popup && document.body.contains(this.popup.el)) { build(this.popup.el); return; }
    this.popup = showPopup(anchor, (popup) => build(popup), { keepOthers: true, onClose: () => { this.popup = null; } });
    void ta;
  }

  private acceptCompletion() {
    const name = this.popupItems[this.popupSel];
    if (!name) return;
    const ta = this.current();
    const before = ta.value.slice(0, ta.selectionStart);
    const m = /([A-Za-zÇĞİÖŞÜçğıöşü_][A-Za-z0-9ÇĞİÖŞÜçğıöşü_.]*)$/.exec(before);
    if (!m) { this.closePopup(); return; }
    const start = ta.selectionStart - m[1].length;
    const after = ta.value.slice(ta.selectionEnd);
    const ins = name + "(";
    ta.value = ta.value.slice(0, start) + ins + after;
    ta.setSelectionRange(start + ins.length, start + ins.length);
    this.closePopup();
    this.onInput(this.inFbar);
    this.showFunctionHint();
  }

  private hint: HTMLElement | null = null;
  private showFunctionHint() {
    this.hideHint();
    if (!this.isFormula()) return;
    const ta = this.current();
    const before = ta.value.slice(0, ta.selectionStart);
    // find the innermost unclosed function call
    let depth = 0, name: string | null = null;
    for (let i = before.length - 1; i >= 0; i--) {
      const ch = before[i];
      if (ch === ")") depth++;
      else if (ch === "(") { if (depth === 0) { const mm = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before.slice(0, i)); if (mm) { name = mm[1].toUpperCase(); break; } } else depth--; }
    }
    if (!name || !FUNC_HELP[name]) return;
    const anchor = (this.inFbar ? this.fbar : this.el).getBoundingClientRect();
    this.hint = el("div", { class: "fn-hint" }, FUNC_HELP[name]);
    document.body.appendChild(this.hint);
    this.hint.style.left = anchor.left + "px";
    this.hint.style.top = (anchor.bottom + 2) + "px";
  }
  private hideHint() { if (this.hint) { this.hint.remove(); this.hint = null; } }

  closePopup() { if (this.popup) { const p = this.popup; this.popup = null; p.close(); } this.hideHint(); }

  destroy() { this.closePopup(); closeAllPopups(); }

  /** Convenience for the app: current formula text's referenced cell under the caret (for status). */
  static describeRef(text: string): string | null {
    const m = /([A-Z]+\d+(?::[A-Z]+\d+)?)/.exec(text);
    return m ? cellRef(0, 0) && m[1] : null;
  }
}
