// Data blocks, filtering (unique-value lists with search, like Sheets/Excel)
// and helpers shared by sort / pivot.
import { Sheet, Range, key, MAXR, MAXC, Cell, isError } from "./model";
import { StyleResolver } from "./render-style";
import { el, showPopup, PopupHandle } from "../ui/widgets";
import { parseLocaleNumber } from "./numfmt";

/** Contiguous block of non-empty cells around (r, c) — what Ctrl+A / filter / sort work on. */
export function dataBlock(sheet: Sheet, r: number, c: number): Range {
  const has = (rr: number, cc: number) => { const cell = sheet.cells.get(key(rr, cc)); return !!cell && cell.v !== null && cell.v !== ""; };
  let r1 = r, r2 = r, c1 = c, c2 = c;
  const maxR = Math.max(sheet.maxRow, 0), maxC = Math.max(sheet.maxCol, 0);
  if (!has(r, c)) {
    // start from a neighbouring filled cell if there is one
    const nb = [[r, c - 1], [r, c + 1], [r - 1, c], [r + 1, c]].find(([rr, cc]) => rr >= 0 && cc >= 0 && has(rr, cc));
    if (!nb) return { r1: r, c1: c, r2: r, c2: c };
    r1 = r2 = nb[0]; c1 = c2 = nb[1];
  }
  let changed = true, guard = 0;
  while (changed && guard++ < 100000) {
    changed = false;
    const rowHas = (rr: number) => { for (let cc = c1; cc <= c2; cc++) if (has(rr, cc)) return true; return false; };
    const colHas = (cc: number) => { for (let rr = r1; rr <= r2; rr++) if (has(rr, cc)) return true; return false; };
    while (r1 > 0 && rowHas(r1 - 1)) { r1--; changed = true; }
    while (r2 < maxR && rowHas(r2 + 1)) { r2++; changed = true; }
    while (c1 > 0 && colHas(c1 - 1)) { c1--; changed = true; }
    while (c2 < maxC && colHas(c2 + 1)) { c2++; changed = true; }
  }
  return { r1, c1, r2, c2 };
}

/** Does the first row of the range look like a header (text above non-text)? */
export function looksLikeHeader(sheet: Sheet, rg: Range): boolean {
  if (rg.r2 <= rg.r1) return false;
  let textTop = 0, cols = 0, differs = 0;
  for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 50); c++) {
    const top = sheet.cells.get(key(rg.r1, c)), next = sheet.cells.get(key(rg.r1 + 1, c));
    if (!top || top.v === null) continue;
    cols++;
    if (typeof top.v === "string" && !top.f) textTop++;
    if (next && next.v !== null && typeof next.v !== typeof top.v) differs++;
  }
  return cols > 0 && textTop === cols && (differs > 0 || rg.r2 - rg.r1 > 2);
}

/** Conditions offered above the value list; the empty op means "no condition". */
export type CondOp =
  | "" | "contains" | "notContains" | "startsWith" | "endsWith" | "isExactly"
  | "isEmpty" | "isNotEmpty"
  | "gt" | "gte" | "lt" | "lte" | "eq" | "neq" | "between"
  | "formula";

export interface Condition { op: CondOp; v1: string; v2: string; }

export const CONDITIONS: { op: CondOp; label: string; args: 0 | 1 | 2 }[] = [
  { op: "", label: "No condition", args: 0 },
  { op: "isEmpty", label: "Is empty", args: 0 },
  { op: "isNotEmpty", label: "Is not empty", args: 0 },
  { op: "contains", label: "Text contains", args: 1 },
  { op: "notContains", label: "Text does not contain", args: 1 },
  { op: "startsWith", label: "Text starts with", args: 1 },
  { op: "endsWith", label: "Text ends with", args: 1 },
  { op: "isExactly", label: "Text is exactly", args: 1 },
  { op: "gt", label: "Greater than", args: 1 },
  { op: "gte", label: "Greater than or equal to", args: 1 },
  { op: "lt", label: "Less than", args: 1 },
  { op: "lte", label: "Less than or equal to", args: 1 },
  { op: "eq", label: "Is equal to", args: 1 },
  { op: "neq", label: "Is not equal to", args: 1 },
  { op: "between", label: "Is between", args: 2 },
  { op: "formula", label: "Custom formula is", args: 1 },
];

// null values = every value passes; null cond = no condition. Both apply together.
export interface ColumnFilter { values: Set<string> | null; cond?: Condition | null; }
export type FilterState = Map<number, ColumnFilter>;

/**
 * Evaluates a custom-formula condition for one row. Supplied by the app (which owns the
 * formula engine) so this module stays free of it; missing means custom formulas never match.
 */
export type FormulaRowTest = (formula: string, row: number) => boolean;

function condMatches(cond: Condition, cell: Cell | undefined, text: string, row: number, test?: FormulaRowTest): boolean {
  const num = (s: string) => parseLocaleNumber(s.trim());
  const cellNum = cell && typeof cell.v === "number" ? cell.v : null;
  const fold = (s: string) => s.toLocaleLowerCase("tr").replace(/i̇/g, "i");
  const t = fold(text), a = fold(cond.v1 || "");
  switch (cond.op) {
    case "isEmpty": return text === "";
    case "isNotEmpty": return text !== "";
    case "contains": return t.includes(a);
    case "notContains": return !t.includes(a);
    case "startsWith": return t.startsWith(a);
    case "endsWith": return t.endsWith(a);
    case "isExactly": return t === a;
    case "formula": return test ? test(cond.v1, row) : false;
    default: break;
  }
  // Numeric comparisons fall back to comparing the displayed text when either side is not a number.
  const b1 = num(cond.v1);
  if (cellNum === null || b1 === null) {
    if (cond.op === "eq") return t === a;
    if (cond.op === "neq") return t !== a;
    return false;
  }
  switch (cond.op) {
    case "gt": return cellNum > b1;
    case "gte": return cellNum >= b1;
    case "lt": return cellNum < b1;
    case "lte": return cellNum <= b1;
    case "eq": return cellNum === b1;
    case "neq": return cellNum !== b1;
    case "between": { const b2 = num(cond.v2); return b2 === null ? false : cellNum >= Math.min(b1, b2) && cellNum <= Math.max(b1, b2); }
    default: return true;
  }
}

function isActive(f: ColumnFilter): boolean { return f.values !== null || !!(f.cond && f.cond.op); }

const filterStates = new WeakMap<Sheet, FilterState>();
export function filterState(sheet: Sheet): FilterState { let s = filterStates.get(sheet); if (!s) { s = new Map(); filterStates.set(sheet, s); } return s; }
export function clearFilterState(sheet: Sheet) { filterStates.delete(sheet); sheet.hiddenRowsByFilter = new Set(); }

/**
 * Bottom row a filter actually covers. A stored autoFilter range (from an .xlsx, or set before
 * rows were appended) can end above the real data - Excel still filters the whole contiguous
 * block, and so must we, or every row past the stored ref stays visible however it is filtered.
 * Extends the range's bottom down through non-blank rows, stopping at the first fully-blank one.
 */
function filterBottom(sheet: Sheet, af: Range): number {
  let r2 = Math.min(af.r2, MAXR - 1);
  const maxR = Math.max(sheet.maxRow, af.r1);
  while (r2 < maxR) {
    let blank = true;
    for (let c = af.c1; c <= af.c2; c++) { const cell = sheet.cells.get(key(r2 + 1, c)); if (cell && cell.v !== null && cell.v !== "") { blank = false; break; } }
    if (blank) break;
    r2++;
  }
  return r2;
}

/** Recompute the set of rows hidden by the sheet's filter criteria. */
export function applyFilters(sheet: Sheet, styles: StyleResolver, test?: FormulaRowTest) {
  const hidden = new Set<number>();
  const af = sheet.autoFilter;
  const st = filterStates.get(sheet);
  if (af && st) {
    const active = Array.from(st.entries()).filter(([, f]) => isActive(f));
    if (active.length) {
      const r2 = filterBottom(sheet, af);
      // A totally empty row (a freshly inserted one) stays visible so it can be typed into. This
      // must look at the WHOLE row, not just the filtered columns: a single-column filter would
      // otherwise treat every existing row whose filtered cell is blank as "empty" and leak it
      // through - blanks showing under a value filter that excludes them.
      const bc2 = Math.max(af.c2, sheet.maxCol);
      for (let r = af.r1 + 1; r <= r2; r++) {
        let blank = true;
        for (let c = 0; c <= bc2; c++) { const cell = sheet.cells.get(key(r, c)); if (cell && cell.v !== null && cell.v !== "") { blank = false; break; } }
        if (blank) continue;
        for (const [c, f] of active) {
          const cell = sheet.cells.get(key(r, c));
          const text = displayOf(cell, styles);
          if (f.values && !f.values.has(text)) { hidden.add(r); break; }
          if (f.cond && f.cond.op && !condMatches(f.cond, cell, text, r, test)) { hidden.add(r); break; }
        }
      }
    }
  }
  sheet.hiddenRowsByFilter = hidden;
}

export function displayOf(cell: Cell | undefined, styles: StyleResolver): string {
  if (!cell || cell.v === null) return "";
  return styles.render(cell, styles.get(cell.s)).text;
}

/** Unique displayed values of a filter column (rows hidden by *other* columns excluded, like Excel). */
export function uniqueValues(sheet: Sheet, af: Range, col: number, styles: StyleResolver, test?: FormulaRowTest): { text: string; count: number; sortKey: number | string }[] {
  const st = filterStates.get(sheet);
  const others = st ? Array.from(st.entries()).filter(([c, f]) => c !== col && isActive(f)) : [];
  const counts = new Map<string, { count: number; sortKey: number | string }>();
  const r2 = filterBottom(sheet, af);
  for (let r = af.r1 + 1; r <= r2; r++) {
    let skip = false;
    for (const [c, f] of others) {
      const other = sheet.cells.get(key(r, c));
      const otherText = displayOf(other, styles);
      if (f.values && !f.values.has(otherText)) { skip = true; break; }
      if (f.cond && f.cond.op && !condMatches(f.cond, other, otherText, r, test)) { skip = true; break; }
    }
    if (skip) continue;
    const cell = sheet.cells.get(key(r, col));
    const text = displayOf(cell, styles);
    const e = counts.get(text);
    if (e) e.count++;
    else counts.set(text, { count: 1, sortKey: cell && typeof cell.v === "number" ? cell.v : cell && isError(cell.v) ? "￿" + text : text.toLocaleLowerCase() });
  }
  const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });
  return Array.from(counts.entries()).map(([text, e]) => ({ text, ...e })).sort((a, b) => {
    if (a.text === "" && b.text !== "") return 1; if (b.text === "" && a.text !== "") return -1;
    if (typeof a.sortKey === "number" && typeof b.sortKey === "number") return a.sortKey - b.sortKey;
    if (typeof a.sortKey === "number") return -1; if (typeof b.sortKey === "number") return 1;
    return collator.compare(a.sortKey as string, b.sortKey as string);
  });
}

export interface FilterPopupCtx {
  sheet: Sheet; col: number; styles: StyleResolver; x: number; y: number;
  header: string;
  /** Row a custom formula is written against (the first data row), for the hint text. */
  firstDataRow: number;
  test?: FormulaRowTest;
  onApply(): void;
  onSort(ascending: boolean): void;
}

/** Excel-style filter dropdown: sort buttons, search, (de)select all, value checkboxes with counts. */
export function showFilterPopup(ctx: FilterPopupCtx): PopupHandle {
  const { sheet, col, styles } = ctx;
  const af = sheet.autoFilter!;
  const st = filterState(sheet);
  const cur: ColumnFilter = st.get(col) || { values: null, cond: null };
  const values = uniqueValues(sheet, af, col, styles, ctx.test);
  const cond: Condition = { op: "", v1: "", v2: "", ...(cur.cond || {}) };
  const selected = new Set<string>(cur.values ? cur.values : values.map((v) => v.text));
  return showPopup({ x: ctx.x, y: ctx.y }, (popup, close) => {
    popup.classList.add("filter-popup");
    popup.append(el("div", { class: "item", style: { fontWeight: "600", padding: "2px 4px 6px" } }, ctx.header ? `Filter: ${ctx.header}` : "Filter"));
    const sortA = el("div", { class: "item" }, "Sort A → Z");
    const sortZ = el("div", { class: "item" }, "Sort Z → A");
    sortA.addEventListener("click", () => { close(); ctx.onSort(true); });
    sortZ.addEventListener("click", () => { close(); ctx.onSort(false); });
    popup.append(sortA, sortZ, el("div", { class: "menu-sep" }));

    // ---- filter by condition
    const opSel = el("select", null, ...CONDITIONS.map((c) => el("option", { value: c.op, selected: c.op === cond.op ? "selected" : null }, c.label)));
    const v1 = el("input", { type: "text", value: cond.v1, placeholder: "Value" });
    const andLbl = el("span", { style: { color: "var(--ui-muted)" } }, "and");
    const v2 = el("input", { type: "text", value: cond.v2, placeholder: "Value" });
    const condHint = el("div", { style: { color: "var(--ui-muted)", fontSize: "12px", padding: "2px 0 0" } }, "");
    const argRow = el("div", { class: "row cond-args" }, v1, andLbl, v2);
    const syncCond = () => {
      const def = CONDITIONS.find((c) => c.op === opSel.value)!;
      argRow.style.display = def.args ? "flex" : "none";
      andLbl.style.display = v2.style.display = def.args === 2 ? "" : "none";
      v1.placeholder = def.op === "formula" ? "=$B2>100" : "Value";
      condHint.textContent = def.op === "formula"
        ? `Written for row ${ctx.firstDataRow + 1} and shifted down each row; a TRUE result keeps the row.`
        : "";
      condHint.style.display = def.op === "formula" ? "" : "none";
    };
    opSel.addEventListener("change", syncCond);
    popup.append(el("div", { class: "cond-label" }, "Filter by condition"), opSel, argRow, condHint, el("div", { class: "menu-sep" }));
    syncCond();

    const search = el("input", { type: "text", placeholder: "Search values" });
    popup.appendChild(search);
    const list = el("div", { class: "list" });
    const rows: { text: string; label: HTMLLabelElement; cb: HTMLInputElement }[] = [];
    const render = () => {
      list.innerHTML = "";
      const q = search.value.toLocaleLowerCase();
      let shown = 0;
      for (const v of values) {
        if (q && !v.text.toLocaleLowerCase().includes(q)) continue;
        shown++;
        if (shown > 1000) { list.appendChild(el("div", { class: "item", style: { color: "var(--ui-muted)" } }, `… ${values.length - 1000} more (use search)`)); break; }
        const cb = el("input", { type: "checkbox" });
        cb.checked = selected.has(v.text);
        cb.addEventListener("change", () => { beforeSearch = null; if (cb.checked) selected.add(v.text); else selected.delete(v.text); });
        const label = el("label", null, cb, el("span", { style: { flex: "1", overflow: "hidden", textOverflow: "ellipsis" } }, v.text === "" ? "(Blanks)" : v.text), el("span", { style: { color: "var(--ui-muted)" } }, String(v.count)));
        list.appendChild(label);
        rows.push({ text: v.text, label, cb });
      }
      if (!shown) list.appendChild(el("div", { class: "item", style: { color: "var(--ui-muted)" } }, "No matches"));
    };
    const allRow = el("div", { class: "row", style: { padding: "4px 0", gap: "12px" } });
    const selAll = el("a", { href: "#", style: { color: "var(--ui-accent)" } }, "Select all");
    const selNone = el("a", { href: "#", style: { color: "var(--ui-accent)" } }, "Clear");
    selAll.addEventListener("click", (e) => { e.preventDefault(); beforeSearch = null; const q = search.value.toLocaleLowerCase(); for (const v of values) if (!q || v.text.toLocaleLowerCase().includes(q)) selected.add(v.text); render(); });
    selNone.addEventListener("click", (e) => { e.preventDefault(); beforeSearch = null; const q = search.value.toLocaleLowerCase(); for (const v of values) if (!q || v.text.toLocaleLowerCase().includes(q)) selected.delete(v.text); render(); });
    allRow.append(selAll, selNone, el("span", { style: { flex: "1" } }), el("span", { style: { color: "var(--ui-muted)" } }, `${values.length} values`));
    popup.append(allRow, list);
    // Typing a search term selects exactly the matches (and nothing else); clearing the box
    // puts back the selection the search started from, so a mistyped search costs nothing.
    let beforeSearch: Set<string> | null = null;
    search.addEventListener("input", () => {
      const q = search.value.toLocaleLowerCase();
      if (q) {
        if (!beforeSearch) beforeSearch = new Set(selected);
        selected.clear();
        for (const v of values) if (v.text.toLocaleLowerCase().includes(q)) selected.add(v.text);
      } else if (beforeSearch) {
        selected.clear();
        for (const t of beforeSearch) selected.add(t);
        beforeSearch = null;
      }
      render();
    });
    search.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { beforeSearch = null; apply(); }   // the matches are already the selection
      else if (e.key === "Escape") { e.stopPropagation(); close(); }
    });
    render();
    const apply = () => {
      const all = selected.size >= values.length && values.every((v) => selected.has(v.text));
      const c: Condition | null = opSel.value ? { op: opSel.value as CondOp, v1: v1.value, v2: v2.value } : null;
      if (all && !c) st.delete(col);
      else st.set(col, { values: all ? null : new Set(selected), cond: c });
      applyFilters(sheet, styles, ctx.test);
      close();
      ctx.onApply();
    };
    const ok = el("button", { class: "primary" }, "OK");
    const cancel = el("button", null, "Cancel");
    const clearBtn = el("button", null, "Clear filter");
    ok.addEventListener("click", apply);
    cancel.addEventListener("click", () => close());
    clearBtn.addEventListener("click", () => { st.delete(col); applyFilters(sheet, styles, ctx.test); close(); ctx.onApply(); });
    popup.appendChild(el("div", { class: "row", style: { justifyContent: "flex-end", padding: "6px 0 0" } }, clearBtn, el("span", { style: { flex: "1" } }), cancel, ok));
    setTimeout(() => search.focus(), 0);
  });
}

/** Whether the column has an active filter (for drawing the funnel icon differently). */
export function columnFiltered(sheet: Sheet, col: number): boolean {
  const st = filterStates.get(sheet);
  const f = st?.get(col);
  return !!f && isActive(f);
}

export { MAXR, MAXC };
