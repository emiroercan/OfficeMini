// Data blocks, filtering (unique-value lists with search, like Sheets/Excel)
// and helpers shared by sort / pivot.
import { Sheet, Range, key, MAXR, MAXC, Cell, isError } from "./model";
import { StyleResolver } from "./render-style";
import { el, showPopup, PopupHandle } from "../ui/widgets";

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

export interface ColumnFilter { values: Set<string> | null; }  // null = no filter on the column
export type FilterState = Map<number, ColumnFilter>;

const filterStates = new WeakMap<Sheet, FilterState>();
export function filterState(sheet: Sheet): FilterState { let s = filterStates.get(sheet); if (!s) { s = new Map(); filterStates.set(sheet, s); } return s; }
export function clearFilterState(sheet: Sheet) { filterStates.delete(sheet); sheet.hiddenRowsByFilter = new Set(); }

/** Recompute the set of rows hidden by the sheet's filter criteria. */
export function applyFilters(sheet: Sheet, styles: StyleResolver) {
  const hidden = new Set<number>();
  const af = sheet.autoFilter;
  const st = filterStates.get(sheet);
  if (af && st) {
    const active = Array.from(st.entries()).filter(([, f]) => f.values !== null);
    if (active.length) {
      const r2 = Math.min(af.r2, Math.max(sheet.maxRow, af.r1));
      for (let r = af.r1 + 1; r <= r2; r++) {
        // Rows that are completely empty inside the filtered columns (freshly inserted rows) stay visible.
        let blank = true;
        for (let c = af.c1; c <= af.c2; c++) { const cell = sheet.cells.get(key(r, c)); if (cell && cell.v !== null && cell.v !== "") { blank = false; break; } }
        if (blank) continue;
        for (const [c, f] of active) {
          const cell = sheet.cells.get(key(r, c));
          const text = displayOf(cell, styles);
          if (!f.values!.has(text)) { hidden.add(r); break; }
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
export function uniqueValues(sheet: Sheet, af: Range, col: number, styles: StyleResolver): { text: string; count: number; sortKey: number | string }[] {
  const st = filterStates.get(sheet);
  const others = st ? Array.from(st.entries()).filter(([c, f]) => c !== col && f.values !== null) : [];
  const counts = new Map<string, { count: number; sortKey: number | string }>();
  const r2 = Math.min(af.r2, Math.max(sheet.maxRow, af.r1));
  for (let r = af.r1 + 1; r <= r2; r++) {
    let skip = false;
    for (const [c, f] of others) { if (!f.values!.has(displayOf(sheet.cells.get(key(r, c)), styles))) { skip = true; break; } }
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
  onApply(): void;
  onSort(ascending: boolean): void;
}

/** Excel-style filter dropdown: sort buttons, search, (de)select all, value checkboxes with counts. */
export function showFilterPopup(ctx: FilterPopupCtx): PopupHandle {
  const { sheet, col, styles } = ctx;
  const af = sheet.autoFilter!;
  const st = filterState(sheet);
  const cur = st.get(col) || { values: null };
  const values = uniqueValues(sheet, af, col, styles);
  const selected = new Set<string>(cur.values ? cur.values : values.map((v) => v.text));
  return showPopup({ x: ctx.x, y: ctx.y }, (popup, close) => {
    popup.classList.add("filter-popup");
    popup.append(el("div", { class: "item", style: { fontWeight: "600", padding: "2px 4px 6px" } }, ctx.header ? `Filter: ${ctx.header}` : "Filter"));
    const sortA = el("div", { class: "item" }, "Sort A → Z");
    const sortZ = el("div", { class: "item" }, "Sort Z → A");
    sortA.addEventListener("click", () => { close(); ctx.onSort(true); });
    sortZ.addEventListener("click", () => { close(); ctx.onSort(false); });
    popup.append(sortA, sortZ, el("div", { class: "menu-sep" }));
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
        cb.addEventListener("change", () => { if (cb.checked) selected.add(v.text); else selected.delete(v.text); });
        const label = el("label", null, cb, el("span", { style: { flex: "1", overflow: "hidden", textOverflow: "ellipsis" } }, v.text === "" ? "(Blanks)" : v.text), el("span", { style: { color: "var(--ui-muted)" } }, String(v.count)));
        list.appendChild(label);
        rows.push({ text: v.text, label, cb });
      }
      if (!shown) list.appendChild(el("div", { class: "item", style: { color: "var(--ui-muted)" } }, "No matches"));
    };
    const allRow = el("div", { class: "row", style: { padding: "4px 0", gap: "12px" } });
    const selAll = el("a", { href: "#", style: { color: "var(--ui-accent)" } }, "Select all");
    const selNone = el("a", { href: "#", style: { color: "var(--ui-accent)" } }, "Clear");
    selAll.addEventListener("click", (e) => { e.preventDefault(); const q = search.value.toLocaleLowerCase(); for (const v of values) if (!q || v.text.toLocaleLowerCase().includes(q)) selected.add(v.text); render(); });
    selNone.addEventListener("click", (e) => { e.preventDefault(); const q = search.value.toLocaleLowerCase(); for (const v of values) if (!q || v.text.toLocaleLowerCase().includes(q)) selected.delete(v.text); render(); });
    allRow.append(selAll, selNone, el("span", { style: { flex: "1" } }), el("span", { style: { color: "var(--ui-muted)" } }, `${values.length} values`));
    popup.append(allRow, list);
    search.addEventListener("input", render);
    search.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        // Enter with a search term: keep only the matching values (Excel behaviour)
        const q = search.value.toLocaleLowerCase();
        if (q) { selected.clear(); for (const v of values) if (v.text.toLocaleLowerCase().includes(q)) selected.add(v.text); }
        apply();
      } else if (e.key === "Escape") { e.stopPropagation(); close(); }
    });
    render();
    const apply = () => {
      const all = selected.size >= values.length && values.every((v) => selected.has(v.text));
      if (all) st.delete(col); else st.set(col, { values: new Set(selected) });
      applyFilters(sheet, styles);
      close();
      ctx.onApply();
    };
    const ok = el("button", { class: "primary" }, "OK");
    const cancel = el("button", null, "Cancel");
    const clearBtn = el("button", null, "Clear filter");
    ok.addEventListener("click", apply);
    cancel.addEventListener("click", () => close());
    clearBtn.addEventListener("click", () => { st.delete(col); applyFilters(sheet, styles); close(); ctx.onApply(); });
    popup.appendChild(el("div", { class: "row", style: { justifyContent: "flex-end", padding: "6px 0 0" } }, clearBtn, el("span", { style: { flex: "1" } }), cancel, ok));
    setTimeout(() => search.focus(), 0);
  });
}

/** Whether the column has an active filter (for drawing the funnel icon differently). */
export function columnFiltered(sheet: Sheet, col: number): boolean {
  const st = filterStates.get(sheet);
  return !!st && (st.get(col)?.values ?? null) !== null;
}

export { MAXR, MAXC };
