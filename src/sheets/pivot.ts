// Pivot tables: a builder dialog (rows, columns, values + aggregate, filter)
// that writes a static summary table into a new sheet. The definition is
// remembered per output sheet so it can be refreshed after the source changes.
import { Workbook, Sheet, Range, key, Cell, Value, isError, cellRef, quoteSheet } from "./model";
import { StyleResolver } from "./render-style";
import { el } from "../ui/widgets";
import { showDialog } from "../ui/dialog-core";
import { dataBlock, looksLikeHeader, displayOf } from "./filter";
import { formatValue } from "./numfmt";

export type Agg = "SUM" | "COUNT" | "COUNTA" | "AVERAGE" | "MIN" | "MAX";
export interface PivotDef {
  sourceSheet: string; source: Range;
  rows: number[]; cols: number[];           // field column offsets within the source
  values: { field: number; agg: Agg }[];
  filter: { field: number; value: string } | null;
  totals: boolean;
}
export const pivotDefs = new WeakMap<Sheet, PivotDef>();

const AGGS: Agg[] = ["SUM", "COUNT", "COUNTA", "AVERAGE", "MIN", "MAX"];

function headers(sheet: Sheet, rg: Range, styles: StyleResolver): string[] {
  const out: string[] = [];
  for (let c = rg.c1; c <= rg.c2; c++) { const t = displayOf(sheet.cells.get(key(rg.r1, c)), styles); out.push(t || `Column ${c - rg.c1 + 1}`); }
  return out;
}

export function pivotDialog(wb: Workbook, sheet: Sheet, sel: Range, styles: StyleResolver, existing?: PivotDef): Promise<PivotDef | null> {
  return new Promise((resolve) => {
    let src = existing ? existing.source : (sel.r1 === sel.r2 && sel.c1 === sel.c2 ? dataBlock(sheet, sel.r1, sel.c1) : sel);
    src = { r1: src.r1, c1: src.c1, r2: Math.min(src.r2, Math.max(sheet.maxRow, src.r1)), c2: Math.min(src.c2, Math.max(sheet.maxCol, src.c1)) };
    if (!looksLikeHeader(sheet, src) && src.r2 > src.r1) { /* still treat first row as header */ }
    const hs = headers(sheet, src, styles);
    const fieldSel = (value: number | null, allowNone = true) => {
      const s = el("select", null, allowNone ? el("option", { value: "-1" }, "(none)") : null, ...hs.map((h, i) => el("option", { value: String(i) }, h)));
      s.value = value === null ? "-1" : String(value);
      return s;
    };
    const rowSel = fieldSel(existing?.rows[0] ?? 0, false);
    const row2Sel = fieldSel(existing?.rows[1] ?? null);
    const colSel = fieldSel(existing?.cols[0] ?? null);
    const valSel = fieldSel(existing?.values[0]?.field ?? Math.min(hs.length - 1, 1), false);
    const aggSel = el("select", null, ...AGGS.map((a) => el("option", { value: a }, a)));
    aggSel.value = existing?.values[0]?.agg || "SUM";
    const val2Sel = fieldSel(existing?.values[1]?.field ?? null);
    const agg2Sel = el("select", null, ...AGGS.map((a) => el("option", { value: a }, a)));
    agg2Sel.value = existing?.values[1]?.agg || "SUM";
    const filtSel = fieldSel(existing?.filter?.field ?? null);
    const filtVal = el("input", { type: "text", value: existing?.filter?.value || "", placeholder: "equals…" });
    const totals = el("input", { type: "checkbox" }); totals.checked = existing ? existing.totals : true;
    const srcLabel = el("div", { style: { color: "var(--ui-muted)", marginBottom: "8px" } }, `Source: ${quoteSheet(sheet.name)}!${cellRef(src.r1, src.c1)}:${cellRef(src.r2, src.c2)} (${src.r2 - src.r1} rows, first row = headers)`);
    const body = el("div", null, srcLabel,
      el("div", { class: "grid2" },
        el("label", null, "Rows"), rowSel, el("label", null, "Then by"), row2Sel,
        el("label", null, "Columns"), colSel, el("label", null, ""), el("span"),
        el("label", null, "Values"), el("span", { style: { display: "flex", gap: "6px" } }, valSel, aggSel),
        el("label", null, "Second value"), el("span", { style: { display: "flex", gap: "6px" } }, val2Sel, agg2Sel),
        el("label", null, "Filter"), el("span", { style: { display: "flex", gap: "6px" } }, filtSel, filtVal)),
      el("label", { style: { display: "flex", alignItems: "center", gap: "6px", marginTop: "10px", color: "var(--ui-fg)" } }, totals, "Show totals"));
    let done = false;
    showDialog(existing ? "Refresh pivot table" : "Create pivot table", body, [
      { label: "Cancel" },
      { label: existing ? "Refresh" : "Create", primary: true, action: () => {
        done = true;
        const rows = [parseInt(rowSel.value, 10)]; if (row2Sel.value !== "-1") rows.push(parseInt(row2Sel.value, 10));
        const cols = colSel.value !== "-1" ? [parseInt(colSel.value, 10)] : [];
        const values = [{ field: parseInt(valSel.value, 10), agg: aggSel.value as Agg }];
        if (val2Sel.value !== "-1") values.push({ field: parseInt(val2Sel.value, 10), agg: agg2Sel.value as Agg });
        resolve({ sourceSheet: sheet.name, source: src, rows, cols, values, filter: filtSel.value !== "-1" && filtVal.value !== "" ? { field: parseInt(filtSel.value, 10), value: filtVal.value } : null, totals: totals.checked });
      } },
    ], { onClose: () => { if (!done) resolve(null); }, width: "520px" });
    void wb;
  });
}

interface Acc { sum: number; n: number; nAll: number; min: number; max: number; }
function acc(): Acc { return { sum: 0, n: 0, nAll: 0, min: Infinity, max: -Infinity }; }
function add(a: Acc, v: Value) { if (v === null || v === "") return; a.nAll++; if (typeof v === "number") { a.sum += v; a.n++; if (v < a.min) a.min = v; if (v > a.max) a.max = v; } }
function result(a: Acc, agg: Agg): Value {
  switch (agg) {
    case "SUM": return a.sum; case "COUNT": return a.n; case "COUNTA": return a.nAll;
    case "AVERAGE": return a.n ? a.sum / a.n : null; case "MIN": return a.n ? a.min : null; case "MAX": return a.n ? a.max : null;
  }
}

/** Compute the pivot as a grid of cells (relative coordinates) with style hints. */
export function computePivot(wb: Workbook, def: PivotDef, styles: StyleResolver): { cells: { r: number; c: number; v: Value; bold?: boolean; fmt?: string }[]; width: number; height: number } {
  const src = wb.sheets.find((s) => s.name.toLowerCase() === def.sourceSheet.toLowerCase());
  if (!src) return { cells: [{ r: 0, c: 0, v: "#REF! (source sheet missing)" }], width: 1, height: 1 };
  const rg = def.source;
  const hs = headers(src, rg, styles);
  const keyText = (r: number, f: number): string => displayOf(src.cells.get(key(r, rg.c1 + f)), styles);
  const rawVal = (r: number, f: number): Value => { const c = src.cells.get(key(r, rg.c1 + f)); return c ? c.v : null; };
  const fmtOf = (f: number): string => { for (let r = rg.r1 + 1; r <= rg.r2; r++) { const c = src.cells.get(key(r, rg.c1 + f)); if (c && typeof c.v === "number") { const fm = styles.get(c.s).numFmt; return fm === "General" ? "#,##0.00" : fm; } } return "#,##0.00"; };
  const rowKeys = new Map<string, { parts: string[]; sort: (number | string)[] }>();
  const colKeys = new Map<string, { parts: string[]; sort: (number | string)[] }>();
  const groups = new Map<string, Map<string, Acc[]>>();
  const sortVal = (r: number, f: number): number | string => { const v = rawVal(r, f); return typeof v === "number" ? v : keyText(r, f).toLocaleLowerCase(); };
  for (let r = rg.r1 + 1; r <= rg.r2; r++) {
    if (def.filter && keyText(r, def.filter.field) !== def.filter.value) continue;
    const rp = def.rows.map((f) => keyText(r, f)), cp = def.cols.map((f) => keyText(r, f));
    const rk = rp.join(""), ck = cp.join("");
    if (!rowKeys.has(rk)) rowKeys.set(rk, { parts: rp, sort: def.rows.map((f) => sortVal(r, f)) });
    if (!colKeys.has(ck)) colKeys.set(ck, { parts: cp, sort: def.cols.map((f) => sortVal(r, f)) });
    let g = groups.get(rk); if (!g) { g = new Map(); groups.set(rk, g); }
    let a = g.get(ck); if (!a) { a = def.values.map(acc); g.set(ck, a); }
    def.values.forEach((v, i) => add(a![i], rawVal(r, v.field)));
  }
  const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });
  const cmp = (a: (number | string)[], b: (number | string)[]) => { for (let i = 0; i < a.length; i++) { const x = a[i], y = b[i]; const d = typeof x === "number" && typeof y === "number" ? x - y : typeof x === "number" ? -1 : typeof y === "number" ? 1 : collator.compare(x as string, y as string); if (d) return d; } return 0; };
  const rks = Array.from(rowKeys.entries()).sort((a, b) => cmp(a[1].sort, b[1].sort));
  const cks = Array.from(colKeys.entries()).sort((a, b) => cmp(a[1].sort, b[1].sort));
  const out: { r: number; c: number; v: Value; bold?: boolean; fmt?: string }[] = [];
  const nv = def.values.length;
  const fmts = def.values.map((v) => (v.agg === "COUNT" || v.agg === "COUNTA" ? "#,##0" : fmtOf(v.field)));
  let r = 0;
  // header rows
  const hasCols = def.cols.length > 0;
  const rowHeadW = def.rows.length;
  if (hasCols) {
    out.push({ r, c: 0, v: hs[def.rows[0]], bold: true });
    cks.forEach((ck, i) => { out.push({ r, c: rowHeadW + i * nv, v: ck[1].parts.join(" / "), bold: true }); });
    if (def.totals) out.push({ r, c: rowHeadW + cks.length * nv, v: "Total", bold: true });
    r++;
  }
  def.rows.forEach((f, i) => out.push({ r, c: i, v: hs[f], bold: true }));
  const valLabel = (v: PivotDef["values"][number]) => `${v.agg} of ${hs[v.field]}`;
  const colCount = hasCols ? cks.length + (def.totals ? 1 : 0) : 1;
  for (let ci = 0; ci < colCount; ci++) def.values.forEach((v, vi) => out.push({ r, c: rowHeadW + ci * nv + vi, v: nv === 1 && hasCols && ci < cks.length ? (ci === 0 ? valLabel(v) : null) : valLabel(v), bold: true }));
  if (nv === 1 && hasCols) { out.length = out.length; }
  r++;
  const colTotals: Acc[][] = Array.from({ length: colCount }, () => def.values.map(acc));
  let prevParts: string[] = [];
  for (const [rk, info] of rks) {
    info.parts.forEach((p, i) => { if (i === info.parts.length - 1 || p !== prevParts[i]) out.push({ r, c: i, v: p }); });
    prevParts = info.parts;
    const g = groups.get(rk)!;
    const rowTotal = def.values.map(acc);
    if (hasCols) {
      cks.forEach((ck, ci) => {
        const a = g.get(ck[0]);
        def.values.forEach((v, vi) => {
          const val = a ? result(a[vi], v.agg) : null;
          out.push({ r, c: rowHeadW + ci * nv + vi, v: val, fmt: fmts[vi] });
          if (a) { mergeAcc(rowTotal[vi], a[vi]); mergeAcc(colTotals[ci][vi], a[vi]); }
        });
      });
      if (def.totals) def.values.forEach((v, vi) => { out.push({ r, c: rowHeadW + cks.length * nv + vi, v: result(rowTotal[vi], v.agg), fmt: fmts[vi], bold: true }); mergeAcc(colTotals[cks.length][vi], rowTotal[vi]); });
    } else {
      const a = g.get("");
      def.values.forEach((v, vi) => { const val = a ? result(a[vi], v.agg) : null; out.push({ r, c: rowHeadW + vi, v: val, fmt: fmts[vi] }); if (a) mergeAcc(colTotals[0][vi], a[vi]); });
    }
    r++;
  }
  if (def.totals) {
    out.push({ r, c: 0, v: "Total", bold: true });
    for (let ci = 0; ci < colCount; ci++) def.values.forEach((v, vi) => out.push({ r, c: rowHeadW + ci * nv + vi, v: result(colTotals[ci][vi], v.agg), fmt: fmts[vi], bold: true }));
    r++;
  }
  return { cells: out, width: rowHeadW + colCount * nv, height: r };
}

function mergeAcc(into: Acc, a: Acc) { into.sum += a.sum; into.n += a.n; into.nAll += a.nAll; if (a.min < into.min) into.min = a.min; if (a.max > into.max) into.max = a.max; }

export { formatValue, isError };
export type { Cell };
