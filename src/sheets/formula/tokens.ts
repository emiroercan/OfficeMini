// Formula tokenizer shared by the parser, reference shifting (copy/paste,
// fill, insert/delete rows) and the formula bar highlighter.
import { parseRef, RefParts, cellRef, colName, quoteSheet, MAXR, MAXC } from "../model";

export type Token =
  | { type: "num"; text: string; value: number }
  | { type: "str"; text: string; value: string }
  | { type: "bool"; text: string; value: boolean }
  | { type: "err"; text: string; value: string }
  | { type: "ref"; text: string; ref: RefParts }
  | { type: "name"; text: string }
  | { type: "func"; text: string; name: string }
  | { type: "op"; text: string }
  | { type: "lparen"; text: string }
  | { type: "rparen"; text: string }
  | { type: "sep"; text: string }
  | { type: "lbrace"; text: string }
  | { type: "rbrace"; text: string }
  | { type: "ws"; text: string }
  | { type: "unknown"; text: string };

const ERRORS = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!"];

export function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  const n = src.length;
  while (i < n) {
    const ch = src[i];
    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r") {
      let j = i; while (j < n && /\s/.test(src[j])) j++;
      out.push({ type: "ws", text: src.slice(i, j) }); i = j; continue;
    }
    if (ch === '"') {
      let j = i + 1, s = "";
      while (j < n) { if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j]; j++; }
      out.push({ type: "str", text: src.slice(i, j + 1), value: s }); i = j + 1; continue;
    }
    if (ch === "#") {
      const err = ERRORS.find((e) => src.startsWith(e, i));
      if (err) { out.push({ type: "err", text: err, value: err }); i += err.length; continue; }
    }
    if (/[0-9]/.test(ch)) {
      // whole-row range "1:1", "2:10"
      const rm = /^\d+:\$?\d+(?![\d.:])/.exec(src.slice(i));
      if (rm) { const ref = parseRef(rm[0]); if (ref) { out.push({ type: "ref", text: rm[0], ref }); i += rm[0].length; continue; } }
    }
    if (/[0-9]/.test(ch) || (ch === "." && /[0-9]/.test(src[i + 1] || ""))) {
      const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(src.slice(i))!;
      out.push({ type: "num", text: m[0], value: parseFloat(m[0]) }); i += m[0].length; continue;
    }
    if (ch === "(") { out.push({ type: "lparen", text: ch }); i++; continue; }
    if (ch === ")") { out.push({ type: "rparen", text: ch }); i++; continue; }
    if (ch === "{") { out.push({ type: "lbrace", text: ch }); i++; continue; }
    if (ch === "}") { out.push({ type: "rbrace", text: ch }); i++; continue; }
    if (ch === "," || ch === ";") { out.push({ type: "sep", text: ch }); i++; continue; }
    if (ch === "<" || ch === ">") {
      const two = src.slice(i, i + 2);
      if (two === "<=" || two === ">=" || two === "<>") { out.push({ type: "op", text: two }); i += 2; continue; }
      out.push({ type: "op", text: ch }); i++; continue;
    }
    if ("+-*/^&=%:".includes(ch)) { out.push({ type: "op", text: ch }); i++; continue; }
    // sheet-qualified or plain reference / name / function
    if (ch === "'" || /[A-Za-z_\\À-￿$]/.test(ch)) {
      let j = i;
      if (ch === "'") { j++; while (j < n) { if (src[j] === "'") { if (src[j + 1] === "'") { j += 2; continue; } break; } j++; } j++; if (src[j] !== "!") { out.push({ type: "unknown", text: src.slice(i, j) }); i = j; continue; } j++; }
      // identifier body (allow $ and : for ranges)
      while (j < n && /[A-Za-z0-9_.$À-￿]/.test(src[j])) j++;
      // unquoted sheet name: Sheet1!B2, Data.2026!A1:A9
      if (ch !== "'" && src[j] === "!" && /[A-Za-z$]/.test(src[j + 1] || "")) { j++; while (j < n && /[A-Za-z0-9_.$À-￿]/.test(src[j])) j++; }
      // range "A1:B2", "A:A", "1:1"
      if (src[j] === ":" && /[A-Za-z$0-9]/.test(src[j + 1] || "")) {
        let k = j + 1;
        while (k < n && /[A-Za-z0-9_.$À-￿]/.test(src[k])) k++;
        const candidate = src.slice(i, k);
        const ref = parseRef(candidate);
        if (ref) { out.push({ type: "ref", text: candidate, ref }); i = k; continue; }
      }
      const text = src.slice(i, j);
      if (src[j] === "(") { out.push({ type: "func", text, name: text.toUpperCase() }); i = j; continue; }
      const upper = text.toUpperCase();
      if (upper === "TRUE" || upper === "FALSE") { out.push({ type: "bool", text, value: upper === "TRUE" }); i = j; continue; }
      const ref = parseRef(text);
      if (ref && ref.r1 < MAXR && ref.c1 < MAXC) { out.push({ type: "ref", text, ref }); i = j; continue; }
      out.push({ type: "name", text }); i = j; continue;
    }
    out.push({ type: "unknown", text: ch }); i++;
  }
  return out;
}

/** Render a reference back to text, keeping the original sheet prefix and $ markers. */
export function refText(ref: RefParts): string {
  const sheet = ref.sheet ? quoteSheet(ref.sheet) + "!" : "";
  if (ref.colOnly) return sheet + (ref.absC1 ? "$" : "") + colName(ref.c1) + ":" + (ref.absC2 ? "$" : "") + colName(ref.c2);
  if (ref.rowOnly) return sheet + (ref.absR1 ? "$" : "") + (ref.r1 + 1) + ":" + (ref.absR2 ? "$" : "") + (ref.r2 + 1);
  if (!ref.isRange) return sheet + cellRef(ref.r1, ref.c1, ref.absR1, ref.absC1);
  return sheet + cellRef(ref.r1, ref.c1, ref.absR1, ref.absC1) + ":" + cellRef(ref.r2, ref.c2, ref.absR2, ref.absC2);
}

/**
 * Shift relative references by (dr, dc) — what copy/paste and fill do.
 * Out-of-range results become #REF!.
 */
export function shiftFormula(formula: string, dr: number, dc: number): string {
  if (!dr && !dc) return formula;
  return tokenize(formula).map((t) => {
    if (t.type !== "ref") return t.text;
    const r = { ...t.ref };
    if (!r.rowOnly) {
      if (!r.absC1) r.c1 += dc;
      if (!r.absC2) r.c2 += dc;
    }
    if (!r.colOnly) {
      if (!r.absR1) r.r1 += dr;
      if (!r.absR2) r.r2 += dr;
    }
    if (r.r1 < 0 || r.c1 < 0 || r.r2 < 0 || r.c2 < 0 || r.r2 >= MAXR || r.c2 >= MAXC) return "#REF!";
    return refText(r);
  }).join("");
}

/**
 * Adjust references for inserted/deleted rows or columns on `sheetName`
 * (references on other sheets that point to this sheet are handled by the caller
 * passing the right `currentSheet`). `at` is the first affected index, `count`
 * positive for insert, negative for delete.
 */
export function adjustFormulaForInsertDelete(formula: string, currentSheet: string, targetSheet: string, axis: "row" | "col", at: number, count: number): string {
  return tokenize(formula).map((t) => {
    if (t.type !== "ref") return t.text;
    const sheetOf = t.ref.sheet || currentSheet;
    if (sheetOf.toLowerCase() !== targetSheet.toLowerCase()) return t.text;
    const r = { ...t.ref };
    const adj = (v: number, isEnd: boolean): number | null => {
      if (count > 0) return v >= at ? v + count : v;
      const delEnd = at - count; // exclusive
      if (v >= at && v < delEnd) return isEnd ? at - 1 : null; // inside deleted block
      return v >= delEnd ? v + count : v;
    };
    if (axis === "row" && !r.colOnly) {
      const a = adj(r.r1, false), b = adj(r.r2, true);
      if (a === null && b === null) return "#REF!";
      if (a === null) { r.r1 = at; } else r.r1 = a;
      if (b === null) return "#REF!"; else r.r2 = b;
      if (r.r2 < r.r1) return "#REF!";
    } else if (axis === "col" && !r.rowOnly) {
      const a = adj(r.c1, false), b = adj(r.c2, true);
      if (a === null && b === null) return "#REF!";
      if (a === null) { r.c1 = at; } else r.c1 = a;
      if (b === null) return "#REF!"; else r.c2 = b;
      if (r.c2 < r.c1) return "#REF!";
    }
    return refText(r);
  }).join("");
}

/**
 * Remap references on `targetSheet` through a column/row index permutation - what reordering
 * columns or rows (drag a header) does. A single cell or a range that stays inside one moved
 * run is exact; a range that only straddles the moved run keeps its bounding box.
 */
export function remapRefs(formula: string, currentSheet: string, targetSheet: string, axis: "row" | "col", map: (v: number) => number): string {
  return tokenize(formula).map((t) => {
    if (t.type !== "ref") return t.text;
    const sheetOf = t.ref.sheet || currentSheet;
    if (sheetOf.toLowerCase() !== targetSheet.toLowerCase()) return t.text;
    const r = { ...t.ref };
    if (axis === "col") {
      if (r.rowOnly) return t.text;
      const a = map(r.c1), b = map(r.c2); r.c1 = Math.min(a, b); r.c2 = Math.max(a, b);
    } else {
      if (r.colOnly) return t.text;
      const a = map(r.r1), b = map(r.r2); r.r1 = Math.min(a, b); r.r2 = Math.max(a, b);
    }
    return refText(r);
  }).join("");
}

/** Rename a sheet inside formulas. */
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  return tokenize(formula).map((t) => {
    if (t.type !== "ref" || !t.ref.sheet || t.ref.sheet.toLowerCase() !== oldName.toLowerCase()) return t.text;
    return refText({ ...t.ref, sheet: newName });
  }).join("");
}

/** All references used by a formula (for dependency tracking / highlighting). */
export function formulaRefs(formula: string): RefParts[] {
  const out: RefParts[] = [];
  for (const t of tokenize(formula)) if (t.type === "ref") out.push(t.ref);
  return out;
}
