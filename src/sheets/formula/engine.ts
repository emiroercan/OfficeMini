// Formula parser and evaluator with a practical Excel function set. Unknown
// functions leave the cell's cached value untouched (Excel recalculates on
// open because the writer sets fullCalcOnLoad).
import { Workbook, Sheet, Cell, Value, CellError, key, isError, MAXR, MAXC, RefParts, parseRef } from "../model";
import { tokenize, Token } from "./tokens";
import { formatValue, parseInput, todaySerial, nowSerial, serialToDate, dateToSerial, locale, parseLocaleNumber } from "../numfmt";

export type Val = number | string | boolean | CellError | null;
export interface Arr { rows: number; cols: number; get(i: number, j: number): Val; }
type Any = Val | Arr;

const ERR = (e: string): CellError => ({ e });
const NA = ERR("#N/A"), VALUE = ERR("#VALUE!"), DIV0 = ERR("#DIV/0!"), REF = ERR("#REF!"), NAME = ERR("#NAME?"), NUM = ERR("#NUM!");
const isArr = (v: Any): v is Arr => !!v && typeof v === "object" && "rows" in v;

// ---- AST ---------------------------------------------------------------------

type Node =
  | { t: "num"; v: number } | { t: "str"; v: string } | { t: "bool"; v: boolean } | { t: "err"; v: string }
  | { t: "ref"; ref: RefParts } | { t: "name"; name: string }
  | { t: "func"; name: string; args: Node[] }
  | { t: "bin"; op: string; l: Node; r: Node } | { t: "un"; op: string; x: Node } | { t: "pct"; x: Node }
  | { t: "array"; rows: Node[][] };

const astCache = new Map<string, Node | CellError>();

class Parser {
  private i = 0;
  private toks: Token[];
  constructor(src: string) { this.toks = tokenize(src).filter((t) => t.type !== "ws"); }
  private peek(): Token | undefined { return this.toks[this.i]; }
  private next(): Token { return this.toks[this.i++]; }
  private isOp(text: string): boolean { const t = this.peek(); return !!t && t.type === "op" && t.text === text; }
  parse(): Node {
    const n = this.comparison();
    if (this.i < this.toks.length) throw new Error("Unexpected " + this.toks[this.i].text);
    return n;
  }
  private comparison(): Node {
    let l = this.concat();
    while (this.peek()?.type === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(this.peek()!.text)) { const op = this.next().text; l = { t: "bin", op, l, r: this.concat() }; }
    return l;
  }
  private concat(): Node { let l = this.additive(); while (this.isOp("&")) { this.next(); l = { t: "bin", op: "&", l, r: this.additive() }; } return l; }
  private additive(): Node { let l = this.term(); while (this.isOp("+") || this.isOp("-")) { const op = this.next().text; l = { t: "bin", op, l, r: this.term() }; } return l; }
  private term(): Node { let l = this.power(); while (this.isOp("*") || this.isOp("/")) { const op = this.next().text; l = { t: "bin", op, l, r: this.power() }; } return l; }
  private power(): Node { let l = this.unary(); while (this.isOp("^")) { this.next(); l = { t: "bin", op: "^", l, r: this.unary() }; } return l; }
  private unary(): Node {
    if (this.isOp("-")) { this.next(); return { t: "un", op: "-", x: this.unary() }; }
    if (this.isOp("+")) { this.next(); return this.unary(); }
    return this.postfix();
  }
  private postfix(): Node { let x = this.primary(); while (this.isOp("%")) { this.next(); x = { t: "pct", x }; } return x; }
  private primary(): Node {
    const t = this.next();
    if (!t) throw new Error("Unexpected end");
    switch (t.type) {
      case "num": return { t: "num", v: t.value };
      case "str": return { t: "str", v: t.value };
      case "bool": return { t: "bool", v: t.value };
      case "err": return { t: "err", v: t.value };
      case "ref": {
        // range operator between two refs: A1:B2 already tokenized; support "A1:" "INDEX(...)"? keep simple
        return { t: "ref", ref: t.ref };
      }
      case "name": return { t: "name", name: t.text };
      case "func": {
        const lp = this.next(); if (!lp || lp.type !== "lparen") throw new Error("Expected (");
        const args: Node[] = [];
        if (this.peek()?.type !== "rparen") {
          for (;;) {
            if (this.peek()?.type === "sep") { args.push({ t: "err", v: "#EMPTY" }); this.next(); continue; }
            args.push(this.comparison());
            if (this.peek()?.type === "sep") { this.next(); if (this.peek()?.type === "rparen") { args.push({ t: "err", v: "#EMPTY" }); break; } continue; }
            break;
          }
        }
        const rp = this.next(); if (!rp || rp.type !== "rparen") throw new Error("Expected )");
        return { t: "func", name: t.name, args };
      }
      case "lparen": { const n = this.comparison(); const rp = this.next(); if (!rp || rp.type !== "rparen") throw new Error("Expected )"); return n; }
      case "lbrace": {
        const rows: Node[][] = [[]];
        while (this.peek() && this.peek()!.type !== "rbrace") {
          const tk = this.peek()!;
          if (tk.type === "sep") { this.next(); if (tk.text === ";") rows.push([]); continue; }
          rows[rows.length - 1].push(this.comparison());
        }
        this.next();
        return { t: "array", rows };
      }
      default: throw new Error("Unexpected " + t.text);
    }
  }
}

export function parseFormula(f: string): Node | CellError {
  let n = astCache.get(f);
  if (n) return n;
  try { n = new Parser(f).parse(); } catch { n = ERR("#ERROR!"); }
  astCache.set(f, n);
  return n;
}

// ---- helpers -----------------------------------------------------------------

function toNum(v: Val): number | CellError {
  if (typeof v === "number") return v;
  if (v === null) return 0;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (isError(v)) return v;
  const p = parseInput(v, locale());
  if (typeof p.value === "number") return p.value;
  return VALUE;
}
function toStr(v: Val): string {
  if (v === null) return "";
  if (typeof v === "string") return v;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (isError(v)) return v.e;
  return numToText(v);
}
/** Excel's number-to-text coercion: up to 15 significant digits, no grouping, scientific only for extreme magnitudes. */
function numToText(v: number): string {
  if (!isFinite(v)) return "#NUM!";
  if (v === 0) return "0";
  const a = Math.abs(v);
  if (a >= 1e21 || a < 1e-9) return v.toExponential(14).replace(/\.?0+e/, "e").replace("e+", "E+").replace("e-", "E-");
  let s = Number(v.toPrecision(15)).toString();
  if (/e/i.test(s)) s = v.toLocaleString("en-US", { useGrouping: false, maximumFractionDigits: 15 });
  return s;
}
function mapAny(v: Any, f: (x: Val) => Val): Any { if (!isArr(v)) return f(v); const a = v; return { rows: a.rows, cols: a.cols, get: (i, j) => f(a.get(i, j)) }; }
/** Element-wise binary operation with Excel broadcasting: single row/column/cell operands stretch, other size mismatches give #N/A. */
function zipAny(l: Any, r: Any, f: (a: Val, b: Val) => Val): Any {
  if (!isArr(l) && !isArr(r)) return f(l, r);
  const la = isArr(l) ? l : null, ra = isArr(r) ? r : null;
  const rows = Math.max(la ? la.rows : 1, ra ? ra.rows : 1), cols = Math.max(la ? la.cols : 1, ra ? ra.cols : 1);
  const pick = (a: Arr | null, v: Any, i: number, j: number): Val => {
    if (!a) return v as Val;
    const ii = a.rows === 1 ? 0 : i, jj = a.cols === 1 ? 0 : j;
    if (ii >= a.rows || jj >= a.cols) return NA;
    return a.get(ii, jj);
  };
  return { rows, cols, get: (i, j) => { const a = pick(la, l, i, j), b = pick(ra, r, i, j); if (isError(a)) return a; if (isError(b)) return b; return f(a, b); } };
}
function toBool(v: Val): boolean | CellError {
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (v === null) return false;
  if (isError(v)) return v;
  const u = v.toUpperCase(); if (u === "TRUE") return true; if (u === "FALSE") return false;
  return VALUE;
}
function first(v: Any): Val { return isArr(v) ? v.get(0, 0) : v; }
function isErrObj(x: unknown): x is CellError { return !!x && typeof x === "object" && typeof (x as any).e === "string" && !("t" in (x as any)); }
function* iterate(v: Any): Generator<Val> {
  if (isArr(v)) { for (let i = 0; i < v.rows; i++) for (let j = 0; j < v.cols; j++) yield v.get(i, j); }
  else yield v;
}

function compare(a: Val, b: Val): number {
  const rank = (v: Val) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 0);
  if (a === null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (b === null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === "number" && typeof b === "number") return a - b;
  if (typeof a === "string" && typeof b === "string") { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  return (a ? 1 : 0) - (b ? 1 : 0);
}

/** Excel criteria: ">10", "<>x", "*ab?", "=", "5", "text" */
function makeCriteria(crit: Val): (v: Val) => boolean {
  // Numeric / boolean criteria compare directly (a text round trip would lose digits of barcodes and IDs)
  if (typeof crit === "number") { const n = crit; return (v) => typeof v === "number" && v === n; }
  if (typeof crit === "boolean") { const b = crit; return (v) => v === b; }
  let s = crit === null ? "" : typeof crit === "string" ? crit : toStr(crit);
  let op = "=";
  const m = /^(<>|>=|<=|=|<|>)(.*)$/.exec(s);
  if (m) { op = m[1]; s = m[2]; }
  const st = s.trim();
  const numCrit = st !== "" && /\d/.test(st) ? (parseLocaleNumber(st, locale()) ?? (/^[+-]?\d+(\.\d+)?$/.test(st) ? parseFloat(st) : null)) : null;
  if (op === "=" || op === "<>") {
    const neg = op === "<>";
    if (/[*?]/.test(s)) {
      const re = new RegExp("^" + s.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i");
      return (v) => { const r = v !== null && re.test(toStr(v)); return neg ? !r : r; };
    }
    return (v) => {
      let r: boolean;
      if (numCrit !== null && typeof v === "number") r = v === numCrit;
      else if (s === "") r = v === null || v === "";
      else r = typeof v !== "number" && v !== null && toStr(v).toLowerCase() === s.toLowerCase();
      return neg ? !r : r;
    };
  }
  return (v) => {
    if (numCrit !== null) { if (typeof v !== "number") return false; return op === ">" ? v > numCrit : op === "<" ? v < numCrit : op === ">=" ? v >= numCrit : v <= numCrit; }
    if (v === null || typeof v === "number") return false;
    const c = compare(v, s);
    return op === ">" ? c > 0 : op === "<" ? c < 0 : op === ">=" ? c >= 0 : c <= 0;
  };
}

function wildcardMatch(pattern: string, s: string): boolean {
  const re = new RegExp("^" + pattern.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i");
  return re.test(s);
}

// ---- engine ------------------------------------------------------------------

export class Engine {
  private values = new Map<string, Val>();     // sheet|r|c -> value in the current pass
  private computing = new Set<string>();
  unsupported = new Set<string>();
  private pass = 0;
  private lookupIndex = new Map<string, Map<string, number>>();
  private inRecalc = false;

  constructor(public wb: Workbook) {}

  /** Normalised key for exact matching: case-insensitive for strings, type-exact otherwise. */
  keyOf(v: Val): string { return typeof v + "\u0000" + (typeof v === "string" ? v.toLowerCase() : String(v)); }

  /**
   * value -> first-index map for a lookup vector, memoised for the whole recalc pass. N exact
   * lookups over the same range then cost O(N) to build once plus O(1) each, instead of O(N)
   * apiece - the XLOOKUP/VLOOKUP/MATCH O(N^2) freeze on large sheets. Returns null when the
   * argument is not a plain cell range (so it has no stable key); the caller scans linearly then.
   */
  exactIndex(node: Node | undefined, arr: Arr, sheet: Sheet, byRow: boolean, colOffset = 0): Map<string, number> | null {
    if (!node || node.t !== "ref") return null;
    const rf = node.ref; const sh = this.sheetByName(rf.sheet, sheet); if (!sh) return null;
    const key = sh.name + "|" + rf.r1 + "," + rf.c1 + "," + rf.r2 + "," + rf.c2 + "|" + (byRow ? "R" : "C") + colOffset;
    let m = this.lookupIndex.get(key);
    if (m) return m;
    m = new Map<string, number>();
    const n = byRow ? arr.rows : arr.cols;
    for (let i = 0; i < n; i++) { const v = (byRow ? arr.get(i, colOffset) : arr.get(colOffset, i)) as Val; const k = this.keyOf(v); if (!m.has(k)) m.set(k, i); }
    this.lookupIndex.set(key, m);
    return m;
  }

  private sheetByName(name: string | null, current: Sheet): Sheet | null {
    if (!name) return current;
    const lower = name.toLowerCase();
    return this.wb.sheets.find((s) => s.name.toLowerCase() === lower) || null;
  }

  /** Value of a cell, evaluating its formula when needed (memoised per pass). */
  cellValue(sheet: Sheet, r: number, c: number): Val {
    const cell = sheet.cells.get(key(r, c));
    if (!cell) return null;
    if (!cell.f) return cell.v as Val;
    const k = sheet.name + "|" + r + "|" + c;
    if (this.values.has(k)) return this.values.get(k)!;
    if (this.computing.has(k)) return ERR("#CIRC!");
    this.computing.add(k);
    let v: Val;
    try { v = this.evaluateFormula(cell.f, sheet, r, c); } catch { v = VALUE; }
    this.computing.delete(k);
    if (isError(v) && (v.e === "#NAME?" || v.e === "#ERROR!") && cell.v !== null && !isError(cell.v)) v = cell.v as Val; // keep cached value for unsupported functions
    if (cell.arr) v = cell.v as Val; // array formulas: cached value only
    this.values.set(k, v);
    return v;
  }

  evaluateFormula(f: string, sheet: Sheet, r: number, c: number): Val {
    if (!this.inRecalc) this.lookupIndex.clear();   // a standalone eval (editing/F9) starts fresh
    const ast = parseFormula(f);
    if (isErrObj(ast)) return ast;
    const v = this.ev(ast, sheet, r, c);
    return first(v);
  }

  /** Recalculate every formula cell; returns the number of cells whose cached value changed. */
  recalcAll(): number {
    this.values.clear();
    this.computing.clear();
    this.unsupported.clear();
    this.lookupIndex.clear();
    this.inRecalc = true;
    this.pass++;
    let changed = 0;
    for (const sheet of this.wb.sheets) {
      // The set of formula cells is cached per sheet (applyCellChanges keeps it current; structural edits reset it),
      // so a plain-value edit on an 800k-cell sheet does not walk every cell.
      if (!sheet.formulaKeys) { const fk = new Set<number>(); for (const [k, cell] of sheet.cells) if (cell.f !== undefined) fk.add(k); sheet.formulaKeys = fk; }
      for (const k of sheet.formulaKeys) {
        const cell = sheet.cells.get(k);
        if (!cell || cell.f === undefined) { sheet.formulaKeys.delete(k); continue; }
        if (cell.arr) continue;
        const r = Math.floor(k / MAXC), c = k % MAXC;
        const v = this.cellValue(sheet, r, c);
        if (!sameVal(v, cell.v as Val)) { sheet.cells.set(k, { ...cell, v: v as Value }); changed++; sheet.dirty = true; }
      }
    }
    this.inRecalc = false;
    return changed;
  }

  private rangeArr(ref: RefParts, sheet: Sheet): Arr | CellError {
    const sh = this.sheetByName(ref.sheet, sheet);
    if (!sh) return REF;
    // Unbounded column/row refs: clamp to used area for iteration speed
    const r1 = ref.r1, c1 = ref.c1;
    const r2 = ref.rowOnly || ref.colOnly ? Math.min(ref.r2, Math.max(sh.maxRow, r1)) : ref.r2;
    const c2 = ref.colOnly || ref.rowOnly ? Math.min(ref.c2, Math.max(sh.maxCol, c1)) : ref.c2;
    const self = this;
    const rows = r2 - r1 + 1, cols = c2 - c1 + 1;
    return { rows, cols, get: (i, j) => (i < 0 || j < 0 || i >= rows || j >= cols ? REF : self.cellValue(sh, r1 + i, c1 + j)) };
  }
  /** Public helpers for functions that work on references (SUBTOTAL, OFFSET, INDIRECT, SUMIF sum_range). */
  sheetFor(name: string | null, current: Sheet): Sheet | null { return this.sheetByName(name, current); }
  rangeOf(ref: RefParts, sheet: Sheet): Arr | CellError { return this.rangeArr(ref, sheet); }
  /** Range of the given size anchored at the top-left of a reference argument (SUMIF's sum_range rule). */
  anchoredArr(node: Node | undefined, rows: number, cols: number, sheet: Sheet): Arr | null {
    if (!node || node.t !== "ref") return null;
    const sh = this.sheetByName(node.ref.sheet, sheet); if (!sh) return null;
    const r1 = node.ref.r1, c1 = node.ref.c1; const self = this;
    return { rows, cols, get: (i, j) => self.cellValue(sh, r1 + i, c1 + j) };
  }

  private ev(n: Node, sheet: Sheet, r: number, c: number): Any {
    switch (n.t) {
      case "num": return n.v;
      case "str": return n.v;
      case "bool": return n.v;
      case "err": return n.v === "#EMPTY" ? null : ERR(n.v);
      case "ref": {
        if (!n.ref.isRange) { const sh = this.sheetByName(n.ref.sheet, sheet); if (!sh) return REF; return this.cellValue(sh, n.ref.r1, n.ref.c1); }
        return this.rangeArr(n.ref, sheet);
      }
      case "name": {
        if (this.scopes.length) { const bound = this.lookupScope(n.name); if (bound !== undefined) return bound; }
        const dn = this.wb.definedNames.find((d) => d.name.toLowerCase() === n.name.toLowerCase());
        if (!dn) return NAME;
        const ref = parseRef(dn.ref.replace(/^=/, ""));
        if (!ref) return NAME;
        return this.ev({ t: "ref", ref }, sheet, r, c);
      }
      case "array": {
        const rows = n.rows.map((row) => row.map((x) => first(this.ev(x, sheet, r, c))));
        return { rows: rows.length, cols: rows[0]?.length || 0, get: (i, j) => rows[i]?.[j] ?? null };
      }
      case "un": return mapAny(this.ev(n.x, sheet, r, c), (x) => { const v = toNum(x); return isError(v) ? v : -v; });
      case "pct": return mapAny(this.ev(n.x, sheet, r, c), (x) => { const v = toNum(x); return isError(v) ? v : v / 100; });
      case "bin": { const l = this.ev(n.l, sheet, r, c), rr = this.ev(n.r, sheet, r, c); return zipAny(l, rr, (a, b) => this.binary(n.op, a, b)); }
      case "func": return this.call(n.name, n.args, sheet, r, c);
    }
  }

  private binary(op: string, a: Val, b: Val): Val {
    if (isError(a)) return a;
    if (isError(b)) return b;
    switch (op) {
      case "&": return toStr(a) + toStr(b);
      case "=": return compare(a, b) === 0;
      case "<>": return compare(a, b) !== 0;
      case "<": return compare(a, b) < 0;
      case ">": return compare(a, b) > 0;
      case "<=": return compare(a, b) <= 0;
      case ">=": return compare(a, b) >= 0;
    }
    const x = toNum(a), y = toNum(b);
    if (isError(x)) return x; if (isError(y)) return y;
    switch (op) {
      case "+": return x + y;
      case "-": return x - y;
      case "*": return x * y;
      case "/": return y === 0 ? DIV0 : x / y;
      case "^": { const p = Math.pow(x, y); return isFinite(p) ? p : NUM; }
    }
    return VALUE;
  }

  /** LET() bindings, innermost last. */
  private scopes: Map<string, Any>[] = [];
  private lookupScope(name: string): Any | undefined {
    const k = name.toUpperCase();
    for (let i = this.scopes.length - 1; i >= 0; i--) { const s = this.scopes[i]; if (s.has(k)) return s.get(k); }
    return undefined;
  }

  private call(name: string, args: Node[], sheet: Sheet, r: number, c: number): Any {
    if (name === "LET") {
      // LET(name1, value1, [name2, value2, ...], expression) — names are visible to later values and the expression
      if (args.length < 3 || args.length % 2 === 0) return VALUE;
      const scope = new Map<string, Any>();
      this.scopes.push(scope);
      try {
        for (let i = 0; i + 1 < args.length - 1; i += 2) {
          const nm = args[i];
          if (nm.t !== "name") return VALUE; // names that look like cell references (x1) are not valid LET names
          scope.set(nm.name.toUpperCase(), this.ev(args[i + 1], sheet, r, c));
        }
        return this.ev(args[args.length - 1], sheet, r, c);
      } finally { this.scopes.pop(); }
    }
    const fn = FUNCS[name];
    if (!fn) { this.unsupported.add(name); return NAME; }
    // lazy args for IF-like functions
    const ctx: Ctx = { eval: (i) => (i < args.length ? this.ev(args[i], sheet, r, c) : undefined), node: (i) => args[i], n: args.length, sheet, r, c, engine: this };
    try { return fn(ctx); } catch (e) { return e && typeof e === "object" && "e" in (e as any) ? (e as CellError) : VALUE; }
  }
}

function sameVal(a: Val, b: Val): boolean {
  if (isError(a) || isError(b)) return isError(a) && isError(b) && a.e === b.e;
  return a === b;
}

// ---- functions ----------------------------------------------------------------

interface Ctx { eval(i: number): Any | undefined; node(i: number): Node | undefined; n: number; sheet: Sheet; r: number; c: number; engine: Engine; }
type Fn = (ctx: Ctx) => Any;

const argVals = (ctx: Ctx, from = 0): Any[] => { const out: Any[] = []; for (let i = from; i < ctx.n; i++) out.push(ctx.eval(i)!); return out; };
const numsOf = (vals: Any[], includeText = false): number[] | CellError => {
  const out: number[] = [];
  for (const v of vals) {
    if (isArr(v)) { for (const x of iterate(v)) { if (isError(x)) return x; if (typeof x === "number") out.push(x); else if (typeof x === "boolean" && includeText) out.push(x ? 1 : 0); } }
    else { if (isError(v)) return v; if (typeof v === "number") out.push(v); else if (typeof v === "boolean") out.push(v ? 1 : 0); else if (typeof v === "string") { const n = parseInput(v, locale()).value; if (typeof n === "number") out.push(n); else if (v !== "") return VALUE; } }
  }
  return out;
};
const need = (ctx: Ctx, i: number): Val => first(ctx.eval(i) ?? null);
const numArg = (ctx: Ctx, i: number, def?: number): number | CellError => { const v = ctx.eval(i); if (v === undefined || v === null) return def !== undefined ? def : 0; return toNum(first(v)); };
const strArg = (ctx: Ctx, i: number, def = ""): string | CellError => { const v = ctx.eval(i); if (v === undefined) return def; const f = first(v); return isError(f) ? f : toStr(f); };
const chk = <T>(v: T | CellError, f: (x: T) => Any): Any => (isError(v as any) ? (v as CellError) : f(v as T));

function round(n: number, d: number, mode: "round" | "up" | "down"): number {
  const m = Math.pow(10, d);
  const x = n * m;
  let r: number;
  if (mode === "round") r = Math.sign(x) * Math.round(Math.abs(x) + 1e-12);
  else if (mode === "up") r = Math.sign(x) * Math.ceil(Math.abs(x) - 1e-12);
  else r = Math.sign(x) * Math.floor(Math.abs(x) + 1e-12);
  return r / m;
}

const dateParts = (serial: number, c?: Ctx) => { const d = serialToDate(Math.floor(serial), !!c?.engine.wb.date1904); return { y: d.getUTCFullYear(), m: d.getUTCMonth(), d: d.getUTCDate(), wd: d.getUTCDay() }; };
const makeDate = (y: number, m: number, d: number, c?: Ctx) => dateToSerial(new Date(Date.UTC(y, m, d)), !!c?.engine.wb.date1904);
function holidaySet(c: Ctx, i: number): Set<number> { const out = new Set<number>(); const v = c.eval(i); if (v === undefined) return out; for (const x of iterate(v)) if (typeof x === "number") out.add(Math.floor(x)); return out; }
function subtotalOf(code: number, vals: number[], cnta: number): Val {
  const n = vals.length, sum = vals.reduce((a, b) => a + b, 0);
  const mean = n ? sum / n : 0;
  const varS = n > 1 ? vals.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - 1) : NaN;
  const varP = n ? vals.reduce((a, b) => a + (b - mean) ** 2, 0) / n : NaN;
  switch (code) {
    case 1: return n ? mean : DIV0; case 2: return n; case 3: return cnta; case 4: return n ? Math.max(...vals) : 0; case 5: return n ? Math.min(...vals) : 0;
    case 6: return vals.reduce((a, b) => a * b, 1); case 7: return n > 1 ? Math.sqrt(varS) : DIV0; case 8: return n ? Math.sqrt(varP) : DIV0; case 9: return sum;
    case 10: return n > 1 ? varS : DIV0; case 11: return n ? varP : DIV0;
  }
  return VALUE;
}

function lookupExact(needle: Val, arr: Arr, mode: number): number {
  const n = arr.rows === 1 ? arr.cols : arr.rows;
  const get = (i: number) => (arr.rows === 1 ? arr.get(0, i) : arr.get(i, 0));
  const ns = typeof needle === "string" ? needle.toLowerCase() : needle;
  for (let i = 0; i < n; i++) {
    const v = get(i);
    if (mode === 2 && typeof needle === "string") { if (typeof v === "string" && wildcardMatch(needle, v)) return i; continue; }
    if (typeof v === "string" && typeof ns === "string") { if (v.toLowerCase() === ns) return i; }
    else if (v === needle) return i;
  }
  if (mode === -1 || mode === 1) {
    // next smaller (-1) or next larger (1)
    let best = -1;
    for (let i = 0; i < n; i++) {
      const v = get(i);
      if (v === null || isError(v)) continue;
      const cmp = compare(v, needle);
      if (mode === -1 && cmp < 0 && (best < 0 || compare(v, get(best)) > 0)) best = i;
      if (mode === 1 && cmp > 0 && (best < 0 || compare(v, get(best)) < 0)) best = i;
    }
    return best;
  }
  return -1;
}
function lookupApprox(needle: Val, arr: Arr, colIdx: number, byRow: boolean): number {
  // last value <= needle (assumes sorted ascending)
  const n = byRow ? arr.rows : arr.cols;
  let best = -1;
  for (let i = 0; i < n; i++) {
    const v = byRow ? arr.get(i, colIdx) : arr.get(colIdx, i);
    if (v === null || isError(v)) continue;
    if (compare(v, needle) <= 0) best = i; else break;
  }
  return best;
}

const FUNCS: Record<string, Fn> = {
  SUM: (c) => chk(numsOf(argVals(c)), (a) => a.reduce((x, y) => x + y, 0)),
  AVERAGE: (c) => chk(numsOf(argVals(c)), (a) => (a.length ? a.reduce((x, y) => x + y, 0) / a.length : DIV0)),
  MIN: (c) => chk(numsOf(argVals(c)), (a) => (a.length ? Math.min(...a) : 0)),
  MAX: (c) => chk(numsOf(argVals(c)), (a) => (a.length ? Math.max(...a) : 0)),
  COUNT: (c) => { let n = 0; for (const v of argVals(c)) for (const x of iterate(v)) if (typeof x === "number") n++; return n; },
  COUNTA: (c) => { let n = 0; for (const v of argVals(c)) for (const x of iterate(v)) if (x !== null && x !== "") n++; return n; },
  COUNTBLANK: (c) => { let n = 0; for (const v of argVals(c)) for (const x of iterate(v)) if (x === null || x === "") n++; return n; },
  PRODUCT: (c) => chk(numsOf(argVals(c)), (a) => a.reduce((x, y) => x * y, 1)),
  ABS: (c) => chk(numArg(c, 0), Math.abs), SQRT: (c) => chk(numArg(c, 0), (x) => (x < 0 ? NUM : Math.sqrt(x))),
  POWER: (c) => chk(numArg(c, 0), (x) => chk(numArg(c, 1), (y) => Math.pow(x, y))),
  MOD: (c) => chk(numArg(c, 0), (x) => chk(numArg(c, 1), (y) => (y === 0 ? DIV0 : x - y * Math.floor(x / y)))),
  INT: (c) => chk(numArg(c, 0), Math.floor), TRUNC: (c) => chk(numArg(c, 0), (x) => chk(numArg(c, 1, 0), (d) => round(x, d, "down"))),
  ROUND: (c) => chk(numArg(c, 0), (x) => chk(numArg(c, 1, 0), (d) => round(x, d, "round"))),
  ROUNDUP: (c) => chk(numArg(c, 0), (x) => chk(numArg(c, 1, 0), (d) => round(x, d, "up"))),
  ROUNDDOWN: (c) => chk(numArg(c, 0), (x) => chk(numArg(c, 1, 0), (d) => round(x, d, "down"))),
  CEILING: (c) => chk(numArg(c, 0), (x) => chk(numArg(c, 1, 1), (s) => (s === 0 ? 0 : Math.ceil(x / s) * s))),
  FLOOR: (c) => chk(numArg(c, 0), (x) => chk(numArg(c, 1, 1), (s) => (s === 0 ? DIV0 : Math.floor(x / s) * s))),
  EXP: (c) => chk(numArg(c, 0), Math.exp), LN: (c) => chk(numArg(c, 0), (x) => (x <= 0 ? NUM : Math.log(x))), LOG10: (c) => chk(numArg(c, 0), (x) => (x <= 0 ? NUM : Math.log10(x))),
  LOG: (c) => chk(numArg(c, 0), (x) => chk(numArg(c, 1, 10), (b) => (x <= 0 || b <= 0 ? NUM : Math.log(x) / Math.log(b)))),
  PI: () => Math.PI, RAND: () => Math.random(), RANDBETWEEN: (c) => chk(numArg(c, 0), (a) => chk(numArg(c, 1), (b) => Math.floor(Math.random() * (b - a + 1)) + a)),
  SIGN: (c) => chk(numArg(c, 0), Math.sign), EVEN: (c) => chk(numArg(c, 0), (x) => Math.sign(x) * Math.ceil(Math.abs(x) / 2) * 2), ODD: (c) => chk(numArg(c, 0), (x) => { const a = Math.ceil(Math.abs(x)); return Math.sign(x || 1) * (a % 2 ? a : a + 1); }),
  SUMPRODUCT: (c) => { const arrs = argVals(c); if (!arrs.length) return 0; const a0 = arrs[0]; if (!isArr(a0)) return chk(numsOf(arrs), (a) => a.reduce((x, y) => x * y, 1)); let s = 0; for (let i = 0; i < a0.rows; i++) for (let j = 0; j < a0.cols; j++) { let p = 1; for (const a of arrs) { const v = isArr(a) ? a.get(i, j) : a; p *= typeof v === "number" ? v : 0; } s += p; } return s; },
  SUMIF: (c) => { const rng = c.eval(0)!, crit = need(c, 1), sumR = c.eval(2); if (!isArr(rng)) return 0; const test = makeCriteria(crit); let s = 0; const sr = c.n > 2 ? (c.engine.anchoredArr(c.node(2), rng.rows, rng.cols, c.sheet) ?? (sumR !== undefined && isArr(sumR) ? sumR : rng)) : rng; for (let i = 0; i < rng.rows; i++) for (let j = 0; j < rng.cols; j++) if (test(rng.get(i, j))) { const v = sr.get(i, j); if (typeof v === "number") s += v; } return s; },
  AVERAGEIF: (c) => { const rng = c.eval(0)!, crit = need(c, 1), avgR = c.eval(2); if (!isArr(rng)) return DIV0; const test = makeCriteria(crit); let s = 0, n = 0; const sr = c.n > 2 ? (c.engine.anchoredArr(c.node(2), rng.rows, rng.cols, c.sheet) ?? (avgR !== undefined && isArr(avgR) ? avgR : rng)) : rng; for (let i = 0; i < rng.rows; i++) for (let j = 0; j < rng.cols; j++) if (test(rng.get(i, j))) { const v = sr.get(i, j); if (typeof v === "number") { s += v; n++; } } return n ? s / n : DIV0; },
  COUNTIF: (c) => { const rng = c.eval(0)!, crit = need(c, 1); if (!isArr(rng)) return makeCriteria(crit)(rng) ? 1 : 0; const test = makeCriteria(crit); let n = 0; for (let i = 0; i < rng.rows; i++) for (let j = 0; j < rng.cols; j++) if (test(rng.get(i, j))) n++; return n; },
  SUMIFS: (c) => multiIf(c, "sum"), COUNTIFS: (c) => multiIf(c, "count"), AVERAGEIFS: (c) => multiIf(c, "avg"), MAXIFS: (c) => multiIf(c, "max"), MINIFS: (c) => multiIf(c, "min"),
  IF: (c) => { const t = toBool(need(c, 0)); if (isError(t)) return t; return t ? (c.n > 1 ? c.eval(1)! : true) : (c.n > 2 ? c.eval(2)! : false); },
  IFS: (c) => { for (let i = 0; i + 1 < c.n; i += 2) { const t = toBool(need(c, i)); if (isError(t)) return t; if (t) return c.eval(i + 1)!; } return NA; },
  IFERROR: (c) => { const v = c.eval(0)!; return isError(first(v)) ? c.eval(1)! : v; },
  IFNA: (c) => { const v = c.eval(0)!; const f = first(v); return isError(f) && f.e === "#N/A" ? c.eval(1)! : v; },
  AND: (c) => { let any = false; for (const v of argVals(c)) for (const x of iterate(v)) { if (x === null || x === "" || (isArr(v) && typeof x === "string")) continue; const b = toBool(x); if (isError(b)) return b; any = true; if (!b) return false; } return any ? true : VALUE; },
  OR: (c) => { let any = false; for (const v of argVals(c)) for (const x of iterate(v)) { if (x === null || x === "" || (isArr(v) && typeof x === "string")) continue; const b = toBool(x); if (isError(b)) return b; any = true; if (b) return true; } return any ? false : VALUE; },
  NOT: (c) => chk(toBool(need(c, 0)), (b) => !b), TRUE: () => true, FALSE: () => false,
  ISBLANK: (c) => { const v = need(c, 0); return v === null; }, ISNUMBER: (c) => typeof need(c, 0) === "number", ISTEXT: (c) => typeof need(c, 0) === "string", ISERROR: (c) => isError(need(c, 0)), ISNA: (c) => { const v = need(c, 0); return isError(v) && v.e === "#N/A"; }, ISLOGICAL: (c) => typeof need(c, 0) === "boolean", ISEVEN: (c) => chk(numArg(c, 0), (x) => Math.floor(x) % 2 === 0), ISODD: (c) => chk(numArg(c, 0), (x) => Math.floor(x) % 2 !== 0),
  NA: () => NA, N: (c) => { const v = need(c, 0); return typeof v === "number" ? v : typeof v === "boolean" ? (v ? 1 : 0) : 0; },
  LEN: (c) => chk(strArg(c, 0), (s) => s.length), LEFT: (c) => chk(strArg(c, 0), (s) => chk(numArg(c, 1, 1), (n) => s.slice(0, Math.max(0, n)))),
  RIGHT: (c) => chk(strArg(c, 0), (s) => chk(numArg(c, 1, 1), (n) => (n <= 0 ? "" : s.slice(-n)))),
  MID: (c) => chk(strArg(c, 0), (s) => chk(numArg(c, 1), (st) => chk(numArg(c, 2), (n) => (st < 1 || n < 0 ? VALUE : s.substr(st - 1, n))))),
  TRIM: (c) => chk(strArg(c, 0), (s) => s.trim().replace(/\s+/g, " ")), UPPER: (c) => chk(strArg(c, 0), (s) => s.toLocaleUpperCase(locale().id === "tr" ? "tr" : "en")), LOWER: (c) => chk(strArg(c, 0), (s) => s.toLocaleLowerCase(locale().id === "tr" ? "tr" : "en")),
  PROPER: (c) => chk(strArg(c, 0), (s) => s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_m, a, b) => a + b.toUpperCase())),
  CONCAT: (c) => { let s = ""; for (const v of argVals(c)) for (const x of iterate(v)) { if (isError(x)) return x; s += toStr(x); } return s; },
  CONCATENATE: (c) => { let s = ""; for (const v of argVals(c)) { const x = first(v); if (isError(x)) return x; s += toStr(x); } return s; },
  TEXTJOIN: (c) => { const d = strArg(c, 0); if (isError(d)) return d; const skip = toBool(need(c, 1)); const parts: string[] = []; for (const v of argVals(c, 2)) for (const x of iterate(v)) { if (isError(x)) return x; if (skip === true && (x === null || x === "")) continue; parts.push(toStr(x)); } return parts.join(d); },
  SUBSTITUTE: (c) => chk(strArg(c, 0), (s) => chk(strArg(c, 1), (o) => chk(strArg(c, 2), (n) => { const inst = c.n > 3 ? numArg(c, 3) : null; if (inst !== null && !isError(inst)) { let idx = -1; for (let k = 0; k < inst; k++) { idx = s.indexOf(o, idx + 1); if (idx < 0) return s; } return s.slice(0, idx) + n + s.slice(idx + o.length); } return o ? s.split(o).join(n) : s; }))),
  REPLACE: (c) => chk(strArg(c, 0), (s) => chk(numArg(c, 1), (st) => chk(numArg(c, 2), (n) => chk(strArg(c, 3), (nw) => s.slice(0, st - 1) + nw + s.slice(st - 1 + n))))),
  FIND: (c) => chk(strArg(c, 0), (f) => chk(strArg(c, 1), (s) => chk(numArg(c, 2, 1), (st) => { const i = s.indexOf(f, st - 1); return i < 0 ? VALUE : i + 1; }))),
  SEARCH: (c) => chk(strArg(c, 0), (f) => chk(strArg(c, 1), (s) => chk(numArg(c, 2, 1), (st) => { const re = new RegExp(f.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, "."), "i"); const m = re.exec(s.slice(st - 1)); return m ? m.index + st : VALUE; }))),
  EXACT: (c) => chk(strArg(c, 0), (a) => chk(strArg(c, 1), (b) => a === b)), REPT: (c) => chk(strArg(c, 0), (s) => chk(numArg(c, 1), (n) => s.repeat(Math.max(0, n)))),
  CHAR: (c) => chk(numArg(c, 0), (n) => String.fromCharCode(n)), CODE: (c) => chk(strArg(c, 0), (s) => (s ? s.charCodeAt(0) : VALUE)), CLEAN: (c) => chk(strArg(c, 0), (s) => s.replace(/[\x00-\x1f]/g, "")),
  TEXT: (c) => { const v = need(c, 0); const f = strArg(c, 1); if (isError(f)) return f; if (isError(v)) return v; return formatValue(v as Value, f).text; },
  VALUE: (c) => chk(strArg(c, 0), (s) => { const p = parseInput(s, locale()); return typeof p.value === "number" ? p.value : VALUE; }),
  NUMBERVALUE: (c) => chk(strArg(c, 0), (s) => { const p = parseInput(s, locale()); return typeof p.value === "number" ? p.value : VALUE; }),
  T: (c) => { const v = need(c, 0); return typeof v === "string" ? v : ""; },
  TODAY: (c) => todaySerial(!!c.engine.wb.date1904), NOW: (c) => nowSerial(!!c.engine.wb.date1904),
  DATE: (c) => chk(numArg(c, 0), (y) => chk(numArg(c, 1), (m) => chk(numArg(c, 2), (d) => makeDate(y < 1900 ? y + 1900 : y, m - 1, d, c)))),
  YEAR: (c) => chk(numArg(c, 0), (s) => dateParts(s, c).y), MONTH: (c) => chk(numArg(c, 0), (s) => dateParts(s, c).m + 1), DAY: (c) => chk(numArg(c, 0), (s) => dateParts(s, c).d),
  HOUR: (c) => chk(numArg(c, 0), (s) => Math.floor(((s % 1) * 24 + 1e-9) % 24)), MINUTE: (c) => chk(numArg(c, 0), (s) => Math.floor(((s % 1) * 1440 + 1e-9) % 60)), SECOND: (c) => chk(numArg(c, 0), (s) => Math.floor(((s % 1) * 86400 + 1e-9) % 60)),
  TIME: (c) => chk(numArg(c, 0), (h) => chk(numArg(c, 1), (m) => chk(numArg(c, 2), (s) => (h * 3600 + m * 60 + s) / 86400))),
  DAYS: (c) => chk(numArg(c, 0), (e) => chk(numArg(c, 1), (s) => Math.floor(e) - Math.floor(s))),
  WEEKDAY: (c) => chk(numArg(c, 0), (s) => chk(numArg(c, 1, 1), (t) => { const wd = dateParts(s, c).wd; if (t === 1 || t === 17) return wd + 1; if (t === 2 || t === 11) return ((wd + 6) % 7) + 1; if (t === 3) return (wd + 6) % 7; if (t >= 12 && t <= 16) { const start = t - 10; return ((wd - start + 7) % 7) + 1; } return NUM; })),
  EDATE: (c) => chk(numArg(c, 0), (s) => chk(numArg(c, 1), (m) => { const p = dateParts(s, c); const d = new Date(Date.UTC(p.y, p.m + m, 1)); const last = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0)).getUTCDate(); return makeDate(d.getUTCFullYear(), d.getUTCMonth(), Math.min(p.d, last), c); })),
  EOMONTH: (c) => chk(numArg(c, 0), (s) => chk(numArg(c, 1), (m) => { const p = dateParts(s, c); const d = new Date(Date.UTC(p.y, p.m + m + 1, 0)); return makeDate(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), c); })),
  DATEVALUE: (c) => chk(strArg(c, 0), (s) => { const p = parseInput(s, locale()); return typeof p.value === "number" ? Math.floor(p.value) : VALUE; }),
  DATEDIF: (c) => chk(numArg(c, 0), (a) => chk(numArg(c, 1), (b) => chk(strArg(c, 2), (u) => { const pa = dateParts(a, c), pb = dateParts(b, c); if (b < a) return NUM; const U = u.toUpperCase(); if (U === "D") return Math.floor(b) - Math.floor(a); let months = (pb.y - pa.y) * 12 + (pb.m - pa.m); if (pb.d < pa.d) months--; if (U === "M") return months; if (U === "Y") return Math.floor(months / 12); if (U === "YM") return months % 12; return VALUE; }))),
  NETWORKDAYS: (c) => chk(numArg(c, 0), (a) => chk(numArg(c, 1), (b) => { const hol = holidaySet(c, 2); let n = 0; const lo = Math.floor(Math.min(a, b)), hi = Math.floor(Math.max(a, b)); for (let s = lo; s <= hi; s++) { const wd = dateParts(s, c).wd; if (wd !== 0 && wd !== 6 && !hol.has(s)) n++; } return a <= b ? n : -n; })),
  WORKDAY: (c) => chk(numArg(c, 0), (a) => chk(numArg(c, 1), (days) => { const hol = holidaySet(c, 2); let s = Math.floor(a); const step = days < 0 ? -1 : 1; let left = Math.abs(Math.trunc(days)); while (left > 0) { s += step; const wd = dateParts(s, c).wd; if (wd !== 0 && wd !== 6 && !hol.has(s)) left--; } return s; })),
  VLOOKUP: (c) => { const v = need(c, 0), tbl = c.eval(1)!, col = numArg(c, 2); if (isError(col)) return col; if (!isArr(tbl)) return NA; const approx = c.n > 3 ? toBool(need(c, 3)) : true; if (col < 1 || col > tbl.cols) return REF; const i = approx === true ? lookupApprox(v, tbl, 0, true) : (() => { const idx = c.engine.exactIndex(c.node(1), tbl, c.sheet, true, 0); if (!idx) return lookupExact(v, { rows: tbl.rows, cols: 1, get: (r) => tbl.get(r, 0) }, 0); const k = c.engine.keyOf(v); return idx.has(k) ? idx.get(k)! : -1; })(); return i < 0 ? NA : tbl.get(i, col - 1); },
  HLOOKUP: (c) => { const v = need(c, 0), tbl = c.eval(1)!, row = numArg(c, 2); if (isError(row)) return row; if (!isArr(tbl)) return NA; const approx = c.n > 3 ? toBool(need(c, 3)) : true; if (row < 1 || row > tbl.rows) return REF; const i = approx === true ? lookupApprox(v, tbl, 0, false) : (() => { const idx = c.engine.exactIndex(c.node(1), tbl, c.sheet, false, 0); if (!idx) return lookupExact(v, { rows: 1, cols: tbl.cols, get: (_r, cc) => tbl.get(0, cc) }, 0); const k = c.engine.keyOf(v); return idx.has(k) ? idx.get(k)! : -1; })(); return i < 0 ? NA : tbl.get(row - 1, i); },
  XLOOKUP: (c) => { const v = need(c, 0), la = c.eval(1)!, ra = c.eval(2)!; if (!isArr(la) || !isArr(ra)) return NA; const mode = c.n > 4 ? numArg(c, 4, 0) : 0; if (isError(mode)) return mode; const i = mode === 0 ? (() => { const idx = c.engine.exactIndex(c.node(1), la, c.sheet, la.rows !== 1, 0); if (!idx) return lookupExact(v, la, 0); const k = c.engine.keyOf(v); return idx.has(k) ? idx.get(k)! : -1; })() : lookupExact(v, la, mode); if (i < 0) return c.n > 3 && c.eval(3) !== undefined && first(c.eval(3)!) !== null ? c.eval(3)! : NA; if (la.rows === 1 && la.cols > 1) { if (ra.rows > 1 && ra.cols === la.cols) return { rows: ra.rows, cols: 1, get: (r) => ra.get(r, i) }; return ra.get(0, Math.min(i, ra.cols - 1)); } if (ra.cols > 1) return { rows: 1, cols: ra.cols, get: (_r, cc) => ra.get(i, cc) }; return ra.get(Math.min(i, ra.rows - 1), 0); },
  INDEX: (c) => { const arr = c.eval(0)!; const r = numArg(c, 1, 0); if (isError(r)) return r; const cc = numArg(c, 2, 0); if (isError(cc)) return cc; if (r < 0 || cc < 0) return VALUE; if (!isArr(arr)) return r <= 1 && cc <= 1 ? arr : REF; if (r === 0 && cc === 0) return arr; if (r === 0) { if (cc > arr.cols) return REF; return { rows: arr.rows, cols: 1, get: (i) => arr.get(i, cc - 1) }; } if (cc === 0) { if (arr.rows === 1) return r > arr.cols ? REF : arr.get(0, r - 1); if (arr.cols === 1) return r > arr.rows ? REF : arr.get(r - 1, 0); if (r > arr.rows) return REF; return { rows: 1, cols: arr.cols, get: (_i, j) => arr.get(r - 1, j) }; } if (r > arr.rows || cc > arr.cols) return REF; return arr.get(r - 1, cc - 1); },
  MATCH: (c) => { const v = need(c, 0), arr = c.eval(1)!, type = numArg(c, 2, 1); if (isError(type)) return type; if (!isArr(arr)) return NA; const i = type === 0 ? ((/[*?]/.test(String(v)) && typeof v === "string") ? lookupExact(v, arr, 2) : (() => { const idx = c.engine.exactIndex(c.node(1), arr, c.sheet, arr.rows !== 1, 0); if (!idx) return lookupExact(v, arr, 0); const k = c.engine.keyOf(v); return idx.has(k) ? idx.get(k)! : -1; })()) : type > 0 ? lookupApprox(v, arr, 0, arr.rows > 1) : lookupExact(v, arr, 1); return i < 0 ? NA : i + 1; },
  CHOOSE: (c) => chk(numArg(c, 0), (i) => (i >= 1 && i < c.n ? c.eval(i)! : VALUE)),
  ROW: (c) => { const nd = c.node(0); if (!nd) return c.r + 1; if (nd.t !== "ref") return VALUE; const rf = nd.ref; return rf.r1 === rf.r2 ? rf.r1 + 1 : { rows: rf.r2 - rf.r1 + 1, cols: 1, get: (i) => rf.r1 + 1 + i }; },
  COLUMN: (c) => { const nd = c.node(0); if (!nd) return c.c + 1; if (nd.t !== "ref") return VALUE; const rf = nd.ref; return rf.c1 === rf.c2 ? rf.c1 + 1 : { rows: 1, cols: rf.c2 - rf.c1 + 1, get: (_i, j) => rf.c1 + 1 + j }; },
  ROWS: (c) => { const a = c.eval(0)!; return isArr(a) ? a.rows : 1; }, COLUMNS: (c) => { const a = c.eval(0)!; return isArr(a) ? a.cols : 1; },
  LARGE: (c) => { const a = numsOf([c.eval(0)!]); if (isErrObj(a)) return a; const k = numArg(c, 1); if (isError(k)) return k; const s = [...a].sort((x, y) => y - x); return k >= 1 && k <= s.length ? s[k - 1] : NUM; },
  SMALL: (c) => { const a = numsOf([c.eval(0)!]); if (isErrObj(a)) return a; const k = numArg(c, 1); if (isError(k)) return k; const s = [...a].sort((x, y) => x - y); return k >= 1 && k <= s.length ? s[k - 1] : NUM; },
  MEDIAN: (c) => chk(numsOf(argVals(c)), (a) => { if (!a.length) return NUM; const s = [...a].sort((x, y) => x - y); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; }),
  STDEV: (c) => chk(numsOf(argVals(c)), (a) => { if (a.length < 2) return DIV0; const m = a.reduce((x, y) => x + y, 0) / a.length; return Math.sqrt(a.reduce((x, y) => x + (y - m) ** 2, 0) / (a.length - 1)); }),
  "STDEV.S": (c) => FUNCS.STDEV(c), "STDEV.P": (c) => chk(numsOf(argVals(c)), (a) => { if (!a.length) return DIV0; const m = a.reduce((x, y) => x + y, 0) / a.length; return Math.sqrt(a.reduce((x, y) => x + (y - m) ** 2, 0) / a.length); }),
  RANK: (c) => chk(numArg(c, 0), (v) => { const a = numsOf([c.eval(1)!]); if (isErrObj(a)) return a; const desc = c.n > 2 ? numArg(c, 2, 0) === 0 : true; const s = [...a].sort((x, y) => (desc ? y - x : x - y)); const i = s.indexOf(v); return i < 0 ? NA : i + 1; }),
  UNIQUE: (c) => { const a = c.eval(0)!; if (!isArr(a)) return a; const seen = new Set<string>(); const rows: Val[][] = []; for (let i = 0; i < a.rows; i++) { const row: Val[] = []; for (let j = 0; j < a.cols; j++) row.push(a.get(i, j)); const k = JSON.stringify(row); if (!seen.has(k)) { seen.add(k); rows.push(row); } } return { rows: rows.length, cols: a.cols, get: (i, j) => rows[i]?.[j] ?? null }; },
  COUNTUNIQUE: (c) => { const seen = new Set<string>(); for (const v of argVals(c)) for (const x of iterate(v)) if (x !== null && x !== "") seen.add(typeof x === "string" ? x.toLowerCase() : JSON.stringify(x)); return seen.size; },
  ISFORMULA: (c) => { void c; return false; },
  HYPERLINK: (c) => (c.n > 1 ? c.eval(1)! : c.eval(0)!),
  TRANSPOSE: (c) => { const a = c.eval(0)!; if (!isArr(a)) return a; return { rows: a.cols, cols: a.rows, get: (i, j) => a.get(j, i) }; },
  ARRAYFORMULA: (c) => c.eval(0)!,
  SPLIT: (c) => chk(strArg(c, 0), (s) => chk(strArg(c, 1, ","), (d) => { const parts = s.split(d); return { rows: 1, cols: parts.length, get: (_i, j) => parts[j] ?? null }; })),
  LET: (c) => c.eval(c.n - 1)!,
  SUBTOTAL: (c) => { const fn = numArg(c, 0); if (isError(fn)) return fn; const ignoreHidden = fn > 100; const code = fn > 100 ? fn - 100 : fn; const vals: number[] = []; let cnta = 0; for (let i = 1; i < c.n; i++) { const nd = c.node(i); if (nd && nd.t === "ref") { const sh = c.engine.sheetFor(nd.ref.sheet, c.sheet); if (!sh) return REF; const r2 = Math.min(nd.ref.r2, Math.max(sh.maxRow, nd.ref.r1)), c2 = Math.min(nd.ref.c2, Math.max(sh.maxCol, nd.ref.c1)); for (let r = nd.ref.r1; r <= r2; r++) { if (sh.hiddenRowsByFilter.has(r) || (ignoreHidden && sh.rows.get(r)?.hidden)) continue; for (let cc = nd.ref.c1; cc <= c2; cc++) { const x = c.engine.cellValue(sh, r, cc); if (isError(x)) return x; if (x !== null && x !== "") cnta++; if (typeof x === "number") vals.push(x); } } } else { const v = c.eval(i)!; for (const x of iterate(v)) { if (isError(x)) return x; if (x !== null && x !== "") cnta++; if (typeof x === "number") vals.push(x); } } } return subtotalOf(code, vals, cnta); },
  OFFSET: (c) => { const nd = c.node(0); if (!nd || nd.t !== "ref") return VALUE; const dr = numArg(c, 1, 0), dc = numArg(c, 2, 0); if (isError(dr)) return dr; if (isError(dc)) return dc; const rf = nd.ref; const hv = c.eval(3), wv = c.eval(4); const h = hv !== undefined && first(hv) !== null ? toNum(first(hv)) : rf.r2 - rf.r1 + 1; const w = wv !== undefined && first(wv) !== null ? toNum(first(wv)) : rf.c2 - rf.c1 + 1; if (isError(h)) return h; if (isError(w)) return w; const r1 = rf.r1 + Math.trunc(dr), c1 = rf.c1 + Math.trunc(dc); if (r1 < 0 || c1 < 0 || h < 1 || w < 1 || r1 + h > MAXR || c1 + w > MAXC) return REF; const sh = c.engine.sheetFor(rf.sheet, c.sheet); if (!sh) return REF; if (h === 1 && w === 1) return c.engine.cellValue(sh, r1, c1); return { rows: h, cols: w, get: (i, j) => c.engine.cellValue(sh, r1 + i, c1 + j) }; },
  INDIRECT: (c) => chk(strArg(c, 0), (s) => { const rf = parseRef(s.trim().replace(/\$/g, "")); if (!rf) return REF; const sh = c.engine.sheetFor(rf.sheet, c.sheet); if (!sh) return REF; if (!rf.isRange) return c.engine.cellValue(sh, rf.r1, rf.c1); return c.engine.rangeOf(rf, c.sheet); }),
  SWITCH: (c) => { const v = need(c, 0); for (let i = 1; i + 1 < c.n; i += 2) { const k = need(c, i); if (typeof v === typeof k && compare(v, k) === 0) return c.eval(i + 1)!; } return c.n % 2 === 0 ? c.eval(c.n - 1)! : NA; },
  ISERR: (c) => { const v = need(c, 0); return isError(v) && v.e !== "#N/A"; },
  XMATCH: (c) => FUNCS.MATCH({ ...c, n: Math.min(c.n, 3), eval: (i) => (i === 2 ? (c.eval(2) === undefined ? 0 : c.eval(2)) : c.eval(i)) }),
  SEQUENCE: (c) => chk(numArg(c, 0), (rows) => chk(numArg(c, 1, 1), (cols) => chk(numArg(c, 2, 1), (start) => chk(numArg(c, 3, 1), (step) => ({ rows: Math.max(1, rows), cols: Math.max(1, cols), get: (i, j) => start + (i * Math.max(1, cols) + j) * step }))))),
};
// Turkish aliases for the most common functions
Object.assign(FUNCS, { TOPLA: FUNCS.SUM, ORTALAMA: FUNCS.AVERAGE, EĞER: FUNCS.IF, EGER: FUNCS.IF, DÜŞEYARA: FUNCS.VLOOKUP, DUSEYARA: FUNCS.VLOOKUP, YATAYARA: FUNCS.HLOOKUP, BİRLEŞTİR: FUNCS.CONCATENATE, BIRLESTIR: FUNCS.CONCATENATE, SOLDAN: FUNCS.LEFT, SAĞDAN: FUNCS.RIGHT, SAGDAN: FUNCS.RIGHT, UZUNLUK: FUNCS.LEN, YUVARLA: FUNCS.ROUND, EĞERSAY: FUNCS.COUNTIF, EGERSAY: FUNCS.COUNTIF, ETOPLA: FUNCS.SUMIF, ÇOKETOPLA: FUNCS.SUMIFS, COKETOPLA: FUNCS.SUMIFS, ÇOKEĞERSAY: FUNCS.COUNTIFS, BUGÜN: FUNCS.TODAY, BUGUN: FUNCS.TODAY, ŞİMDİ: FUNCS.NOW, SIMDI: FUNCS.NOW, TARİH: FUNCS.DATE, TARIH: FUNCS.DATE, YIL: FUNCS.YEAR, AY: FUNCS.MONTH, GÜN: FUNCS.DAY, GUN: FUNCS.DAY, MİN: FUNCS.MIN, MAK: FUNCS.MAX, BAĞ_DEĞ_SAY: FUNCS.COUNT, BAĞ_DEĞ_DOLU_SAY: FUNCS.COUNTA, İNDİS: FUNCS.INDEX, INDIS: FUNCS.INDEX, KAÇINCI: FUNCS.MATCH, KACINCI: FUNCS.MATCH, EĞERHATA: FUNCS.IFERROR, EGERHATA: FUNCS.IFERROR, VE: FUNCS.AND, YADA: FUNCS.OR, DEĞİL: FUNCS.NOT, METNEÇEVİR: FUNCS.TEXT, KIRP: FUNCS.TRIM, BÜYÜKHARF: FUNCS.UPPER, KÜÇÜKHARF: FUNCS.LOWER, PARÇAAL: FUNCS.MID, MBUL: FUNCS.FIND, BUL: FUNCS.FIND, YERİNEKOY: FUNCS.SUBSTITUTE, ÇAPRAZARA: FUNCS.XLOOKUP, CAPRAZARA: FUNCS.XLOOKUP, ÇOKETOPLA2: FUNCS.SUMIFS });

function multiIf(c: Ctx, kind: "sum" | "count" | "avg" | "max" | "min"): Any {
  const start = kind === "count" ? 0 : 1;
  const target = kind === "count" ? null : c.eval(0)!;
  if (target !== null && !isArr(target)) return VALUE;
  const pairs: { rng: Arr; test: (v: Val) => boolean }[] = [];
  for (let i = start; i + 1 < c.n; i += 2) { const rng = c.eval(i)!; if (!isArr(rng)) return VALUE; pairs.push({ rng, test: makeCriteria(need(c, i + 1)) }); }
  if (!pairs.length) return VALUE;
  const rows = pairs[0].rng.rows, cols = pairs[0].rng.cols;
  for (const p of pairs) if (p.rng.rows !== rows || p.rng.cols !== cols) return VALUE;
  if (target && ((target as Arr).rows !== rows || (target as Arr).cols !== cols)) return VALUE;
  let s = 0, n = 0, mx = -Infinity, mn = Infinity;
  for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
    let ok = true;
    for (const p of pairs) if (!p.test(p.rng.get(i, j))) { ok = false; break; }
    if (!ok) continue;
    if (kind === "count") { n++; continue; }
    const v = (target as Arr).get(i, j);
    if (typeof v !== "number") continue;
    s += v; n++; mx = Math.max(mx, v); mn = Math.min(mn, v);
  }
  if (kind === "count") return n;
  if (kind === "sum") return s;
  if (kind === "avg") return n ? s / n : DIV0;
  if (kind === "max") return n ? mx : 0;
  return n ? mn : 0;
}

export const FUNCTION_NAMES = Object.keys(FUNCS).filter((k) => /^[A-ZÇĞİÖŞÜ][A-Z0-9._ÇĞİÖŞÜ]*$/.test(k)).sort();
export type { Node };
export { toStr as valueToText };
