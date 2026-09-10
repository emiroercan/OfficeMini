// CSV / TSV reading and writing with delimiter and encoding detection.
import { Sheet, Workbook, newSheet, key, updateExtent, Value } from "./model";
import { parseInput, locale, Locale } from "./numfmt";
import { Package } from "../docx/zip";
import { emptyStyles } from "./xlsx-read";

export interface CsvOptions { delimiter: string; encoding: string; bom: boolean; }

export function decodeBytes(bytes: Uint8Array): { text: string; encoding: string; bom: boolean } {
  const bom = bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
  if (bom) return { text: new TextDecoder("utf-8").decode(bytes.subarray(3)), encoding: "utf-8", bom: true };
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) return { text: new TextDecoder("utf-16le").decode(bytes.subarray(2)), encoding: "utf-16le", bom: true };
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return { text: new TextDecoder("utf-16be").decode(bytes.subarray(2)), encoding: "utf-16be", bom: true };
  try {
    return { text: new TextDecoder("utf-8", { fatal: true }).decode(bytes), encoding: "utf-8", bom: false };
  } catch {
    // Not valid UTF-8: Turkish Windows exports are usually Windows-1254.
    return { text: new TextDecoder("windows-1254").decode(bytes), encoding: "windows-1254", bom: false };
  }
}

/** Pick the delimiter whose field count is most consistent over the first lines (title lines and decimal commas tolerated). */
export function detectDelimiter(text: string): string {
  const lines: string[] = [];
  let pos = 0;
  while (lines.length < 40 && pos < text.length) { let nl = text.indexOf("\n", pos); if (nl < 0) nl = text.length; const line = text.slice(pos, nl).replace(/\r$/, ""); pos = nl + 1; if (line.trim()) lines.push(line); }
  const dec = locale().decimal;
  let best = dec === "," ? ";" : ",", bestScore = -1;
  for (const d of [",", ";", "\t", "|"]) {
    const counts = lines.map((l) => countOutsideQuotes(l, d)).filter((n) => n > 0);
    if (!counts.length) continue;
    const freq = new Map<number, number>();
    for (const n of counts) freq.set(n, (freq.get(n) || 0) + 1);
    let mode = 0, modeN = 0;
    for (const [n, f] of freq) if (f > modeN || (f === modeN && n > mode)) { mode = n; modeN = f; }
    let score = modeN * 10 + Math.min(mode, 9) + counts.length;
    if (d === dec) score -= 8; // splitting on the locale's decimal separator is unlikely to be right
    if (score > bestScore) { bestScore = score; best = d; }
  }
  return best;
}

let cp1254: Map<string, number> | null = null;
/** Encode CSV text in the encoding the file was read with (Windows-1254 stays Windows-1254); falls back to UTF-8 when a character does not fit. */
export function encodeCsv(text: string, encoding: string, bom: boolean): Uint8Array {
  if (encoding === "windows-1254") {
    if (!cp1254) { cp1254 = new Map(); const dec = new TextDecoder("windows-1254"); for (let b = 0; b < 256; b++) cp1254.set(dec.decode(new Uint8Array([b])), b); }
    const out = new Uint8Array(text.length);
    let n = 0, lossy = false;
    for (const ch of text) { const b = cp1254.get(ch); if (b === undefined) { lossy = true; break; } out[n++] = b; }
    if (!lossy) return out.subarray(0, n);
  }
  if (encoding === "utf-16le") { const buf = new Uint8Array(2 + text.length * 2); buf[0] = 0xff; buf[1] = 0xfe; for (let i = 0; i < text.length; i++) { const code = text.charCodeAt(i); buf[2 + i * 2] = code & 0xff; buf[3 + i * 2] = code >> 8; } return buf; }
  const enc = new TextEncoder().encode(text);
  if (!bom) return enc;
  const out = new Uint8Array(enc.length + 3); out.set([0xef, 0xbb, 0xbf]); out.set(enc, 3);
  return out;
}

function countOutsideQuotes(line: string, d: string): number {
  let n = 0, q = false;
  for (const ch of line) { if (ch === '"') q = !q; else if (ch === d && !q) n++; }
  return n;
}

/** RFC 4180 parser (quotes, doubled quotes, embedded newlines). */
export function parseCsv(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let i = 0, q = false;
  const n = text.length;
  while (i < n) {
    const ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i += 2; continue; } q = false; i++; continue; }
      field += ch; i++; continue;
    }
    if (ch === '"') { q = true; i++; continue; }
    if (ch === delimiter) { row.push(field); field = ""; i++; continue; }
    if (ch === "\r") { if (text[i + 1] !== "\n") { row.push(field); rows.push(row); row = []; field = ""; } i++; continue; } // CRLF or classic Mac CR
    if (ch === "\n") { row.push(field); rows.push(row); row = []; field = ""; i++; continue; }
    // fast path: consume a run of plain characters
    let j = i + 1;
    while (j < n) { const c = text[j]; if (c === '"' || c === delimiter || c === "\n" || c === "\r") break; j++; }
    field += text.slice(i, j);
    i = j;
  }
  if (field !== "" || row.length) { row.push(field); rows.push(row); }
  return rows;
}

/** Convert a text field to a typed value; numbers keep a format when they carry separators/dates. */
function typedValue(s: string, loc: Locale, formats: Map<string, number>, wb: Workbook): { v: Value; s: number } {
  if (s === "") return { v: null, s: 0 };
  // Keep leading zeros / identifiers as text (barcodes, IDs)
  if (/^0\d+$/.test(s) || /^\d{13,}$/.test(s)) return { v: s, s: 0 };
  if ((/^[\d.,\s+%₺$€£-]+$/.test(s) || /^[\d.,\s+-]+\s*(TL|TRY|USD|EUR|GBP)$/i.test(s)) && /\d/.test(s)) {
    const p = parseInput(s, loc);
    if (typeof p.value === "number") return { v: p.value, s: p.format ? styleFor(p.format, formats, wb) : 0 };
    return { v: s, s: 0 };
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(s) || /^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}/.test(s)) {
    const p = parseInput(s, loc);
    if (typeof p.value === "number") return { v: p.value, s: p.format ? styleFor(p.format, formats, wb) : 0 };
  }
  if (/^(true|false)$/i.test(s)) return { v: s.toLowerCase() === "true", s: 0 };
  return { v: s, s: 0 };
}

function styleFor(code: string, cache: Map<string, number>, wb: Workbook): number {
  let idx = cache.get(code);
  if (idx !== undefined) return idx;
  // add a numFmt + xf
  let fmtId = 0;
  for (const [id, c] of wb.styles.numFmts) if (c === code) fmtId = id;
  if (!fmtId) {
    const builtin = Object.entries({ "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10, "m/d/yyyy": 14, "h:mm": 20, "hh:mm": 20, "h:mm:ss": 21, "@": 49 }).find(([c]) => c === code);
    if (builtin) fmtId = builtin[1];
    else { fmtId = 164; for (const id of wb.styles.numFmts.keys()) fmtId = Math.max(fmtId, id + 1); wb.styles.numFmts.set(fmtId, code); }
  }
  idx = wb.styles.xfs.length;
  wb.styles.xfs.push({ numFmtId: fmtId, fontId: 0, fillId: 0, borderId: 0, xfId: 0, alignment: null, protection: null, applyNumberFormat: true });
  cache.set(code, idx);
  return idx;
}

/** Build a one-sheet workbook from CSV bytes. */
export function csvToWorkbook(bytes: Uint8Array, name: string): Workbook {
  const { text, encoding, bom } = decodeBytes(bytes);
  const delimiter = detectDelimiter(text);
  const rows = parseCsv(text, delimiter);
  const wb = blankWorkbook(name.replace(/\.(csv|tsv|txt)$/i, "") || "Sheet1");
  wb.kind = "csv";
  wb.csvOptions = { delimiter, encoding, bom };
  const sheet = wb.sheets[0];
  const loc = locale();
  const fmtCache = new Map<string, number>();
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    for (let c = 0; c < row.length; c++) {
      const s = row[c];
      if (s === "") continue;
      const tv = typedValue(s, loc, fmtCache, wb);
      if (tv.v === null) continue;
      sheet.cells.set(key(r, c), { v: tv.v, s: tv.s });
      updateExtent(sheet, r, c);
    }
  }
  // Header row frozen by default for viewing.
  if (sheet.maxRow > 0) { sheet.freeze = { rows: 1, cols: 0 }; sheet.autoFreeze = true; }
  return wb;
}

export function blankWorkbook(sheetName = "Sheet1"): Workbook {
  const pkg = Package.empty();
  const sheet = newSheet(sheetName, 1, "rId1", "xl/worksheets/sheet1.xml");
  sheet.view.tabSelected = true;
  sheet.dirty = true;
  return {
    sheets: [sheet], active: 0, date1904: false, definedNames: [], styles: emptyStyles(), sst: [], sstRaw: [], sstIndex: new Map(),
    themeColors: ["000000", "FFFFFF", "44546A", "E7E6E6", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"],
    pkg, workbookPart: "xl/workbook.xml", workbookXml: "", workbookRels: "", calcChainPart: null, stylesPart: "xl/styles.xml", sstPart: null, path: null, kind: "xlsx",
  };
}

/** Serialize a sheet to CSV text (values as displayed text for numbers with formats, raw otherwise). */
export function sheetToCsv(sheet: Sheet, delimiter: string, render: (r: number, c: number) => string): string {
  const lines: string[] = [];
  const needsQuote = (s: string) => s.includes(delimiter) || s.includes('"') || s.includes("\n") || s.includes("\r");
  for (let r = 0; r <= sheet.maxRow; r++) {
    const fields: string[] = [];
    let last = -1;
    for (let c = 0; c <= sheet.maxCol; c++) {
      const cell = sheet.cells.get(key(r, c));
      const s = cell ? render(r, c) : "";
      fields.push(needsQuote(s) ? '"' + s.replace(/"/g, '""') + '"' : s);
      if (s !== "") last = c;
    }
    lines.push(fields.slice(0, last + 1).join(delimiter));
  }
  return lines.join("\r\n") + "\r\n";
}
