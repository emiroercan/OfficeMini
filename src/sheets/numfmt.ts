// Excel number formats: rendering values with format codes, locale-aware
// separators and month names, date serials, and parsing what the user types.
import { Value, isError } from "./model";

export interface Locale { id: string; decimal: string; thousands: string; dateOrder: "dmy" | "mdy" | "ymd"; dateSep: string; months: string[]; monthsShort: string[]; days: string[]; daysShort: string[]; currency: string; currencyPos: "prefix" | "suffix"; lcid: string; }

export const LOCALES: Record<string, Locale> = {
  tr: { id: "tr", decimal: ",", thousands: ".", dateOrder: "dmy", dateSep: ".", months: ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"], monthsShort: ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"], days: ["Pazar", "Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi"], daysShort: ["Paz", "Pzt", "Sal", "Çar", "Per", "Cum", "Cmt"], currency: "₺", currencyPos: "prefix", lcid: "41F" },
  us: { id: "us", decimal: ".", thousands: ",", dateOrder: "mdy", dateSep: "/", months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"], monthsShort: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], days: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], daysShort: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], currency: "$", currencyPos: "prefix", lcid: "409" },
  eu: { id: "eu", decimal: ",", thousands: ".", dateOrder: "dmy", dateSep: ".", months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"], monthsShort: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], days: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], daysShort: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], currency: "€", currencyPos: "suffix", lcid: "407" },
  uk: { id: "uk", decimal: ".", thousands: ",", dateOrder: "dmy", dateSep: "/", months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"], monthsShort: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], days: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"], daysShort: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], currency: "£", currencyPos: "prefix", lcid: "809" },
};

let current: Locale = LOCALES.tr;
export function setLocale(id: string) { current = LOCALES[id] || LOCALES.us; }
export function locale(): Locale { return current; }
export function detectLocale(): string {
  const l = (navigator.language || "en-US").toLowerCase();
  if (l.startsWith("tr")) return "tr";
  if (l.startsWith("en-us")) return "us";
  if (l.startsWith("en-gb")) return "uk";
  if (/^(de|fr|it|es|nl|pt|pl|cs|da|fi|nb|sv|ro|hu|el)/.test(l)) return "eu";
  return "us";
}

// ---- built-in formats -------------------------------------------------------

export const BUILTIN: Record<number, string> = {
  0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
  14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm",
  37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
};

/** Locale-dependent short date used for built-in id 14/22 (Excel shows the system format). */
function localeShortDate(): string {
  const l = current;
  if (l.dateOrder === "mdy") return "m/d/yyyy";
  if (l.dateOrder === "ymd") return "yyyy-mm-dd";
  return l.dateSep === "." ? "dd.mm.yyyy" : "dd/mm/yyyy";
}
export function formatCodeFor(id: number, custom: Map<number, string>): string {
  const c = custom.get(id);
  if (c !== undefined) return c;
  if (id === 14) return localeShortDate();
  if (id === 22) return localeShortDate() + " h:mm";
  return BUILTIN[id] ?? "General";
}

// ---- date serials -----------------------------------------------------------

const MS_DAY = 86400000;
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);

export function serialToDate(serial: number, date1904 = false): Date {
  let s = serial;
  if (!date1904 && s < 61) s += 1; // Lotus 1900 leap-year bug: serials before 1 Mar 1900 are off by one
  return new Date((date1904 ? EPOCH_1904 : EPOCH_1900) + s * MS_DAY);
}
export function dateToSerial(d: Date, date1904 = false): number {
  const ms = Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds(), d.getUTCMilliseconds());
  let s = (ms - (date1904 ? EPOCH_1904 : EPOCH_1900)) / MS_DAY;
  if (!date1904 && s < 61) s -= 1;
  return s;
}
export function todaySerial(date1904 = false): number { const n = new Date(); return Math.floor(dateToSerial(new Date(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate())), date1904)); }
export function nowSerial(date1904 = false): number { const n = new Date(); return dateToSerial(new Date(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate(), n.getHours(), n.getMinutes(), n.getSeconds())), date1904); }

// ---- format code parsing ----------------------------------------------------

interface Section {
  cond: { op: string; val: number } | null;
  color: string | null;
  tokens: Tok[];
  isDate: boolean;
  isText: boolean;
  general: boolean;
  percent: boolean;
  thousandsSep: boolean;   // "," between digits
  scale: number;           // trailing commas scale by 1000 each
  intDigits: number;       // count of 0/#/? before decimal point
  intMin: number;          // count of "0" before the decimal point (minimum digits, zero padded)
  fracDigits: number;
  fracMin: number;         // count of "0" after point
  exp: { plus: boolean; digits: number } | null;
  fraction: { num: string; den: string } | null;
  hasAmPm: boolean;
  elapsed: string | null;  // [h] [m] [s]
}
type Tok = { t: "lit"; s: string } | { t: "num" } | { t: "date"; s: string } | { t: "text" } | { t: "fill"; s: string } | { t: "pad"; s: string } | { t: "general" };

const cache = new Map<string, Section[]>();

function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = "", inQ = false, inB = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"' && !inB) inQ = !inQ;
    else if (!inQ && ch === "[") inB = true;
    else if (!inQ && ch === "]") inB = false;
    if (ch === ";" && !inQ && !inB) { out.push(cur); cur = ""; continue; }
    if (ch === "\\" && i + 1 < code.length && !inQ) { cur += ch + code[i + 1]; i++; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

const COLORS: Record<string, string> = { black: "#000000", blue: "#0000ff", cyan: "#00ffff", green: "#00b050", magenta: "#ff00ff", red: "#e00000", white: "#ffffff", yellow: "#c8a000" };

function parseSection(src: string): Section {
  const sec: Section = { cond: null, color: null, tokens: [], isDate: false, isText: false, general: false, percent: false, thousandsSep: false, scale: 0, intDigits: 0, intMin: 0, fracDigits: 0, fracMin: 0, exp: null, fraction: null, hasAmPm: false, elapsed: null };
  let i = 0;
  let numTokenAdded = false;
  let seenDigit = false, afterPoint = false;
  let pendingCommas = 0;
  const litBuf = { s: "" };
  const flushLit = () => { if (litBuf.s) { sec.tokens.push({ t: "lit", s: litBuf.s }); litBuf.s = ""; } };
  const addNum = () => { if (!numTokenAdded) { flushLit(); sec.tokens.push({ t: "num" }); numTokenAdded = true; } };
  const s = src;
  while (i < s.length) {
    const ch = s[i];
    if (ch === "[") {
      const end = s.indexOf("]", i);
      if (end < 0) { litBuf.s += ch; i++; continue; }
      const body = s.slice(i + 1, end);
      i = end + 1;
      const lower = body.toLowerCase();
      if (COLORS[lower]) { sec.color = COLORS[lower]; continue; }
      if (/^color\d+$/i.test(body)) { sec.color = "#000000"; continue; }
      const cm = /^([<>=!]+)\s*(-?[\d.]+)$/.exec(body);
      if (cm) { sec.cond = { op: cm[1], val: parseFloat(cm[2]) }; continue; }
      if (/^\$/.test(body)) {
        // [$₺-41F] currency symbol with locale, or [$-409] locale only
        const m = /^\$([^-]*)(?:-(.*))?$/.exec(body);
        if (m && m[1]) { litBuf.s += m[1]; }
        continue;
      }
      if (/^(h+|m+|s+)$/i.test(body)) { flushLit(); sec.tokens.push({ t: "date", s: "[" + body.toLowerCase() + "]" }); sec.isDate = true; sec.elapsed = body[0].toLowerCase(); continue; }
      continue; // unknown bracket: ignore
    }
    if (ch === '"') {
      const end = s.indexOf('"', i + 1);
      litBuf.s += s.slice(i + 1, end < 0 ? s.length : end);
      i = end < 0 ? s.length : end + 1;
      continue;
    }
    if (ch === "\\") { litBuf.s += s[i + 1] || ""; i += 2; continue; }
    if (ch === "_") { flushLit(); sec.tokens.push({ t: "pad", s: s[i + 1] || " " }); i += 2; continue; }
    if (ch === "*") { flushLit(); sec.tokens.push({ t: "fill", s: s[i + 1] || " " }); i += 2; continue; }
    if (ch === "@") { flushLit(); sec.tokens.push({ t: "text" }); sec.isText = true; i++; continue; }
    if (/^general/i.test(s.slice(i))) { flushLit(); sec.tokens.push({ t: "general" }); sec.general = true; i += 7; continue; }
    // date/time tokens
    const dm = /^(yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|am\/pm|A\/P|a\/p)/i.exec(s.slice(i));
    if (dm && !/^[0#?]/.test(ch)) {
      flushLit();
      let tok = dm[0];
      if (/^(am\/pm|a\/p)$/i.test(tok)) { sec.hasAmPm = true; tok = tok.length > 3 ? "AM/PM" : "A/P"; }
      sec.tokens.push({ t: "date", s: tok });
      sec.isDate = true;
      i += dm[0].length;
      continue;
    }
    if (ch === "0" || ch === "#" || ch === "?") {
      addNum();
      if (afterPoint) { sec.fracDigits++; if (ch === "0") sec.fracMin++; } else { sec.intDigits++; if (ch === "0") sec.intMin++; }
      if (pendingCommas && seenDigit) { sec.thousandsSep = true; pendingCommas = 0; }
      seenDigit = true;
      i++;
      continue;
    }
    if (ch === "." && !afterPoint && (seenDigit || /^[0#?]/.test(s[i + 1] || ""))) { addNum(); afterPoint = true; i++; continue; }
    if (ch === "," && seenDigit) {
      // comma between digits = thousands; trailing commas (also after the decimals) = scaling by 1000
      if (!afterPoint && /^[0#?]/.test(s[i + 1] || "")) sec.thousandsSep = true; else sec.scale++;
      i++; continue;
    }
    if (ch === "%") { flushLit(); sec.tokens.push({ t: "lit", s: "%" }); sec.percent = true; i++; continue; }
    if ((ch === "E" || ch === "e") && /^[+-]/.test(s[i + 1] || "") && seenDigit) {
      const plus = s[i + 1] === "+";
      let j = i + 2, digits = 0;
      while (j < s.length && /[0#?]/.test(s[j])) { digits++; j++; }
      sec.exp = { plus, digits: Math.max(1, digits) };
      i = j; continue;
    }
    if (ch === "/" && seenDigit && !afterPoint && /[?#0]/.test(s[i + 1] || "")) {
      // fraction: read denominator pattern
      let j = i + 1, den = "";
      while (j < s.length && /[?#0]/.test(s[j])) { den += s[j]; j++; }
      sec.fraction = { num: "?", den };
      i = j; continue;
    }
    litBuf.s += ch;
    i++;
  }
  flushLit();
  if (sec.tokens.some((t) => t.t === "date")) sec.isDate = true;
  return sec;
}

export function parseFormat(code: string): Section[] {
  let secs = cache.get(code);
  if (secs) return secs;
  secs = splitSections(code).map(parseSection);
  cache.set(code, secs);
  return secs;
}

// ---- rendering ----------------------------------------------------------------

function pad(n: number, len: number): string { let s = String(Math.abs(Math.trunc(n))); while (s.length < len) s = "0" + s; return s; }

function fmtGeneral(v: number): string {
  if (!isFinite(v)) return isError({ e: "" }) ? "#NUM!" : String(v);
  if (v === 0) return "0";
  const abs = Math.abs(v);
  let s: string;
  if (abs >= 1e11 || abs < 1e-9) s = v.toExponential(5).replace(/\.?0+e/, "e").replace("e+", "E+").replace("e-", "E-");
  else {
    // up to 11 significant digits like Excel's General
    const digits = Math.max(0, 10 - Math.floor(Math.log10(abs)));
    s = v.toFixed(Math.min(digits, 20));
    if (s.includes(".")) s = s.replace(/\.?0+$/, "");
  }
  return s.replace(".", current.decimal);
}

function groupInt(intStr: string): string {
  let out = "";
  for (let i = 0; i < intStr.length; i++) {
    const fromEnd = intStr.length - i;
    out += intStr[i];
    if (fromEnd > 1 && (fromEnd - 1) % 3 === 0) out += current.thousands;
  }
  return out;
}

function renderNumber(sec: Section, v: number): string {
  let n = Math.abs(v);
  if (sec.percent) n *= 100;
  if (sec.scale) n /= Math.pow(1000, sec.scale);
  let numText: string;
  if (sec.exp) {
    let e = n === 0 ? 0 : Math.floor(Math.log10(n));
    let mant = n / Math.pow(10, e);
    mant = parseFloat(mant.toFixed(sec.fracDigits));
    if (mant >= 10) { mant /= 10; e++; }
    let m = mant.toFixed(sec.fracDigits);
    if (sec.fracDigits > sec.fracMin) m = m.replace(/\.?0+$/, "");
    numText = m.replace(".", current.decimal) + "E" + (e < 0 ? "-" : sec.exp.plus ? "+" : "") + pad(Math.abs(e), sec.exp.digits);
  } else if (sec.fraction) {
    const whole = Math.floor(n);
    const frac = n - whole;
    const maxDen = Math.pow(10, sec.fraction.den.length) - 1;
    let bestN = 0, bestD = 1, bestErr = frac;
    for (let d = 1; d <= maxDen; d++) { const nn = Math.round(frac * d); const err = Math.abs(frac - nn / d); if (err < bestErr) { bestErr = err; bestN = nn; bestD = d; if (err === 0) break; } }
    numText = (whole ? whole + " " : "") + (bestN ? `${bestN}/${bestD}` : whole ? "" : "0");
  } else {
    const fixed = n.toFixed(Math.min(sec.fracDigits, 20));
    let [intPart, fracPart = ""] = fixed.split(".");
    if (sec.fracDigits > sec.fracMin) fracPart = fracPart.replace(new RegExp(`0{0,${sec.fracDigits - sec.fracMin}}$`), "");
    if (intPart === "0" && sec.intMin === 0) intPart = "";
    while (intPart.length < sec.intMin && intPart.length < 30) intPart = "0" + intPart;
    // Only "#"/"?" placeholders and a zero value: Excel shows nothing (e.g. the `"-"??` accounting section).
    if (n === 0 && sec.intMin === 0 && sec.fracMin === 0) { numText = ""; return numText; }
    if (sec.thousandsSep) intPart = groupInt(intPart);
    numText = intPart + (fracPart ? current.decimal + fracPart : sec.fracDigits > 0 && sec.fracMin > 0 ? "" : "");
    if (!numText) numText = "0";
  }
  return numText;
}

function renderDate(sec: Section, v: number, date1904: boolean): string {
  if (v < 0 || v >= 2958466) return "#".repeat(8);
  const totalSec = Math.round(v * 86400);
  const d = serialToDate(Math.floor(v), date1904);
  const daySec = ((totalSec % 86400) + 86400) % 86400;
  const h24 = Math.floor(daySec / 3600), mi = Math.floor((daySec % 3600) / 60), ss = daySec % 60;
  const y = d.getUTCFullYear(), mo = d.getUTCMonth(), dd = d.getUTCDate(), wd = d.getUTCDay();
  let out = "";
  const toks = sec.tokens;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.t === "lit") { out += t.s; continue; }
    if (t.t === "pad") { out += " "; continue; }
    if (t.t !== "date") continue;
    const s = t.s;
    const lower = s.toLowerCase();
    if (lower === "[h]") { out += String(Math.floor(v * 24)); continue; }
    if (lower === "[m]") { out += String(Math.floor(v * 1440)); continue; }
    if (lower === "[s]") { out += String(Math.floor(v * 86400)); continue; }
    if (lower === "yyyy") out += String(y);
    else if (lower === "yy") out += pad(y % 100, 2);
    else if (lower === "mmmmm") out += current.months[mo][0];
    else if (lower === "mmmm") out += current.months[mo];
    else if (lower === "mmm") out += current.monthsShort[mo];
    else if (lower === "mm" || lower === "m") {
      // minutes when adjacent to h/s, else month
      const prev = toks.slice(0, i).reverse().find((x) => x.t === "date") as { s: string } | undefined;
      const next = toks.slice(i + 1).find((x) => x.t === "date") as { s: string } | undefined;
      const isMinute = (prev && /^(h|hh|\[h\])$/i.test(prev.s)) || (next && /^(s|ss)$/i.test(next.s));
      if (isMinute) out += lower === "mm" ? pad(mi, 2) : String(mi);
      else out += lower === "mm" ? pad(mo + 1, 2) : String(mo + 1);
    }
    else if (lower === "dddd") out += current.days[wd];
    else if (lower === "ddd") out += current.daysShort[wd];
    else if (lower === "dd") out += pad(dd, 2);
    else if (lower === "d") out += String(dd);
    else if (lower === "hh" || lower === "h") { const h = sec.hasAmPm ? (h24 % 12 || 12) : h24; out += lower === "hh" ? pad(h, 2) : String(h); }
    else if (lower === "ss" || lower === "s") out += lower === "ss" ? pad(ss, 2) : String(ss);
    else if (lower === "am/pm") out += h24 < 12 ? "AM" : "PM";
    else if (lower === "a/p") out += h24 < 12 ? "A" : "P";
  }
  return out;
}

export interface Rendered { text: string; color: string | null; align: "left" | "right" | "center" | null; }

/** Format a cell value with an Excel format code. */
export function formatValue(v: Value, code: string, date1904 = false): Rendered {
  if (v === null || v === undefined) return { text: "", color: null, align: null };
  if (isError(v)) return { text: v.e, color: null, align: "center" };
  if (typeof v === "boolean") return { text: v ? "TRUE" : "FALSE", color: null, align: "center" };
  const secs = parseFormat(code || "General");
  if (typeof v === "string") {
    const textSec = secs.find((s) => s.isText) || (secs.length === 4 ? secs[3] : null);
    if (!textSec) return { text: v, color: null, align: "left" };
    let out = "";
    for (const t of textSec.tokens) { if (t.t === "lit") out += t.s; else if (t.t === "text") out += v; else if (t.t === "pad") out += " "; }
    return { text: out, color: textSec.color, align: "left" };
  }
  // number
  let sec: Section | undefined;
  let negHandled = false;
  const numSecs = secs.filter((s) => !s.isText || s.general);
  if (numSecs.some((s) => s.cond)) {
    sec = numSecs.find((s) => s.cond && cmp(v, s.cond.op, s.cond.val));
    if (!sec) sec = numSecs.find((s) => !s.cond);
    negHandled = false; // Excel keeps the minus sign in conditional sections
  } else if (numSecs.length >= 2) {
    if (v < 0) { sec = numSecs[1]; negHandled = true; }
    else if (v === 0 && numSecs.length >= 3) sec = numSecs[2];
    else sec = numSecs[0];
  } else sec = numSecs[0];
  if (!sec) return { text: fmtGeneral(v), color: null, align: "right" };
  if (sec.isDate) return { text: renderDate(sec, v, date1904), color: sec.color, align: "right" };
  if (sec.general || !sec.tokens.some((t) => t.t === "num")) {
    let out = "";
    for (const t of sec.tokens) { if (t.t === "lit") out += t.s; else if (t.t === "general") out += fmtGeneral(Math.abs(negHandled ? v : v)); else if (t.t === "pad") out += " "; }
    if (!sec.tokens.length) out = secs.length > 1 ? "" : fmtGeneral(v); // ";;;" style empty sections hide the number
    else if (!negHandled && v < 0 && !sec.tokens.some((t) => t.t === "general")) out = "-" + out;
    else if (sec.tokens.some((t) => t.t === "general") && v < 0 && !negHandled) out = out.replace(fmtGeneral(Math.abs(v)), fmtGeneral(v));
    return { text: out, color: sec.color, align: "right" };
  }
  let out = "";
  for (const t of sec.tokens) {
    if (t.t === "lit") out += t.s;
    else if (t.t === "num") out += renderNumber(sec, v);
    else if (t.t === "pad") out += " ";
  }
  if (v < 0 && !negHandled) out = "-" + out;
  return { text: out, color: sec.color, align: "right" };
}

function cmp(v: number, op: string, x: number): boolean {
  switch (op) { case "<": return v < x; case "<=": return v <= x; case ">": return v > x; case ">=": return v >= x; case "=": return v === x; case "<>": case "!=": return v !== x; }
  return false;
}

export function isDateFormat(code: string): boolean {
  return parseFormat(code).some((s) => s.isDate);
}

// ---- format presets (toolbar menu) ---------------------------------------------

export interface Preset { label: string; code: string; sample?: number; }
export function presets(): { group: string; items: Preset[] }[] {
  const l = current;
  const cur = (sym: string, pos: "prefix" | "suffix", dec = 2) => {
    const num = dec ? "#,##0.00" : "#,##0";
    return pos === "prefix" ? `"${sym}"${num}` : `${num} "${sym}"`;
  };
  return [
    { group: "General", items: [{ label: "Automatic", code: "General" }, { label: "Plain text", code: "@" }] },
    { group: "Numbers", items: [
      { label: "Number 1,234.56", code: "#,##0.00", sample: 1234.56 }, { label: "Number 1234.56", code: "0.00", sample: 1234.56 }, { label: "Integer 1,234", code: "#,##0", sample: 1234 },
      { label: "Percent 12.3%", code: "0.0%", sample: 0.123 }, { label: "Percent 12%", code: "0%", sample: 0.12 },
      { label: "Scientific 1.23E+03", code: "0.00E+00", sample: 1234 }, { label: "Accounting", code: `_("${l.currency}"* #,##0.00_);_("${l.currency}"* (#,##0.00);_("${l.currency}"* "-"??_);_(@_)`, sample: 1234.56 },
      { label: "Negative in red (1,234.56)", code: "#,##0.00;[Red](#,##0.00)", sample: -1234.56 },
    ] },
    { group: "Currency", items: [
      { label: "₺ Turkish lira", code: cur("₺", "prefix"), sample: 1234.5 }, { label: "TL suffix", code: cur("TL", "suffix"), sample: 1234.5 },
      { label: "$ US dollar", code: cur("$", "prefix"), sample: 1234.5 }, { label: "€ Euro", code: cur("€", "suffix"), sample: 1234.5 }, { label: "€ Euro prefix", code: cur("€", "prefix"), sample: 1234.5 },
      { label: "£ Pound", code: cur("£", "prefix"), sample: 1234.5 }, { label: "Rounded ₺", code: cur("₺", "prefix", 0), sample: 1234.5 },
    ] },
    { group: "Dates", items: [
      { label: "09.09.2026 (TR / EU)", code: "dd.mm.yyyy", sample: 46274 }, { label: "09/09/2026 (UK)", code: "dd/mm/yyyy", sample: 46274 }, { label: "9/9/2026 (US)", code: "m/d/yyyy", sample: 46274 },
      { label: "2026-09-09 (ISO)", code: "yyyy-mm-dd", sample: 46274 }, { label: "9 September 2026", code: "d mmmm yyyy", sample: 46274 }, { label: "9 Eyl 2026", code: "d mmm yyyy", sample: 46274 },
      { label: "Wednesday, 9 September 2026", code: "dddd, d mmmm yyyy", sample: 46274 }, { label: "Sep 2026", code: "mmm yyyy", sample: 46274 },
      { label: "Date and time", code: "dd.mm.yyyy hh:mm", sample: 46274.6 }, { label: "Time 14:30", code: "hh:mm", sample: 0.604 }, { label: "Time 2:30 PM", code: "h:mm AM/PM", sample: 0.604 }, { label: "Duration [h]:mm", code: "[h]:mm", sample: 1.5 },
    ] },
  ];
}

// ---- parsing user input --------------------------------------------------------

export interface Parsed { value: Value; format?: string; }

const MONTH_LOOKUP: Record<string, number> = {};
for (const loc of Object.values(LOCALES)) {
  loc.months.forEach((m, i) => { MONTH_LOOKUP[m.toLowerCase()] = i; });
  loc.monthsShort.forEach((m, i) => { MONTH_LOOKUP[m.toLowerCase()] = i; });
}

/** Interpret typed text as number / date / boolean / text using the active locale. */
function monthIndex(name: string): number | undefined {
  const a = MONTH_LOOKUP[name.toLocaleLowerCase("tr")];
  return a !== undefined ? a : MONTH_LOOKUP[name.toLowerCase()];
}

export function parseInput(text: string, loc: Locale = current, date1904 = false): Parsed {
  const s = text.trim();
  if (s === "") return { value: null };
  if (/^(true|false)$/i.test(s)) return { value: s.toLowerCase() === "true" };
  if (s.length <= 6 && /^[dDyY]/.test(s)) { const tl = s.toLocaleLowerCase("tr"); if (tl === "doğru" || tl === "yanlış") return { value: tl === "doğru" }; } // locale lowercasing is slow: only for candidates
  // percent
  let m = /^(-?)\s*([\d.,\s]+)\s*%$/.exec(s);
  if (m) { const n = parseLocaleNumber(m[2], loc); if (n !== null) return { value: (m[1] ? -n : n) / 100, format: /[.,]\d/.test(m[2]) ? "0.00%" : "0%" }; }
  // currency: symbol before or after, ₺ $ € £ TL USD EUR
  m = /^(-?)\s*(₺|\$|€|£|TL|USD|EUR|GBP|TRY)?\s*([\d.,\s]+)\s*(₺|\$|€|£|TL|USD|EUR|GBP|TRY)?$/i.exec(s);
  if (m && (m[2] || m[4])) {
    const n = parseLocaleNumber(m[3], loc);
    if (n !== null) {
      const sym = (m[2] || m[4] || "").toUpperCase();
      const symbol = sym === "TL" || sym === "TRY" || sym === "₺" ? "₺" : sym === "USD" || sym === "$" ? "$" : sym === "EUR" || sym === "€" ? "€" : "£";
      const suffix = !!m[4];
      return { value: m[1] ? -n : n, format: suffix ? `#,##0.00 "${symbol}"` : `"${symbol}"#,##0.00` };
    }
  }
  // A day.month.year pattern is a date even where "." is the thousands separator (09.09.2026).
  const dm = /^(\d{1,2})([./-])(\d{1,2})\2(\d{2}|\d{4})$/.exec(s);
  const looksLikeDate = !!dm && (+dm[3] >= 1 && +dm[3] <= 12 || +dm[1] >= 1 && +dm[1] <= 12) && +dm[1] <= 31 && +dm[3] <= 31;
  // plain number
  const n = looksLikeDate ? null : parseLocaleNumber(s, loc);
  if (n !== null) {
    const thousands = new RegExp("\\" + loc.thousands).test(s) && !/^[+-]?\d{1,3}([.,]\d+)?$/.test(s);
    return { value: n, format: thousands ? (/[.,]\d+$/.test(s.replace(new RegExp("\\" + loc.thousands, "g"), "")) && s.split(loc.decimal).length > 1 ? "#,##0.00" : "#,##0") : undefined };
  }
  // ISO date / datetime
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
  if (m) return dateResult(+m[1], +m[2] - 1, +m[3], m[4] ? +m[4] : 0, m[5] ? +m[5] : 0, m[6] ? +m[6] : 0, m[4] ? "yyyy-mm-dd hh:mm" : "yyyy-mm-dd", date1904);
  // numeric date with separators in locale order
  m = /^(\d{1,4})([./-])(\d{1,2})\2(\d{1,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
  if (m) {
    const a = +m[1], b = +m[3], c = +m[4];
    let y: number, mo: number, d: number;
    if (m[1].length === 4) { y = a; mo = b - 1; d = c; }
    else if (loc.dateOrder === "mdy" && m[2] === "/") { mo = a - 1; d = b; y = c; }
    else { d = a; mo = b - 1; y = c; }
    if (y < 100) y += y < 30 ? 2000 : 1900;
    if (mo >= 0 && mo < 12 && d >= 1 && d <= 31) {
      const sep = m[2];
      const code = m[1].length === 4 ? "yyyy-mm-dd" : loc.dateOrder === "mdy" && sep === "/" ? "m/d/yyyy" : `dd${sep}mm${sep}yyyy`;
      return dateResult(y, mo, d, m[5] ? +m[5] : 0, m[6] ? +m[6] : 0, m[7] ? +m[7] : 0, m[5] ? code + " hh:mm" : code, date1904);
    }
  }
  // "9 Eylül 2026", "9 Sep 2026", "September 9, 2026"
  m = /^(\d{1,2})\s+([\p{L}]+)\.?\s+(\d{4})$/u.exec(s);
  if (m && monthIndex(m[2]) !== undefined) return dateResult(+m[3], monthIndex(m[2])!, +m[1], 0, 0, 0, "d mmmm yyyy", date1904);
  m = /^([\p{L}]+)\.?\s+(\d{1,2}),?\s+(\d{4})$/u.exec(s);
  if (m && monthIndex(m[1]) !== undefined) return dateResult(+m[3], monthIndex(m[1])!, +m[2], 0, 0, 0, "mmmm d, yyyy", date1904);
  // time
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$/i.exec(s);
  if (m) {
    let h = +m[1];
    if (m[4]) { if (m[4].toLowerCase() === "pm" && h < 12) h += 12; if (m[4].toLowerCase() === "am" && h === 12) h = 0; }
    return { value: (h * 3600 + +m[2] * 60 + (m[3] ? +m[3] : 0)) / 86400, format: m[4] ? "h:mm AM/PM" : m[3] ? "hh:mm:ss" : "hh:mm" };
  }
  return { value: text };
}

function dateResult(y: number, mo: number, d: number, h: number, mi: number, s: number, code: string, date1904: boolean): Parsed {
  const serial = dateToSerial(new Date(Date.UTC(y, mo, d, h, mi, s)), date1904);
  return { value: serial, format: code };
}

/** Parse a number written with the locale's separators; also accepts the other convention when unambiguous. */
export function parseLocaleNumber(s: string, loc: Locale = current): number | null {
  let t = s.replace(/\s/g, "");
  if (!t || !/\d/.test(t) || !/^[+-]?[\d.,]+(e[+-]?\d+)?$/i.test(t)) return null;
  const dec = loc.decimal, tho = loc.thousands;
  const hasDec = t.includes(dec), hasTho = t.includes(tho);
  if (hasDec && hasTho) {
    // decimal must come last
    if (t.lastIndexOf(dec) < t.lastIndexOf(tho)) return null;
    t = t.split(tho).join("").replace(dec, ".");
  } else if (hasDec) {
    // a single separator: decimal, unless it looks like grouping (e.g. "1.234" in TR could be 1234 or 1,234...)
    const parts = t.split(dec);
    if (parts.length > 2) t = parts.join(""); // multiple => grouping in the other convention? treat as thousands
    else t = t.replace(dec, ".");
  } else if (hasTho) {
    const parts = t.split(tho);
    // "1,234" (US typed in TR) with exactly 3 digits after -> thousands; otherwise other-locale decimal
    if (parts.length === 2 && parts[1].length !== 3) t = parts.join(".");
    else t = parts.join("");
  }
  const n = Number(t);
  return isNaN(n) ? null : n;
}

/** Text shown in the formula bar / cell editor for a value. */
export function editText(v: Value, code: string, date1904 = false): string {
  if (v === null || v === undefined) return "";
  if (isError(v)) return v.e;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "string") return v;
  if (/%/.test(code.split(";")[0]) && !/"[^"]*%[^"]*"/.test(code)) {
    let p = String(Math.round(v * 100 * 1e10) / 1e10);
    if (/e/i.test(p)) p = (v * 100).toLocaleString("en-US", { useGrouping: false, maximumFractionDigits: 15 });
    return p.replace(".", current.decimal) + "%";
  }
  if (isDateFormat(code)) {
    const secs = parseFormat(code);
    const toks = secs.flatMap((s) => s.tokens.filter((t) => t.t === "date") as { s: string }[]);
    const isMinute = (i: number) => (toks[i - 1] && /^(h|hh|\[h\])$/i.test(toks[i - 1].s)) || (toks[i + 1] && /^(s|ss)$/i.test(toks[i + 1].s));
    const hasDate = toks.some((t, i) => /^(yy|yyyy|d|dd|ddd|dddd|mmm|mmmm|mmmmm)$/i.test(t.s) || (/^(m|mm)$/i.test(t.s) && !isMinute(i)));
    const elapsed = toks.some((t) => /^\[(h|m|s)\]$/i.test(t.s));
    if (elapsed) return formatValue(v, "[h]:mm:ss", date1904).text;
    if (!hasDate) return formatValue(v, "hh:mm:ss", date1904).text;
    const hasTime = toks.some((t) => /^(h|hh|s|ss|am\/pm|a\/p)$/i.test(t.s));
    const dateOnly = Number.isInteger(v);
    const dfmt = current.dateOrder === "mdy" ? "m/d/yyyy" : current.dateOrder === "ymd" ? "yyyy-mm-dd" : `dd${current.dateSep}mm${current.dateSep}yyyy`;
    return formatValue(v, hasTime && !dateOnly ? dfmt + " hh:mm:ss" : dfmt, date1904).text;
  }
  // numbers: full precision with the locale decimal separator
  let s = String(v);
  if (/e/i.test(s)) s = v.toLocaleString("en-US", { useGrouping: false, maximumFractionDigits: 15 });
  return s.replace(".", current.decimal);
}
