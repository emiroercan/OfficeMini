// Typed run/paragraph/table properties parsed from OOXML, plus CSS generation.
import { NS, child, children, wattr, wint, onOff } from "./xml";
import { fmt, twipsToPt, halfPtToPt } from "./units";

export interface Border { val: string; sz: number; color: string; space: number; }
export interface TabStop { val: string; pos: number; leader?: string; }

export interface RunProps {
  font?: string | null;
  size?: number | null; // half-points
  b?: boolean;
  i?: boolean;
  u?: string | null; // underline style, "none" clears
  strike?: boolean;
  dstrike?: boolean;
  color?: string | null; // RRGGBB or "auto"
  highlight?: string | null; // name or "none"
  shd?: string | null; // fill RRGGBB
  vertAlign?: string | null; // superscript | subscript | baseline
  caps?: boolean;
  smallCaps?: boolean;
  vanish?: boolean;
  spacing?: number | null; // twips
  rStyle?: string | null;
  rtl?: boolean;
  lang?: string | null;
}

export interface ParaProps {
  pStyle?: string | null;
  jc?: string | null;
  indLeft?: number | null;
  indRight?: number | null;
  indHanging?: number | null;
  indFirstLine?: number | null;
  spBefore?: number | null;
  spAfter?: number | null;
  spLine?: number | null;
  spLineRule?: string | null;
  spBeforeAuto?: boolean;
  spAfterAuto?: boolean;
  contextual?: boolean;
  numId?: number | null;
  ilvl?: number | null;
  shd?: string | null;
  bdrTop?: Border | null;
  bdrBottom?: Border | null;
  bdrLeft?: Border | null;
  bdrRight?: Border | null;
  bdrBetween?: Border | null;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  tabs?: TabStop[] | null;
  bidi?: boolean;
  outlineLvl?: number | null;
  rPr?: RunProps | null; // paragraph mark run properties
}

export interface TableProps {
  tblStyle?: string | null;
  width?: { w: number; type: string } | null;
  jc?: string | null;
  indent?: number | null; // twips
  borders?: Partial<Record<"top" | "bottom" | "left" | "right" | "insideH" | "insideV", Border>> | null;
  cellMar?: { top?: number; bottom?: number; left?: number; right?: number } | null; // twips
  layout?: string | null; // fixed | autofit
  look?: string | null;
  shd?: string | null;
  cellSpacing?: number | null;
}

export interface RowProps {
  height?: number | null; // twips
  hRule?: string | null; // exact | atLeast | auto
  header?: boolean;
  cantSplit?: boolean;
  jc?: string | null;
}

export interface CellProps {
  width?: { w: number; type: string } | null;
  gridSpan?: number;
  vMerge?: string | null; // "restart" | "continue"
  borders?: Partial<Record<"top" | "bottom" | "left" | "right", Border>> | null;
  shd?: string | null;
  vAlign?: string | null;
  mar?: { top?: number; bottom?: number; left?: number; right?: number } | null;
  noWrap?: boolean;
  textDirection?: string | null;
}

export function parseBorder(el: Element | null): Border | null {
  if (!el) return null;
  const val = wattr(el, "val") || "single";
  return {
    val,
    sz: wint(el, "sz") ?? 4,
    color: wattr(el, "color") || "auto",
    space: wint(el, "space") ?? 0,
  };
}

/** Parse w:rFonts + theme references. `theme` gives the resolved major/minor latin fonts. */
export function parseFont(rFonts: Element | null, theme: { major: string; minor: string }): string | null | undefined {
  if (!rFonts) return undefined;
  const at = wattr(rFonts, "asciiTheme") || wattr(rFonts, "hAnsiTheme");
  if (at) {
    if (at.startsWith("major")) return theme.major;
    if (at.startsWith("minor")) return theme.minor;
  }
  const f = wattr(rFonts, "ascii") || wattr(rFonts, "hAnsi") || wattr(rFonts, "cs") || wattr(rFonts, "eastAsia");
  return f || undefined;
}

export function parseRPr(el: Element | null, theme: { major: string; minor: string }): RunProps {
  const p: RunProps = {};
  if (!el) return p;
  for (const c of children(el, NS.w)) {
    switch (c.localName) {
      case "rFonts": { const f = parseFont(c, theme); if (f !== undefined) p.font = f; break; }
      case "sz": { const v = wint(c, "val"); if (v !== null) p.size = v; break; }
      case "b": p.b = onOff(c) ?? true; break;
      case "i": p.i = onOff(c) ?? true; break;
      case "u": p.u = wattr(c, "val") || "single"; break;
      case "strike": p.strike = onOff(c) ?? true; break;
      case "dstrike": p.dstrike = onOff(c) ?? true; break;
      case "color": p.color = wattr(c, "val") || "auto"; break;
      case "highlight": p.highlight = wattr(c, "val") || "none"; break;
      case "shd": {
        const fill = wattr(c, "fill");
        const val = wattr(c, "val");
        if (fill && fill.toLowerCase() !== "auto") p.shd = fill;
        else if (val && val !== "clear" && val !== "nil") p.shd = "D9D9D9";
        break;
      }
      case "vertAlign": p.vertAlign = wattr(c, "val"); break;
      case "caps": p.caps = onOff(c) ?? true; break;
      case "smallCaps": p.smallCaps = onOff(c) ?? true; break;
      case "vanish": p.vanish = onOff(c) ?? true; break;
      case "spacing": { const v = wint(c, "val"); if (v !== null) p.spacing = v; break; }
      case "rStyle": p.rStyle = wattr(c, "val"); break;
      case "rtl": p.rtl = onOff(c) ?? true; break;
      case "lang": p.lang = wattr(c, "val"); break;
    }
  }
  return p;
}

export function parsePPr(el: Element | null, theme: { major: string; minor: string }): ParaProps {
  const p: ParaProps = {};
  if (!el) return p;
  for (const c of children(el, NS.w)) {
    switch (c.localName) {
      case "pStyle": p.pStyle = wattr(c, "val"); break;
      case "jc": p.jc = wattr(c, "val"); break;
      case "ind": {
        const l = wint(c, "left") ?? wint(c, "start");
        const r = wint(c, "right") ?? wint(c, "end");
        const h = wint(c, "hanging");
        const f = wint(c, "firstLine");
        if (l !== null) p.indLeft = l;
        if (r !== null) p.indRight = r;
        if (h !== null) { p.indHanging = h; p.indFirstLine = null; }
        else if (f !== null) { p.indFirstLine = f; p.indHanging = null; }
        break;
      }
      case "spacing": {
        const b = wint(c, "before"), a = wint(c, "after"), l = wint(c, "line");
        if (b !== null) p.spBefore = b;
        if (a !== null) p.spAfter = a;
        if (l !== null) p.spLine = l;
        const lr = wattr(c, "lineRule"); if (lr) p.spLineRule = lr;
        const ba = wattr(c, "beforeAutospacing"); if (ba !== null) p.spBeforeAuto = ba === "1" || ba === "true";
        const aa = wattr(c, "afterAutospacing"); if (aa !== null) p.spAfterAuto = aa === "1" || aa === "true";
        break;
      }
      case "contextualSpacing": p.contextual = onOff(c) ?? true; break;
      case "numPr": {
        const numId = wint(child(c, NS.w, "numId"), "val");
        const ilvl = wint(child(c, NS.w, "ilvl"), "val");
        if (numId !== null) p.numId = numId;
        p.ilvl = ilvl ?? 0;
        break;
      }
      case "shd": {
        const fill = wattr(c, "fill");
        if (fill && fill.toLowerCase() !== "auto") p.shd = fill;
        break;
      }
      case "pBdr": {
        p.bdrTop = parseBorder(child(c, NS.w, "top"));
        p.bdrBottom = parseBorder(child(c, NS.w, "bottom"));
        p.bdrLeft = parseBorder(child(c, NS.w, "left"));
        p.bdrRight = parseBorder(child(c, NS.w, "right"));
        p.bdrBetween = parseBorder(child(c, NS.w, "between"));
        break;
      }
      case "keepNext": p.keepNext = onOff(c) ?? true; break;
      case "keepLines": p.keepLines = onOff(c) ?? true; break;
      case "pageBreakBefore": p.pageBreakBefore = onOff(c) ?? true; break;
      case "widowControl": p.widowControl = onOff(c) ?? true; break;
      case "tabs": {
        const tabs: TabStop[] = [];
        for (const t of children(c, NS.w, "tab")) {
          const pos = wint(t, "pos");
          if (pos === null) continue;
          tabs.push({ val: wattr(t, "val") || "left", pos, leader: wattr(t, "leader") || undefined });
        }
        p.tabs = tabs;
        break;
      }
      case "bidi": p.bidi = onOff(c) ?? true; break;
      case "outlineLvl": p.outlineLvl = wint(c, "val"); break;
      case "rPr": p.rPr = parseRPr(c, theme); break;
    }
  }
  return p;
}

function parseMar(el: Element | null): { top?: number; bottom?: number; left?: number; right?: number } | null {
  if (!el) return null;
  const m: { top?: number; bottom?: number; left?: number; right?: number } = {};
  for (const c of children(el, NS.w)) {
    const w = wint(c, "w");
    if (w === null) continue;
    const type = wattr(c, "type");
    if (type && type !== "dxa") continue;
    if (c.localName === "top") m.top = w;
    else if (c.localName === "bottom") m.bottom = w;
    else if (c.localName === "left" || c.localName === "start") m.left = w;
    else if (c.localName === "right" || c.localName === "end") m.right = w;
  }
  return m;
}

function parseBorders<K extends string>(el: Element | null, keys: K[]): Partial<Record<K, Border>> | null {
  if (!el) return null;
  const out: Partial<Record<K, Border>> = {};
  for (const k of keys) {
    const name = k === "left" ? "left" : k === "right" ? "right" : k;
    let b = parseBorder(child(el, NS.w, name));
    if (!b && k === "left") b = parseBorder(child(el, NS.w, "start"));
    if (!b && k === "right") b = parseBorder(child(el, NS.w, "end"));
    if (b) out[k] = b;
  }
  return out;
}

export function parseTblPr(el: Element | null): TableProps {
  const p: TableProps = {};
  if (!el) return p;
  for (const c of children(el, NS.w)) {
    switch (c.localName) {
      case "tblStyle": p.tblStyle = wattr(c, "val"); break;
      case "tblW": { const w = wint(c, "w"); if (w !== null) p.width = { w, type: wattr(c, "type") || "dxa" }; break; }
      case "jc": p.jc = wattr(c, "val"); break;
      case "tblInd": { const w = wint(c, "w"); if (w !== null) p.indent = w; break; }
      case "tblBorders": p.borders = parseBorders(c, ["top", "bottom", "left", "right", "insideH", "insideV"]); break;
      case "tblCellMar": p.cellMar = parseMar(c); break;
      case "tblLayout": p.layout = wattr(c, "type"); break;
      case "tblLook": p.look = wattr(c, "val"); break;
      case "shd": { const fill = wattr(c, "fill"); if (fill && fill.toLowerCase() !== "auto") p.shd = fill; break; }
      case "tblCellSpacing": { const w = wint(c, "w"); if (w !== null) p.cellSpacing = w; break; }
    }
  }
  return p;
}

export function parseTrPr(el: Element | null): RowProps {
  const p: RowProps = {};
  if (!el) return p;
  for (const c of children(el, NS.w)) {
    switch (c.localName) {
      case "trHeight": { const v = wint(c, "val"); if (v !== null) p.height = v; p.hRule = wattr(c, "hRule") || "atLeast"; break; }
      case "tblHeader": p.header = onOff(c) ?? true; break;
      case "cantSplit": p.cantSplit = onOff(c) ?? true; break;
      case "jc": p.jc = wattr(c, "val"); break;
    }
  }
  return p;
}

export function parseTcPr(el: Element | null): CellProps {
  const p: CellProps = {};
  if (!el) return p;
  for (const c of children(el, NS.w)) {
    switch (c.localName) {
      case "tcW": { const w = wint(c, "w"); if (w !== null) p.width = { w, type: wattr(c, "type") || "dxa" }; break; }
      case "gridSpan": p.gridSpan = wint(c, "val") ?? 1; break;
      case "vMerge": p.vMerge = wattr(c, "val") || "continue"; break;
      case "tcBorders": p.borders = parseBorders(c, ["top", "bottom", "left", "right"]); break;
      case "shd": { const fill = wattr(c, "fill"); if (fill && fill.toLowerCase() !== "auto") p.shd = fill; break; }
      case "vAlign": p.vAlign = wattr(c, "val"); break;
      case "tcMar": p.mar = parseMar(c); break;
      case "noWrap": p.noWrap = onOff(c) ?? true; break;
      case "textDirection": p.textDirection = wattr(c, "val"); break;
    }
  }
  return p;
}

/** Shallow merge: keys in `over` that are defined win. */
export function merge<T extends object>(base: T | null | undefined, over: T | null | undefined): T {
  const out: any = Object.assign({}, base || {});
  if (over) for (const k of Object.keys(over)) { const v = (over as any)[k]; if (v !== undefined) out[k] = v; }
  return out as T;
}

// ---------------------------------------------------------------------------
// CSS generation

const FONT_STACKS: Record<string, string> = {
  "calibri": '"Calibri","Carlito",sans-serif',
  "calibri light": '"Calibri Light","Calibri","Carlito",sans-serif',
  "cambria": '"Cambria","Caladea",serif',
  "aptos": '"Aptos","Calibri","Carlito",sans-serif',
  "aptos display": '"Aptos Display","Aptos","Calibri Light","Carlito",sans-serif',
  "arial": '"Arial","Liberation Sans","Arimo",sans-serif',
  "helvetica": '"Helvetica","Arial","Liberation Sans",sans-serif',
  "times new roman": '"Times New Roman","Liberation Serif","Tinos",serif',
  "courier new": '"Courier New","Liberation Mono","Cousine",monospace',
  "consolas": '"Consolas","DejaVu Sans Mono",monospace',
  "segoe ui": '"Segoe UI","Noto Sans","DejaVu Sans",sans-serif',
  "verdana": '"Verdana","DejaVu Sans",sans-serif',
  "tahoma": '"Tahoma","DejaVu Sans",sans-serif',
  "georgia": '"Georgia","Gelasio",serif',
  "garamond": '"Garamond","EB Garamond",serif',
  "open sans": '"Open Sans","Noto Sans",sans-serif',
  "symbol": '"Symbol",serif',
  "wingdings": '"Wingdings",sans-serif',
};

export function fontStack(name: string | null | undefined): string {
  if (!name) return FONT_STACKS["calibri"];
  const key = name.toLowerCase();
  if (FONT_STACKS[key]) return FONT_STACKS[key];
  const generic = /serif|roman|garamond|georgia|book/i.test(name) && !/sans/i.test(name) ? "serif" : /mono|courier|consolas/i.test(name) ? "monospace" : "sans-serif";
  return `"${name.replace(/"/g, "")}",${generic}`;
}

/** Approximate ratio of a font's natural line height to its em size (used for Word line spacing). */
export function normalLineRatio(font: string | null | undefined): number {
  const k = (font || "calibri").toLowerCase();
  if (k.startsWith("calibri")) return 1.22;
  if (k.startsWith("cambria")) return 1.17;
  if (k.startsWith("aptos")) return 1.2;
  if (k === "arial" || k === "helvetica") return 1.15;
  if (k === "times new roman") return 1.15;
  if (k === "courier new") return 1.13;
  if (k === "segoe ui") return 1.33;
  if (k === "open sans") return 1.36;
  if (k === "verdana" || k === "tahoma") return 1.215;
  if (k === "georgia") return 1.14;
  return 1.2;
}

const HIGHLIGHTS: Record<string, string> = {
  yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff", blue: "#0000ff", red: "#ff0000",
  darkBlue: "#000080", darkCyan: "#008080", darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000",
  darkYellow: "#808000", darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000", white: "#ffffff",
};

// ---------------------------------------------------------------------------
// Dark mode: document colours are remapped by luminance so dark text becomes
// light and light fills become dark, while hue and saturation are kept.

let darkMode = false;
export function setDarkMode(on: boolean) { darkMode = on; }
export function isDarkMode() { return darkMode; }

function hexToHsl(hex: string): [number, number, number] {
  const r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
  }
  return [h, s, l];
}
function hslToHex(h: number, s: number, l: number): string {
  const hue2rgb = (p: number, q: number, t: number) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
  let r: number, g: number, b: number;
  if (s === 0) r = g = b = l;
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q; r = hue2rgb(p, q, h + 1 / 3); g = hue2rgb(p, q, h); b = hue2rgb(p, q, h - 1 / 3); }
  return "#" + [r, g, b].map((v) => Math.round(v * 255).toString(16).padStart(2, "0")).join("");
}

/** Text/line colour for the current theme. */
export function cssTextColor(c: string | null | undefined, fallback = ""): string {
  if (!c || c === "auto") return fallback;
  if (!/^[0-9a-fA-F]{6}$/.test(c)) return c;
  if (!darkMode) return "#" + c.toLowerCase();
  const [h, s, l] = hexToHsl(c);
  if (l >= 0.55) return "#" + c.toLowerCase();
  if (l <= 0.06 && s < 0.2) return "var(--paper-text)";
  // black -> light grey, mid-dark colours -> lighter tints of the same hue
  return hslToHex(h, s, Math.min(0.9, 0.9 - l * 0.5));
}

/** Fill/background colour for the current theme. */
export function cssBgColor(c: string | null | undefined, fallback = ""): string {
  if (!c || c === "auto") return fallback;
  if (!/^[0-9a-fA-F]{6}$/.test(c)) return c;
  if (!darkMode) return "#" + c.toLowerCase();
  const [h, s, l] = hexToHsl(c);
  if (l <= 0.45) return "#" + c.toLowerCase();
  if (l >= 0.96 && s < 0.2) return "var(--paper)";
  // white -> near page colour, light tints -> dark tints of the same hue
  return hslToHex(h, Math.min(s, 0.6), 0.14 + (1 - l) * 0.55);
}

/** Highlight colour (marker pen): translucent in dark mode so light text stays readable. */
export function cssHighlight(css: string): string {
  if (!darkMode) return css;
  const m = /^#([0-9a-f]{6})$/i.exec(css);
  if (!m) return css;
  const r = parseInt(m[1].slice(0, 2), 16), g = parseInt(m[1].slice(2, 4), 16), b = parseInt(m[1].slice(4, 6), 16);
  return `rgba(${r},${g},${b},0.38)`;
}

/** @deprecated use cssTextColor / cssBgColor */
export function cssColor(c: string | null | undefined, fallback = ""): string { return cssTextColor(c, fallback); }

export function runCss(p: RunProps): string {
  let s = "";
  if (p.font) s += `font-family:${fontStack(p.font)};`;
  if (p.size) s += `font-size:${fmt(halfPtToPt(p.size))}pt;`;
  if (p.b !== undefined) s += `font-weight:${p.b ? "bold" : "normal"};`;
  if (p.i !== undefined) s += `font-style:${p.i ? "italic" : "normal"};`;
  const deco: string[] = [];
  if (p.u && p.u !== "none") deco.push("underline");
  if (p.strike || p.dstrike) deco.push("line-through");
  if (deco.length) {
    s += `text-decoration-line:${deco.join(" ")};`;
    if (p.u === "double" || p.dstrike) s += "text-decoration-style:double;";
    else if (p.u === "dotted" || p.u === "dottedHeavy") s += "text-decoration-style:dotted;";
    else if (p.u && /dash/.test(p.u)) s += "text-decoration-style:dashed;";
    else if (p.u && /wave/.test(p.u)) s += "text-decoration-style:wavy;";
  } else if (p.u === "none" || p.strike === false) s += "text-decoration-line:none;";
  if (p.color !== undefined) s += `color:${cssTextColor(p.color, "inherit")};`;
  if (p.highlight && p.highlight !== "none") s += `background-color:${cssHighlight(HIGHLIGHTS[p.highlight] || p.highlight)};`;
  else if (p.shd) s += `background-color:${cssBgColor(p.shd)};`;
  if (p.vertAlign === "superscript") s += "vertical-align:super;font-size:smaller;";
  else if (p.vertAlign === "subscript") s += "vertical-align:sub;font-size:smaller;";
  if (p.caps) s += "text-transform:uppercase;";
  if (p.smallCaps) s += "font-variant-caps:small-caps;";
  if (p.vanish) s += "display:none;";
  if (p.spacing) s += `letter-spacing:${fmt(twipsToPt(p.spacing))}pt;`;
  if (p.rtl) s += "direction:rtl;unicode-bidi:embed;";
  return s;
}

const BORDER_STYLES: Record<string, string> = {
  single: "solid", thick: "solid", double: "double", dotted: "dotted", dashed: "dashed", dashSmallGap: "dashed",
  dotDash: "dashed", dotDotDash: "dashed", triple: "double", thinThickSmallGap: "double", thickThinSmallGap: "double",
  thinThickThinSmallGap: "double", thinThickMediumGap: "double", thickThinMediumGap: "double", wave: "solid",
  doubleWave: "double", inset: "inset", outset: "outset", nil: "none", none: "none",
};

export function borderCss(b: Border | null | undefined): string {
  if (!b) return "";
  const style = BORDER_STYLES[b.val] || "solid";
  if (style === "none") return "none";
  const w = Math.max(0.75, b.sz / 8) * (96 / 72);
  return `${fmt(w)}px ${style} ${cssTextColor(b.color, "currentColor")}`;
}

export const AUTO_SPACING_TWIPS = 280; // Word's "auto" paragraph spacing (14pt)

/**
 * CSS for a paragraph from its effective props. `font`/`size` are the effective
 * run font and size in the paragraph (used for line-height calculations).
 */
export function paraCss(p: ParaProps, effFont: string | null | undefined, effSizeHalfPt: number, inTable: boolean): string {
  let s = "";
  switch (p.jc) {
    case "center": s += "text-align:center;"; break;
    case "right": case "end": s += "text-align:right;"; break;
    case "both": case "distribute": s += "text-align:justify;"; break;
    case "left": case "start": s += "text-align:left;"; break;
  }
  const left = p.indLeft || 0, right = p.indRight || 0;
  let textIndent = 0;
  if (p.indHanging) textIndent = -p.indHanging;
  else if (p.indFirstLine) textIndent = p.indFirstLine;
  if (left) s += `margin-left:${fmt(twipsToPt(left))}pt;`;
  if (right) s += `margin-right:${fmt(twipsToPt(right))}pt;`;
  if (textIndent) s += `text-indent:${fmt(twipsToPt(textIndent))}pt;`;
  const before = p.spBeforeAuto ? AUTO_SPACING_TWIPS : (p.spBefore || 0);
  const after = p.spAfterAuto ? (inTable ? 0 : AUTO_SPACING_TWIPS) : (p.spAfter || 0);
  s += `margin-top:${fmt(twipsToPt(before))}pt;margin-bottom:${fmt(twipsToPt(after))}pt;`;
  const rule = p.spLineRule || "auto";
  const line = p.spLine ?? 240;
  if (rule === "auto") {
    if (line === 240) s += "line-height:normal;";
    else s += `line-height:${fmt((line / 240) * normalLineRatio(effFont), 3)};`;
  } else if (rule === "exact") {
    s += `line-height:${fmt(twipsToPt(line))}pt;`;
  } else {
    // atLeast: browsers can't express it; use the larger of natural and requested.
    const natural = halfPtToPt(effSizeHalfPt) * normalLineRatio(effFont);
    const req = twipsToPt(line);
    s += req > natural ? `line-height:${fmt(req)}pt;` : "line-height:normal;";
  }
  if (p.shd) s += `background-color:${cssBgColor(p.shd)};`;
  const bt = borderCss(p.bdrTop), bb = borderCss(p.bdrBottom), bl = borderCss(p.bdrLeft), br = borderCss(p.bdrRight);
  if (bt && bt !== "none") s += `border-top:${bt};padding-top:${p.bdrTop!.space}pt;`;
  if (bb && bb !== "none") s += `border-bottom:${bb};padding-bottom:${p.bdrBottom!.space}pt;`;
  if (bl && bl !== "none") s += `border-left:${bl};padding-left:${p.bdrLeft!.space}pt;`;
  if (br && br !== "none") s += `border-right:${br};padding-right:${p.bdrRight!.space}pt;`;
  if (p.bidi) s += "direction:rtl;";
  return s;
}
