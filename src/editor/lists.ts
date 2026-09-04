// List labels (bullets / numbers) rendered through node decorations, plus
// Word's "contextual spacing" (no space between paragraphs of the same style).
import { Plugin, PluginKey, EditorState } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { Node as PMNode } from "prosemirror-model";
import { schema, paragraphStyle } from "../schema";
import { ctx, NumLevel } from "../docx/styles";
import { fontStack, cssTextColor } from "../docx/props";
import { mapSymbolChar } from "../docx/parse";
import { twipsToPt, fmt, halfPtToPt } from "../docx/units";

export const numberingKey = new PluginKey<DecorationSet>("numbering");

interface ParaInfo {
  numId: number | null; ilvl: number; hanging: number; firstLine: number; left: number;
  pStyle: string | null; contextual: boolean; before: number; after: number;
}
const infoCache = new WeakMap<PMNode, ParaInfo>();

function info(node: PMNode): ParaInfo {
  let i = infoCache.get(node);
  if (i) return i;
  const eff = paragraphStyle(node).pPr;
  i = {
    numId: eff.numId || null, ilvl: eff.ilvl || 0,
    hanging: eff.indHanging || 0, firstLine: eff.indFirstLine || 0, left: eff.indLeft || 0,
    pStyle: eff.pStyle || ctx.defaultPara, contextual: !!eff.contextual,
    before: eff.spBefore || 0, after: eff.spAfter || 0,
  };
  infoCache.set(node, i);
  return i;
}

export function clearNumberingCache() { /* WeakMap: nothing to do, nodes are new per document */ }

function toRoman(n: number): string {
  if (n <= 0 || n >= 4000) return String(n);
  const v = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
  const s = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
  let out = "";
  for (let i = 0; i < v.length; i++) while (n >= v[i]) { out += s[i]; n -= v[i]; }
  return out;
}
function toLetter(n: number): string {
  if (n <= 0) return String(n);
  // Word: A..Z, AA..ZZ (repeat letter), not base-26.
  const idx = (n - 1) % 26, rep = Math.floor((n - 1) / 26) + 1;
  return String.fromCharCode(65 + idx).repeat(rep);
}
const ONES = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"];
const TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"];
function cardinal(n: number): string {
  if (n < 20) return ONES[n];
  if (n < 100) return TENS[Math.floor(n / 10)] + (n % 10 ? "-" + ONES[n % 10] : "");
  return String(n);
}

export function formatNumber(fmtName: string, n: number): string {
  switch (fmtName) {
    case "decimal": return String(n);
    case "decimalZero": return n < 10 ? "0" + n : String(n);
    case "upperRoman": return toRoman(n);
    case "lowerRoman": return toRoman(n).toLowerCase();
    case "upperLetter": return toLetter(n);
    case "lowerLetter": return toLetter(n).toLowerCase();
    case "ordinal": { const s = ["th", "st", "nd", "rd"], v = n % 100; return n + (s[(v - 20) % 10] || s[v] || s[0]); }
    case "cardinalText": { const c = cardinal(n); return c.charAt(0).toUpperCase() + c.slice(1); }
    case "ordinalText": { const c = cardinal(n); return (c.charAt(0).toUpperCase() + c.slice(1)).replace(/y$/, "ie") + (c.endsWith("t") ? "h" : "th"); }
    case "bullet": return "";
    case "none": return "";
    case "decimalEnclosedCircle": return n >= 1 && n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n);
    case "decimalEnclosedParen": case "decimalEnclosedFullstop": return String(n);
    default: return String(n);
  }
}

interface Counters { [abstractId: number]: (number | undefined)[]; }

function labelFor(lvl: NumLevel, counters: (number | undefined)[], ilvl: number): string {
  if (lvl.numFmt === "bullet") return mapSymbolChar(lvl.rPr.font, lvl.lvlText || "•");
  if (lvl.numFmt === "none") return "";
  let text = lvl.lvlText || `%${ilvl + 1}.`;
  text = text.replace(/%(\d)/g, (_, d) => {
    const li = parseInt(d, 10) - 1;
    const ref = ctxLevel(lvl, li);
    const val = counters[li] ?? (ref ? ref.start : 1);
    const fmtName = lvl.isLgl && li < ilvl ? "decimal" : ref ? ref.numFmt : "decimal";
    return formatNumber(fmtName, val);
  });
  return text;
}

// The level definitions for other levels are needed when formatting "%1.%2";
// we resolve through a small per-walk cache set by compute().
let currentNumId = 0;
function ctxLevel(_lvl: NumLevel, ilvl: number): NumLevel | null { return ctx.numLevel(currentNumId, ilvl); }

export function computeNumbering(doc: PMNode): DecorationSet {
  const decos: Decoration[] = [];
  const counters: Counters = {};
  const seenNums = new Set<number>();
  doc.descendants((node, pos, parent, index) => {
    if (node.type !== schema.nodes.paragraph) return true;
    const inf = info(node);
    let style = "";
    let cls = "";
    let label: string | null = null;
    if (inf.numId) {
      const num = ctx.nums.get(inf.numId);
      const lvl = ctx.numLevel(inf.numId, inf.ilvl);
      if (num && lvl) {
        const absId = num.abstractNumId;
        let c = counters[absId];
        if (!c) c = counters[absId] = [];
        if (!seenNums.has(inf.numId)) {
          seenNums.add(inf.numId);
          for (const [l, ov] of num.overrides) if (ov.start !== null) c[l] = ov.start - 1;
        }
        const startOf = (l: number) => ctx.numStart(inf.numId!, l);
        const level = inf.ilvl;
        c[level] = (c[level] ?? startOf(level) - 1) + 1;
        for (let l = level + 1; l < 9; l++) c[l] = undefined;
        currentNumId = inf.numId;
        label = labelFor(lvl, c, level);
        if (lvl.suff === "space") label += " ";
        cls = "om-num";
        // Label box: the hanging indent (or a default tab stop) between number and text.
        if (inf.hanging > 0) style += `--lbl-w:${fmt(twipsToPt(inf.hanging))}pt;`;
        else if (lvl.suff !== "nothing") style += `--lbl-pad:${fmt(twipsToPt(ctx.defaultTabStop || 720) / 2)}pt;`;
        const r = lvl.rPr;
        if (r.font) style += `--lbl-font:${fontStack(r.font)};`;
        if (r.size) style += `--lbl-size:${fmt(halfPtToPt(r.size))}pt;`;
        if (r.b !== undefined) style += `--lbl-bold:${r.b ? "bold" : "normal"};`;
        if (r.i !== undefined) style += `--lbl-italic:${r.i ? "italic" : "normal"};`;
        if (r.color && r.color !== "auto") style += `--lbl-color:${cssTextColor(r.color)};`;
        if (lvl.lvlJc === "right") style += "--lbl-jc:right;";
      }
    }
    // Contextual spacing against siblings of the same style.
    if (inf.contextual && parent) {
      const prev = index > 0 ? parent.child(index - 1) : null;
      const next = index + 1 < parent.childCount ? parent.child(index + 1) : null;
      if (prev && prev.type === schema.nodes.paragraph && info(prev).pStyle === inf.pStyle && inf.before) style += "margin-top:0;";
      if (next && next.type === schema.nodes.paragraph && info(next).pStyle === inf.pStyle && inf.after) style += "margin-bottom:0;";
    }
    if (label !== null || style) {
      const attrs: Record<string, string> = {};
      if (label !== null) { attrs["data-lbl"] = label; attrs.class = cls; }
      if (style) attrs.style = style;
      decos.push(Decoration.node(pos, pos + node.nodeSize, attrs));
    }
    return true;
  });
  return DecorationSet.create(doc, decos);
}

export function numberingPlugin(): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key: numberingKey,
    state: {
      init: (_, state) => computeNumbering(state.doc),
      apply: (tr, old) => (tr.docChanged || tr.getMeta(numberingKey) === "refresh") ? computeNumbering(tr.doc) : old,
    },
    props: {
      decorations(state: EditorState) { return numberingKey.getState(state); },
    },
  });
}
