// .xlsx reader: workbook, shared strings, styles, theme, worksheets. Unknown
// worksheet children are kept as raw XML for a faithful round trip.
import { Package, parseRels, resolveTarget, Relationship } from "../docx/zip";
import { parseXml, children, child, serialize, rootNamespaceDecls, attr as xattr } from "../docx/xml";
import {
  Workbook, Sheet, Styles, Xf, Font, Fill, Border, BorderSide, Color, Alignment, DefinedName, RawBlock, Cell, Value,
  newSheet, key, colIndex, updateExtent, parseRef, DEFAULT_FONT,
} from "./model";
import { shiftFormula } from "./formula/tokens";
import { dateToSerial } from "./numfmt";

const NS_SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_SST = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const REL_CALC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
const REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

const a = (el: Element | null | undefined, name: string): string | null => (el ? el.getAttribute(name) : null);
const num = (el: Element | null | undefined, name: string): number | null => { const v = a(el, name); if (v === null) return null; const n = parseFloat(v); return isNaN(n) ? null : n; };
const bool = (el: Element | null | undefined, name: string, def = false): boolean => { const v = a(el, name); if (v === null) return def; return v === "1" || v === "true"; };
const kids = (el: Element | null | undefined, name?: string) => children(el, NS_SS, name);
const kid = (el: Element | null | undefined, name: string) => child(el, NS_SS, name);

let sstRawRef: string[] = [];

export function emptyStyles(): Styles {
  return {
    numFmts: new Map(),
    fonts: [{ ...DEFAULT_FONT }],
    fills: [{ patternType: "none", fg: null, bg: null }, { patternType: "gray125", fg: null, bg: null }],
    borders: [{ left: null, right: null, top: null, bottom: null, diagonal: null }],
    cellStyleXfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, alignment: null, protection: null }],
    xfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, alignment: null, protection: null }],
    rawTail: [{ name: "cellStyles", xml: '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' }],
    rootAttrs: `xmlns="${NS_SS}"`,
  };
}

function parseColor(el: Element | null): Color | null {
  if (!el) return null;
  const c: Color = {};
  const rgb = a(el, "rgb"); if (rgb) c.rgb = rgb;
  const theme = num(el, "theme"); if (theme !== null) c.theme = theme;
  const tint = num(el, "tint"); if (tint !== null) c.tint = tint;
  const indexed = num(el, "indexed"); if (indexed !== null) c.indexed = indexed;
  if (bool(el, "auto")) c.auto = true;
  return c;
}

function parseFont(el: Element, decls: Map<string, string>): Font {
  const f: Font = { ...DEFAULT_FONT, raw: serialize(el, decls) };
  for (const c of kids(el)) {
    switch (c.localName) {
      case "b": f.bold = bool(c, "val", true); break;
      case "i": f.italic = bool(c, "val", true); break;
      case "u": f.underline = a(c, "val") || "single"; break;
      case "strike": f.strike = bool(c, "val", true); break;
      case "sz": f.size = num(c, "val") ?? 11; break;
      case "color": f.color = parseColor(c); break;
      case "name": f.name = a(c, "val") || "Calibri"; break;
      case "family": f.family = num(c, "val") ?? undefined; break;
      case "scheme": f.scheme = a(c, "val") || undefined; break;
      case "charset": f.charset = num(c, "val") ?? undefined; break;
    }
  }
  return f;
}

function parseFill(el: Element, decls: Map<string, string>): Fill {
  const p = kid(el, "patternFill");
  const fill: Fill = { patternType: a(p, "patternType") || (p ? "solid" : "none"), fg: parseColor(kid(p, "fgColor")), bg: parseColor(kid(p, "bgColor")), raw: serialize(el, decls) };
  if (!p) { const g = kid(el, "gradientFill"); if (g) { fill.patternType = "gradient"; const stops = kids(g, "stop"); if (stops.length) fill.fg = parseColor(kid(stops[0], "color")); } }
  return fill;
}

function parseSide(el: Element | null): BorderSide | null {
  if (!el) return null;
  const style = a(el, "style");
  if (!style || style === "none") return null;
  return { style, color: parseColor(kid(el, "color")) };
}
function parseBorder(el: Element, decls: Map<string, string>): Border {
  return { left: parseSide(kid(el, "left") || kid(el, "start")), right: parseSide(kid(el, "right") || kid(el, "end")), top: parseSide(kid(el, "top")), bottom: parseSide(kid(el, "bottom")), diagonal: parseSide(kid(el, "diagonal")), diagonalUp: bool(el, "diagonalUp"), diagonalDown: bool(el, "diagonalDown"), raw: serialize(el, decls) };
}

function parseXf(el: Element): Xf {
  const al = kid(el, "alignment");
  let alignment: Alignment | null = null;
  if (al) {
    alignment = {};
    const h = a(al, "horizontal"); if (h) alignment.horizontal = h;
    const v = a(al, "vertical"); if (v) alignment.vertical = v;
    if (bool(al, "wrapText")) alignment.wrapText = true;
    const ind = num(al, "indent"); if (ind) alignment.indent = ind;
    const rot = num(al, "textRotation"); if (rot) alignment.textRotation = rot;
    if (bool(al, "shrinkToFit")) alignment.shrinkToFit = true;
    const ro = num(al, "readingOrder"); if (ro) alignment.readingOrder = ro;
  }
  const pr = kid(el, "protection");
  const protection = pr ? { locked: a(pr, "locked") === null ? undefined : bool(pr, "locked", true), hidden: bool(pr, "hidden") } : null;
  const xf: Xf = { numFmtId: num(el, "numFmtId") ?? 0, fontId: num(el, "fontId") ?? 0, fillId: num(el, "fillId") ?? 0, borderId: num(el, "borderId") ?? 0, xfId: num(el, "xfId") ?? 0, alignment, protection };
  for (const k of ["applyNumberFormat", "applyFont", "applyFill", "applyBorder", "applyAlignment", "applyProtection", "quotePrefix"] as const) if (a(el, k) !== null) (xf as any)[k] = bool(el, k);
  return xf;
}

export function parseStyles(xml: string | undefined): Styles {
  const st = emptyStyles();
  if (!xml) return st;
  const doc = parseXml(xml);
  const root = doc.documentElement;
  const decls = rootNamespaceDecls(root);
  st.rootAttrs = Array.from(root.attributes).map((at) => `${at.name}="${at.value.replace(/"/g, "&quot;")}"`).join(" ");
  st.rawTail = [];
  st.fonts = []; st.fills = []; st.borders = []; st.cellStyleXfs = []; st.xfs = [];
  for (const c of kids(root)) {
    switch (c.localName) {
      case "numFmts": for (const f of kids(c, "numFmt")) { const id = num(f, "numFmtId"); const code = a(f, "formatCode"); if (id !== null && code !== null) st.numFmts.set(id, code); } break;
      case "fonts": for (const f of kids(c, "font")) st.fonts.push(parseFont(f, decls)); break;
      case "fills": for (const f of kids(c, "fill")) st.fills.push(parseFill(f, decls)); break;
      case "borders": for (const b of kids(c, "border")) st.borders.push(parseBorder(b, decls)); break;
      case "cellStyleXfs": for (const x of kids(c, "xf")) st.cellStyleXfs.push(parseXf(x)); break;
      case "cellXfs": for (const x of kids(c, "xf")) st.xfs.push(parseXf(x)); break;
      default: st.rawTail.push({ name: c.localName, xml: serialize(c, decls) });
    }
  }
  if (!st.fonts.length) st.fonts.push({ ...DEFAULT_FONT });
  if (!st.fills.length) st.fills.push({ patternType: "none", fg: null, bg: null }, { patternType: "gray125", fg: null, bg: null });
  if (!st.borders.length) st.borders.push({ left: null, right: null, top: null, bottom: null, diagonal: null });
  if (!st.cellStyleXfs.length) st.cellStyleXfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, alignment: null, protection: null });
  if (!st.xfs.length) st.xfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, alignment: null, protection: null });
  return st;
}

export function parseTheme(xml: string | undefined): string[] {
  const def = ["000000", "FFFFFF", "44546A", "E7E6E6", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"];
  if (!xml) return def;
  try {
    const doc = parseXml(xml);
    const scheme = doc.getElementsByTagNameNS(NS_A, "clrScheme")[0];
    if (!scheme) return def;
    const order = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
    const out = [...def];
    for (const c of children(scheme, NS_A)) {
      const i = order.indexOf(c.localName);
      if (i < 0) continue;
      const srgb = child(c, NS_A, "srgbClr")?.getAttribute("val");
      const sys = child(c, NS_A, "sysClr")?.getAttribute("lastClr");
      const v = srgb || sys;
      if (v) out[i] = v.toUpperCase();
    }
    return out;
  } catch { return def; }
}

function parseSharedStrings(xml: string | undefined, decls?: Map<string, string>): { sst: string[]; sstRaw: string[] } {
  const sst: string[] = [], sstRaw: string[] = [];
  if (!xml) return { sst, sstRaw };
  const doc = parseXml(xml);
  const rootDecls = decls || rootNamespaceDecls(doc.documentElement);
  for (const si of kids(doc.documentElement, "si")) {
    let text = "";
    let rich = false;
    for (const c of kids(si)) {
      if (c.localName === "t") text += c.textContent || "";
      else if (c.localName === "r") { rich = true; for (const t of kids(c, "t")) text += t.textContent || ""; }
    }
    sst.push(text);
    sstRaw.push(rich ? serialize(si, rootDecls) : "");
  }
  return { sst, sstRaw };
}

// ---- worksheet ---------------------------------------------------------------

interface SheetCtx { wb: Workbook; rels: Map<string, Relationship>; }

function parseSheetXml(sheet: Sheet, xml: string, ctxs: SheetCtx) {
  const doc = parseXml(xml);
  const root = doc.documentElement;
  const decls = rootNamespaceDecls(root);
  sheet.rawRootAttrs = Array.from(root.attributes).map((at) => `${at.name}="${at.value.replace(/"/g, "&quot;")}"`).join(" ");
  const sst = ctxs.wb.sst;
  const sharedMasters = new Map<number, { f: string; r: number; c: number }>();
  for (const el of kids(root)) {
    switch (el.localName) {
      case "dimension": break; // recomputed from cells
      case "sheetViews": {
        const sv = kid(el, "sheetView");
        if (sv) {
          sheet.view.showGridLines = bool(sv, "showGridLines", true);
          sheet.view.zoom = num(sv, "zoomScale") ?? 100;
          sheet.view.tabSelected = bool(sv, "tabSelected");
          sheet.view.topLeft = a(sv, "topLeftCell") || undefined;
          sheet.view.rightToLeft = bool(sv, "rightToLeft");
          const pane = kid(sv, "pane");
          if (pane && a(pane, "state") === "frozen") {
            sheet.freeze = { rows: num(pane, "ySplit") ?? 0, cols: num(pane, "xSplit") ?? 0 };
          }
          const sel = kids(sv, "selection");
          const s0 = sel.find((s) => a(s, "pane") === (pane ? a(pane, "activePane") : null)) || sel[sel.length - 1];
          if (s0) { sheet.view.active = a(s0, "activeCell") || undefined; sheet.view.selection = a(s0, "sqref") || undefined; }
        }
        break;
      }
      case "sheetFormatPr":
        sheet.defaultRowHeight = num(el, "defaultRowHeight") ?? 15;
        { const w = num(el, "defaultColWidth"); const base = num(el, "baseColWidth"); if (w !== null) sheet.defaultColWidth = w; else if (base !== null) sheet.defaultColWidth = base + 0.71; }
        break;
      case "cols":
        for (const c of kids(el, "col")) {
          const min = (num(c, "min") ?? 1) - 1, max = (num(c, "max") ?? 1) - 1;
          const width = num(c, "width");
          sheet.cols.push({ min, max, width: width ?? undefined, hidden: bool(c, "hidden") || undefined, style: num(c, "style") ?? undefined, customWidth: bool(c, "customWidth") || undefined, bestFit: bool(c, "bestFit") || undefined, level: num(c, "outlineLevel") ?? undefined });
        }
        break;
      case "sheetData": {
        let r = -1;
        for (const rowEl of kids(el, "row")) {
          const rn = num(rowEl, "r");
          r = rn !== null ? rn - 1 : r + 1;
          const ht = num(rowEl, "ht"), hidden = bool(rowEl, "hidden"), style = num(rowEl, "s"), custom = bool(rowEl, "customHeight"), level = num(rowEl, "outlineLevel");
          if (ht !== null || hidden || style !== null || level) sheet.rows.set(r, { height: ht ?? undefined, hidden: hidden || undefined, style: style ?? undefined, customHeight: custom || undefined, level: level ?? undefined });
          let c = -1;
          for (const cEl of kids(rowEl, "c")) {
            const ref = a(cEl, "r");
            if (ref) { const m = /^([A-Za-z]+)(\d+)$/.exec(ref); if (m) { c = colIndex(m[1].toUpperCase()); r = parseInt(m[2], 10) - 1; } else c++; } else c++;
            const cell = parseCell(cEl, sst, ctxs.wb.date1904, decls, sharedMasters, r, c);
            if (cell) { sheet.cells.set(key(r, c), cell); updateExtent(sheet, r, c); }
          }
        }
        break;
      }
      case "mergeCells":
        for (const m of kids(el, "mergeCell")) { const ref = parseRef(a(m, "ref") || ""); if (ref) sheet.merges.push({ r1: ref.r1, c1: ref.c1, r2: ref.r2, c2: ref.c2 }); }
        break;
      case "autoFilter": { const ref = parseRef(a(el, "ref") || ""); if (ref) sheet.autoFilter = { r1: ref.r1, c1: ref.c1, r2: ref.r2, c2: ref.c2 }; sheet.rawBlocks.push({ name: "autoFilter", xml: serialize(el, decls) }); break; }
      case "hyperlinks":
        for (const h of kids(el, "hyperlink")) {
          const rId = xattr(h, NS_R, "id");
          const rel = rId ? ctxs.rels.get(rId) : undefined;
          sheet.hyperlinks.push({ ref: a(h, "ref") || "", rId: rId || undefined, target: rel?.target, location: a(h, "location") || undefined, display: a(h, "display") || undefined, tooltip: a(h, "tooltip") || undefined });
        }
        break;
      case "sheetPr": {
        const tc = kid(el, "tabColor");
        if (tc) { const col = parseColor(tc); sheet.tabColor = col?.rgb ? col.rgb.slice(-6) : null; }
        sheet.rawBlocks.push({ name: "sheetPr", xml: serialize(el, decls) });
        break;
      }
      default:
        sheet.rawBlocks.push({ name: el.localName, xml: serialize(el, decls) });
    }
  }
  // Frozen header for viewing when the file has no pane.
  if (!sheet.freeze && sheet.maxRow > 0) { sheet.freeze = { rows: 1, cols: 0 }; sheet.autoFreeze = true; }
}

function parseCell(cEl: Element, sst: string[], date1904: boolean, decls: Map<string, string>, sharedMasters: Map<number, { f: string; r: number; c: number }>, r: number, c: number): Cell | null {
  const t = a(cEl, "t") || "n";
  const s = num(cEl, "s") ?? 0;
  let v: Value = null;
  let f: string | undefined;
  let sh: Cell["sh"];
  let arr: string | undefined;
  let rich: string | undefined;
  const fEl = kid(cEl, "f");
  if (fEl) {
    const ft = a(fEl, "t");
    const text = fEl.textContent || "";
    if (ft === "shared") {
      const si = num(fEl, "si") ?? 0;
      const ref = a(fEl, "ref");
      if (ref) { sharedMasters.set(si, { f: text, r, c }); f = text; sh = { si, ref, master: true }; }
      else {
        const m = sharedMasters.get(si);
        f = m ? shiftFormula(m.f, r - m.r, c - m.c) : text;
        sh = { si };
      }
    } else if (ft === "array") { f = text; arr = a(fEl, "ref") || undefined; }
    else f = text;
  }
  const vEl = kid(cEl, "v");
  const vText = vEl ? vEl.textContent || "" : null;
  switch (t) {
    case "s": { const i = vText !== null ? parseInt(vText, 10) : -1; v = i >= 0 && i < sst.length ? sst[i] : ""; if (i >= 0 && sstRawRef[i]) rich = sstRawRef[i].replace(/^<(\w+:)?si\b/, "<is").replace(/<\/(\w+:)?si>$/, "</is>"); break; }
    case "str": v = vText ?? ""; break;
    case "inlineStr": {
      const is = kid(cEl, "is");
      let text = "";
      let isRich = false;
      if (is) for (const k of kids(is)) { if (k.localName === "t") text += k.textContent || ""; else if (k.localName === "r") { isRich = true; for (const tt of kids(k, "t")) text += tt.textContent || ""; } }
      v = text;
      if (isRich && is) rich = serialize(is, decls);
      break;
    }
    case "b": v = vText === "1" || vText === "true"; break;
    case "e": v = { e: vText || "#VALUE!" }; break;
    case "d": { const d = vText ? new Date(vText) : null; v = d && !isNaN(d.getTime()) ? dateToSerial(d, date1904) : null; break; }
    default: v = vText !== null && vText !== "" ? parseFloat(vText) : null;
  }
  if (v === null && !f && s === 0) return null;
  const cell: Cell = { v, s };
  if (f !== undefined) cell.f = f;
  if (sh) cell.sh = sh;
  if (arr) cell.arr = arr;
  if (rich) cell.rich = rich;
  return cell;
}

// ---- workbook ------------------------------------------------------------------

export function loadXlsx(bytes: Uint8Array, path: string | null): Workbook {
  const pkg = Package.fromBytes(bytes);
  const pkgRels = parseRels(pkg.text("_rels/.rels"));
  let wbPart = "xl/workbook.xml";
  for (const r of pkgRels.values()) if (r.type === REL_OFFICE) { wbPart = resolveTarget("", r.target); break; }
  const workbookXml = pkg.text(wbPart);
  if (!workbookXml) throw new Error("Not a spreadsheet (missing " + wbPart + ")");
  const wbDir = wbPart.split("/").slice(0, -1).join("/");
  const relsPart = wbDir + "/_rels/" + wbPart.split("/").pop() + ".rels";
  const workbookRels = pkg.text(relsPart) || "";
  const rels = parseRels(workbookRels);
  const partFor = (type: string): string | null => { for (const r of rels.values()) if (r.type === type) return resolveTarget(wbPart, r.target); return null; };
  const stylesPart = partFor(REL_STYLES) || wbDir + "/styles.xml";
  const sstPart = partFor(REL_SST);
  const themePart = partFor(REL_THEME);
  const calcChainPart = partFor(REL_CALC);

  const styles = parseStyles(pkg.text(stylesPart));
  const themeColors = parseTheme(themePart ? pkg.text(themePart) : undefined);
  const { sst, sstRaw } = parseSharedStrings(sstPart ? pkg.text(sstPart) : undefined);
  const sstIndex = new Map<string, number>();
  sst.forEach((s, i) => { if (!sstRaw[i] && !sstIndex.has(s)) sstIndex.set(s, i); }); // rich entries are never matched by plain text
  sstRawRef = sstRaw;

  const wbDoc = parseXml(workbookXml);
  const wbRoot = wbDoc.documentElement;
  const pr = kid(wbRoot, "workbookPr");
  const date1904 = bool(pr, "date1904");
  const wb: Workbook = {
    sheets: [], active: 0, date1904, definedNames: [], styles, sst, sstRaw, sstIndex, themeColors, pkg,
    workbookPart: wbPart, workbookXml, workbookRels, calcChainPart, stylesPart, sstPart, path, kind: "xlsx",
  };
  const bv = kid(kid(wbRoot, "bookViews"), "workbookView");
  const activeTab = num(bv, "activeTab") ?? 0;
  const dn = kid(wbRoot, "definedNames");
  if (dn) for (const d of kids(dn, "definedName")) wb.definedNames.push({ name: a(d, "name") || "", ref: d.textContent || "", localSheetId: num(d, "localSheetId"), hidden: bool(d, "hidden"), raw: serialize(d, rootNamespaceDecls(wbRoot)) });

  const sheetsEl = kid(wbRoot, "sheets");
  let idx = 0;
  for (const s of kids(sheetsEl, "sheet")) {
    const rId = xattr(s, NS_R, "id") || "";
    const rel = rels.get(rId);
    const part = rel ? resolveTarget(wbPart, rel.target) : wbDir + "/worksheets/sheet" + (idx + 1) + ".xml";
    const sheet = newSheet(a(s, "name") || "Sheet" + (idx + 1), num(s, "sheetId") ?? idx + 1, rId, part);
    const state = a(s, "state"); if (state === "hidden" || state === "veryHidden") sheet.state = state;
    const xml = pkg.text(part);
    const sheetRelsPart = part.split("/").slice(0, -1).join("/") + "/_rels/" + part.split("/").pop() + ".rels";
    sheet.relsXml = pkg.text(sheetRelsPart) || null;
    const sheetRels = parseRels(sheet.relsXml || undefined);
    if (xml) parseSheetXml(sheet, xml, { wb, rels: sheetRels });
    wb.sheets.push(sheet);
    idx++;
  }
  if (!wb.sheets.length) throw new Error("Workbook has no sheets");
  wb.active = Math.min(activeTab, wb.sheets.length - 1);
  const firstVisible = wb.sheets.findIndex((s) => s.state === "visible");
  if (wb.sheets[wb.active].state !== "visible" && firstVisible >= 0) wb.active = firstVisible;
  for (const d of wb.definedNames) d.sheet = d.localSheetId !== null ? wb.sheets[d.localSheetId] || null : null;
  return wb;
}

export { REL_WORKSHEET, REL_HYPERLINK, NS_SS };
